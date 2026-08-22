"""
FitOrder 출력 생성기 — 장부 / 작업지시서 / 경영박사

    from output import build_all
    build_all(orders, ship_date="목", out_dir="../out")

orders = extract_order() 또는 parsers 결과 리스트
"""
import re
from copy import copy
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from rules import (CLIENT_INFO, ERP_CLIENT_NAME, NO_DITTO, NO_MERGE_PLACE,
                   PACKING_CLIENTS, PACKING_ITEM,
                   Master, calc_erp, default_handle_length,
                   ledger_color, normalize_handle)

# ── 서식 상수 (원본 장부.xls 에서 추출) ──
FONT_DATA = "THE행복열매"
FONT_UI = "맑은 고딕"
ROW_H = 20.1
LEDGER_ROWS = 35              # 3행 ~ 37행
SHEET_ROWS = 16               # 작업지시서 블록당 데이터 행 수
COL_W = [3.5, 15.1, 16.6, 6.6, 2.6, 6.6, 6.1, 6.0, 6.0, 6.1, 6.1, 6.1]
#        A         B(상호)   C(색상)  D(가로)  E(X)      F(세로)
#        G(수량)   H(모형1)  I(모형2) J(기재)  K(기재2)  L(출고일)
ALIGN = ["center", "right", "left", "right", "center", "left",
         "center", "left", "right", "center", "center", "center"]
DETAIL_COL_W = [3.5, 14.125, 16.625, 6.625, 2.625, 6.625, 4.875,
                5.125, 4.875, 5.625, 5.625, 5.625, 5.625]
# 세부 장부: A / 상호 / 색상 / 가로 / X / 세로 / 수량 / 방향 / 길이 /
#           특이 / 기재사항1 / 기재사항2 / 출고일
DETAIL_ALIGN = ["center", "right", "left", "right", "center", "left",
                "left", "center", "center", "right", "center", "center",
                "center"]
# 그룹 안쪽에는 세로선이 없다: 규격(D,E,F) / 모형(H,I) / 기재사항(J,K)
NO_LEFT = {4, 5, 8, 10}      # 0-based col index
NO_RIGHT = {3, 4, 7, 9}
HAIR = Side(style="hair")
THIN = Side(style="thin")
GREEN, RED, BLUE = "FF008000", "FFFF0000", "FF0000FF"


def _al(i):
    h = ALIGN[i]
    return Alignment(horizontal=(None if h == "general" else h),
                     vertical="center", shrink_to_fit=True)


def _border(r, c, last_row, top=None):
    left = None if c in NO_LEFT else (THIN if c == 0 else HAIR)
    right = None if c in NO_RIGHT else (THIN if c == 11 else HAIR)
    return Border(top=(top if top is not None else (THIN if r == 2 else HAIR)),
                  bottom=HAIR if r < last_row else THIN,
                  left=left, right=right)


def _tb(text, color=None, sz=14):
    """글꼴을 명시한 텍스트 블록 (기본 글꼴로 떨어지는 것을 방지)"""
    return TextBlock(InlineFont(rFont=FONT_DATA, sz=sz, color=color), text)


def _rich(text, word, color):
    """text 안의 word 부분만 색을 넣은 리치텍스트"""
    if not text or word not in text:
        return text
    i = text.index(word)
    parts = []
    if i:
        parts.append(_tb(text[:i]))
    parts.append(_tb(word, color))
    tail = text[i + len(word):]
    if tail:
        parts.append(_tb(tail))
    return CellRichText(*parts)


def _client_cell(client, mark):
    """상호: 코드는 14pt, 내부표시 (K) 는 11pt"""
    if not client:
        return None
    if not mark:
        return client
    return CellRichText(_tb(client), _tb(f"        {mark}", sz=11))


def _color_cell(text):
    """색상 열: '원코드' 초록, '셔터' 빨강"""
    for w, c in (("원코드", GREEN), ("셔터", RED)):
        if text and w in text:
            return _rich(text, w, c)
    return text


def _note_cell(text):
    """기재사항: '틀안' 파랑"""
    return _rich(text, "틀안", BLUE) if text and "틀안" in text else text


def _special_cell(text):
    """특이 칸: '틀안'은 파랑, 연창 #는 일반 색상."""
    return _rich(text, "틀안", BLUE) if text and "틀안" in text else text


def _take_tlean(*values):
    """기재사항에서 '틀안'을 빼고 (틀안 여부, 정리된 값들)을 반환."""
    found = False
    cleaned = []
    for value in values:
        parts = []
        for part in str(value or "").split("/"):
            part = part.strip()
            if not part:
                continue
            if "틀안" in part:
                found = True
                part = re.sub(r"\s*틀안\s*", " ", part).strip()
            if part:
                parts.append(part)
        cleaned.append("/".join(parts) or None)
    return found, cleaned


# ─────────────────────────────────────────────
# 1. 주문 -> 장부 행
# ─────────────────────────────────────────────
JULBONG = re.compile(r"(?:줄|봉)\s*(\d+)")
HANDLE_FRACTIONS = {"½": "1/2", "⅓": "1/3", "¼": "1/4"}


def handle_split(value):
    """1/2·1/3·1/4 손잡이 분할의 총 창 수. 아니면 None."""
    text = str(value or "").strip()
    text = HANDLE_FRACTIONS.get(text, text).replace(" ", "")
    m = re.fullmatch(r"1/(\d+)", text)
    return int(m.group(1)) if m and int(m.group(1)) >= 2 else None


def normalized_ledger_qty(value):
    n = handle_split(value)
    return f"1/{n}" if n else value


def unify_handle_word(text):
    """기재사항에 남은 '줄140' · '봉50' 표기를 '손140' 형태로 통일"""
    if not text:
        return text
    return JULBONG.sub(lambda m: f"손{m.group(1)}", str(text))


def _short_dir(d, client=None):
    # 우측 손잡이는 DI만 표기하고 나머지 업체는 기본값으로 생략한다.
    return d if d == "좌" or (client == "DI" and d == "우") else None


def _handle(kind, height, length):
    n, _ = normalize_handle(length)
    if not n:
        return None
    # 장부의 길이 전용 칸이므로 '손170' 대신 숫자 '170'만 적는다.
    txt = None if n == default_handle_length(kind, height) else n
    return txt


def sort_di_items(items, M):
    """DI 장부 정렬 — 손잡이실 색상으로 묶고 그룹 안에서 슬랫 번호 오름차순.
       그룹 간 순서는 고정이 아니므로 첫 등장 순서를 유지한다."""
    order_of = {}
    for it in items:
        t = M.thread_color(it.get("품목코드")) or "?"
        order_of.setdefault(t, len(order_of))
    return sorted(
        items,
        key=lambda it: (order_of[M.thread_color(it.get("품목코드")) or "?"],
                        str(it.get("품목코드") or "")))


def to_rows(order, ship_label, M=None):
    """주문 1건 -> 장부 행 리스트(dict)"""
    client = order.get("거래처")
    mark = CLIENT_INFO.get(client, (None, None, None, ""))[3]
    order_no = str(order.get("주문번호") or "").strip()
    if client == "SP" and order_no:
        m = re.search(r"(?:^|-)\s*(\d+)\s*$", order_no)
        mark = m.group(1) if m else ""
    rows = []
    items = [it for it in order.get("items", []) if not it.get("예외품목")]
    if client == "DI" and M is not None:
        items = sort_di_items(items, M)
    prev_color = prev_h = None
    prev_joint = False
    common = order.get("전체기재사항") or ""
    if client == "SP" and order_no:
        common = "/".join(x for x in common.split("/")
                          if x.strip() and x.strip() != order_no)
    sp_notices = []
    if client == "SP":
        common_parts = [x.strip() for x in common.split("/") if x.strip()]
        sp_notices = [x for x in common_parts if "공지" in x]
        common = "/".join(x for x in common_parts if "공지" not in x)
    body = "/".join(x for x in common.split("/") if x and x != "포장비용")
    cust = (order.get("고객명") or "").replace("/", "-").strip()
    head = "/".join(x for x in (body, cust) if x)
    single = len(items) == 1 and client not in NO_MERGE_PLACE

    for i, it in enumerate(items):
        color = ledger_color(it.get("품목코드"), it.get("종류"), it.get("타입"),
                             it.get("_ledger_prefix", "B"))
        same_color = color == prev_color
        h = it.get("세로")
        ditto = (same_color and h == prev_h and client not in NO_DITTO)
        place = it.get("설치장소")
        note = it.get("기재사항") or ""
        # 창이 1개면 기재사항 한 칸에 '/' 로 합쳐 적는다.
        # 여러 개면 공통값은 J열(첫 행), 방 이름은 K열에 행마다 적는다.
        if client == "SP":
            # SP: 주문번호 -> 개별 창 특징 -> 주문 전체 특징
            note1 = "/".join(x for x in (
                "/".join(sp_notices) if i == 0 else None,
                order_no if i == 0 else None,
                note, place,
                body if i == 0 else None,
            ) if x) or None
            note2 = None
        elif client in ("보노", "미래가공"):
            note1 = "/".join(x for x in (note, place) if x) or None
            note2 = order_no if i == 0 else None
        elif client == "MS":
            note1 = "/".join(x for x in (place, head if i == 0 else None, note)
                             if x) or None
            note2 = None
        elif single:
            note1 = "/".join(x for x in (head, note, place) if x) or None
            note2 = None
        else:
            note1 = "/".join(x for x in (head, note) if x) if i == 0 else \
                    (note or None)
            note2 = place
        note1, note2 = unify_handle_word(note1), unify_handle_word(note2)
        has_tlean, (note1, note2) = _take_tlean(note1, note2)
        joint = bool(it.get("종류") == "원코드" and it.get("연창"))
        joint_start = joint and (not prev_joint or not same_color)
        special = " ".join(x for x in (
            "틀안" if has_tlean else None,
            "#" if joint_start else None,
        ) if x) or None
        handle_text = _handle(it.get("종류"), h, it.get("손잡이길이"))
        if client == "미래가공" and normalize_handle(it.get("손잡이길이"))[0] == 150:
            handle_text = None
        rows.append({
            "거래처": client if i == 0 else None,
            "내부표시": mark if i == 0 else None,
            "색상": None if same_color else color,
            "가로": it.get("가로"),
            "세로": '"' if ditto else h,
            "수량": normalized_ledger_qty(it.get("수량")),
            "모형1": _short_dir(it.get("손잡이방향"), client),
            "모형2": handle_text,
            "특이": special,
            "기재사항": note1,
            "기재사항2": note2,
            "출고일": ship_label if i == 0 else None,
            "_같은색": same_color,
            "_특수": None,
        })
        prev_color, prev_h, prev_joint = color, (h if not ditto else prev_h), joint

    d = order.get("배송") or {}
    if client == "유앤":
        for accessory in order.get("_unit_accessories") or []:
            rows.append({"_특수": "부속", "문구": accessory.get("품명"),
                         "수량": f"X{accessory.get('세트', 1)}set"})
    if client == "아지트":
        rows.append({"_특수": "☆", "문구": order.get("_az_place") or "시온가공소"})
    if "포장비용" in common:
        rows.append({"_특수": "#", "문구": "포장비용"})
    if d.get("방식") in ("택배", "화물", "배달") and d.get("주소"):
        rows.append({"_특수": "☆", "문구": d["주소"]})
        who = " ".join(x for x in (d.get("수령인"), d.get("연락처")) if x)
        if who:
            rows.append({"_특수": "", "문구": who,
                         "선불": d.get("선불착불")})
        if d.get("전달사항"):
            rows.append({"_특수": "", "문구": f"전달 : {d['전달사항']}"})
        if d.get("발신"):
            rows.append({"_특수": "", "문구": f"발신 : {d['발신']}"})
    return rows


# ─────────────────────────────────────────────
# 2. 장부 워크북
# ─────────────────────────────────────────────
HEAD = {1: "상      호", 2: "색       상", 3: "규     격", 6: "수량",
        7: "모   형", 9: "기재사항", 11: "출고일"}
DETAIL_HEAD = {1: "상      호", 2: "색       상", 3: "규     격", 6: "수량",
               7: "방향", 8: "길이", 9: "특이", 10: "기재사항", 12: "출고일"}


def _init_sheet(ws, total_rows, detailed=False):
    widths = DETAIL_COL_W if detailed else COL_W
    for c, width in enumerate(widths):
        ws.column_dimensions[get_column_letter(c + 1)].width = width
    for r in range(1, total_rows + 1):
        ws.row_dimensions[r].height = ROW_H


def _block(ws, top, nrows, header_date=None, detailed=False):
    """top 행에 날짜, top+1 에 헤더, 그 아래 nrows 개 데이터 영역을 만든다.
       데이터 시작 행을 반환."""
    last_col = 13 if detailed else 12
    head = DETAIL_HEAD if detailed else HEAD
    aligns = DETAIL_ALIGN if detailed else ALIGN
    no_left = {4, 5, 11} if detailed else NO_LEFT
    no_right = {3, 4, 10} if detailed else NO_RIGHT
    ws.merge_cells(start_row=top, end_row=top, start_column=8, end_column=last_col)
    d = header_date or date.today()
    c = ws.cell(top, 8, f"{d.month} 월   {d.day} 일        번째")
    c.font = Font(name=FONT_UI, size=11, bold=True)
    c.alignment = Alignment(horizontal="right", vertical="center")

    hr = top + 1
    for i in range(last_col):
        cell = ws.cell(hr, i + 1, head.get(i))
        cell.font = Font(name=FONT_UI, size=8 if i == last_col - 1 else 11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(top=THIN, bottom=HAIR,
                             left=None if i in no_left else (THIN if i == 0 else HAIR),
                             right=None if i in no_right else (THIN if i == last_col - 1 else HAIR))
    merges = ((4, 6), (11, 12)) if detailed else ((4, 6), (8, 9), (10, 11))
    for a, b in merges:
        ws.merge_cells(start_row=hr, end_row=hr, start_column=a, end_column=b)

    first, last = hr + 1, hr + nrows
    for r in range(first, last + 1):
        for i in range(last_col):
            cell = ws.cell(r, i + 1)
            cell.font = Font(name=FONT_UI if i == 4 else FONT_DATA,
                             size=11 if i == 4 else (18 if i == 0 else 14))
            h_align = aligns[i]
            cell.alignment = Alignment(horizontal=h_align, vertical="center",
                                       shrink_to_fit=True)
            left = None if i in no_left else (THIN if i == 0 else HAIR)
            right = None if i in no_right else (THIN if i == last_col - 1 else HAIR)
            cell.border = Border(top=THIN if r == first else HAIR,
                                 bottom=HAIR, left=left, right=right)
            if r == last:
                cell.border = Border(top=HAIR, bottom=THIN,
                                     left=cell.border.left, right=cell.border.right)
        ws.cell(r, 5, "X")
    return first


def _new_sheet(wb, title, nrows, header_date=None, detailed=False):
    ws = wb.create_sheet(title) if wb.sheetnames != ["Sheet"] else wb.active
    ws.title = title
    _init_sheet(ws, nrows + 2, detailed)
    _block(ws, 1, nrows, header_date, detailed)
    return ws


def _write_rows(ws, rows, start=3, center_head=False, detailed=False):
    """rows 를 start 행부터 기록. 모형_2·기재사항은 조건에 맞으면 세로 병합"""
    r = start
    span = []
    for row in rows:
        if row.get("_특수") is not None:
            if row.get("_특수") == "부속":
                ws.cell(r, 3, row.get("문구"))
                ws.cell(r, 3).alignment = Alignment(horizontal="left",
                                                     vertical="center",
                                                     shrink_to_fit=True)
                ws.cell(r, 7, row.get("수량"))
                ws.cell(r, 7).alignment = Alignment(horizontal="center",
                                                     vertical="center",
                                                     shrink_to_fit=True)
                r += 1
                continue
            ws.cell(r, 3, row["_특수"])
            ws.cell(r, 3).alignment = Alignment(horizontal="right", vertical="center")
            ws.cell(r, 4, row.get("문구"))
            ws.cell(r, 4).alignment = Alignment(horizontal="left", vertical="center",
                                                shrink_to_fit=True)
            if row.get("선불"):
                pay_col = 12 if detailed else 11
                ws.merge_cells(start_row=r, end_row=r, start_column=4,
                               end_column=pay_col - 1)
                ws.cell(r, pay_col, row["선불"]).font = Font(
                    name=FONT_DATA, size=14, color=RED)
                ws.cell(r, pay_col).alignment = Alignment(horizontal="center",
                                                          vertical="center")
            else:
                ws.merge_cells(start_row=r, end_row=r, start_column=4,
                               end_column=12 if detailed else 11)
        else:
            c2 = ws.cell(r, 2, _client_cell(row["거래처"], row["내부표시"]))
            if center_head:
                c2.alignment = Alignment(horizontal="center",
                                         vertical="center", shrink_to_fit=True)
            c3 = ws.cell(r, 3, _color_cell(row["색상"]))
            if center_head:
                c3.alignment = Alignment(horizontal="center",
                                         vertical="center", shrink_to_fit=True)
            ws.cell(r, 4, row["가로"])
            ws.cell(r, 6, row["세로"])
            ws.cell(r, 7, row["수량"])
            ws.cell(r, 8, row["모형1"])
            ws.cell(r, 9, row["모형2"])
            if detailed:
                ws.cell(r, 10, _special_cell(row.get("특이")))
                # 연창 표시 #을 포함한 특이 칸은 가운데 정렬한다.
                ws.cell(r, 10).alignment = Alignment(horizontal="center",
                                                      vertical="center",
                                                      shrink_to_fit=True)
                ws.cell(r, 11, _note_cell(row["기재사항"]))
                ws.cell(r, 12, row["기재사항2"])
                note_col, note2_col, date_col = 11, 12, 13
            else:
                ws.cell(r, 10, _note_cell(row["기재사항"]))
                ws.cell(r, 11, row["기재사항2"])
                note_col, note2_col, date_col = 10, 11, 12
            if row["기재사항"] and not row["기재사항2"]:
                span.append(r)          # 방 이름이 없으면 J:K 를 합친다
            ws.cell(r, date_col, row["출고일"])
        r += 1

    for rr in span:                 # 기재사항 가로 병합 (J:K)
        ws.merge_cells(start_row=rr, end_row=rr,
                       start_column=11 if detailed else 10,
                       end_column=12 if detailed else 11)

    # 출고일은 같은 주문의 제품 행끼리 세로로 병합한다.
    # 페이지가 나뉘면 각 페이지 안의 구간만 병합된다.
    dated = [(start + i, row) for i, row in enumerate(rows)
             if row.get("_oi") is not None]
    i = 0
    while i < len(dated):
        r0, row0 = dated[i]
        order_id = row0.get("_oi")
        j = i + 1
        while j < len(dated) and dated[j][1].get("_oi") == order_id:
            j += 1
        if j - i > 1 and row0.get("출고일"):
            ws.merge_cells(start_row=r0, end_row=dated[j - 1][0],
                           start_column=13 if detailed else 12,
                           end_column=13 if detailed else 12)
            ws.cell(r0, 13 if detailed else 12).alignment = Alignment(horizontal="center",
                                                   vertical="center")
        i = j

    # 세로 병합: 색상이 이어지는 구간에서 같은 값이 반복되면 묶는다
    prod = [(start + i, row) for i, row in enumerate(rows)
            if row.get("_특수") is None]
    spanset = set(span)
    merge_specs = ((9, "모형2"), (11, "기재사항")) if detailed \
                  else ((9, "모형2"), (10, "기재사항"))
    for col, key in merge_specs:
        i = 0
        while i < len(prod):
            r0, row0 = prod[i]
            j = i + 1
            while (j < len(prod) and prod[j][1].get("_같은색")
                   and prod[j][1].get(key) in (None, "", row0.get(key))):
                j += 1
            if col == (11 if detailed else 10) and any(rr in spanset for rr in
                                 (p[0] for p in prod[i:j])):
                i = j                # 가로 병합한 행은 세로 병합에서 제외
                continue
            if j - i > 1 and row0.get(key):
                ws.merge_cells(start_row=r0, end_row=prod[j - 1][0],
                               start_column=col, end_column=col)
            i = j
    return r


def build_ledger(all_rows, path, header_date=None):
    wb = openpyxl.Workbook()
    pages = [all_rows[i:i + LEDGER_ROWS]
             for i in range(0, max(len(all_rows), 1), LEDGER_ROWS)] or [[]]
    for n, chunk in enumerate(pages, 1):
        ws = _new_sheet(wb, "장부" if n == 1 else f"장부 ({n})",
                        LEDGER_ROWS, header_date, detailed=True)
        _write_rows(ws, chunk, detailed=True)
    wb.save(path)
    return path


# ─────────────────────────────────────────────
# 3. 장부 -> 작업지시서
# ─────────────────────────────────────────────
INNER = re.compile(r"\s*\((?:[A-Z]|\d+)\)\s*$")


def to_worksheet_rows(rows):
    """내부표시 제거, 색상 앞 'B ' 제거, 특수 행 제외"""
    out = []
    for row in rows:
        if row.get("_특수") is not None:
            continue
        r = dict(row)
        r["내부표시"] = None                      # (K), (N), (1) 등 삭제
        if r.get("색상"):
            r["색상"] = re.sub(r"^\s*B\s+", "", r["색상"]).strip()
        if "#" in str(r.get("특이") or ""):
            r["모형2"] = f"#{r.get('모형2') or ''}"
        if "틀안" in str(r.get("특이") or ""):
            r["기재사항"] = "/".join(x for x in ("틀안", r.get("기재사항")) if x)
        out.append(r)
    return out


def _group_len(rows):
    """거래처가 채워진 지점을 기준으로 업체별 묶음 길이"""
    idx = [i for i, r in enumerate(rows) if r.get("거래처")] + [len(rows)]
    return [(idx[i], idx[i + 1] - idx[i]) for i in range(len(idx) - 1)]


def build_worksheet(rows, path, header_date=None):
    """업체 묶음을 자르지 않고 SHEET_ROWS 씩 분할.
       한 시트 = 제작용 블록 + 확인용 블록 (내용 동일, 각각 날짜·헤더 포함)"""
    ws_rows = to_worksheet_rows(rows)
    pages, cur = [], []
    for start, n in _group_len(ws_rows):
        block = ws_rows[start:start + n]
        if cur and len(cur) + n > SHEET_ROWS:
            pages.append(cur)
            cur = []
        cur.extend(block)
    if cur:
        pages.append(cur)

    wb = openpyxl.Workbook()
    for i, chunk in enumerate(pages or [[]], 1):
        ws = wb.create_sheet("현장" if i == 1 else f"현장 ({i})") \
             if wb.sheetnames != ["Sheet"] else wb.active
        ws.title = "현장" if i == 1 else f"현장 ({i})"
        total = (SHEET_ROWS + 2) * 2
        _init_sheet(ws, total)
        f1 = _block(ws, 1, SHEET_ROWS, header_date)
        _write_rows(ws, chunk, start=f1, center_head=True)
        f2 = _block(ws, SHEET_ROWS + 3, SHEET_ROWS, header_date)
        _write_rows(ws, chunk, start=f2, center_head=True)      # 확인용
    wb.save(path)
    return path


# ─────────────────────────────────────────────
# 4. 경영박사 전표
# ─────────────────────────────────────────────
ERP_HEAD = ["날짜", "전표번호", "계정코드", "계정", "거래처관리코드", "상호",
            "대체_코드", "대체_상호", "품목관리코드", "품명", "규격", "수량",
            "단가", "금액", "부가세", "전표적요", "사원코드", "사원"]


def build_erp(orders, M, path, order_date=None):
    """경영박사 EDI용 Excel 97-2003 파일을 만든다.

    EDI는 첫 행부터 자료로 읽으므로 헤더를 쓰지 않는다. 품목관리코드는
    실제 코드를 확정하기 전까지 공란으로 둔다.
    """
    try:
        import xlwt
    except ImportError as e:
        raise RuntimeError(
            "EDI .xls 생성 모듈이 없습니다. requirements.txt의 xlwt를 설치해 주세요."
        ) from e

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("전표")
    out_row = 0
    for voucher_no, o in enumerate(orders, 1):
        d = o.get("_ship_date") or order_date or date.today()
        if isinstance(d, str):
            try:
                d = date.fromisoformat(d[:10])
            except ValueError:
                d = order_date or date.today()
        dstr = f"{d.year}.{d.month:02d}.{d.day:02d}"
        client = o.get("거래처")
        erp_client = ERP_CLIENT_NAME.get(client, client)
        code, _, _, _, _ = CLIENT_INFO.get(client, ("", "", "", "", ""))
        valid_count = 0
        for it in o.get("items", []):
            if it.get("예외품목"):
                continue
            hit = M.find_item(it.get("품목코드"), it.get("종류"),
                              it.get("타입"), client)
            if not hit:
                continue
            w, h = it.get("가로"), it.get("세로")
            if not (w and h):
                continue
            calc = calc_erp(w, h, hit["단가"])
            memo = f"{float(w):.1f}*{float(h):.1f}/1EA"
            extra = "/".join(x for x in (it.get("기재사항"),
                                         it.get("설치장소")) if x)
            if extra:
                memo += f" {extra}"
            split_n = handle_split(it.get("수량"))
            if not split_n:
                erp_parts = [(it.get("손잡이방향") or "우",
                              calc["수량"], calc["금액"], calc["부가세"])]
            else:
                # 분수 손잡이는 경영박사에 창별로 나눈다.
                # 1/2=좌우, 1/3=좌우우, 1/4=좌좌우우. 수량의 0.01
                # 나머지는 앞의 좌측 행부터 배분해 합계를 맞춘다.
                left_n = split_n // 2
                directions = ["좌"] * left_n + ["우"] * (split_n - left_n)
                qty_cents = int((Decimal(calc["수량"]) * 100).to_integral_value())
                q_base, q_rem = divmod(qty_cents, split_n)
                amount_total, vat_total = int(calc["금액"]), int(calc["부가세"])
                a_base, a_rem = divmod(amount_total, split_n)
                v_base, v_rem = divmod(vat_total, split_n)
                erp_parts = []
                for part_i, direction in enumerate(directions):
                    qty = Decimal(q_base + (1 if part_i < q_rem else 0)) / 100
                    amount = a_base + (1 if part_i < a_rem else 0)
                    vat = v_base + (1 if part_i < v_rem else 0)
                    erp_parts.append((direction, qty, amount, vat))
            for dir_, qty, amount, vat in erp_parts:
                values = [dstr, voucher_no, 3, "외출", code, erp_client, "", "", "",
                          hit["품명"], hit["규격"], float(qty), int(calc["단가"]),
                          int(amount), int(vat), memo,
                          1 if dir_ == "좌" else 2, dir_]
                for col, value in enumerate(values):
                    ws.write(out_row, col, value)
                out_row += 1
            valid_count += 1

        # 포장비도 해당 주문과 같은 전표번호로 기록한다.
        if client in PACKING_CLIENTS and valid_count:
            amount = PACKING_ITEM["단가"] * valid_count
            vat = int((Decimal(amount) / 10).quantize(
                Decimal("1"), rounding=ROUND_HALF_UP))
            values = [dstr, voucher_no, 3, "외출", code, erp_client, "", "", "",
                      PACKING_ITEM["품명"], PACKING_ITEM["규격"], float(valid_count),
                      PACKING_ITEM["단가"], amount, vat, "", "", ""]
            for col, value in enumerate(values):
                ws.write(out_row, col, value)
            out_row += 1

    wb.save(path)
    return path


# ─────────────────────────────────────────────
def build_all(orders, ship_label, master_path, out_dir="."):
    from pathlib import Path
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    M = Master(master_path)
    rows = []
    for o in orders:
        rows.extend(to_rows(o, ship_label))
    return {
        "장부": build_ledger(rows, f"{out_dir}/장부.xlsx"),
        "작업지시서": build_worksheet(rows, f"{out_dir}/작업지시서.xlsx"),
        "경영박사": build_erp(orders, M, f"{out_dir}/경영박사_EDI.xls"),
    }


def _ledger_client(value):
    """장부 상호 셀에서 내부표시를 제외한 FitOrder 거래처를 찾는다."""
    text = str(value or "").strip()
    aliases = {"두창": "DU", "두창블라인드": "DU",
               "대일": "DI", "대일산업": "DI", "루임트": "RT",
               "스페이스": "SP", "유앤아이티엔에스": "유앤",
               "윈도우투모로우": "WT", "트루갤러리": "인천)트루",
               "미성텍스": "MS", "이끌림": "보노"}
    for name in sorted(CLIENT_INFO, key=len, reverse=True):
        if re.match(rf"^{re.escape(name)}(?:\s|\(|$)", text):
            return name
    for name, code in aliases.items():
        if text.startswith(name):
            return code
    return None


def _ledger_number(value):
    if value in (None, ""):
        return None
    try:
        number = float(str(value).replace(",", "").strip())
        return int(number) if number.is_integer() else number
    except (TypeError, ValueError):
        return None


def read_ledger(source, M=None, filename=None):
    """FitOrder 장부.xlsx를 작업지시서/EDI 생성용 데이터로 역변환한다.

    반환값은 (장부행, 주문, 오류문구)이다. 장부의 여러 시트도 순서대로 읽는다.
    """
    suffix = Path(filename or getattr(source, "name", "")).suffix.lower()
    if suffix == ".xls":
        try:
            import xlrd
        except ImportError as e:
            raise RuntimeError("기존 .xls 장부 처리 모듈이 없습니다.") from e
        contents = source.getvalue() if hasattr(source, "getvalue") else None
        book = xlrd.open_workbook(filename=None if contents is not None else str(source),
                                  file_contents=contents)
        raw_sheets = [ws for ws in book.sheets()
                      if ws.name == "장부" or re.fullmatch(r"장부 \(\d+\)", ws.name)]

        def sheet_info(ws):
            def get_cell(row, col):
                if row < 1 or col < 1 or row > ws.nrows or col > ws.ncols:
                    return None
                return ws.cell_value(row - 1, col - 1)
            return ws.name, ws.nrows, get_cell
        sheets = [sheet_info(ws) for ws in raw_sheets]
    else:
        wb = openpyxl.load_workbook(source, data_only=True)
        raw_sheets = [ws for ws in wb.worksheets
                      if ws.title == "장부" or re.fullmatch(r"장부 \(\d+\)", ws.title)]

        def sheet_info(ws):
            return ws.title, ws.max_row, lambda row, col: ws.cell(row, col).value
        sheets = [sheet_info(ws) for ws in raw_sheets]
    if not sheets:
        raise ValueError("FitOrder 장부 시트를 찾지 못했습니다.")

    rows, orders, errors = [], [], []
    current = None
    prev_color = prev_height = None
    for sheet_name, max_row, get_cell in sheets:
        header = [str(get_cell(2, c) or "").strip() for c in range(1, 14)]
        detailed = "특이" in header
        for excel_row in range(3, max_row + 1):
            values = [get_cell(excel_row, c) for c in range(1, 14)]
            client_cell, color_cell = values[1], values[2]
            if str(client_cell or "").replace(" ", "") in ("상호", "현장용"):
                current = None
                prev_color = prev_height = None
                continue
            width, height = values[3], values[5]
            qty, direction, length = values[6], values[7], values[8]
            if detailed:
                special, note1, note2 = values[9], values[10], values[11]
            else:
                # 기존 공장 장부는 I열 하나에 길이와 연창 #을 함께 적는다.
                special = "#" if "#" in str(length or "") else None
                note1, note2 = values[9], values[10]
            # 빈 장부 행에도 E열의 'X'가 미리 들어 있으므로 E열은 제외한다.
            meaningful = [values[i] for i in (1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12)]
            if not any(v not in (None, "") for v in meaningful):
                continue

            # 규격이 없는 #/☆ 행은 장부 특수 행으로 보존하되 EDI 품목에서는 제외한다.
            if height in (None, "") and str(color_cell or "").strip() in ("#", "☆", "") \
                    and isinstance(width, str):
                marker = str(color_cell or "").strip()
                rows.append({"_특수": marker, "문구": width})
                continue

            if client_cell not in (None, ""):
                client = _ledger_client(client_cell)
                if not client:
                    errors.append(f"{sheet_name} {excel_row}행: 거래처 '{client_cell}'를 찾을 수 없습니다.")
                    current = None
                    continue
                current = {"거래처": client, "주문번호": None, "고객명": None,
                           "전체기재사항": None, "배송": {}, "items": []}
                orders.append(current)
                prev_color = prev_height = None
            if current is None:
                errors.append(f"{sheet_name} {excel_row}행: 주문의 첫 행에 상호가 없습니다.")
                continue

            color_text = str(color_cell or prev_color or "").strip()
            if color_cell not in (None, ""):
                prev_color = color_cell
            if str(height or "").strip() in ('"', '”'):
                height = prev_height
            elif height not in (None, ""):
                prev_height = height
            code, kind, type_ = parse_ledger_color(color_text)
            w, h = _ledger_number(width), _ledger_number(height)
            if not code:
                errors.append(f"{sheet_name} {excel_row}행: 품목 '{color_text}'를 해석할 수 없습니다.")
            if w is None or h is None:
                errors.append(f"{sheet_name} {excel_row}행: 가로·세로 규격을 확인해 주세요.")
            if code and M is not None and not M.find_item(code, kind, type_, current["거래처"]):
                errors.append(f"{sheet_name} {excel_row}행: 마스터에 없는 품목 {color_text}")

            item = {"품목코드": code, "종류": kind, "타입": type_,
                    "가로": w, "세로": h, "수량": qty or 1,
                    "손잡이방향": str(direction).strip() if direction in ("좌", "우") else None,
                    "손잡이길이": parse_handle(length)[0],
                    "연창": "#" in str(special or ""),
                    "설치장소": note2, "기재사항": note1,
                    "예외품목": False}
            current["items"].append(item)
            rows.append({"거래처": current["거래처"] if len(current["items"]) == 1 else None,
                         "내부표시": None, "색상": color_text,
                         "가로": w, "세로": h, "수량": qty,
                         "모형1": item["손잡이방향"], "모형2": length,
                         "특이": special, "기재사항": note1,
                         "기재사항2": note2, "출고일": None,
                         "_같은색": color_cell in (None, ""), "_특수": None})
    return rows, orders, errors


# ─────────────────────────────────────────────
# 편집 반영: 장부 표시값 -> 주문 항목
# ─────────────────────────────────────────────
def parse_ledger_color(text):
    """'B L18-원코드 200' -> ('200', '원코드', 'L자')"""
    s = (text or "").strip()
    if not s:
        return None, None, None
    type_ = "L자" if re.search(r"L(18|21)-", s) else "C자"
    kind = "투코드"
    for k in ("원코드", "셔터", "투코드"):
        if k in s:
            kind = k
            break
    m = re.search(r"(\d{3}[A-Za-z]*)\s*$", s)
    return (m.group(1) if m else None), kind, type_


def parse_handle(text):
    """'#손140' -> (140, True) / '손120' -> (120, False)"""
    s = str(text or "").strip()
    yeon = s.startswith("#")
    m = re.search(r"(\d+)", s)
    return (int(m.group(1)) if m else None), yeon


def apply_edit(order, item_idx, field, value, color_prefix="B"):
    """장부 셀 편집값을 주문 항목에 되돌려 적용"""
    it = order["items"][item_idx]
    v = None if value in (None, "", "None") else value
    if field == "상호":
        raw = str(v or "").strip()
        code = raw.split()[0] if raw else None
        aliases = {"두창": "DU", "두창블라인드": "DU",
                   "대일": "DI", "대일산업": "DI", "루임트": "RT"}
        code = aliases.get(code, code)
        if code in CLIENT_INFO:
            order["거래처"] = code
    elif field == "색상":
        code, kind, type_ = parse_ledger_color(v)
        if code:
            it["품목코드"], it["종류"], it["타입"] = code, kind, type_
            # 사용자가 H/R/C를 직접 적었으면 그 값을 우선하고,
            # 숫자만 적었으면 화면 위의 품목 기본 표시를 적용한다.
            m = re.match(r"\s*([BHRC])(?:\s|$)", str(v or ""), re.I)
            prefix = m.group(1).upper() if m else str(color_prefix or "B").upper()
            it["_ledger_prefix"] = prefix if prefix in {"B", "H", "R", "C"} else "B"
    elif field in ("가로", "세로"):
        if v is not None and str(v).strip() not in ('"', "”"):
            try:
                it[field] = float(str(v).replace(",", ""))
            except ValueError:
                pass
    elif field == "수량":
        it["수량"] = v
    elif field in ("모형1", "방향"):
        if str(v or "").strip() == "ㅈ":
            v = "좌"
        it["손잡이방향"] = v if v in ("좌", "우") else None
    elif field in ("모형2", "길이"):
        n, yeon = parse_handle(v)
        it["손잡이길이"] = n
        if yeon:
            it["연창"] = True
    elif field == "특이":
        text = str(v or "")
        it["연창"] = "#" in text
        old = str(it.get("기재사항") or "")
        parts = [x.strip() for x in old.split("/")
                 if x.strip() and x.strip() != "틀안"]
        if "틀안" in text:
            parts.insert(0, "틀안")
        it["기재사항"] = "/".join(parts) or None
    elif field == "기재사항":
        it["기재사항"] = v
    elif field == "기재사항2":
        it["설치장소"] = v
