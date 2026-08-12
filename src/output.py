"""
FitOrder 출력 생성기 — 장부 / 작업지시서 / 경영박사

    from output import build_all
    build_all(orders, ship_date="목", out_dir="../out")

orders = extract_order() 또는 parsers 결과 리스트
"""
import re
from copy import copy
from datetime import date

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from rules import (CLIENT_INFO, NO_DITTO, NO_MERGE_PLACE,
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


# ─────────────────────────────────────────────
# 1. 주문 -> 장부 행
# ─────────────────────────────────────────────
JULBONG = re.compile(r"(?:줄|봉)\s*(\d+)")


def unify_handle_word(text):
    """기재사항에 남은 '줄140' · '봉50' 표기를 '손140' 형태로 통일"""
    if not text:
        return text
    return JULBONG.sub(lambda m: f"손{m.group(1)}", str(text))


def _short_dir(d):
    return d if d == "좌" else None          # 우는 적지 않음


def _handle(kind, height, length, yeon):
    n, _ = normalize_handle(length)
    if not n:
        return "#" if yeon else None
    txt = None if n == default_handle_length(kind, height) else f"손{n}"
    if yeon:
        return f"#{txt}" if txt else "#"
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
    rows = []
    items = [it for it in order.get("items", []) if not it.get("예외품목")]
    if client == "DI" and M is not None:
        items = sort_di_items(items, M)
    prev_color = prev_h = None
    common = order.get("전체기재사항") or ""
    body = "/".join(x for x in common.split("/") if x and x != "포장비용")
    cust = (order.get("고객명") or "").replace("/", "-").strip()
    head = "/".join(x for x in (body, cust) if x)
    single = len(items) == 1 and client not in NO_MERGE_PLACE

    for i, it in enumerate(items):
        color = ledger_color(it.get("품목코드"), it.get("종류"), it.get("타입"))
        same_color = color == prev_color
        h = it.get("세로")
        ditto = (same_color and h == prev_h and client not in NO_DITTO)
        place = it.get("설치장소")
        note = it.get("기재사항") or ""
        # 창이 1개면 기재사항 한 칸에 '/' 로 합쳐 적는다.
        # 여러 개면 공통값은 J열(첫 행), 방 이름은 K열에 행마다 적는다.
        if single:
            note1 = "/".join(x for x in (head, note, place) if x) or None
            note2 = None
        else:
            note1 = "/".join(x for x in (head, note) if x) if i == 0 else \
                    (note or None)
            note2 = place
        note1, note2 = unify_handle_word(note1), unify_handle_word(note2)
        rows.append({
            "거래처": client if i == 0 else None,
            "내부표시": mark if i == 0 else None,
            "색상": None if same_color else color,
            "가로": it.get("가로"),
            "세로": '"' if ditto else h,
            "수량": it.get("수량"),
            "모형1": _short_dir(it.get("손잡이방향")),
            "모형2": _handle(it.get("종류"), h, it.get("손잡이길이"),
                             it.get("종류") == "원코드" and it.get("연창")),
            "기재사항": note1,
            "기재사항2": note2,
            "출고일": ship_label if i == 0 else None,
            "_같은색": same_color,
            "_특수": None,
        })
        prev_color, prev_h = color, (h if not ditto else prev_h)

    d = order.get("배송") or {}
    if "포장비용" in common:
        rows.append({"_특수": "#", "문구": "포장비용"})
    if d.get("방식") in ("택배", "화물") and d.get("주소"):
        rows.append({"_특수": "☆", "문구": d["주소"]})
        who = " ".join(x for x in (d.get("수령인"), d.get("연락처")) if x)
        rows.append({"_특수": "", "문구": who, "선불": d.get("선불착불")})
        if d.get("전달사항"):
            rows.append({"_특수": "", "문구": f"전달 : {d['전달사항']}"})
    return rows


# ─────────────────────────────────────────────
# 2. 장부 워크북
# ─────────────────────────────────────────────
HEAD = {1: "상      호", 2: "색       상", 3: "규     격", 6: "수량",
        7: "모   형", 9: "기재사항", 11: "출고일"}


def _init_sheet(ws, total_rows):
    for c in range(12):
        ws.column_dimensions[get_column_letter(c + 1)].width = COL_W[c]
    for r in range(1, total_rows + 1):
        ws.row_dimensions[r].height = ROW_H


def _block(ws, top, nrows, header_date=None):
    """top 행에 날짜, top+1 에 헤더, 그 아래 nrows 개 데이터 영역을 만든다.
       데이터 시작 행을 반환."""
    ws.merge_cells(start_row=top, end_row=top, start_column=8, end_column=12)
    d = header_date or date.today()
    c = ws.cell(top, 8, f"{d.month} 월   {d.day} 일        번째")
    c.font = Font(name=FONT_UI, size=11, bold=True)
    c.alignment = Alignment(horizontal="right", vertical="center")

    hr = top + 1
    for i in range(12):
        cell = ws.cell(hr, i + 1, HEAD.get(i))
        cell.font = Font(name=FONT_UI, size=8 if i == 11 else 11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(top=THIN, bottom=HAIR,
                             left=None if i in NO_LEFT else (THIN if i == 0 else HAIR),
                             right=None if i in NO_RIGHT else (THIN if i == 11 else HAIR))
    for a, b in ((4, 6), (8, 9), (10, 11)):
        ws.merge_cells(start_row=hr, end_row=hr, start_column=a, end_column=b)

    first, last = hr + 1, hr + nrows
    for r in range(first, last + 1):
        for i in range(12):
            cell = ws.cell(r, i + 1)
            cell.font = Font(name=FONT_UI if i == 4 else FONT_DATA,
                             size=11 if i == 4 else (18 if i == 0 else 14))
            cell.alignment = _al(i)
            cell.border = _border(2 if r == first else 3, i, 99,
                                  top=THIN if r == first else HAIR)
            if r == last:
                cell.border = Border(top=HAIR, bottom=THIN,
                                     left=cell.border.left, right=cell.border.right)
        ws.cell(r, 5, "X")
    return first


def _new_sheet(wb, title, nrows, header_date=None):
    ws = wb.create_sheet(title) if wb.sheetnames != ["Sheet"] else wb.active
    ws.title = title
    _init_sheet(ws, nrows + 2)
    _block(ws, 1, nrows, header_date)
    return ws


def _write_rows(ws, rows, start=3, center_head=False):
    """rows 를 start 행부터 기록. 모형_2·기재사항은 조건에 맞으면 세로 병합"""
    r = start
    span = []
    for row in rows:
        if row.get("_특수") is not None:
            ws.cell(r, 3, row["_특수"])
            ws.cell(r, 4, row.get("문구"))
            ws.merge_cells(start_row=r, end_row=r, start_column=4, end_column=11)
            if row.get("선불"):
                ws.cell(r, 12, row["선불"]).font = Font(
                    name=FONT_DATA, size=14, color=RED)
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
            ws.cell(r, 10, _note_cell(row["기재사항"]))
            ws.cell(r, 11, row["기재사항2"])
            if row["기재사항"] and not row["기재사항2"]:
                span.append(r)          # 방 이름이 없으면 J:K 를 합친다
            ws.cell(r, 12, row["출고일"])
        r += 1

    for rr in span:                 # 기재사항 가로 병합 (J:K)
        ws.merge_cells(start_row=rr, end_row=rr, start_column=10, end_column=11)

    # 세로 병합: 색상이 이어지는 구간에서 같은 값이 반복되면 묶는다
    prod = [(start + i, row) for i, row in enumerate(rows)
            if row.get("_특수") is None]
    spanset = set(span)
    for col, key in ((9, "모형2"), (10, "기재사항")):
        i = 0
        while i < len(prod):
            r0, row0 = prod[i]
            j = i + 1
            while (j < len(prod) and prod[j][1].get("_같은색")
                   and prod[j][1].get(key) in (None, "", row0.get(key))):
                j += 1
            if col == 10 and any(rr in spanset for rr in
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
                        LEDGER_ROWS, header_date)
        _write_rows(ws, chunk)
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
ERP_HEAD = ["날짜", "계정코드", "계정", "거래처관리코드", "상호", "대체_코드",
            "대체_상호", "품명", "규격", "수량", "단가", "금액", "부가세",
            "전표적요", "사원코드", "사원"]


def build_erp(orders, M, path, order_date=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "전표"
    ws.append(ERP_HEAD)
    for c in range(1, len(ERP_HEAD) + 1):
        ws.cell(1, c).font = Font(name=FONT_UI, bold=True)
        ws.column_dimensions[get_column_letter(c)].width = 14
    d = order_date or date.today()
    dstr = f"{d.year}.{d.month:02d}.{d.day:02d}."
    for o in orders:
        client = o.get("거래처")
        code, _, _, _, _ = CLIENT_INFO.get(client, ("", "", "", "", ""))
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
            dir_ = it.get("손잡이방향") or "우"
            memo = f"{float(w):.1f}*{float(h):.1f}/1EA"
            extra = "/".join(x for x in (it.get("기재사항"),
                                         it.get("설치장소")) if x)
            if extra:
                memo += f" {extra}"
            ws.append([dstr, 3, "외출", code, client, None, None,
                       hit["품명"], hit["규격"], float(calc["수량"]),
                       int(calc["단가"]), int(calc["금액"]),
                       int(calc["부가세"]), memo,
                       1 if dir_ == "좌" else 2, dir_])
    # 포장비용 — 거래처별 창 수만큼, 전표 맨 아래에
    from decimal import Decimal, ROUND_HALF_UP
    pack = {}
    for o in orders:
        cl = o.get("거래처")
        if cl not in PACKING_CLIENTS:
            continue
        n = sum(1 for it in o.get("items", [])
                if not it.get("예외품목") and it.get("가로") and it.get("세로"))
        if n:
            pack[cl] = pack.get(cl, 0) + n
    for cl, n in pack.items():
        code = CLIENT_INFO.get(cl, ("",))[0]
        amt = PACKING_ITEM["단가"] * n
        vat = int((Decimal(amt) / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        ws.append([dstr, 3, "외출", code, cl, None, None,
                   PACKING_ITEM["품명"], PACKING_ITEM["규격"], float(n),
                   PACKING_ITEM["단가"], amt, vat, None, None, None])

    for r in range(2, ws.max_row + 1):
        for c in (10, 11, 12, 13):
            ws.cell(r, c).number_format = "#,##0.00" if c == 10 else "#,##0"
        for c in range(1, len(ERP_HEAD) + 1):
            ws.cell(r, c).font = Font(name=FONT_UI, size=10)
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
        "경영박사": build_erp(orders, M, f"{out_dir}/경영박사.xlsx"),
    }


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
    s = (text or "").strip()
    yeon = s.startswith("#")
    m = re.search(r"(\d+)", s)
    return (int(m.group(1)) if m else None), yeon


def apply_edit(order, item_idx, field, value):
    """장부 셀 편집값을 주문 항목에 되돌려 적용"""
    it = order["items"][item_idx]
    v = None if value in (None, "", "None") else value
    if field == "색상":
        code, kind, type_ = parse_ledger_color(v)
        if code:
            it["품목코드"], it["종류"], it["타입"] = code, kind, type_
    elif field in ("가로", "세로"):
        if v is not None and str(v).strip() not in ('"', "”"):
            try:
                it[field] = float(str(v).replace(",", ""))
            except ValueError:
                pass
    elif field == "수량":
        it["수량"] = v
    elif field == "모형1":
        it["손잡이방향"] = v if v in ("좌", "우") else None
    elif field == "모형2":
        n, yeon = parse_handle(v)
        it["손잡이길이"] = n
        if yeon:
            it["연창"] = True
    elif field == "기재사항":
        it["기재사항"] = v
    elif field == "기재사항2":
        it["설치장소"] = v
