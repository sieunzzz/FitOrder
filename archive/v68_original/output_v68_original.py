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
                   DI_CODE_ORDER, clean_delivery_notice,
                   PACKING_CLIENTS, PACKING_ITEM, BONO_PACKING_ITEM, TRUE_ACCESSORY_PRODUCTS,
                   Master, calc_erp, default_handle_length, di_mix_info,
                   ledger_color, normalize_handle)
from holding import (HOLDING_FEATURE_ENABLED, accessory_qty, holding_accessory_from_text, holding_calc,
                     holding_ledger_text, holding_magnet_charge_qty,
                     holding_product_from_text, looks_like_holding_operation,
                     normalize_holding_operation)

# ── 서식 상수 (원본 장부.xls 에서 추출) ──
FONT_DATA = "THE행복열매"
FONT_UI = "맑은 고딕"
ROW_H = 20.1
LEDGER_ROWS = 35              # 3행 ~ 37행
SHEET_ROWS = 16               # 작업지시서 블록당 데이터 행 수
DI_TARGET_ROWS = 35           # DI: 한 장의 기본 목표
DI_MAX_ROWS = 40              # DI: 같은 수령인을 유지하는 최대 행 수
DI_SHEET_GROUPS = (
    ("C 원코드", ("C자", "원코드")),
    ("C 투코드", ("C자", "투코드")),
    ("C 셔터", ("C자", "셔터")),
    ("L 원코드", ("L자", "원코드")),
    ("L 투코드", ("L자", "투코드")),
)
COL_W = [3.5, 15.1, 16.6, 6.6, 2.6, 6.6, 6.1, 6.0, 6.0, 6.1, 6.1, 8.4]
#        A         B(상호)   C(색상)  D(가로)  E(X)      F(세로)
#        G(수량)   H(모형1)  I(모형2) J(기재)  K(기재2)  L(출고일)
ALIGN = ["center", "right", "left", "right", "center", "left",
         "center", "left", "right", "center", "center", "center"]
DETAIL_COL_W = [3.5, 14.125, 16.625, 6.625, 2.625, 6.625, 4.875,
                5.125, 4.875, 5.625, 5.625, 5.625, 8.375]
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


def _true_accessory_catalog(label):
    key = str(label or "").strip()
    if key in TRUE_ACCESSORY_PRODUCTS:
        return dict(TRUE_ACCESSORY_PRODUCTS[key])
    normalized = re.sub(r"\s+", "", key).lower()
    if normalized in {"b원코드브라켓", "b원코드브라캣"}:
        return dict(TRUE_ACCESSORY_PRODUCTS["B 원코드 브라켓"])
    if normalized in {"노피스(1)", "노피스1", "노피스브라켓(1ea)-커튼박스용"}:
        return dict(TRUE_ACCESSORY_PRODUCTS["노피스(1)"])
    if normalized in {"노피스(2)", "노피스2", "노피스브라켓(2ea)"}:
        return dict(TRUE_ACCESSORY_PRODUCTS["노피스(2)"])
    return None


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
        return mark or None
    if not mark:
        return client
    return CellRichText(_tb(client), _tb(f"        {mark}", sz=11))


def _color_cell(text):
    """색상 열 부분색상.

    - 원코드: `원코드` 글자만 초록
    - 셔터: `셔터` 글자만 빨강
    - P/FP: 품명 전체가 아니라 `029FP`, `102P` 같은 실제 색상코드만 파랑

    예: `B 원코드 029FP` -> `원코드` 초록 + `029FP` 파랑.
    """
    if not text:
        return text
    raw = str(text)
    token_re = re.compile(r"원코드|셔터|\d{1,4}(?:FP|P)(?![A-Za-z0-9])", re.I)
    matches = list(token_re.finditer(raw))
    if not matches:
        return raw
    parts = []
    pos = 0
    for m in matches:
        if m.start() > pos:
            parts.append(_tb(raw[pos:m.start()]))
        token = m.group(0)
        if re.fullmatch(r"\d{1,4}(?:FP|P)", token, re.I):
            color = BLUE
        elif token == "원코드":
            color = GREEN
        else:
            color = RED
        parts.append(_tb(token, color))
        pos = m.end()
    if pos < len(raw):
        parts.append(_tb(raw[pos:]))
    return CellRichText(*parts)


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


def expand_same_size_directions(order):
    """한 사이즈에 좌/우가 섞인 경우 방향별 행으로 펼친다.

    좌개수/우개수가 모두 존재하면 방향별 한 행씩 만들고 그 방향의 창개수를 보존한다.
    예: 창개수3, 좌1/우2 -> 좌 1창 행 + 우 2창 행.
    `좌우`처럼 방향 문자열만 두 개 이상이면 각 방향 1창으로 펼친다.
    한쪽 방향만 여러 개인 경우(우2)는 한 행 + 창개수2를 유지한다.
    """
    if not order or not order.get("items"):
        return order
    expanded_items = []
    for it in order.get("items") or []:
        if it.get("_product_group") == "holding" or it.get("_holding_accessory"):
            expanded_items.append(it)
            continue
        try:
            left_n = int(float(it.get("좌개수") or 0))
            right_n = int(float(it.get("우개수") or 0))
        except (TypeError, ValueError):
            left_n = right_n = 0
        if left_n > 0 and right_n > 0:
            for direction, count in (("좌", left_n), ("우", right_n)):
                clone = dict(it)
                clone["손잡이방향"] = direction
                # 방향별 집계행은 그 방향의 실제 창 수만 보존한다.
                # 원본 수량을 그대로 두면 EDI에서 다시 곱해지는 문제가 생긴다.
                clone["수량"] = count
                clone["창개수"] = count
                clone["좌개수"] = count if direction == "좌" else 0
                clone["우개수"] = count if direction == "우" else 0
                expanded_items.append(clone)
            continue

        raw = " ".join(str(x or "") for x in (it.get("손잡이방향"), it.get("원문")))
        seqs = re.findall(r"[좌우]{2,10}", raw)
        if seqs and left_n + right_n == 0:
            directions = list(seqs[-1])
            for direction in directions:
                clone = dict(it)
                clone["손잡이방향"] = direction
                clone["수량"] = 1
                clone["창개수"] = 1
                clone["좌개수"] = 1 if direction == "좌" else 0
                clone["우개수"] = 1 if direction == "우" else 0
                expanded_items.append(clone)
            continue
        expanded_items.append(it)
    order["items"] = expanded_items
    return order



def _di_recipient_key(it):
    return str(it.get("_DIrecipient_context") or it.get("기재사항") or "").strip()

def _di_item_is_urgent(it):
    """DI 한 행이 긴급 건인지 최종 편집값까지 포함해 판정한다."""
    return "긴급" in " ".join(str(it.get(k) or "") for k in
                               ("_수동특이", "기재사항", "원문"))

def _propagate_di_urgent_order(order):
    if order.get("거래처") != "DI":
        return order
    groups = {}
    for it in order.get("items") or []:
        # Excel/GPT/수동 수정 어느 경로에서 들어왔든 `긴급 ---`처럼
        # 긴급 뒤에 구분선만 붙은 경우는 최종 출력 전에 `긴급`으로 통일한다.
        special = str(it.get("_수동특이") or "").strip()
        if special:
            special = re.sub(r"(?<=긴급)\s*[-–—_=~.·ㆍ]+\s*(?=$|/)", "", special)
            special = special.strip(" -–—_=~.·ㆍ")
            it["_수동특이"] = special or None
        groups.setdefault(_di_recipient_key(it), []).append(it)
    for _recipient, items in groups.items():
        if not any("긴급" in " ".join(str(it.get(k) or "") for k in
                   ("_수동특이", "기재사항", "원문")) for it in items):
            continue
        for it in items:
            parts = [x.strip() for x in str(it.get("_수동특이") or "").split("/") if x.strip()]
            if "긴급" not in parts:
                parts.append("긴급")
            it["_수동특이"] = "/".join(parts) or None
    return order

def synchronize_order_for_outputs(order):
    """사전점검에서 확정된 주문을 세 출력물이 동일하게 쓰도록 마지막 공통 정규화."""
    if not order:
        return order
    delivery = order.setdefault("배송", {})
    delivery["전달사항"] = clean_delivery_notice(delivery.get("전달사항"))
    _propagate_di_urgent_order(order)
    # 누락된 손잡이길이는 원문/기재에서 결정 가능한 숫자만 보완한다.
    for it in order.get("items") or []:
        if it.get("손잡이길이") in (None, ""):
            raw = " ".join(str(it.get(k) or "") for k in
                           ("손잡이방향", "원문", "기재사항", "_수동특이"))
            m = re.search(r"(?:손잡이(?:길이)?|손|줄|봉|각도봉)\s*[:=/-]?\s*(\d{2,3})", raw)
            if m:
                it["손잡이길이"] = int(m.group(1))
    return order


def normalized_ledger_qty(value):
    n = handle_split(value)
    return f"1/{n}" if n else value


def _display_size(value):
    """장부 치수 표시에서 불필요한 .0을 제거한다."""
    if value in (None, ""):
        return value
    try:
        f = float(value)
        return int(f) if f.is_integer() else f
    except (TypeError, ValueError):
        text = str(value).strip()
        return re.sub(r"(?<=\d)\.0$", "", text)


def unify_handle_word(text):
    """기재사항에 남은 '줄140' · '봉50' 표기를 '손140' 형태로 통일"""
    if not text:
        return text
    return JULBONG.sub(lambda m: f"손{m.group(1)}", str(text))


def _jo_clean_note_text(value):
    text = str(value or "").strip()
    if not text:
        return None
    text = re.sub(r"(?:손잡이(?:길이)?|손|줄|봉)\s*[:=]?\s*\d{2,3}", " ", text)
    text = re.sub(r"\s*[/|,]+\s*", "/", text)
    text = re.sub(r"/{2,}", "/", text).strip(" /,.-")
    return text or None


def _jo_note2_parts(order, item, customer=None, first=False):
    """JO 기재사항2: 주문번호와 손길이를 제외한 나머지 정보."""
    out = []
    order_no = str(order.get("주문번호") or "").strip()
    values = []
    if first:
        values.extend((customer, order.get("전체기재사항")))
    values.extend((item.get("기재사항"), item.get("설치장소")))
    for value in values:
        value = _jo_clean_note_text(value)
        for part in str(value or "").split("/"):
            part = part.strip()
            if not part:
                continue
            if order_no and re.sub(r"[()（）\s]", "", part) == re.sub(r"[()（）\s]", "", order_no):
                continue
            if part not in out:
                out.append(part)
    return out


def _blind_counts(it):
    """표형 발주의 창개수/좌/우를 정수로 정규화."""
    def n(v):
        try:
            return max(0, int(float(v))) if v not in (None, "") else 0
        except (TypeError, ValueError):
            return 0
    left_n, right_n, count = n(it.get("좌개수")), n(it.get("우개수")), n(it.get("창개수"))
    if count <= 0 and left_n + right_n > 0:
        count = left_n + right_n
    if left_n + right_n == 0 and count > 1:
        if it.get("손잡이방향") == "좌":
            left_n = count
        elif it.get("손잡이방향") == "우":
            right_n = count
    return left_n, right_n, max(1, count or 1)


def _effective_delivery_mode(order):
    """실제 배송방식이 비어 있으면 거래처 기본배송을 보조로 사용한다.

    포장비용 공통 규칙은 택배/화물일 때만 적용한다. RT처럼 판별이 필요한
    거래처나 DI처럼 기본값이 택배/화물로 확정되지 않은 거래처는 실제 판독값이
    있을 때만 포장비를 붙인다.
    """
    delivery = order.get("배송") or {}
    raw = str(delivery.get("방식") or "").strip()
    if raw:
        if "택배" in raw:
            return "택배"
        if "화물" in raw:
            return "화물"
        # 실제 배송이 배달/내사 등으로 명시됐으면 거래처 기본배송으로 덮지 않는다.
        return raw
    client = order.get("거래처")
    default = CLIENT_INFO.get(client, (None, None, None, None, None))[2]
    default = str(default or "").strip()
    if "택배" in default:
        return "택배"
    if "화물" in default:
        return "화물"
    return raw or default


def _needs_window_packing(order):
    """전 업체 공통: 배송이 택배/화물이면 창당 포장비를 부과한다."""
    return _effective_delivery_mode(order) in {"택배", "화물"}


def _prepend_note_part(note, part):
    """기재사항1 맨 앞에 값을 중복 없이 넣는다."""
    part = str(part or "").strip()
    if not part:
        return note
    parts = [x.strip() for x in str(note or "").split("/") if x.strip()]
    parts = [x for x in parts if x != part]
    return "/".join([part] + parts) or None


def _short_dir(d, client=None):
    # 장부/작업지시서는 좌측만 표기한다. 우측은 기본값이므로 숨기되
    # 원 주문 데이터에는 남겨 경영박사 EDI의 방향코드에는 그대로 사용한다.
    return "좌" if d == "좌" else None


def _has_non_di_mix(item, client=None):
    if client == "DI":
        return False
    if item.get("_mix_word"):
        return True
    raw = " ".join(str(item.get(k) or "") for k in
                   ("색상원문", "원문", "기재사항"))
    return bool(re.search(r"(?<![A-Za-z])MIX(?![A-Za-z])", raw, re.I))


def _non_di_mix_ledger_text(item, base):
    """대일 외 MIX 표기는 기존 코드/타입을 유지하면서 MIX만 추가한다."""
    if not base:
        return base
    if "MIX" in base.upper():
        return base
    kind = str(item.get("종류") or "투코드")
    text = str(base)
    if "원코드" in text:
        return text.replace("원코드", "원코드 MIX", 1)
    if "셔터" in text:
        return text.replace("셔터", "셔터 MIX", 1)
    # L 투코드는 `L18-투코드` 형태이므로 투코드 자리를 MIX로 바꾼다.
    if "투코드" in text:
        return text.replace("투코드", "MIX", 1)
    m = re.match(r"(\s*[BHRC])(?:\s+)(.*)", text, re.I)
    return f"{m.group(1)} MIX {m.group(2)}" if m else f"MIX {text}"


JL_PACKAGING = re.compile(r"(?:비닐\s*포장|겉\s*비닐|걷\s*비닐)")


def _jl_customer_name(value):
    """JL 고객명 끝에 붙는 내부 표기 DW는 출력/EDI에서 제거한다."""
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return re.sub(r"\s*DW\s*$", "", text, flags=re.I).strip()


def _jl_order_mark(value):
    """JL 장부 발주번호는 반드시 괄호를 포함해 표시한다."""
    text = str(value or "").strip()
    if not text:
        return None
    text = re.sub(r"^[（(]\s*|\s*[）)]$", "", text).strip()
    return f"({text})" if text else None


def _split_note_parts(*values):
    parts = []
    for value in values:
        for part in str(value or "").split("/"):
            part = part.strip()
            if part and part not in parts:
                parts.append(part)
    return parts


def _note_text(*values):
    """`/` 단위 중복을 제거한 기재사항 문자열."""
    parts = _split_note_parts(*values)
    return "/".join(parts) or None


def _shared_item_note_parts(items):
    """모든 제작행에 반복되는 item 기재사항을 공통 기재로 승격한다."""
    note_sets = []
    first_order = []
    for it in items:
        if it.get("_holding_accessory"):
            continue
        parts = _split_note_parts(it.get("기재사항"))
        if not note_sets:
            first_order = list(parts)
        note_sets.append(set(parts))
    if not note_sets:
        return []
    shared = set.intersection(*note_sets) if note_sets else set()
    return [p for p in first_order if p in shared]


def _total_window_count(items):
    """기재사항 분배용 실제 창 수. 수량을 다시 곱하지 않는다."""
    total = 0
    for it in items:
        if it.get("_holding_accessory"):
            continue
        if it.get("_product_group") == "holding":
            total += _holding_count(it)
        else:
            total += _blind_counts(it)[2]
    return max(1, total)


def _without_parts(parts, excluded):
    blocked = set(excluded or [])
    return [p for p in parts if p not in blocked]


def _jl_clean_parts(parts, customer=None, order_no=None, keep_piece=False):
    """JL 메모에서 포장 문구/중복 이름/DW/주문번호를 제거한다."""
    out = []
    customer = _jl_customer_name(customer)
    order_no = re.sub(r"^[（(]\s*|\s*[）)]$", "", str(order_no or "").strip())
    for part in parts:
        text = str(part or "").strip()
        if not text or JL_PACKAGING.search(text):
            continue
        bare_order = re.sub(r"^[（(]\s*|\s*[）)]$", "", text).strip()
        if order_no and bare_order == order_no:
            continue
        # '(기재: 이름 DW)' 같은 내부 DW 표기는 고객명만 남기고 중복 제거한다.
        text_name = _jl_customer_name(re.sub(r"^(?:기재\s*[:：]?)\s*", "", text))
        if customer and text_name == customer:
            continue
        if re.fullmatch(r"DW", text, re.I):
            continue
        if "피스" in text:
            if keep_piece and "피스" not in out:
                out.append("피스")
            continue
        if text not in out:
            out.append(text)
    return out


def _handle(kind, height, length):
    n, _ = normalize_handle(length)
    if not n:
        return None
    # 장부의 길이 전용 칸이므로 '손170' 대신 숫자 '170'만 적는다.
    txt = None if n == default_handle_length(kind, height) else n
    return txt


def sort_di_items(items, M):
    """DI 블라인드는 기존 규칙으로 정렬하고 홀딩도어는 원래 순서를 보존한다.

    홀딩도어의 레일연결부속/라운드부속은 바로 앞 본품과 순서 관계가 중요하므로
    색상실 기준 블라인드 정렬에 섞지 않는다.
    """
    holding = [(pos, it) for pos, it in enumerate(items)
               if it.get("_product_group") == "holding" or it.get("_holding_accessory")]
    normal = [(pos, it) for pos, it in enumerate(items)
              if not (it.get("_product_group") == "holding" or it.get("_holding_accessory"))]

    group_order = {("C자", "원코드"): 0, ("C자", "투코드"): 1,
                   ("C자", "셔터"): 2, ("L자", "원코드"): 3,
                   ("L자", "투코드"): 4}
    buckets = {}
    for pos, it in normal:
        key = group_order.get((it.get("타입"), it.get("종류")), 99)
        buckets.setdefault(key, []).append((pos, it))

    result = []
    for group in sorted(buckets):
        result.extend(_sort_di_group(buckets[group], M))
    result.extend(it for _, it in sorted(holding, key=lambda x: x[0]))
    return result

def _sort_di_group(indexed_items, M):
    """하나의 DI 타입/종류 그룹 안을 색상·수령인 단위로 정렬."""
    people, sequence = {}, []
    first_position = {}
    for pos, it in indexed_items:
        recipient = str(it.get("기재사항") or "").strip() or f"__row_{pos}"
        if recipient not in people:
            people[recipient] = []
            sequence.append(recipient)
            first_position[recipient] = pos
        people[recipient].append(it)

    thread_order = {}
    for recipient in sequence:
        for it in people[recipient]:
            thread = M.thread_color(it.get("품목코드")) or "?"
            thread_order.setdefault(thread, len(thread_order))

    def code_key(code):
        code = str(code or "").upper()
        # 믹스 조합도 첫 실제 슬랫 코드 기준으로 기존 DI 품목 순서에 배치한다.
        primary = code.split("+")[0]
        return (DI_CODE_ORDER.get(primary, 999999), code)

    def person_key(recipient):
        block = people[recipient]
        first = block[0]
        thread = M.thread_color(first.get("품목코드")) or "?"
        codes = {str(x.get("품목코드") or "") for x in block}
        # 긴급 건은 같은 DI 유형/시트 안에서 가장 먼저 나오게 한다.
        # 한 사람의 여러 창 중 하나만 긴급이어도 propagate 단계에서 전부 긴급이므로
        # 사람 묶음 전체가 함께 위로 올라간다.
        urgent_rank = 0 if any(_di_item_is_urgent(x) for x in block) else 1
        # 다색 주문자는 주 색상 그룹의 뒤쪽으로 보낸다.
        return (urgent_rank, thread_order[thread], code_key(first.get("품목코드")),
                1 if len(codes) > 1 else 0, first_position[recipient])

    result = []
    for recipient in sorted(sequence, key=person_key):
        result.extend(people[recipient])       # 사람 안의 주문 순서는 원본 유지
    return result


def _holding_display_count(it):
    """홀딩도어 장부 G열에 적을 실제 개수. 작동방식 분수는 수량이 아니다."""
    raw = it.get("수량")
    if raw not in (None, "") and not looks_like_holding_operation(raw):
        return raw
    count = it.get("창개수")
    try:
        n = int(float(count)) if count not in (None, "") else 1
    except (TypeError, ValueError):
        n = 1
    return n if n > 1 else None


def _holding_count(it):
    """경영박사 계산용 홀딩도어 실물 개수."""
    count = it.get("창개수")
    try:
        n = int(float(count)) if count not in (None, "") else None
    except (TypeError, ValueError):
        n = None
    if n and n > 0:
        return n
    raw = it.get("수량")
    if raw not in (None, "") and not looks_like_holding_operation(raw):
        try:
            n = int(float(str(raw).replace("X", "").replace("x", "")))
            return max(1, n)
        except (TypeError, ValueError):
            pass
    return 1


def to_rows(order, ship_label, M=None, sort_di=True):
    """주문 1건 -> 장부 행 리스트(dict). 블라인드와 홀딩도어를 함께 지원한다.

    사전점검에서 사람이 확정한 값이 장부/작업지시서/EDI 사이에서 다시
    원본값으로 되돌아가지 않도록 출력 직전 공통 정규화를 반드시 거친다.
    """
    synchronize_order_for_outputs(order)
    client = order.get("거래처")
    mark = CLIENT_INFO.get(client, (None, None, None, ""))[3]
    order_no = str(order.get("주문번호") or "").strip()
    if client == "SP" and order_no:
        m = re.search(r"(?:^|-)\s*(\d+)\s*$", order_no)
        mark = m.group(1) if m else ""
    elif client in {"JL", "DU"} and order_no:
        # JL/두창은 주문번호를 반드시 괄호 포함으로 장부에 표시한다.
        mark = _jl_order_mark(order_no)
    rows = []
    source_index = {id(it): idx for idx, it in enumerate(order.get("items", []))}
    items = [it for it in order.get("items", []) if not it.get("예외품목")]
    if client == "DI" and M is not None and sort_di:
        items = sort_di_items(items, M)
    prev_color = prev_h = prev_di_group = None
    prev_joint = False
    prev_di_recipient = None
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
    if client == "JL":
        cust = _jl_customer_name(cust)

    # v68 기재사항 공통 규칙
    # - 주문 전체에 반복되는 값은 기재사항1(공통)
    # - 개별 창에만 있는 값은 기재사항2(개인창)
    # - 같은 문구가 전체기재/고객명/item 기재에 중복돼도 한 번만 출력
    base_common_parts = _split_note_parts(body, cust)
    shared_item_parts = _shared_item_note_parts(items)
    common_note_parts = _split_note_parts(*base_common_parts, *shared_item_parts)
    total_windows = _total_window_count(items)
    multi_windows = total_windows > 1
    single = not multi_windows and client not in NO_MERGE_PLACE

    for i, it in enumerate(items):
        is_holding = it.get("_product_group") == "holding" or it.get("_holding_accessory")
        is_accessory = bool(it.get("_holding_accessory"))
        if is_holding:
            color = holding_ledger_text(it)
        else:
            color = ledger_color(it.get("품목코드"), it.get("종류"), it.get("타입"),
                                 it.get("_ledger_prefix", "B"),
                                 it.get("_mix_name"), it.get("_mix_codes"))
            if _has_non_di_mix(it, client):
                color = _non_di_mix_ledger_text(it, color)
        di_recipient = (str(it.get("_DIrecipient_context") or it.get("기재사항") or "").strip()
                        if client == "DI" else None)
        same_person = client != "DI" or di_recipient == prev_di_recipient
        di_group = ((it.get("타입"), it.get("종류"))
                    if client == "DI" and not is_holding else None)
        same_color = (not is_accessory and same_person and color == prev_color
                      and (client != "DI" or is_holding or di_group == prev_di_group))
        h = it.get("세로")
        ditto = (not is_accessory and same_color and h == prev_h
                 and client not in NO_DITTO)
        place = it.get("설치장소")
        note = it.get("기재사항") or ""
        if client == "DI" and it.get("_generic_mix") and it.get("_mix_codes"):
            # 미등록 조합 MIX는 품목 칸에는 B MIX/B 원코드 MIX만 쓰고
            # 실제 조합은 기재사항의 맨 앞에 남긴다.
            combo = str(it.get("_mix_codes") or "").replace(" ", "")
            note_parts = [x.strip() for x in str(note).split("/") if x.strip()]
            note = "/".join([combo] + [x for x in note_parts if x != combo])
        # ── v68 기재사항 분배 규칙 ──
        if client == "DI":
            # DI는 기재사항1만 사용: 이름 + MIX 정보. 기재사항2는 사용하지 않는다.
            recv = str(it.get("_DIrecipient_context") or cust or
                       (order.get("배송") or {}).get("수령인") or "").strip()
            mix_note = None
            if it.get("_mix_codes"):
                codes = str(it.get("_mix_codes") or "").replace(" ", "")
                mix_name = str(it.get("_mix_name") or "").strip()
                mix_note = f"{mix_name}{codes}" if mix_name else codes
            note1 = _note_text(recv, mix_note)
            note2 = None
        elif client == "휴안":
            # 휴안도 기재사항1만 사용한다. 이름과 노피스/앙카 등 필요한 제작 메모를
            # 한 칸에 모으고 기재사항2는 비운다.
            recv = str(cust or (order.get("배송") or {}).get("수령인") or "").strip()
            item_parts = _split_note_parts(note)
            note1 = _note_text(recv, *item_parts)
            note2 = None
        elif client == "RT":
            # 루임트는 기존 확정 규칙을 유지한다.
            head = _note_text(body, cust)
            if single:
                note1 = _note_text(head, note, place)
                note2 = None
            else:
                note1 = _note_text(head if i == 0 else None, note)
                note2 = place or None
        elif client == "JO":
            # JO: 주문번호는 항상 기재사항1의 첫 값.
            # 1~2창은 나머지 내용도 기재사항1에 합치고, 3창 이상부터 기재사항2를 쓴다.
            rest = _jo_note2_parts(order, it, cust, first=True)
            if total_windows >= 3:
                note1 = _note_text(order_no)
                note2 = _note_text(*rest)
            else:
                note1 = _note_text(order_no, *rest)
                note2 = None
        elif client == "SP":
            # SP도 여러 창일 때 공통/개인창 규칙을 따른다. 공지·주문번호는 공통.
            sp_common = _split_note_parts(*sp_notices, order_no, *common_note_parts)
            personal = _without_parts(_split_note_parts(note, place), sp_common)
            if multi_windows:
                note1 = _note_text(*sp_common)
                note2 = _note_text(*personal)
            else:
                note1 = _note_text(*sp_common, *personal)
                note2 = None
        elif client == "보노":
            # 보노 기존 규칙: 기재사항1=받는사람/화물지점, 기재사항2=오더명/시공위치.
            note1 = _note_text(order.get("_bono_note1"))
            note2 = _note_text(order_no, place)
        elif client == "미래가공":
            # 미래가공은 기존 업체 전용 배치를 유지한다.
            note1 = _note_text(note, place)
            note2 = order_no if i == 0 else None
        elif client == "DU":
            # 두창: 고객명은 공통, 발주내 기재내용은 개인창.
            note1 = _note_text(cust, (order.get("배송") or {}).get("수령인"))
            detail_parts = _without_parts(
                _split_note_parts(order.get("전체기재사항"), note, place),
                _split_note_parts(note1))
            note2 = _note_text(*detail_parts)
        elif client == "JL":
            # JL: 피스/이름은 공통 기재사항1, 설치장소는 개인창 기재사항2.
            raw_parts = _split_note_parts(common, note)
            has_piece = any("피스" in x for x in raw_parts)
            fixed = (["피스"] if has_piece else []) + ([cust] if cust else [])
            item_parts = _jl_clean_parts(_split_note_parts(note), cust, order_no)
            common_parts = _jl_clean_parts(_split_note_parts(common), cust, order_no)
            note1 = _note_text(*(fixed + item_parts + common_parts))
            note2 = place or None
        else:
            # 전체 공통: 여러 창이면 기재1=공통, 기재2=개인창.
            # 모든 item에 반복된 기재문구는 공통으로 승격하여 대동의
            # `공학1관/공학1관` 같은 중복을 방지한다.
            personal = _without_parts(_split_note_parts(note, place), common_note_parts)
            if multi_windows:
                note1 = _note_text(*common_note_parts)
                note2 = _note_text(*personal)
            else:
                note1 = _note_text(*common_note_parts, *personal)
                note2 = None
        # 택배/화물은 받는 명칭을 기재사항1에서 확인할 수 있게 한다.
        # JO/JL/DU는 별도 고정 규칙이 있으므로 해당 규칙을 우선한다.
        if client not in {"JO", "JL", "DU", "보노", "DI", "휴안"} and _effective_delivery_mode(order) in {"택배", "화물"}:
            recv = str((order.get("배송") or {}).get("수령인") or "").strip()
            if recv:
                note1 = _prepend_note_part(note1, recv)

        # 사전점검에서 직접 수정한 기재사항1/2는 업체별 자동 조합보다 우선한다.
        # 키의 존재 자체를 override 신호로 사용하므로 사용자가 빈칸으로 지운 경우도
        # 파일 생성 때 자동값으로 되살아나지 않는다.
        if "_manual_note1" in it:
            note1 = it.get("_manual_note1")
        if "_manual_note2" in it:
            note2 = it.get("_manual_note2")
        note1, note2 = unify_handle_word(note1), unify_handle_word(note2)
        # 모든 업체에서 `/` 단위 동일 문구는 한 번만 남긴다.
        # 예: 공학1관/공학1관 -> 공학1관
        note1 = _note_text(note1)
        note2 = _note_text(note2)
        # 홀딩도어는 OCR/장부 복원 과정에서 고객명과 기재사항이 같은 값으로
        # 동시에 들어올 수 있다. 장부에는 같은 문구를 두 번 쓰지 않는다.
        if is_holding and note1:
            unique_parts = []
            seen_parts = set()
            for part in str(note1).split("/"):
                part = part.strip()
                key = re.sub(r"\s+", " ", part)
                if part and key not in seen_parts:
                    unique_parts.append(part)
                    seen_parts.add(key)
            note1 = "/".join(unique_parts) or None
        has_tlean, (note1, note2) = _take_tlean(note1, note2)
        joint = bool((not is_holding) and it.get("종류") == "원코드" and it.get("연창"))
        joint_start = joint and (not prev_joint or not same_color)
        manual_special = str(it.get("_수동특이") or "").strip()
        special = " ".join(x for x in (
            "틀안" if has_tlean else None,
            "#" if joint_start else None,
            manual_special or None,
        ) if x) or None

        if is_accessory:
            raw_qty = it.get("수량") if it.get("수량") not in (None, "") else it.get("창개수")
            aq = accessory_qty(raw_qty)
            qty_display = str(raw_qty).strip() if str(raw_qty or "").strip().lower().startswith("x") else f"X{aq:g}"
            model1 = model2 = None
            width_out = height_out = None
        elif is_holding:
            qty_display = _holding_display_count(it)
            model1 = normalize_holding_operation(
                it.get("_holding_operation") or
                (it.get("수량") if looks_like_holding_operation(it.get("수량")) else None))
            model2 = str(it.get("_holding_rail") or "").strip() or None
            width_out, height_out = _display_size(it.get("가로")), ('  \"' if ditto else _display_size(h))
        else:
            qty_display = normalized_ledger_qty(
                it.get("수량") if it.get("수량") not in (None, "")
                else (it.get("창개수") if client in {"JL", "JO", "RT"} else None))
            model1 = _short_dir(it.get("손잡이방향"), client)
            model2 = _handle(it.get("종류"), h, it.get("손잡이길이"))
            if client == "미래가공" and normalize_handle(it.get("손잡이길이"))[0] == 150:
                model2 = None
            width_out, height_out = _display_size(it.get("가로")), ('  \"' if ditto else _display_size(h))

        rows.append({
            "거래처": client if i == 0 else None,
            "내부표시": mark if i == 0 else None,
            "색상": None if same_color else color,
            "가로": width_out,
            "세로": height_out,
            "수량": qty_display,
            "모형1": model1,
            "모형2": model2,
            "특이": special,
            "기재사항": note1,
            "기재사항2": note2,
            "출고일": ship_label if i == 0 else None,
            "_같은색": same_color,
            "_특수": None,
            "_거래처코드": client,
            "_source_item_index": source_index.get(id(it)),
            "_product_group": "holding" if is_holding else "blind",
            "_DI수령인": (str(it.get("_DIrecipient_context") or it.get("기재사항") or "").strip()
                           if client == "DI" else None),
            "_DI품목코드": (str(it.get("품목코드") or "").upper()
                            if client == "DI" else None),
            "_DI유형": di_group,
            "_DI실색": (M.thread_color(it.get("품목코드"))
                       if client == "DI" and M is not None and not is_holding else None),
            # JL 동일 고객의 여러 창은 기재사항을 한 셀로 세로 병합한다.
            "_JL고객": cust if client == "JL" and cust else None,
            "_JO주문번호": order_no if client == "JO" and order_no else None,
            "_JO기재2사용": bool(client == "JO" and total_windows >= 3),
        })
        if not is_accessory:
            prev_color = color
            if h not in (None, ""):
                prev_h = h if not ditto else prev_h
        prev_joint = joint
        prev_di_recipient = di_recipient
        prev_di_group = di_group

    d = order.get("배송") or {}
    if client == "인천)트루":
        for accessory in order.get("_true_accessories") or []:
            label = str(accessory.get("표시") or "").strip()
            if label:
                # 숫자 규칙을 추정하지 않고 사용자가 요청한 1개로 고정한다.
                rows.append({"_특수": "부속", "문구": label, "수량": 1})
    if client == "유앤":
        for accessory in order.get("_unit_accessories") or []:
            label = str(accessory.get("품명") or "").strip()
            if accessory.get("피스수"):
                label += f" ({accessory['피스수']})"
            rows.append({"_특수": "부속", "문구": label,
                         "수량": f"X{accessory.get('세트', 1)}set"})
    if str(order.get("_추가부속") or "").strip():
        rows.append({"_특수": "부속", "문구": str(order.get("_추가부속")).strip(),
                     "수량": None})
    if client == "아지트":
        rows.append({"_특수": "☆", "문구": order.get("_az_place") or "시온가공소"})
    # 전 업체 공통: 택배/화물 주문은 창당 포장비용을 추가한다.
    # 사람이 작성한 기존 장부에서 `포장비용`이 명시된 경우도 보존한다.
    pack_windows = 0
    for pit in items:
        if pit.get("_holding_accessory") or pit.get("_product_group") == "holding":
            continue
        try:
            pack_windows += _blind_counts(pit)[2]
        except Exception:
            pack_windows += 1
    if pack_windows and (_needs_window_packing(order) or "포장비용" in common):
        rows.append({"_특수": "#", "문구": "포장비용", "수량": pack_windows})
    if d.get("방식") in ("택배", "화물", "배달") and d.get("주소"):
        rows.append({"_특수": "☆", "문구": d["주소"]})
        who = " ".join(x for x in (d.get("수령인"), d.get("연락처")) if x)
        # 장부/작업지시서에는 선불만 우측 빨간 글씨로 표시한다.
        # 착불은 내부 배송값에는 보존해 EDI 배송품목 판별에 쓰되 화면/장부에는 적지 않는다.
        pay_mark = "선불" if str(d.get("선불착불") or "").strip() == "선불" else None
        if who or (client == "RT" and pay_mark):
            rows.append({"_특수": "", "문구": who or "",
                         "선불": pay_mark})
        delivery_notice = clean_delivery_notice(d.get("전달사항")) or ""
        if delivery_notice:
            rows.append({"_특수": "", "문구": f"전달 : {delivery_notice}"})
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
            # 샘플 장부 규칙: 헤더 아래 첫 데이터 행도 내부선(hair)이다.
            cell.border = Border(top=HAIR,
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
    previous_order_client = None
    for row in rows:
        if row.get("_특수") is not None:
            if row.get("_특수") == "부속":
                ws.cell(r, 3, row.get("문구"))
                ws.cell(r, 3).alignment = Alignment(horizontal="center",
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
            client_text = row["거래처"]
            mark_text = row["내부표시"]
            # 같은 업체 주문이 연속되면 첫 주문에만 업체명을 쓰고
            # 다음 주문부터는 내부표시만 남긴다. 내부표시가 없는 업체(DI 등)는
            # 공통 반복 표시 (K)를 사용한다. SP/JL처럼 주문별 표시가 있는 경우는 보존한다.
            if client_text:
                same_client = previous_order_client == client_text
                display_client = None if same_client else client_text
                if same_client and not mark_text:
                    mark_text = "(K)"
                previous_order_client = client_text
            else:
                display_client = None
            c2 = ws.cell(r, 2, _client_cell(display_client, mark_text))
            if center_head:
                c2.alignment = Alignment(horizontal="center",
                                         vertical="center", shrink_to_fit=True)
            c3 = ws.cell(r, 3, _color_cell(row["색상"]))
            if center_head:
                c3.alignment = Alignment(horizontal="center",
                                         vertical="center", shrink_to_fit=True)
            ws.cell(r, 4, row["가로"])
            ws.cell(r, 6, row["세로"])
            # 같은 세로의 따옴표 표시는 앞 공백 2칸을 유지하고 셀 중앙에 정렬한다.
            if str(row.get("세로") or "").strip() == '"':
                ws.cell(r, 6).alignment = Alignment(horizontal="center", vertical="center",
                                                    shrink_to_fit=True)
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
            if (row["기재사항"] and not row["기재사항2"]
                    and not row.get("_DI수령인") and not row.get("_JL고객")
                    and not row.get("_JO주문번호")):
                span.append(r)          # 방 이름이 없으면 J:K 를 합친다
            ws.cell(r, date_col, row["출고일"])
        r += 1

    # DI 수령인은 색상이 달라도 같은 사람의 연속 행 전체를 하나로
    # 병합한다. 세부 양식에서는 기재사항 두 칸(K:L)을 직사각형으로 묶는다.
    i = 0
    while i < len(rows):
        recipient = rows[i].get("_DI수령인")
        if not recipient or rows[i].get("_특수") is not None:
            i += 1
            continue
        j = i + 1
        first_note = str(rows[i].get("기재사항") or "").strip()
        while (j < len(rows) and rows[j].get("_특수") is None
               and rows[j].get("_DI수령인") == recipient
               and str(rows[j].get("기재사항") or "").strip() == first_note):
            j += 1
        ws.merge_cells(start_row=start + i, end_row=start + j - 1,
                       start_column=11 if detailed else 10,
                       end_column=12 if detailed else 11)
        i = j

    # JL은 같은 주문/고객의 여러 제품행에서 `피스/이름` 같은 공통
    # 기재사항을 한 번만 보이게 한다. 기재사항2(설치장소)가 전부 비어 있으면
    # K:L 전체를 직사각형으로, 설치장소가 있으면 K열만 세로 병합한다.
    i = 0
    while i < len(rows):
        row0 = rows[i]
        customer = row0.get("_JL고객")
        order_id = row0.get("_oi")
        if not customer or row0.get("_특수") is not None:
            i += 1
            continue
        j = i + 1
        while (j < len(rows) and rows[j].get("_특수") is None
               and rows[j].get("_JL고객") == customer
               and rows[j].get("_oi") == order_id):
            j += 1
        if j - i > 1:
            # 병합할 공통 기재사항은 첫 행의 값을 기준으로 하고, 하위 행에서
            # 고객명/피스만 빠진 형태도 같은 공통값으로 간주한다.
            first_note = str(row0.get("기재사항") or "").strip()
            if first_note:
                all_note2_empty = all(not str(rows[k].get("기재사항2") or "").strip()
                                      for k in range(i, j))
                if all_note2_empty:
                    ws.merge_cells(start_row=start + i, end_row=start + j - 1,
                                   start_column=11 if detailed else 10,
                                   end_column=12 if detailed else 11)
                else:
                    ws.merge_cells(start_row=start + i, end_row=start + j - 1,
                                   start_column=11 if detailed else 10,
                                   end_column=11 if detailed else 10)
                ws.cell(start + i, 11 if detailed else 10).alignment = Alignment(
                    horizontal="center", vertical="center", shrink_to_fit=True)
        i = j

    # JO 주문번호는 기재사항1 열에만 두고, 같은 주문의 여러 행이면 세로 병합한다.
    i = 0
    while i < len(rows):
        row0 = rows[i]
        order_no = row0.get("_JO주문번호")
        order_id = row0.get("_oi")
        if not order_no or row0.get("_특수") is not None:
            i += 1
            continue
        j = i + 1
        while (j < len(rows) and rows[j].get("_특수") is None
               and rows[j].get("_oi") == order_id
               and rows[j].get("_JO주문번호") == order_no):
            j += 1
        # JO는 3창 이상일 때만 기재사항1을 주문번호 전용으로 세로 병합한다.
        # 1~2창은 각 행의 주문번호+개별 내용을 보존해야 하므로 병합하지 않는다.
        if j - i > 1 and row0.get("_JO기재2사용"):
            col = 11 if detailed else 10
            ws.merge_cells(start_row=start + i, end_row=start + j - 1,
                           start_column=col, end_column=col)
            ws.cell(start + i, col).alignment = Alignment(
                horizontal="center", vertical="center", shrink_to_fit=True)
        i = j

    # 전체 공통: 같은 주문 안에서 기재사항 값이 연속해서 완전히 같으면
    # 색상이 달라도 한 셀로 세로 병합한다. DI/JL/JO는 위 전용 병합 규칙을 우선한다.
    for note_col, note_key in ((11 if detailed else 10, "기재사항"),
                               (12 if detailed else 11, "기재사항2")):
        i = 0
        while i < len(rows):
            row0 = rows[i]
            value = str(row0.get(note_key) or "").strip()
            order_id = row0.get("_oi")
            if (row0.get("_특수") is not None or not value or order_id is None
                    or row0.get("_DI수령인") or row0.get("_JL고객")
                    or row0.get("_JO주문번호")):
                i += 1
                continue
            j = i + 1
            while (j < len(rows) and rows[j].get("_특수") is None
                   and rows[j].get("_oi") == order_id
                   and str(rows[j].get(note_key) or "").strip() == value
                   and not rows[j].get("_DI수령인")
                   and not rows[j].get("_JL고객")
                   and not rows[j].get("_JO주문번호")):
                j += 1
            if j - i > 1:
                try:
                    ws.merge_cells(start_row=start + i, end_row=start + j - 1,
                                   start_column=note_col, end_column=note_col)
                    ws.cell(start + i, note_col).alignment = Alignment(
                        horizontal="center", vertical="center", shrink_to_fit=True)
                except ValueError:
                    pass
            i = j

    # DI의 같은 수령인·같은 품목코드·같은 유형이 연속되면
    # 색상(C열) 칸을 세로로 병합한다. P/FP는 별도 코드이므로 서로 합치지 않는다.
    i = 0
    while i < len(rows):
        row0 = rows[i]
        recipient = row0.get("_DI수령인")
        code = row0.get("_DI품목코드")
        type_group = row0.get("_DI유형")
        if not recipient or not code or row0.get("_특수") is not None:
            i += 1
            continue
        j = i + 1
        while (j < len(rows) and rows[j].get("_특수") is None
               and rows[j].get("_DI수령인") == recipient
               and rows[j].get("_DI품목코드") == code
               and rows[j].get("_DI유형") == type_group):
            j += 1
        if j - i > 1:
            ws.merge_cells(start_row=start + i, end_row=start + j - 1,
                           start_column=3, end_column=3)
            ws.cell(start + i, 3).alignment = Alignment(
                horizontal="center" if center_head else "left",
                vertical="center", shrink_to_fit=True)
        i = j

    for rr in span:                 # 기재사항 가로 병합 (J:K)
        ws.merge_cells(start_row=rr, end_row=rr,
                       start_column=11 if detailed else 10,
                       end_column=12 if detailed else 11)

    # 출고일은 같은 주문(휴안·유앤은 한 사람)의 제품·포장·배송·주소·
    # 연락처·전달사항 전체를 세로로 병합한다.
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
            if key == "기재사항" and (row0.get("_DI수령인") or row0.get("_JL고객")
                                          or row0.get("_JO주문번호")):
                i += 1
                continue
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
    is_di = bool(all_rows) and all(r.get("_거래처코드") == "DI" for r in all_rows)
    if is_di:
        first_sheet = True
        for group_name, group_key in DI_SHEET_GROUPS:
            group_rows = [r for r in all_rows if r.get("_DI유형") == group_key]
            for page_no, chunk in enumerate(_di_paginate(group_rows) or [[]], 1):
                title = group_name if page_no == 1 else f"{group_name} ({page_no})"
                page_rows = max(DI_MAX_ROWS, len(chunk))
                if first_sheet:
                    ws = wb.active
                    ws.title = title
                    _init_sheet(ws, page_rows + 2, detailed=True)
                    _block(ws, 1, page_rows, header_date, detailed=True)
                    first_sheet = False
                else:
                    ws = _new_sheet(wb, title, page_rows, header_date,
                                    detailed=True)
                _write_rows(ws, chunk, detailed=True)
    else:
        pages = _paginate_order_rows(all_rows, LEDGER_ROWS)
        for n, chunk in enumerate(pages, 1):
            ws = _new_sheet(wb, "장부" if n == 1 else f"장부 ({n})",
                            LEDGER_ROWS, header_date, detailed=True)
            _write_rows(ws, chunk, detailed=True)
    wb.save(path)
    return path


def _paginate_order_rows(rows, limit=LEDGER_ROWS):
    """같은 주문/사람을 가능한 한 자르지 않고 장부 페이지를 나눈다."""
    if not rows:
        return [[]]
    blocks, i = [], 0
    while i < len(rows):
        order_id = rows[i].get("_oi")
        j = i + 1
        while j < len(rows) and rows[j].get("_oi") == order_id:
            j += 1
        blocks.append(rows[i:j])
        i = j

    pages, current = [], []
    for block in blocks:
        if current and len(current) + len(block) > limit:
            pages.append(current)
            current = []
        if len(block) > limit:
            if current:
                pages.append(current)
                current = []
            pages.extend(block[i:i + limit] for i in range(0, len(block), limit))
        else:
            current.extend(block)
    if current:
        pages.append(current)
    return pages or [[]]


# ─────────────────────────────────────────────
# 3. 장부 -> 작업지시서
# ─────────────────────────────────────────────
INNER = re.compile(r"\s*\((?:[A-Z]|\d+)\)\s*$")


def to_worksheet_rows(rows):
    """내부표시 제거, 색상 앞 'B ' 제거, 특수 행 제외.

    작업지시서도 세부 양식을 사용하므로 #·틀안은 길이/기재사항에
    중복하지 않고 특이 칸에만 유지한다.
    """
    out = []
    for row in rows:
        r = dict(row)
        if r.get("_특수") is None:
            r["내부표시"] = None                  # (K), (N), (1) 등 삭제
            if r.get("색상"):
                r["색상"] = re.sub(r"^\s*[BHRC]\s+", "", r["색상"], flags=re.I).strip()
        # 포장비/부속/주소/전달사항도 장부와 작업지시서가 서로 어긋나지 않게 보존한다.
        out.append(r)
    return out


def _group_len(rows):
    """거래처가 채워진 지점을 기준으로 업체별 묶음 길이"""
    idx = [i for i, r in enumerate(rows) if r.get("거래처")] + [len(rows)]
    return [(idx[i], idx[i + 1] - idx[i]) for i in range(len(idx) - 1)]


def _di_paginate(rows):
    """DI 행을 35행 목표/최대 40행으로 나누되 수령인은 자르지 않는다."""
    blocks = []
    for row in rows:
        recipient = row.get("_DI수령인") or "__unknown"
        if blocks and blocks[-1][0] == recipient:
            blocks[-1][1].append(row)
        else:
            blocks.append((recipient, [row]))

    if not blocks:
        return []

    # 108행처럼 큰 수령인 묶음이 뒤쪽에 있어도 마지막에 2~3행짜리
    # 추가 페이지가 생기지 않도록, 필요 페이지 수를 먼저 정하고 수령인
    # 블록을 균등 배치한다. 페이지 안에서는 원래 색상 순서를 유지한다.
    total = sum(len(block) for _, block in blocks)
    page_count = max(1, (total + DI_MAX_ROWS - 1) // DI_MAX_ROWS)
    indexed = [(i, recipient, block) for i, (recipient, block) in enumerate(blocks)]
    desired = (total + page_count - 1) // page_count
    # 최윤동 26행처럼 큰 묶음은 전용 페이지 틀을 먼저 잡아 둔다.
    large = [x for x in indexed if len(x[2]) > desired / 2]
    reserved = [{"load": len(block), "blocks": [(i, block)]}
                for i, _, block in large]
    large_ids = {i for i, _, _ in large}
    regular_count = max(0, page_count - len(reserved))
    regular = [{"load": 0, "blocks": []} for _ in range(regular_count)]

    leftovers = []
    bi = 0
    for index, recipient, block in indexed:
        if index in large_ids:
            continue
        if bi < regular_count:
            cur = regular[bi]
            if (cur["load"] >= DI_TARGET_ROWS
                    or (cur["load"] + len(block) > desired
                        and cur["load"] >= DI_TARGET_ROWS)):
                bi += 1
            if bi < regular_count and regular[bi]["load"] + len(block) <= DI_MAX_ROWS:
                regular[bi]["blocks"].append((index, block))
                regular[bi]["load"] += len(block)
                continue
        leftovers.append((index, block))

    # 앞장을 채우고 남은 작은 묶음은 큰 수령인 페이지에 붙인다.
    bins = regular + reserved
    for index, block in leftovers:
        candidates = [b for b in bins if b["load"] + len(block) <= DI_MAX_ROWS]
        chosen = min(candidates or bins, key=lambda b: b["load"])
        chosen["blocks"].append((index, block))
        chosen["load"] += len(block)

    # 최초 정렬 순서가 빠른 페이지부터 나오게 하고, 페이지 내부도 유지한다.
    bins.sort(key=lambda b: min((i for i, _ in b["blocks"]), default=999999))
    pages = []
    for b in bins:
        page = []
        for _, block in sorted(b["blocks"], key=lambda x: x[0]):
            page.extend(block)
        pages.append(page)
    return pages


def sort_di_items_in_ledger_order(items, M):
    """DI 품목을 최종 장부의 시트·페이지·행 순서로 반환한다.

    단순 색상 정렬뿐 아니라 종류별 시트 분리와 35~40행 페이지 배치까지
    장부 생성과 똑같이 적용한다. 경영박사 EDI가 발주서 원본 순서가 아닌
    실제 장부에 보이는 순서를 따라야 할 때 사용한다.
    """
    sorted_items = sort_di_items(list(items), M)
    ordered, included = [], set()
    for _, group_key in DI_SHEET_GROUPS:
        group = []
        for item in sorted_items:
            if (item.get("타입"), item.get("종류")) != group_key:
                continue
            group.append({
                "_DI수령인": str(item.get("기재사항") or "").strip(),
                "_EDI품목": item,
            })
            included.add(id(item))
        for page in _di_paginate(group):
            ordered.extend(row["_EDI품목"] for row in page)

    # 검증 전의 미확정 유형이 있더라도 EDI에서 행이 사라지지는 않게 한다.
    ordered.extend(item for item in sorted_items if id(item) not in included)
    return ordered


def build_worksheet(rows, path, header_date=None):
    """업체 묶음을 자르지 않고 SHEET_ROWS 씩 분할.
       한 시트 = 제작용 블록 + 확인용 블록 (내용 동일, 각각 날짜·헤더 포함).
       변경된 장부와 동일하게 방향·길이·특이·기재사항을 분리한 13열
       세부 양식을 사용한다."""
    ws_rows = to_worksheet_rows(rows)
    is_di = bool(ws_rows) and all(r.get("_거래처코드") == "DI" for r in ws_rows)
    if is_di:
        wb = openpyxl.Workbook()
        first_sheet = True
        for group_name, group_key in DI_SHEET_GROUPS:
            group_rows = [r for r in ws_rows if r.get("_DI유형") == group_key]
            for page_no, chunk in enumerate(_di_paginate(group_rows) or [[]], 1):
                title = group_name if page_no == 1 else f"{group_name} ({page_no})"
                ws = wb.active if first_sheet else wb.create_sheet()
                first_sheet = False
                ws.title = title
                page_rows = max(DI_MAX_ROWS, len(chunk))
                _init_sheet(ws, page_rows + 2, detailed=True)
                first = _block(ws, 1, page_rows, header_date, detailed=True)
                _write_rows(ws, chunk, start=first, center_head=True,
                            detailed=True)
        wb.save(path)
        return path

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
        _init_sheet(ws, total, detailed=True)
        f1 = _block(ws, 1, SHEET_ROWS, header_date, detailed=True)
        _write_rows(ws, chunk, start=f1, center_head=True, detailed=True)
        f2 = _block(ws, SHEET_ROWS + 3, SHEET_ROWS, header_date,
                    detailed=True)
        _write_rows(ws, chunk, start=f2, center_head=True,
                    detailed=True)      # 확인용
    wb.save(path)
    return path


# ─────────────────────────────────────────────
# 4. 경영박사 전표
# ─────────────────────────────────────────────
ERP_HEAD = ["날짜", "전표번호", "계정코드", "계정", "거래처관리코드", "상호",
            "대체_코드", "대체_상호", "품목관리코드", "품명", "규격", "수량",
            "단가", "금액", "부가세", "전표적요", "사원코드", "사원"]


def _erp_text_width(text):
    """경영박사 인쇄 폭: 한글 2칸, 영문·숫자·기호 1칸."""
    return sum(2 if "\uac00" <= ch <= "\ud7a3" else 1 for ch in str(text))


def _split_erp_token(token, limit=30):
    """공백 없이 30칸을 넘는 토큰을 문자 경계에서 나눈다."""
    parts, cur, width = [], "", 0
    for ch in str(token):
        cw = 2 if "\uac00" <= ch <= "\ud7a3" else 1
        if cur and width + cw > limit:
            parts.append(cur)
            cur, width = "", 0
        cur += ch
        width += cw
    if cur:
        parts.append(cur)
    return parts


def wrap_erp_text(text, limit=30):
    """EDI ** 행의 적요를 인쇄 폭에 맞게 나눈다.

    `명지로 106`, `도원로 45`, `산호대로25길 44`처럼 도로명과
    바로 뒤의 건물번호는 하나의 단위로 취급해 서로 떨어지지 않게 한다.
    """
    tokens = str(text or "").strip().split()
    units, i = [], 0
    while i < len(tokens):
        token = tokens[i]
        road = token.rstrip(",")
        if (re.search(r"(?:로|길)$", road) and i + 1 < len(tokens)
                and re.match(r"^\d", tokens[i + 1])):
            units.append(token + " " + tokens[i + 1])
            i += 2
        else:
            units.append(token)
            i += 1

    expanded = []
    for unit in units:
        expanded.extend([unit] if _erp_text_width(unit) <= limit
                        else _split_erp_token(unit, limit))

    # 30칸을 그냥 앞에서부터 채우면 마지막 줄에 한 단어만
    # 남을 수 있다. 각 줄의 남는 폭 제곱합이 작아지도록
    # 동적 계획법으로 균형 있게 나눈다.
    n = len(expanded)
    best = [(float("inf"), None)] * (n + 1)
    best[n] = (0, n)
    for i in range(n - 1, -1, -1):
        line = ""
        for j in range(i, n):
            line = expanded[j] if j == i else f"{line} {expanded[j]}"
            width = _erp_text_width(line)
            if width > limit:
                break
            score = (limit - width) ** 2 + best[j + 1][0]
            if score < best[i][0]:
                best[i] = (score, j + 1)

    lines, i = [], 0
    while i < n:
        j = best[i][1] or i + 1
        lines.append(" ".join(expanded[i:j]))
        i = j
    return lines


def _erp_address_text(value):
    """경영박사 EDI 전용 택배/화물 주소 정리.

    - 번지/도로번호 뒤 쉼표 제거
    - 괄호에는 아파트/건물명만 남기고 상세 전달문구는 제거
    - 101동202호 -> 101-202
    장부 원문은 변경하지 않는다.
    """
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    if not text:
        return ""
    dongho = None
    m = re.search(r"(\d+)\s*동\s*(\d+)\s*호", text)
    if m:
        dongho = f"{m.group(1)}-{m.group(2)}"
        text = text[:m.start()] + text[m.end():]

    def clean_paren(m):
        inside = re.sub(r"\s+", " ", m.group(1)).strip()
        # 동/호, 쉼표, 배송메모 앞은 버리고 건물/아파트 이름만 유지한다.
        inside = re.split(r"\d+\s*동|\d+\s*호|[,;/]", inside, maxsplit=1)[0].strip()
        inside = re.sub(r"\s+(?:경비실|문앞|배송|연락|부재|세대).*", "", inside).strip()
        return f"({inside})" if inside else ""

    text = re.sub(r"\(([^()]*)\)", clean_paren, text)
    text = re.sub(r"(\d+(?:-\d+)?(?:번지)?)\s*,\s*", r"\1 ", text, count=1)
    text = re.sub(r"\s{2,}", " ", text).strip(" ,")
    if dongho and dongho not in text:
        text = f"{text} {dongho}".strip()
    return text


def build_erp(orders, M, path, order_date=None):
    """경영박사 EDI용 Excel 97-2003 파일을 만든다.

    블라인드는 기존 계산식을 유지한다. 홀딩도어는 사용자 제공 정답 전표에 맞춰
    세로 최소 150cm, 최종 최소 2.5㎡를 적용하고 작동방식/부속을 별도 품목으로 쓴다.
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

    def write_row(values):
        nonlocal out_row
        for col, value in enumerate(values):
            ws.write(out_row, col, value)
        out_row += 1

    def money(unit, qty, vat_zero=False):
        amount = (Decimal(str(unit)) * Decimal(str(qty))).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP)
        vat = Decimal("0") if vat_zero else (amount / 10).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP)
        return int(amount), int(vat)

    for voucher_no, o in enumerate(orders, 1):
        synchronize_order_for_outputs(o)
        d = o.get("_ship_date") or order_date or date.today()
        if isinstance(d, str):
            try:
                d = date.fromisoformat(d[:10])
            except ValueError:
                d = order_date or date.today()
        dstr = f"{d.year}.{d.month:02d}.{d.day:02d}"
        client = o.get("거래처")
        erp_client = ERP_CLIENT_NAME.get(client, client)
        delivery = o.get("배송") or {}
        _common_parts = [x.strip() for x in str(o.get("전체기재사항") or "").split("/")
                         if x.strip() and x.strip() != "포장비용"]
        extra_accessory = str(o.get("_추가부속") or "").strip()
        if extra_accessory and extra_accessory not in _common_parts:
            _common_parts.append(extra_accessory)
        order_common_note = "/".join(_common_parts)
        # 장부에서 최종 확정되는 기재사항1/2를 EDI도 그대로 재사용한다.
        # 이렇게 하면 같은 수정값이 장부에는 보이고 EDI에서는 빠지는 문제를 막는다.
        ledger_note_by_index = {}
        try:
            for lr in to_rows(o, "", M, sort_di=False):
                idx = lr.get("_source_item_index")
                if isinstance(idx, int) and lr.get("_특수") is None:
                    ledger_note_by_index[idx] = (lr.get("기재사항"), lr.get("기재사항2"))
        except Exception:
            ledger_note_by_index = {}
        blind_valid_count = 0
        blind_window_count = 0
        total_valid_count = 0
        holding_pack_qty = Decimal("0")
        holding_pack_recipient = None
        di_pack_recipient = None
        di_pack_qty = Decimal("0")

        def holding_recipient(it=None):
            it = it or {}
            if client == "DI":
                return str(it.get("_DIrecipient_context") or it.get("기재사항") or
                           di_pack_recipient or "").strip()
            if client == "휴안":
                # 장부 기재사항의 현장/상호명이 경영박사 적요명이며, 배송 수령인은
                # 주소행에 더 짧게 적힐 수 있다. 제공 정답 EDI는 홀딩 적요명의
                # 띄어쓰기를 제거한다(예: 제이원 송길수 부장님 -> 제이원송길수부장님).
                name = str(it.get("기재사항") or delivery.get("수령인") or
                           o.get("고객명") or "").strip()
                return re.sub(r"\s+", "", name)
            return str(it.get("기재사항") or o.get("고객명") or
                       delivery.get("수령인") or "").strip()

        def write_holding_pack(recipient, qty):
            if client == "SP" or not qty or Decimal(str(qty)) <= 0:
                return
            q = Decimal(str(qty)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            unit = 550 if client == "DI" else 1000
            amount, vat = money(unit, q, vat_zero=(client == "DI"))
            name = "포장비용(H)"
            values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                      "", "", name, name, "", float(q), unit,
                      amount, vat, str(recipient or "").strip(), "", ""]
            write_row(values)

        source_items = list(o.get("items", []))
        source_idx_by_id = {id(x): i for i, x in enumerate(source_items)}
        items = source_items
        if client == "DI":
            items = sort_di_items_in_ledger_order(items, M)

        for it in items:
            if it.get("예외품목"):
                continue
            is_holding = it.get("_product_group") == "holding" or it.get("_holding_accessory")
            is_accessory = bool(it.get("_holding_accessory"))

            # ── 홀딩도어 ────────────────────────────
            if is_holding:
                hit = M.find_order_item(it, client)
                if not hit or not hit.get("단가"):
                    continue

                recipient = holding_recipient(it)
                if client == "DI" and not is_accessory:
                    if di_pack_recipient is not None and recipient != di_pack_recipient:
                        write_holding_pack(di_pack_recipient, di_pack_qty)
                        di_pack_qty = Decimal("0")
                    di_pack_recipient = recipient

                if is_accessory:
                    q = Decimal(str(accessory_qty(
                        it.get("수량") if it.get("수량") not in (None, "")
                        else it.get("창개수")))).quantize(Decimal("0.01"))
                    amount, vat = money(hit["단가"], q, vat_zero=(client == "DI"))
                    q_text = f"{float(q):g}EA"
                    memo_parts = [q_text]
                    for value in (it.get("_수동특이"), it.get("기재사항"),
                                  it.get("설치장소"), recipient):
                        for part in str(value or "").split("/"):
                            part = part.strip()
                            if part and part not in memo_parts:
                                memo_parts.append(part)
                    memo = " ".join(memo_parts)
                    values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                              "", "", hit.get("관리코드") or hit["품명"], hit["품명"],
                              hit.get("규격") or "", float(q), int(hit["단가"]),
                              amount, vat, memo, "", ""]
                    write_row(values)
                    total_valid_count += 1
                    continue

                w, h = it.get("가로"), it.get("세로")
                if not (w and h):
                    continue
                count = _holding_count(it)
                base = holding_calc(w, h, hit["단가"])
                q = (Decimal(base["수량"]) * count).quantize(
                    Decimal("0.01"), rounding=ROUND_HALF_UP)
                amount, vat = money(hit["단가"], q, vat_zero=(client == "DI"))
                op = normalize_holding_operation(
                    it.get("_holding_operation") or
                    (it.get("수량") if looks_like_holding_operation(it.get("수량")) else None))
                rail = str(it.get("_holding_rail") or "").strip()
                memo_base = f"{float(w):.1f}*{float(h):.1f}/{op}"
                if rail:
                    memo_base += f"/{rail}"

                note_parts = []
                if client == "DI":
                    if recipient:
                        note_parts.append(recipient)
                    for value in (it.get("_수동특이"), it.get("설치장소")):
                        value = str(value or "").strip()
                        if value and value not in note_parts:
                            note_parts.append(value)
                elif client == "휴안":
                    recipient_key = re.sub(r"\s+", "", recipient)
                    for value in (it.get("_수동특이"), it.get("기재사항"), it.get("설치장소")):
                        for part in str(value or "").split("/"):
                            part = part.strip()
                            if part and re.sub(r"\s+", "", part) != recipient_key \
                                    and part not in note_parts:
                                note_parts.append(part)
                    if recipient:
                        note_parts.append(recipient)
                else:
                    for value in (it.get("기재사항"), it.get("설치장소"), it.get("_수동특이")):
                        for part in str(value or "").split("/"):
                            part = part.strip()
                            if part and part not in note_parts:
                                note_parts.append(part)
                    if recipient:
                        recipient_parts = [x.strip() for x in str(recipient).split("/") if x.strip()]
                        if not recipient_parts or not all(x in note_parts for x in recipient_parts):
                            if recipient not in note_parts:
                                note_parts.append(recipient)
                if note_parts:
                    memo_base += "/" + "/".join(note_parts)
                memo = ("상하로라>" if it.get("_holding_upper_roller") else "") + memo_base
                values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                          "", "", hit.get("관리코드") or hit["품명"], hit["품명"],
                          hit.get("규격") or "", float(q), int(hit["단가"]),
                          amount, vat, memo, "", ""]
                write_row(values)
                total_valid_count += 1

                # 포장비용은 부속이 아닌 본품의 청구 ㎡ 합계로 계산한다.
                if client == "DI":
                    di_pack_qty += q
                elif client != "SP":
                    holding_pack_qty += q
                    if recipient:
                        holding_pack_recipient = recipient

                if it.get("_holding_upper_roller"):
                    ex = M.find_holding_extra("H상하로라(가로m당)", client)
                    if ex and ex.get("단가"):
                        # 제공 휴안 정답: 가로 88cm도 1.0m, DI 270cm는 2.7m.
                        uq = max(Decimal(str(w)) / 100, Decimal("1.0")) * count
                        uq = uq.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                        uamount, uvat = money(ex["단가"], uq, vat_zero=(client == "DI"))
                        values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                                  "", "", ex.get("관리코드") or ex["품명"], ex["품명"],
                                  ex.get("규격") or "", float(uq), int(ex["단가"]),
                                  uamount, uvat, recipient, "", ""]
                        write_row(values)

                charge_each = holding_magnet_charge_qty(op)
                if charge_each:
                    ex = M.find_holding_extra("H추가비용(+자석바1)", client)
                    if ex and ex.get("단가"):
                        mq = Decimal(charge_each * count)
                        mamount, mvat = money(ex["단가"], mq, vat_zero=(client == "DI"))
                        values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                                  "", "", ex.get("관리코드") or ex["품명"], ex["품명"],
                                  ex.get("규격") or "", float(mq), int(ex["단가"]),
                                  mamount, mvat, recipient, "", ""]
                        write_row(values)
                continue

            # ── 기존 블라인드 ───────────────────────
            hit = M.find_order_item(it, client)
            if not hit:
                continue
            w, h = it.get("가로"), it.get("세로")
            if not (w and h):
                continue
            calc = calc_erp(w, h, hit["단가"])
            memo = f"{float(w):.1f}*{float(h):.1f}/1EA"
            extra_parts = []
            if client == "휴안":
                receiver = str(delivery.get("수령인") or
                               o.get("고객명") or "").strip()
                manual_special = str(it.get("_수동특이") or "").strip()
                if manual_special:
                    extra_parts.append(manual_special)
                for value in (it.get("기재사항"), it.get("설치장소")):
                    for part in str(value or "").split("/"):
                        part = part.strip()
                        if part and part != receiver and part not in extra_parts:
                            extra_parts.append(part)
                if receiver:
                    extra_parts.append(receiver)
            elif client == "DI":
                receiver = str(it.get("기재사항") or "").strip()
                if receiver:
                    extra_parts.append(receiver)
                manual_special = str(it.get("_수동특이") or "").strip()
                if manual_special and manual_special not in extra_parts:
                    extra_parts.append(manual_special)
                handle = normalize_handle(it.get("손잡이길이"))[0]
                if handle:
                    extra_parts.append(f"손{handle}")
                if it.get("_mix_name") and it.get("_mix_codes"):
                    mix_codes = str(it["_mix_codes"]).replace(" ", "")
                    extra_parts.append(
                        mix_codes if it.get("_generic_mix") or str(it.get("_mix_name")).upper() == "MIX"
                        else f"{it['_mix_name']}{mix_codes}")
                place = str(it.get("설치장소") or "").strip()
                if place and place not in extra_parts:
                    extra_parts.append(place)
            elif client == "JO":
                # 제이원 EDI에도 주문번호가 반드시 보이고, 나머지 기재사항은 그 뒤에 둔다.
                order_no = str(o.get("주문번호") or "").strip()
                if order_no:
                    extra_parts.append(order_no)
                manual_special = str(it.get("_수동특이") or "").strip()
                if manual_special:
                    extra_parts.append(manual_special)
                for part in _jo_note2_parts(o, it, o.get("고객명") or delivery.get("수령인"), first=True):
                    if part not in extra_parts:
                        extra_parts.append(part)
            elif client == "보노":
                # 보노 EDI도 장부와 같은 기재사항 구조를 사용한다.
                # 손잡이길이는 제작정보로 앞에, 기재사항1/2는 그 뒤에 둔다.
                handle = normalize_handle(it.get("손잡이길이"))[0]
                if handle:
                    extra_parts.append(f"손{handle}")
                note1 = str(o.get("_bono_note1") or "").strip()
                if note1:
                    extra_parts.append(note1)
                order_no = str(o.get("주문번호") or "").strip()
                place = str(it.get("설치장소") or "").strip()
                for part in (order_no, place):
                    if part and part not in extra_parts:
                        extra_parts.append(part)
            elif client == "JL":
                common = str(o.get("전체기재사항") or "")
                receiver = _jl_customer_name(o.get("고객명") or delivery.get("수령인") or "")
                all_parts = _split_note_parts(it.get("_수동특이"), common,
                                             it.get("기재사항"), it.get("설치장소"))
                # JL EDI 적요 순서 고정: 손잡이길이 -> 피스 -> 이름 -> 나머지 기재사항.
                handle = normalize_handle(it.get("손잡이길이"))[0]
                if handle:
                    extra_parts.append(f"손{handle}")
                if any("피스" in x for x in all_parts):
                    extra_parts.append("피스")
                if receiver:
                    extra_parts.append(receiver)
                # 나머지 실제 기재사항은 항상 그 뒤에 붙인다.
                for part in _jl_clean_parts(all_parts, receiver, o.get("주문번호")):
                    if part not in extra_parts:
                        extra_parts.append(part)
            else:
                # 화면에서 직접 입력한 특이사항도 모든 거래처의 EDI 적요에 반영한다.
                manual_special = str(it.get("_수동특이") or "").strip()
                if manual_special:
                    extra_parts.append(manual_special)
                if _has_non_di_mix(it, client):
                    extra_parts.append("MIX")
                # 기재사항은 전산 적요의 맨 뒤에 보존한다.
                # 전체기재사항 -> 행 기재사항 -> 설치장소 순서로 붙여
                # `퇴로로20/주방`처럼 장부에는 보이던 앞부분이 EDI에서 빠지지 않게 한다.
                tail_parts = []
                for value in (order_common_note, it.get("기재사항"), it.get("설치장소")):
                    for part in str(value or "").split("/"):
                        part = part.strip()
                        if part and part not in tail_parts:
                            tail_parts.append(part)
                extra_parts = [x for x in extra_parts if x not in tail_parts]
                extra_parts.extend(tail_parts)
            if client == "DI" and order_common_note:
                for part in [x.strip() for x in order_common_note.split("/") if x.strip()]:
                    extra_parts = [x for x in extra_parts if x != part]
                    extra_parts.append(part)
            # 기재사항1 -> 기재사항2 순서를 EDI 맨 뒤에 강제한다.
            src_idx = source_idx_by_id.get(id(it))
            ledger_notes = ledger_note_by_index.get(src_idx, ()) if src_idx is not None else ()
            ledger_tail = []
            for value in ledger_notes:
                for part in str(value or "").split("/"):
                    part = part.strip()
                    if part and part not in ledger_tail:
                        ledger_tail.append(part)
            if ledger_tail:
                extra_parts = [x for x in extra_parts if x not in ledger_tail]
                extra_parts.extend(ledger_tail)
            extra = "/".join(extra_parts)
            if extra:
                memo += f" {extra}"
            split_n = handle_split(it.get("수량"))
            if not split_n:
                left_n, right_n, window_count = _blind_counts(it)
                erp_parts = []
                dir_counts = []
                if left_n or right_n:
                    if left_n:
                        dir_counts.append(("좌", left_n))
                    if right_n:
                        dir_counts.append(("우", right_n))
                else:
                    dir_counts.append((it.get("손잡이방향") or "우", window_count))
                for direction, count_n in dir_counts:
                    qty = (Decimal(str(calc["수량"])) * count_n).quantize(Decimal("0.01"))
                    amount = int((Decimal(str(calc["단가"])) * qty).quantize(
                        Decimal("1"), rounding=ROUND_HALF_UP))
                    vat = int((Decimal(amount) / 10).quantize(
                        Decimal("1"), rounding=ROUND_HALF_UP))
                    erp_parts.append((direction, qty, amount, vat))
            else:
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
                if client == "DI":
                    vat = 0
                values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                          "", "", hit.get("관리코드") or hit["품명"], hit["품명"],
                          hit["규격"], float(qty), int(calc["단가"]),
                          int(amount), int(vat), memo,
                          1 if dir_ == "좌" else 2, dir_]
                write_row(values)
            blind_valid_count += 1
            try:
                _lpack, _rpack, _npack = _blind_counts(it)
                blind_window_count += max(1, int(_npack or 1))
            except Exception:
                blind_window_count += 1
            total_valid_count += 1

        # 인천)트루 추가부속은 장부/작업지시서와 같은 주문에서 경영박사에도 1개씩 출력한다.
        # 창틀용/커튼박스용 무타공은 사용자 제공 품목코드표의 공식 관리코드로 연결하며,
        # 로컬 마스터에 단가가 없으면 안전하게 0원 + 사전점검 확인으로 남긴다.
        if client == "인천)트루":
            for accessory in o.get("_true_accessories") or []:
                catalog = _true_accessory_catalog(accessory.get("표시")) or accessory
                edi_name = str(catalog.get("품명") or catalog.get("관리코드") or catalog.get("표시") or "").strip()
                if not edi_name:
                    continue
                hit = M.find_named_item(edi_name, client)
                qty = 1.0
                if hit:
                    unit = int(hit.get("단가") or 0)
                    spec = hit.get("규격") or ""
                    code = hit.get("관리코드") or edi_name
                    pname = hit.get("품명") or edi_name
                else:
                    unit = 0
                    spec = ""
                    code = str(catalog.get("관리코드") or edi_name)
                    pname = edi_name
                amount = unit
                vat = int((Decimal(amount) / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                values = [dstr, voucher_no, 3, "외출", erp_client, erp_client, "", "",
                          code, pname, spec, qty, unit, amount, vat, "", "", ""]
                write_row(values)

        # DI는 수령인별 홀딩도어 본품 합계 바로 뒤에 포장비를 붙인다.
        if client == "DI" and di_pack_recipient is not None:
            write_holding_pack(di_pack_recipient, di_pack_qty)
        elif client != "DI" and client != "SP" and holding_pack_qty:
            write_holding_pack(holding_pack_recipient or holding_recipient(), holding_pack_qty)

        # 전 업체 공통 포장비용. 택배/화물 주문이면 실제 출력되는 창 수만큼
        # 한 줄로 합산한다. 보노는 사용자 지정 품목/단가(700원)를 유지하고,
        # 그 외 거래처는 기존 공통 포장비용(25mm, 500원)을 사용한다.
        # 기존 장부에서 `# 포장비용`을 직접 넣은 경우도 역변환 시 보존한다.
        explicit_pack = "포장비용" in str(o.get("전체기재사항") or "")
        if blind_window_count and (_needs_window_packing(o) or explicit_pack):
            pack = BONO_PACKING_ITEM if client == "보노" else PACKING_ITEM
            amount = int(pack["단가"] * blind_window_count)
            vat = int((Decimal(amount) / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
            values = [dstr, voucher_no, 3, "외출", erp_client, erp_client, "", "",
                      pack["품명"], pack["품명"], pack["규격"], float(blind_window_count),
                      pack["단가"], amount, vat, "", "", ""]
            write_row(values)

        # 휴안·유앤·보노·인천)트루 배송/주소 행.
        # 트루는 제품만 EDI에 나오고 주소가 빠지는 문제가 있었으므로 같은 배송행 구조를 사용한다.
        has_true_accessory = client == "인천)트루" and bool(o.get("_true_accessories"))
        cleaned_notice = clean_delivery_notice(delivery.get("전달사항"))
        has_shipping_info = any(str(delivery.get(k) or "").strip() for k in ("주소", "수령인", "연락처", "발신")) or bool(cleaned_notice)
        if (total_valid_count or has_true_accessory) and has_shipping_info and _effective_delivery_mode(o) in {"택배", "화물"}:
            address = _erp_address_text(delivery.get("주소"))
            method = str(delivery.get("방식") or "택배").strip()
            prepaid = str(delivery.get("선불착불") or "선불").strip()
            is_freight = "화물" in method
            is_collect = "착불" in prepaid
            if is_freight:
                ship_code = ("#화물(*파손무책)" if is_collect else
                             "#화물(선불★/파손무책/청구分)")
                ship_spec = "화물/착불" if is_collect else "화물/선불"
            else:
                ship_code = ("#택배(*파손무책/대신)" if is_collect else
                             "#택배(선불★/파손무책/청구分)")
                ship_spec = "택배/착불" if is_collect else "택배/선불"

            address_lines = wrap_erp_text(address) if address else [""]
            values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                      "", "", ship_code, ship_code, ship_spec, 0, 0, 0, 0,
                      address_lines[0], "", ""]
            write_row(values)
            for memo_line in address_lines[1:]:
                values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                          "", "", "**", "**", "", 0, 0, 0, 0,
                          memo_line, "", ""]
                write_row(values)

            if client == "보노" and o.get("_bono_receiver_is_bono"):
                # 보노 수령 건은 주소에 -보노가 들어가므로 다음 줄에는 전화만 반복한다.
                receiver = str(delivery.get("연락처") or "").strip()
            else:
                receiver = " ".join(x for x in (
                    str(delivery.get("수령인") or "").strip(),
                    str(delivery.get("연락처") or "").strip(),
                ) if x)
            if receiver:
                for memo_line in wrap_erp_text(receiver):
                    values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                              "", "", "**", "**", "", 0, 0, 0, 0,
                              memo_line, "", ""]
                    write_row(values)

            sender = str(delivery.get("발신") or "").strip()
            if sender:
                for memo_line in wrap_erp_text(sender):
                    values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                              "", "", "발신", "발신", "", 0, 0, 0, 0,
                              memo_line, "", ""]
                    write_row(values)

            notice = cleaned_notice or ""
            if notice:
                notice_code = f"전달사항({'화물' if is_freight else '택배'})"
                for memo_line in wrap_erp_text(notice):
                    values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                              "", "", notice_code, notice_code, "", 0, 0, 0, 0,
                              memo_line, "", ""]
                    write_row(values)

    wb.save(path)
    return path


# ─────────────────────────────────────────────
def build_all(orders, ship_label, master_path, out_dir="."):
    from pathlib import Path
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    M = Master(master_path)
    rows = []
    for o in orders:
        rows.extend(to_rows(o, ship_label, M))
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
               "대동산업": "DD", "대동": "DD", "유앤아이티엔에스": "유앤",
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
    """사람이 작성한 기존 장부 또는 FitOrder 장부를 작업지시서/EDI 데이터로 읽는다.

    `장부`, `장부 (2)` 시트를 우선한다. 그런 이름이 없으면 상호/색상/규격 헤더가
    있는 첫 실제 장부 시트를 자동으로 찾는다. 홀딩도어는 H 표기를 공식 품목으로 복원한다.
    """
    suffix = Path(filename or getattr(source, "name", "")).suffix.lower()

    def find_header(get_cell, max_row):
        for r in range(1, min(max_row, 12) + 1):
            vals = [re.sub(r"\s+", "", str(get_cell(r, c) or "")) for c in range(1, 14)]
            if "상호" in vals and any(v.startswith("색") for v in vals) \
                    and any(v.startswith("규격") for v in vals):
                return r
        return None

    if suffix == ".xls":
        try:
            import xlrd
        except ImportError as e:
            raise RuntimeError("기존 .xls 장부 처리 모듈이 없습니다.") from e
        contents = source.getvalue() if hasattr(source, "getvalue") else None
        book = xlrd.open_workbook(filename=None if contents is not None else str(source),
                                  file_contents=contents)

        def make_info(ws):
            def get_cell(row, col):
                if row < 1 or col < 1 or row > ws.nrows or col > ws.ncols:
                    return None
                return ws.cell_value(row - 1, col - 1)
            hr = find_header(get_cell, ws.nrows)
            return (ws.name, ws.nrows, get_cell, hr) if hr else None

        preferred = [ws for ws in book.sheets()
                     if ws.name == "장부" or re.fullmatch(r"장부 \(\d+\)", ws.name)]
        infos = [make_info(ws) for ws in preferred]
        sheets = [x for x in infos if x]
        if not sheets:
            # FitOrder DI 장부는 C 원코드/C 투코드/...처럼 여러 실제 시트로
            # 나뉘므로 헤더가 있는 모든 장부 시트를 읽는다. 견본/샘플/현장은 제외한다.
            for ws in book.sheets():
                if re.search(r"견본|샘플|현장", ws.name):
                    continue
                info = make_info(ws)
                if info:
                    sheets.append(info)
    else:
        wb = openpyxl.load_workbook(source, data_only=True)

        def make_info(ws):
            get_cell = lambda row, col: ws.cell(row, col).value
            hr = find_header(get_cell, ws.max_row)
            return (ws.title, ws.max_row, get_cell, hr) if hr else None

        preferred = [ws for ws in wb.worksheets
                     if ws.title == "장부" or re.fullmatch(r"장부 \(\d+\)", ws.title)]
        infos = [make_info(ws) for ws in preferred]
        sheets = [x for x in infos if x]
        if not sheets:
            for ws in wb.worksheets:
                if re.search(r"견본|샘플|현장", ws.title):
                    continue
                info = make_info(ws)
                if info:
                    sheets.append(info)
    if not sheets:
        raise ValueError("상호/색상/규격이 있는 장부 시트를 찾지 못했습니다.")

    rows, orders, errors = [], [], []
    current = None
    prev_color = prev_height = None
    prev_di_recipient_ctx = None
    for sheet_name, max_row, get_cell, header_row in sheets:
        header = [str(get_cell(header_row, c) or "").strip() for c in range(1, 14)]
        detailed = "특이" in header
        for excel_row in range(header_row + 1, max_row + 1):
            values = [get_cell(excel_row, c) for c in range(1, 14)]
            client_cell, color_cell = values[1], values[2]
            if str(client_cell or "").replace(" ", "") in ("상호", "현장용"):
                current = None
                prev_color = prev_height = None
                prev_di_recipient_ctx = None
                continue
            width, height = values[3], values[5]
            qty, direction, length = values[6], values[7], values[8]
            if detailed:
                special, note1, note2 = values[9], values[10], values[11]
            else:
                special = "#" if "#" in str(length or "") else None
                note1, note2 = values[9], values[10]
            meaningful = [values[i] for i in (1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12)]
            if not any(v not in (None, "") for v in meaningful):
                continue

            # 인천)트루의 별도 부속행을 사람이 작성한 장부 -> EDI에서도 복원한다.
            # 장부 표시명은 현장용이고, EDI는 공식 품명/관리코드로 변환한다.
            if current is not None and current.get("거래처") == "인천)트루" \
                    and height in (None, "") and width in (None, ""):
                acc = _true_accessory_catalog(color_cell)
                if acc:
                    acc["수량"] = 1
                    current.setdefault("_true_accessories", [])
                    if not any(x.get("표시") == acc.get("표시") for x in current["_true_accessories"]):
                        current["_true_accessories"].append(acc)
                    rows.append({"_특수": "부속", "문구": acc.get("표시"), "수량": 1})
                    continue

            # 루임트 MIX 조합의 길이 보조행(예: 030 / )128cm, 990 / )32cm)은
            # 별도 제품이 아니다. 바로 위 `B 원코드 030+990`의 구성 정보이므로 건너뛴다.
            if current is not None and current.get("거래처") == "RT" \
                    and height in (None, "") and color_cell in (None, ""):
                width_token = re.sub(r"\.0$", "", str(width or "").strip())
                qty_token = str(qty or "").strip()
                if re.fullmatch(r"\d{2,3}", width_token) and re.search(r"cm", qty_token, re.I):
                    continue

            # 주소/수령인/전달 등 장부 특수행. 사람이 작성한 장부를 EDI로
            # 되돌릴 때 배송 정보도 같이 복원한다. ☆행은 주소가 비어 있고
            # 우측에 선불만 적힌 경우도 있으므로 width가 문자열일 때만으로 제한하지 않는다.
            marker = str(color_cell or "").strip()
            is_special_row = (height in (None, "") and marker in ("#", "☆", "") and
                              (isinstance(width, str) or marker == "☆"))
            if is_special_row:
                text = str(width or "").strip()
                rows.append({"_특수": marker, "문구": width})
                if current is not None:
                    # 사람이 작성한 장부의 `# 포장비용`도 장부→EDI에서 잃지 않는다.
                    if marker == "#" and re.sub(r"\s+", "", text) == "포장비용":
                        parts = [x.strip() for x in str(current.get("전체기재사항") or "").split("/") if x.strip()]
                        if "포장비용" not in parts:
                            parts.append("포장비용")
                        current["전체기재사항"] = "/".join(parts)
                    delivery = current.setdefault("배송", {})
                    prepaid = str(note2 or note1 or "").strip()
                    if prepaid in {"선불", "착불"}:
                        delivery["선불착불"] = prepaid
                    if marker == "☆":
                        if text:
                            delivery["주소"] = text
                            default_delivery = CLIENT_INFO.get(current.get("거래처"), (None, None, None))[2]
                            if default_delivery in {"택배", "화물", "배달"}:
                                delivery["방식"] = default_delivery
                            elif not delivery.get("방식"):
                                delivery["방식"] = "택배"
                            if current.get("거래처") == "보노" and not delivery.get("선불착불"):
                                delivery["선불착불"] = "착불"
                            # 이 주소 행까지가 현재 수령인의 주문 묶음. 다음 제품행은
                            # 상호가 공란이어도 같은 거래처의 새 주문으로 시작한다.
                            current["_split_next_product"] = True
                    elif marker == "":
                        # 전달/발신 행에 전화번호가 있어도 수령인으로 오인하지 않는다.
                        if text.startswith("전달"):
                            delivery["전달사항"] = text.split(":", 1)[-1].strip()
                        elif text.startswith("발신"):
                            delivery["발신"] = text.split(":", 1)[-1].strip()
                        else:
                            phone = re.search(r"(01\d[- ]?\d{3,4}[- ]?\d{4})", text)
                            if phone:
                                number = phone.group(1).replace(" ", "")
                                receiver = text[:phone.start()].strip(" /,-")
                                delivery["연락처"] = number
                                if receiver:
                                    delivery["수령인"] = receiver
                                    current["고객명"] = receiver
                                    if current.get("거래처") == "보노" and "보노" in receiver.replace(" ", ""):
                                        current["_bono_receiver_is_bono"] = True
                continue

            # ☆ 주소행을 지난 뒤 상호가 공란인 다음 제품은 같은 거래처의
            # 새 수령인 주문으로 분리한다. (휴안 정답 장부 형태)
            if client_cell in (None, "") and current is not None \
                    and current.pop("_split_next_product", False):
                client = current.get("거래처")
                current = {"거래처": client, "주문번호": None, "고객명": None,
                           "전체기재사항": None, "배송": {}, "items": []}
                orders.append(current)
                prev_color = prev_height = None
                prev_di_recipient_ctx = None

            if client_cell not in (None, "") and str(client_cell).strip():
                raw_client_mark = str(client_cell).strip()
                client = _ledger_client(client_cell)
                marker_only = False
                marker_order_no = None
                if not client and current is not None:
                    # 같은 업체를 연속 작성할 때 상호 대신 (K)/(F) 같은 내부표시만
                    # 쓰거나, SP/JL은 주문번호만 쓰는 장부를 다시 읽을 수 있게 한다.
                    paren_mark = re.fullmatch(r"\(([A-Za-z0-9-]+)\)", raw_client_mark)
                    if paren_mark:
                        client = current.get("거래처")
                        marker_only = True
                        if current.get("거래처") == "JL" and paren_mark.group(1).upper() != "K":
                            marker_order_no = paren_mark.group(1)
                    elif current.get("거래처") in {"SP", "JL"} and \
                            re.fullmatch(r"[A-Za-z0-9-]+", raw_client_mark):
                        client = current.get("거래처")
                        marker_only = True
                        marker_order_no = raw_client_mark
                if not client:
                    errors.append(f"{sheet_name} {excel_row}행: 거래처 '{client_cell}'를 찾을 수 없습니다.")
                    current = None
                    prev_color = prev_height = None
                    continue
                current = {"거래처": client, "주문번호": marker_order_no, "고객명": None,
                           "전체기재사항": None, "배송": {}, "items": []}
                orders.append(current)
                prev_color = prev_height = None
                prev_di_recipient_ctx = None
            if current is None:
                errors.append(f"{sheet_name} {excel_row}행: 주문의 첫 행에 상호가 없습니다.")
                continue

            raw_color = str(color_cell or "").strip()
            color_text = raw_color or str(prev_color or "").strip()
            if HOLDING_FEATURE_ENABLED:
                accessory_hit = holding_accessory_from_text(color_text)
                product_hit = None if accessory_hit else holding_product_from_text(color_text)
                is_holding = bool(accessory_hit or product_hit)
            else:
                # 홀딩 자동 판별 OFF: H/홀딩처럼 보여도 일반 장부 행으로 계속 처리한다.
                accessory_hit = product_hit = None
                is_holding = False

            if str(height or "").strip() in ('"', '”'):
                height = prev_height
            elif height not in (None, ""):
                prev_height = height

            if is_holding:
                if accessory_hit:
                    # 부속은 앞 본품의 색상 반복 기준을 바꾸지 않는다.
                    aq = accessory_qty(qty)
                    item = {
                        "품목코드": accessory_hit.get("코드"), "색상원문": color_text,
                        "종류": "투코드", "타입": "C자", "가로": None, "세로": None,
                        "수량": qty if qty not in (None, "") else aq,
                        "손잡이방향": None, "손잡이길이": None, "연창": False,
                        "설치장소": note2, "기재사항": note1, "예외품목": None,
                        "창개수": None, "좌개수": None, "우개수": None,
                        "_product_group": "holding", "_holding_accessory": True,
                        "_holding_product_name": accessory_hit["품명"],
                        "_holding_accessory_label": accessory_hit.get("장부표시"),
                        "_holding_accessory_color": accessory_hit.get("부속색상") or "화이트",
                        "_ledger_prefix": "H",
                    }
                    if current.get("거래처") == "DI":
                        item["_DIrecipient_context"] = prev_di_recipient_ctx
                    if M is not None and not M.find_order_item(item, current["거래처"]):
                        errors.append(f"{sheet_name} {excel_row}행: 등록되지 않은 홀딩 부속 {color_text}")
                    current["items"].append(item)
                    rows.append({"거래처": current["거래처"] if len(current["items"]) == 1 else None,
                                 "내부표시": None, "색상": color_text,
                                 "가로": None, "세로": None, "수량": qty,
                                 "모형1": None, "모형2": None, "특이": special,
                                 "기재사항": note1, "기재사항2": note2, "출고일": None,
                                 "_같은색": False, "_특수": None,
                                 "_product_group": "holding"})
                    continue

                prev_color = raw_color or prev_color
                w, h = _ledger_number(width), _ledger_number(height)
                if w is None or h is None:
                    errors.append(f"{sheet_name} {excel_row}행: 홀딩도어 가로·세로 규격을 확인해 주세요.")
                op_source = direction if looks_like_holding_operation(direction) else \
                    (qty if looks_like_holding_operation(qty) else None)
                operation = normalize_holding_operation(op_source)
                rail = str(length or "").strip() if "레일" in str(length or "") else None
                raw_qty = None if looks_like_holding_operation(qty) else qty
                try:
                    count = int(float(raw_qty)) if raw_qty not in (None, "") else 1
                except (TypeError, ValueError):
                    count = 1
                di_ctx = None
                if current.get("거래처") == "DI":
                    visible_recipient = str(note1 or "").strip()
                    if visible_recipient:
                        prev_di_recipient_ctx = visible_recipient
                    di_ctx = prev_di_recipient_ctx
                item = {
                    "품목코드": product_hit.get("코드") if product_hit else None,
                    "색상원문": color_text, "종류": "투코드", "타입": "C자",
                    "가로": w, "세로": h, "수량": raw_qty,
                    "손잡이방향": None, "손잡이길이": None, "연창": False,
                    "설치장소": note2, "기재사항": note1, "예외품목": None,
                    "창개수": count, "좌개수": None, "우개수": None,
                    "_수동특이": re.sub(r"(?:#|틀안)", " ", str(special or "")).strip() or None,
                    "_product_group": "holding", "_ledger_prefix": "H",
                    "_DIrecipient_context": di_ctx,
                    "_holding_product_name": product_hit.get("품명") if product_hit else None,
                    "_holding_label": product_hit.get("장부표시") if product_hit else None,
                    "_holding_operation": operation,
                    "_holding_rail": rail,
                    # +상하로라는 현재 행에 실제로 적혀 있을 때만 적용한다.
                    "_holding_upper_roller": "+상하로라" in raw_color.replace(" ", ""),
                }
                if not product_hit:
                    errors.append(f"{sheet_name} {excel_row}행: 홀딩도어 품목 '{color_text}'를 찾지 못했습니다.")
                elif M is not None and not M.find_order_item(item, current["거래처"]):
                    errors.append(f"{sheet_name} {excel_row}행: 등록되지 않은 홀딩도어 품목 {color_text}")
                current["items"].append(item)
                rows.append({"거래처": current["거래처"] if len(current["items"]) == 1 else None,
                             "내부표시": None, "색상": raw_color or None,
                             "가로": w, "세로": ('"' if str(values[5] or "").strip() in ('"','”') else h),
                             "수량": raw_qty, "모형1": operation, "모형2": rail,
                             "특이": special, "기재사항": note1, "기재사항2": note2,
                             "출고일": None, "_같은색": not bool(raw_color), "_특수": None,
                             "_product_group": "holding"})
                continue

            # ── 기존 블라인드 장부 ──
            if current.get("거래처") == "JO":
                # 사람이 작성한 JO 장부: 기재사항1=주문번호, 기재사항2=나머지 메모.
                if note1 not in (None, ""):
                    jo_no = str(note1).strip()
                    if jo_no:
                        current["주문번호"] = jo_no
                note1 = _jo_clean_note_text(note2)
                note2 = None
            elif current.get("거래처") == "보노":
                # 보노 장부 역변환: 기재사항1=받는사람/화물지점,
                # 기재사항2=오더명/시공위치.
                if note1 not in (None, ""):
                    current["_bono_note1"] = str(note1).strip()
                note2_parts = [x.strip() for x in str(note2 or "").split("/") if x.strip()]
                if note2_parts:
                    if not current.get("주문번호"):
                        current["주문번호"] = note2_parts[0]
                    note2 = "/".join(note2_parts[1:]) or None
                note1 = None
            if raw_color:
                prev_color = raw_color
            code, kind, type_ = parse_ledger_color(color_text, note1)
            mix_name, mix_codes = parse_ledger_mix(color_text, note1)
            if mix_name == "MIX" and mix_codes:
                # 사람이 작성한 장부의 B MIX/B 원코드 MIX는 조합을 기재사항에서 읽고,
                # EDI 수령인/메모에는 조합 문자열을 중복시키지 않는다.
                note_parts = [x.strip() for x in str(note1 or "").split("/") if x.strip()]
                note1 = "/".join(x for x in note_parts if x.replace(" ", "") != mix_codes) or None
            w, h = _ledger_number(width), _ledger_number(height)
            if not code:
                errors.append(f"{sheet_name} {excel_row}행: 품목 '{color_text}'를 해석할 수 없습니다.")
            if w is None or h is None:
                errors.append(f"{sheet_name} {excel_row}행: 가로·세로 규격을 확인해 주세요.")

            if current.get("거래처") == "JL":
                # JL 장부의 `피스/강종민` 같은 병합 기재사항에서 고객명을 복원한다.
                # 두 번째 이후 행의 기재사항이 병합으로 비어 있어도 첫 행에서 저장한
                # 고객명을 현재 주문 전체에 사용한다.
                for part in _split_note_parts(note1):
                    candidate = _jl_customer_name(part)
                    if not candidate or "피스" in candidate or JL_PACKAGING.search(candidate):
                        continue
                    if re.fullmatch(r"손\d{2,3}", candidate):
                        continue
                    if candidate not in {"틀안", "포장X", "브라켓"}:
                        current["고객명"] = candidate
                        break

            item = {"품목코드": code, "종류": kind, "타입": type_,
                    "가로": w, "세로": h, "수량": qty or 1,
                    "손잡이방향": (str(direction).strip()
                                    if re.fullmatch(r"[좌우]{1,10}", str(direction or "").strip())
                                    else None),
                    "손잡이길이": parse_handle(length)[0],
                    "연창": "#" in str(special or ""),
                    "_수동특이": re.sub(r"(?:#|틀안)", " ",
                                      str(special or "")).strip() or None,
                    "설치장소": note2, "기재사항": note1,
                    "예외품목": None}
            if mix_name and mix_codes:
                item["_mix_name"], item["_mix_codes"] = mix_name, mix_codes
                if mix_name == "MIX":
                    item["_generic_mix"] = True
            if code and M is not None and not M.find_order_item(item, current["거래처"]):
                errors.append(f"{sheet_name} {excel_row}행: 마스터에 없는 품목 {color_text}")
            current["items"].append(item)
            rows.append({"거래처": current["거래처"] if len(current["items"]) == 1 else None,
                         "내부표시": None, "색상": color_text,
                         "가로": w, "세로": h, "수량": qty,
                         "모형1": item["손잡이방향"], "모형2": length,
                         "특이": special, "기재사항": note1,
                         "기재사항2": note2, "출고일": None,
                         "_같은색": color_cell in (None, ""), "_특수": None})
    for _order in orders:
        expand_same_size_directions(_order)
    return rows, orders, errors


# ─────────────────────────────────────────────
# 편집 반영: 장부 표시값 -> 주문 항목
# ─────────────────────────────────────────────
def _ledger_mix_codes(*values):
    for value in values:
        m = re.search(r"(?<!\d)(\d{3}(?:\s*\+\s*\d{3})+)(?!\d)", str(value or ""))
        if m:
            return re.sub(r"\s+", "", m.group(1))
    return None


def parse_ledger_mix(text, note=None):
    """정식 MIX 또는 사람이 적은 `B MIX` + 기재사항 코드 조합을 복원."""
    s = str(text or "").strip()
    m = re.search(r"([가-힣A-Za-z0-9]+)\s*\(\s*(\d{3}(?:\s*\+\s*\d{3})+)\s*\)", s)
    if m:
        return m.group(1), re.sub(r"\s+", "", m.group(2))
    if re.search(r"\bMIX\b", s, re.I):
        return "MIX", _ledger_mix_codes(s, note)
    # 사용자가 코드 괄호를 생략해도 제공된 색상표에 있는 이름이면 복원한다.
    name_only = re.sub(r"^\s*[BHRC]\s+", "", s, flags=re.I)
    name_only = re.sub(r"^L(?:18|21)-(?:원코드|투코드|셔터)\s+", "", name_only, flags=re.I)
    name_only = re.sub(r"^(?:원코드|투코드|셔터)\s+", "", name_only)
    known = di_mix_info(name_only)
    return known if known else (None, None)


def parse_ledger_color(text, note=None):
    """장부 색상 문자열 -> (코드/믹스조합, 종류, 타입)."""
    s = str(text or "").strip()
    if not s:
        return None, None, None
    type_ = "L자" if re.search(r"L(18|21)-", s) else "C자"
    kind = "투코드"
    for k in ("원코드", "셔터", "투코드"):
        if k in s:
            kind = k
            break
    mix_name, mix_codes = parse_ledger_mix(s, note)
    if mix_name and mix_codes:
        return mix_codes, kind, type_
    m = re.search(r"(\d{3,4}[A-Za-z]*)\s*$", s)
    return (m.group(1) if m else None), kind, type_


def parse_handle(text):
    """'#손140' -> (140, True) / '손120' -> (120, False)"""
    s = str(text or "").strip()
    yeon = s.startswith("#")
    m = re.search(r"(\d+)", s)
    return (int(m.group(1)) if m else None), yeon


def apply_edit(order, item_idx, field, value, color_prefix="B"):
    """장부 셀 편집값을 주문 항목에 되돌려 적용. 홀딩도어도 같은 편집 화면을 사용한다."""
    it = order["items"][item_idx]
    v = None if value in (None, "", "None") else value
    is_holding = it.get("_product_group") == "holding" or it.get("_holding_accessory")
    if field == "상호":
        raw = str(v or "").strip()
        code = raw.split()[0] if raw else None
        aliases = {"두창": "DU", "두창블라인드": "DU",
                   "대일": "DI", "대일산업": "DI", "루임트": "RT"}
        code = aliases.get(code, code)
        if code in CLIENT_INFO:
            order["거래처"] = code
    elif field == "색상":
        text = str(v or "").strip()
        accessory = holding_accessory_from_text(text) if HOLDING_FEATURE_ENABLED else None
        product = (None if accessory else holding_product_from_text(text)) if HOLDING_FEATURE_ENABLED else None
        if accessory or product:
            it["_product_group"] = "holding"
            it["_ledger_prefix"] = "H"
            it["색상원문"] = text
            it["타입"], it["종류"] = "C자", "투코드"
            it["손잡이방향"] = it["손잡이길이"] = None
            it["예외품목"] = None
            if accessory:
                it["_holding_accessory"] = True
                it["_holding_product_name"] = accessory["품명"]
                it["_holding_accessory_label"] = accessory.get("장부표시")
                it["_holding_accessory_color"] = accessory.get("부속색상") or "화이트"
                it["품목코드"] = accessory.get("코드")
            else:
                it.pop("_holding_accessory", None)
                it["_holding_product_name"] = product["품명"]
                it["_holding_label"] = product.get("장부표시")
                it["_holding_upper_roller"] = "+상하로라" in text.replace(" ", "")
                it["품목코드"] = product.get("코드")
                it.setdefault("_holding_operation", "편")
        else:
            # 사용자가 다시 일반 블라인드 품명을 입력하면 홀딩 상태를 제거한다.
            for key in list(it):
                if key.startswith("_holding_"):
                    it.pop(key, None)
            it.pop("_product_group", None)
            code, kind, type_ = parse_ledger_color(v)
            if code:
                it["품목코드"], it["종류"], it["타입"] = code, kind, type_
                mix_name, mix_codes = parse_ledger_mix(v, it.get("기재사항"))
                if re.search(r"\bMIX\b", text, re.I) and not mix_codes:
                    # 색상 칸만 B MIX로 고친 경우 기존에 판독한 코드 조합은 유지한다.
                    mix_name, mix_codes = "MIX", it.get("_mix_codes")
                if mix_name and mix_codes:
                    known = di_mix_info(mix_name)
                    if known and known[1] == mix_codes:
                        mix_name, mix_codes = known
                    it["_mix_name"], it["_mix_codes"] = mix_name, mix_codes
                    it["_generic_mix"] = str(mix_name).upper() == "MIX"
                    it["색상원문"] = mix_name
                else:
                    it.pop("_mix_name", None)
                    it.pop("_mix_codes", None)
                    it.pop("_generic_mix", None)
                m = re.match(r"\s*([BHRC])(?:\s|$)", str(v or ""), re.I)
                prefix = m.group(1).upper() if m else str(color_prefix or "B").upper()
                it["_ledger_prefix"] = prefix if prefix in {"B", "H", "R", "C"} else "B"
        is_holding = it.get("_product_group") == "holding" or it.get("_holding_accessory")
    elif field in ("가로", "세로"):
        if v is not None and str(v).strip() not in ('"', "”"):
            try:
                it[field] = float(str(v).replace(",", ""))
            except ValueError:
                pass
    elif field == "수량":
        it["수량"] = v
        # 사전점검 수량 수정은 장부뿐 아니라 EDI의 창개수에도 같은 값으로 반영한다.
        try:
            count = int(float(v)) if v not in (None, "") and not handle_split(v) else None
        except (TypeError, ValueError):
            count = None
        if count is not None and count >= 1:
            it["수량"] = count
            it["창개수"] = count
            if it.get("손잡이방향") == "좌":
                it["좌개수"], it["우개수"] = count, 0
            elif it.get("손잡이방향") == "우":
                it["좌개수"], it["우개수"] = 0, count
    elif field in ("모형1", "방향"):
        if is_holding and not it.get("_holding_accessory"):
            it["_holding_operation"] = normalize_holding_operation(v)
            it["손잡이방향"] = None
        else:
            if str(v or "").strip() == "ㅈ":
                v = "좌"
            it["손잡이방향"] = v if v in ("좌", "우") else None
            try:
                count = max(1, int(float(it.get("창개수") or it.get("수량") or 1)))
            except (TypeError, ValueError):
                count = 1
            if it["손잡이방향"] == "좌":
                it["좌개수"], it["우개수"] = count, 0
            elif it["손잡이방향"] == "우":
                it["좌개수"], it["우개수"] = 0, count
    elif field in ("모형2", "길이"):
        if is_holding and not it.get("_holding_accessory"):
            it["_holding_rail"] = str(v or "").strip() or None
            it["손잡이길이"] = None
        else:
            n, yeon = parse_handle(v)
            it["손잡이길이"] = n
            if yeon:
                it["연창"] = True
    elif field == "특이":
        text = str(v or "")
        if not is_holding:
            it["연창"] = "#" in text
        it["_수동특이"] = re.sub(r"(?:#|틀안)", " ", text).strip() or None
        if order.get("거래처") == "DI" and "긴급" in text:
            _propagate_di_urgent_order(order)
        old = str(it.get("기재사항") or "")
        parts = [x.strip() for x in old.split("/")
                 if x.strip() and x.strip() != "틀안"]
        if "틀안" in text:
            parts.insert(0, "틀안")
        it["기재사항"] = "/".join(parts) or None
    elif field == "기재사항":
        text = str(v or "").strip()
        # 사전점검에 보이는 최종 문자열을 그대로 보존한다. 업체별 재조합 때문에
        # 장부/EDI에서 다시 달라지는 것을 막기 위한 명시적 override다.
        it["_manual_note1"] = text or None
        if it.get("_generic_mix"):
            combo = _ledger_mix_codes(text)
            if combo:
                it["_mix_codes"] = combo
                it["품목코드"] = combo
                parts = [x.strip() for x in text.split("/") if x.strip()]
                text = "/".join(x for x in parts if x.replace(" ", "") != combo)
        it["기재사항"] = text or None
    elif field == "기재사항2":
        text = str(v or "").strip()
        it["_manual_note2"] = text or None
        it["설치장소"] = text or None
    elif field == "출고일":
        text = str(v or "").strip().replace(".", "-").replace("/", "-")
        try:
            order["_ship_date"] = date.fromisoformat(text).isoformat()
        except ValueError:
            pass

