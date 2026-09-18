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

from rules import (BONO_LIKE_CLIENTS, CLIENT_INFO, ERP_CLIENT_NAME, NO_DITTO, NO_MERGE_PLACE,
                   DI_CODE_ORDER, clean_delivery_notice, strip_vinyl_text,
                   MIX_WORD_RX, MIX_LAYOUT_RX, mix_extra_cost_name, mix_length_text, mix_memo_text, parse_mix_parts,
                   PACKING_CLIENTS, PACKING_ITEM, BONO_PACKING_ITEM, BONO_STYLE_CLIENTS, TRUE_ACCESSORY_PRODUCTS,
                   Master, calc_erp, calc_roll_erp, default_handle_length, di_mix_info,
                   is_client_staff_name, ledger_color, normalize_handle, normalize_roll_handle)
from roll_combo import ROLL_CLIENT_INFO, _roll_note_ignored

from holding import (HOLDING_FEATURE_ENABLED, accessory_qty, holding_accessory_from_text, holding_calc,
                     holding_ledger_text, holding_magnet_charge_qty, holding_magnet_extra_names,
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
THICK = Side(style="thick")
GREEN, RED, BLUE, BLACK = "FF008000", "FFFF0000", "FF0000FF", "FF000000"


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
    """색상 열 리치텍스트.

    - P/FP 품목: ``029FP``/``102P``처럼 P/FP가 붙은 코드 전체를 파란색
    - ``원코드``와 바로 뒤의 ``_``는 초록, ``셔터``는 빨강
    - 나머지 문자는 셀의 기존 글꼴색을 상속하지 않고 검정으로 고정

    예) ``B 원코드_029FP`` -> B(검정) / 원코드_(초록) / 029FP(파랑)
    예) ``B 원코드_820``   -> B(검정) / 원코드_(초록) / 820(검정)
    """
    if not text:
        return text
    s = str(text)
    spans = []

    # P/FP 색상코드는 숫자와 접미사를 포함한 코드 전체를 파란색으로 칠한다.
    # 예: 029FP -> 전체 파랑, 102P -> 전체 파랑.
    # 앞에 '_'나 공백, 재질 접두사가 붙어 있어도 실제 P/FP 코드만 잡는다.
    for m in re.finditer(r"\d+(?:FP|P)(?![A-Za-z0-9])", s, re.I):
        spans.append((m.start(), m.end(), BLUE, 30))

    # 원코드 뒤에 '_'가 붙은 표기는 '_'까지 초록으로 묶는다.
    for m in re.finditer(r"원코드_?", s):
        spans.append((m.start(), m.end(), GREEN, 20))
    for m in re.finditer(r"셔터", s):
        spans.append((m.start(), m.end(), RED, 20))
    # 롤/콤비 방염/필증 표기는 제품명 셀에서 눈에 띄도록 빨간색으로 표시한다.
    for m in re.finditer(r"방염", s):
        spans.append((m.start(), m.end(), RED, 25))
    for m in re.finditer(r"<필증>", s):
        spans.append((m.start(), m.end(), RED, 25))

    if not spans:
        return s

    # 겹치면 P/FP 코드 전체(우선순위 30)를 우선한다.
    spans.sort(key=lambda x: (x[0], -x[3], -(x[1] - x[0])))
    accepted = []
    for a, b, color, priority in spans:
        overlap = False
        for aa, bb, _, pp in accepted:
            if not (b <= aa or a >= bb):
                overlap = True
                if priority > pp:
                    accepted = [x for x in accepted if (b <= x[0] or a >= x[1])]
                    overlap = False
                break
        if not overlap:
            accepted.append((a, b, color, priority))

    accepted.sort(key=lambda x: (x[0], x[1]))
    parts = []
    cursor = 0
    for a, b, color, _ in accepted:
        if a < cursor:
            continue
        if a > cursor:
            parts.append(_tb(s[cursor:a], BLACK))
        parts.append(_tb(s[a:b], color, sz=11 if s[a:b] == "<필증>" else 14))
        cursor = b
    if cursor < len(s):
        parts.append(_tb(s[cursor:], BLACK))
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
    """좌/우 개수 또는 `좌 우` 의미를 실제 창 단위 item으로 펼친다."""
    if not order or not order.get("items"):
        return order
    client = order.get("거래처")
    prepared = []
    for it in order.get("items") or []:
        if it.get("_product_group") == "holding" or it.get("_holding_accessory"):
            prepared.append(it)
            continue
        def _n(v):
            try:
                return max(0, int(float(v))) if v not in (None, "") else 0
            except (TypeError, ValueError):
                return 0
        left_n, right_n = _n(it.get("좌개수")), _n(it.get("우개수"))
        # 수량과 좌/우 합계가 다를 때는 어느 쪽이 맞는지 프로그램이 임의로 고치지 않는다.
        # 원값을 보존하고 검증 단계에서 반드시 확인 대상으로 올린다.
        declared = _n(it.get("창개수"))
        if declared > 0 and left_n + right_n > 0 and declared != left_n + right_n:
            it["_count_lr_mismatch"] = (declared, left_n, right_n)
        else:
            it.pop("_count_lr_mismatch", None)
        raw_dir = re.sub(r"\s+", "", str(it.get("손잡이방향") or ""))
        if left_n + right_n == 0 and re.fullmatch(r"[좌우]{2,10}", raw_dir):
            it["좌개수"] = raw_dir.count("좌")
            it["우개수"] = raw_dir.count("우")
            it["창개수"] = len(raw_dir)
            it["손잡이방향"] = None
        prepared.append(it)
    order["items"] = prepared
    return _split_integer_windows(order)

def _di_recipient_key(it):
    return str(it.get("_DIrecipient_context") or it.get("기재사항") or "").strip()

def _di_item_is_urgent(it):
    """DI 한 행이 긴급 건인지 최종 편집값까지 포함해 판정한다."""
    return bool(it.get("_DI긴급")) or "긴급" in " ".join(
        str(it.get(k) or "") for k in ("_수동특이", "기재사항", "원문"))

def _propagate_di_urgent_order(order):
    """DI 긴급은 특이 칸이 아니라 기재사항1/정렬용 플래그로만 보존한다."""
    if order.get("거래처") != "DI":
        return order
    groups = {}
    for it in order.get("items") or []:
        # 원문·기재사항·수동특이 어디에서 왔든 긴급 여부만 보존한다.
        raw = " ".join(str(it.get(k) or "") for k in ("_수동특이", "기재사항", "원문"))
        urgent = "긴급" in raw
        if urgent:
            it["_DI긴급"] = True
        # DI 특이사항은 출력하지 않으므로 수동특이에서도 긴급/구분선을 제거한다.
        special = str(it.get("_수동특이") or "")
        special = re.sub(r"\b긴급\b\s*[-–—_=~.·ㆍ]*", " ", special)
        special = re.sub(r"\s*/\s*", "/", special).strip(" /-–—_=~.·ㆍ")
        it["_수동특이"] = special or None
        groups.setdefault(_di_recipient_key(it), []).append(it)
    # 같은 수령인 묶음 중 한 행이라도 긴급이면 묶음 전체를 긴급 주문으로 취급한다.
    for _recipient, items in groups.items():
        if any(it.get("_DI긴급") or _di_item_is_urgent(it) for it in items):
            for it in items:
                it["_DI긴급"] = True
    return order

def _mix_group_key(it):
    """같은 MIX 묶음(구성행을 한 번만 붙이는 같은 규격 창들)을 판별하는 키."""
    return (it.get("_mix_combo"),
            tuple((p.get("코드"), p.get("길이")) for p in it.get("_mix_parts") or []),
            it.get("세로"), it.get("종류"), it.get("타입"))


def _prepare_blind_mix_items(order):
    """비대일 블라인드 MIX 품목의 섞인 코드와 세로 비례 길이를 계산한다(대일은 기존 믹스 형식 유지)."""
    if order.get("거래처") == "DI" or order.get("_product_mode") in {"holding", "roll_combo"}:
        return
    for it in order.get("items") or []:
        if (it.get("_product_group") in {"holding", "roll_combo"} or it.get("_holding_accessory")
                or it.get("예외품목")):
            continue
        codes_text = it.get("_mix_codes") if str(it.get("_mix_name") or "").upper() == "MIX" else None
        source = it.get("_mix_source")
        if source and codes_text and parse_mix_parts(source, None) and \
                "+".join(p["코드"] for p in parse_mix_parts(source, None)) != str(codes_text).upper():
            source = None  # 사용자가 품목 칸에서 조합을 바꾼 경우 예전 비율은 쓰지 않는다.
        sources = [source, codes_text, it.get("색상원문"), it.get("품목코드"), it.get("원문"), it.get("기재사항")]
        joined = " ".join(str(x or "") for x in sources)
        is_mix = bool(it.get("_mix_word") or it.get("_blind_mix") or it.get("_mix_parts_manual")
                      or str(it.get("_mix_name") or "").upper() == "MIX" or MIX_WORD_RX.search(joined)
                      or MIX_LAYOUT_RX.search(joined))
        if not is_mix and order.get("거래처") == "JL" and not it.get("_mix_name"):
            # JL: 발주서 같은 칸(색상)에 여러 색 코드가 있으면 MIX 로 본다.
            jl_codes = list(dict.fromkeys(
                c.upper() for c in re.findall(r"(?<!\d)(\d{3,4}(?:FP|P)?)(?!\d)",
                                              str(it.get("색상원문") or ""), re.I)))
            if len(jl_codes) >= 2:
                sources.insert(0, "+".join(jl_codes))
                is_mix = True
        if not is_mix:
            continue
        parts = None
        if it.get("_mix_parts_manual") and len(it.get("_mix_parts") or []) >= 2:
            parts = [dict(p) for p in it["_mix_parts"]]
        else:
            for src in sources:
                parts = parse_mix_parts(src, it.get("세로"))
                if parts:
                    if src is not codes_text:
                        it["_mix_source"] = str(src)
                    break
        if not parts:
            it["_mix_unresolved"] = True
            it.pop("_blind_mix", None)
            continue
        it.pop("_mix_unresolved", None)
        combo = "+".join(p["코드"] for p in parts)
        it.update({"_blind_mix": True, "_mix_parts": parts, "_mix_combo": combo,
                   "_mix_name": "MIX", "_mix_codes": combo, "_generic_mix": True})
        if not it.get("종류"):
            it["종류"] = "투코드"
        if not it.get("품목코드"):
            it["품목코드"] = parts[0]["코드"]


_VINYL_ORDER_FIELDS = ("전체기재사항", "고객명", "_추가부속")
_VINYL_DELIVERY_FIELDS = ("발신",)
_VINYL_ITEM_FIELDS = ("기재사항", "설치장소", "_수동특이", "_manual_note1", "_manual_note2")


def _strip_vinyl_fields(order):
    """주문의 출력 대상 문구에서 비닐 포장 관련 칸을 뺀다(사전점검 직접 입력값 포함)."""
    for key in _VINYL_ORDER_FIELDS:
        if order.get(key) is not None:
            order[key] = strip_vinyl_text(order[key])
    delivery = order.get("배송") or {}
    for key in _VINYL_DELIVERY_FIELDS:
        if delivery.get(key) is not None:
            delivery[key] = strip_vinyl_text(delivery[key])
    for it in order.get("items") or []:
        for key in _VINYL_ITEM_FIELDS:
            if it.get(key) is not None:
                it[key] = strip_vinyl_text(it[key])


def synchronize_order_for_outputs(order):
    """사전점검에서 확정된 주문을 세 출력물이 동일하게 쓰도록 마지막 공통 정규화."""
    if not order:
        return order
    _strip_vinyl_fields(order)
    delivery = order.setdefault("배송", {})
    delivery["전달사항"] = clean_delivery_notice(delivery.get("전달사항"))
    # 선불/착불 정보가 없으면 택배·화물은 기본적으로 착불 처리한다.
    # 명시된 선불/착불 값은 그대로 우선한다.
    if not str(delivery.get("선불착불") or "").strip():
        mode = _effective_delivery_mode(order)
        if mode in {"택배", "화물"}:
            delivery["선불착불"] = "착불"
    expand_same_size_directions(order)
    _propagate_di_urgent_order(order)
    _prepare_blind_mix_items(order)
    # 누락된 손잡이길이는 원문/기재에서 결정 가능한 숫자만 보완한다.
    # 단, 사전점검에서 사용자가 길이를 직접 수정(또는 빈칸으로 삭제)한 경우에는
    # 그 값을 최종값으로 보고 다시 자동 추출하지 않는다.
    for it in order.get("items") or []:
        if "_manual_handle_length" in it:
            it["손잡이길이"] = it.get("_manual_handle_length")
            continue
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


def _clean_pay_words(text):
    """주소/화물지점 문자열에서 선불·착불·후불 같은 결제 문구를 제거한다."""
    text = str(text or "")
    text = re.sub(r"(?:택배비|배송비|화물비)?\s*(?:선불|착불|후불)", " ", text)
    return re.sub(r"\s+", " ", text).strip(" ,/-")


def _short_city_address(value):
    """장부/사전점검 택배 주소를 현장 표기 규칙으로 줄인다."""
    text = _clean_pay_words(value)
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"^(?:경기도|강원(?:특별자치)?도|충청북도|충청남도|충북|충남|전라북도|전라남도|전북(?:특별자치도)?|전남|경상북도|경상남도|경북|경남|제주(?:특별자치)?도)\s+", "", text)
    city_map = {
        "서울특별시":"서울시", "부산광역시":"부산시", "대구광역시":"대구시",
        "인천광역시":"인천시", "광주광역시":"광주시", "대전광역시":"대전시",
        "울산광역시":"울산시", "세종특별자치시":"세종시",
    }
    for src, dst in city_map.items():
        if text.startswith(src):
            text = dst + text[len(src):]
            break
    text = re.sub(r"\s*\([^()]*\)\s*", " ", text)
    text = re.sub(r"\s+", " ", text).strip(" ,")
    m = re.search(r"(?<![\d-])(\d+\s*동\s*\d+\s*호)\s*$", text)
    if m and m.start() > 0:
        before = text[:m.start()].rstrip(" ,")
        detail = re.sub(r"\s+", "", m.group(1))
        text = f"{before}, {detail}"
    return text


def _freight_ledger_text(delivery, phone_sep=" / "):
    """화물 장부 주소행: `화물지점 - 받는사람 / 전화번호`.

    보노 계열은 전화번호 앞에 `/`를 쓰지 않는다(phone_sep=" ").

    도로명 주소가 함께 들어 있어도 화물은 반드시 `화물지점`을 우선한다.
    """
    d = delivery or {}
    branch = _clean_pay_words(d.get("화물지점") or d.get("주소"))
    receiver = str(d.get("수령인") or "").strip()
    phone = str(d.get("연락처") or "").strip()
    left = branch
    if receiver and receiver not in left:
        left = f"{left} - {receiver}" if left else receiver
    if phone:
        return f"{left}{phone_sep}{phone}" if left else phone
    return left


def _erp_ledger_handle(item):
    """전산 적요에 넣을 손잡이길이 `손120`.

    장부 길이 칸에 적히는 값과 같다: 기본 길이(종류·세로로 정해지는 값)는 적지 않고,
    사전점검에서 사람이 직접 고친 길이는 기본값과 같아도 적는다(2026-09-18 사용자 확정).
    """
    if "_manual_handle_length" in item:
        n = normalize_handle(item.get("_manual_handle_length"))[0]
    else:
        n = _handle(item.get("종류"), item.get("세로"), item.get("손잡이길이"))
    return f"손{n}" if n else None


def _erp_memo_join(parts):
    """경영박사 적요: `가로*세로/1EA` 다음은 슬러시 없이 한 칸 띄우고 나머지를 `/`로 잇는다."""
    head, rest = (parts[0], parts[1:]) if parts else ("", [])
    return f"{head} {'/'.join(rest)}" if rest else head


def parse_freight_ledger_text(value):
    """사전점검에서 수정한 화물 주소 한 줄을 구조화 배송정보로 되돌린다."""
    text = _clean_pay_words(value)
    if not text:
        return None, None, None
    phone = None
    pm = re.search(r"((?:01[016789]|0\d{1,2})[-.\s]?\d{3,4}[-.\s]?\d{4}|\d{8,11})\s*$", text)
    if pm:
        phone = re.sub(r"\s+", "", pm.group(1)).strip(" /")
        text = text[:pm.start()].rstrip(" /,-")
    receiver = None
    branch = text
    if " - " in text:
        branch, receiver = [x.strip() or None for x in text.split(" - ", 1)]
    elif "-" in text:
        a, b = text.split("-", 1)
        if a.strip() and b.strip():
            branch, receiver = a.strip(), b.strip()
    return branch or None, receiver or None, phone or None


def _short_bono_branch(value):
    """보노/미래가공 EDI용 지점명. `안양 박달`→`박달`, `안산초지점`→`초지`."""
    text = str(value or "").strip()
    text = re.sub(r"(?:대신화물|경동화물|경동택배)", " ", text)
    text = re.sub(r"\s+지점\s*$", "", text)
    text = re.sub(r"점\s*$", "", text).strip(" -/,()")
    parts = [x for x in re.split(r"[\s/,-]+", text) if x]
    if len(parts) >= 2:
        return parts[-1]
    token = parts[0] if parts else text
    # 붙여 쓴 `안산초지`, `안양박달`도 현장명만 남긴다.
    city_prefixes = (
        "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종",
        "수원", "성남", "안양", "안산", "고양", "용인", "부천", "화성",
        "평택", "시흥", "광명", "김포", "군포", "의왕", "하남", "남양주",
        "의정부", "양주", "파주", "포천", "춘천", "원주", "청주", "충주",
        "천안", "아산", "전주", "익산", "군산", "목포", "순천", "여수",
        "포항", "경주", "구미", "김천", "창원", "김해", "양산", "진주",
    )
    for city in city_prefixes:
        if token.startswith(city) and len(token) > len(city):
            return token[len(city):]
    return token


def _bono_erp_identity(order):
    """보노/미래가공 EDI 적요 첫 항목.

    받는사람이 거래처 본인이 아니면 받는사람, 본인이면 화물 지점의 현장명만 사용한다.
    """
    delivery = order.get("배송") or {}
    if order.get("_bono_receiver_is_bono"):
        return _short_bono_branch(order.get("_bono_branch") or delivery.get("화물지점"))
    return str(delivery.get("수령인") or "").strip()


def _split_integer_windows(order):
    """일반 블라인드의 창개수/정수 수량 N을 실제 N개 창(item)으로 펼친다."""
    out = []
    for it in list(order.get("items") or []):
        # 홀딩도어 본품도 실제 창 수만큼 한 줄씩 펼친다.
        # 예: 같은 규격 2창이면 수량 2 한 줄이 아니라 수량 1인 두 줄.
        # 부속품은 EA 수량을 유지하므로 펼치지 않는다.
        if it.get("_holding_accessory"):
            out.append(it)
            continue
        if it.get("_product_group") == "holding":
            def _hn(v):
                try:
                    return max(0, int(float(v))) if v not in (None, "") else 0
                except (TypeError, ValueError):
                    return 0
            count = _hn(it.get("창개수"))
            if count <= 1 and not looks_like_holding_operation(it.get("수량")):
                count = _hn(str(it.get("수량") or "").replace("X", "").replace("x", ""))
            if count > 1:
                for _ in range(count):
                    clone = dict(it)
                    clone["창개수"] = 1
                    # 작동방식이 수량 칸에 들어온 원본은 그대로 보존하고, 숫자 수량만 1로 바꾼다.
                    if not looks_like_holding_operation(clone.get("수량")):
                        clone["수량"] = 1
                    clone["_expanded_window"] = True
                    out.append(clone)
            else:
                out.append(it)
            continue
        split_n = handle_split(it.get("수량"))
        if split_n:
            if order.get("거래처") == "SP":
                # 스페이스 수량 `1/2`는 손잡이 분할이 아니라 **가로를 2등분**한다는 뜻이다.
                # 가로 200 · 1/2 -> 가로 100짜리 2창. 청구도 나눈 창 각각 계산한다(2026-09-18).
                width = _as_float(it.get("가로"))
                if width:
                    part = width / split_n
                    it["가로"] = int(part) if float(part).is_integer() else round(part, 1)
                it["수량"] = None
                def _sn(v):
                    try:
                        return max(0, int(float(v))) if v not in (None, "") else 0
                    except (TypeError, ValueError):
                        return 0
                if _sn(it.get("좌개수")) + _sn(it.get("우개수")) != split_n:
                    it["창개수"] = split_n
                # 아래 공통 로직이 창개수/좌우 기준으로 실제 창을 펼친다.
            else:
                out.append(it)
                continue
        def _n(v):
            try:
                return max(0, int(float(v))) if v not in (None, "") else 0
            except (TypeError, ValueError):
                return 0
        left_n, right_n = _n(it.get("좌개수")), _n(it.get("우개수"))
        declared = _n(it.get("창개수"))
        # 수량/좌우가 서로 모순되면 창을 임의 증식/축소하지 않는다.
        # 사전점검의 오류 상태에서 사람이 원본을 보고 수정하도록 원행을 그대로 둔다.
        if declared > 0 and left_n + right_n > 0 and declared != left_n + right_n:
            it["_count_lr_mismatch"] = (declared, left_n, right_n)
            out.append(it)
            continue
        if left_n + right_n > 1:
            for direction, count in (("좌", left_n), ("우", right_n)):
                for _ in range(count):
                    clone = dict(it)
                    clone["손잡이방향"] = direction
                    clone["수량"] = 1
                    clone["창개수"] = 1
                    clone["좌개수"] = 1 if direction == "좌" else 0
                    clone["우개수"] = 1 if direction == "우" else 0
                    clone["_expanded_window"] = True
                    out.append(clone)
            continue
        count = _n(it.get("창개수"))
        if count <= 1:
            raw_qty = it.get("수량")
            if isinstance(raw_qty, (int, float)) or (isinstance(raw_qty, str) and re.fullmatch(r"\s*\d+(?:\.0+)?\s*", raw_qty)):
                count = _n(raw_qty)
        if count > 1:
            direction = str(it.get("손잡이방향") or "").strip() or None
            for _ in range(count):
                clone = dict(it)
                clone["수량"] = 1
                clone["창개수"] = 1
                clone["좌개수"] = 1 if direction == "좌" else 0
                clone["우개수"] = 1 if direction == "우" else 0
                clone["_expanded_window"] = True
                out.append(clone)
            continue
        out.append(it)
    order["items"] = out
    return order


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
    """전 업체 공통: 배송이 택배/화물이면 창당 포장비를 부과한다.

    PySide6 사전점검에서 사용자가 포장행을 직접 수정한 경우에는
    `_manual_packing_enabled` 값을 최우선으로 사용한다. 자동 규칙과 사람의
    최종 확인값이 서로 다시 덮어쓰지 않게 하기 위한 명시적 override다.
    """
    if "_manual_packing_enabled" in order:
        return bool(order.get("_manual_packing_enabled"))
    return _effective_delivery_mode(order) in {"택배", "화물"}


def _prepend_note_part(note, part):
    """기재사항1 맨 앞에 값을 중복 없이 넣는다. 피스가 있으면 `피스/이름` 순서(2026-09-14)."""
    part = str(part or "").strip()
    if not part:
        return note
    parts = [x.strip() for x in str(note or "").split("/") if x.strip()]
    parts = [x for x in parts if x != part]
    pieces = [x for x in parts if x.startswith("피스")]
    rest = [x for x in parts if not x.startswith("피스")]
    return "/".join(pieces + [part] + rest) or None


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
    return bool(MIX_WORD_RX.search(raw))


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


def _note_key(value):
    """기재사항 중복 비교용 키. 공백 차이 때문에 같은 문구가 반복되는 것을 막는다."""
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def _unique_note_parts(*values):
    out, seen = [], set()
    for part in _split_note_parts(*values):
        key = _note_key(part)
        if key and key not in seen:
            out.append(part)
            seen.add(key)
    return out


def _remove_note_parts(parts, excluded):
    blocked = {_note_key(x) for x in excluded if str(x or "").strip()}
    return [x for x in parts if _note_key(x) not in blocked]


def _order_window_count(items):
    """행 수가 아니라 실제 창 수 기준. 수량 3인 한 행도 3창으로 본다."""
    total = 0
    for it in items or []:
        if it.get("예외품목") or it.get("_holding_accessory"):
            continue
        if it.get("_product_group") == "holding":
            total += _holding_count(it)
            continue
        _l, _r, count = _blind_counts(it)
        # 창개수/좌우개수가 따로 없고 수량만 숫자로 적힌 발주도 실제 창 수로 본다.
        # (출력 행을 다시 증식시키는 용도가 아니라 기재사항 분배 기준 계산에만 사용)
        if count <= 1 and not any(it.get(k) not in (None, "") for k in ("창개수", "좌개수", "우개수")):
            raw_qty = it.get("수량")
            try:
                qn = int(float(raw_qty)) if raw_qty not in (None, "") else 1
                count = max(count, qn)
            except (TypeError, ValueError):
                pass
        total += max(1, count)
    return max(0, total)


def _common_and_personal_note_parts(order, items, body=None, customer=None):
    """공통 기재 / 개인창 기재를 분리한다.

    전체기재사항·고객명은 공통으로 보고, 모든 창의 `기재사항`에 반복해서
    들어간 같은 문구도 공통으로 승격한다. 설치장소는 창별 정보로 유지한다.
    이 과정에서 대동산업처럼 `공학1관/공학1관`이 되는 중복을 제거한다.
    """
    body = order.get("전체기재사항") if body is None else body
    customer = order.get("고객명") if customer is None else customer
    common = _unique_note_parts(body, customer)

    item_note_lists = []
    for it in items or []:
        if it.get("예외품목") or it.get("_holding_accessory"):
            continue
        parts = _unique_note_parts(it.get("기재사항"))
        if parts:
            item_note_lists.append(parts)
    if len(item_note_lists) >= 2:
        intersection = {_note_key(x) for x in item_note_lists[0]}
        for parts in item_note_lists[1:]:
            intersection &= {_note_key(x) for x in parts}
        for part in item_note_lists[0]:
            if _note_key(part) in intersection and _note_key(part) not in {_note_key(x) for x in common}:
                common.append(part)

    personal = {}
    for it in items or []:
        parts = _unique_note_parts(it.get("기재사항"), it.get("설치장소"))
        personal[id(it)] = _remove_note_parts(parts, common)
    return common, personal


def _jl_clean_parts(parts, customer=None, order_no=None, keep_piece=False):
    """JL 메모에서 포장 문구/중복 이름/DW/주문번호를 제거한다."""
    out = []
    customer = _jl_customer_name(customer)
    order_no = re.sub(r"^[（(]\s*|\s*[）)]$", "", str(order_no or "").strip())
    for part in parts:
        text = str(part or "").strip()
        if not text or JL_PACKAGING.search(text):
            continue
        # 거래처 담당자 이름(예: JL 김현경)은 주문 정보가 아니므로 적지 않는다.
        if is_client_staff_name("JL", text):
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


def _holding_ledger_operation(value):
    # 홀딩 작동방식은 세 출력물 모두 같은 짧은 표기를 사용한다:
    # 편개=편, 양개=양, 이등분=1/2, 양자석=양자석.
    return normalize_holding_operation(value)


def _holding_display_count(it):
    """홀딩도어 장부 G열에 적을 실제 개수. 작동방식 분수는 수량이 아니다."""
    raw = it.get("수량")
    if raw not in (None, "") and not looks_like_holding_operation(raw):
        try:
            n = float(raw)
            return None if n == 1 else (int(n) if n.is_integer() else n)
        except (TypeError, ValueError):
            return raw
    count = it.get("창개수")
    try:
        n = int(float(count)) if count not in (None, "") else 1
    except (TypeError, ValueError):
        n = 1
    return n if n > 1 else None


def _as_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _dimension_attention_flags(item, is_roll=False, is_holding=False):
    """장부/사전점검 치수 파란색 표시 여부.

    - 홀딩도어: 세로 300cm부터 세로 파랑
    - 일반 블라인드: 원코드 가로 36cm 미만, 투코드/셔터 30cm 미만 가로 파랑
      그리고 세로 370cm 초과는 세로 파랑
    - 롤/콤비는 별도 치수 파랑 규칙을 적용하지 않는다.
    """
    w = _as_float(item.get("가로"))
    h = _as_float(item.get("세로"))
    if is_holding:
        return False, bool(h is not None and h >= 300)
    if is_roll:
        return False, False
    kind = re.sub(r"\s+", "", str(item.get("종류") or ""))
    blue_w = bool(w is not None and (
        ("원코드" in kind and w < 36) or
        (("투코드" in kind or "셔터" in kind) and w < 30)
    ))
    blue_h = bool(h is not None and h > 370)
    return blue_w, blue_h


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
    is_roll_order = order.get("_product_mode") == "roll_combo"
    # 롤/콤비 장부도 상호 옆 내부표시 (K)를 사용한다.
    mark = "(K)" if is_roll_order else CLIENT_INFO.get(client, (None, None, None, ""))[3]
    order_no = str(order.get("주문번호") or "").strip()
    if client == "SP" and order_no:
        # 스페이스 장부 상호 칸은 주문번호 끝 숫자를 괄호로 묶어 적는다(`C-9` -> `(9)`).
        m = re.search(r"(?:^|-)\s*(\d+)\s*$", order_no)
        mark = f"({m.group(1)})" if m else ""
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
        if is_client_staff_name(client, cust):
            cust = ""
    head = "/".join(x for x in (body, cust) if x)
    window_count = _order_window_count(items)
    single = window_count <= 1 and client not in NO_MERGE_PLACE
    generic_common_parts, generic_personal_parts = _common_and_personal_note_parts(
        order, items, body=body, customer=cust)
    for i, it in enumerate(items):
        is_holding = it.get("_product_group") == "holding" or it.get("_holding_accessory")
        is_accessory = bool(it.get("_holding_accessory"))
        is_roll = it.get("_product_group") == "roll_combo" or is_roll_order
        if is_roll:
            color = it.get("_ledger_text_manual") if "_ledger_text_manual" in it else it.get("_ledger_text")
            color = color or str(it.get("롤품명") or it.get("품명") or "미확인")
        elif is_holding:
            color = holding_ledger_text(it)
        else:
            color = ledger_color(it.get("품목코드"), it.get("종류"), it.get("타입"),
                                 it.get("_ledger_prefix", "B"),
                                 it.get("_mix_name"), it.get("_mix_codes"))
            if _has_non_di_mix(it, client):
                color = _non_di_mix_ledger_text(it, color)
            if it.get("_blind_mix") and client != "DI":
                # 비대일 MIX 품목 칸: `B Mix 200+125`. 4가지 이상 배합은 `B 원코드 Mix`만 적고
                # 조합(280+160+280+160)은 첫 구성행 색상 칸에 적는다.
                color = ledger_color("Mix", it.get("종류") or "투코드", it.get("타입"),
                                     it.get("_ledger_prefix", "B"))
                if len(it.get("_mix_parts") or []) < 4:
                    color = f"{color} {it.get('_mix_combo')}"
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
        if is_roll:
            # 닫힌창/닫힌사이즈는 롤·콤비 기재사항으로 쓰지 않는다.
            note = "/".join(x.strip() for x in str(note).split("/")
                            if x.strip() and not _roll_note_ignored(x))
            if place and _roll_note_ignored(place):
                place = None
        if client == "DI" and it.get("_generic_mix") and it.get("_mix_codes"):
            # 미등록 조합 MIX는 품목 칸에는 B MIX/B 원코드 MIX만 쓰고
            # 실제 조합은 기재사항의 맨 앞에 남긴다.
            combo = str(it.get("_mix_codes") or "").replace(" ", "")
            note_parts = [x.strip() for x in str(note).split("/") if x.strip()]
            note = "/".join([combo] + [x for x in note_parts if x != combo])
        if is_roll:
            # 롤/콤비는 거래처 공통 형식. 안산)보노는 현재 정답 장부의
            # 오더명/시공위치 구조를 그대로 사용한다. 연창은 별도로 만들지 않는다.
            if client == "안산)보노":
                note1 = place or None
                note2 = order_no or None
            else:
                common_roll = _unique_note_parts(order_no, body, cust)
                personal_roll = _unique_note_parts(place, note)
                if window_count <= 1:
                    note1 = "/".join(_unique_note_parts(*common_roll, *personal_roll)) or None
                    note2 = None
                else:
                    note1 = "/".join(common_roll) or None
                    note2 = "/".join(personal_roll) or None
        elif client == "DI":
            # DI는 기재사항1만 사용하고 특이 칸은 항상 비운다.
            # '긴급 ---' 등은 긴급 한 단어만 기재사항1에 1회 남긴다.
            mix_detail = None
            if it.get("_mix_codes"):
                combo = str(it.get("_mix_codes") or "").replace(" ", "")
                mix_name = str(it.get("_mix_name") or "").strip()
                mix_detail = (combo if it.get("_generic_mix") or mix_name.upper() == "MIX"
                              else f"{mix_name}{combo}" if mix_name else combo)
            urgent = "긴급" if _di_item_is_urgent(it) else None
            note1 = "/".join(_unique_note_parts(urgent, di_recipient, mix_detail)) or None
            note2 = None
        elif client == "휴안":
            # 휴안도 기재사항1 한 칸만 사용한다. 기존의 창별 메모/부속/설치장소는
            # 잃지 않도록 한 칸에 합치고, 수령인/고객명은 마지막에 한 번만 둔다.
            receiver = str((order.get("배송") or {}).get("수령인") or cust or "").strip()
            hu_parts = _unique_note_parts(body, note, place)
            hu_parts = [clean_delivery_notice(x) for x in hu_parts]
            hu_parts = [x for x in hu_parts if x]
            if receiver and _note_key(receiver) not in {_note_key(x) for x in hu_parts}:
                hu_parts.append(receiver)
            note1 = "/".join(hu_parts) or None
            note2 = None
        elif client == "RT":
            # 루임트는 기존 주소/받는명칭 규칙을 그대로 유지한다.
            if single:
                note1 = "/".join(_unique_note_parts(head, note, place)) or None
                note2 = None
            else:
                note1 = "/".join(_unique_note_parts(head, note)) if i == 0 else (note or None)
                note2 = place
        elif client == "JO":
            # 제이원: 기재사항1은 창 수와 관계없이 주문번호만 쓴다.
            jo_other = _jo_note2_parts(order, it, cust, first=(i == 0))
            note1 = order_no or None
            note2 = ("/".join(_unique_note_parts(*jo_other)) or None) if window_count >= 3 else None
        elif client == "JL":
            # JL 공통기재는 모든 창에 같은 값으로 만든 뒤 화면/Excel에서 병합한다.
            all_note_parts = _split_note_parts(common, *(x.get("기재사항") for x in items))
            has_piece = any("피스" in x for x in all_note_parts)
            common_clean = _jl_clean_parts(_split_note_parts(common), cust, order_no)
            common_clean = [x for x in common_clean if "피스" not in x]
            ordered = []
            for part in ((["피스"] if has_piece else []) + ([cust] if cust else []) + common_clean):
                if part and part not in ordered:
                    ordered.append(part)
            note1 = "/".join(ordered) or None
            item_parts = [x for x in _jl_clean_parts(_split_note_parts(note), cust, order_no)
                          if "피스" not in x and x not in ordered]
            note2 = "/".join(_unique_note_parts(place, *item_parts)) or None
        else:
            # 공통 규칙: 여러 창이면 기재사항1=주문 공통정보,
            # 기재사항2=각 개인창 정보. 한 창이면 두 정보를 기재사항1에 합친다.
            common_parts = list(generic_common_parts)
            personal_parts = list(generic_personal_parts.get(id(it), []))

            # 업체별로 이미 확정된 공통/개인 정보의 의미는 유지한다.
            if client == "SP":
                common_parts = _unique_note_parts(
                    "/".join(sp_notices), order_no, body, cust, *common_parts)
            elif client in BONO_STYLE_CLIENTS:
                # 보노/미래가공 장부에는 받는사람/지점명을 넣지 않는다.
                # 장부는 오더명 + 창별 시공위치만 확정값으로 사용한다.
                common_parts = _unique_note_parts(order_no)
                personal_parts = _unique_note_parts(place, *personal_parts)
            elif client == "DU":
                common_parts = _unique_note_parts(cust or (order.get("배송") or {}).get("수령인"))
                personal_parts = _unique_note_parts(body, note, place)

            if single:
                note1 = "/".join(_unique_note_parts(*common_parts, *personal_parts)) or None
                note2 = None
            else:
                note1 = "/".join(_unique_note_parts(*common_parts)) or None
                note2 = "/".join(_unique_note_parts(*personal_parts)) or None
        # 택배/화물은 받는 명칭을 기재사항1에서 확인할 수 있게 한다.
        # JO/JL/DU는 별도 고정 규칙이 있으므로 해당 규칙을 우선한다.
        # 휴안은 위에서 이름을 맨 뒤에 두었다(추가부속 `석고앙카3/이름`, 2026-09-14). 앞으로 옮기지 않는다.
        if (not is_roll) and (not is_holding) and client not in ({"JO", "JL", "DU", "휴안"} | BONO_STYLE_CLIENTS) and _effective_delivery_mode(order) in {"택배", "화물"}:
            recv = str((order.get("배송") or {}).get("수령인") or "").strip()
            if client == "RT" and recv == "더커튼":
                recv = "더"   # 루임트: 기재사항에는 `더`, 주소에는 `더커튼`
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
        # 공통 규칙: 기재사항이 하나만 있으면 기재사항1에 두고 1·2 칸을 병합한다.
        if not note1 and note2:
            note1, note2 = note2, None
        # 공통 규칙: 기재사항1·2가 같으면 좌우로 합친다(사용자가 직접 고친 칸은 자동으로 합치지 않음).
        if (note1 and note2 and _note_key(note1) == _note_key(note2)
                and "_manual_note1" not in it and "_manual_note2" not in it):
            note2 = None
        # 홀딩도어 배송 수령인/연락처는 아래의 주소·수령인 전달행에서 따로 보여 준다.
        # 같은 이름이 기재사항1/2에도 반복되는 문제(예: 홀랜드/홀랜드)를 제거하고,
        # 두 기재사항 열 사이의 중복도 함께 제거한다. 제작 메모 자체는 보존한다.
        if is_holding:
            delivery = order.get("배송") or {}
            blocked = [delivery.get("수령인"), delivery.get("연락처")]
            blocked_keys = {_note_key(x) for x in blocked if str(x or "").strip()}

            def _clean_holding_note(value):
                parts = _unique_note_parts(value)
                parts = [x for x in parts if _note_key(x) not in blocked_keys]
                return "/".join(parts) or None

            note1 = _clean_holding_note(note1)
            note2 = _clean_holding_note(note2)
            if note1 and note2:
                note1_keys = {_note_key(x) for x in _unique_note_parts(note1)}
                note2_parts = [x for x in _unique_note_parts(note2) if _note_key(x) not in note1_keys]
                note2 = "/".join(note2_parts) or None
            if not note1 and note2:
                note1, note2 = note2, None
        has_tlean, (note1, note2) = _take_tlean(note1, note2)
        joint = bool((not is_holding) and (not is_roll) and it.get("종류") == "원코드" and it.get("연창"))
        joint_start = joint and (not prev_joint or not same_color)
        manual_special = str(it.get("_수동특이") or "").strip()
        # 보노 계열: 발주서 작동방식필증 칸 내용을 특이사항에 적는다.
        operation_special = (str(it.get("작동방식필증") or "").strip()
                             if client in BONO_LIKE_CLIENTS else "")
        special = " ".join(dict.fromkeys(x for x in (
            "틀안" if has_tlean else None,
            "#" if joint_start else None,
            manual_special or None,
            operation_special or None,
        ) if x)) or None
        if client == "DI":
            # 대일은 자동 판독값으로 특이사항을 채우지 않는다.
            # 다만 사전점검에서 사용자가 직접 입력한 값은 최종값으로 존중한다.
            special = manual_special or None if it.get("_manual_special_override") else None

        if is_accessory:
            raw_qty = it.get("수량") if it.get("수량") not in (None, "") else it.get("창개수")
            aq = accessory_qty(raw_qty)
            qty_display = str(raw_qty).strip() if str(raw_qty or "").strip().lower().startswith("x") else f"X{aq:g}"
            model1 = model2 = None
            width_out = height_out = None
        elif is_roll:
            raw_qty = it.get("수량") if it.get("수량") not in (None, "") else it.get("창개수")
            try:
                qn = int(float(raw_qty)) if raw_qty not in (None, "") else 1
            except (TypeError, ValueError):
                qn = 1
            # 롤/콤비 정답 장부는 1개 창의 수량을 비우고, 2개 이상만 수량 표시.
            qty_display = qn if qn > 1 else None
            direction = it.get("손잡이방향")
            model1 = _short_dir(direction, client)
            raw_len = it.get("_manual_handle_length") if "_manual_handle_length" in it else it.get("손잡이길이")
            # 롤/콤비 손잡이길이는 5단위 내림(2026-09-14).
            hn = normalize_roll_handle(raw_len)[0]
            # 롤/콤비/트리플은 기본 손잡이 150을 장부/작업지시서에 쓰지 않는다.
            # 별도 길이가 150이 아닌 경우만 숫자로 표시한다.
            model2 = hn if hn is not None and hn != 150 else None
            width_out, height_out = _display_size(it.get("가로")), ('  "' if ditto else _display_size(h))
        elif is_holding:
            # 홀딩도어만 수량/방향 열을 반대로 사용한다.
            # G(수량)에는 작동방식(편/양/1/2/양자석), H(방향)에는 실제 개수를 적는다.
            # 다만 홀딩 본품은 창 단위로 펼치므로 1개는 공통 규칙에 따라 공란이다.
            operation = _holding_ledger_operation(
                it.get("_holding_operation") or
                (it.get("수량") if looks_like_holding_operation(it.get("수량")) else None))
            count_display = _holding_display_count(it)
            qty_display = operation
            model1 = count_display
            model2 = str(it.get("_holding_rail") or "").strip() or None
            width_out, height_out = _display_size(it.get("가로")), ('  \"' if ditto else _display_size(h))
        else:
            # 실제 창 단위로 펼친 뒤에는 일반 블라인드 행도 수량을 비워 두지 않는다.
            # 원 발주가 수량 2이면 두 행으로 분리되고 각 행은 수량 1로 표시된다.
            qty_source = it.get("수량")
            if qty_source in (None, ""):
                qty_source = it.get("창개수") if it.get("창개수") not in (None, "") else 1
            qty_display = normalized_ledger_qty(qty_source)
            # 공통 규칙: 가로·세로가 있는 제작제품은 수량 1을 장부에 쓰지 않는다.
            # 부속품은 별도 분기에서 수량 1도 유지한다.
            try:
                if it.get("가로") not in (None, "") and it.get("세로") not in (None, "") \
                        and float(qty_source) == 1:
                    qty_display = None
            except (TypeError, ValueError):
                pass
            direction = it.get("손잡이방향")
            if direction in (None, ""):
                try:
                    ln = int(float(it.get("좌개수") or 0))
                    rn = int(float(it.get("우개수") or 0))
                except (TypeError, ValueError):
                    ln = rn = 0
                if ln == 1 and rn == 0:
                    direction = "좌"
                elif rn == 1 and ln == 0:
                    direction = "우"
            model1 = _short_dir(direction, client)
            if "_manual_handle_length" in it:
                # 사용자가 사전점검에서 직접 고친 길이는 기본길이와 같더라도
                # 숨기지 않고 입력한 값 자체를 장부/작업지시서에 표시한다.
                manual_len = normalize_handle(it.get("_manual_handle_length"))[0]
                model2 = manual_len or None
            else:
                model2 = _handle(it.get("종류"), h, it.get("손잡이길이"))
            width_out, height_out = _display_size(it.get("가로")), ('  \"' if ditto else _display_size(h))

        # 롤/콤비 품명은 내부 문자열 공백을 한 칸으로 정리한 뒤,
        # FM 장부처럼 셀 맨 앞에 공백을 정확히 한 칸 둔다.
        # <필증>은 제품명 셀 안에 넣지 않고 build_ledger 단계에서 별도 행으로 만든다.
        if is_roll and color:
            color = re.sub(r"[ \t]+", " ", str(color)).strip()
        display_color = None if same_color else color
        if is_roll and display_color:
            display_color = " " + str(display_color).lstrip()

        blue_width, blue_height = _dimension_attention_flags(
            it, is_roll=is_roll, is_holding=(is_holding and not is_accessory))

        rows.append({
            "거래처": client if i == 0 else None,
            "내부표시": mark if i == 0 else None,
            "색상": display_color,
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
            "_product_group": "roll_combo" if is_roll else ("holding" if is_holding else "blind"),
            "_roll_fireproof": bool(is_roll and it.get("_roll_fireproof")),
            "_roll_fire_key": (re.sub(r"[ \t]+", " ", str(color or "")).strip() if is_roll else None),
            "_blue_width": blue_width,
            "_blue_height": blue_height,
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
            # 자동 병합은 최초 판독값에만 적용한다. 사용자가 사전점검에서
            # 해당 기재사항 칸을 한 번이라도 수정했다면 값이 같아도 자동 병합 금지.
            "_manual_note1": ("_manual_note1" in it),
            "_manual_note2": ("_manual_note2" in it),
            "_연창": bool(joint),
        })
        if not is_accessory:
            prev_color = color
            if h not in (None, ""):
                prev_h = h if not ditto else prev_h
        prev_joint = joint
        prev_di_recipient = di_recipient
        prev_di_group = di_group
        if it.get("_blind_mix") and client != "DI" and it.get("_mix_parts"):
            # 같은 규격 창을 모두 적은 뒤 섞인 코드 수만큼 구성행을 한 번만 붙인다.
            nxt = items[i + 1] if i + 1 < len(items) else None
            if not (nxt is not None and _mix_group_key(nxt) == _mix_group_key(it)):
                parts = it.get("_mix_parts") or []
                for k, part in enumerate(parts):
                    rows.append({
                        "_특수": "믹스", "코드": part.get("코드"), "길이": part.get("길이"),
                        "조합": it.get("_mix_combo") if (k == 0 and len(parts) >= 4) else None,
                        "_mix_item_index": source_index.get(id(it)), "_mix_part_index": k,
                    })
                # 구성행 다음 제품은 같은 품목이어도 품목명/세로를 다시 적는다(정답 장부 형식).
                prev_color = prev_h = None

    d = order.get("배송") or {}
    if client == "인천)트루":
        for accessory in order.get("_true_accessories") or []:
            label = str(accessory.get("표시") or "").strip()
            if label:
                # 숫자 규칙을 추정하지 않고 사용자가 요청한 1개로 고정한다.
                rows.append({"_특수": "부속", "문구": label, "수량": 1})
    if client == "유앤":
        # 피스 종류는 창 가로별 피스 수를 더해 `10EA`, 노피스브라켓은 `2set` (2026-09-14).
        for line in _unit_accessory_lines(order):
            rows.append({"_특수": "부속", "문구": line["품명"], "수량": line["표시수량"]})
    if client == "휴안":
        # 휴안 추가부속: 색상 칸 `노피스(3)`, 수량 칸 `2EA`의 별도 행.
        for accessory in order.get("_huan_accessories") or []:
            label = str(accessory.get("표시") or "").strip()
            if label:
                rows.append({"_특수": "부속", "문구": label, "수량": accessory.get("수량")})
    if str(order.get("_추가부속") or "").strip():
        rows.append({"_특수": "부속", "문구": str(order.get("_추가부속")).strip(),
                     "수량": None})
    if is_roll_order:
        # 주문서의 제작 공통 메모는 새 장부에서도 잃지 않는다.
        # 연창 기호(#)와 혼동되지 않도록 별도 일반 메모행으로만 보존한다.
        seen_roll_notes = set()
        for raw_note in str(order.get("전체기재사항") or "").split("/"):
            raw_note = raw_note.strip()
            if (raw_note and raw_note != "포장비용" and not _roll_note_ignored(raw_note)
                    and raw_note not in seen_roll_notes):
                rows.append({"_특수": "", "문구": raw_note, "_roll_common_note": True})
                seen_roll_notes.add(raw_note)
    if client == "아지트":
        # 가공소 표시는 배송 주소가 아니므로 작업지시서에도 남긴다(_az_place).
        rows.append({"_특수": "☆", "문구": order.get("_az_place") or "시온가공소", "_az_place": True})
    # 전 업체 공통: 택배/화물 주문은 창당 포장비용을 추가한다.
    # 사람이 작성한 기존 장부에서 `포장비용`이 명시된 경우도 보존한다.
    pack_windows = 0
    for pit in items:
        if pit.get("_holding_accessory") or pit.get("_product_group") == "holding":
            continue
        if (pit.get("_product_group") == "roll_combo" or is_roll_order) and client not in BONO_LIKE_CLIENTS:
            # 롤·콤비 창당 포장비는 보노 계열(안산)보노·미래가공)만 붙인다.
            continue
        try:
            pack_windows += _blind_counts(pit)[2]
        except Exception:
            pack_windows += 1
    if pack_windows and (_needs_window_packing(order) or "포장비용" in common):
        manual_pack_qty = order.get("_manual_packing_qty")
        try:
            manual_pack_qty = int(float(manual_pack_qty)) if manual_pack_qty not in (None, "") else None
        except (TypeError, ValueError):
            manual_pack_qty = None
        pack_qty = max(1, manual_pack_qty) if manual_pack_qty else pack_windows
        rows.append({"_특수": "#", "문구": "포장비용", "수량": pack_qty})
    # 화물은 도로명 주소 없이 화물지점만 있어도 주소행을 만든다(예: 루임트 `오산삼미 - 더커튼`).
    if d.get("방식") in ("택배", "화물", "배달") and (
            d.get("주소") or ("화물" in str(d.get("방식")) and d.get("화물지점"))):
        method = _effective_delivery_mode(order)
        pay_mark = "선불" if str(d.get("선불착불") or "").strip() == "선불" else None
        phone_sep = " " if client in BONO_LIKE_CLIENTS else " / "
        address_text = (_freight_ledger_text(d, phone_sep) if method == "화물"
                        else _short_city_address(d.get("주소")))
        who = (" ".join(x for x in (d.get("수령인"), d.get("연락처")) if x)
               if method != "화물" else "")
        # 선불은 이름·전화번호 행의 마지막 칸(기재사항2)에 적는다. 그 행이 없으면 주소행에 적는다.
        rows.append({"_특수": "☆", "문구": address_text, "선불": None if who else pay_mark})
        if who:
            rows.append({"_특수": "", "문구": who, "선불": pay_mark})
        delivery_notice = clean_delivery_notice(d.get("전달사항")) or ""
        if delivery_notice:
            rows.append({"_특수": "", "문구": f"전달 : {delivery_notice}"})
        if d.get("발신"):
            rows.append({"_특수": "", "문구": f"발신 : {d['발신']}"})
    # 마지막 안전장치: 어떤 경로로 만들어진 장부 문구에도 비닐 포장 문구를 남기지 않는다.
    for row in rows:
        for key in ("기재사항", "기재사항2", "특이", "문구"):
            if row.get(key) is not None:
                row[key] = strip_vinyl_text(row[key])
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


def _safe_merge(ws, start_row, end_row, start_column, end_column):
    """겹치는 mergeCells 레코드를 만들지 않아 Excel 복구 경고를 방지한다."""
    if start_row == end_row and start_column == end_column:
        return True
    for rng in list(ws.merged_cells.ranges):
        if (rng.min_row == start_row and rng.max_row == end_row and
                rng.min_col == start_column and rng.max_col == end_column):
            return True
        overlap = not (end_row < rng.min_row or rng.max_row < start_row
                       or end_column < rng.min_col or rng.max_col < start_column)
        if overlap:
            return False
    ws.merge_cells(start_row=start_row, end_row=end_row,
                   start_column=start_column, end_column=end_column)
    return True


def _is_mix_part(row):
    return row.get("_특수") == "믹스"


def _write_mix_part_row(ws, r, row):
    """MIX 구성행: 규격(D:F)=코드(굵은 테두리), 수량·방향(G:H)=`) 108cm`, 4가지 이상 배합은 첫 행 C열에 조합."""
    if row.get("조합"):
        combo = ws.cell(r, 3, row["조합"])
        combo.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    code = ws.cell(r, 4, row.get("코드"))
    code.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    for col in (4, 5, 6):
        ws.cell(r, col).border = Border(left=THICK if col == 4 else None,
                                        right=THICK if col == 6 else None,
                                        top=THICK, bottom=THICK)
    _safe_merge(ws, r, r, 4, 6)
    length = mix_length_text(row.get("길이"))
    amount = ws.cell(r, 7, f") {length}cm" if length else ")")
    amount.alignment = Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
    _safe_merge(ws, r, r, 7, 8)


def _write_rows(ws, rows, start=3, center_head=False, detailed=False):
    """rows 를 start 행부터 기록. 모형_2·기재사항은 조건에 맞으면 세로 병합"""
    r = start
    span = []
    previous_order_client = None
    for row in rows:
        if row.get("_특수") is not None:
            if _is_mix_part(row):
                _write_mix_part_row(ws, r, row)
                r += 1
                continue
            if row.get("_특수") == "__필증__":
                # 기존에 재사용할 칸이 전혀 없을 때만 생기는 필증 전용 행.
                cert = ws.cell(r, 3, "<필증>")
                cert.font = Font(name=FONT_DATA, size=11, color=RED)
                cert.alignment = Alignment(horizontal="center", vertical="top")
                if row.get("문구"):
                    ws.cell(r, 4, row.get("문구"))
                    ws.cell(r, 4).alignment = Alignment(horizontal="left", vertical="center",
                                                        shrink_to_fit=True)
                    _safe_merge(ws, r, r, 4, 12 if detailed else 11)
                r += 1
                continue
            if row.get("_roll_cert_here"):
                # FM: 이미 존재하는 공통메모 행의 C열을 <필증> 칸으로 재사용한다.
                cert = ws.cell(r, 3, "<필증>")
                cert.font = Font(name=FONT_DATA, size=11, color=RED)
                cert.alignment = Alignment(horizontal="center", vertical="top")
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
                _safe_merge(ws, r, r, 4, pay_col - 1)
                ws.cell(r, pay_col, row["선불"]).font = Font(
                    name=FONT_DATA, size=14, color=RED)
                ws.cell(r, pay_col).alignment = Alignment(horizontal="center",
                                                          vertical="center")
            else:
                _safe_merge(ws, r, r, 4, 12 if detailed else 11)
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
            if row.get("_roll_cert_here"):
                # 같은 방염 품목의 연속 제작행이 이미 있으면 빈 색상 셀을
                # 새 행 대신 FM 필증 칸으로 사용한다.
                c3.value = "<필증>"
                c3.font = Font(name=FONT_DATA, size=11, color=RED)
                c3.alignment = Alignment(horizontal="center", vertical="top")
            if "\n" in str(row.get("색상") or ""):
                ws.row_dimensions[r].height = max(ws.row_dimensions[r].height or ROW_H, 34)
            if center_head:
                c3.alignment = Alignment(horizontal="center",
                                         vertical="center", shrink_to_fit=True)
            c_width = ws.cell(r, 4, row["가로"])
            c_height = ws.cell(r, 6, row["세로"])
            if row.get("_blue_width"):
                f = copy(c_width.font); f.color = BLUE; c_width.font = f
            if row.get("_blue_height"):
                f = copy(c_height.font); f.color = BLUE; c_height.font = f
            # 같은 세로의 따옴표 표시는 앞 공백 2칸을 유지하고 왼쪽+아래쪽 맞춤(2026-09-14).
            if str(row.get("세로") or "").strip() == '"':
                ws.cell(r, 6).alignment = Alignment(horizontal="left", vertical="bottom",
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
            if row["기재사항"] and not row["기재사항2"]:
                # 공통 규칙: 기재사항이 하나뿐이면 기재사항1·2를 가로 병합한다.
                span.append(r)
            date_cell = ws.cell(r, date_col, row["출고일"])
            if "\n" in str(row.get("출고일") or ""):
                # 배송방식은 날짜 아래 줄에 표시한다. 다만 같은 주문의 출고일 셀이
                # 여러 행에 걸쳐 세로 병합되는 경우 첫 제품행만 높아지면 장부 행높이가
                # 들쭉날쭉해진다. 여기서는 줄바꿈/정렬만 설정하고, 실제 행높이 증가는
                # 아래 출고일 병합 단계에서 '단일 행 주문'에만 적용한다.
                date_cell.alignment = Alignment(horizontal="center", vertical="center",
                                                shrink_to_fit=True, wrap_text=True)
        r += 1

    # DI 수령인은 색상이 달라도 같은 사람의 연속 행 전체를 하나로
    # 병합한다. 세부 양식에서는 기재사항 두 칸(K:L)을 직사각형으로 묶는다.
    i = 0
    while i < len(rows):
        recipient = rows[i].get("_DI수령인")
        if (not recipient or rows[i].get("_특수") is not None
                or rows[i].get("_manual_note1") or rows[i].get("_manual_note2")):
            i += 1
            continue
        j = i + 1
        while (j < len(rows) and rows[j].get("_특수") is None
               and rows[j].get("_DI수령인") == recipient
               and not rows[j].get("_manual_note1")
               and not rows[j].get("_manual_note2")):
            j += 1
        _safe_merge(ws, start + i, start + j - 1, 11 if detailed else 10, 12 if detailed else 11)
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
        while j < len(rows) and (
                (rows[j].get("_특수") is None
                 and rows[j].get("_JL고객") == customer
                 and rows[j].get("_oi") == order_id)
                or (_is_mix_part(rows[j]) and rows[j].get("_oi") == order_id)):
            j += 1
        # MIX 구성행은 같은 주문 기재사항 병합 안에 포함하되 값 비교에서는 제외한다.
        product_k = [k for k in range(i, j) if rows[k].get("_특수") is None]
        if j - i > 1:
            # 최초 자동 판독값만 병합한다. 사용자가 기재사항1/2를 직접 수정한
            # 행이 하나라도 섞이면 같은 문자열이어도 자동으로 다시 병합하지 않는다.
            first_note = str(row0.get("기재사항") or "").strip()
            same_note1 = all(str(rows[k].get("기재사항") or "").strip() == first_note
                             for k in product_k)
            note1_manual = any(rows[k].get("_manual_note1") for k in product_k)
            note2_manual = any(rows[k].get("_manual_note2") for k in product_k)
            if first_note and same_note1 and not note1_manual:
                all_note2_empty = all(not str(rows[k].get("기재사항2") or "").strip()
                                      for k in product_k)
                if all_note2_empty and not note2_manual:
                    _safe_merge(ws, start + i, start + j - 1, 11 if detailed else 10, 12 if detailed else 11)
                else:
                    _safe_merge(ws, start + i, start + j - 1, 11 if detailed else 10, 11 if detailed else 10)
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
        while j < len(rows) and (
                (rows[j].get("_특수") is None
                 and rows[j].get("_oi") == order_id
                 and rows[j].get("_JO주문번호") == order_no)
                or (_is_mix_part(rows[j]) and rows[j].get("_oi") == order_id)):
            j += 1
        if j - i > 1:
            col = 11 if detailed else 10
            # 3창 이상 JO는 기재사항1이 순수 주문번호이므로 세로 병합한다.
            # 1~2창은 주문번호 뒤에 창별 메모가 함께 들어갈 수 있어 병합하지 않는다.
            values = [str(rows[k].get("기재사항") or "").strip() for k in range(i, j)]
            pure_order_no = all((not v) or v == str(order_no).strip() for v in values)
            note1_manual = any(rows[k].get("_manual_note1") for k in range(i, j))
            if pure_order_no and not note1_manual:
                note2_empty = all(not str(rows[k].get("기재사항2") or "").strip()
                                  and not rows[k].get("_manual_note2") for k in range(i, j))
                # 기재사항2가 모두 비었으면 1·2 칸을 직사각형으로 병합한다(2만 따로 남지 않게).
                _safe_merge(ws, start + i, start + j - 1, col, col + 1 if note2_empty else col)
                ws.cell(start + i, col).alignment = Alignment(
                    horizontal="center", vertical="center", shrink_to_fit=True)
        i = j

    # 원코드 연창의 # 표시는 연창 묶음 전체를 하나의 특이 셀로 병합한다.
    if detailed:
        i = 0
        while i < len(rows):
            row0 = rows[i]
            if row0.get("_특수") is not None or not row0.get("_연창"):
                i += 1
                continue
            order_id = row0.get("_oi")
            j = i + 1
            while (j < len(rows) and rows[j].get("_특수") is None
                   and rows[j].get("_oi") == order_id and rows[j].get("_연창")):
                j += 1
            if j - i > 1 and str(row0.get("특이") or "").strip() == "#":
                _safe_merge(ws, start + i, start + j - 1, 10, 10)
                ws.cell(start + i, 10).alignment = Alignment(horizontal="center", vertical="center")
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
            manual_key = "_manual_note1" if note_key == "기재사항" else "_manual_note2"
            # JL/JO 전용 병합 규칙은 기재사항1에만 적용한다. 기재사항2는 이 공통 세로 병합을 쓴다.
            dedicated = bool(row0.get("_DI수령인")) or (note_key == "기재사항" and bool(
                row0.get("_JL고객") or row0.get("_JO주문번호")))
            if (row0.get("_특수") is not None or not value or order_id is None
                    or dedicated or row0.get(manual_key)):
                i += 1
                continue
            j = i + 1
            while j < len(rows) and (
                    (rows[j].get("_특수") is None
                     and rows[j].get("_oi") == order_id
                     and str(rows[j].get(note_key) or "").strip() == value
                     and not rows[j].get("_DI수령인")
                     and not (note_key == "기재사항"
                              and (rows[j].get("_JL고객") or rows[j].get("_JO주문번호")))
                     and not rows[j].get(manual_key))
                    or (_is_mix_part(rows[j]) and rows[j].get("_oi") == order_id)):
                j += 1
            if j - i > 1:
                end_col = note_col
                if note_key == "기재사항" and all(
                        not str(rows[k].get("기재사항2") or "").strip()
                        and not rows[k].get("_manual_note2") for k in range(i, j)):
                    # 기재사항2가 모두 비었으면 1·2 칸을 직사각형으로 병합한다(1만 합쳐지고 2가 남던 문제).
                    end_col = note_col + 1
                try:
                    _safe_merge(ws, start + i, start + j - 1, note_col, end_col)
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
            _safe_merge(ws, start + i, start + j - 1, 3, 3)
            ws.cell(start + i, 3).alignment = Alignment(
                horizontal="center" if center_head else "left",
                vertical="center", shrink_to_fit=True)
        i = j

    for rr in span:                 # 기재사항 가로 병합 (J:K)
        _safe_merge(ws, rr, rr, 11 if detailed else 10, 12 if detailed else 11)

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
        if row0.get("출고일"):
            date_col = 13 if detailed else 12
            if j - i > 1:
                _safe_merge(ws, r0, dated[j - 1][0], date_col, date_col)
                ws.cell(r0, date_col).alignment = Alignment(horizontal="center",
                                                            vertical="center", wrap_text=True)
                # 여러 행에 걸쳐 병합된 출고일은 병합 영역 자체에 충분한 높이가 있으므로
                # 날짜가 들어간 첫 행도 다른 제작행과 동일한 기본 높이를 유지한다.
                # (예: 한길 본품행만 유독 넓어지던 현상 방지)
                if ws.row_dimensions[r0].height == 31:
                    ws.row_dimensions[r0].height = ROW_H
            elif "\n" in str(row0.get("출고일") or ""):
                # 주문이 실제로 한 행뿐이면 날짜+배송방식 두 줄을 보여줄 공간이 필요하다.
                ws.row_dimensions[r0].height = max(ws.row_dimensions[r0].height or ROW_H, 31)
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
                                          or row0.get("_JO주문번호")
                                          or row0.get("_manual_note1")):
                i += 1
                continue
            j = i + 1
            # 길이(모형2)는 실제로 같은 값만 병합한다. 기본 길이라 비워 둔 창까지 위 길이와 병합하면
            # 장부에서 두 창이 같은 길이로 보이는 오류가 생긴다(예: 120 / 기본 130).
            same_value = ((lambda v: v == row0.get(key)) if key == "모형2"
                          else (lambda v: v in (None, "", row0.get(key))))
            while (j < len(prod) and prod[j][1].get("_같은색")
                   and same_value(prod[j][1].get(key))
                   and not (key == "기재사항" and prod[j][1].get("_manual_note1"))):
                j += 1
            if col == (11 if detailed else 10) and any(rr in spanset for rr in
                                 (p[0] for p in prod[i:j])):
                i = j                # 가로 병합한 행은 세로 병합에서 제외
                continue
            if j - i > 1 and row0.get(key):
                _safe_merge(ws, r0, prod[j - 1][0], col, col)
            i = j
    return r


def _ledger_fireproof_cert_rows(rows):
    """롤/콤비 방염 ``<필증>``을 FM 장부의 남는 C열 칸에 배치한다.

    사용자 정답 장부의 우선순위는 다음과 같다.
    1) 같은 방염 품목이 여러 제작행이면 마지막 연속 제작행의 비어 있는 C열을 사용한다.
    2) 한 제작행뿐이고 바로 다음에 공통 메모행이 있으면 그 메모행의 C열을 사용한다.
    3) 위 둘 다 없을 때만 ``<필증>`` 전용 행을 새로 만든다.

    즉 이미 아래쪽에 사용할 수 있는 칸이 있는데 필증 때문에 행을 하나 더 만들지 않는다.
    """
    rows = [dict(r) for r in rows]
    out = []
    i = 0
    while i < len(rows):
        row = rows[i]
        if not (row.get("_특수") is None and row.get("_roll_fireproof")):
            out.append(row)
            i += 1
            continue

        key = row.get("_roll_fire_key")
        oi = row.get("_oi")
        j = i + 1
        while j < len(rows):
            nxt = rows[j]
            if not (nxt.get("_특수") is None
                    and nxt.get("_roll_fireproof")
                    and nxt.get("_roll_fire_key") == key
                    and (oi is None or nxt.get("_oi") == oi)):
                break
            j += 1

        group = rows[i:j]
        if len(group) >= 2:
            # 같은 품목의 2번째 이후 제작행은 색상칸이 원래 비어 있으므로
            # 마지막 제작행 C열을 필증 칸으로 재사용한다.
            group[-1]["_roll_cert_here"] = True
            out.extend(group)
            i = j
            continue

        out.extend(group)
        if j < len(rows) and rows[j].get("_roll_common_note"):
            # 다음 공통 메모행이 이미 있으면 그 행 C열에 필증을 두고
            # 메모는 기존대로 D열 이후를 사용한다.
            note_row = dict(rows[j])
            note_row["_roll_cert_here"] = True
            out.append(note_row)
            i = j + 1
            continue

        # 사용할 기존 행이 전혀 없을 때만 필증 전용 행을 추가한다.
        out.append({
            "_특수": "__필증__", "_ledger_only": True,
            "_거래처코드": row.get("_거래처코드"),
            "_oi": oi, "문구": None,
        })
        i = j
    return out


def build_ledger(all_rows, path, header_date=None):
    wb = openpyxl.Workbook()
    di_rows = [r for r in all_rows if r.get("_거래처코드") == "DI"]
    general_rows = _ledger_fireproof_cert_rows(
        [r for r in all_rows if r.get("_거래처코드") != "DI"]
    )
    used_active = False

    # DI는 다른 업체와 같은 배치에 있어도 항상 C/L·원/투/셔터 전용 시트 규칙을 유지한다.
    if di_rows:
        for group_name, group_key in DI_SHEET_GROUPS:
            group_rows = [r for r in di_rows if r.get("_DI유형") == group_key]
            if not group_rows:
                continue
            for page_no, chunk in enumerate(_di_paginate(group_rows), 1):
                title = group_name if page_no == 1 else f"{group_name} ({page_no})"
                page_rows = max(DI_MAX_ROWS, len(chunk))
                if not used_active:
                    ws = wb.active
                    ws.title = title
                    _init_sheet(ws, page_rows + 2, detailed=True)
                    _block(ws, 1, page_rows, header_date, detailed=True)
                    used_active = True
                else:
                    ws = _new_sheet(wb, title, page_rows, header_date, detailed=True)
                _write_rows(ws, chunk, detailed=True)
        # 유형값이 비정상인 DI 행도 유실하지 않는다.
        known = {key for _, key in DI_SHEET_GROUPS}
        unknown_di = [r for r in di_rows if r.get("_DI유형") not in known]
        if unknown_di:
            title = "DI 기타"
            if not used_active:
                ws = wb.active
                ws.title = title
                _init_sheet(ws, max(DI_MAX_ROWS, len(unknown_di)) + 2, detailed=True)
                _block(ws, 1, max(DI_MAX_ROWS, len(unknown_di)), header_date, detailed=True)
                used_active = True
            else:
                ws = _new_sheet(wb, title, max(DI_MAX_ROWS, len(unknown_di)), header_date, detailed=True)
            _write_rows(ws, unknown_di, detailed=True)

    if general_rows or not used_active:
        pages = _paginate_order_rows(general_rows, LEDGER_ROWS) if general_rows else [[]]
        for n, chunk in enumerate(pages, 1):
            title = "장부" if n == 1 else f"장부 ({n})"
            if not used_active:
                ws = wb.active
                ws.title = title
                _init_sheet(ws, LEDGER_ROWS + 2, detailed=True)
                _block(ws, 1, LEDGER_ROWS, header_date, detailed=True)
                used_active = True
            else:
                ws = _new_sheet(wb, title, LEDGER_ROWS, header_date, detailed=True)
            _write_rows(ws, chunk, detailed=True)
    wb.save(path)
    return path


class _PlanFont:
    color = None


class _PlanCell:
    """ledger_merge_ranges 전용 가짜 셀. 값/서식 대입만 받고 아무것도 그리지 않는다."""

    def __init__(self, value=None):
        self.value = value
        self.font = _PlanFont()
        self.alignment = None


class _PlanRowDimension:
    height = None


class _PlanRange:
    def __init__(self, min_row, max_row, min_col, max_col):
        self.min_row, self.max_row = min_row, max_row
        self.min_col, self.max_col = min_col, max_col


class _PlanMergedCells:
    def __init__(self):
        self.ranges = []


class _MergePlanSheet:
    """_write_rows 를 실제 시트 없이 실행해 병합 범위만 기록한다."""

    def __init__(self):
        self.merged_cells = _PlanMergedCells()
        self._dims = {}
        self.row_dimensions = self

    def __getitem__(self, row):
        return self._dims.setdefault(row, _PlanRowDimension())

    def cell(self, row, column, value=None):
        return _PlanCell(value)

    def merge_cells(self, start_row, end_row, start_column, end_column):
        self.merged_cells.ranges.append(_PlanRange(start_row, end_row, start_column, end_column))


def ledger_merge_ranges(output_rows):
    """장부 엑셀(build_ledger → _write_rows)과 같은 규칙으로 병합 범위를 계산한다.

    사전점검 화면이 장부와 다른 자체 병합 규칙을 갖지 않도록 하는 단일 기준이다.
    반환: [(첫 행 _output_index, 끝 행 _output_index, 시작 열, 끝 열)]
    열 번호는 장부 세부양식 기준(B=2 상호 … M=13 출고일). 페이지 분할은 반영하지 않는다.
    """
    di_rows = [r for r in output_rows if r.get("_거래처코드") == "DI"]
    general_rows = _ledger_fireproof_cert_rows(
        [r for r in output_rows if r.get("_거래처코드") != "DI"])
    known = {key for _, key in DI_SHEET_GROUPS}
    chunks = [[r for r in di_rows if r.get("_DI유형") == key] for _, key in DI_SHEET_GROUPS]
    chunks.append([r for r in di_rows if r.get("_DI유형") not in known])
    chunks.append(general_rows)
    ranges = []
    for chunk in chunks:
        if not chunk:
            continue
        sheet = _MergePlanSheet()
        _write_rows(sheet, chunk, start=0, detailed=True)
        for rng in sheet.merged_cells.ranges:
            indices = [chunk[i].get("_output_index") for i in range(rng.min_row, rng.max_row + 1)]
            indices = [i for i in indices if isinstance(i, int)]
            if indices:
                ranges.append((min(indices), max(indices), rng.min_col, rng.max_col))
    return ranges


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
def _worksheet_excluded(row):
    """작업지시서에는 주소·이름/전화번호·발신·포장비용을 넣지 않는다(2026-09-14 사용자 확정).

    전달사항, 부속, MIX 구성행, 롤·콤비 공통메모, 아지트 가공소 표시는 남긴다.
    """
    mark = row.get("_특수")
    text = str(row.get("문구") or "").strip()
    if mark == "#":
        return re.sub(r"\s+", "", text) == "포장비용"
    if mark == "☆":
        return not row.get("_az_place")
    if mark == "":
        return not (row.get("_roll_common_note") or text.startswith("전달"))
    return False


def to_worksheet_rows(rows):
    """내부표시 제거, 색상 앞 'B ' 제거, 특수 행 제외.

    작업지시서도 세부 양식을 사용하므로 #·틀안은 길이/기재사항에
    중복하지 않고 특이 칸에만 유지한다.
    """
    out = []
    for row in rows:
        if row.get("_ledger_only") or _worksheet_excluded(row):
            continue
        r = dict(row)
        if r.get("_특수") is None:
            r["내부표시"] = None                  # (K), (N), (1) 등 삭제
            if r.get("색상"):
                color_text = str(r.get("색상") or "")
                if color_text.strip() == "<필증>":
                    # 작업지시서는 품명을 주문행 전체에 병합하므로 별도 필증 표시를 품명칸에 두지 않는다.
                    r["색상"] = None
                    r["_같은색"] = True
                else:
                    color_text = color_text.replace("\n<필증>", "")
                    r["색상"] = re.sub(r"^\s*[BHRCS]\s+", "", color_text, flags=re.I).strip()
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


def _merge_worksheet_order_identity(ws, rows, start_row):
    """작업지시서에서 업체명과 품명을 주문 단위로 세로 병합한다.

    업체명(B열)은 같은 주문의 제작품 행 전체, 품명(C열)은 같은 주문 안에서
    같은 품명이 이어지는 구간만 병합한다. 부속/주소 같은 특수행은 제외한다.
    """
    # MIX 구성행은 같은 주문의 일부라 업체명 병합에 포함하고, 품명 병합에서는 끊는다(4색 조합 문구가 C열에 있음).
    product_rows = [(idx, row) for idx, row in enumerate(rows)
                    if row.get("_특수") is None or _is_mix_part(row)]
    i = 0
    while i < len(product_rows):
        idx0, row0 = product_rows[i]
        oi = row0.get("_oi")
        j = i + 1
        while j < len(product_rows):
            idxj, rowj = product_rows[j]
            if rowj.get("_oi") != oi or idxj != product_rows[j - 1][0] + 1:
                break
            j += 1
        if j - i > 1:
            _safe_merge(ws, start_row + idx0, start_row + product_rows[j - 1][0], 2, 2)
        # 품명은 같은 주문 안에서 같은 색상/제품 묶음별로 병합
        k = i
        while k < j:
            idxk, rowk = product_rows[k]
            if _is_mix_part(rowk):
                k += 1
                continue
            base = rowk.get("색상")
            m = k + 1
            while m < j:
                idxm, rowm = product_rows[m]
                if idxm != product_rows[m - 1][0] + 1 or _is_mix_part(rowm):
                    break
                if rowm.get("색상") not in (None, "", base) and not rowm.get("_같은색"):
                    break
                m += 1
            if m - k > 1 and base:
                _safe_merge(ws, start_row + idxk, start_row + product_rows[m - 1][0], 3, 3)
            k = m
        i = j


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
        groups = [(name, key, [r for r in ws_rows if r.get("_DI유형") == key])
                  for name, key in DI_SHEET_GROUPS]
        # 장부와 같이 주문이 없는 유형(예: 셔터 주문 없음)은 작업지시서 시트도 만들지 않는다.
        groups = [g for g in groups if g[2]] or groups[:1]
        for group_name, group_key, group_rows in groups:
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
                _merge_worksheet_order_identity(ws, chunk, first)
        wb.save(path)
        return path

    pages, cur = [], []
    for start, n in _group_len(ws_rows):
        block = ws_rows[start:start + n]
        if cur and len(cur) + n > SHEET_ROWS:
            pages.append(cur)
            cur = []
        if n > SHEET_ROWS:
            # 한 업체 묶음이 블록(16행)보다 길면 아래 확인용 블록을 침범해
            # 병합 셀에 헤더를 쓰다 실패한다. 장부와 같이 블록 크기로 나눈다.
            pages.extend(block[k:k + SHEET_ROWS] for k in range(0, n, SHEET_ROWS))
            continue
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
        _merge_worksheet_order_identity(ws, chunk, f1)
        f2 = _block(ws, SHEET_ROWS + 3, SHEET_ROWS, header_date,
                    detailed=True)
        _write_rows(ws, chunk, start=f2, center_head=True,
                    detailed=True)      # 확인용
        _merge_worksheet_order_identity(ws, chunk, f2)
    wb.save(path)
    return path


# ─────────────────────────────────────────────
# 4. 경영박사 전표
# ─────────────────────────────────────────────
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
    """경영박사 EDI 주소: 결제문구 없이 택배 주소 규칙으로 정리한다."""
    return _short_city_address(value)


def _erp_address_lines(address):
    """경영박사 택배 주소 줄: 번지 뒤 쉼표 없이 상세주소를 새 줄로 쓴다(2026-09-14).

    예) `태안군 근흥면 근흥로 687-1, 서울한약방` → `태안군 근흥면` / `근흥로 687-1` / `서울한약방`
    장부 주소행은 쉼표를 그대로 둔다.
    """
    lines = wrap_erp_text(address) if str(address or "").strip() else []
    for k, line in enumerate(lines):
        if "," not in line:
            continue
        # 쉼표 앞(번지)까지는 기존 줄 나눔 그대로, 쉼표 뒤 상세주소부터 새 줄로 다시 나눈다.
        head, _, tail = line.partition(",")
        detail = " ".join([tail.strip()] + lines[k + 1:]).strip()
        return (lines[:k] + ([head.strip()] if head.strip() else [])
                + (wrap_erp_text(detail) if detail else [])) or [""]
    return lines or [""]


def _erp_order_name(order):
    """추가부속·포장비용 경영박사 적요에 적는 이름(2026-09-14). 제품 적요의 이름 규칙을 따른다."""
    client = order.get("거래처")
    if client == "DI":
        return ""                       # DI는 한 주문에 수령인이 여럿이다.
    if client in BONO_LIKE_CLIENTS:
        return _bono_erp_identity(order)
    delivery = order.get("배송") or {}
    receiver = str(delivery.get("수령인") or "").strip()
    customer = str(order.get("고객명") or "").strip()
    if client == "JL":
        name = _jl_customer_name(customer or receiver)
        return "" if is_client_staff_name(client, name) else name
    if client == "RT" and receiver == "더커튼":
        receiver = "더"
    # 배송 정보가 없는 자기 배송은 받는 이름이 없다(장부에서 되돌릴 수 없는 고객명으로 채우지 않음).
    return receiver or (customer if str(delivery.get("주소") or "").strip() else "")


def _unit_accessory_lines(order):
    """유앤 추가부속을 품명별로 합친다(2026-09-14 사용자 확정).

    - 피스 종류(피스(석고앙카)/피스(석고날개)/피스(콘크리트)…): 창마다 가로에 맞는 피스 수를 더해 `EA`.
    - 노피스브라켓: 세트 수를 더해 `set`.
    반환: [{"품명", "단위"(EA/set), "수량", "표시수량"}]
    """
    lines = {}
    for acc in order.get("_unit_accessories") or []:
        name = str(acc.get("품명") or "").strip()
        if not name:
            continue
        try:
            sets = max(1, int(float(acc.get("세트") or 1)))
        except (TypeError, ValueError):
            sets = 1
        if "노피스" in name:
            unit, n = "set", sets
        elif acc.get("수량EA") not in (None, ""):
            unit, n = "EA", int(acc["수량EA"])
        else:
            try:
                pieces = int(float(acc.get("피스수") or 0))
            except (TypeError, ValueError):
                pieces = 0
            unit, n = "EA", sets * pieces
        entry = lines.setdefault(name, {"품명": name, "단위": unit, "수량": 0})
        entry["수량"] += n
    return [dict(x, 표시수량=f"{x['수량']}{x['단위']}") for x in lines.values()]

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

    special_font = xlwt.Font()
    special_font.name = "맑은 고딕"
    special_font.height = 11 * 20
    special_align = xlwt.Alignment()
    special_align.horz = xlwt.Alignment.HORZ_CENTER
    special_align.vert = xlwt.Alignment.VERT_TOP
    special_style = xlwt.XFStyle()
    special_style.font = special_font
    special_style.alignment = special_align

    def write_row(values, style=None):
        nonlocal out_row
        for col, value in enumerate(values):
            if style is None:
                ws.write(out_row, col, value)
            else:
                ws.write(out_row, col, value, style)
        out_row += 1

    def money(unit, qty, vat_zero=False):
        amount = (Decimal(str(unit)) * Decimal(str(qty))).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP)
        vat = Decimal("0") if vat_zero else (amount / 10).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP)
        return int(amount), int(vat)

    for voucher_no, o in enumerate(orders, 1):
        synchronize_order_for_outputs(o)
        # 호출자가 지정한 출고일이 최우선이다. 원본/이전 장부 날짜가 다시 덮지 못한다.
        d = order_date or o.get("_ship_date") or date.today()
        if isinstance(d, str):
            try:
                d = date.fromisoformat(d[:10])
            except ValueError:
                d = order_date or date.today()
        dstr = f"{d.year}.{d.month:02d}.{d.day:02d}"
        client = o.get("거래처")
        is_roll_order = o.get("_product_mode") == "roll_combo"
        erp_client = ((ROLL_CLIENT_INFO.get(client) or {}).get("erp") if is_roll_order else None) or ERP_CLIENT_NAME.get(client, client)
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
        roll_window_count = 0
        total_valid_count = 0
        holding_pack_qty = Decimal("0")
        holding_pack_recipient = None
        di_pack_recipient = None
        di_pack_qty = Decimal("0")
        roll_fireproof_needed = False

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
            is_roll = it.get("_product_group") == "roll_combo" or is_roll_order

            # ── 롤/콤비/쉐이드 ───────────────────────
            if is_roll:
                hit = M.find_order_item(it, client)
                if not hit or not hit.get("단가"):
                    continue
                w, h = it.get("가로"), it.get("세로")
                if not (w and h):
                    continue
                count = 1
                try:
                    count = max(1, int(float(it.get("창개수") or it.get("수량") or 1)))
                except (TypeError, ValueError):
                    count = 1
                # 롤/콤비/쉐이드: 세로 최소 150cm, 창 1개당 최소 2.0㎡ 청구.
                calc = calc_roll_erp(w, h, hit["단가"])
                qty = (Decimal(str(calc["수량"])) * count).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                amount = int((Decimal(str(hit["단가"])) * qty).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                vat = int((Decimal(amount) / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                memo_parts = [f"{float(w):.1f}*{float(h):.1f}/{count}EA"]
                # 롤/콤비 손잡이길이는 장부와 같이 5단위 내림(2026-09-14).
                handle_n = normalize_roll_handle(it.get("손잡이길이"))[0]
                if handle_n and handle_n != 150:
                    memo_parts.append(f"손{handle_n}")
                if client in BONO_LIKE_CLIENTS:
                    # 보노 계열 경영박사 적요 순서: 손 → 지점 → 방 → 고객정보(오더명).
                    identity = _bono_erp_identity(o)
                    if identity and identity not in memo_parts:
                        memo_parts.append(identity)
                for value in (it.get("설치장소"), o.get("주문번호"), it.get("기재사항"), it.get("_수동특이")):
                    for part in str(value or "").split("/"):
                        part = part.strip()
                        if part and part not in memo_parts:
                            memo_parts.append(part)
                direction = it.get("손잡이방향") or "우"
                values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                          "", "", hit.get("관리코드") or hit.get("품명"), hit.get("품명") or hit.get("관리코드"),
                          hit.get("규격") or "", float(qty), int(hit["단가"]),
                          amount, vat, _erp_memo_join(memo_parts),
                          1 if direction == "좌" else 2, direction]
                write_row(values)
                if it.get("_roll_fireproof"):
                    roll_fireproof_needed = True
                roll_window_count += count
                total_valid_count += 1
                continue

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
                    # 수령인이 `현장B/베란다`처럼 여러 값이면 각 값을 메모에서 빼고 마지막에 한 번만 붙인다.
                    recipient_keys = {re.sub(r"\s+", "", x) for x in str(recipient or "").split("/") if x.strip()}
                    for value in (it.get("_수동특이"), it.get("기재사항"), it.get("설치장소")):
                        for part in str(value or "").split("/"):
                            part = part.strip()
                            if part and re.sub(r"\s+", "", part) not in recipient_keys \
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

                # 자석바 추가비용은 사용자 제공 추가비용 품목장의 실제 SKU를 사용한다.
                # 세로 300cm 이상이면 /300cm이상 품목, 추가 자석바 N개면 +자석바N 품목.
                for magnet_name in holding_magnet_extra_names(op, h):
                    ex = M.find_holding_extra(magnet_name, client)
                    if ex and ex.get("단가"):
                        mq = Decimal(count)
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
            # 사전점검에서 확정된 장부 기재사항1/2를 이 품목의 최종 메모 원본으로 쓴다.
            src_idx = source_idx_by_id.get(id(it))
            has_final_notes = src_idx is not None and src_idx in ledger_note_by_index
            ledger_notes = ledger_note_by_index.get(src_idx, ()) if has_final_notes else ()
            ledger_tail = []
            for value in ledger_notes:
                for part in str(value or "").split("/"):
                    part = part.strip()
                    if part and part not in ledger_tail:
                        ledger_tail.append(part)
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
                # DI는 장부와 동일하게 기재사항1(이름/MIX) 한 칸만 사용한다.
                # 설치장소는 EDI 기재사항에 별도로 다시 붙이지 않는다.
            elif client == "JO":
                # JO EDI도 사전점검/장부의 기재사항1·2를 최종 기준으로 사용한다.
                # 별도로 고객명/현장명을 다시 붙이지 않아 화면과 전산이 달라지지 않게 한다.
                extra_parts = []
            elif client in BONO_STYLE_CLIENTS:
                # 보노/미래가공 EDI: 받는사람이 외부인이면 그 이름을,
                # 거래처 본인이면 화물지점의 현장명만 맨 앞에 둔다.
                # 그 뒤에는 사전점검/장부에서 확정된 오더명·시공위치를 그대로 붙인다.
                handle = normalize_handle(it.get("손잡이길이"))[0]
                if handle:
                    extra_parts.append(f"손{handle}")
                identity = _bono_erp_identity(o)
                if identity:
                    extra_parts.append(identity)
                # 경영박사 적요 순서(사용자 확정): 손 → 지점 → 방 → 고객정보(오더명). 장부는 그대로.
                order_name = str(o.get("주문번호") or "").strip()
                ordered_tail = ([p for p in ledger_tail if p != order_name]
                                + ([order_name] if order_name in ledger_tail else []))
                for part in ordered_tail:
                    if part not in extra_parts:
                        extra_parts.append(part)
            elif client == "SP":
                # 스페이스 적요 순서(2026-09-18 사용자 확정):
                # 피스 -> 주문번호 -> 줄길이 -> 창별 세부 기재사항 -> 공통 기재사항
                # 예) 피스/3-17/작은방/우미린아파트
                # 장부 기재사항1(공통)·2(세부)를 기준으로 만들어 발주서/장부 어느 쪽에서 와도 같게 한다.
                sp_order_no = str(o.get("주문번호") or "").strip()
                if has_final_notes:
                    common_parts = _split_note_parts(ledger_notes[0] if ledger_notes else None)
                    detail_parts = _split_note_parts(ledger_notes[1] if len(ledger_notes) > 1 else None)
                else:
                    common_parts = _split_note_parts(order_common_note, o.get("고객명"))
                    detail_parts = _split_note_parts(it.get("설치장소"), it.get("기재사항"))
                detail_parts = _split_note_parts(it.get("_수동특이")) + detail_parts
                if any("피스" in x for x in common_parts + detail_parts):
                    extra_parts.append("피스")
                if sp_order_no:
                    extra_parts.append(sp_order_no)
                handle_text = _erp_ledger_handle(it)      # 기본 길이가 아닌 줄길이
                if handle_text:
                    extra_parts.append(handle_text)
                if _has_non_di_mix(it, client):
                    extra_parts.append("MIX")
                for part in detail_parts + common_parts:
                    if "피스" in part or part == sp_order_no or part in extra_parts:
                        continue
                    extra_parts.append(part)
            elif client == "JL":
                common = str(o.get("전체기재사항") or "")
                receiver = _jl_customer_name(o.get("고객명") or delivery.get("수령인") or "")
                if is_client_staff_name(client, receiver):
                    receiver = ""
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
                # 기본 길이가 아닌 손잡이길이는 전 업체 전산 적요에 적는다(2026-09-18).
                handle_text = _erp_ledger_handle(it)
                if handle_text:
                    extra_parts.append(handle_text)
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
            # 장부의 최종 기재사항이 존재하면 오래된 원본 전체기재/설치장소를 다시 조합하지 않는다.
            # 제작에 필요한 특이/MIX/손잡이 정보만 보존하고 기재사항 부분은 ledger_tail로 교체한다.
            if has_final_notes and client != "SP" and client not in BONO_STYLE_CLIENTS:
                production_parts = []
                if client == "DI":
                    if it.get("_manual_special_override"):
                        sp = str(it.get("_수동특이") or "").strip()
                        if sp:
                            production_parts.append(sp)
                    handle = normalize_handle(it.get("손잡이길이"))[0]
                    if handle:
                        production_parts.append(f"손{handle}")
                    if it.get("_mix_name") and it.get("_mix_codes"):
                        mix_codes = str(it["_mix_codes"]).replace(" ", "")
                        mix_text = (mix_codes if it.get("_generic_mix") or str(it.get("_mix_name")).upper() == "MIX"
                                    else f"{it['_mix_name']}{mix_codes}")
                        production_parts.append(mix_text)
                elif client == "JL":
                    handle = normalize_handle(it.get("손잡이길이"))[0]
                    if handle:
                        production_parts.append(f"손{handle}")
                    sp = str(it.get("_수동특이") or "").strip()
                    if sp:
                        production_parts.append(sp)
                else:
                    handle_text = _erp_ledger_handle(it)
                    if handle_text:
                        production_parts.append(handle_text)
                    sp = str(it.get("_수동특이") or "").strip()
                    if sp:
                        production_parts.append(sp)
                    if _has_non_di_mix(it, client):
                        production_parts.append("MIX")
                extra_parts = []
                for part in production_parts + ledger_tail:
                    if part and part not in extra_parts:
                        extra_parts.append(part)
            if it.get("_blind_mix") and client != "DI":
                # 비대일 MIX: 적요 맨 앞에 `200(108)+125(12)`, 그 뒤에 나머지 기재사항.
                mix_text = mix_memo_text(it)
                extra_parts = [mix_text] + [x for x in extra_parts if x not in (mix_text, "MIX")]
            if (it.get("종류") == "원코드" and it.get("연창")
                    and it.get("_product_group") not in {"holding", "roll_combo"}):
                # 장부 특이 칸 연창 # 과 같은 창: 경영박사 적요 맨 앞에 # 을 붙인다.
                extra_parts = ["#"] + [x for x in extra_parts if x != "#"]
            # 포장비용은 별도 품목행이다. 제품 적요에는 적지 않는다(장부에서 불러온 주문 포함).
            extra_parts = [x for x in extra_parts if re.sub(r"\s+", "", str(x)) != "포장비용"]
            if client == "휴안":
                # 휴안 경영박사: 석고앙카N 은 `앙카N` (세로가 가장 긴 창에만 기록됨).
                extra_parts = [re.sub(r"^석고앙카", "앙카", x) for x in extra_parts]
            extra = strip_vinyl_text("/".join(extra_parts))
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
            if it.get("_blind_mix") and client != "DI":
                # MIX 추가비용: 창 1개당 1개, 단가 0원(사용자 확정).
                extra_name = mix_extra_cost_name(it.get("타입"))
                mix_windows = 1 if split_n else max(1, _blind_counts(it)[2])
                write_row([dstr, voucher_no, 3, "외출", erp_client, erp_client, "", "",
                           extra_name, extra_name, "", float(mix_windows), 0, 0, 0, "", "", ""])
            blind_valid_count += 1
            try:
                _lpack, _rpack, _npack = _blind_counts(it)
                blind_window_count += max(1, int(_npack or 1))
            except Exception:
                blind_window_count += 1
            total_valid_count += 1

        # 롤/콤비 방염 제품이 하나라도 있으면 필증을 주문당 1행 추가한다.
        if is_roll_order and roll_fireproof_needed:
            cert = "▶필증(방염)◀"
            values = [dstr, voucher_no, 3, "외출", erp_client, erp_client, "", "",
                      cert, cert, "", 1.0, 0, 0, 0, "", "", ""]
            write_row(values, special_style)

        # 추가부속·포장비용 적요에 적는 이름(추가부속은 `수량/이름`, 포장비용은 `이름`).
        order_name = _erp_order_name(o)

        def accessory_memo(qty_text):
            return "/".join(x for x in (qty_text, order_name) if x)

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
                          code, pname, spec, qty, unit, amount, vat, accessory_memo(f"{qty:g}EA"), "", ""]
                write_row(values)

        # 휴안 추가부속: 장부 `노피스(3)` / `2EA` → 경영박사 `노피스브라켓(3EA)` 수량 2.
        # 품목명은 기존 등록 품목명 규칙(노피스브라켓(2EA)/(3EA))을 따르고, 단가가 없으면 0원으로 둔다.
        if client == "휴안":
            for accessory in o.get("_huan_accessories") or []:
                m = re.fullmatch(r"노피스\((\d+)\)", str(accessory.get("표시") or "").strip())
                if not m:
                    continue
                edi_name = f"노피스브라켓({m.group(1)}EA)"
                qm = re.search(r"\d+", str(accessory.get("수량") or ""))
                qty = float(qm.group(0)) if qm else 1.0
                hit = M.find_named_item(edi_name, client)
                unit = int((hit or {}).get("단가") or 0)
                amount = int(unit * qty)
                vat = int((Decimal(amount) / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                write_row([dstr, voucher_no, 3, "외출", erp_client, erp_client, "", "",
                           (hit or {}).get("관리코드") or edi_name, (hit or {}).get("품명") or edi_name,
                           (hit or {}).get("규격") or "", qty, unit, amount, vat,
                           accessory_memo(f"{qty:g}set"), "", ""])

        # 유앤 추가부속(2026-09-14): 장부와 같은 합계. 피스 종류는 수량 0·단가 0에 적요 `10EA/이름`,
        # 노피스브라켓은 세트 수를 수량으로 쓰고 적요 `2set/이름`. 품목장에 없으면 0원.
        if client == "유앤":
            for line in _unit_accessory_lines(o):
                name = line["품명"]
                hit = M.find_named_item(name, client) or {}
                unit = int(hit.get("단가") or 0)
                qty = float(line["수량"]) if line["단위"] == "set" else 0.0
                amount = int((Decimal(unit) * Decimal(str(qty))).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                vat = int((Decimal(amount) / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                write_row([dstr, voucher_no, 3, "외출", erp_client, erp_client, "", "",
                           hit.get("관리코드") or name, hit.get("품명") or name, hit.get("규격") or "",
                           qty, unit, amount, vat, accessory_memo(line["표시수량"]), "", ""])

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
        pack_window_count = blind_window_count + (roll_window_count if client in BONO_LIKE_CLIENTS else 0)
        if ((not is_roll_order or client in BONO_LIKE_CLIENTS) and pack_window_count
                and (_needs_window_packing(o) or explicit_pack)):
            pack = BONO_PACKING_ITEM if client in BONO_LIKE_CLIENTS else PACKING_ITEM
            manual_pack_qty = o.get("_manual_packing_qty")
            try:
                manual_pack_qty = int(float(manual_pack_qty)) if manual_pack_qty not in (None, "") else None
            except (TypeError, ValueError):
                manual_pack_qty = None
            pack_qty = max(1, manual_pack_qty) if manual_pack_qty else pack_window_count
            amount = int(pack["단가"] * pack_qty)
            vat = int((Decimal(amount) / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
            values = [dstr, voucher_no, 3, "외출", erp_client, erp_client, "", "",
                      pack["품명"], pack["품명"], pack["규격"], float(pack_qty),
                      pack["단가"], amount, vat, order_name, "", ""]
            write_row(values)

        # 휴안·유앤·보노·인천)트루 배송/주소 행.
        # 트루는 제품만 EDI에 나오고 주소가 빠지는 문제가 있었으므로 같은 배송행 구조를 사용한다.
        has_true_accessory = client == "인천)트루" and bool(o.get("_true_accessories"))
        cleaned_notice = clean_delivery_notice(delivery.get("전달사항"))
        has_shipping_info = any(str(delivery.get(k) or "").strip() for k in ("주소", "수령인", "연락처", "발신")) or bool(cleaned_notice)
        if (total_valid_count or has_true_accessory) and has_shipping_info and _effective_delivery_mode(o) in {"택배", "화물"}:
            method = str(delivery.get("방식") or "택배").strip()
            is_freight = "화물" in method
            bono_like = client in BONO_LIKE_CLIENTS
            if is_freight and bono_like:
                # 보노 계열: 첫 줄 `지점 - 받는사람`, 다음 ** 줄에 전화번호(앞에 / 없음).
                address = _freight_ledger_text({**delivery, "연락처": None})
            else:
                address = (_freight_ledger_text(delivery) if is_freight
                           else _erp_address_text(delivery.get("주소")))
            prepaid = str(delivery.get("선불착불") or "").strip()
            is_collect = prepaid == "착불"
            is_prepaid = prepaid == "선불"
            if is_freight:
                if is_collect:
                    ship_code, ship_spec = "#화물(*파손무책)", "화물/착불"
                elif is_prepaid:
                    ship_code, ship_spec = "#화물(선불★/파손무책/청구分)", "화물/선불"
                else:
                    ship_code, ship_spec = "#화물(*파손무책)", "화물"
            else:
                if is_collect:
                    ship_code, ship_spec = "#택배(*파손무책/대신)", "택배/착불"
                elif is_prepaid:
                    ship_code, ship_spec = "#택배(선불★/파손무책/청구分)", "택배/선불"
                else:
                    ship_code, ship_spec = "#택배(*파손무책/대신)", "택배"

            address_lines = (wrap_erp_text(address) if address else [""]) if is_freight \
                else _erp_address_lines(address)
            values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                      "", "", ship_code, ship_code, ship_spec, 1, 0, 0, 0,   # 화물/택배 행 수량 1 (** 줄은 0)
                      address_lines[0], "", ""]
            write_row(values)
            for memo_line in address_lines[1:]:
                values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                          "", "", "**", "**", "", 0, 0, 0, 0,
                          memo_line, "", ""]
                write_row(values)

            if is_freight and bono_like and str(delivery.get("연락처") or "").strip():
                write_row([dstr, voucher_no, 3, "외출", erp_client, erp_client,
                           "", "", "**", "**", "", 0, 0, 0, 0,
                           str(delivery.get("연락처")).strip(), "", ""])
            # 화물 주소행 자체에 `지점 - 받는사람 / 전화번호`가 모두 들어가므로
            # 같은 수령인/전화번호를 다음 줄에 중복 기록하지 않는다. 택배만 별도 수령인 행을 유지한다.
            if not is_freight:
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
                              "", "", "└ 발 신:", "└ 발 신:", "", 0, 0, 0, 0,
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

        # 제주 롤/콤비 택배·화물 출고는 주소 입력 여부와 관계없이 주의 품목을 1행 추가한다.
        if (is_roll_order and str(client or "").startswith("제주)")
                and _effective_delivery_mode(o) in {"택배", "화물"}):
            caution = "파손주의 스티커 부착"
            values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                      "", "", caution, caution, "", 1.0, 0, 0, 0,
                      "", "", ""]
            write_row(values, special_style)

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


# 장부 역변환 때 세로 병합된 칸을 아래 행에도 같은 값으로 채우는 열 범위(J 특이 ~ L 기재사항2).
# 사전점검/장부가 같은 기재사항을 세로로 병합하므로 병합 아래 행이 빈 값으로 읽히면 안 된다.
# I열(길이)은 같은 값끼리만 병합하므로(기본 길이로 비운 칸은 병합하지 않음) 함께 채운다.
_FILL_MIN_COL, _FILL_MAX_COL = 9, 12
# FitOrder DI 장부의 유형별 시트 이름. 시트 첫 행에 상호가 없을 수 있다.
_DI_SHEET_NAMES = {name for name, _ in DI_SHEET_GROUPS} | {"DI 기타"}


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
        try:
            # 병합 정보(merged_cells)를 읽으려면 formatting_info가 필요하다.
            book = xlrd.open_workbook(filename=None if contents is not None else str(source),
                                      file_contents=contents, formatting_info=True)
        except Exception:
            book = xlrd.open_workbook(filename=None if contents is not None else str(source),
                                      file_contents=contents)

        def make_info(ws):
            filled = {}
            for rlo, rhi, clo, chi in getattr(ws, "merged_cells", None) or []:
                if rhi - rlo > 1:
                    for c in range(max(clo + 1, _FILL_MIN_COL), min(chi, _FILL_MAX_COL) + 1):
                        top = ws.cell_value(rlo, c - 1)
                        for rr in range(rlo + 2, rhi + 1):
                            filled[(rr, c)] = top

            def get_cell(row, col):
                if (row, col) in filled:
                    return filled[(row, col)]
                if row < 1 or col < 1 or row > ws.nrows or col > ws.ncols:
                    return None
                return ws.cell_value(row - 1, col - 1)
            hr = find_header(get_cell, ws.nrows)
            return (ws.name, ws.nrows, get_cell, hr) if hr else None

        # DI 유형별 시트(C 원코드 …)를 먼저 읽고, 같은 파일의 `장부` 시트(배송·포장 행)를 이어서 읽는다.
        preferred = ([ws for ws in book.sheets()
                      if re.sub(r"\s*\(\d+\)$", "", ws.name) in _DI_SHEET_NAMES]
                     + [ws for ws in book.sheets()
                        if ws.name == "장부" or re.fullmatch(r"장부 \(\d+\)", ws.name)])
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
            filled = {}
            for rng in ws.merged_cells.ranges:
                if rng.max_row > rng.min_row:
                    for c in range(max(rng.min_col, _FILL_MIN_COL), min(rng.max_col, _FILL_MAX_COL) + 1):
                        top = ws.cell(rng.min_row, c).value
                        for rr in range(rng.min_row + 1, rng.max_row + 1):
                            filled[(rr, c)] = top
            get_cell = lambda row, col: filled[(row, col)] if (row, col) in filled else ws.cell(row, col).value
            hr = find_header(get_cell, ws.max_row)
            return (ws.title, ws.max_row, get_cell, hr) if hr else None

        # DI 유형별 시트(C 원코드 …)를 먼저 읽고, 같은 파일의 `장부` 시트(배송·포장 행)를 이어서 읽는다.
        preferred = ([ws for ws in wb.worksheets
                      if re.sub(r"\s*\(\d+\)$", "", ws.title) in _DI_SHEET_NAMES]
                     + [ws for ws in wb.worksheets
                        if ws.title == "장부" or re.fullmatch(r"장부 \(\d+\)", ws.title)])
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
        if re.sub(r"\s*\(\d+\)$", "", sheet_name) in _DI_SHEET_NAMES:
            # DI 유형별 시트는 서로 다른 주문 묶음이다. 앞 시트 주문에 이어 붙이지 않는다.
            current = None
            prev_color = prev_height = None
            prev_di_recipient_ctx = None
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

            # MIX 구성행(규격 칸=코드, 수량 칸=`) 108cm`)은 별도 제품이 아니라 바로 위 MIX 품목의 구성 정보다.
            if current is not None and current.get("items") and height in (None, ""):
                width_token = re.sub(r"\.0$", "", str(width or "").strip())
                qty_token = str(qty or "").strip()
                if (re.fullmatch(r"\d{3,4}(?:FP|P)?", width_token, re.I)
                        and re.fullmatch(r"\)?\s*\d+(?:\.\d+)?\s*cm", qty_token, re.I)):
                    length = float(re.search(r"\d+(?:\.\d+)?", qty_token).group(0))
                    current["items"][-1].setdefault("_ledger_mix_parts", []).append(
                        {"코드": width_token.upper(), "길이": length})
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
                        if text and current.get("거래처") == "아지트" and not re.search(r"\d", text):
                            # 아지트 ☆시온가공소/☆금빛커텐은 배송 주소가 아니라 가공소 표시다.
                            # 주문 묶음 끝 표시는 주소행과 같이 유지한다.
                            current["_az_place"] = text
                            current["_split_next_product"] = True
                            continue
                        if text:
                            if not current.get("_ledger_delivery_label"):
                                # 거래처 기본값이 아니라 장부 모양으로 배송 방식을 되돌린다.
                                # 화물 주소행은 `지점 - 받는사람 / 전화`처럼 한 줄에 받는사람/전화가 있고,
                                # 택배는 주소 다음 행에 이름·전화번호가 따로 있다.
                                looks_freight = bool(
                                    re.search(r"01\d[-.\s]?\d{3,4}[-.\s]?\d{4}", text)
                                    or re.search(r"\S\s+-\s+\S", text))
                                delivery["방식"] = "화물" if looks_freight else "택배"
                            if "화물" in str(delivery.get("방식") or ""):
                                branch, receiver, phone = parse_freight_ledger_text(text)
                                delivery["화물지점"] = branch
                                delivery["주소"] = branch
                                if receiver:
                                    delivery["수령인"] = receiver
                                if phone:
                                    delivery["연락처"] = phone
                            else:
                                delivery["주소"] = text
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
                                    # 받는사람은 배송 수령인일 뿐 주문 고객명(기재사항 이름)을 덮어쓰지 않는다.
                                    delivery["수령인"] = receiver
                                    if current.get("거래처") in BONO_STYLE_CLIENTS:
                                        compact = receiver.replace(" ", "")
                                        vendor = str(current.get("거래처") or "").replace(" ", "")
                                        if vendor and vendor in compact:
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
                if client in {"SP", "JL"}:
                    # 상호 칸은 `SP   9`, `JL   (17-9)`처럼 거래처명과 주문번호 표시가 함께 있다.
                    # 거래처명을 알아봤더라도 뒤의 번호를 주문번호 후보로 살린다.
                    tail = re.sub(rf"^{re.escape(client)}", "", raw_client_mark).strip()
                    tail = re.sub(r"^[（(]\s*|\s*[）)]$", "", tail).strip()
                    if tail and tail.upper() != "K" and re.fullmatch(r"[A-Za-z0-9-]+", tail):
                        marker_order_no = tail
                if not client and current is not None:
                    # 같은 업체를 연속 작성할 때 상호 대신 (K)/(F) 같은 내부표시만
                    # 쓰거나, SP/JL은 주문번호만 쓰는 장부를 다시 읽을 수 있게 한다.
                    paren_mark = re.fullmatch(r"\(([A-Za-z0-9-]+)\)", raw_client_mark)
                    if paren_mark:
                        client = current.get("거래처")
                        marker_only = True
                        # JL `(17-9)`, SP `(9)`처럼 괄호 안 숫자는 주문번호 표시다.
                        if current.get("거래처") in {"JL", "SP"} and paren_mark.group(1).upper() != "K":
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
            if current is None and re.sub(r"\s*\(\d+\)$", "", sheet_name) in _DI_SHEET_NAMES:
                current = {"거래처": "DI", "주문번호": None, "고객명": None,
                           "전체기재사항": None, "배송": {}, "items": []}
                orders.append(current)
                prev_color = prev_height = None
                prev_di_recipient_ctx = None
            if current is None:
                errors.append(f"{sheet_name} {excel_row}행: 주문의 첫 행에 상호가 없습니다.")
                continue
            ship_mode = re.search(r"\((택배|화물|배달|내사)\)", str(values[12] or ""))
            if ship_mode:
                # 출고일 옆 (택배)/(화물) 표시가 있으면 그것이 가장 확실한 배송 방식이다.
                current.setdefault("배송", {})["방식"] = ship_mode.group(1)
                current["_ledger_delivery_label"] = True

            # 규격 없이 색상 칸에 부속명만 있는 부속행: 휴안 노피스(3)/2EA, 유앤 노피스브라켓, 수동 추가부속(석고앙카2).
            accessory_text = str(color_cell or "").strip()
            if (accessory_text and width in (None, "") and height in (None, "")
                    and not parse_ledger_color(accessory_text)[0]
                    and not holding_accessory_from_text(accessory_text)):
                qty_text = str(qty or "").strip()
                client_now = current.get("거래처")
                huan_m = re.fullmatch(r"노피스\((\d+)\)", accessory_text)
                unit_m = re.fullmatch(r"(.+?)\s*\((\d+)\)", accessory_text)   # 예전 형식: 품명 (피스수)
                set_m = re.fullmatch(r"[Xx]?\s*(\d+)\s*set", qty_text, re.I)
                ea_m = re.fullmatch(r"(\d+)\s*EA", qty_text, re.I)
                if client_now == "휴안" and huan_m:
                    # 예전 장부의 `2EA`도 set 단위로 읽는다(2026-09-14).
                    qm = re.search(r"\d+", qty_text)
                    current.setdefault("_huan_accessories", []).append(
                        {"표시": accessory_text, "수량": f"{qm.group(0) if qm else 1}set"})
                elif client_now == "유앤" and ea_m:
                    current.setdefault("_unit_accessories", []).append(
                        {"품명": accessory_text, "수량EA": int(ea_m.group(1))})
                elif client_now == "유앤" and set_m:
                    current.setdefault("_unit_accessories", []).append(
                        {"품명": unit_m.group(1).strip(), "세트": int(set_m.group(1)),
                         "피스수": int(unit_m.group(2))} if unit_m and qty_text.upper().startswith("X")
                        else {"품명": accessory_text, "세트": int(set_m.group(1))})
                else:
                    current["_추가부속"] = "/".join(_split_note_parts(current.get("_추가부속"), accessory_text))
                rows.append({"_특수": "부속", "문구": accessory_text, "수량": qty})
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
                jo_parts = _split_note_parts(note1)
                # 숫자가 있는 첫 값만 주문번호다. 주문번호 없는 주문의 기재사항1(수령인 이름 등)은 메모로 둔다.
                if jo_parts and re.search(r"\d", jo_parts[0]) and \
                        current.get("주문번호") in (None, "", jo_parts[0]):
                    current["주문번호"] = jo_parts[0]
                    jo_parts = jo_parts[1:]
                note1 = "/".join(_split_note_parts(*jo_parts, _jo_clean_note_text(note2))) or None
                note2 = None
            elif current.get("거래처") in BONO_STYLE_CLIENTS:
                # v90 장부 역변환: 기재사항1의 오더명과 기재사항2의 창별 시공위치를 분리한다.
                # 기재사항1이 세로 병합되어 두 번째 행부터 비어 있어도 시공위치를 버리지 않는다.
                n1 = [x.strip() for x in str(note1 or "").split("/") if x.strip()]
                n2 = [x.strip() for x in str(note2 or "").split("/") if x.strip()]
                rest = []
                if n1:
                    # 오더명(예: 부천67327013)은 숫자를 포함한다. 이름만 있는 기재사항은 오더명으로 쓰지 않는다.
                    if not current.get("주문번호") and re.search(r"\d", n1[0]):
                        current["주문번호"] = n1[0]
                    # 단일창에서는 `오더명/시공위치`가 기재사항1 한 칸에 합쳐질 수 있다.
                    if current.get("주문번호") and n1[0] == str(current.get("주문번호")):
                        rest.extend(n1[1:])
                    else:
                        rest.extend(n1)
                rest.extend(n2)
                note1 = None
                note2 = "/".join(rest) or None
            elif current.get("거래처") == "SP":
                sp_parts = _split_note_parts(note1)
                sp_no = str(current.get("주문번호") or "")
                if re.fullmatch(r"\d+", sp_no):
                    # 상호 칸에는 주문번호 끝 숫자만 적힌다(C-7 → 7). 기재사항의 전체 주문번호로 복원한다.
                    full_no = next((p for p in sp_parts if re.fullmatch(rf"[A-Za-z0-9]+-0*{sp_no}", p)), None)
                    if full_no:
                        current["주문번호"] = full_no
                notices = [p for p in sp_parts if "공지" in p]
                if notices:
                    # 공지는 장부 첫 행에만 1회 적히므로 주문 공통값으로 되돌린다(EDI는 모든 창에 적용).
                    current["전체기재사항"] = "/".join(
                        _split_note_parts(current.get("전체기재사항"), *notices))
                # 기재사항1은 발주번호 + 공통기재사항이다. 발주번호·공지를 뺀 나머지는
                # 아래 _restore_sp_ledger_notes 에서 공통/창별로 나눈다(2026-09-18).
                order_no_now = str(current.get("주문번호") or "")
                note1 = "/".join(p for p in sp_parts
                                 if "공지" not in p and p != order_no_now) or None
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

            item = {"품목코드": code, "종류": kind, "타입": type_,
                    "색상원문": color_text,   # `B MIX 029FP`의 MIX 표시 등 원래 품목 칸을 보존
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
            # 원래 장부 칸 값: 홀딩 장부를 사전점검으로 불러올 때 작동방식/개수/레일 복원용
            item.update({"_ledger_color_text": color_text, "_ledger_qty_text": qty,
                         "_ledger_direction_text": direction, "_ledger_length_text": length})
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
        if _order.get("거래처") == "SP":
            _restore_sp_ledger_notes(_order)
        if _order.get("거래처") == "JL":
            _restore_jl_ledger_notes(_order)
        elif _order.get("거래처") == "DU":
            _restore_du_ledger_customer(_order)
        # 구성행은 같은 MIX 묶음의 마지막 창 아래에만 있으므로 앞쪽 같은 규격 창에도 같은 구성을 준다.
        carried = None
        for item in reversed(_order.get("items") or []):
            own = item.pop("_ledger_mix_parts", None)
            if own:
                carried = (own, item.get("_mix_codes"), item.get("세로"))
            if carried and (own or (item.get("_mix_codes") == carried[1]
                                    and item.get("세로") == carried[2])):
                item["_mix_parts"] = [dict(p) for p in carried[0]]
                item["_mix_parts_manual"] = True
            else:
                carried = None
        expand_same_size_directions(_order)
    return rows, orders, errors


# 설치장소(방 이름) 단어. AI 판독 안내문(prompt.py 8번)의 목록과 같다.
ROOM_WORDS = ("주방픽스", "드레스룸", "다용도실", "작은방", "입구방", "알파룸", "아이방",
              "주방", "거실", "안방", "큰방", "베란다", "현관", "서재", "침실", "복도", "욕실")
_ROOM_RX = re.compile(rf"^(?:{'|'.join(ROOM_WORDS)})(?:\s*\d+(?:-\d+)?)?$")


def split_room_parts(value):
    """`현장B/베란다` → (['현장B'], ['베란다']). 방 이름(거실, 거실1-1 등)만 설치장소로 나눈다."""
    notes, rooms = [], []
    for part in _split_note_parts(value):
        (rooms if _ROOM_RX.match(part) else notes).append(part)
    return notes, rooms


def _restore_du_ledger_customer(order):
    """두창 장부 기재사항1의 첫 값은 고객명(없으면 수령인)이다.

    모든 창에 공통인 첫 값만 고객명으로 되돌리고 창 메모에서는 뺀다. 창마다 이름이 다르면(최수령/정수령)
    고객명으로 쓰지 않고 각 창의 메모로 남긴다.
    """
    items = [it for it in order.get("items") or [] if not it.get("예외품목")]
    if not items or order.get("고객명"):
        return
    part_lists = [_split_note_parts(it.get("기재사항")) for it in items]
    first = part_lists[0][0] if part_lists[0] else None
    if not first or not all(first in parts for parts in part_lists[1:]):
        return
    order["고객명"] = first
    for it, parts in zip(items, part_lists):
        it["기재사항"] = "/".join(p for p in parts if p != first) or None


def _restore_sp_ledger_notes(order):
    """SP 장부 기재사항1을 공통/창별로 나눈다(2026-09-18).

    기재사항1은 원래 `발주번호 + 공통기재사항`이지만, 공통이 하나도 없는 주문에서는
    창별 메모가 기재사항1로 올라온다. 모든 창에 같은 값만 공통으로 되돌리고
    창마다 다른 값은 그 창의 메모로 남긴다.
    """
    items = [it for it in order.get("items") or [] if not it.get("예외품목")]
    if not items:
        return
    part_lists = [_split_note_parts(it.get("기재사항")) for it in items]
    shared = [p for p in part_lists[0] if all(p in parts for parts in part_lists[1:])]
    if shared:
        order["전체기재사항"] = "/".join(
            _split_note_parts(order.get("전체기재사항"), *shared))
    for it, parts in zip(items, part_lists):
        it["기재사항"] = "/".join(p for p in parts if p not in shared) or None


def _restore_jl_ledger_notes(order):
    """JL 장부 기재사항1을 주문 공통(피스/고객명/공통기재)과 창별 메모로 되돌린다.

    JL 기재사항1은 주문 공통값이라 모든 창에 같다. 모든 창에 공통인 값만 공통으로 보고
    (첫 후보를 고객명), 창마다 다른 값(예: 최수령/정수령)은 그 창의 메모로 남긴다.
    """
    items = [it for it in order.get("items") or [] if not it.get("예외품목")]
    if not items:
        return
    part_lists = [_split_note_parts(it.get("기재사항")) for it in items]
    shared = [p for p in part_lists[0] if all(p in parts for parts in part_lists[1:])]
    for part in shared:
        candidate = _jl_customer_name(part)
        if (not candidate or "피스" in candidate or JL_PACKAGING.search(candidate)
                or is_client_staff_name("JL", candidate)
                or re.fullmatch(r"손\d{2,3}", candidate) or candidate in {"틀안", "포장X", "브라켓"}):
            continue
        order["고객명"] = candidate
        break
    common = [p for p in shared if _jl_customer_name(p) != order.get("고객명")]
    if common:
        order["전체기재사항"] = "/".join(_split_note_parts(order.get("전체기재사항"), *common))
    for it, parts in zip(items, part_lists):
        it["기재사항"] = "/".join(p for p in parts if p not in shared) or None


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
    name_only = re.sub(r"^\s*[BHRCS]\s+", "", s, flags=re.I)
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
            # 키의 존재 자체가 수동 override 신호다. 사용자가 빈칸으로 지운 경우도
            # 원문에서 길이를 다시 자동 복구하지 않는다.
            it["_manual_handle_length"] = n
            it["손잡이길이"] = n
            if yeon:
                it["연창"] = True
    elif field == "특이":
        text = str(v or "")
        if order.get("거래처") == "DI":
            # 대일의 자동 특이사항은 비워 두되 사람이 직접 수정한 값은 보존한다.
            # '긴급'은 기존 규칙대로 특이 칸이 아니라 기재사항1/정렬 신호로 처리한다.
            if "긴급" in text:
                it["_DI긴급"] = True
            cleaned = re.sub(r"\b긴급\b\s*[-–—_=~.·ㆍ]*", " ", text)
            cleaned = re.sub(r"\s+", " ", cleaned).strip(" /-–—_=~.·ㆍ")
            it["_manual_special_override"] = True
            it["_수동특이"] = cleaned or None
            it["연창"] = False
            _propagate_di_urgent_order(order)
        else:
            if not is_holding:
                it["연창"] = "#" in text
            it["_수동특이"] = re.sub(r"(?:#|틀안)", " ", text).strip() or None
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

