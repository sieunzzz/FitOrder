"""주문 1건 -> 장부 행(dict). 블라인드·홀딩도어, 업체별 기재사항/표시 규칙.

2026-09-13 뼈대 정리: v68 output.py 에서 코드 변경 없이 분리.
"""
import re
from types import SimpleNamespace

from rules import (
    CLIENT_INFO, NO_DITTO, NO_MERGE_PLACE, DI_CODE_ORDER, clean_delivery_notice,
    TRUE_ACCESSORY_PRODUCTS, default_handle_length, ledger_color, normalize_handle)
from holding import (
    accessory_qty, holding_ledger_text, looks_like_holding_operation,
    normalize_holding_operation)

from clients import get_client_rules
from order_text import _has_non_di_mix, _note_text, _prepend_note_part, _split_note_parts


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


def _short_dir(d, client=None):
    # 장부/작업지시서는 좌측만 표기한다. 우측은 기본값이므로 숨기되
    # 원 주문 데이터에는 남겨 경영박사 EDI의 방향코드에는 그대로 사용한다.
    return "좌" if d == "좌" else None


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
    client_rules = get_client_rules(client)
    mark = CLIENT_INFO.get(client, (None, None, None, ""))[3]
    order_no = str(order.get("주문번호") or "").strip()
    # 업체별 내부표시(SP: 주문번호 끝 숫자, JL/두창: 괄호 주문번호) — clients/*.ledger_mark
    mark = client_rules.ledger_mark(order_no, mark)
    rows = []
    source_index = {id(it): idx for idx, it in enumerate(order.get("items", []))}
    items = [it for it in order.get("items", []) if not it.get("예외품목")]
    if client == "DI" and M is not None and sort_di:
        items = sort_di_items(items, M)
    prev_color = prev_h = prev_di_group = None
    prev_joint = False
    prev_di_recipient = None
    common = order.get("전체기재사항") or ""
    # 업체별 전체기재사항 가공(SP: 주문번호·공지 분리) — clients/*.ledger_common
    common, sp_notices = client_rules.ledger_common(common, order_no)
    body = "/".join(x for x in common.split("/") if x and x != "포장비용")
    cust = (order.get("고객명") or "").replace("/", "-").strip()
    cust = client_rules.ledger_customer(cust)

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
        # ── v68 기재사항 분배 규칙 — 업체별: clients/*.ledger_notes ──
        note1, note2 = client_rules.ledger_notes(SimpleNamespace(
            order=order, it=it, i=i, client=client, cust=cust, body=body, common=common,
            order_no=order_no, sp_notices=sp_notices, common_note_parts=common_note_parts,
            total_windows=total_windows, multi_windows=multi_windows, single=single,
            note=note, place=place))
        # 택배/화물은 받는 명칭을 기재사항1에서 확인할 수 있게 한다.
        # JO/JL/DU/보노/DI/휴안은 별도 고정 규칙이 있어 제외한다(prepend_receiver_to_note1=False).
        if client_rules.prepend_receiver_to_note1 and _effective_delivery_mode(order) in {"택배", "화물"}:
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
                else (it.get("창개수") if client_rules.ledger_qty_uses_window_count else None))
            model1 = _short_dir(it.get("손잡이방향"), client)
            model2 = _handle(it.get("종류"), h, it.get("손잡이길이"))
            if client_rules.hide_handle_150 and normalize_handle(it.get("손잡이길이"))[0] == 150:
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
    # 업체 전용 부속 행(트루/유앤) — clients/*.ledger_accessory_rows
    rows.extend(client_rules.ledger_accessory_rows(order))
    if str(order.get("_추가부속") or "").strip():
        rows.append({"_특수": "부속", "문구": str(order.get("_추가부속")).strip(),
                     "수량": None})
    # 업체 전용 특수 행(아지트 ☆가공처) — clients/*.ledger_place_rows
    rows.extend(client_rules.ledger_place_rows(order))
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
        if who or (client_rules.show_prepaid_without_receiver and pay_mark):
            rows.append({"_특수": "", "문구": who or "",
                         "선불": pay_mark})
        delivery_notice = clean_delivery_notice(d.get("전달사항")) or ""
        if delivery_notice:
            rows.append({"_특수": "", "문구": f"전달 : {delivery_notice}"})
        if d.get("발신"):
            rows.append({"_특수": "", "문구": f"발신 : {d['발신']}"})
    return rows
