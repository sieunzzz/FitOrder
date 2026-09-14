"""기재사항/적요 문자열 도우미 — 장부·EDI·장부역변환·업체 규칙(clients/)이 함께 쓴다.

2026-09-13 뼈대 정리: v68 output.py 에서 코드 변경 없이 분리.
"""
import re


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


def _prepend_note_part(note, part):
    """기재사항1 맨 앞에 값을 중복 없이 넣는다."""
    part = str(part or "").strip()
    if not part:
        return note
    parts = [x.strip() for x in str(note or "").split("/") if x.strip()]
    parts = [x for x in parts if x != part]
    return "/".join([part] + parts) or None


def _has_non_di_mix(item, client=None):
    if client == "DI":
        return False
    if item.get("_mix_word"):
        return True
    raw = " ".join(str(item.get(k) or "") for k in
                   ("색상원문", "원문", "기재사항"))
    return bool(re.search(r"(?<![A-Za-z])MIX(?![A-Za-z])", raw, re.I))


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
