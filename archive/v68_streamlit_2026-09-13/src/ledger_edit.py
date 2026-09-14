"""사전점검 표의 셀 편집값 -> 주문 항목 반영.

2026-09-13 뼈대 정리: v68 output.py 에서 코드 변경 없이 분리.
"""
import re
from datetime import date
from rules import CLIENT_INFO, di_mix_info
from holding import (
    HOLDING_FEATURE_ENABLED, holding_accessory_from_text, holding_product_from_text,
    normalize_holding_operation)

from ledger_rows import _propagate_di_urgent_order, handle_split
from ledger_import import _ledger_mix_codes, parse_handle, parse_ledger_color, parse_ledger_mix


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
