from __future__ import annotations
import math, re

HOLDING_FEATURE_ENABLED = True

def looks_like_holding_operation(value):
    s = str(value or "").strip().replace(" ","")
    return bool(re.fullmatch(r"(?:편|양|양개|양자석|1/\d+)", s))

def normalize_holding_operation(value):
    s = str(value or "").strip().replace("양개","양")
    return s or None

def holding_product_from_text(text):
    s = str(text or "")
    return {"code":"H","name":"홀딩도어"} if any(k in s for k in ("홀딩","홀딩도어","자바라")) or re.search(r"(^|\s)H(\s|$)", s) else None

def holding_accessory_from_text(text):
    s = str(text or "")
    if "레일연결" in s: return {"코드":"H레일연결부속","표시":"레일연결부속"}
    if "라운드" in s: return {"코드":"H라운드부속","표시":"라운드부속"}
    if "상하로라" in s: return {"코드":"H상하로라","표시":"+상하로라"}
    if "자석바" in s: return {"코드":"H추가비용(+자석바1)","표시":"자석바"}
    return None

def accessory_qty(value):
    try: return max(1, int(float(value)))
    except: return 1

def holding_calc(width_cm, height_cm, count=1):
    w = float(width_cm) / 100
    h = max(float(height_cm) / 100, 1.5)
    charge = max(w * h, 2.5)
    return round(charge * max(1, int(count)), 3)

def holding_magnet_charge_qty(magnet_bars):
    try: n = int(magnet_bars)
    except: n = 0
    return max(0, n - 1)

def holding_ledger_text(item):
    op = normalize_holding_operation(item.get("작동방식") or item.get("수량"))
    return f"H {op}".strip() if op else "H"
