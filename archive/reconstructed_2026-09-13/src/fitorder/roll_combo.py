from __future__ import annotations
from .rules import ROLL_COMBO_CLIENTS

def normalize_roll_combo_item(client: str, item: dict) -> dict:
    out = dict(item)
    conf = ROLL_COMBO_CLIENTS.get(client, {})
    if conf.get("swap_wh"):
        out["가로"], out["세로"] = out.get("세로"), out.get("가로")
    # 롤/콤비/트리플 손잡이 150은 기본값이므로 출력 생략
    if str(out.get("손잡이길이") or "").strip() in {"150","150.0"}:
        out["손잡이길이"] = None
    out["품명"] = " ".join(str(out.get("품명") or out.get("품목코드") or "").split())
    return out

def billed_area(width_cm, height_cm, count=1):
    w = float(width_cm) / 100
    h = max(float(height_cm), 150.0) / 100
    each = max(w * h, 2.0)
    return round(each * max(1, int(count)), 3)

def fireproof_labels(is_fireproof: bool):
    if not is_fireproof:
        return None, None
    return "<필증>", "▶필증(방염)◀"

def jeju_erp_notice(client: str):
    return "파손주의 스티커 부착" if client in {"제주)우림","제주)중문"} else None
