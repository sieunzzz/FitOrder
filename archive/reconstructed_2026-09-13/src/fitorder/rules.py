from __future__ import annotations
import math, re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .clients import get_client_rule

CLIENTS = [
    "DI","휴안","JL","DU","인천)트루","WT","DD","MS","유앤아이티엔에스",
    "아지트","루임트","JO","보노","미래가공","RT","SP",
]

CLIENT_ALIASES = {
    "대일":"DI","대일산업":"DI","DI":"DI",
    "휴안":"휴안","제이엘":"JL","JL":"JL","두창":"DU","DU":"DU",
    "트루":"인천)트루","인천)트루":"인천)트루","윈도우투모로우":"WT","WT":"WT",
    "대동":"DD","DD":"DD","미성텍스":"MS","MS":"MS",
    "루임트":"루임트","제이원":"JO","JO":"JO","이끌림":"보노","보노":"보노",
    "미래가공":"미래가공","RT":"RT","스페이스":"SP","SP":"SP",
}

# (ERP상호, 내부표시, 기본배송, 메모)
CLIENT_INFO = {
    "DI": ("DI", "", None, "기재사항1만"),
    "휴안": ("휴안", "", "택배", "기재사항1만"),
    "JL": ("JL", "", None, "발주번호 괄호"),
    "DU": ("DU", "", None, ""),
    "인천)트루": ("인천)트루", "K", "택배", ""),
    "WT": ("WT", "", None, ""),
    "DD": ("DD", "", None, ""),
    "MS": ("MS", "", None, ""),
    "유앤아이티엔에스": ("유앤아이티엔에스", "", None, ""),
    "아지트": ("아지트", "", None, ""),
    "루임트": ("루임트", "", None, ""),
    "JO": ("JO", "K", None, "3창 이상일 때 기재2"),
    "보노": ("보노", "K", "택배", "이끌림 규칙"),
    "미래가공": ("미래가공", "K", "택배", "보노 동일 규칙"),
    "RT": ("RT", "", None, ""),
    "SP": ("SP", "", None, ""),
}

ROLL_COMBO_CLIENTS = {
    "안산)보노": {},
    "미래가공": {},
    "천안)채원": {},
    "세일인테리어": {},
    "진천)충북혁신": {},
    "음성)충북혁신": {},
    "논공)현대종합장식": {},
    "원주)창과방": {},
    "피카소": {"swap_wh": True},
    "제주)우림": {"jeju": True},
    "제주)중문": {"jeju": True},
}

TRUE_ACCESSORY_PRODUCTS = {
    "B 원코드 브라켓": {"code":"B 원코드 브라켓", "qty":1},
    "노피스(1)": {"code":"노피스(1)", "qty":1},
    "노피스(2)": {"code":"노피스(2)", "qty":1},
}

def normalize_client(value: str | None) -> str | None:
    s = str(value or "").strip()
    if not s:
        return None
    return CLIENT_ALIASES.get(s, s)

def display_size(value):
    if value in (None, ""):
        return ""
    try:
        f = float(value)
        return str(int(f)) if f.is_integer() else str(f).rstrip("0").rstrip(".")
    except Exception:
        return str(value).strip()

def normalize_handle(value):
    if value in (None, ""):
        return None
    s = str(value).strip()
    s = s.replace("½","1/2").replace("⅓","1/3").replace("¼","1/4")
    m = re.search(r"(\d{2,3})", s)
    return int(m.group(1)) if m and "/" not in s else s

def default_handle_length(height):
    try:
        h = float(height)
    except Exception:
        return None
    if h <= 160: return 100
    if h <= 190: return 120
    if h <= 230: return 140
    return 150

def clean_delivery_notice(text):
    s = str(text or "").strip()
    return s or None

def effective_delivery_mode(order: dict) -> str | None:
    delivery = order.get("배송") or {}
    raw = str(delivery.get("방식") or "").strip()
    if raw:
        for mode in ("택배","화물","배달","내사"):
            if mode in raw:
                return mode
        return raw
    info = CLIENT_INFO.get(order.get("거래처"))
    return info[2] if info else None

def needs_packing(order: dict) -> bool:
    if "_manual_packing_enabled" in order:
        return bool(order["_manual_packing_enabled"])
    return effective_delivery_mode(order) in {"택배","화물"}

def packing_rule(client: str, area_m2: float, windows: int = 1):
    client = normalize_client(client) or client
    if client == "SP":
        return 0
    if client == "DI":
        return 550 * max(1, windows)
    if client in {"보노","미래가공"}:
        return 1000 * max(1, windows)   # 실제 품목코드는 master에서 치환
    return int(math.ceil(max(0, area_m2) * 1000))

def color_tokens(text: str):
    """UI/output 부분색상용 토큰. P/FP는 숫자코드만 blue."""
    s = str(text or "")
    spans = []
    for m in re.finditer(r"원코드_?", s):
        spans.append((m.start(), m.end(), "green"))
    for m in re.finditer(r"셔터", s):
        spans.append((m.start(), m.end(), "red"))
    for m in re.finditer(r"\d{3,4}(?:FP|P)\b", s, re.I):
        spans.append((m.start(), m.end(), "blue"))
    return spans

def dimension_flags(product_group: str, kind: str | None, width, height):
    flags = {"width_blue": False, "height_blue": False, "confirm": []}
    try: w = float(width) if width not in (None,"") else None
    except: w = None
    try: h = float(height) if height not in (None,"") else None
    except: h = None

    if product_group == "holding":
        if h is not None and h >= 270:
            flags["confirm"].append("홀딩도어 세로 270cm 이상 제작 확인")
        if h is not None and h >= 300:
            flags["height_blue"] = True
        return flags

    if w is not None:
        if kind == "원코드" and w < 36: flags["width_blue"] = True
        if kind in {"투코드","셔터"} and w < 30: flags["width_blue"] = True
    if h is not None and h > 370:
        flags["height_blue"] = True
    return flags

def note_fields(order: dict, item: dict, item_index: int, total_items: int):
    """기재사항1/2. 업체별 규칙은 `fitorder/clients/<업체>.py`에 있다."""
    client = normalize_client(order.get("거래처")) or order.get("거래처")
    return get_client_rule(client).notes(order, item, item_index, total_items)

@dataclass
class Master:
    path: Path | None = None
    def item_exists(self, code: str | None) -> bool:
        return bool(str(code or "").strip())
    def thread_color(self, code: str | None) -> str | None:
        return None
    def client_info(self, client: str):
        return CLIENT_INFO.get(client)
