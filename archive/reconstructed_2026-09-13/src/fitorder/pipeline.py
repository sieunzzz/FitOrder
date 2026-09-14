"""주문 → 장부행 계산.

사전점검 UI, 장부, 작업지시서, ERP가 모두 이 결과 하나를 사용한다.
계산 순서: 원본 주문 → 공통/업체 규칙 → 사용자 수정값(override, 항상 마지막).
"""
from __future__ import annotations

from .overrides import apply_overrides
from .rules import display_size, effective_delivery_mode, needs_packing, note_fields

LEDGER_COLUMNS = ["상호","색상","가로","X","세로","수량","방향","길이","특이","기재사항1","기재사항2","출고일"]
WEEKDAYS = ["월","화","수","목","금","토","일"]


def ship_label(ship_date):
    if hasattr(ship_date, "weekday"):
        return WEEKDAYS[ship_date.weekday()]
    return str(ship_date) if ship_date else ""


def product_color(item):
    code = str(item.get("품목코드") or "").strip()
    kind = str(item.get("종류") or "").strip()
    if not code: return kind
    if kind and kind not in code:
        return f"B {kind} {code}".strip()
    return code


def _item_rows(oi, order, label):
    items = order.get("items") or []
    rows = []
    for ii, item in enumerate(items):
        n1, n2 = note_fields(order, item, ii, len(items))
        row = {
            "상호": order.get("거래처") if ii == 0 else "",
            "색상": product_color(item),
            "가로": display_size(item.get("가로")),
            "X": "X",
            "세로": display_size(item.get("세로")),
            "수량": item.get("창개수") or item.get("수량") or 1,
            "방향": item.get("손잡이방향") or "",
            "길이": item.get("손잡이길이") or "",
            "특이": item.get("특이") or ("#" if item.get("연창") else ""),
            "기재사항1": n1 or "",
            "기재사항2": n2 or "",
            "출고일": label if ii == 0 else "",
            "_oi": oi, "_ii": ii, "_특수": None,
            "_row_key": ("item", oi, ii),
            "_품명": item.get("품목코드") or item.get("종류"),
        }
        rows.append(apply_overrides(row, order, item))
    return rows


def _special_row(oi, kind, detail_kind, **values):
    row = {"상호":"","색상":"","가로":"","X":"","세로":"","수량":"",
           "방향":"","길이":"","특이":"","기재사항1":"","기재사항2":"","출고일":""}
    row.update(values)
    row.update({"_oi": oi, "_ii": None, "_특수": kind, "_detail_kind": detail_kind, "_row_key": (kind, oi)})
    return row


def _special_rows(oi, order):
    items = order.get("items") or []
    d = order.get("배송") or {}
    rows = []
    if needs_packing(order):
        rows.append(_special_row(oi, "packing", "packing", 색상="#", 가로="포장비용", 수량=len(items) or 1))
    if effective_delivery_mode(order) in {"택배","화물"}:
        rows.append(_special_row(oi, "address", "address", 색상="☆", 가로=d.get("주소") or "",
                                 기재사항1=d.get("선불착불") or ""))
        if d.get("수령인") or d.get("연락처"):
            rows.append(_special_row(oi, "recipient", "recipient_contact",
                                     가로=" ".join(x for x in (d.get("수령인"), d.get("연락처")) if x)))
        if d.get("전달사항"):
            rows.append(_special_row(oi, "notice", "delivery_notice", 가로="전달: " + str(d["전달사항"])))
    return [apply_overrides(r, order) for r in rows]


def build_rows(orders, ship_date=None):
    label = ship_label(ship_date)
    rows = []
    for oi, order in enumerate(orders):
        rows.extend(_item_rows(oi, order, label))
        rows.extend(_special_rows(oi, order))
    return rows
