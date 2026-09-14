"""사용자 사전점검 수정값(override).

원칙: 사용자가 사전점검에서 입력한 값은 자동판독/규칙 결과보다 항상 우선한다.
- 모든 셀 수정은 그 셀의 override로 저장되어, 규칙을 다시 계산해도 입력한 글자 그대로 표시된다.
- 빈 값("")도 '사용자가 지운 값'으로 유지한다.
- 원본 필드와 1:1로 대응하는 셀(가로, 세로, 주소 등)은 원본 필드에도 반영해
  다른 규칙이 수정값을 기준으로 계산하게 한다.
- 계산으로 만들어지는 셀(기재사항1/2, 수령인+연락처 행)은 원본 필드를 바꾸지 않는다.
- 수정은 그 행에만 적용된다. 상호/출고일은 주문의 첫 품목행 값이다.

저장 위치(주문 dict 안에 함께 저장되므로 행 삭제·Undo와 같이 움직인다):
- item["_override"]                 = {컬럼: 값}          품목행
- order["_override_head"]           = {"상호"/"출고일": 값} 첫 품목행
- order["_override_rows"][종류]      = {컬럼: 값}          포장/주소/수령인/전달 행
"""
from __future__ import annotations

ITEM_KEY = "_override"
HEAD_KEY = "_override_head"
SPECIAL_KEY = "_override_rows"

HEAD_COLUMNS = {"상호", "출고일"}

ITEM_SOURCE_FIELDS = {
    "색상": "품목코드", "가로": "가로", "세로": "세로", "수량": "창개수",
    "방향": "손잡이방향", "길이": "손잡이길이", "특이": "특이",
}

SPECIAL_SOURCE_FIELDS = {
    ("address", "가로"): "주소",
    ("address", "기재사항1"): "선불착불",
    ("delivery_notice", "가로"): "전달사항",
}


def _row_kind(row):
    return None if row.get("_특수") is None else row.get("_detail_kind")


def apply_cell_edit(orders, row, column, value):
    """사전점검 셀 수정을 주문 데이터에 기록한다."""
    oi = row.get("_oi")
    if column == "X" or not isinstance(oi, int) or not (0 <= oi < len(orders)):
        return
    order = orders[oi]
    kind = _row_kind(row)

    if kind is None:
        ii = row.get("_ii")
        items = order.get("items") or []
        if not isinstance(ii, int) or not (0 <= ii < len(items)):
            return
        if column in HEAD_COLUMNS and ii == 0:
            order.setdefault(HEAD_KEY, {})[column] = value
            return
        item = items[ii]
        item.setdefault(ITEM_KEY, {})[column] = value
        if column in ITEM_SOURCE_FIELDS:
            item[ITEM_SOURCE_FIELDS[column]] = value
        return

    if kind == "packing" and column == "수량":
        # 기존 동작 유지: 비우거나 0이면 포장행 제외.
        order["_manual_packing_enabled"] = str(value).strip() not in {"", "0", "False", "false"}
    order.setdefault(SPECIAL_KEY, {}).setdefault(kind, {})[column] = value
    field = SPECIAL_SOURCE_FIELDS.get((kind, column))
    if field:
        delivery = order.setdefault("배송", {})
        delivery[field] = str(value).removeprefix("전달: ").strip() if field == "전달사항" else value


def overrides_for_row(order, row, item=None):
    kind = _row_kind(row)
    if kind is None:
        values = dict((item or {}).get(ITEM_KEY) or {})
        if row.get("_ii") == 0:
            values.update(order.get(HEAD_KEY) or {})
        return values
    return dict((order.get(SPECIAL_KEY) or {}).get(kind) or {})


def apply_overrides(row, order, item=None):
    """자동 계산된 행에 사용자 수정값을 마지막으로 덮어쓴다."""
    values = overrides_for_row(order, row, item)
    if values:
        row.update(values)
        row["_overridden"] = sorted(values)
    return row
