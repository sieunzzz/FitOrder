"""유앤 유앤아이티엔에스(UNITNS) — v68 동작.

장부: 노피스/피스 부속을 `품명 (피스수)` + `X{세트}set` 별도 행으로 표시.
(부속 인식은 parsers.parse_unitns 에 있다.)
"""
from .base import ClientRules


class Unit(ClientRules):
    code = "유앤"

    def ledger_accessory_rows(self, order):
        rows = []
        for accessory in order.get("_unit_accessories") or []:
            label = str(accessory.get("품명") or "").strip()
            if accessory.get("피스수"):
                label += f" ({accessory['피스수']})"
            rows.append({"_특수": "부속", "문구": label,
                         "수량": f"X{accessory.get('세트', 1)}set"})
        return rows
