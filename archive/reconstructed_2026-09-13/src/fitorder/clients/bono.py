"""보노(이끌림).

RULES_MASTER 규칙:
- [구현] 받는사람 ≠ 보노 → 이름/상호를 기재1.
- [부분] 받는사람이 보노 → 지점 표기(예: 안산초지). `order["지점"]`을 채우는 파서가 없음.
- [확인 필요] 기재2 = 오더명/시공위치(예: `부천67327013/거실1-1`). 현재는 주문번호/개인창.
- [미구현] ERP에도 동일.
- [미구현] 포장비는 창당 `포장비용(B/R/C/S/HC)` 코드 체계.
"""
from .base import ClientRule


class Bono(ClientRule):
    name = "보노"

    def notes(self, order, item, item_index, total_items):
        delivery = order.get("배송") or {}
        receiver = str(delivery.get("수령인") or self.customer(order) or "").strip()
        note1 = None
        if receiver and "보노" not in receiver.replace(" ", ""):
            note1 = receiver
        elif "지점" in order:
            note1 = str(order.get("지점") or "").strip() or None
        note2_parts = [x for x in (self.order_no(order), self.personal(item)) if x]
        return note1, "/".join(note2_parts) or None
