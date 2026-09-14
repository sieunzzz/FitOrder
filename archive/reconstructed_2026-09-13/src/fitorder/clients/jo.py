"""JO 제이원.

RULES_MASTER 규칙:
- [구현] 기재1 = 주문번호만(경영박사 포함).
- [구현] 3창 이상일 때만 기재2.
- [부분] 기재2 = 나머지 정보. 현재는 개인창/공통/고객명 중 첫 값 하나만 사용.
- [미구현] 표 구조 3창→2창 오인식 주의(파서).
"""
from .base import ClientRule


class JO(ClientRule):
    name = "JO"

    def notes(self, order, item, item_index, total_items):
        note1 = self.order_no(order) or None
        note2 = None
        if total_items >= 3:
            note2 = self.personal(item) or self.common(order) or self.customer(order) or None
        return note1, note2
