"""휴안.

RULES_MASTER 규칙:
- [구현] 기재사항1만 사용.
- [미구현] `★이지픽스메탈…` 메시지 → 노피스/석고앙카를 가장 긴 창 첫머리에.

DI와 현재 기재사항 로직이 같지만 업체 규칙이 다르므로 별도 클래스로 유지한다.
"""
from .base import ClientRule


class Huan(ClientRule):
    name = "휴안"

    def notes(self, order, item, item_index, total_items):
        return self.join_unique((self.order_no(order), self.customer(order), self.common(order))), None
