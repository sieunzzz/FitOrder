"""DI 대일.

RULES_MASTER 규칙:
- [구현] 기재사항1만 사용.
- [미구현] 2코드 이상 혼합 → MIX. `330(ㅇㄴ)-102(SJ)`도 MIX.
- [미구현] `B MIX B 원코드 MIX` → 기재사항 `102+520` 형태.
- [미구현] `긴급 ---` → 텍스트는 `긴급`만, 장부 제일 상단에 긴급건 1회 집계.
- [미구현] 포장 550원 (창당/㎡당 확인 필요). 홀딩 포장은 청구㎡×550원.
"""
from .base import ClientRule


class DI(ClientRule):
    name = "DI"

    def notes(self, order, item, item_index, total_items):
        return self.join_unique((self.order_no(order), self.customer(order), self.common(order))), None
