"""JL.

RULES_MASTER 규칙:
- [구현] 공통 `피스/고객명`, 개인창은 기재사항2.
- [미구현·충돌] 우상단 괄호 발주번호 필수(괄호 포함). 현재 장부 어디에도 출력되지 않음.
- [미구현] 원코드 위치 고정 아님(파서).
- [미구현] 같은 사람 2건 병합 시 `피스/이름`.
- [미구현] `dw` 생략.
- [미구현] 손잡이길이 → 피스 → 손잡이 순서.
"""
from .base import ClientRule


class JL(ClientRule):
    name = "JL"

    def notes(self, order, item, item_index, total_items):
        return self.join_unique((self.common(order), self.customer(order))), self.personal(item) or None
