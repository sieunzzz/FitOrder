"""SP 스페이스. 기재사항은 현재 일반 업체 규칙.

RULES_MASTER 규칙:
- [확인 필요] 주문번호 동일 = 동일건 (자동 병합 금지와의 관계).
- [미구현] `공지`는 기재사항 맨 앞 1회.
- [부분] 포장 0원(rules.packing_rule, 출력 미연결). 홀딩 포장 없음.
"""
from .base import ClientRule


class SP(ClientRule):
    name = "SP"
