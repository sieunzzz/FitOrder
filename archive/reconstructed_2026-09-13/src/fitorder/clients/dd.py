"""DD 대동. 기재사항은 현재 일반 업체 규칙.

RULES_MASTER 규칙:
- [미구현·충돌] 기재사항1이 여러 번 반복되지 않게 한 주문에서 공통값 1회 중심으로 출력.
  현재는 모든 품목행에 기재1이 반복된다(전 업체 공통 동작).
"""
from .base import ClientRule


class DD(ClientRule):
    name = "DD"
