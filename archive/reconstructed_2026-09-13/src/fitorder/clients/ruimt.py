"""루임트. 기재사항은 현재 일반 업체 규칙.

RULES_MASTER 규칙:
- [미구현] 주소 끝 글자 = 받는사람, `더커튼` → `더`.
- [미구현] 선불은 빨간색만.
- [미구현] 착불 생략 가능.
"""
from .base import ClientRule


class Ruimt(ClientRule):
    name = "루임트"
