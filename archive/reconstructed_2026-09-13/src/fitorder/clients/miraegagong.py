"""미래가공.

RULES_MASTER 규칙:
- 보노와 동일한 발주/기재/포장 규칙 → Bono를 상속.
- 단 거래처명·관리코드·ERP 상호는 각 업체 유지.
"""
from .bono import Bono


class Miraegagong(Bono):
    name = "미래가공"
