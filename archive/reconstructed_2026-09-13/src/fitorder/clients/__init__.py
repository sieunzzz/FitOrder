"""거래처별 규칙 모음.

새 업체 추가: 파일을 만들고 ClientRule을 상속한 클래스를 아래 `_RULES`에 등록한다.
등록되지 않은 거래처(미확인 등)는 일반 업체 규칙(ClientRule)을 사용한다.
"""
from __future__ import annotations

from .base import ClientRule
from .bono import Bono
from .dd import DD
from .di import DI
from .general import DU, MS, RT, WT, Ajit, UniTNS
from .huan import Huan
from .incheon_true import IncheonTrue
from .jl import JL
from .jo import JO
from .miraegagong import Miraegagong
from .ruimt import Ruimt
from .sp import SP

_RULES: dict[str, ClientRule] = {
    rule.name: rule
    for rule in (
        DI(), Huan(), JL(), DU(), IncheonTrue(), WT(), DD(), MS(), UniTNS(),
        Ajit(), Ruimt(), JO(), Bono(), Miraegagong(), RT(), SP(),
    )
}

_DEFAULT = ClientRule()


def get_client_rule(client: str | None) -> ClientRule:
    return _RULES.get(client or "", _DEFAULT)


def registered_clients() -> list[str]:
    return list(_RULES)


__all__ = ["ClientRule", "get_client_rule", "registered_clients"]
