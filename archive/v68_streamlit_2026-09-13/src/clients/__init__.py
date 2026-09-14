"""거래처별 규칙.

`get_client_rules(거래처코드)` 로 업체 규칙 객체를 얻는다. 등록되지 않은 거래처는 ClientRules(공통).
새 업체: 파일을 만들고 ClientRules 를 상속한 클래스를 `_RULES` 에 등록한다.
"""
from .azit import Azit
from .base import ClientRules
from .bono import Bono
from .di import DI
from .du import DU
from .general import DD, MS, WT, Midas
from .huan import Huan
from .incheon_true import IncheonTrue
from .jl import JL
from .jo import JO
from .miraegagong import Miraegagong
from .rt import RT
from .sp import SP
from .unit import Unit

_RULES = {
    rule.code: rule
    for rule in (
        DI(), SP(), Huan(), Midas(), DU(), RT(), JO(), JL(), DD(), Unit(), Azit(),
        WT(), IncheonTrue(), Miraegagong(), Bono(), MS(),
    )
}
_DEFAULT = ClientRules()


def get_client_rules(client):
    return _RULES.get(client, _DEFAULT)


def registered_clients():
    return list(_RULES)
