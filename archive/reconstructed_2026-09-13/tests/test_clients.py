from fitorder.clients import ClientRule, get_client_rule, registered_clients
from fitorder.rules import CLIENTS


def test_every_blind_client_has_its_own_rule_object():
    # 업체를 CLIENTS에 추가하면 규칙 클래스도 반드시 등록해야 한다.
    assert sorted(registered_clients()) == sorted(CLIENTS)
    rules = [get_client_rule(c) for c in CLIENTS]
    assert all(type(r) is not ClientRule for r in rules)
    assert len({id(r) for r in rules}) == len(CLIENTS)


def test_unknown_client_uses_general_rule():
    assert type(get_client_rule("미확인")) is ClientRule
    assert type(get_client_rule(None)) is ClientRule


def test_miraegagong_follows_bono_rules_but_keeps_name():
    bono, mirae = get_client_rule("보노"), get_client_rule("미래가공")
    order = {"주문번호": "부천1", "배송": {"수령인": "홍길동"}}
    item = {"설치장소": "거실1-1"}
    assert mirae.notes(order, item, 0, 1) == bono.notes(order, item, 0, 1)
    assert mirae.name == "미래가공"
