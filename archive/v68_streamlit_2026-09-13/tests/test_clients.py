from clients import ClientRules, get_client_rules, registered_clients
from rules import CLIENT_INFO


def test_every_client_has_its_own_rules_object():
    # CLIENT_INFO 에 업체를 추가하면 clients/ 에도 반드시 등록한다.
    assert sorted(registered_clients()) == sorted(CLIENT_INFO)
    objs = [get_client_rules(c) for c in CLIENT_INFO]
    assert all(type(o) is not ClientRules for o in objs)
    assert len({type(o) for o in objs}) == len(CLIENT_INFO)


def test_unknown_client_uses_common_rules():
    assert type(get_client_rules("미확인")) is ClientRules
    assert type(get_client_rules(None)) is ClientRules
