"""수정 전후로 PySide 108 출력이 같은지 확인한다.

실패하면 의도한 변경인지 먼저 확인한다.
의도한 변경일 때만 `python tests/_generate_golden.py <항목>`으로 갱신한다.
"""
import json

import pytest

from _cases import GOLDEN_PATH, SNAPSHOTS


@pytest.fixture(scope="module")
def golden():
    return json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))


@pytest.mark.parametrize("name", list(SNAPSHOTS))
def test_matches_golden(golden, name):
    actual = SNAPSHOTS[name]()
    expected = golden[name]
    assert actual.keys() == expected.keys()
    for key in expected:
        assert actual[key] == expected[key], f"{name}: {key}"
