"""장부 → 경영박사 연결 점검 (2026-09-13).

같은 주문으로
  A) 주문 → 경영박사 EDI
  B) 주문 → 장부 → 장부 불러오기(read_ledger → 사전점검 준비) → 경영박사 EDI
를 만들어 비교한다. 장부를 거쳐도 경영박사 내용이 같아야 한다.

장부 형식상 되돌릴 수 없는 값은 비교에서 제외한다.
- 기본 손잡이 길이: 장부에는 기본 길이를 적지 않으므로 `손130` 같은 기본값 표시는 복원되지 않는다.
- DI: 유형별 시트(C 원코드 …)로 나뉘고 특이(연창 #) 칸이 없어 주문 묶음/연창이 복원되지 않는다.
"""
import copy
import re
from collections import Counter
from pathlib import Path

import pytest
import xlrd

import _cases
import desktop_workflow as dw
import output
from master_registry import get_master_for_mode

EXACT_CLIENTS = ["SP", "휴안", "DU", "RT", "JO", "DD", "유앤", "아지트", "WT", "인천)트루", "MS"]
HANDLE_DEFAULT_CLIENTS = ["JL", "보노", "미래가공"]


def _edi(orders, M, path):
    for o in orders:
        output.synchronize_order_for_outputs(o)
    output.build_erp(orders, M, path, None)
    sheet = xlrd.open_workbook(str(path)).sheet_by_index(0)
    return [tuple(sheet.row_values(r)[8:18]) for r in range(sheet.nrows)]


def _via_ledger(orders, M, mode, tmp_path):
    rows = dw.build_output_rows(copy.deepcopy(orders), _cases.SHIP, M=M, sort_di=True)
    ledger = tmp_path / "장부.xlsx"
    output.build_ledger(rows, ledger, header_date=_cases.HEADER)
    _, read, _ = output.read_ledger(str(ledger), M, ledger.name)
    ready = dw.ledger_orders_for_precheck(read, ledger, mode, _cases.SHIP)
    return _edi(ready, M, tmp_path / "b.xls")


def _drop_handle(rows):
    return [row[:7] + (re.sub(r"(?<=[ /])손\d{2,3}/?", "", str(row[7])).rstrip("/"),) + row[8:]
            for row in rows]


@pytest.mark.parametrize("client", EXACT_CLIENTS + HANDLE_DEFAULT_CLIENTS)
def test_blind_ledger_path_gives_same_edi(tmp_path, client):
    M = get_master_for_mode("blind")
    orders = _cases.blind_orders(client)
    direct = _edi(copy.deepcopy(orders), M, tmp_path / "a.xls")
    via = _via_ledger(orders, M, "blind", tmp_path)
    if client in HANDLE_DEFAULT_CLIENTS:
        direct, via = _drop_handle(direct), _drop_handle(via)
    assert Counter(via) == Counter(direct)


@pytest.mark.parametrize("client", ["M", "휴안", "구미)경남", "DU", "창문애", "한길"])
def test_holding_ledger_path_gives_same_edi(tmp_path, client):
    M = get_master_for_mode("holding")
    orders = _cases.holding_orders(client)
    direct = _edi(copy.deepcopy(orders), M, tmp_path / "a.xls")
    via = _via_ledger(orders, M, "holding", tmp_path)
    # 휴안은 기재사항1 한 칸이라 `베란다/현장B`의 순서만 달라질 수 있다(내용은 같아야 함).
    norm = (lambda rows: [r[:7] + ("/".join(sorted(str(r[7]).split("/"))),) + r[8:] for r in rows])
    assert Counter(norm(via)) == Counter(norm(direct))


def test_di_ledger_products_are_readable(tmp_path):
    M = get_master_for_mode("blind")
    orders = _cases.blind_orders("DI")
    direct = [r for r in _edi(copy.deepcopy(orders), M, tmp_path / "a.xls") if r[8] in (1.0, 2.0)]
    via = [r for r in _via_ledger(orders, M, "blind", tmp_path) if r[8] in (1.0, 2.0)]
    key = (lambda r: (r[0], r[2], r[3], r[4], r[9]))   # 품목/규격/수량/단가/방향
    assert Counter(map(key, via)) == Counter(map(key, direct))
