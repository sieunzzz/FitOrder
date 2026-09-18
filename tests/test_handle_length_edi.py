"""기본 길이가 아닌 손잡이길이는 전 업체 전산 적요에 넣는다(2026-09-18 사용자 확정).

장부 길이 칸에 적히는 값과 같은 값을 쓰므로, 발주서로 넣든 장부로 불러오든 결과가 같다.
"""
import copy
from pathlib import Path

import pytest
import xlrd

import _cases
import desktop_workflow as dw
import output
from master_registry import get_master_for_mode

M = get_master_for_mode("blind")


def _order(client, handle, kind="투코드", h=150):
    order = {"거래처": client, "주문번호": "7", "고객명": "임지애", "배송": {},
             "items": [{"품목코드": "102", "종류": kind, "타입": "C자", "가로": 100, "세로": h,
                        "손잡이방향": "좌", "손잡이길이": handle, "기재사항": "거실"}]}
    order = dw._normalize_ingested_order(order, Path("t.png"), "blind")
    order["_ship_date"] = _cases.SHIP.isoformat()
    return order


def _memos(orders, path):
    for o in orders:
        output.synchronize_order_for_outputs(o)
    output.build_erp(orders, M, path, None)
    sheet = xlrd.open_workbook(str(path)).sheet_by_index(0)
    return [sheet.row_values(r)[15] for r in range(sheet.nrows)
            if str(sheet.row_values(r)[8]).startswith("B")]


def _via_ledger(orders, tmp_path):
    rows = dw.build_output_rows(copy.deepcopy(orders), _cases.SHIP, M=M, sort_di=True)
    ledger = tmp_path / "장부.xlsx"
    output.build_ledger(rows, ledger, header_date=_cases.HEADER)
    _, read, _ = output.read_ledger(str(ledger), M, ledger.name)
    ready = dw.ledger_orders_for_precheck(read, ledger, "blind", _cases.SHIP)
    return _memos(ready, tmp_path / "b.xls")


@pytest.mark.parametrize("client", ["DD", "휴안", "JO", "DU", "WT", "MS", "유앤", "인천)트루", "SP", "RT"])
def test_non_default_handle_in_edi_for_every_client(tmp_path, client):
    orders = [_order(client, 120)]                       # 투코드 세로 150 기본값은 100
    direct = _memos(copy.deepcopy(orders), tmp_path / "a.xls")
    assert direct
    if client == "SP":
        # 스페이스는 `피스/주문번호` 다음에 줄길이가 온다(업체 고유 순서).
        assert all("손120" in m for m in direct)
    else:
        assert all(m.split(" ", 1)[1].startswith("손120") for m in direct)
    assert _via_ledger(orders, tmp_path) == direct       # 장부로 불러와도 같다


@pytest.mark.parametrize("client,kind,handle", [("DD", "투코드", 100), ("휴안", "원코드", 130),
                                                ("JO", "투코드", 100)])
def test_default_handle_stays_out_of_edi(tmp_path, client, kind, handle):
    orders = [_order(client, handle, kind=kind)]
    direct = _memos(copy.deepcopy(orders), tmp_path / "a.xls")
    assert direct and not any("손" in m for m in direct)
    assert _via_ledger(orders, tmp_path) == direct


def test_manual_handle_edit_is_written_even_if_default(tmp_path):
    orders = [_order("DD", 120)]
    orders[0]["items"][0]["_manual_handle_length"] = 100     # 사람이 기본값으로 직접 고침
    assert all("손100" in m for m in _memos(orders, tmp_path / "a.xls"))
