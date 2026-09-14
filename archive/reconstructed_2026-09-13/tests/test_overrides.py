"""사용자 사전점검 수정값이 규칙 재계산 후에도 유지되고, 장부·ERP까지 전달되는지 확인."""
import csv

import pytest

from _cases import ALL_CLIENTS, SHIP_DATE, _full_order, all_orders
from fitorder.desktop_workflow import apply_preview_cell_edit, preview_statuses
from fitorder.output import build_erp, build_ledger, read_ledger
from fitorder.pipeline import build_rows


def edit(orders, index, column, value):
    rows = build_rows(orders, SHIP_DATE)
    apply_preview_cell_edit(orders, rows[index], column, value)
    return build_rows(orders, SHIP_DATE)


def row_of(rows, kind):
    return next(i for i, r in enumerate(rows) if r.get("_detail_kind") == kind)


def one_item_order(client):
    order = _full_order(client)
    order["items"] = order["items"][:1]
    return order


@pytest.mark.parametrize("client", ALL_CLIENTS)
def test_note_edits_are_kept_for_every_client(client):
    orders = [one_item_order(client)]
    edit(orders, 0, "기재사항1", "사용자기재1")
    rows = edit(orders, 0, "기재사항2", "사용자기재2")
    assert rows[0]["기재사항1"] == "사용자기재1"
    assert rows[0]["기재사항2"] == "사용자기재2"


def test_note_edit_applies_only_to_that_row():
    orders = [_full_order("DU")]
    before = build_rows(orders, SHIP_DATE)
    after = edit(orders, 0, "기재사항1", "첫행만")
    assert after[0]["기재사항1"] == "첫행만"
    assert after[1]["기재사항1"] == before[1]["기재사항1"]


def test_cleared_cells_stay_empty():
    orders = [{"거래처": "DU", "items": [
        {"품목코드": "102", "가로": 100, "세로": 100, "수량": 5, "연창": True, "기재사항": "안방"}]}]
    for column in ("특이", "수량", "기재사항2"):
        rows = edit(orders, 0, column, "")
        assert rows[0][column] == "", column


def test_head_columns_edit_first_item_row():
    orders = [_full_order("JL")]
    edit(orders, 0, "상호", "JL(발주1)")
    rows = edit(orders, 0, "출고일", "화")
    assert (rows[0]["상호"], rows[0]["출고일"]) == ("JL(발주1)", "화")
    assert (rows[1]["상호"], rows[1]["출고일"]) == ("", "")


def test_source_fields_follow_one_to_one_edits():
    orders = [_full_order("DU")]
    edit(orders, 0, "가로", "120")
    rows = edit(orders, 0, "색상", "105")
    item = orders[0]["items"][0]
    assert (item["가로"], item["품목코드"]) == ("120", "105")
    assert (rows[0]["가로"], rows[0]["색상"]) == ("120", "105")


def test_recipient_edit_does_not_duplicate_contact():
    orders = [_full_order("보노")]
    rows = build_rows(orders, SHIP_DATE)
    rows = edit(orders, row_of(rows, "recipient_contact"), "가로", "박수령 010-1111-2222")
    assert rows[row_of(rows, "recipient_contact")]["가로"] == "박수령 010-1111-2222"
    assert orders[0]["배송"]["수령인"] == "김수령"


def test_address_and_notice_edits_update_delivery():
    orders = [_full_order("DU")]
    rows = build_rows(orders, SHIP_DATE)
    rows = edit(orders, row_of(rows, "address"), "가로", "서울시 수정로 2")
    rows = edit(orders, row_of(rows, "delivery_notice"), "가로", "전달: 경비실")
    assert orders[0]["배송"]["주소"] == "서울시 수정로 2"
    assert orders[0]["배송"]["전달사항"] == "경비실"
    assert rows[row_of(rows, "delivery_notice")]["가로"] == "전달: 경비실"


def test_packing_zero_hides_packing_row_as_before():
    orders = [_full_order("DU")]
    rows = build_rows(orders, SHIP_DATE)
    rows = edit(orders, row_of(rows, "packing"), "수량", "0")
    assert all(r.get("_detail_kind") != "packing" for r in rows)


def test_status_tooltip_shows_user_edit():
    orders = [_full_order("DU")]
    rows = edit(orders, 0, "기재사항2", "수정")
    assert "사용자 수정: 기재사항2" in preview_statuses(rows, orders)[0][2]


def _erp_rows(orders, tmp_path):
    path = build_erp(orders, SHIP_DATE, tmp_path / "erp.csv")
    with path.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def test_edits_reach_ledger_and_erp(tmp_path):
    orders = [_full_order("JO")]
    for column, value in (("가로", "77"), ("수량", "4"), ("기재사항1", "장부수정1"), ("기재사항2", "장부수정2")):
        edit(orders, 0, column, value)
    ledger = read_ledger(build_ledger(orders, SHIP_DATE, tmp_path / "ledger.xlsx"))[0]
    erp = _erp_rows(orders, tmp_path)[0]
    assert (ledger["가로"], ledger["수량"], ledger["기재사항1"], ledger["기재사항2"]) == ("77", "4", "장부수정1", "장부수정2")
    assert (erp["규격"], erp["수량"], erp["기재사항1"], erp["기재사항2"]) == ("77x124", "4", "장부수정1", "장부수정2")


def test_ledger_and_erp_item_values_are_consistent(tmp_path):
    orders = all_orders()
    ledger = [r for r in build_rows(orders, SHIP_DATE) if r.get("_특수") is None]
    erp = _erp_rows(orders, tmp_path)
    assert len(ledger) == len(erp)
    for lr, er in zip(ledger, erp):
        assert er["규격"] == f"{lr['가로']}x{lr['세로']}"
        assert er["수량"] == str(lr["수량"])
        assert (er["기재사항1"], er["기재사항2"]) == (lr["기재사항1"], lr["기재사항2"])
