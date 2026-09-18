"""스페이스(SP) 손글씨 작업일지 규칙 — 2026-09-18 사용자 확정."""
from pathlib import Path

import copy

import xlrd

import _cases
import desktop_workflow as dw
import extract
import output
from master_registry import get_master_for_mode

MB = get_master_for_mode("blind")


def _order(items, order_no="3-17", **extra):
    order = {"거래처": "SP", "주문번호": order_no, "고객명": "우미린아파트",
             "배송": {}, "items": items}
    order.update(extra)
    order = dw._normalize_ingested_order(order, Path("t.png"), "blind")
    order["_ship_date"] = _cases.SHIP.isoformat()
    return order


def _item(w=100, h=150, **extra):
    return {"품목코드": "102", "종류": "투코드", "타입": "C자",
            "가로": w, "세로": h, **extra}


def _rows(orders):
    return dw.build_output_rows(orders, _cases.SHIP, M=MB, sort_di=True)


def _products(rows):
    return [r for r in rows if r.get("_특수") is None]


def _edi(orders, path):
    for o in orders:
        output.synchronize_order_for_outputs(o)
    output.build_erp(orders, MB, path, None)
    sheet = xlrd.open_workbook(str(path)).sheet_by_index(0)
    return [sheet.row_values(r) for r in range(sheet.nrows)]


# ─────────────────── 주문번호 보정 ───────────────────
# 주문번호는 `현장전달사항` 칸에 `< 2-12 >`로 적히고 `주문날짜의 일 - 동그라미 숫자`다.
def test_bracket_number_overrides_model_reading():
    """꺾쇠 표기가 있으면 모델이 다르게 읽어도 그 값을 쓴다."""
    data = {"주문번호": "2-2", "전체기재사항": "<2-12>", "items": []}
    extract._fix_sp_order_no(data)
    assert data["주문번호"] == "2-12"
    assert data["전체기재사항"] is None      # 번호는 기재사항에 남기지 않는다


def test_bracket_number_found_in_item_text_and_keeps_other_notes():
    data = {"주문번호": None, "전체기재사항": "호계푸르지오/<2-12>", "items": []}
    extract._fix_sp_order_no(data)
    assert (data["주문번호"], data["전체기재사항"]) == ("2-12", "호계푸르지오")


def test_number_without_bracket_is_normalized_and_deduped():
    data = {"주문번호": " 02 - 05 ", "전체기재사항": "주문번호 2-5/피스", "items": []}
    extract._fix_sp_order_no(data)
    assert (data["주문번호"], data["전체기재사항"]) == ("2-5", "피스")


def test_date_cell_text_is_not_a_note():
    data = {"주문번호": "1-24", "전체기재사항": "발송날짜: 9월 4일", "items": []}
    extract._fix_sp_order_no(data)
    assert (data["주문번호"], data["전체기재사항"]) == ("1-24", None)


def test_site_number_is_not_mistaken_for_an_order_number():
    """`<111-1806>`은 동·호수다. 일(day)이 31을 넘으므로 주문번호가 아니다."""
    data = {"주문번호": "23", "전체기재사항": "한강숲중흥 <111-1806>", "items": []}
    extract._fix_sp_order_no(data)
    assert (data["주문번호"], data["전체기재사항"]) == ("23", "한강숲중흥 <111-1806>")


# ─────────────────── 장부 ───────────────────
def test_ledger_mark_is_order_tail_in_parentheses():
    """장부 상호 칸은 `SP` 옆에 주문번호 끝 숫자를 괄호로 적는다 → `SP   (17)`."""
    rows = _rows([_order([_item()])])
    assert _products(rows)[0]["내부표시"] == "(17)"


def test_ledger_note1_is_order_number_plus_common_note2_is_personal():
    """기재사항1 = 발주번호 + 공통, 기재사항2 = 창별 개별 정보."""
    order = _order([_item(설치장소="작은방"), _item(설치장소="거실")],
                   전체기재사항="피스")
    products = _products(_rows([order]))
    assert all(r["기재사항"] == "3-17/피스/우미린아파트" for r in products)
    assert [r["기재사항2"] for r in products] == ["작은방", "거실"]


# ─────────────────── 경영박사 적요 ───────────────────
def test_edi_memo_order_is_piece_number_detail_common(tmp_path):
    """적요 순서: 피스 → 번호 → 세부 기재사항 → 공통 기재사항."""
    order = _order([_item(설치장소="작은방")], 전체기재사항="피스")
    edi = _edi([order], tmp_path / "e.xls")
    memo = next(r[15] for r in edi if str(r[8]).startswith("B102"))
    assert memo.endswith("피스/3-17/작은방/우미린아파트")


def test_edi_matches_ledger_roundtrip(tmp_path):
    """발주서에서 바로 만든 EDI와 장부를 거쳐 만든 EDI가 같아야 한다.

    주문번호가 장부 상호 칸의 `(17)`로만 남으므로 되돌릴 때 잃기 쉽다.
    """
    def make():
        return [_order([_item(설치장소="작은방"), _item(설치장소="거실")],
                       전체기재사항="피스")]

    direct = _edi(make(), tmp_path / "a.xls")
    ledger = tmp_path / "장부.xlsx"
    output.build_ledger(_rows(make()), ledger, header_date=_cases.HEADER)
    _, read, _ = output.read_ledger(str(ledger), MB, ledger.name)
    ready = dw.ledger_orders_for_precheck(read, ledger, "blind", _cases.SHIP)
    via = _edi(ready, tmp_path / "b.xls")
    assert [r[15] for r in via] == [r[15] for r in direct]
    assert any("3-17" in str(r[15]) for r in via)


# ─────────────────── 수량 1/2 = 가로 2등분 ───────────────────
def test_half_quantity_splits_the_width_into_two_windows():
    """스페이스 수량 `1/2`는 손잡이 분할이 아니라 가로를 나눈다는 뜻이다."""
    rows = _products(_rows([_order([_item(w=200, h=150, 수량="1/2")])]))
    assert len(rows) == 2                       # 200짜리 1창이 아니라 100짜리 2창
    assert [r["가로"] for r in rows] == [100, 100]
    assert rows[0]["세로"] == 150                # 둘째 줄 세로는 장부 반복표시 `"`


def test_split_windows_are_billed_separately(tmp_path):
    """나눈 창은 각각 청구한다(창당 최소 1.5㎡가 각각 걸린다)."""
    edi = _edi([_order([_item(w=200, h=150, 수량="1/2")])], tmp_path / "e.xls")
    products = [r for r in edi if str(r[8]).startswith("B102")]
    assert len(products) == 2


# ─────────────────── 여러 장 합치기 ───────────────────
def test_pages_with_the_same_order_number_merge_into_one():
    a = _order([_item(w=100)], order_no="3-17")
    b = _order([_item(w=110)], order_no="3-17")
    merged = dw.merge_sp_page_orders([a, b])
    assert len(merged) == 1
    assert [it["가로"] for it in merged[0]["items"]] == [100, 110]


def test_pages_with_different_numbers_stay_apart():
    orders = [_order([_item()], order_no="3-17"), _order([_item()], order_no="3-18")]
    assert len(dw.merge_sp_page_orders(orders)) == 2


def test_pages_without_a_number_are_left_alone():
    """번호를 못 읽은 장은 남의 주문에 섞지 않는다."""
    orders = [_order([_item()], order_no=None), _order([_item()], order_no=None)]
    assert len(dw.merge_sp_page_orders(orders)) == 2
