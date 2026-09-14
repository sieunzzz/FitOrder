"""사용자 확정 규칙(2026-09-13)
- 비닐 포장 관련 문구는 장부·작업지시서·사전점검·경영박사 EDI 어디에도 적지 않는다.
- 미더스는 블라인드 거래처에서 제외한다(홀딩도어 미더스 (K)는 유지).
"""
import json
import tempfile
from pathlib import Path

import pytest

import _cases
import desktop_workflow as dw
import output
from holding import HOLDING_CLIENTS
from master_registry import get_master_for_mode
from rules import clean_delivery_notice, strip_vinyl_text

VINYL_WORDS = ("비닐", "비 닐")


def test_strip_vinyl_text_keeps_other_text_unchanged():
    assert strip_vinyl_text("피스/공학1관") == "피스/공학1관"
    assert strip_vinyl_text(120) == 120
    assert strip_vinyl_text("피스/비닐포장 따로 해주세요/공학1관") == "피스/공학1관"
    assert strip_vinyl_text("창별 비닐포장") is None
    assert strip_vinyl_text("겉비닐\n거실") == "거실"
    assert clean_delivery_notice("문앞/비닐 겉면 기재 : 홍길동") == "문앞"


def _add_vinyl(order):
    order["전체기재사항"] = "/".join(x for x in (order.get("전체기재사항"), "비닐포장 따로 해주세요") if x)
    order.setdefault("배송", {})["전달사항"] = "문앞/비닐 겉면 기재 : 홍길동"
    for i, it in enumerate(order.get("items") or []):
        it["기재사항"] = "/".join(x for x in (it.get("기재사항"), "겉비닐") if x)
        if i == 0:
            it["설치장소"] = "창별 비닐포장"
            it["_수동특이"] = "비닐X"
    return order


def _all_strings(value):
    if isinstance(value, dict):
        for v in value.values():
            yield from _all_strings(v)
    elif isinstance(value, (list, tuple)):
        for v in value:
            yield from _all_strings(v)
    elif isinstance(value, str):
        yield value


CASES = [("blind", c) for c in ("DU", "JL", "휴안", "DI", "보노", "RT", "JO", "SP")] + \
        [("holding", "휴안"), ("roll_combo", "천안)채원")]


@pytest.mark.parametrize("mode,client", CASES)
def test_vinyl_text_never_reaches_any_output(tmp_path, mode, client):
    M = get_master_for_mode(mode)
    orders = [_add_vinyl(o) for o in _cases.MODES[mode][1](client)]
    preview, _ = dw.build_preview_rows(orders, _cases.SHIP, M)
    # 사전점검에서 사용자가 비닐 문구를 직접 입력해도 출력에는 남지 않는다.
    product = next(r for r in preview if r.get("_특수") is None)
    dw.apply_preview_cell_edit(orders, product, "기재사항1", "사용자메모/비닐포장")
    preview, _ = dw.build_preview_rows(orders, _cases.SHIP, M)
    shown = [dw.row_display_values(r, i + 1) for i, r in enumerate(preview)]

    paths, _ = dw.generate_outputs(orders, _cases.SHIP, M, tmp_path / "out", tag="v")
    books = {k: _cases.read_book(paths[k]) for k in ("장부", "작업지시서")}
    import xlrd
    sheet = xlrd.open_workbook(str(paths["경영박사"])).sheet_by_index(0)
    edi = [sheet.row_values(r) for r in range(sheet.nrows)]

    texts = list(_all_strings(json.loads(json.dumps([shown, books, edi], ensure_ascii=False, default=str))))
    leaked = [t for t in texts if any(w in t for w in VINYL_WORDS)]
    assert leaked == []
    assert any("사용자메모" in t for t in texts)


def test_midas_is_not_a_blind_client_but_stays_holding():
    assert "M" not in dw.BLIND_CLIENTS
    assert "M" in HOLDING_CLIENTS


def test_blind_workspace_client_list_excludes_midas(qapp_module):
    import qt_app
    ws = qt_app.OrderWorkspace("blind")
    try:
        assert "M" not in ws.mode_clients
        assert [ws.client_combo.itemData(i) for i in range(ws.client_combo.count())].count("M") == 0
    finally:
        ws.deleteLater()


@pytest.fixture(scope="module")
def qapp_module():
    import os
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtWidgets import QApplication
    return QApplication.instance() or QApplication([])
