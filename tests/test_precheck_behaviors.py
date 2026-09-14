"""사전점검 편집·병합·장부 불러오기·출고일 표시의 목표 동작.

1. 직접 작성한 장부 → 사전점검에 불러와 확인 → [파일 생성]으로 장부·작업지시서·EDI
2. 사전점검을 엑셀처럼: 화면 병합 = 장부 병합, 병합 셀 수정은 병합된 모든 행에 반영,
   행을 추가/이동해도 병합 유지, 해제한 병합은 다시 자동 병합되지 않음, 셀 복사/붙여넣기
3. 출고일은 날짜로 고르고 장부·작업지시서·사전점검에는 요일로 표시
"""
import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import _cases  # noqa: E402
import desktop_workflow as dw  # noqa: E402
import output  # noqa: E402
from master_registry import get_master_for_mode  # noqa: E402


# ───────────────────────── 3. 출고일 ─────────────────────────
@pytest.mark.parametrize("mode,client", [("blind", "DU"), ("holding", "휴안"), ("roll_combo", "천안)채원")])
def test_ship_date_is_written_as_weekday(mode, client):
    orders = _cases.MODES[mode][1](client)
    labels = [r["출고일"] for r in dw.build_output_rows(orders, _cases.SHIP) if r.get("출고일")]
    assert labels and all(str(x).split("\n")[0] == "화" for x in labels)


# ───────────────────────── 행 고유 ID ─────────────────────────
def test_row_keys_are_unique_and_survive_row_insert():
    orders = _cases.blind_orders("DU")
    keys = [r["_row_key"] for r in dw.build_output_rows(orders, _cases.SHIP)]
    assert len(keys) == len(set(keys))
    before = {k for k in keys if k[0] == "product"}
    orders[0]["items"].insert(1, dict(orders[0]["items"][0]))   # 복사된 행도 새 ID를 받아야 한다
    keys2 = [r["_row_key"] for r in dw.build_output_rows(orders, _cases.SHIP)]
    assert len(keys2) == len(set(keys2))
    assert before <= set(keys2)


# ───────────────────────── 1. 장부 → 사전점검 ─────────────────────────
def _ledger_file(tmp_path, orders, M):
    rows = dw.build_output_rows(orders, _cases.SHIP, M=M, sort_di=True)
    path = tmp_path / "직접작성_장부.xlsx"
    output.build_ledger(rows, path, header_date=_cases.HEADER)
    return path


@pytest.mark.parametrize("client", ["DU", "JL", "휴안", "보노", "DI"])
def test_blind_ledger_loads_into_precheck_and_exports(tmp_path, client):
    M = get_master_for_mode("blind")
    path = _ledger_file(tmp_path, _cases.blind_orders(client), M)
    _, orders, _ = output.read_ledger(str(path), M, path.name)
    ready = dw.ledger_orders_for_precheck(orders, path, "blind", _cases.SHIP)
    assert ready
    assert all(o["_product_mode"] == "blind" and o["_ship_date"] == _cases.SHIP.isoformat() for o in ready)
    preview, _ = dw.build_preview_rows(ready, _cases.SHIP, M)
    assert any(r.get("_특수") is None for r in preview)
    paths, _ = dw.generate_outputs(ready, _cases.SHIP, M, tmp_path / "out", tag="t")
    assert all(p.exists() for p in paths.values())


def test_holding_ledger_loads_into_precheck_with_products(tmp_path):
    M = get_master_for_mode("holding")
    source = _cases.holding_orders("휴안")
    path = _ledger_file(tmp_path, source, M)
    _, orders, _ = output.read_ledger(str(path), M, path.name)
    ready = dw.ledger_orders_for_precheck(orders, path, "holding", _cases.SHIP)
    items = [it for o in ready for it in o["items"]]
    src_items = [it for o in source for it in o["items"]]
    assert [it.get("_holding_product_name") for it in items] == \
           [it.get("_holding_product_name") for it in src_items]
    assert [it.get("_holding_operation") for it in items if not it.get("_holding_accessory")] == \
           [it.get("_holding_operation") for it in src_items if not it.get("_holding_accessory")]


# ───────────────────────── 2. 사전점검 화면 (PySide6) ─────────────────────────
@pytest.fixture(scope="module")
def qapp():
    from PySide6.QtWidgets import QApplication
    return QApplication.instance() or QApplication([])


@pytest.fixture
def workspace(qapp, monkeypatch):
    from PySide6.QtWidgets import QMessageBox
    import qt_app

    monkeypatch.setattr(QMessageBox, "question",
                        staticmethod(lambda *a, **k: QMessageBox.StandardButton.Yes))
    for name in ("information", "warning", "critical"):
        monkeypatch.setattr(QMessageBox, name,
                            staticmethod(lambda *a, **k: QMessageBox.StandardButton.Ok))
    created = []

    def make(mode, client):
        ws = qt_app.OrderWorkspace(mode)
        ws.orders = _cases.MODES[mode][1](client)
        ws.rebuild_table(preserve=False)
        created.append(ws)
        return ws

    yield make
    for ws in created:
        ws.deleteLater()


def _col():
    import qt_app
    return qt_app.COL


def _select(ws, top, bottom, left, right):
    from PySide6.QtWidgets import QTableWidgetSelectionRange
    ws.table.setCurrentCell(top, left)
    ws.table.clearSelection()
    ws.table.setRangeSelected(QTableWidgetSelectionRange(top, left, bottom, right), True)


def _products(ws):
    return [i for i, r in enumerate(ws.preview_rows) if r.get("_특수") is None]


def _item_by_key(ws, key):
    for order in ws.orders:
        for item in order.get("items") or []:
            if key[0] == "product" and item.get("_uid") == key[2]:
                return item
    raise AssertionError(f"item not found: {key}")


@pytest.mark.parametrize("mode,client", [("blind", "JL"), ("blind", "DD"), ("blind", "보노"),
                                         ("blind", "DI"), ("holding", "휴안"), ("roll_combo", "천안)채원")])
def test_preview_auto_merges_equal_ledger_merges(workspace, mode, client):
    ws = workspace(mode, client)
    col = _col()
    out_to_table = {r["_output_index"]: i for i, r in enumerate(ws.preview_rows)
                    if isinstance(r.get("_output_index"), int)}
    expected = set()
    for o1, o2, c1, c2 in output.ledger_merge_ranges(ws.output_rows):
        t1, t2 = out_to_table[o1], out_to_table[o2]
        if ws.preview_rows[t1].get("_특수") is not None or c1 == col["출고일"]:
            continue
        expected.add((t1, c1, t2 - t1 + 1, c2 - c1 + 1))
    shown = {s for s in ws._current_spans()
             if ws.preview_rows[s[0]].get("_특수") is None and s[1] != col["출고일"]}
    assert shown == expected


def test_edit_on_merged_cell_updates_every_merged_row(workspace):
    ws = workspace("blind", "DD")
    col = _col()
    a, b = _products(ws)[:2]
    keys = [ws.preview_rows[a]["_row_key"], ws.preview_rows[b]["_row_key"]]
    _select(ws, a, b, col["기재사항2"], col["기재사항2"])
    ws.merge_selected()
    assert ws.table.rowSpan(a, col["기재사항2"]) == 2
    ws.table.item(a, col["기재사항2"]).setText("공통수정")
    assert [_item_by_key(ws, k).get("_manual_note2") for k in keys] == ["공통수정", "공통수정"]
    assert ws.table.rowSpan(a, col["기재사항2"]) == 2


def test_delete_on_merged_cell_clears_every_merged_row(workspace):
    ws = workspace("blind", "DD")
    col = _col()
    a, b = _products(ws)[:2]
    keys = [ws.preview_rows[a]["_row_key"], ws.preview_rows[b]["_row_key"]]
    _select(ws, a, b, col["기재사항2"], col["기재사항2"])
    ws.merge_selected()
    _select(ws, a, a, col["기재사항2"], col["기재사항2"])
    ws.delete_selected_cells()
    assert [_item_by_key(ws, k).get("_manual_note2") for k in keys] == [None, None]
    assert all("_manual_note2" in _item_by_key(ws, k) for k in keys)


def test_manual_merge_survives_row_added_below(workspace):
    ws = workspace("blind", "DD")
    col = _col()
    a, b = _products(ws)[:2]
    _select(ws, a, b, col["기재사항2"], col["기재사항2"])
    ws.merge_selected()
    _select(ws, b, b, col["가로"], col["가로"])
    ws.add_row_below()
    assert ws.manual_merges
    assert ws.table.rowSpan(a, col["기재사항2"]) == 2


def test_row_added_inside_manual_merge_extends_it(workspace):
    ws = workspace("blind", "DD")
    col = _col()
    a, b = _products(ws)[:2]
    _select(ws, a, b, col["기재사항2"], col["기재사항2"])
    ws.merge_selected()
    _select(ws, a, a, col["가로"], col["가로"])
    ws.add_row_below()
    assert ws.table.rowSpan(a, col["기재사항2"]) == 3


def test_unmerged_auto_merge_stays_unmerged_after_other_edits(workspace):
    ws = workspace("blind", "JL")
    col = _col()
    a = _products(ws)[0]
    assert ws.table.rowSpan(a, col["기재사항1"]) > 1
    _select(ws, a, a, col["기재사항1"], col["기재사항1"])
    ws.unmerge_selected()
    assert ws.table.rowSpan(a, col["기재사항1"]) == 1
    ws.table.item(a, col["가로"]).setText("151")
    assert ws.table.rowSpan(a, col["기재사항1"]) == 1


def test_merge_selection_overlapping_existing_merge_is_not_ignored(workspace):
    ws = workspace("blind", "DI")
    col = _col()
    a, b = _products(ws)[:2]
    _select(ws, a, b, col["기재사항2"], col["기재사항2"])
    ws.merge_selected()
    assert ws.manual_merges


def test_copy_and_paste_cells_like_excel(workspace, qapp):
    from PySide6.QtWidgets import QApplication
    ws = workspace("blind", "DU")
    col = _col()
    a, b = _products(ws)[:2]
    keys = [ws.preview_rows[a]["_row_key"], ws.preview_rows[b]["_row_key"]]

    _select(ws, a, b, col["가로"], col["가로"])
    ws.copy_selected_cells()
    copied = QApplication.clipboard().text().splitlines()
    assert copied == [ws.table.item(a, col["가로"]).text(), ws.table.item(b, col["가로"]).text()]

    QApplication.clipboard().setText("150\n160")
    _select(ws, a, a, col["세로"], col["세로"])
    ws.paste_cells_from_clipboard()
    assert [_item_by_key(ws, k).get("세로") for k in keys] == [150.0, 160.0]


def test_ledger_button_loads_ledger_into_precheck(workspace, tmp_path):
    ws = workspace("blind", "DU")
    path = _ledger_file(tmp_path, _cases.blind_orders("DU"), ws.M)
    before = len(ws.orders)
    ws.pending_files = [str(path)]
    ws.convert_selected_ledger()
    assert len(ws.orders) > before
    assert ws.pending_files == []
