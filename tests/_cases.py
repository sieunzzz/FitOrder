"""PySide 108 현재 동작을 고정하는 특성(characterization) 테스트 케이스.

- golden 파일은 "정답"이 아니라 "현재 동작"이다. 수정 중 출력이 바뀌면 실패한다.
- 규칙/동작을 의도적으로 바꾼 경우에만 `python tests/_generate_golden.py <항목>`으로 갱신하고
  바뀐 부분을 사용자에게 보고한다.
- 가상의 주문 값만 사용한다(실제 고객 정보 금지). 품목장은 data/master 의 실제 파일을 쓴다.
"""
from __future__ import annotations

import copy
import json
import os
import re
import sys
import tempfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import openpyxl  # noqa: E402
from openpyxl.cell.rich_text import CellRichText  # noqa: E402

import desktop_workflow as dw  # noqa: E402
import output  # noqa: E402
from holding import HOLDING_CLIENTS  # noqa: E402
from master_registry import get_master_for_mode  # noqa: E402
from roll_combo import ROLL_CLIENTS, normalize_roll_order  # noqa: E402

GOLDEN_PATH = Path(__file__).resolve().parent / "golden" / "characterization.json"
SHIP = date(2026, 9, 15)      # 화요일
HEADER = date(2026, 9, 14)
FIXTURE = Path("fixture_order.png")
BLIND_CLIENTS = list(dw.BLIND_CLIENTS) + ["미확인"]


_ROW_ID_RX = re.compile(r"#[oi]\d+")


def _jsonable(value):
    """JSON 비교용 변환. 실행 순서에 따라 번호가 달라지는 행 고유 ID(#o12, #i34)는
    처음 나온 순서대로 다시 번호를 매겨 결과를 결정적으로 만든다."""
    text = json.dumps(value, ensure_ascii=False, default=str)
    mapping = {}

    def canon(match):
        token = match.group(0)
        if token not in mapping:
            mapping[token] = f"#{token[1]}{len(mapping) + 1}"
        return mapping[token]

    return json.loads(_ROW_ID_RX.sub(canon, text))


# ─────────────────────────────────────────────
# 블라인드 주문
# ─────────────────────────────────────────────
def _blind_full(client):
    return {
        "거래처": client, "주문번호": "17-9", "고객명": "홍길동 DW",
        "전체기재사항": "피스/공학1관/공지 금요일 휴무",
        "전체원문": "총 5창", "변경요청": True, "변경문구": "사이즈 이걸로 변경되나요",
        "배송": {"방식": "택배", "선불착불": "선불",
                 "주소": "서울특별시 테스트구 테스트로 12, (테스트동, 가나빌) 101동 202호",
                 "수령인": "김수령", "연락처": "010-0000-0000", "전달사항": "배송전연락/문앞"},
        "_true_accessories": [{"표시": "노피스(2)", "품명": "노피스브라켓(2EA)", "수량": 1}],
        "_unit_accessories": [{"품명": "노피스브라켓(3EA)", "세트": 2, "피스수": 3}],
        "_추가부속": "석고앙카2",
        "items": [
            {"품목코드": "102", "종류": "원코드", "타입": "C자", "가로": 150.0, "세로": 124,
             "창개수": 1, "손잡이방향": "좌", "손잡이길이": 120, "연창": True,
             "설치장소": "거실", "기재사항": "공학1관"},
            {"품목코드": "102", "종류": "원코드", "타입": "C자", "가로": 90, "세로": 124,
             "창개수": 1, "손잡이방향": "우", "손잡이길이": 130, "연창": True,
             "설치장소": "안방", "기재사항": "공학1관"},
            {"품목코드": "029FP", "종류": "투코드", "타입": "C자", "가로": 29.5, "세로": 371,
             "창개수": 2, "좌개수": 1, "우개수": 1, "손잡이길이": 145,
             "기재사항": "공학1관/틀안", "_수동특이": "긴급 ---", "원문": "MIX 029FP"},
            {"품목코드": "200", "종류": "셔터", "타입": "L자", "가로": 80, "세로": 90,
             "수량": "1/2", "설치장소": "주방", "기재사항": "공학1관/줄140"},
        ],
    }


def _blind_minimal(client):
    return {"거래처": client, "고객명": "이고객", "배송": {},
            "items": [{"품목코드": "330", "종류": "투코드", "타입": "C자",
                       "가로": 100, "세로": 150.5, "손잡이방향": "우"}]}


def _blind_cargo(client):
    return {"거래처": client, "주문번호": "C-7", "고객명": "박현장",
            "배송": {"방식": "화물", "선불착불": "착불", "화물지점": "오산삼미", "주소": "오산삼미",
                     "수령인": "보노 안산", "연락처": "010-1111-2222", "발신": "테스트 010-2222-3333"},
            "items": [
                {"품목코드": "200", "종류": "투코드", "타입": "C자", "가로": 100, "세로": 150,
                 "손잡이방향": "좌우", "설치장소": "작은방"},
                {"품목코드": "200", "종류": "투코드", "타입": "C자", "가로": 110, "세로": 150,
                 "창개수": 3, "손잡이방향": "우", "설치장소": "작은방"},
                {"품목코드": None, "종류": "투코드", "타입": "C자", "가로": None, "세로": 130,
                 "예외품목": "수리"},
            ]}


def _blind_mix(client):
    return {"거래처": client, "배송": {"방식": None},
            "items": [
                {"품목코드": "200+290", "_mix_name": "밀크코코아", "_mix_codes": "200+290",
                 "종류": "원코드", "타입": "C자", "가로": 120, "세로": 200, "기재사항": "최수령"},
                {"품목코드": "102+520", "_mix_name": "MIX", "_mix_codes": "102+520",
                 "_generic_mix": True, "종류": "투코드", "타입": "C자", "가로": 60, "세로": 200,
                 "기재사항": "최수령", "_수동특이": "긴급", "손잡이길이": 125},
                {"품목코드": "102", "종류": "투코드", "타입": "C자", "가로": 70, "세로": 100,
                 "기재사항": "정수령"},
                {"품목코드": "102", "종류": "투코드", "타입": "C자", "가로": 75, "세로": 100,
                 "기재사항": "정수령"},
            ]}


def blind_orders(client):
    out = []
    for make in (_blind_full, _blind_minimal, _blind_cargo, _blind_mix):
        order = dw._normalize_ingested_order(make(client), FIXTURE, "blind")
        order["_ship_date"] = SHIP.isoformat()
        out.append(order)
    return out


# ─────────────────────────────────────────────
# 홀딩도어 주문
# ─────────────────────────────────────────────
def _holding_order(client):
    return {"거래처": client, "주문번호": "H-1", "고객명": "제이원 송길수 부장님",
            "배송": {"방식": "택배", "주소": "대구광역시 테스트구 홀딩로 3", "수령인": "홀랜드",
                     "연락처": "010-3333-4444"},
            "items": [
                {"색상원문": "H003 화이트", "품목코드": "H003", "홀딩방식": "양자석", "가로": 88,
                 "세로": 310, "창개수": 1, "기재사항": "현장A",
                 "원문": "H003 화이트 88x310 양자석 +상하로라"},
                {"색상원문": "아이보리", "품목코드": "105", "수량": "1/3", "가로": 250, "세로": 120,
                 "창개수": 2, "설치장소": "베란다", "기재사항": "현장B", "원문": "홀딩 아이보리 250x120 1/3 레일"},
                {"색상원문": "레일연결부속", "원문": "레일연결부속 블랙", "홀딩부속": "레일연결부속",
                 "홀딩부속색상": "블랙", "수량": 2, "기재사항": "현장B"},
            ]}


def holding_orders(client):
    order = dw._normalize_holding_order(_holding_order(client), FIXTURE)
    order["_ship_date"] = SHIP.isoformat()
    return [order]


# ─────────────────────────────────────────────
# 롤·콤비 주문
# ─────────────────────────────────────────────
def _roll_order(client):
    return {"거래처": client, "주문번호": "부천123", "고객명": "롤고객",
            "전체기재사항": "피스/닫힌사이즈",
            "배송": {},
            "items": [
                {"롤품명": "달리", "롤색상": "그레이", "가로": 120, "세로": 180,
                 "손잡이방향": "좌", "설치장소": "거실1-1"},
                {"롤품명": "디어암막", "롤색상": "모카", "작동방식필증": "방염", "가로": 90, "세로": 140,
                 "창개수": 2, "손잡이방향": "우", "손잡이길이": 150, "기재사항": "닫힌창/안방"},
                {"롤품명": "없는제품", "롤색상": "핑크", "롤구분": "R", "가로": 60, "세로": 100},
            ]}


def roll_orders(client):
    M = get_master_for_mode("roll_combo")
    order = normalize_roll_order(_roll_order(client), M, client)
    order["_file"] = FIXTURE.name
    order["_source_path"] = str(FIXTURE)
    for it in order.get("items") or []:
        it["연창"] = False
    output.synchronize_order_for_outputs(order)
    order["_ship_date"] = SHIP.isoformat()
    return [order]


MODES = {
    "blind": (BLIND_CLIENTS, blind_orders),
    "holding": (list(HOLDING_CLIENTS), holding_orders),
    "roll_combo": (list(ROLL_CLIENTS), roll_orders),
}


def _each_case():
    for mode, (clients, make) in MODES.items():
        M = get_master_for_mode(mode)
        for client in clients:
            yield f"{mode}/{client}", mode, client, M, make


# ─────────────────────────────────────────────
# 사전점검 / 출력
# ─────────────────────────────────────────────
def _preview_payload(orders, M):
    preview, _ = dw.build_preview_rows(orders, SHIP, M, sort_di=False)
    statuses = dw.preview_statuses(preview, orders, M)
    rows = []
    for i, r in enumerate(preview):
        rows.append({
            "values": dw.row_display_values(r, i + 1),
            "editable": sorted(dw.editable_columns_for_row(r)),
            "kind": r.get("_detail_kind"), "key": r.get("_row_key"),
            "ui_only": bool(r.get("_ui_only")), "out": r.get("_output_index"),
            "status": list(statuses[i]),
        })
    return rows


def snapshot_preview():
    return _jsonable({key: _preview_payload(make(client), M)
                      for key, mode, client, M, make in _each_case()})


def _cell(c):
    v = c.value
    if isinstance(v, CellRichText):
        parts = []
        for b in v:
            if isinstance(b, str):
                parts.append([b, None, None])
            else:
                color = getattr(getattr(b.font, "color", None), "rgb", None) if b.font else None
                size = getattr(b.font, "sz", None) if b.font else None
                parts.append([b.text, color if isinstance(color, str) else None, size])
        return {"rich": parts}
    rgb = getattr(c.font.color, "rgb", None) if c.font and c.font.color else None
    if isinstance(rgb, str) and rgb not in ("FF000000",) and v not in (None, ""):
        return {"v": v, "color": rgb}
    return v


def read_book(path):
    wb = openpyxl.load_workbook(path, rich_text=True)
    sheets = {}
    for ws in wb.worksheets:
        values = [[_cell(c) for c in row] for row in ws.iter_rows()]
        while values and all(x in (None, "", "X") for x in values[-1]):
            values.pop()
        sheets[ws.title] = {"values": values,
                            "merged": sorted(str(r) for r in ws.merged_cells.ranges)}
    return sheets


def _books(builder):
    out = {}
    for key, mode, client, M, make in _each_case():
        rows = dw.build_output_rows(make(client), SHIP, M=M, sort_di=True)
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "book.xlsx"
            builder(rows, path, header_date=HEADER)
            out[key] = read_book(path)
    return _jsonable(out)


def snapshot_ledger():
    return _books(output.build_ledger)


def snapshot_worksheet():
    return _books(output.build_worksheet)


def snapshot_erp():
    import xlrd
    out = {}
    for key, mode, client, M, make in _each_case():
        orders = make(client)
        for order in orders:
            output.synchronize_order_for_outputs(order)
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "edi.xls"
            output.build_erp(orders, M, path, None)
            sheet = xlrd.open_workbook(str(path)).sheet_by_index(0)
            out[key] = [sheet.row_values(r) for r in range(sheet.nrows)]
    return _jsonable(out)


EDIT_STEPS = [
    ("product", "색상", "B 원코드 105"), ("product", "가로", "77"), ("product", "세로", "150"),
    ("product", "수량", "3"), ("product", "방향", "좌"), ("product", "길이", "140"),
    ("product", "특이", "틀안"), ("product", "기재사항1", "사용자1"), ("product", "기재사항2", ""),
    ("address", "가로", "서울특별시 수정구 수정로 9"), ("address", "특이", "화물"),
    ("recipient_contact", "가로", "홍길순 010-1234-5678"),
    ("delivery_notice", "가로", "전달 : 경비실"), ("packing", "수량", "0"),
]


def snapshot_edits():
    out = {}
    for key, mode, client, M, make in _each_case():
        orders = make(client)[:1]
        steps = []
        for kind, column, value in EDIT_STEPS:
            preview, _ = dw.build_preview_rows(orders, SHIP, M)
            target = next((r for r in preview if r.get("_detail_kind") == kind), None)
            if target is None or column not in dw.editable_columns_for_row(target):
                steps.append([kind, column, value, "skip"])
                continue
            dw.apply_preview_cell_edit(orders, target, column, value)
            preview, _ = dw.build_preview_rows(orders, SHIP, M)
            steps.append([kind, column, value,
                          [dw.row_display_values(r, i + 1) for i, r in enumerate(preview)]])
        out[key] = {"steps": steps, "items": copy.deepcopy(orders[0].get("items"))}
    return _jsonable(out)


def snapshot_read_ledger():
    out = {}
    for key, mode, client, M, make in _each_case():
        if mode == "roll_combo":
            continue
        rows = dw.build_output_rows(make(client), SHIP, M=M, sort_di=True)
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "ledger.xlsx"
            output.build_ledger(rows, path, header_date=HEADER)
            _, orders, errors = output.read_ledger(str(path), M, "ledger.xlsx")
            out[key] = {"orders": orders, "errors": errors}
    return _jsonable(out)


def snapshot_ship_labels():
    return _jsonable({
        "ship_label": [[str(d), m, dw.ship_label(d, m)] for d in (SHIP, date(2026, 9, 19))
                       for m in (None, "택배", "화물", "배달", "내사", "없음")],
        "roll_output_date": {c: [r.get("출고일") for r in dw.build_output_rows(roll_orders(c), SHIP)
                                 if r.get("출고일")] for c in ROLL_CLIENTS},
    })


# ─────────────────────────────────────────────
# 사전점검 화면 병합 시나리오 (PySide6 offscreen)
# ─────────────────────────────────────────────
UI_SCENARIOS = [("blind", "JL"), ("blind", "DI"), ("blind", "DD"), ("blind", "보노"),
                ("holding", "휴안"), ("roll_combo", "천안)채원")]


def _ui_state(ws, cols):
    table = ws.table
    texts = []
    for r in range(table.rowCount()):
        texts.append([(table.item(r, c).text() if table.item(r, c) else "") for c in cols])
    # 뒤 단계에서 목록이 제자리 수정되어도 이 단계 기록이 바뀌지 않도록 복사해 둔다.
    return copy.deepcopy({
        "spans": ws._current_spans(),
        "manual_merges": ws.manual_merges,
        "manual_unmerges": ws.manual_unmerges,
        "texts": texts,
        "items": [[it.get("기재사항"), it.get("설치장소"), it.get("_manual_note1"),
                   it.get("_manual_note2")] for o in ws.orders for it in o.get("items") or []],
    })


def snapshot_ui_merges():
    from PySide6.QtWidgets import QApplication, QMessageBox, QTableWidgetSelectionRange
    import qt_app

    app = QApplication.instance() or QApplication([])  # noqa: F841
    saved = {n: getattr(QMessageBox, n) for n in ("question", "information", "warning", "critical")}
    QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.StandardButton.Yes)
    for n in ("information", "warning", "critical"):
        setattr(QMessageBox, n, staticmethod(lambda *a, **k: QMessageBox.StandardButton.Ok))
    COL = qt_app.COL
    cols = list(range(COL["상호"], len(qt_app.HEADERS)))

    def select(ws, top, bottom, left, right):
        ws.table.setCurrentCell(top, left)
        ws.table.clearSelection()
        ws.table.setRangeSelected(QTableWidgetSelectionRange(top, left, bottom, right), True)

    out = {}
    try:
        for mode, client in UI_SCENARIOS:
            make = MODES[mode][1]
            ws = qt_app.OrderWorkspace(mode)
            ws.orders = make(client)
            ws.rebuild_table(preserve=False)
            rec = {"initial": _ui_state(ws, cols)}
            prod = [i for i, r in enumerate(ws.preview_rows) if r.get("_특수") is None]
            a, b = prod[0], prod[1]

            select(ws, a, b, COL["기재사항2"], COL["기재사항2"])
            ws.merge_selected()
            rec["merge_note2_rows01"] = _ui_state(ws, cols)

            item = ws.table.item(a, COL["기재사항2"])
            item.setText("병합후수정")
            rec["edit_merged_anchor"] = _ui_state(ws, cols)

            select(ws, a, a, COL["기재사항1"], COL["기재사항1"])
            ws.unmerge_selected()
            rec["unmerge_note1_row0"] = _ui_state(ws, cols)

            select(ws, a, a, COL["가로"], COL["가로"])
            ws.add_row_below()
            rec["add_row_below_row0"] = _ui_state(ws, cols)

            select(ws, a, b, COL["기재사항1"], COL["기재사항1"])
            ws.merge_selected()
            rows = dw.build_output_rows(ws.orders, SHIP, M=ws.M, sort_di=True)
            merges, n1 = dw.resolve_stable_merge_specs(rows, ws.manual_merges)
            unmerges, n2 = dw.resolve_stable_merge_specs(rows, ws.manual_unmerges)
            with tempfile.TemporaryDirectory() as d:
                path = Path(d) / "ledger.xlsx"
                output.build_ledger(rows, path, header_date=HEADER)
                notes = dw.apply_excel_merge_overrides(path, rows, merges, unmerges)
                book = read_book(path)
            rec["export"] = {"notes": n1 + n2 + notes,
                             "merged": {k: v["merged"] for k, v in book.items()}}
            out[f"{mode}/{client}"] = rec
            ws.deleteLater()
    finally:
        for n, fn in saved.items():
            setattr(QMessageBox, n, fn)
    return _jsonable(out)


SNAPSHOTS = {
    "preview": snapshot_preview,
    "ledger": snapshot_ledger,
    "worksheet": snapshot_worksheet,
    "erp": snapshot_erp,
    "edits": snapshot_edits,
    "read_ledger": snapshot_read_ledger,
    "ship_labels": snapshot_ship_labels,
    "ui_merges": snapshot_ui_merges,
}
