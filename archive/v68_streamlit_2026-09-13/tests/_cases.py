"""v68 원본 코드의 현재 동작을 고정하는 특성(characterization) 테스트 케이스.

- golden 파일은 "정답"이 아니라 "현재 동작"이다. 구조 정리 중 출력이 바뀌면 실패한다.
- 규칙을 의도적으로 바꾼 경우에만 `python tests/_generate_golden.py <항목>`으로 갱신하고
  바뀐 부분을 사용자에게 보고한다.
- 가상의 값만 사용한다(실제 고객 정보 금지).
- 실제 품목 마스터(fitorder_master.xlsx)가 없으므로 StubMaster(고정 가짜 단가)로
  경영박사 EDI·validate·DI 정렬의 "로직"만 고정한다. 실제 단가/품명 검증은 마스터 확보 후 추가.
- AI 추출은 FakeOpenAI로 고정 응답을 넣어 업체별 후처리 로직만 고정한다.
"""
from __future__ import annotations

import copy
import json
import sys
import tempfile
from datetime import date
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))

import openpyxl  # noqa: E402
from openpyxl.cell.rich_text import CellRichText  # noqa: E402

import holding  # noqa: E402
import output  # noqa: E402
import rules  # noqa: E402

GOLDEN_PATH = Path(__file__).resolve().parent / "golden" / "characterization.json"
HEADER_DATE = date(2026, 9, 14)
CLIENTS = list(rules.CLIENT_INFO) + ["미확인"]


# ─────────────────────────────────────────────
# 가짜 마스터
# ─────────────────────────────────────────────
class StubMaster:
    """rules.Master 와 같은 조회 메서드를 제공하는 결정적 가짜 마스터."""
    THREAD = {"102": "백색", "105": "백색", "200": "아이보리", "290": "베이지", "330": "핑크"}

    @staticmethod
    def _price(name):
        return 10000 + sum(ord(ch) for ch in str(name)) % 9000

    def thread_color(self, code):
        return self.THREAD.get(str(code or "").split("+")[0])

    def find_order_item(self, item, client):
        if item.get("_product_group") == "holding" or item.get("_holding_accessory"):
            name = item.get("_holding_product_name") or item.get("품목코드")
        elif item.get("_mix_name") and item.get("_mix_codes"):
            name = f"B{item['_mix_name']}({item['_mix_codes']})-{item.get('종류')}/{item.get('타입')}"
        elif item.get("품목코드"):
            name = f"B{item['품목코드']}-{item.get('종류')}/{item.get('타입')}"
        else:
            return None
        price = 0 if str(item.get("품목코드")) == "330" else self._price(name)
        return {"품명": name, "관리코드": name, "규격": "25", "단가": price, "단가등급": "출고I가"}

    def find_holding_extra(self, name, client):
        return {"품명": name, "관리코드": name, "규격": "", "단가": 4000,
                "_inferred_tiers": {"출고I가"}}

    def find_named_item(self, name, client):
        if "2EA" in str(name):
            return None
        return {"품명": name, "관리코드": name, "규격": "", "단가": 3000, "단가등급": "출고A가"}


STUB = StubMaster()


# ─────────────────────────────────────────────
# 주문 형태
# ─────────────────────────────────────────────
def _full(client):
    """택배·선불·주소·전달사항, 원코드 연창·FP·L자 셔터·분수 수량, 업체별 부속 필드."""
    return {
        "거래처": client, "주문번호": "17-9", "고객명": "홍길동 DW",
        "전체기재사항": "피스/공학1관/공지 금요일 휴무",
        "전체원문": "총 5창", "변경요청": True, "변경문구": "사이즈 이걸로 변경되나요",
        "배송": {"방식": "택배", "선불착불": "선불",
                 "주소": "서울특별시 테스트구 테스트로 12, (테스트동, 가나빌) 101동 202호",
                 "수령인": "김수령", "연락처": "010-0000-0000", "전달사항": "배송전연락/문앞"},
        "_true_accessories": [{"표시": "노피스(2)", "품명": "노피스브라켓(2EA)", "수량": 1},
                              {"표시": "B 원코드 브라켓", "품명": "B브라켓(25mm원코드/구)", "수량": 1}],
        "_unit_accessories": [{"품명": "노피스브라켓(3EA)", "세트": 2, "피스수": 3}],
        "_추가부속": "석고앙카2",
        "_az_place": "금빛커텐",
        "_bono_note1": "받는상호",
        "items": [
            {"품목코드": "102", "종류": "원코드", "타입": "C자", "가로": 150.0, "세로": 124,
             "창개수": 1, "손잡이방향": "좌", "손잡이길이": 120, "연창": True,
             "설치장소": "거실", "기재사항": "공학1관",
             "확신도": {"가로": 0.5, "세로": 1.0, "품목코드": 1.0, "손잡이": 0.2}},
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


def _minimal(client):
    return {"거래처": client, "고객명": "이고객", "배송": {},
            "items": [{"품목코드": "330", "종류": "투코드", "타입": "C자",
                       "가로": "100", "세로": "150.5", "손잡이방향": "우"}]}


def _cargo(client):
    """화물·착불·주소 있음, 같은 색/같은 세로 반복(ditto), 좌우 문자열 펼침."""
    return {"거래처": client, "주문번호": "C-7", "고객명": "박현장",
            "배송": {"방식": "화물", "선불착불": "착불", "주소": "오산삼미",
                     "수령인": "보노 안산", "연락처": "010-1111-2222", "발신": "테스트 010-2222-3333"},
            "_bono_note1": "안산초지", "_bono_receiver_is_bono": True,
            "items": [
                {"품목코드": "200", "종류": "투코드", "타입": "C자", "가로": 100, "세로": 150,
                 "손잡이방향": "좌우", "설치장소": "작은방"},
                {"품목코드": "200", "종류": "투코드", "타입": "C자", "가로": 110, "세로": 150,
                 "창개수": 3, "손잡이방향": "우", "설치장소": "작은방"},
                {"품목코드": None, "종류": "투코드", "타입": "C자", "가로": None, "세로": 130,
                 "예외품목": "수리"},
            ]}


def _mix_holding(client):
    """DI MIX(정식·일반), 기재사항 수령인 2명, 홀딩 본품/부속(내부 표식 보존 경로)."""
    return {"거래처": client, "배송": {"방식": None},
            "items": [
                {"품목코드": "200+290", "_mix_name": "밀크코코아", "_mix_codes": "200+290",
                 "종류": "원코드", "타입": "C자", "가로": 120, "세로": 200, "기재사항": "최수령"},
                {"품목코드": "102+520", "_mix_name": "MIX", "_mix_codes": "102+520",
                 "_generic_mix": True, "종류": "투코드", "타입": "C자", "가로": 60, "세로": 200,
                 "기재사항": "최수령", "_수동특이": "긴급", "손잡이길이": 125},
                {"품목코드": "102", "종류": "투코드", "타입": "C자", "가로": 70, "세로": 100,
                 "기재사항": "정수령"},
                {"품목코드": "H003", "_product_group": "holding", "_ledger_prefix": "H",
                 "_holding_product_name": "H003 (화이트)", "_holding_label": "화이트 I",
                 "_holding_operation": "양자석", "_holding_upper_roller": True,
                 "가로": 88, "세로": 310, "창개수": 1, "기재사항": "정수령"},
                {"품목코드": "H레일연결부속", "_product_group": "holding", "_holding_accessory": True,
                 "_holding_product_name": "H레일연결부속(화이트)",
                 "_holding_accessory_label": "레일연결부속", "_holding_accessory_color": "화이트",
                 "수량": 2, "기재사항": "정수령"},
                {"품목코드": "H005", "_product_group": "holding", "_ledger_prefix": "H",
                 "_holding_product_name": "H005 (아이보리/9)", "_holding_operation": "1/3",
                 "가로": 250, "세로": 120, "창개수": 2, "기재사항": "문수령", "설치장소": "베란다"},
            ]}


VARIANTS = {"full": _full, "minimal": _minimal, "cargo": _cargo, "mix_holding": _mix_holding}


def _jsonable(value):
    return json.loads(json.dumps(value, ensure_ascii=False, default=str))


def _rows_for(client, M=None, sort_di=False, variants=None):
    """app.rebuild_rows 와 같은 방식으로 주문별 장부행을 만들고 _oi 를 붙인다."""
    rows = []
    for oi, make in enumerate((variants or VARIANTS).values()):
        for r in output.to_rows(make(client), "월", M, sort_di=sort_di):
            r["_oi"] = oi
            rows.append(r)
    return rows


# ─────────────────────────────────────────────
# 장부행 / 엑셀
# ─────────────────────────────────────────────
def snapshot_to_rows():
    out = {}
    for client in CLIENTS:
        for name, make in VARIANTS.items():
            order = make(client)
            rows = output.to_rows(order, "월", None, sort_di=False)
            out[f"{client}/{name}"] = {"rows": rows, "order_after": order}
    return _jsonable(out)


def snapshot_to_rows_di_sorted():
    out = {}
    for name, make in VARIANTS.items():
        out[name] = output.to_rows(make("DI"), "월", STUB, sort_di=True)
    return _jsonable(out)


def _cell(c):
    v = c.value
    if isinstance(v, CellRichText):
        parts = []
        for b in v:
            if isinstance(b, str):
                parts.append([b, None])
            else:
                color = getattr(getattr(b.font, "color", None), "rgb", None) if b.font else None
                size = getattr(b.font, "sz", None) if b.font else None
                parts.append([b.text, color if isinstance(color, str) else None, size])
        return {"rich": parts}
    rgb = getattr(c.font.color, "rgb", None) if c.font and c.font.color else None
    if isinstance(rgb, str) and rgb not in ("FF000000",) and v not in (None, ""):
        return {"v": v, "color": rgb}
    return v


def _read_book(path):
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
    for client in CLIENTS:
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "book.xlsx"
            builder(_rows_for(client), path, HEADER_DATE)
            out[client] = _read_book(path)
    return _jsonable(out)


def snapshot_ledger():
    return _books(output.build_ledger)


def snapshot_worksheet():
    return _books(output.build_worksheet)


DI_ONLY = {"minimal": _minimal, "mix_holding": _mix_holding}


def _di_books(builder):
    with tempfile.TemporaryDirectory() as d:
        path = Path(d) / "book.xlsx"
        # DI 전용 시트 분할은 모든 행이 DI일 때만 동작한다(특수행 없는 주문).
        rows = []
        for rep in range(12):  # 페이지 분할(35~40행)을 타도록 반복
            for r in _rows_for("DI", STUB, sort_di=True, variants=DI_ONLY):
                r = dict(r)
                r["_oi"] = r["_oi"] + rep * 10
                rows.append(r)
        builder(rows, path, HEADER_DATE)
        return _jsonable(_read_book(path))


def snapshot_ledger_di_sorted():
    return _di_books(output.build_ledger)


def snapshot_worksheet_di_sorted():
    return _di_books(output.build_worksheet)


# ─────────────────────────────────────────────
# 편집 / 역변환
# ─────────────────────────────────────────────
EDITS = [
    ("상호", "JL"), ("색상", "B 원코드 105"), ("가로", "77"), ("세로", '"'),
    ("수량", "3"), ("방향", "ㅈ"), ("길이", "#손140"), ("특이", "틀안 # 긴급"),
    ("기재사항", "사용자기재1"), ("기재사항2", ""), ("출고일", "2026.09.15"),
    ("색상", "B MIX"), ("기재사항", "102+520/최수령"),
]


def snapshot_apply_edit():
    out = {}
    for client in CLIENTS:
        order = _full(client)
        steps = []
        for field, value in EDITS:
            output.apply_edit(order, 0, field, value)
            output.synchronize_order_for_outputs(order)
            steps.append([field, value, copy.deepcopy(order["items"][0]), order.get("거래처"),
                          order.get("_ship_date")])
        rows = output.to_rows(order, "월", None, sort_di=False)
        out[client] = {"steps": steps, "rows": rows}
    return _jsonable(out)


def snapshot_read_ledger():
    out = {}
    for client in CLIENTS:
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "ledger.xlsx"
            output.build_ledger(_rows_for(client), path, HEADER_DATE)
            rows, orders, errors = output.read_ledger(str(path), STUB, "ledger.xlsx")
            out[client] = {"rows": rows, "orders": orders, "errors": errors}
    return _jsonable(out)


# ─────────────────────────────────────────────
# 경영박사 EDI / 검증 (StubMaster)
# ─────────────────────────────────────────────
def _read_xls(path):
    import xlrd
    sheet = xlrd.open_workbook(str(path)).sheet_by_index(0)
    return [sheet.row_values(r) for r in range(sheet.nrows)]


def snapshot_erp():
    out = {}
    for client in CLIENTS:
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "edi.xls"
            orders = [make(client) for make in VARIANTS.values()]
            output.build_erp(orders, STUB, path, HEADER_DATE)
            out[client] = _read_xls(path)
    return _jsonable(out)


def snapshot_validate():
    # 주의(v68 버그 기록): validate 는 가로/세로가 문자열("100")이면 TypeError 로 멈춘다.
    # 원본 동작을 바꾸지 않는 단계이므로 여기서는 숫자 치수로만 고정한다.
    out = {}
    for client in CLIENTS:
        for name, make in VARIANTS.items():
            order = make(client)
            for it in order["items"]:
                for field in ("가로", "세로"):
                    if isinstance(it.get(field), str):
                        it[field] = float(it[field])
            out[f"{client}/{name}"] = rules.validate(order, STUB)
    return _jsonable(out)


# ─────────────────────────────────────────────
# AI 추출 후처리 (FakeOpenAI)
# ─────────────────────────────────────────────
def _fake_extraction(client):
    item = {
        "제품군": "블라인드", "홀딩방식": None, "홀딩레일": None, "홀딩상하로라": False,
        "홀딩부속": None, "홀딩부속색상": None, "품목코드": "102", "색상원문": "알루미늄 화이트 MIX",
        "타입": "L자", "종류": "투코드", "가로": 100, "세로": 150, "수량": "½",
        "손잡이방향": "ㅈ", "손잡이길이": None, "연창": False, "설치장소": "작은방",
        "기재사항": "더커튼/동탄푸르1502/택배비 선불/틀안/피스/비닐포장/내일 출고 부탁드립니다/줄120",
        "예외품목": None, "창개수": None, "좌개수": 0, "우개수": 2,
        "원문": "원코드 좌,조절110 손130 이타입 H",
        "확신도": {"가로": 1, "세로": 1, "품목코드": 1, "손잡이": 1},
    }
    second = dict(item, 품목코드="200", 기재사항="스냅 2개/노피스(2)", 원문="L18 투코드", 수량=None)
    return {
        "거래처": "DI", "발주일": None, "주문번호": "(1234)", "고객명": "임지애 DW",
        "items": [item, second],
        "배송": {"방식": None, "화물지점": None, "주소": "인천 테스트로 1", "수령인": "더커튼",
                 "연락처": None, "선불착불": None, "전달사항": "착불/문앞"},
        "전체원문": ("주문 번 호 : 226095 (5678)\n원코드 피스\n창틀용 무타공 커튼박스용 무타공 택배비선불\n"
                     "안방베란다82,154,82×220 IV200\nL타입좌,우,우원코드\n에어캡+포장"),
        "전체기재사항": "226095/피스/임지애 DW/비닐포장/줄120/☆시온가공소/금요일 출고 부탁드립니다",
        "변경요청": False, "변경문구": None,
    }


class _FakeOpenAI:
    def __init__(self, payload):
        self.payload = payload
        self.chat = SimpleNamespace(completions=SimpleNamespace(create=self._create))

    def _create(self, **kwargs):
        return SimpleNamespace(
            choices=[SimpleNamespace(message=SimpleNamespace(content=json.dumps(self.payload, ensure_ascii=False)))],
            model="fake-model", usage=SimpleNamespace(prompt_tokens=1, completion_tokens=2))


def snapshot_extract_order():
    import extract
    from PIL import Image
    out = {}
    original = extract.client
    try:
        with tempfile.TemporaryDirectory() as d:
            img = Path(d) / "order.png"
            Image.new("RGB", (60, 40), "white").save(img)
            for client in CLIENTS:
                extract.client = _FakeOpenAI(_fake_extraction(client))
                result = extract.extract_order([img], client_hint=None if client == "미확인" else client)
                result["_meta"]["files"] = ["order.png"]
                out[client] = result
    finally:
        extract.client = original
    return _jsonable(out)


def snapshot_extract_helpers():
    import extract
    pages = [{"주문번호": "<17-9>", "items": [{"n": 1}], "전체기재사항": "공지/A", "_source_pages": [1]},
             {"주문번호": "17 - 9", "items": [{"n": 2}], "전체기재사항": "A/B", "_source_pages": [2]},
             {"주문번호": None, "items": [{"n": 3}], "_source_pages": [3]}]
    holding_items = [
        {"제품군": "홀딩도어", "품목코드": "H003", "_product_group": "holding", "예외품목": "홀딩",
         "_holding_label": "x"},
        {"품목코드": "102", "색상원문": "H 화이트 I"},
    ]
    return _jsonable({
        "sp_key": [[v, extract._sp_order_key(v)] for v in ("<17-9>", "17 – 09", "abc", None)],
        "merge_sp": extract._merge_sp_orders(copy.deepcopy(pages)),
        "clean_jo_note": [[v, extract._clean_jo_note(v)] for v in ("손120/거실", "줄 90, 안방", None)],
        "clean_rt_note": [[v, extract._clean_rt_note_text(v)] for v in ("택배비 선불/거실", "착불", None)],
        "clean_true_note": [[v, extract._clean_true_accessory_note(v)]
                            for v in ("피스추가+택배비선불+빠른출고", "창틀용 무타공 2세트/노피스(2)/피스/피스")],
        "normalize_holding_off": [extract._normalize_holding_item(copy.deepcopy(x)) for x in holding_items],
        "default_handle": [[k, h, extract.default_handle_length(k, h)]
                           for k in ("원코드", "셔터") for h in (None, 95, 210)],
        "color_text": [extract.color_text("투코드", "102"), extract.color_text("원코드", "200")],
        "ledger_color": [extract.ledger_color(x) for x in (
            {"품목코드": "102"}, {"품목코드": "200", "종류": "원코드", "타입": "L자"},
            {"_mix_name": "밀크코코아", "_mix_codes": "200+290", "종류": "원코드"})],
        "gate": [holding.apply_holding_feature_gate(copy.deepcopy({"items": holding_items}))],
    })


# ─────────────────────────────────────────────
# 순수 함수
# ─────────────────────────────────────────────
def snapshot_pure_functions():
    sizes = [(100, 100), (80, 150), (29.5, 371), (250, 90), (88, 310), (300, 270)]
    ops = ["", "편", "편개", "양", "양개", "양자석", "양개양자석", "고정", "1/2", "1/3양자석",
           "1/10", "½", "(편+양자석)", "이동분1/4"]
    texts = ["H003", "H 화이트 I", "화이트II", "H5", "투명+블랙부속★", "방염 F2", "레일연결부속 블랙",
             "라운드부속", "홀딩도어 월넛/590", "102", "B 원코드 029FP"]
    addresses = ["경상남도 진주시 초전북로 151 (초전동) 인사인광고기획",
                 "서울특별시 영등포구 도신로 161 (도림동, 정성빌) 201호",
                 "대구시 테스트로 25, 101동 202호 (가나아파트 101동, 경비실)"]
    return _jsonable({
        "calc_erp": [[w, h, rules.calc_erp(w, h, 12000)] for w, h in sizes],
        "holding_calc": [[w, h, holding.holding_calc(w, h, 16000)] for w, h in sizes],
        "operations": [[op, holding.normalize_holding_operation(op),
                        holding.looks_like_holding_operation(op),
                        holding.holding_magnet_bars(op), holding.holding_magnet_charge_qty(op)]
                       for op in ops],
        "holding_text": [[t, holding.holding_product_from_text(t),
                          holding.holding_accessory_from_text(t)] for t in texts],
        "ledger_color": [[args, rules.ledger_color(*args)] for args in [
            ("102", "투코드", "C자"), ("029FP", "원코드", "C자"), ("200", "셔터", "L자"),
            ("200+290", "원코드", "C자", "B", "밀크코코아", "200+290"),
            ("102+520", "원코드", "C자", "B", "MIX", "102+520"), ("105", "투코드", "C자", "X")]],
        "default_handle": [[k, h, rules.default_handle_length(k, h)]
                           for k in ("원코드", "투코드", "셔터", "기타") for h in (None, 90, 150, 210)],
        "normalize_handle": [[v, rules.normalize_handle(v)] for v in (None, 0, 120, 125)],
        "clean_notice": [[v, rules.clean_delivery_notice(v)] for v in
                         (None, "배송전연락", "배송전 연락/문앞\n던지지마세요", "경비실/경비실")],
        "di_mix": [[v, rules.di_mix_info(v), rules.di_mix_from_codes(v)] for v in
                   ("포인트믹스 밀크코코아", "밀크코코아", "102+330", "330+102", "290+200")],
        "erp_address": [[a, output._erp_address_text(a), output.wrap_erp_text(output._erp_address_text(a))]
                        for a in addresses],
        "handle_split": [[v, output.handle_split(v)] for v in ("1/2", "½", "1/1", "2", None)],
        "parse_ledger": [[t, output.parse_ledger_color(t), output.parse_ledger_mix(t, "102+520/홍"),
                          output.parse_handle(t)] for t in
                         (" B 원코드 029FP", "B L18-투코드 200", "B MIX", "B 밀크코코아 (200+290)", "#손140")],
        "expand_directions": [output.expand_same_size_directions(copy.deepcopy(o)) for o in (
            {"items": [{"좌개수": 1, "우개수": 2, "창개수": 3}]},
            {"items": [{"손잡이방향": "좌우우"}]},
            {"items": [{"손잡이방향": "우", "창개수": 2}]},
            {"items": [{"_product_group": "holding", "손잡이방향": "좌우"}]})],
        "duplicates": rules.find_duplicates(
            [_minimal("DI")], [("09-10 출고", "DI", _minimal("DI")["items"][0])]),
        "changes": rules.find_changes(
            [_minimal("DI")], [("09-10 출고", "DI", dict(_minimal("DI")["items"][0], 가로="100.5", 세로=160), True)]),
    })


SNAPSHOTS = {
    "to_rows": snapshot_to_rows,
    "to_rows_di_sorted": snapshot_to_rows_di_sorted,
    "ledger": snapshot_ledger,
    "worksheet": snapshot_worksheet,
    "ledger_di_sorted": snapshot_ledger_di_sorted,
    "worksheet_di_sorted": snapshot_worksheet_di_sorted,
    "apply_edit": snapshot_apply_edit,
    "read_ledger": snapshot_read_ledger,
    "erp": snapshot_erp,
    "validate": snapshot_validate,
    "extract_order": snapshot_extract_order,
    "extract_helpers": snapshot_extract_helpers,
    "pure_functions": snapshot_pure_functions,
}
