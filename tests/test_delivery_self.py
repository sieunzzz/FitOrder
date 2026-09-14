"""사용자 확정 규칙(2026-09-13): 기본 배송지가 발주 업체 자신인 업체(기본 배송 택배/화물)가
배송 정보를 아무것도 적지 않았으면 장부·경영박사에 택배/화물을 적지 않는다. 포장비용은 적는다."""
import tempfile
from pathlib import Path

import pytest
import xlrd

import _cases
import desktop_workflow as dw
import output
from master_registry import get_master_for_mode
from rules import CLIENT_INFO
from roll_combo import ROLL_CLIENT_INFO, normalize_roll_order

BLIND = [c for c in dw.BLIND_CLIENTS if CLIENT_INFO[c][2] in ("택배", "화물")]
HOLDING = ["휴안", "구미)경남", "창문애"]
ROLL = [c for c, v in ROLL_CLIENT_INFO.items() if v["delivery"] in ("택배", "화물")]


def _order(mode, client, delivery):
    base = {"거래처": client, "주문번호": "S-1", "고객명": "현장", "배송": dict(delivery),
            "items": [{"품목코드": "102", "종류": "투코드", "타입": "C자", "가로": 100, "세로": 120}]}
    if mode == "blind":
        order = dw._normalize_ingested_order(base, Path("t.png"), "blind")
    elif mode == "holding":
        base["items"] = [{"색상원문": "H003 화이트", "품목코드": "H003", "홀딩방식": "편",
                          "가로": 100, "세로": 200}]
        order = dw._normalize_holding_order(base, Path("t.png"))
    else:
        base["items"] = [{"롤품명": "달리", "롤색상": "그레이", "가로": 100, "세로": 150}]
        order = normalize_roll_order(base, get_master_for_mode("roll_combo"), client)
    order["_ship_date"] = _cases.SHIP.isoformat()
    return order


CASES = [("blind", c) for c in BLIND] + [("holding", c) for c in HOLDING] + [("roll_combo", c) for c in ROLL]


@pytest.mark.parametrize("mode,client", CASES)
@pytest.mark.parametrize("delivery", [{}, "default"])
def test_self_delivery_writes_no_parcel_or_freight(mode, client, delivery):
    default = CLIENT_INFO[client][2] if mode != "roll_combo" else ROLL_CLIENT_INFO[client]["delivery"]
    delivery = {"방식": default} if delivery == "default" else delivery
    M = get_master_for_mode(mode)
    orders = [_order(mode, client, delivery)]
    rows = dw.build_output_rows(orders, _cases.SHIP, M=M, sort_di=True)
    ledger_text = " ".join(str(r.get(k) or "") for r in rows for k in ("출고일", "문구"))
    assert "택배" not in ledger_text and "화물" not in ledger_text
    assert not any(r.get("_특수") == "☆" for r in rows)
    with tempfile.TemporaryDirectory() as d:
        path = Path(d) / "edi.xls"
        for order in orders:
            output.synchronize_order_for_outputs(order)
        output.build_erp(orders, M, path, None)
        sheet = xlrd.open_workbook(str(path)).sheet_by_index(0)
        codes = [str(sheet.cell_value(r, 8)) for r in range(sheet.nrows)]
    assert not any(c.startswith(("#택배", "#화물", "**", "전달사항")) for c in codes)
    if mode != "roll_combo":
        assert any("포장" in c for c in codes), "기본 배송이 택배/화물이면 포장비용은 적는다"
