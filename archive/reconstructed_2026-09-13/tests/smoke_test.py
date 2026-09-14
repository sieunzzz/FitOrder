from pathlib import Path
import json
from fitorder.desktop_workflow import build_preview_rows
from fitorder.rules import note_fields, dimension_flags
from fitorder.roll_combo import billed_area

def test_preview():
    p = Path(__file__).parents[1] / "data" / "samples" / "sample_order.json"
    orders = json.loads(p.read_text(encoding="utf-8"))
    rows,_ = build_preview_rows(orders, "월")
    assert len(rows) >= 2
    assert rows[0]["상호"] == "JL"

def test_jo_note2_only_three_windows():
    o={"거래처":"JO","주문번호":"123","전체기재사항":"공통","items":[{},{},{}]}
    n1,n2=note_fields(o,{"설치장소":"거실"},0,3)
    assert n1=="123"
    assert n2=="거실"

def test_roll_combo_minimum():
    assert billed_area(100,100,1) == 2.0

def test_dimension_flags():
    assert dimension_flags("blind","원코드",35,180)["width_blue"] is True
    assert dimension_flags("holding",None,100,300)["height_blue"] is True
