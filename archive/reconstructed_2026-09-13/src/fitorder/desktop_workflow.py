from __future__ import annotations
import copy
from datetime import date
from pathlib import Path
from .overrides import apply_cell_edit
from .pipeline import LEDGER_COLUMNS, build_rows
from .rules import CLIENTS, Master, normalize_client
from .parsers import parse_json_order, parse_xlsx_generic

__all__ = [
    "CLIENTS", "LEDGER_COLUMNS", "Master", "apply_preview_cell_edit", "build_preview_rows",
    "client_label", "default_ship", "editable_columns_for_row", "generate_outputs",
    "ingest_file", "preview_statuses", "row_display_values",
]

def default_ship():
    return date.today()

def client_label(client):
    return client

def ingest_file(path, master: Master | None = None, forced_client: str | None = None):
    p = Path(path)
    suf = p.suffix.lower()
    if suf == ".json":
        orders = parse_json_order(p)
    elif suf in {".xlsx",".xls"}:
        if suf == ".xls":
            raise ValueError(".xls는 업체별 legacy 변환기 연결이 필요합니다. Claude 작업 시 xlrd→xlsx 임시변환을 추가하세요.")
        orders = parse_xlsx_generic(p, forced_client)
    elif suf in {".png",".jpg",".jpeg",".webp",".bmp"}:
        from .ai_extract import extract_with_openai
        orders = extract_with_openai(p, forced_client)
    else:
        raise ValueError(f"지원하지 않는 입력 형식: {suf}")
    for o in orders:
        o["거래처"] = normalize_client(o.get("거래처")) or forced_client or "미확인"
        o.setdefault("배송", {})
        o.setdefault("items", [])
        o.setdefault("_source_path", str(p.resolve()))
    return orders

def build_preview_rows(orders, ship_date=None, master=None, sort_di=False):
    rows = build_rows(orders, ship_date)
    return rows, copy.deepcopy(rows)

def row_display_values(row, no=None):
    return {k: row.get(k,"") for k in LEDGER_COLUMNS}

def editable_columns_for_row(row):
    if row.get("_특수") is None:
        return set(LEDGER_COLUMNS) - {"X"}
    kind = row.get("_detail_kind")
    return {
        "address":{"가로","기재사항1"},
        "recipient_contact":{"가로"},
        "delivery_notice":{"가로"},
        "packing":{"수량"},
    }.get(kind, {"가로"})

def apply_preview_cell_edit(orders, row, column, value):
    apply_cell_edit(orders, row, column, value)

def preview_statuses(rows, orders, master=None):
    out = []
    for r in rows:
        edited = f" · 사용자 수정: {', '.join(r['_overridden'])}" if r.get("_overridden") else ""
        if r.get("_특수") is not None:
            if r.get("_detail_kind") == "address" and not str(r.get("가로") or "").strip():
                out.append(("yellow","확인","택배/화물 주소가 비어 있습니다."+edited))
            else:
                out.append(("info","정보","부가정보/특수행"+edited))
            continue
        problems = []
        if not str(r.get("색상") or "").strip(): problems.append("품목/색상 누락")
        if not str(r.get("가로") or "").strip(): problems.append("가로 누락")
        if not str(r.get("세로") or "").strip(): problems.append("세로 누락")
        if problems: out.append(("red","오류"," / ".join(problems)+edited))
        else: out.append(("ok","정상","규칙상 즉시 확인되는 오류 없음"+edited))
    return out

def resolve_stable_merge_specs(*args, **kwargs):
    return []

def apply_excel_merge_overrides(*args, **kwargs):
    return None

def generate_outputs(orders, ship_date, out_dir, master=None):
    from .output import build_all
    return build_all(orders, ship_date=ship_date, out_dir=out_dir)
