from __future__ import annotations
import json
from pathlib import Path
from typing import Any

def _num(v):
    if v in (None,""): return None
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except Exception:
        return v

def parse_json_order(path: str | Path):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    return data if isinstance(data, list) else [data]

def parse_xlsx_generic(path: str | Path, forced_client: str | None = None):
    """일반적인 표형 발주서 fallback.
    업체별 parser가 없을 때만 사용한다.
    """
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = {}
    for ri, row in enumerate(rows):
        norm = [str(x or "").strip().lower() for x in row]
        if any(x in norm for x in ("가로","가로(cm)","width")) and any(x in norm for x in ("세로","세로(cm)","height")):
            for ci, val in enumerate(norm):
                if val: headers[val] = ci
            start = ri + 1
            break
    else:
        raise ValueError("표 헤더(가로/세로)를 찾지 못했습니다.")

    order = {
        "거래처": forced_client or "미확인",
        "주문번호": None,
        "고객명": None,
        "전체기재사항": None,
        "배송": {},
        "items": [],
        "_source_path": str(Path(path).resolve()),
    }
    def col(row, *names):
        for n in names:
            if n.lower() in headers:
                i = headers[n.lower()]
                return row[i] if i < len(row) else None
        return None
    for row in rows[start:]:
        w = col(row,"가로","가로(cm)","width")
        h = col(row,"세로","세로(cm)","height")
        if w in (None,"") and h in (None,""):
            continue
        order["items"].append({
            "품목코드": col(row,"품목","품목코드","product"),
            "종류": col(row,"종류","타입","type") or "원코드",
            "가로": _num(w), "세로": _num(h),
            "창개수": int(_num(col(row,"수량","창개수","qty")) or 1),
            "손잡이방향": col(row,"방향","손잡이방향"),
            "손잡이길이": _num(col(row,"손잡이","손잡이길이","길이")),
            "기재사항": col(row,"기재사항","설치장소","메모"),
        })
    return [order]
