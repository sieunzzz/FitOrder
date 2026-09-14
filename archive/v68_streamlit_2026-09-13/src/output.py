"""
FitOrder 출력 생성기 — 장부 / 작업지시서 / 경영박사

    from output import build_all
    build_all(orders, ship_date="목", out_dir="../out")

orders = extract_order() 또는 parsers 결과 리스트

구조 (2026-09-13 뼈대 정리, 동작 변경 없음):
    ledger_rows.py    주문 -> 장부행 (to_rows, 업체별 기재사항)
    ledger_book.py    장부 / 작업지시서 엑셀 쓰기
    erp_edi.py        경영박사 EDI
    ledger_import.py  기존 장부 역변환 (read_ledger)
    ledger_edit.py    사전점검 편집 반영 (apply_edit)

기존 `from output import ...` 코드가 그대로 동작하도록 분리된 이름을 모두 다시 내보낸다.
"""
from pathlib import Path
from rules import Master

from ledger_rows import (
    _blind_counts, _di_item_is_urgent, _di_recipient_key, _display_size,
    _effective_delivery_mode, _handle, _holding_count, _holding_display_count,
    _needs_window_packing, _non_di_mix_ledger_text, _propagate_di_urgent_order,
    _shared_item_note_parts, _short_dir, _sort_di_group, _take_tlean, _total_window_count,
    _true_accessory_catalog, expand_same_size_directions, HANDLE_FRACTIONS, handle_split,
    JULBONG, normalized_ledger_qty, sort_di_items, synchronize_order_for_outputs, to_rows,
    unify_handle_word)  # noqa: F401
from order_text import (
    _has_non_di_mix, _jl_clean_parts, _jl_customer_name, _jl_order_mark, _jo_clean_note_text,
    _jo_note2_parts, _note_text, _prepend_note_part, _split_note_parts, _without_parts,
    JL_PACKAGING)  # noqa: F401
from ledger_book import (
    _al, _block, _border, _client_cell, _color_cell, _di_paginate, _group_len, _init_sheet,
    _new_sheet, _note_cell, _paginate_order_rows, _rich, _special_cell, _tb, _write_rows,
    ALIGN, BLUE, build_ledger, build_worksheet, COL_W, DETAIL_ALIGN, DETAIL_COL_W, DETAIL_HEAD,
    DI_MAX_ROWS, DI_SHEET_GROUPS, DI_TARGET_ROWS, FONT_DATA, FONT_UI, GREEN, HAIR, HEAD, INNER,
    LEDGER_ROWS, NO_LEFT, NO_RIGHT, RED, ROW_H, SHEET_ROWS, sort_di_items_in_ledger_order,
    THIN, to_worksheet_rows)  # noqa: F401
from erp_edi import (
    _erp_address_text, _erp_text_width, _split_erp_token, build_erp, ERP_HEAD, wrap_erp_text)  # noqa: F401
from ledger_import import (
    _ledger_client, _ledger_mix_codes, _ledger_number, parse_handle, parse_ledger_color,
    parse_ledger_mix, read_ledger)  # noqa: F401
from ledger_edit import apply_edit  # noqa: F401


# ─────────────────────────────────────────────
def build_all(orders, ship_label, master_path, out_dir="."):
    from pathlib import Path
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    M = Master(master_path)
    rows = []
    for o in orders:
        rows.extend(to_rows(o, ship_label, M))
    return {
        "장부": build_ledger(rows, f"{out_dir}/장부.xlsx"),
        "작업지시서": build_worksheet(rows, f"{out_dir}/작업지시서.xlsx"),
        "경영박사": build_erp(orders, M, f"{out_dir}/경영박사_EDI.xls"),
    }
