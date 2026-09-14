"""JL — v68 동작.

장부: 내부표시 = 괄호 포함 발주번호, 고객명 끝 DW 제거, 기재사항1 = 피스/이름/기타, 기재사항2 = 설치장소.
수량 칸은 창개수 표시, 수령인 자동 앞붙임 없음.
EDI: 손잡이길이 → 피스 → 이름 → 나머지 기재사항.
(같은 고객 기재사항 세로 병합은 ledger_book._write_rows, 판독 보정은 extract.py 에 있다.)
"""
from order_text import (_jl_clean_parts, _jl_customer_name, _jl_order_mark, _note_text,
                        _split_note_parts)
from rules import normalize_handle

from .base import ClientRules


class JL(ClientRules):
    code = "JL"
    prepend_receiver_to_note1 = False
    ledger_qty_uses_window_count = True

    def ledger_mark(self, order_no, mark):
        # JL/두창은 주문번호를 반드시 괄호 포함으로 장부에 표시한다.
        if order_no:
            mark = _jl_order_mark(order_no)
        return mark

    def ledger_customer(self, cust):
        return _jl_customer_name(cust)

    def ledger_notes(self, c):
        # JL: 피스/이름은 공통 기재사항1, 설치장소는 개인창 기재사항2.
        raw_parts = _split_note_parts(c.common, c.note)
        has_piece = any("피스" in x for x in raw_parts)
        fixed = (["피스"] if has_piece else []) + ([c.cust] if c.cust else [])
        item_parts = _jl_clean_parts(_split_note_parts(c.note), c.cust, c.order_no)
        common_parts = _jl_clean_parts(_split_note_parts(c.common), c.cust, c.order_no)
        note1 = _note_text(*(fixed + item_parts + common_parts))
        note2 = c.place or None
        return note1, note2

    def erp_blind_memo_parts(self, c):
        it = c.it
        extra_parts = []
        common = str(c.o.get("전체기재사항") or "")
        receiver = _jl_customer_name(c.o.get("고객명") or c.delivery.get("수령인") or "")
        all_parts = _split_note_parts(it.get("_수동특이"), common,
                                     it.get("기재사항"), it.get("설치장소"))
        # JL EDI 적요 순서 고정: 손잡이길이 -> 피스 -> 이름 -> 나머지 기재사항.
        handle = normalize_handle(it.get("손잡이길이"))[0]
        if handle:
            extra_parts.append(f"손{handle}")
        if any("피스" in x for x in all_parts):
            extra_parts.append("피스")
        if receiver:
            extra_parts.append(receiver)
        # 나머지 실제 기재사항은 항상 그 뒤에 붙인다.
        for part in _jl_clean_parts(all_parts, receiver, c.o.get("주문번호")):
            if part not in extra_parts:
                extra_parts.append(part)
        return extra_parts
