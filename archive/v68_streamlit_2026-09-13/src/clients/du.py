"""DU 두창블라인드 — v68 동작.

장부: 내부표시 = 괄호 포함 주문번호, 기재사항1 = 고객명/수령인, 기재사항2 = 나머지 기재내용.
수령인 자동 앞붙임 없음.
"""
from order_text import _jl_order_mark, _note_text, _split_note_parts, _without_parts

from .base import ClientRules


class DU(ClientRules):
    code = "DU"
    prepend_receiver_to_note1 = False

    def ledger_mark(self, order_no, mark):
        # JL/두창은 주문번호를 반드시 괄호 포함으로 장부에 표시한다.
        if order_no:
            mark = _jl_order_mark(order_no)
        return mark

    def ledger_notes(self, c):
        # 두창: 고객명은 공통, 발주내 기재내용은 개인창.
        note1 = _note_text(c.cust, (c.order.get("배송") or {}).get("수령인"))
        detail_parts = _without_parts(
            _split_note_parts(c.order.get("전체기재사항"), c.note, c.place),
            _split_note_parts(note1))
        note2 = _note_text(*detail_parts)
        return note1, note2
