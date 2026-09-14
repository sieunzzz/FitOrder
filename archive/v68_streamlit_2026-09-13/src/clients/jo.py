"""JO 제이원 — v68 동작.

장부: 기재사항1 첫 값은 주문번호. 1~2창은 나머지도 기재1, 3창 이상부터 기재2 사용.
수량 칸은 창개수 표시, 수령인 자동 앞붙임 없음.
EDI: 주문번호 → 특이 → 나머지 기재사항.
"""
from order_text import _jo_note2_parts, _note_text

from .base import ClientRules


class JO(ClientRules):
    code = "JO"
    prepend_receiver_to_note1 = False
    ledger_qty_uses_window_count = True

    def ledger_notes(self, c):
        # JO: 주문번호는 항상 기재사항1의 첫 값.
        # 1~2창은 나머지 내용도 기재사항1에 합치고, 3창 이상부터 기재사항2를 쓴다.
        rest = _jo_note2_parts(c.order, c.it, c.cust, first=True)
        if c.total_windows >= 3:
            note1 = _note_text(c.order_no)
            note2 = _note_text(*rest)
        else:
            note1 = _note_text(c.order_no, *rest)
            note2 = None
        return note1, note2

    def erp_blind_memo_parts(self, c):
        it = c.it
        extra_parts = []
        # 제이원 EDI에도 주문번호가 반드시 보이고, 나머지 기재사항은 그 뒤에 둔다.
        order_no = str(c.o.get("주문번호") or "").strip()
        if order_no:
            extra_parts.append(order_no)
        manual_special = str(it.get("_수동특이") or "").strip()
        if manual_special:
            extra_parts.append(manual_special)
        for part in _jo_note2_parts(c.o, it, c.o.get("고객명") or c.delivery.get("수령인"), first=True):
            if part not in extra_parts:
                extra_parts.append(part)
        return extra_parts
