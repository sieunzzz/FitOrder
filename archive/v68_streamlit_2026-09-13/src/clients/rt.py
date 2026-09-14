"""RT 루임트 — v68 동작.

장부: 한 창이면 기재1에 전체/고객/기재/설치장소, 여러 창이면 첫 행에만 공통 + 설치장소는 기재2.
수량 칸은 창개수 표시, 수령인이 없어도 선불 표시행 생성.
(주소 끝 받는사람·더커튼→더·선불/착불 정리는 extract.py 후처리에 있다.)
"""
from order_text import _note_text

from .base import ClientRules


class RT(ClientRules):
    code = "RT"
    ledger_qty_uses_window_count = True
    show_prepaid_without_receiver = True

    def ledger_notes(self, c):
        # 루임트는 기존 확정 규칙을 유지한다.
        head = _note_text(c.body, c.cust)
        if c.single:
            note1 = _note_text(head, c.note, c.place)
            note2 = None
        else:
            note1 = _note_text(head if c.i == 0 else None, c.note)
            note2 = c.place or None
        return note1, note2
