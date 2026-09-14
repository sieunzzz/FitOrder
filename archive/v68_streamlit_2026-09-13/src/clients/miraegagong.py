"""미래가공 — v68 동작 (보노와 다른 전용 배치).

장부: 기재사항1 = 기재/설치장소, 기재사항2 = 첫 행에만 주문번호. 손잡이 150 은 길이 칸에 표시 안 함.
"""
from order_text import _note_text

from .base import ClientRules


class Miraegagong(ClientRules):
    code = "미래가공"
    hide_handle_150 = True

    def ledger_notes(self, c):
        # 미래가공은 기존 업체 전용 배치를 유지한다.
        note1 = _note_text(c.note, c.place)
        note2 = c.order_no if c.i == 0 else None
        return note1, note2
