"""SP 스페이스 — v68 동작.

장부: 내부표시 = 주문번호 끝 숫자(17-9 → 9), `공지`는 공통 기재사항 맨 앞.
여러 창이면 기재1=공지/주문번호/공통, 기재2=개인창.
EDI: 홀딩도어 포장비 없음.
(여러 SP 주문의 공지를 장부 첫 SP 행에 1회 모으는 처리는 app.rebuild_rows 에 있다.)
"""
import re

from order_text import _note_text, _split_note_parts, _without_parts

from .base import ClientRules


class SP(ClientRules):
    code = "SP"
    holding_pack_enabled = False

    def ledger_mark(self, order_no, mark):
        if order_no:
            m = re.search(r"(?:^|-)\s*(\d+)\s*$", order_no)
            mark = m.group(1) if m else ""
        return mark

    def ledger_common(self, common, order_no):
        if order_no:
            common = "/".join(x for x in common.split("/")
                              if x.strip() and x.strip() != order_no)
        common_parts = [x.strip() for x in common.split("/") if x.strip()]
        sp_notices = [x for x in common_parts if "공지" in x]
        common = "/".join(x for x in common_parts if "공지" not in x)
        return common, sp_notices

    def ledger_notes(self, c):
        # SP도 여러 창일 때 공통/개인창 규칙을 따른다. 공지·주문번호는 공통.
        sp_common = _split_note_parts(*c.sp_notices, c.order_no, *c.common_note_parts)
        personal = _without_parts(_split_note_parts(c.note, c.place), sp_common)
        if c.multi_windows:
            note1 = _note_text(*sp_common)
            note2 = _note_text(*personal)
        else:
            note1 = _note_text(*sp_common, *personal)
            note2 = None
        return note1, note2
