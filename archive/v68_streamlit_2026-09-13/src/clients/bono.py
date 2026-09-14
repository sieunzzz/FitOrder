"""보노(이끌림) — v68 동작.

장부: 기재사항1 = 받는사람/화물지점(_bono_note1), 기재사항2 = 오더명(주문번호)/시공위치.
수령인 자동 앞붙임 없음. EDI: 손잡이 → 기재1 → 오더명/시공위치, 포장비 품목 `포장비용(B/R/C/S/HC)`.
(받는사람=보노일 때 EDI 연락처 행 처리는 erp_edi.build_erp 에 있다.)
"""
from order_text import _note_text
from rules import BONO_PACKING_ITEM, normalize_handle

from .base import ClientRules


class Bono(ClientRules):
    code = "보노"
    prepend_receiver_to_note1 = False
    blind_packing_item = BONO_PACKING_ITEM

    def ledger_notes(self, c):
        # 보노 기존 규칙: 기재사항1=받는사람/화물지점, 기재사항2=오더명/시공위치.
        note1 = _note_text(c.order.get("_bono_note1"))
        note2 = _note_text(c.order_no, c.place)
        return note1, note2

    def erp_blind_memo_parts(self, c):
        it = c.it
        extra_parts = []
        # 보노 EDI도 장부와 같은 기재사항 구조를 사용한다.
        # 손잡이길이는 제작정보로 앞에, 기재사항1/2는 그 뒤에 둔다.
        handle = normalize_handle(it.get("손잡이길이"))[0]
        if handle:
            extra_parts.append(f"손{handle}")
        note1 = str(c.o.get("_bono_note1") or "").strip()
        if note1:
            extra_parts.append(note1)
        order_no = str(c.o.get("주문번호") or "").strip()
        place = str(it.get("설치장소") or "").strip()
        for part in (order_no, place):
            if part and part not in extra_parts:
                extra_parts.append(part)
        return extra_parts
