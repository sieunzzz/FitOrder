"""휴안 — v68 동작.

장부: 기재사항1만(이름 + 노피스/앙카 등 제작 메모), 수령인 자동 앞붙임 없음.
EDI: 블라인드 적요 = 특이 → 기재/설치장소 → 수령인, 홀딩 적요명은 공백 제거.
"""
import re

from order_text import _note_text, _split_note_parts

from .base import ClientRules


class Huan(ClientRules):
    code = "휴안"
    prepend_receiver_to_note1 = False

    def ledger_notes(self, c):
        # 휴안도 기재사항1만 사용한다. 이름과 노피스/앙카 등 필요한 제작 메모를
        # 한 칸에 모으고 기재사항2는 비운다.
        recv = str(c.cust or (c.order.get("배송") or {}).get("수령인") or "").strip()
        item_parts = _split_note_parts(c.note)
        note1 = _note_text(recv, *item_parts)
        note2 = None
        return note1, note2

    def holding_recipient(self, o, it, delivery, di_pack_recipient):
        # 장부 기재사항의 현장/상호명이 경영박사 적요명이며, 배송 수령인은
        # 주소행에 더 짧게 적힐 수 있다. 제공 정답 EDI는 홀딩 적요명의
        # 띄어쓰기를 제거한다(예: 제이원 송길수 부장님 -> 제이원송길수부장님).
        name = str(it.get("기재사항") or delivery.get("수령인") or
                   o.get("고객명") or "").strip()
        return re.sub(r"\s+", "", name)

    def holding_note_parts(self, it, recipient):
        note_parts = []
        recipient_key = re.sub(r"\s+", "", recipient)
        for value in (it.get("_수동특이"), it.get("기재사항"), it.get("설치장소")):
            for part in str(value or "").split("/"):
                part = part.strip()
                if part and re.sub(r"\s+", "", part) != recipient_key \
                        and part not in note_parts:
                    note_parts.append(part)
        if recipient:
            note_parts.append(recipient)
        return note_parts

    def erp_blind_memo_parts(self, c):
        it = c.it
        extra_parts = []
        receiver = str(c.delivery.get("수령인") or
                       c.o.get("고객명") or "").strip()
        manual_special = str(it.get("_수동특이") or "").strip()
        if manual_special:
            extra_parts.append(manual_special)
        for value in (it.get("기재사항"), it.get("설치장소")):
            for part in str(value or "").split("/"):
                part = part.strip()
                if part and part != receiver and part not in extra_parts:
                    extra_parts.append(part)
        if receiver:
            extra_parts.append(receiver)
        return extra_parts
