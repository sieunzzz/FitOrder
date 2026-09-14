"""DI 대일산업 — v68 동작.

장부: 기재사항1만(수령인 + MIX 조합), 수령인 자동 앞붙임 없음.
EDI: 부가세 0, 장부 시트 순서로 정렬, 홀딩 포장비는 수령인별 ㎡×550원.
(수령인 묶음·긴급·MIX·시트 분할 등 DI 장부 배치는 아직 ledger_rows/ledger_book 에 있다.)
"""
from order_text import _note_text
from rules import normalize_handle

from .base import ClientRules


class DI(ClientRules):
    code = "DI"
    prepend_receiver_to_note1 = False
    erp_vat_zero = True
    erp_sort_in_ledger_order = True
    holding_pack_by_recipient = True
    holding_pack_unit = 550

    def ledger_notes(self, c):
        # DI는 기재사항1만 사용: 이름 + MIX 정보. 기재사항2는 사용하지 않는다.
        recv = str(c.it.get("_DIrecipient_context") or c.cust or
                   (c.order.get("배송") or {}).get("수령인") or "").strip()
        mix_note = None
        if c.it.get("_mix_codes"):
            codes = str(c.it.get("_mix_codes") or "").replace(" ", "")
            mix_name = str(c.it.get("_mix_name") or "").strip()
            mix_note = f"{mix_name}{codes}" if mix_name else codes
        note1 = _note_text(recv, mix_note)
        note2 = None
        return note1, note2

    def holding_recipient(self, o, it, delivery, di_pack_recipient):
        return str(it.get("_DIrecipient_context") or it.get("기재사항") or
                   di_pack_recipient or "").strip()

    def holding_note_parts(self, it, recipient):
        note_parts = []
        if recipient:
            note_parts.append(recipient)
        for value in (it.get("_수동특이"), it.get("설치장소")):
            value = str(value or "").strip()
            if value and value not in note_parts:
                note_parts.append(value)
        return note_parts

    def erp_blind_memo_parts(self, c):
        it = c.it
        extra_parts = []
        receiver = str(it.get("기재사항") or "").strip()
        if receiver:
            extra_parts.append(receiver)
        manual_special = str(it.get("_수동특이") or "").strip()
        if manual_special and manual_special not in extra_parts:
            extra_parts.append(manual_special)
        handle = normalize_handle(it.get("손잡이길이"))[0]
        if handle:
            extra_parts.append(f"손{handle}")
        if it.get("_mix_name") and it.get("_mix_codes"):
            mix_codes = str(it["_mix_codes"]).replace(" ", "")
            extra_parts.append(
                mix_codes if it.get("_generic_mix") or str(it.get("_mix_name")).upper() == "MIX"
                else f"{it['_mix_name']}{mix_codes}")
        place = str(it.get("설치장소") or "").strip()
        if place and place not in extra_parts:
            extra_parts.append(place)
        if c.order_common_note:
            for part in [x.strip() for x in c.order_common_note.split("/") if x.strip()]:
                extra_parts = [x for x in extra_parts if x != part]
                extra_parts.append(part)
        return extra_parts
