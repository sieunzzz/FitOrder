"""거래처 규칙 기본값 = v68 의 '그 외 업체' 공통 동작.

업체별 파일은 ClientRules 를 상속해 **다른 부분만** 재정의한다.
겉보기에 비슷해도 업체끼리 합치지 않는다(예: v68 에서 보노와 미래가공은 서로 다른 규칙).

2026-09-13 뼈대 정리: output.py 의 `if client == ...` 분기를 내용 변경 없이 옮겼다.
"""
from order_text import _has_non_di_mix, _note_text, _split_note_parts, _without_parts
from rules import PACKING_ITEM


class ClientRules:
    code = None

    # ───────────── 장부 (ledger_rows.to_rows) ─────────────
    # 택배/화물이면 배송 수령인을 기재사항1 맨 앞에 넣는다.
    prepend_receiver_to_note1 = True
    # 수량 칸이 비었을 때 창개수를 장부 수량 칸에 표시한다.
    ledger_qty_uses_window_count = False
    # 손잡이길이 150 을 장부 길이 칸에 표시하지 않는다.
    hide_handle_150 = False
    # 수령인/연락처가 없어도 선불 표시행을 만든다.
    show_prepaid_without_receiver = False

    def ledger_mark(self, order_no, mark):
        """상호 옆 내부표시. 기본은 CLIENT_INFO 의 내부표시 그대로."""
        return mark

    def ledger_common(self, common, order_no):
        """(가공한 전체기재사항, 공지 목록)."""
        return common, []

    def ledger_customer(self, cust):
        return cust

    def ledger_notes(self, c):
        """장부 기재사항1/2. c = to_rows 현재 행 문맥(SimpleNamespace)."""
        # 전체 공통: 여러 창이면 기재1=공통, 기재2=개인창.
        # 모든 item에 반복된 기재문구는 공통으로 승격하여 대동의
        # `공학1관/공학1관` 같은 중복을 방지한다.
        personal = _without_parts(_split_note_parts(c.note, c.place), c.common_note_parts)
        if c.multi_windows:
            note1 = _note_text(*c.common_note_parts)
            note2 = _note_text(*personal)
        else:
            note1 = _note_text(*c.common_note_parts, *personal)
            note2 = None
        return note1, note2

    def ledger_accessory_rows(self, order):
        """제품행 뒤 · 추가부속 행 앞에 붙는 업체 전용 부속 행."""
        return []

    def ledger_place_rows(self, order):
        """추가부속 행 뒤 · 포장비용 행 앞에 붙는 업체 전용 특수 행."""
        return []

    # ───────────── 경영박사 EDI (erp_edi.build_erp) ─────────────
    erp_vat_zero = False              # 부가세 0 원 처리
    erp_sort_in_ledger_order = False  # 장부 시트/페이지 순서로 EDI 행 정렬
    holding_pack_enabled = True       # 홀딩도어 포장비용(H) 사용
    holding_pack_by_recipient = False  # 수령인이 바뀔 때마다 포장비용(H) 마감
    holding_pack_unit = 1000          # 홀딩 포장비 ㎡당 단가
    blind_packing_item = PACKING_ITEM  # 블라인드 택배/화물 창당 포장비 품목

    def holding_recipient(self, o, it, delivery, di_pack_recipient):
        """홀딩도어 EDI 적요에 쓸 이름/현장."""
        return str(it.get("기재사항") or o.get("고객명") or
                   delivery.get("수령인") or "").strip()

    def holding_note_parts(self, it, recipient):
        """홀딩도어 본품 EDI 적요의 `가로*세로/작동방식` 뒤에 붙는 값들."""
        note_parts = []
        for value in (it.get("기재사항"), it.get("설치장소"), it.get("_수동특이")):
            for part in str(value or "").split("/"):
                part = part.strip()
                if part and part not in note_parts:
                    note_parts.append(part)
        if recipient:
            recipient_parts = [x.strip() for x in str(recipient).split("/") if x.strip()]
            if not recipient_parts or not all(x in note_parts for x in recipient_parts):
                if recipient not in note_parts:
                    note_parts.append(recipient)
        return note_parts

    def erp_blind_memo_parts(self, c):
        """블라인드 EDI 적요의 `가로*세로/1EA` 뒤에 붙는 값들(장부 기재사항1/2는 build_erp 가 맨 뒤에 붙인다).

        c = SimpleNamespace(o, it, client, delivery, order_common_note)
        """
        it = c.it
        extra_parts = []
        # 화면에서 직접 입력한 특이사항도 모든 거래처의 EDI 적요에 반영한다.
        manual_special = str(it.get("_수동특이") or "").strip()
        if manual_special:
            extra_parts.append(manual_special)
        if _has_non_di_mix(it, c.client):
            extra_parts.append("MIX")
        # 기재사항은 전산 적요의 맨 뒤에 보존한다.
        # 전체기재사항 -> 행 기재사항 -> 설치장소 순서로 붙여
        # `퇴로로20/주방`처럼 장부에는 보이던 앞부분이 EDI에서 빠지지 않게 한다.
        tail_parts = []
        for value in (c.order_common_note, it.get("기재사항"), it.get("설치장소")):
            for part in str(value or "").split("/"):
                part = part.strip()
                if part and part not in tail_parts:
                    tail_parts.append(part)
        extra_parts = [x for x in extra_parts if x not in tail_parts]
        extra_parts.extend(tail_parts)
        return extra_parts
