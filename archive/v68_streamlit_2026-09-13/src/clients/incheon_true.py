"""인천)트루(트루갤러리) — v68 동작.

장부: 부속(노피스(1)/노피스(2)/B 원코드 브라켓) 별도 행, 수량 1 고정.
(부속 인식은 extract._postprocess_true, EDI 부속 행·주소행은 erp_edi.build_erp 에 있다.)
"""
from .base import ClientRules


class IncheonTrue(ClientRules):
    code = "인천)트루"

    def ledger_accessory_rows(self, order):
        rows = []
        for accessory in order.get("_true_accessories") or []:
            label = str(accessory.get("표시") or "").strip()
            if label:
                # 숫자 규칙을 추정하지 않고 사용자가 요청한 1개로 고정한다.
                rows.append({"_특수": "부속", "문구": label, "수량": 1})
        return rows
