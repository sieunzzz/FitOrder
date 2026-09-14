"""아지트 — v68 동작.

장부: 추가부속 행 뒤에 `☆금빛커텐` / `☆시온가공소` 행(기본 시온가공소).
(에어캡+포장 판별은 app.ingest 에 있다.)
"""
from .base import ClientRules


class Azit(ClientRules):
    code = "아지트"

    def ledger_place_rows(self, order):
        return [{"_특수": "☆", "문구": order.get("_az_place") or "시온가공소"}]
