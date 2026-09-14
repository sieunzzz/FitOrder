"""v68 에서 장부/EDI 전용 분기가 없는 업체들.

업체마다 클래스를 따로 두어, 규칙을 추가할 때 다른 업체에 영향이 없게 한다.
(판독 후처리 분기는 extract.py / app.ingest 에 남아 있다: MS 는 extract, DD 는 app.ingest)
"""
from .base import ClientRules


class Midas(ClientRules):
    code = "M"      # 미더스


class DD(ClientRules):
    code = "DD"     # 대동산업


class WT(ClientRules):
    code = "WT"     # 윈도우투모로우


class MS(ClientRules):
    code = "MS"     # 미성텍스
