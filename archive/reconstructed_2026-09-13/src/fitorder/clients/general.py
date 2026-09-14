"""RULES_MASTER에 개별 규칙이 아직 없는 업체들.

지금은 모두 일반 업체 규칙을 쓰지만 업체마다 클래스를 따로 둔다.
개별 규칙이 확정되면 해당 클래스를 별도 파일로 옮겨 재정의한다.
"""
from .base import ClientRule


class DU(ClientRule):
    name = "DU"


class WT(ClientRule):
    name = "WT"


class MS(ClientRule):
    name = "MS"


class UniTNS(ClientRule):
    name = "유앤아이티엔에스"


class Ajit(ClientRule):
    name = "아지트"


class RT(ClientRule):
    name = "RT"
