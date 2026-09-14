from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import Any

@dataclass
class Delivery:
    방식: str | None = None
    선불착불: str | None = None
    주소: str | None = None
    수령인: str | None = None
    연락처: str | None = None
    전달사항: str | None = None
    발신: str | None = None

@dataclass
class Item:
    품목코드: str | None = None
    종류: str | None = None
    타입: str | None = None
    가로: float | None = None
    세로: float | None = None
    창개수: int = 1
    수량: Any = None
    좌개수: int = 0
    우개수: int = 0
    손잡이방향: str | None = None
    손잡이길이: Any = None
    연창: bool = False
    기재사항: str | None = None
    설치장소: str | None = None
    특이: str | None = None
    원문: str | None = None
    extra: dict[str, Any] = field(default_factory=dict)

@dataclass
class Order:
    거래처: str
    주문번호: str | None = None
    고객명: str | None = None
    전체기재사항: str | None = None
    출고일: str | None = None
    배송: Delivery = field(default_factory=Delivery)
    items: list[Item] = field(default_factory=list)
    source_path: str | None = None
    source_pages: list[int] = field(default_factory=list)
    extra: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict:
        return asdict(self)
