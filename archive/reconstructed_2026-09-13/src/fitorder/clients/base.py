"""거래처 규칙 기본 클래스 = 일반 업체 동작.

업체별 규칙은 이 클래스를 상속해 **다른 부분만** 재정의한다.
겉보기에 비슷한 업체라도 하나로 합치지 않는다. 문서(RULES_MASTER)에
"동일 규칙"이라고 명시된 경우에만 상속으로 위임한다(예: 미래가공 → 보노).
"""
from __future__ import annotations


class ClientRule:
    name: str = ""

    # ---- 주문/품목 값 읽기 (업체 규칙에서 공통으로 쓰는 헬퍼) ----
    @staticmethod
    def order_no(order: dict) -> str:
        return str(order.get("주문번호") or "").strip()

    @staticmethod
    def common(order: dict) -> str:
        return str(order.get("전체기재사항") or "").strip()

    @staticmethod
    def customer(order: dict) -> str:
        return str(order.get("고객명") or "").strip()

    @staticmethod
    def personal(item: dict) -> str:
        return str(item.get("설치장소") or item.get("기재사항") or "").strip()

    @staticmethod
    def join_unique(parts) -> str | None:
        """빈 값 제외, 중복 제거 후 `/`로 연결."""
        return "/".join(dict.fromkeys(x for x in parts if x)) or None

    # ---- 규칙 hook ----
    def notes(self, order: dict, item: dict, item_index: int, total_items: int):
        """기재사항1/2. 일반 업체: 기재1 = 주문번호/공통/고객명, 기재2 = 개인창."""
        return (
            self.join_unique((self.order_no(order), self.common(order), self.customer(order))),
            self.personal(item) or None,
        )
