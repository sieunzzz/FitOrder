"""제품군별 품목장 객체를 한 곳에서 생성/캐시한다."""
from __future__ import annotations

from functools import lru_cache


VALID_PRODUCT_MODES = {"blind", "roll_combo", "holding"}


def normalize_product_mode(mode: str | None) -> str:
    mode = str(mode or "blind").strip().lower()
    if mode not in VALID_PRODUCT_MODES:
        raise ValueError(f"알 수 없는 제품군입니다: {mode}")
    return mode


@lru_cache(maxsize=1)
def _blind_master():
    from rules import Master
    return Master()


@lru_cache(maxsize=1)
def _roll_master():
    from roll_combo import RollComboMaster, resolve_roll_master_dir
    return RollComboMaster(resolve_roll_master_dir())


def get_master_for_mode(mode: str = "blind"):
    mode = normalize_product_mode(mode)
    if mode == "roll_combo":
        return _roll_master()
    # 블라인드/홀딩은 거래처 단가 체계와 EDI 공통 API를 공유한다.
    # 홀딩 품목 데이터 자체는 holding.py의 외부 HoldingCatalog에서 조회한다.
    return _blind_master()
