"""홀딩도어 품목/장부/EDI 규칙.

v100부터 품목·단가·장부표시는 코드에 하드코딩하지 않고
`data/master/holding/holding_master.xlsx`를 공통 품목장 로더로 읽는다.
SRC만 교체한 환경에서는 `src/holding_master/holding_master.xlsx`를 폴백으로 쓴다.

업무 규칙(편개/양개/자석바/헤베 계산 등)은 이 모듈에 유지하고,
품목 데이터만 외부 품목장으로 분리한다.
"""
from __future__ import annotations

import re
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from master_loader import read_rows, resolve_master_file, master_dir_candidates


# 블라인드 모드에서 홀딩도어를 자동 분류하지 않는 안전장치.
# 홀딩도어 전용 모드는 desktop_workflow가 명시적으로 normalize한다.
HOLDING_FEATURE_ENABLED = False

# 홀딩도어 전용 거래처. 블라인드/롤·콤비 목록과 분리해 화면에서 다른 업체가 섞이지 않게 한다.
HOLDING_CLIENTS = ("M", "휴안", "DI", "구미)경남", "DU", "창문애", "한길")
HOLDING_CLIENT_LABELS = {
    "M": "미더스 (K)",
    "휴안": "휴안 (K)",
    "DI": "대일",
    "구미)경남": "구미)경남 (K)",
    "DU": "두창 (K)",
    "창문애": "창문애 (K)",
    "한길": "한길 (K)",
}


def holding_client_label(client):
    return HOLDING_CLIENT_LABELS.get(str(client or ""), str(client or ""))

UNICODE_FRACTIONS = {"½":"1/2", "⅓":"1/3", "¼":"1/4", "⅕":"1/5", "⅙":"1/6", "⅐":"1/7", "⅛":"1/8", "⅑":"1/9", "⅒":"1/10"}


def normalize_holding_text(value):
    s = str(value or "").strip()
    for k, v in UNICODE_FRACTIONS.items():
        s = s.replace(k, v)
    return re.sub(r"\s+", " ", s).strip()


def _alias_key(value):
    s = normalize_holding_text(value).upper().replace("Ⅰ", "I").replace("Ⅱ", "II")
    # 발주서의 자연어 겹수 표기를 품목장의 장부 표기(I/II)로 먼저 통일한다.
    # 예: 회색 한겹 -> 회색I, 화이트 이중 -> 화이트II.
    # 이 변환을 별칭 key 단계에서 처리하면 Excel/이미지/수동수정이 모두 같은 매칭 경로를 탄다.
    s = (s.replace("한겹", "I").replace("1겹", "I")
           .replace("이중", "II").replace("두겹", "II").replace("2겹", "II"))
    s = re.sub(r"\s+", "", s).replace("홀딩도어", "").replace("홀딩", "").replace("자바라", "")
    if s.startswith("H") and not re.match(r"^H(?:\d|방염)", s):
        s = s[1:]
    return s.replace("+상하로라", "").strip()


def _num(value):
    if value in (None, ""):
        return 0
    try:
        f = float(value)
        return int(f) if f.is_integer() else f
    except (TypeError, ValueError):
        return value


def _optional_holding_extra_path():
    """사용자 제공 홀딩 추가비용 품목장을 찾는다.

    배포 ZIP의 한글 파일명 문제를 피하기 위해 실제 배포명은
    ``holding_extra_costs.xls``로 고정한다. 과거 파일명도 읽기만 지원한다.
    """
    names = ("holding_extra_costs.xls", "홀딩도어 추가비용.xls",
             "holding_extra_costs.xlsx", "홀딩도어 추가비용.xlsx")
    for directory in master_dir_candidates("holding"):
        for name in names:
            path = directory / name
            if path.is_file():
                return path
    return None


def _catalog_row(raw, default_category="홀딩도어"):
    name = str(raw.get("품명") or raw.get("관리코드") or "").strip()
    if not name:
        return None
    row = {
        "품명": name,
        "관리코드": str(raw.get("관리코드") or name).strip(),
        "대분류": str(raw.get("대분류") or default_category).strip(),
        "규격": raw.get("규격") or "",
        "비고2": raw.get("비고2") or "",
    }
    for tier in ("A","B","C","D","E","F","G","H","I","J","L"):
        row[f"출고{tier}가"] = _num(raw.get(f"출고{tier}가")) or 0
    inferred = str(raw.get("추정단가등급") or "").strip()
    if inferred:
        row["_inferred_tiers"] = {x.strip() for x in inferred.split(",") if x.strip()}
    return row


class HoldingCatalog:
    """외부 홀딩 품목장을 메모리 인덱스로 변환한다."""

    def __init__(self, path: str | Path | None = None):
        self.path = Path(path) if path else resolve_master_file(
            "holding", ("holding_master.xlsx", "홀딩도어_품목장.xlsx", "홀딩도어 품목 리스트.xlsx", "홀딩도어 품목 리스트.xls")
        )
        rows = read_rows(self.path, required_headers=("관리코드", "품명"), preferred_sheet="품목")
        self.products_by_name: dict[str, dict] = {}
        self.default_by_code: dict[str, str] = {}
        self.ledger_labels: dict[str, str] = {}
        self.aliases: dict[str, str] = {}
        self.extra_products: dict[str, dict] = {}

        for raw in rows:
            row = _catalog_row(raw)
            if not row:
                continue
            name = row["품명"]
            kind = str(raw.get("유형") or "제작품").strip()
            default_code = str(raw.get("기본코드") or "").strip()
            ledger = str(raw.get("장부표시") or "").strip()
            if kind == "추가비용":
                self.extra_products[name] = row
                continue

            self.products_by_name[name] = row
            if default_code:
                self.default_by_code[default_code] = name
                if ledger:
                    self.ledger_labels[default_code] = ledger

            # 품명/관리코드/기본코드 자체도 항상 별칭으로 등록한다.
            for alias in (name, row["관리코드"], default_code):
                if alias:
                    self.aliases[_alias_key(alias)] = name
            for alias in re.split(r"\s*\|\s*", str(raw.get("별칭") or "")):
                if alias.strip():
                    self.aliases[_alias_key(alias)] = name

        # 사용자 제공 `홀딩도어 추가비용.xls`를 별도 마스터로 읽는다.
        # 기본 holding_master에 있던 임시/추정 추가비용과 이름이 같으면
        # 사용자 제공 원본의 실제 가격으로 덮어쓴다.
        extra_path = _optional_holding_extra_path()
        self.extra_master_path = extra_path
        if extra_path:
            extra_rows = read_rows(extra_path, required_headers=("관리코드", "품명"))
            for raw in extra_rows:
                row = _catalog_row(raw, default_category="부속")
                if row:
                    self.extra_products[row["품명"]] = row

        if not self.products_by_name:
            raise RuntimeError(f"홀딩도어 품목장이 비어 있습니다: {self.path}")

    def product(self, name: str | None):
        row = self.products_by_name.get(str(name or ""))
        return dict(row) if row else None


_CATALOG: HoldingCatalog | None = None

# 하위 호환용 공개 dict. 최초 사용 시 _sync_public_maps()로 채운다.
HOLDING_PRODUCTS_BY_NAME: dict[str, dict] = {}
HOLDING_DEFAULT_BY_CODE: dict[str, str] = {}
HOLDING_LEDGER_LABELS: dict[str, str] = {}
HOLDING_ALIASES: dict[str, str] = {}
HOLDING_EXTRA_PRODUCTS: dict[str, dict] = {}


def get_holding_catalog() -> HoldingCatalog:
    global _CATALOG
    if _CATALOG is None:
        _CATALOG = HoldingCatalog()
        _sync_public_maps()
    return _CATALOG


def _sync_public_maps():
    if _CATALOG is None:
        return
    HOLDING_PRODUCTS_BY_NAME.clear(); HOLDING_PRODUCTS_BY_NAME.update(_CATALOG.products_by_name)
    HOLDING_DEFAULT_BY_CODE.clear(); HOLDING_DEFAULT_BY_CODE.update(_CATALOG.default_by_code)
    HOLDING_LEDGER_LABELS.clear(); HOLDING_LEDGER_LABELS.update(_CATALOG.ledger_labels)
    # 공개 별칭 map은 사람이 읽을 필요가 없으므로 정규화 key 기준으로 유지한다.
    HOLDING_ALIASES.clear(); HOLDING_ALIASES.update(_CATALOG.aliases)
    HOLDING_EXTRA_PRODUCTS.clear(); HOLDING_EXTRA_PRODUCTS.update(_CATALOG.extra_products)


def _with_meta(name):
    cat = get_holding_catalog()
    row = cat.products_by_name.get(str(name or ""))
    if not row:
        return None
    m = re.match(r"^(H[^\s(]+)", str(name))
    code = m.group(1) if m else str(name)
    return dict(row, 코드=code, 장부표시=cat.ledger_labels.get(code), _holding_product_name=name)


def _holding_layer_from_text(value):
    """발주서 자연어에서 한겹/이중을 장부 표기 I/II로 판별한다."""
    raw = normalize_holding_text(value)
    compact = re.sub(r"\s+", "", raw).upper().replace("Ⅰ", "I").replace("Ⅱ", "II")
    if any(token in compact for token in ("이중", "두겹", "2겹")):
        return "II"
    if any(token in compact for token in ("한겹", "1겹")):
        return "I"
    # 이미 표준 표기로 입력된 '화이트 II', '회색 I'도 지원한다.
    if re.search(r"(?:^|[^A-Z])II(?:$|[^A-Z])", raw.upper()):
        return "II"
    if re.search(r"(?:^|[^A-Z])I(?:$|[^A-Z])", raw.upper()):
        return "I"
    return None


def _holding_name_from_color_layer(cat, value):
    """긴 문장 안의 색상+겹수로 홀딩 품목을 찾는다.

    예: '홀딩도어(한겹)/3(화이트) 281x252 1/2 레일연결부속포함'
        -> H003 (화이트)
    정확한 코드/별칭 매칭이 먼저이고 이 함수는 그 다음 안전 폴백이다.
    """
    layer = _holding_layer_from_text(value)
    if not layer:
        return None
    raw_key = _alias_key(value)
    candidates = []
    # 품목장의 별칭 중 '색상I/색상II' 형태를 재사용한다. 코드 숫자 별칭은 제외한다.
    for alias_key, name in cat.aliases.items():
        suffix = "II" if alias_key.endswith("II") else ("I" if alias_key.endswith("I") else None)
        if suffix != layer:
            continue
        base = alias_key[:-len(suffix)]
        if len(base) < 2 or not re.search(r"[가-힣A-Z]", base):
            continue
        if base in raw_key:
            candidates.append((len(base), name))
    if not candidates:
        return None
    candidates.sort(reverse=True)
    return candidates[0][1]


def holding_product_from_text(value):
    cat = get_holding_catalog()
    raw = normalize_holding_text(value)
    name = cat.aliases.get(_alias_key(raw))
    if name:
        return _with_meta(name)
    m = re.search(r"\bH\s*0*(\d{1,3}(?:-1)?)\b", raw, re.I)
    if m:
        num = m.group(1)
        if "-" in num:
            base, suf = num.split("-", 1)
            code = f"H{int(base):03d}-{suf}"
        else:
            code = f"H{int(num):03d}"
        return _with_meta(cat.default_by_code.get(code))
    fm = re.search(r"H?\s*방염\s*F?([123])", raw, re.I)
    if fm:
        return _with_meta(cat.default_by_code.get(f"H방염F{fm.group(1)}"))
    fuzzy_name = _holding_name_from_color_layer(cat, raw)
    return _with_meta(fuzzy_name) if fuzzy_name else None


def holding_product_by_item(item):
    cat = get_holding_catalog()
    name = item.get("_holding_product_name")
    if name in cat.products_by_name:
        return _with_meta(name)
    code = str(item.get("품목코드") or "").strip()
    if code in cat.default_by_code:
        return _with_meta(cat.default_by_code[code])
    return holding_product_from_text(item.get("색상원문") or code)


def holding_accessory_from_text(value, default_color="화이트"):
    s = normalize_holding_text(value)
    if "레일연결부속" not in s and "라운드부속" not in s:
        return None
    color = "블랙" if "블랙" in s else (default_color or "화이트")
    if "레일연결부속" in s:
        name = f"H레일연결부속({color})"
    else:
        name = f"H라운드부속({color}/2022)"
    row = _with_meta(name)
    if row:
        row.update(
            장부표시=("레일연결부속" if "레일연결" in name else "라운드부속"),
            부속색상=color,
            _holding_accessory=True,
        )
    return row


def holding_extra_product(name):
    row = get_holding_catalog().extra_products.get(str(name or ""))
    return dict(row) if row else None


def apply_holding_feature_gate(order):
    """블라인드 모드에서는 잘못 붙은 홀딩 내부 표식을 제거한다."""
    if HOLDING_FEATURE_ENABLED or not isinstance(order, dict):
        return 0
    changed = 0
    for item in order.get("items") or []:
        if not (item.get("_product_group") == "holding" or item.get("_holding_accessory")
                or str(item.get("제품군") or "").strip() == "홀딩도어"):
            continue
        item["제품군"] = "블라인드"
        exc = str(item.get("예외품목") or "").strip()
        if exc and ("홀딩" in exc or "자바라" in exc):
            item["예외품목"] = None
        item.pop("_product_group", None)
        item.pop("_ledger_prefix", None)
        item.pop("_holding_disabled", None)
        for key in list(item):
            if key.startswith("_holding_"):
                item.pop(key, None)
        item["홀딩방식"] = None
        item["홀딩레일"] = None
        item["홀딩상하로라"] = False
        item["홀딩부속"] = None
        item["홀딩부속색상"] = None
        changed += 1
    order.pop("_holding_disabled_count", None)
    return changed


def holding_ledger_text(item):
    if item.get("_holding_unmatched"):
        return " H [품목확인]"
    if item.get("_holding_accessory"):
        label = item.get("_holding_accessory_label") or (
            "레일연결부속" if "레일연결" in str(item.get("_holding_product_name")) else "라운드부속"
        )
        color = item.get("_holding_accessory_color")
        return f" H {label}" + (f"({color})" if color == "블랙" else "")
    hit = holding_product_by_item(item)
    label = item.get("_holding_label") or (hit.get("장부표시") if hit else None) \
        or str(item.get("색상원문") or item.get("품목코드") or "").strip()
    text = f" H {label}".rstrip()
    if item.get("_holding_upper_roller") and "+상하로라" not in text:
        text += " +상하로라"
    return text


def normalize_holding_operation(value):
    """홀딩 작동방식을 장부/작업지시서/EDI 공통 표기로 정규화한다.

    편개 -> 편, 양개 -> 양, 이등분/2등분 -> 1/2, 양자석 -> 양자석.
    1/3~1/10 및 양자석 조합은 기존 분수 표기를 그대로 유지한다.
    """
    s = normalize_holding_text(value)
    if not s or s in {"-", "0"}:
        return "편"
    s = re.sub(r"\s+", "", s).replace("편개", "편").replace("양개양자석", "양개 양자석")
    if s in {"이등분", "2등분"}:
        return "1/2"
    if s in {"양개", "양"}:
        return "양"
    if s == "양자석":
        return "양자석"
    if s.startswith("고정"):
        return "고정"
    return s


def looks_like_holding_operation(value):
    s = normalize_holding_operation(value)
    return bool(s in {"편", "양", "양자석", "고정", "양개 양자석"} or re.match(r"^1/(?:[2-9]|10)", s))


def holding_magnet_bars(operation):
    op = normalize_holding_operation(operation)
    if op == "고정":
        return 0
    m = re.search(r"\((.+)\)", op)
    if m and "+" in m.group(1):
        total = 0
        for seg in m.group(1).split("+"):
            if "양자석" in seg:
                total += 2
            elif re.search(r"(?:양개|양)(?!자석)", seg):
                total += 2
            else:
                total += 1
        return total
    frac = re.match(r"^1/(10|[2-9])", op)
    if frac:
        n = int(frac.group(1))
        return 2 * n if "양자석" in op else n
    if "양개 양자석" in op or "양개양자석" in op:
        return 4
    if op == "양자석":
        return 2
    if op in {"양", "양개"}:
        return 2
    return 1


def holding_magnet_charge_qty(operation):
    return max(0, holding_magnet_bars(operation) - 1)


def holding_magnet_extra_names(operation, height_cm=None):
    """추가 자석바 수에 맞는 실제 경영박사 품목명을 반환한다.

    최초 자석바 1개는 편개 기본 구성이라 청구하지 않는다. 추가분은
    사용자 제공 품목장의 ``+자석바1``~``+자석바10``을 사용하며,
    세로 300cm 이상이면 ``/300cm이상`` 품목을 사용한다. 10개 초과는
    +10 품목과 나머지 품목으로 안전하게 나눈다.
    """
    remain = holding_magnet_charge_qty(operation)
    try:
        tall = float(height_cm) >= 300 if height_cm not in (None, "") else False
    except (TypeError, ValueError):
        tall = False
    suffix = "/300cm이상" if tall else ""
    names = []
    while remain > 0:
        n = min(10, remain)
        names.append(f"H추가비용(+자석바{n}{suffix})")
        remain -= n
    return names


def holding_calc(width_cm, height_cm, unit_price):
    h = max(Decimal(str(height_cm)), Decimal("150"))
    raw = (Decimal(str(width_cm)) / 100) * (h / 100)
    hebe = raw.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    qty = max(hebe, Decimal("2.5"))
    amount = (Decimal(str(unit_price)) * qty).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    vat = (amount / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return {"헤베":hebe, "수량":qty, "계산세로":h, "단가":unit_price, "금액":amount, "부가세":vat}


def accessory_qty(value):
    if value in (None, ""):
        return 1
    m = re.search(r"(\d+(?:\.\d+)?)", str(value))
    if not m:
        return 1
    n = float(m.group(1))
    return int(n) if n.is_integer() else n
