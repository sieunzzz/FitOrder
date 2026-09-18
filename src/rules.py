"""
FitOrder 규칙 엔진 — 마스터 조회 / 금액 계산 / 검증

사용:
    from rules import Master, calc_erp, validate
    M = Master("../data/master/fitorder_master.xlsx")
    row = M.find_item("200", "원코드", "L자", "DU")
    issues = validate(order_dict, M)
"""
import re
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from master_loader import read_rows, resolve_master_file
from holding import (holding_accessory_from_text, holding_extra_product,
                     holding_product_by_item, holding_magnet_extra_names)


# 기본 배송 안내는 제작/EDI 메모가 아니라 업체의 상시 안내 문구이므로
# 사전점검·장부·작업지시서·EDI 어느 곳에도 반복 출력하지 않는다.
DEFAULT_DELIVERY_NOTICES = {"배송전연락", "던지지마세요"}

# 사용자 확정 규칙: 비닐 포장 관련 문구(비닐포장·겉비닐·걷비닐·창별 비닐포장·비닐 겉면 기재 등)는
# 발주서·사전점검 어디에서 들어와도 장부·작업지시서·경영박사 EDI 어디에도 적지 않는다.
VINYL_RX = re.compile(r"비\s*닐")


def strip_vinyl_text(value):
    """`/`·줄바꿈으로 나뉜 문구 중 비닐 포장 관련 칸을 통째로 뺀다.

    비닐 문구가 없으면 값을 그대로(자료형·공백 포함) 돌려준다. 모두 빠지면 None.
    """
    if value is None or not VINYL_RX.search(str(value)):
        return value
    kept = [part.strip() for part in re.split(r"[\n/]+", str(value))
            if part.strip() and not VINYL_RX.search(part)]
    return "/".join(kept) or None


# ─────────────────────────────────────────────
# 비대일 블라인드 MIX (2026-09-13 사용자 확정 규칙)
# - "믹스/mix/MIX" 단어 + `200+125` 조합이 있으면 MIX 로 처리한다. 대일(DI)은 기존 믹스 형식을 유지한다.
# - 섞이는 길이: 비율(%)이면 세로에 비례, 길이(cm)면 그대로, 없으면 똑같이 나눈다. 0.5 단위 반올림.
# ─────────────────────────────────────────────
MIX_WORD_RX = re.compile(r"(?<![A-Za-z])mix(?![A-Za-z])|믹스", re.I)
_MIX_CODE = r"\d{3,4}(?:FP|P)?"
_MIX_UNIT = r"(?:%|cm|㎝)?"
# 수량 표기는 괄호 `102(70%)` 와 꺾쇠 `102<7>` 를 모두 지원한다(꺾쇠는 비율 표기).
_MIX_AMOUNT = (rf"(?:\s*(?:\(\s*\d+(?:\.\d+)?\s*{_MIX_UNIT}\s*\)"
               rf"|[<〈＜]\s*\d+(?:\.\d+)?\s*{_MIX_UNIT}\s*[>〉＞]))?")
# 코드 사이 구분은 `+`, `:`, `/` 를 인정하고 `상`/`하` 같은 위치 표기는 건너뛴다.
# 예: `상 102<7> : 하 870<3>`
MIX_COMBO_RX = re.compile(
    rf"(?<![\d.]){_MIX_CODE}{_MIX_AMOUNT}"
    rf"(?:\s*[+＋:：/]\s*(?:[상하]\s*[:：]?\s*)?{_MIX_CODE}{_MIX_AMOUNT})+", re.I)
_MIX_PART_RX = re.compile(
    rf"({_MIX_CODE})\s*(?:([(<〈＜])\s*(\d+(?:\.\d+)?)\s*({_MIX_UNIT})\s*[)>〉＞])?", re.I)
# `상 102<7> : 하 870<3>` 처럼 상/하 색상 조합이나 꺾쇠 비율이 보이면
# "믹스" 글자가 없어도 MIX 주문으로 본다(2026-09-16).
MIX_LAYOUT_RX = re.compile(
    rf"(?:[상하]\s*[:：]?\s*{_MIX_CODE}{_MIX_AMOUNT}\s*[+＋:：/]\s*[상하]\s*[:：]?\s*{_MIX_CODE}"
    rf"|{_MIX_CODE}\s*[<〈＜]\s*\d+(?:\.\d+)?\s*{_MIX_UNIT}\s*[>〉＞]"
    rf"\s*[+＋:：/]\s*(?:[상하]\s*[:：]?\s*)?{_MIX_CODE})", re.I)


def round_half(value):
    """0.5 단위 반올림. 예: 120.6→120.5, 13.4→13.5, 49.5→49.5, 108→108."""
    return int(float(value) * 2 + 0.5) / 2


def mix_length_text(value):
    if value in (None, ""):
        return ""
    v = round_half(value)
    return str(int(v)) if float(v).is_integer() else f"{v:g}"


def parse_mix_parts(text, height):
    """`200(90%)+125(10%)` / `200(108cm)+125(12cm)` / `102+500` → [{"코드", "길이"}].

    코드가 2개 미만이면 None. 세로를 알 수 없으면 길이는 None.
    """
    match = MIX_COMBO_RX.search(str(text or ""))
    if not match:
        return None
    raw = [(code.upper(), (bracket or ""), float(num) if num else None, (unit or "").lower())
           for code, bracket, num, unit in _MIX_PART_RX.findall(match.group(0))]
    if len(raw) < 2:
        return None
    try:
        h = float(height)
    except (TypeError, ValueError):
        h = None
    units = {u for _, _, n, u in raw if n is not None and u}
    nums = [n for _, _, n, _ in raw if n is not None]
    total = sum(nums)
    # 단위(%/cm) 표기가 없는 숫자는 길이 합이 세로와 같을 때만 길이로 본다.
    # 합이 다르면 비율이다: `7:3`, `<7>·<3>`, `70+30` 모두 세로를 그 비율로 나눈다.
    as_ratio = bool(nums and not units and len(nums) == len(raw) and total > 0
                    and h is not None and abs(total - h) > 1.0)
    lengths = []
    for _, _bracket, num, unit in raw:
        if num is None:
            lengths.append(None)
        elif unit == "%":
            lengths.append(h * num / 100 if h is not None else None)
        elif as_ratio:
            lengths.append(h * num / total)
        else:
            lengths.append(num)
    missing = [i for i, v in enumerate(lengths) if v is None]
    if missing and h is not None:
        known = sum(v for v in lengths if v is not None)
        share = max(0.0, h - known) / len(missing)
        for i in missing:
            lengths[i] = share
    return [{"코드": code, "길이": (round_half(v) if v is not None else None)}
            for (code, _, _, _), v in zip(raw, lengths)]


def mix_memo_text(item):
    """경영박사 적요/상태 표시용 `200(108)+125(12)`."""
    parts = item.get("_mix_parts") or []
    return "+".join(f"{p.get('코드')}({mix_length_text(p.get('길이'))})"
                    if p.get("길이") is not None else str(p.get("코드")) for p in parts)


def generic_mix_product_name(kind, type_, variant=None):
    """경영박사 일반 MIX 품목명. 예: B-MIX-투코드/25mm, B-MIX-L18/원코드/25mm."""
    kind = kind or "투코드"
    family = variant if variant in ("L18", "L21") else ("L18" if type_ == "L자" else None)
    return f"B-MIX-{family}/{kind}/25mm" if family else f"B-MIX-{kind}/25mm"


def mix_extra_cost_name(type_):
    return "B추가비용-MIX/Ltype" if type_ == "L자" else "B추가비용-MIX"


def clean_delivery_notice(value):
    """기본 상시 배송문구와 비닐 포장 문구를 제거하고 실제 전달사항만 보존한다."""
    text = str(value or "").strip()
    if not text:
        return None
    parts = re.split(r"[\n/|]+", text)
    kept = []
    for part in parts:
        part = part.strip(" ,.-")
        if not part or VINYL_RX.search(part):
            continue
        normalized = re.sub(r"[^가-힣A-Za-z0-9]", "", part)
        if normalized in DEFAULT_DELIVERY_NOTICES:
            continue
        if part not in kept:
            kept.append(part)
    return "/".join(kept) or None

# ─────────────────────────────────────────────
# 거래처
# ─────────────────────────────────────────────
CLIENT_INFO = {
    # 상호: (관리코드, 단가컬럼, 기본배송, 내부표시, 입력형태)
    "DI":   ("대구",       "출고L가", "DI/SP",   "",    "excel"),
    "SP":   ("대구",       "출고H가", "배달",     "",    "pdf"),
    "휴안": ("광주",       "출고I가", "택배",     "(K)", "excel"),
    "M":    ("대구",       "출고I가", None,      "(K)", "image"),
    "DU":   ("대구광역시", "출고I가", None,      "(K)", "image"),
    "RT":   ("경기도",     "출고I가", "판별필요", "(K)", "image"),
    "JO":   ("대구광역시", "출고I가", None,      "(K)", "image"),
    "JL":   ("대구",       "출고I가", None,      "",    "image"),
    "DD":   ("대구",       "출고I가", "배달",    "(K)", "image"),
    "유앤": ("대구광역시", "출고A가", "택배",    "(K)", "excel"),
    "아지트": ("대구",     "출고B가", "배달",    "(K)", "image"),
    "WT":   ("경상북도",   "출고I가", "내사",    "(F)", "image"),
    "인천)트루": ("인천",  "출고A가", "택배",    "(K)", "image"),
    "미래가공": ("대구",   "출고A가", "화물",    "(K)", "excel"),
    "보노": ("경기도",     "출고A가", "화물",    "(K)", "excel"),
    "MS":   ("대구광역시", "출고A가", "배달",    "(K)", "image"),
    # 홀딩도어 전용 거래처 (경영박사 거래처 화면 기준)
    "구미)경남": ("구미)경남블라인드", "출고A가", "택배", "(K)", "image"),
    "창문애": ("창문애아트블라인드", "출고I가", "택배", "(K)", "image"),
    "한길": ("경산)한길산업", "출고A가", "내사", "(K)", "image"),
}

# AI 이미지/PDF 판독의 배송 방식은 모델 추측보다 이 표의 기본값을 우선한다.
# 단, 원문에 배송 방식이 명시되어 있을 때만 해당 방식으로 덮어쓴다.
# 이렇게 하면 "내일 출고", "빠른 출고" 같은 제작/일정 문구가 택배·화물로
# 오판되는 것을 막으면서, 실제로 "택배", "화물", "경동", "대신화물" 등이
# 적힌 주문은 예외 배송으로 처리할 수 있다.
_FIXED_DELIVERY_DEFAULTS = {"배달", "택배", "화물", "내사"}


def _delivery_evidence_text(order):
    """가공된 기재/배송 필드는 제외하고 발주서 원문 필드만 모은다."""
    chunks = [order.get("전체원문")]
    for item in order.get("items") or []:
        chunks.append(item.get("원문"))
    return " ".join(str(x or "") for x in chunks)


def explicit_delivery_mode(order):
    """발주서 원문에 *명시된* 배송 방식만 반환한다.

    여러 배송 라벨이 동시에 보이는 양식(예: 화물( ) / 택배(O))은
    모델이 체크 위치를 읽어 낸 현재 배송값이 원문 후보 안에 있을 때만
    그 값을 보조적으로 인정한다. 아무 명시가 없으면 None이다.
    """
    text = _delivery_evidence_text(order)
    compact = re.sub(r"\s+", "", text)
    if not compact:
        return None

    found = set()
    carrier_freight = bool(
        re.search(r"경동(?:택배|화물)?", compact)
        or re.search(r"대신(?:택배|화물)", compact)
        or re.search(r"(?:선불|착불)대신[-:/]", compact)
    )
    # 경동택배/대신택배는 공장 배송 구분에서는 화물 계열로 취급하므로,
    # 그 상호 안의 '택배' 글자를 일반 택배 근거로 중복 집계하지 않는다.
    generic_text = re.sub(r"(?:경동|대신)\s*택\s*배", " ", text)
    if re.search(r"택\s*배", generic_text):
        found.add("택배")
    if re.search(r"화\s*물", text) or carrier_freight:
        found.add("화물")
    if re.search(r"내\s*사(?:\s*픽업)?|픽\s*업", text):
        found.add("내사")
    if re.search(r"배\s*달|퀵", text):
        found.add("배달")

    if len(found) == 1:
        return next(iter(found))

    # 납품구분 표처럼 택배/화물 문자가 둘 다 인쇄된 경우에는
    # 모델이 읽은 O 체크 결과가 원문 후보 중 하나일 때만 사용한다.
    current = str((order.get("배송") or {}).get("방식") or "").strip()
    if current in found:
        return current
    return None


def apply_delivery_policy(order):
    """업체 기본배송 + 원문 명시 예외로 AI 배송 판독을 결정적으로 보정한다.

    CLIENT_INFO의 기본배송이 배달/택배/화물/내사로 확정된 업체만 대상이다.
    RT(판별필요), 기본값 None, DI의 특수 기본표기 같은 경우는 기존 판독을 유지한다.
    이 함수는 AI 이미지/PDF 추출 직후에 호출하며, 사전점검에서 사용자가
    직접 수정한 배송값에는 다시 적용하지 않는다.
    """
    client = order.get("거래처")
    default = CLIENT_INFO.get(client, (None, None, None))[2]
    if default not in _FIXED_DELIVERY_DEFAULTS:
        return order

    delivery = order.setdefault("배송", {})
    explicit = explicit_delivery_mode(order)
    delivery["방식"] = explicit or default
    return order

# 사용 중단 거래처. 화면·자동 판별·추출 스키마에는 포함하지 않지만
# 나중에 복구할 수 있도록 설정값은 보관한다.
DISABLED_CLIENT_INFO = {}

# 경영박사 EDI에 기록할 실제 등록 상호. 장부/FitOrder 표시는 CLIENT_INFO의 키다.
ERP_CLIENT_NAME = {
    "유앤": "주식회사 유앤아이티엔에스",
    "아지트": "아지트(GB)",
    "WT": "WT",
    "인천)트루": "인천)주식회사 트루갤러리",
    "미래가공": "주식회사 미래가공",
    "보노": "안산)주식회사 보노",
    "MS": "MS",
    "구미)경남": "구미)경남블라인드",
    "창문애": "창문애아트블라인드",
    "한길": "경산)한길산업",
}

# 세로 규격 반복표시 제외 업체. 휴안도 같은 사람·같은 색상·같은
# 세로가 이어지면 두 번째 창부터 "를 사용한다.
NO_DITTO = set()

# 창이 1개여도 설치장소를 기재사항에 합치지 않는 거래처
# (JO 는 기재사항 열에 주문번호만, 설치장소는 다음 열에 적는다)
# 한 창짜리 주문이어도 개별 기재사항(설치장소)을 기재사항2에 따로 적는 거래처.
# SP: 기재1 = 발주번호 + 공통기재사항, 기재2 = 개별 기재사항 (2026-09-18)
NO_MERGE_PLACE = {"JO", "SP"}

# 일반 블라인드 공통 포장비용. v64부터 거래처와 무관하게 실제 배송이
# 택배/화물이면 창당 1회 부과한다. 보노는 아래 전용 품목/단가를 사용한다.
PACKING_ITEM = {"품명": "포장비용(25mm)", "규격": "창당", "단가": 500}
# 구버전 호환용 상수(실제 적용 여부는 output._needs_window_packing에서 결정).
PACKING_CLIENTS = set()
# 보노 실제 경영박사 사용 품목. 창당 700원.
BONO_PACKING_ITEM = {"품명": "포장비용(B/R/C/S/HC)", "규격": "창당/히트", "단가": 700}
# 미래가공은 보노와 동일한 발주/기재/포장 규칙을 사용하되 거래처명/ERP 상호는 각각 유지한다.
BONO_STYLE_CLIENTS = {"보노", "미래가공"}
# 보노 계열: 블라인드 보노/미래가공 + 롤·콤비 안산)보노/미래가공(같은 발주/기재/포장/경영박사 규칙).
BONO_LIKE_CLIENTS = BONO_STYLE_CLIENTS | {"안산)보노"}

# 장부 출고일 옆 배송방식 표기를 생략하는 거래처.
# 이 업체들은 기본 배송지뿐 아니라 별도 배송지로 보내는 경우에도
# 현장 장부에는 `(택배)` / `(화물)`을 적지 않는 것이 정답 양식이다.
# 롤/콤비의 안산)보노도 같은 보노 규칙을 사용한다.
NO_DELIVERY_LABEL_CLIENTS = {"휴안", "보노", "안산)보노", "미래가공"}

# 거래처 담당자(발주를 넣어 주는 직원) 이름. 주문마다 반복해서 들어오지만
# 실제 주문 정보가 아니므로 장부/작업지시서/경영박사 기재사항·적요에는 적지 않는다.
# (2026-09-16 사용자 확정: JL 담당자 김현경)
CLIENT_STAFF_NAMES = {"JL": {"김현경"}}


# 이름 뒤에 붙는 호칭. `김현경님`, `김현경 대리`도 같은 사람으로 본다.
STAFF_TITLE_RX = re.compile(r"\s*(?:님|씨|담당(?:자)?|대리|주임|과장|차장|부장|실장|이사|사장)\s*$")


def is_client_staff_name(client, value):
    """거래처 담당자 이름이면 True. 공백과 뒤에 붙은 호칭은 무시한다."""
    names = CLIENT_STAFF_NAMES.get(client)
    if not names:
        return False
    text = STAFF_TITLE_RX.sub("", str(value or "").strip())
    text = re.sub(r"\s+", "", text)
    return bool(text) and any(text == re.sub(r"\s+", "", n) for n in names)

# 인천)트루 전용 추가부속. 장부/작업지시서에는 사용자 현장 표기를,
# 경영박사에는 실제 등록 품목명을 사용한다. 노피스브라켓(2EA)은
# 사용자 제공 품목코드표에는 있으나 현재 로컬 마스터에 단가가 없으므로
# 단가를 임의로 만들지 않는다.
TRUE_ACCESSORY_PRODUCTS = {
    # 사용자 제공 인천)트루 품목코드표 기준. 숫자 규칙은 추정하지 않고 모두 1개로 고정한다.
    "노피스(2)": {
        "표시": "노피스(2)", "품명": "노피스브라켓(2EA)",
        "관리코드": "노피스브라켓(2EA)", "수량": 1,
    },
    "노피스(1)": {
        "표시": "노피스(1)", "품명": "노피스브라켓(1EA)-커튼박스용",
        "관리코드": "노피스브라켓(1EA)-커튼박스용", "수량": 1,
    },
    "B 원코드 브라켓": {
        "표시": "B 원코드 브라켓", "품명": "B브라켓(25mm원코드/구)",
        "관리코드": "B브라켓(25mm원코드/구)", "수량": 1,
    },
}

MIN_WIDTH = {"원코드": 36, "투코드": 10, "셔터": 10}
CONFIDENCE_THRESHOLD = 0.8

# DI 는 일반 색상명으로 발주되기도 하므로 이름 -> 슬랫 번호를 보조 매핑한다.
# 믹스 품목은 대표 코드 하나로 축약하지 않고 DI_MIX의 정확한 조합을 사용한다.
DI_COLOR = {
    "화이트": "102", "백아이보리": "105", "아이보리": "200", "밀크로즈": "220",
    "크림베이지": "230", "바닐라": "260", "밀크티": "270", "머쉬룸": "280",
    "베이지": "290", "연핑크": "330", "옐로우": "500", "머스타드": "590",
    "딥블루": "690", "멜론": "740", "그린": "760", "딥그린": "790",
    "연그레이": "820", "그레이": "860", "스톤그레이": "870",
    "모던그레이": "880", "다크그레이": "890", "차콜그레이": "920",
    "블랙": "990", "민트": "023", "아쿠아": "026", "오렌지": "025",
    "글로시실버": "030", "카멜": "140",
    # 우측 수식 결과가 저장되지 않은 DI 파일의 예비 변환
    "펄실버": "027", "브라운": "150", "펄다크그레이": "003", "라임": "720",
}

# 2026-08-30 알루미늄 색상표. 키는 공백을 제거한 발주서 표기다.
# value = (마스터 품목명에 쓰이는 믹스명, 정확한 슬랫 조합)
DI_MIX = {
    # 포인트믹스
    "포인트믹스밀크코코아": ("밀크코코아", "200+290"),
    "포인트믹스애쉬그레이": ("애쉬그레이", "102+870"),
    "포인트믹스라즈베리": ("라즈베리", "102+380"),
    "포인트믹스아네모네": ("아네모네", "102+460"),
    "포인트믹스블랙야크": ("블랙야크", "102+990"),
    "포인트믹스딸기샤베트": ("딸기샤베트", "102+330"),
    "포인트믹스오렌지카라멜": ("오렌지카라멜", "200+025"),
    "포인트믹스봄내음": ("봄내음", "590+760"),
    "포인트믹스플라멩고": ("플라멩고", "125+990"),
    # 더블믹스
    "더블믹스달콤코코아": ("달콤코코아", "260+290"),
    "더블믹스스카이블루": ("스카이블루", "620+670"),
    "더블믹스스윗베베": ("스윗베베", "102+330"),
    "더블믹스감성그린": ("감성그린", "760+740"),
    "더블믹스사이다": ("사이다", "102+760"),
    "더블믹스레드칵테일": ("레드칵테일", "102+125"),
    "더블믹스오란씨": ("오란씨", "102+025"),
    "더블믹스모히칸": ("모히칸", "990+125"),
    "더블믹스도시남녀": ("도시남녀", "870+990"),
    # 트리플믹스
    "트리플믹스차도남": ("차도남", "870+102+990"),
    "트리플믹스차도녀": ("차도녀", "102+870+990"),
    "트리플믹스바닐라크림": ("바닐라크림", "102+200+260"),
    "트리플믹스레몬칵테일": ("레몬칵테일", "200+500+590"),
    "트리플믹스파프리카": ("파프리카", "590+025+125"),
    "트리플믹스오션블루": ("오션블루", "102+620+670"),
    "트리플믹스핑크레이디": ("핑크레이디", "102+330+380"),
    "트리플믹스베를린": ("베를린", "990+125+590"),
    "트리플믹스마쉬멜로우": ("마쉬멜로우", "620+102+330"),
    "트리플믹스아이리스": ("아이리스", "102+620+460"),
    "트리플믹스모히또": ("모히또", "102+740+760"),
    "트리플믹스오렌지블라썸": ("오렌지블라썸", "102+590+025"),
}

# 믹스 접두사가 생략된 수동 편집도 정확한 조합으로 복원할 수 있게 한다.
DI_MIX_BY_NAME = {name: (name, codes) for name, codes in DI_MIX.values()}


def _mix_code_signature(codes):
    """MIX 조합 비교용. 발주서의 코드 순서/하이픈/괄호 메모와 무관하게 비교한다."""
    parts = re.findall(r"\d{3}", str(codes or ""))
    return tuple(sorted(parts))


# 같은 코드 조합이 서로 다른 이름으로 등록된 경우(예: 102+330)는
# 코드만으로 품명을 단정하지 않고 일반 B MIX로 처리하기 위해 목록으로 보존한다.
DI_MIX_BY_CODES = {}
for _mix_name, _mix_codes in DI_MIX.values():
    DI_MIX_BY_CODES.setdefault(_mix_code_signature(_mix_codes), []).append(
        (_mix_name, _mix_codes))


def di_mix_from_codes(codes):
    """숫자 코드 조합만으로 유일하게 식별되는 공식 MIX면 (이름, 조합)을 반환."""
    candidates = DI_MIX_BY_CODES.get(_mix_code_signature(codes), [])
    return candidates[0] if len(candidates) == 1 else None


# 사용자 제공 단가표의 25mm 일반 MIX 품목. 50mm 항목은 의도적으로 제외한다.
# 정식 이름을 특정할 수 없는 2개 이상 코드 조합은 아래 품목으로 청구하고,
# 실제 조합(예: 102+520)은 장부/작업지시서/EDI 기재사항에 보존한다.
DI_GENERIC_MIX_PRODUCTS = {
    ("C자", "투코드"): {
        "품명": "B-MIX-투코드/25mm", "관리코드": "B-MIX-투코드/25mm",
        "대분류": "블라인드-C", "규격": "", "출고L가": 30000, "비고2": "",
    },
    ("C자", "원코드"): {
        "품명": "B-MIX-원코드/25mm", "관리코드": "B-MIX-원코드/25mm",
        "대분류": "블라인드-C", "규격": "", "출고L가": 35000, "비고2": "",
    },
    ("C자", "셔터"): {
        "품명": "B-MIX-셔터/25mm", "관리코드": "B-MIX-셔터/25mm",
        "대분류": "블라인드-C", "규격": "", "출고L가": 40000, "비고2": "",
    },
    ("L자", "투코드"): {
        "품명": "B-MIX-L18/투코드/25mm", "관리코드": "B-MIX-L18/투코드/25mm",
        "대분류": "블라인드-L18", "규격": "", "출고L가": 38500, "비고2": "",
    },
    ("L자", "원코드"): {
        "품명": "B-MIX-L18/원코드/25mm", "관리코드": "B-MIX-L18/원코드/25mm",
        "대분류": "블라인드-L18", "규격": "", "출고L가": 44000, "비고2": "",
    },
    ("L18", "투코드"): {
        "품명": "B-MIX-L18/투코드/25mm", "관리코드": "B-MIX-L18/투코드/25mm",
        "대분류": "블라인드-L18", "규격": "", "출고L가": 38500, "비고2": "",
    },
    ("L18", "원코드"): {
        "품명": "B-MIX-L18/원코드/25mm", "관리코드": "B-MIX-L18/원코드/25mm",
        "대분류": "블라인드-L18", "규격": "", "출고L가": 44000, "비고2": "",
    },
    ("L18", "셔터"): {
        "품명": "B-MIX-L18/셔터/25mm",
        # 사용자 제공 경영박사 화면의 관리코드를 그대로 따른다.
        "관리코드": "B-MIX-L18/원코드/25mm",
        "대분류": "블라인드-L18", "규격": "", "출고L가": 49500, "비고2": "",
    },
    ("L21", "원코드"): {
        "품명": "B-MIX-L21/원코드/25mm", "관리코드": "B-MIX-L21/원코드/25mm",
        "대분류": "블라인드-L21", "규격": "", "출고L가": 55000, "비고2": "",
    },
}


# 2026-08-30 경영박사 「대일산업 믹스 리스트.xls」 실제 등록 품목.
# key = (믹스명, 전체 슬랫 조합, 종류, 타입)
# 관리코드는 현재 경영박사에서 품명과 동일하다.
DI_MIX_PRODUCTS = {
    ('감성그린', '760+740', '투코드', 'C자'): {'품명': 'B감성그린(760+740)', '관리코드': 'B감성그린(760+740)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('감성그린', '760+740', '원코드', 'C자'): {'품명': 'B감성그린-원코드(760+740)', '관리코드': 'B감성그린-원코드(760+740)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('달콤코코아', '260+290', '투코드', 'C자'): {'품명': 'B달콤코코아(260+290)', '관리코드': 'B달콤코코아(260+290)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('달콤코코아', '260+290', '원코드', 'C자'): {'품명': 'B달콤코코아-원코드(260+290)', '관리코드': 'B달콤코코아-원코드(260+290)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('도시남녀', '870+990', '투코드', 'C자'): {'품명': 'B도시남녀(870+990)', '관리코드': 'B도시남녀(870+990)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('도시남녀', '870+990', '원코드', 'C자'): {'품명': 'B도시남녀-원코드(870+990)', '관리코드': 'B도시남녀-원코드(870+990)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('딸기샤베트', '102+330', '투코드', 'C자'): {'품명': 'B딸기샤베트(102+330)', '관리코드': 'B딸기샤베트(102+330)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('딸기샤베트', '102+330', '원코드', 'C자'): {'품명': 'B딸기샤베트-원코드(102+330)', '관리코드': 'B딸기샤베트-원코드(102+330)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('라즈베리', '102+380', '투코드', 'C자'): {'품명': 'B라즈베리(102+380)', '관리코드': 'B라즈베리(102+380)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('라즈베리', '102+380', '원코드', 'C자'): {'품명': 'B라즈베리-원코드(102+380)', '관리코드': 'B라즈베리-원코드(102+380)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('레드칵테일', '102+125', '투코드', 'C자'): {'품명': 'B레드칵테일(102+125)', '관리코드': 'B레드칵테일(102+125)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('레드칵테일', '102+125', '원코드', 'C자'): {'품명': 'B레드칵테일-원코드(102+125)', '관리코드': 'B레드칵테일-원코드(102+125)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('레몬칵테일', '200+500+590', '투코드', 'C자'): {'품명': 'B레몬칵테일(200+500+590)', '관리코드': 'B레몬칵테일(200+500+590)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('레몬칵테일', '200+500+590', '원코드', 'C자'): {'품명': 'B레몬칵테일-원코드(200+500+590)', '관리코드': 'B레몬칵테일-원코드(200+500+590)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('마쉬멜로우', '620+102+330', '투코드', 'C자'): {'품명': 'B마쉬멜로우(620+102+330)', '관리코드': 'B마쉬멜로우(620+102+330)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('마쉬멜로우', '620+102+330', '원코드', 'C자'): {'품명': 'B마쉬멜로우-원코드(620+102+330)', '관리코드': 'B마쉬멜로우-원코드(620+102+330)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('모히또', '102+740+760', '투코드', 'C자'): {'품명': 'B모히또(102+740+760)', '관리코드': 'B모히또(102+740+760)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('모히또', '102+740+760', '원코드', 'C자'): {'품명': 'B모히또-원코드(102+740+760)', '관리코드': 'B모히또-원코드(102+740+760)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('모히칸', '990+125', '투코드', 'C자'): {'품명': 'B모히칸(990+125)', '관리코드': 'B모히칸(990+125)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('모히칸', '990+125', '원코드', 'C자'): {'품명': 'B모히칸-원코드(990+125)', '관리코드': 'B모히칸-원코드(990+125)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('밀크코코아', '200+290', '투코드', 'C자'): {'품명': 'B밀크코코아(200+290)', '관리코드': 'B밀크코코아(200+290)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('밀크코코아', '200+290', '원코드', 'C자'): {'품명': 'B밀크코코아-원코드(200+290)', '관리코드': 'B밀크코코아-원코드(200+290)', '대분류': '블라인드-C/DI', '규격': '추가비용청구', '출고L가': 17600, '비고2': ''},
    ('바닐라크림', '102+200+260', '투코드', 'C자'): {'품명': 'B바닐라크림(102+200+260)', '관리코드': 'B바닐라크림(102+200+260)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('바닐라크림', '102+200+260', '원코드', 'C자'): {'품명': 'B바닐라크림-원코드(102+200+260)', '관리코드': 'B바닐라크림-원코드(102+200+260)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('베를린', '990+125+590', '투코드', 'C자'): {'품명': 'B베를린(990+125+590)', '관리코드': 'B베를린(990+125+590)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('베를린', '990+125+590', '원코드', 'C자'): {'품명': 'B베를린-원코드(990+125+590)', '관리코드': 'B베를린-원코드(990+125+590)', '대분류': '블라인드-C/DI', '규격': '추가비용청구★', '출고L가': 17600, '비고2': ''},
    ('봄내음', '590+760', '투코드', 'C자'): {'품명': 'B봄내음(590+760)', '관리코드': 'B봄내음(590+760)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('봄내음', '590+760', '원코드', 'C자'): {'품명': 'B봄내음-원코드(590+760)', '관리코드': 'B봄내음-원코드(590+760)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('블랙야크', '102+990', '투코드', 'C자'): {'품명': 'B블랙야크(102+990)', '관리코드': 'B블랙야크(102+990)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('블랙야크', '102+990', '원코드', 'C자'): {'품명': 'B블랙야크-원코드(102+990)', '관리코드': 'B블랙야크-원코드(102+990)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('사이다', '102+760', '투코드', 'C자'): {'품명': 'B사이다(102+760)', '관리코드': 'B사이다(102+760)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('사이다', '102+760', '원코드', 'C자'): {'품명': 'B사이다-원코드(102+760)', '관리코드': 'B사이다-원코드(102+760)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('스윗베베', '102+330', '투코드', 'C자'): {'품명': 'B스윗베베(102+330)', '관리코드': 'B스윗베베(102+330)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('스윗베베', '102+330', '원코드', 'C자'): {'품명': 'B스윗베베-원코드(102+330)', '관리코드': 'B스윗베베-원코드(102+330)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('스카이블루', '620+670', '투코드', 'C자'): {'품명': 'B스카이블루(620+670)', '관리코드': 'B스카이블루(620+670)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('스카이블루', '620+670', '원코드', 'C자'): {'품명': 'B스카이블루-원코드(620+670)', '관리코드': 'B스카이블루-원코드(620+670)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('아네모네', '102+460', '투코드', 'C자'): {'품명': 'B아네모네(102+460)', '관리코드': 'B아네모네(102+460)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('아네모네', '102+460', '원코드', 'C자'): {'품명': 'B아네모네-원코드(102+460)', '관리코드': 'B아네모네-원코드(102+460)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('아이리스', '102+620+460', '투코드', 'C자'): {'품명': 'B아이리스(102+620+460)', '관리코드': 'B아이리스(102+620+460)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('아이리스', '102+620+460', '원코드', 'C자'): {'품명': 'B아이리스-원코드(102+620+460)', '관리코드': 'B아이리스-원코드(102+620+460)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('애쉬그레이', '102+870', '투코드', 'C자'): {'품명': 'B애쉬그레이(102+870)', '관리코드': 'B애쉬그레이(102+870)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('애쉬그레이', '102+870', '원코드', 'C자'): {'품명': 'B애쉬그레이-원코드(102+870)', '관리코드': 'B애쉬그레이-원코드(102+870)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('오란씨', '102+025', '투코드', 'C자'): {'품명': 'B오란씨(102+025)', '관리코드': 'B오란씨(102+025)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('오란씨', '102+025', '원코드', 'C자'): {'품명': 'B오란씨-원코드(102+025)', '관리코드': 'B오란씨-원코드(102+025)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('오렌지블라썸', '102+590+025', '투코드', 'C자'): {'품명': 'B오렌지블라썸(102+590+025)', '관리코드': 'B오렌지블라썸(102+590+025)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('오렌지블라썸', '102+590+025', '원코드', 'C자'): {'품명': 'B오렌지블라썸-원코드(102+590+025)', '관리코드': 'B오렌지블라썸-원코드(102+590+025)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('오렌지카라멜', '200+025', '투코드', 'C자'): {'품명': 'B오렌지카라멜(200+025)', '관리코드': 'B오렌지카라멜(200+025)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('오렌지카라멜', '200+025', '원코드', 'C자'): {'품명': 'B오렌지카라멜-원코드(200+025)', '관리코드': 'B오렌지카라멜-원코드(200+025)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('오션블루', '102+620+670', '투코드', 'C자'): {'품명': 'B오션블루(102+620+670)', '관리코드': 'B오션블루(102+620+670)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('오션블루', '102+620+670', '원코드', 'C자'): {'품명': 'B오션블루-원코드(102+620+670)', '관리코드': 'B오션블루-원코드(102+620+670)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('차도남', '870+102+990', '투코드', 'C자'): {'품명': 'B차도남(870+102+990)', '관리코드': 'B차도남(870+102+990)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('차도남', '870+102+990', '원코드', 'C자'): {'품명': 'B차도남-원코드(870+102+990)', '관리코드': 'B차도남-원코드(870+102+990)', '대분류': '블라인드-C/DI', '규격': '추가비용청구', '출고L가': 17600, '비고2': ''},
    ('차도녀', '102+870+990', '투코드', 'C자'): {'품명': 'B차도녀(102+870+990)', '관리코드': 'B차도녀(102+870+990)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('차도녀', '102+870+990', '원코드', 'C자'): {'품명': 'B차도녀-원코드(102+870+990)', '관리코드': 'B차도녀-원코드(102+870+990)', '대분류': '블라인드-C/DI', '규격': '추가비용청구', '출고L가': 17600, '비고2': ''},
    ('파프리카', '590+025+125', '투코드', 'C자'): {'품명': 'B파프리카(590+025+125)', '관리코드': 'B파프리카(590+025+125)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('파프리카', '590+025+125', '원코드', 'C자'): {'품명': 'B파프리카-원코드(590+025+125)', '관리코드': 'B파프리카-원코드(590+025+125)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('플라멩고', '125+990', '투코드', 'C자'): {'품명': 'B플라멩고(125+990)', '관리코드': 'B플라멩고(125+990)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('플라멩고', '125+990', '원코드', 'C자'): {'품명': 'B플라멩고-원코드(125+990)', '관리코드': 'B플라멩고-원코드(125+990)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 17600, '비고2': ''},
    ('핑크레이디', '102+330+380', '투코드', 'C자'): {'품명': 'B핑크레이디(102+330+380)', '관리코드': 'B핑크레이디(102+330+380)', '대분류': '블라인드-C/DI', '규격': '', '출고L가': 12650, '비고2': ''},
    ('핑크레이디', '102+330+380', '원코드', 'C자'): {'품명': 'B핑크레이디-원코드(102+330+380)', '관리코드': 'B핑크레이디-원코드(102+330+380)', '대분류': '블라인드-C/DI', '규격': '추가비용청구★', '출고L가': 17600, '비고2': ''},
}

DI_COLOR_UNSURE = set()


def normalize_di_color_name(value):
    """DI 색상명을 비교용으로 정규화(공백·하이픈·언더바 제거)."""
    return re.sub(r"[\s_-]+", "", str(value or "").strip())


def di_mix_info(value):
    """발주/수동 입력 색상명 -> (믹스명, '102+870') 또는 None."""
    key = normalize_di_color_name(value)
    return DI_MIX.get(key) or DI_MIX_BY_NAME.get(key)

# 대일산업 25mm 알루미늄 블라인드 품목표의 '순서'.
# 숫자 크기가 아니라 공장에서 사용하는 이 표의 순서로 DI를 정렬한다.
DI_CODE_SEQUENCE = (
    "102", "105", "200", "220", "230", "260", "270", "280", "290",
    "310", "330", "380", "410", "460", "500", "590", "620", "670",
    "680", "690", "720", "740", "760", "790", "023", "026", "025",
    "125", "140", "150", "160", "180", "820", "860", "870", "880",
    "890", "920", "990", "027", "028", "001", "002", "003", "004",
    "030", "031", "032",
    "102P", "200P", "260P", "330P", "380P", "410P", "460P", "500P",
    "590P", "620P", "670P", "720P", "740P", "760P", "023P", "990P",
    "027P", "028P",
    "010", "011", "012", "013",
    "104FP", "240FP", "220FP", "870FP", "990FP", "003FP", "028FP",
    "029FP", "050", "051", "052", "053", "054", "055",
    "4453", "4454", "4455", "4456",
)
DI_CODE_ORDER = {code: index for index, code in enumerate(DI_CODE_SEQUENCE)}


# ─────────────────────────────────────────────
# 마스터
# ─────────────────────────────────────────────
class Master:
    def __init__(self, path=None):
        if path is None:
            path = resolve_master_file("blind", ("fitorder_master.xlsx", "fitorder_master_출고B가추가.xlsx"))
        path = Path(path)
        if not path.is_file():
            # 설치본의 data/master와 SRC 폴백을 공통 방식으로 다시 탐색한다.
            path = resolve_master_file("blind", (path.name, "fitorder_master.xlsx", "fitorder_master_출고B가추가.xlsx"))
        self.path = path
        raw_rows = read_rows(path, required_headers=("품명", "대분류"), preferred_sheet="품목")
        self.tier_col = {}
        self.items = {}          # 품명 -> dict
        self.by_code = {}        # 코드 -> 기본 품명 (접미사 없는 것)
        self.mix_items = {}      # (믹스명, 조합, 종류, 타입) -> 품명
        for raw in raw_rows:
            name = raw.get("품명")
            if not name:
                continue
            item = {
                "품명": name,
                "대분류": raw.get("대분류") or "",
                "규격": raw.get("규격") or "",
                "비고2": raw.get("비고2") or "",
            }
            for key, value in raw.items():
                if str(key).startswith("출고"):
                    item[key] = value or 0
                    self.tier_col[key] = key
            self.items[name] = item
            # 일반 슬랫은 3자리, 우드(WD4453~4456)는 4자리다.
            m = re.match(r"^B([0-9]{3,4}[A-Za-z]*)\(([^)]*)\)$", str(name))
            if m:
                code, inner = m.group(1), m.group(2)
                if code in self.by_code and ("/" in inner or "無" in inner):
                    continue
                self.by_code[code] = name
            mm = re.match(r"^B(.+?)(-원코드)?\((\d{3}(?:\+\d{3})+)\)$", str(name))
            if mm:
                mix_name = mm.group(1)
                kind = "원코드" if mm.group(2) else "투코드"
                self.mix_items[(mix_name, mm.group(3), kind, "C자")] = name

    # 손잡이실 색상 (DI 정렬용)
    def thread_color(self, code):
        # 믹스는 첫 번째 실제 슬랫의 실색을 정렬 기준으로만 사용한다.
        code = str(code or "").split("+")[0]
        base = self.by_code.get(code)
        if not base:
            return None
        m = re.search(r"실\s*:?\s*([가-힣A-Za-z]+)", self.items[base]["비고2"])
        return m.group(1) if m else None

    def build_name(self, code, kind, type_, client):
        """품명 조합: B200(IV)-L18/원코드 형태"""
        base = self.by_code.get(code)
        if not base:
            return None
        if type_ == "L자":                       # L자는 DI도 동일 단가
            return f"{base}-L18/{kind}"
        # DI의 C자 원/투코드만 DI 전용 품명을 쓴다.
        # 셔터와 L자는 다른 업체와 같은 공통 품명을 사용한다.
        if client == "DI" and kind != "셔터":
            return base + ("-DI" if kind == "투코드" else f"-{kind}/DI")
        return base if kind == "투코드" else f"{base}-{kind}"

    def _priced_item(self, name, client):
        if not name or name not in self.items:
            return None
        it = dict(self.items[name])
        it["관리코드"] = name
        tier = CLIENT_INFO.get(client, (None, "출고I가"))[1]
        it["단가"] = it.get(tier, 0)
        it["단가등급"] = tier
        return it

    def find_item(self, code, kind, type_, client):
        name = self.build_name(code, kind, type_, client)
        return self._priced_item(name, client)

    def find_mix_item(self, mix_name, mix_codes, kind, type_, client, variant=None):
        """색상표의 정확한 믹스 조합을 찾는다.

        DI C자 믹스는 2026-08-30 경영박사 실제 등록 목록을 최우선으로 사용한다.
        따라서 로컬 마스터가 구버전이라 원코드 MIX 행이 빠져 있어도
        품명/관리코드/출고L가를 정확히 복원한다. 다른 거래처나 L자는 기존
        마스터만 사용하며 임의 조합을 생성하지 않는다.
        """
        key = (str(mix_name or "").strip(), str(mix_codes or "").replace(" ", ""),
               kind or "투코드", type_ or "C자")
        if client == "DI":
            official = DI_MIX_PRODUCTS.get(key)
            if official:
                it = dict(official)
                it["단가"] = it["출고L가"]
                it["단가등급"] = "출고L가"
                return it
            # 코드가 2개 이상이지만 공식 이름 조합표에 없거나 코드만으로
            # 이름을 하나로 특정할 수 없는 경우에도 MIX 자체는 정상 품목이다.
            # L자/셔터는 제공된 일반 MIX 단가표의 범위가 아니므로 임의 적용하지 않는다.
            generic_key = (variant or type_ or "C자", kind or "투코드")
            if str(mix_name or "").upper() == "MIX" \
                    and generic_key in DI_GENERIC_MIX_PRODUCTS:
                # 사용자 제공 MIX 단가표의 25mm 품목만 사용한다. 50mm는 제외.
                it = dict(DI_GENERIC_MIX_PRODUCTS[generic_key])
                it["단가"] = it["출고L가"]
                it["단가등급"] = "출고L가"
                return it
        if str(mix_name or "").upper() == "MIX" and client != "DI":
            # 비대일 블라인드 MIX는 경영박사 일반 MIX 품목(B-MIX-…/25mm)으로 청구한다.
            # MIX 품목 단가는 품목장 값(사람이 직접 입력)을 쓰고, 믹스 추가비용은 별도 행으로 붙는다.
            it = self._priced_item(generic_mix_product_name(kind, type_, variant), client)
            if it and it["품명"] == "B-MIX-L18/셔터/25mm":
                # 사용자 제공 경영박사 화면의 관리코드(DI_GENERIC_MIX_PRODUCTS와 동일).
                it["관리코드"] = "B-MIX-L18/원코드/25mm"
            return it
        return self._priced_item(self.mix_items.get(key), client)

    def find_holding_item(self, item, client):
        """홀딩도어/부속은 사용자 제공 공식 품목표를 소스 스냅샷으로 조회한다."""
        hit = None
        if item.get("_holding_accessory"):
            hit = holding_accessory_from_text(
                item.get("_holding_product_name") or item.get("색상원문") or item.get("품목코드"),
                item.get("_holding_accessory_color") or "화이트")
        else:
            hit = holding_product_by_item(item)
        if not hit:
            return None
        out = dict(hit)
        tier = CLIENT_INFO.get(client, (None, "출고I가"))[1]
        out["단가"] = out.get(tier, 0) or 0
        out["단가등급"] = tier
        return out

    def find_holding_extra(self, name, client):
        row = holding_extra_product(name)
        if not row:
            return None
        out = dict(row)
        tier = CLIENT_INFO.get(client, (None, "출고I가"))[1]
        out["단가"] = out.get(tier, 0) or 0
        out["단가등급"] = tier
        return out

    def find_named_item(self, name, client):
        """마스터의 정확한 품명으로 거래처 단가를 조회한다."""
        return self._priced_item(name, client)

    def find_order_item(self, item, client):
        if item.get("_product_group") == "holding" or item.get("_holding_accessory"):
            return self.find_holding_item(item, client)
        mix_name = item.get("_mix_name")
        mix_codes = item.get("_mix_codes")
        if mix_name and mix_codes:
            return self.find_mix_item(mix_name, mix_codes, item.get("종류"),
                                      item.get("타입"), client,
                                      item.get("_di_mix_family"))
        return self.find_item(item.get("품목코드"), item.get("종류"),
                              item.get("타입"), client)


# ─────────────────────────────────────────────
# 계산
# ─────────────────────────────────────────────
def calc_erp(width_cm, height_cm, unit_price):
    """헤베·금액 계산.

    실제 세로가 150cm 이하이면 계산 세로만 150cm로 적용한다.
    장부·전표적요의 실제 규격은 바꾸지 않는다. 헤베는 둘째 자리,
    최소 수량은 1.5, 금액·부가세는 원 단위로 반올림한다.
    """
    actual_height = Decimal(str(height_cm))
    billing_height = max(actual_height, Decimal("150"))
    raw = (Decimal(str(width_cm)) / 100) * (billing_height / 100)
    hebe = raw.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    qty = max(hebe, Decimal("1.5"))
    amount = (Decimal(str(unit_price)) * qty).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP)
    vat = (amount / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return {"헤베": hebe, "수량": qty, "계산세로": billing_height,
            "단가": unit_price,
            "금액": amount, "부가세": vat}


def calc_roll_erp(width_cm, height_cm, unit_price):
    """롤/콤비/쉐이드 경영박사 청구 계산.

    실제 세로가 150cm 미만이면 계산 세로만 150cm로 올리고,
    계산 면적이 2.0㎡ 미만이면 청구 수량을 2.0㎡로 적용한다.
    장부와 전표적요에는 주문서의 실제 가로/세로를 그대로 남긴다.
    """
    actual_height = Decimal(str(height_cm))
    billing_height = max(actual_height, Decimal("150"))
    raw = (Decimal(str(width_cm)) / 100) * (billing_height / 100)
    hebe = raw.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    qty = max(hebe, Decimal("2.0"))
    amount = (Decimal(str(unit_price)) * qty).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP)
    vat = (amount / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return {"헤베": hebe, "수량": qty, "계산세로": billing_height,
            "단가": unit_price, "금액": amount, "부가세": vat}


def default_handle_length(kind, height):
    if height is None:
        return None
    if kind == "원코드":
        return 150 if height >= 210 else 130
    if kind == "투코드":
        return 150 if height >= 210 else 100
    if kind == "셔터":
        if height >= 210:
            return 150
        if height >= 100:
            return 100
        return int(height // 10) * 10 - 10
    return None


def normalize_handle(length):
    """10단위 내림. 보정 여부 반환"""
    if not length:
        return None, False
    r = (int(length) // 10) * 10
    return r, r != int(length)


def normalize_roll_handle(length):
    """롤/콤비 손잡이길이: 5단위 내림(2026-09-14 사용자 확정, 127→125). 보정 여부 반환"""
    try:
        n = int(float(length)) if length not in (None, "") else None
    except (TypeError, ValueError):
        return None, False
    if not n:
        return None, False
    r = (n // 5) * 5
    return r, r != n


def ledger_color(code, kind, type_, prefix="B", mix_name=None, mix_codes=None):
    """장부 색상 열 문자열. 믹스는 이름 + 전체 코드 조합을 보존한다."""
    if not code and not (mix_name and mix_codes):
        return None
    prefix = str(prefix or "B").strip().upper()
    if prefix not in {"B", "H", "R", "C"}:
        prefix = "B"
    if mix_name and mix_codes:
        if str(mix_name).upper() == "MIX":
            # 조합표에 없는 MIX는 타입/조합을 품목 칸에 풀어 쓰지 않는다.
            # 사용자가 지정한 표기대로 B MIX / B 원코드 MIX만 표시하고
            # 실제 코드 조합은 기재사항에 남긴다.
            text = f"{prefix} 원코드 MIX" if kind == "원코드" else f"{prefix} MIX"
        else:
            mix = f"{mix_name} ({str(mix_codes).replace(' ', '')})"
            if type_ == "L자":
                text = f"{prefix} L18-{kind} {mix}"
            else:
                text = f"{prefix} {mix}" if kind == "투코드" else f"{prefix} {kind} {mix}"
    elif type_ == "L자":
        text = f"{prefix} L18-{kind} {code}"
    else:
        text = f"{prefix} {code}" if kind == "투코드" else f"{prefix} {kind} {code}"
    # 장부 품목 칸에서만 맨 앞에 여백 한 칸을 둔다.
    return " " + text



# ─────────────────────────────────────────────
# 검증
# ─────────────────────────────────────────────
def validate(order, M):
    """order = extract_order() 또는 파서 결과. [(등급, 코드, 메시지, 행번호)]"""
    out = []
    client = order.get("거래처")
    if client not in CLIENT_INFO:
        out.append(("red", "거래처불명", "거래처를 판별하지 못했습니다", None))

    items = order.get("items") or []
    if not items:
        out.append(("red", "항목없음", "추출된 주문 항목이 없습니다", None))

    if client == "JO" and not str(order.get("주문번호") or "").strip():
        out.append(("red", "JO주문번호누락",
                    "제이원 발주서의 상단 주문번호를 읽지 못했습니다. 원본의 주문번호를 확인해 주세요.", None))
    if client == "DU" and not str(order.get("주문번호") or "").strip():
        out.append(("yellow", "DU주문번호누락",
                    "두창 발주서의 주문번호를 읽지 못했습니다. 장부에는 (주문번호) 형식이 필요하므로 원본을 확인해 주세요.", None))

    # 택배/화물 주소는 사전점검에서 누락을 즉시 알린다.
    delivery = order.get("배송") or {}
    mode = str(delivery.get("방식") or CLIENT_INFO.get(client, (None, None, ""))[2] or "")
    if "화물" in mode:
        has_destination = bool(str(delivery.get("화물지점") or delivery.get("주소") or "").strip())
    else:
        has_destination = bool(str(delivery.get("주소") or "").strip())
    # 배송 목적지는 '기본 거래처 배송지와 다른 곳'으로 보낼 때만 별도 입력한다.
    # 거래처의 등록 주소1/기본 화물지점에서 이미 택배·화물 방식이 결정되는 경우에는
    # 별도 목적지가 비어 있어도 등록 기본배송지를 사용하므로 누락 경고를 띄우지 않는다.
    default_mode = str(CLIENT_INFO.get(client, (None, None, ""))[2] or "")
    uses_registered_destination = (
        not has_destination
        and mode in {"택배", "화물"}
        and mode == default_mode
    )
    if ("택배" in mode or "화물" in mode) and not has_destination and not uses_registered_destination:
        out.append(("yellow", "배송주소누락",
                    "기본 거래처 배송지가 아닌 택배/화물 주문인데 배송 목적지가 없습니다. "
                    "사전점검의 별도 주소/화물지점을 확인해 주세요.", None))
    # 선불/착불 미기재는 공통 규칙상 착불로 자동 처리하므로 확인사항으로 띄우지 않는다.
    if client == "RT" and mode in {"", "판별필요"}:
        out.append(("yellow", "배송방식확인",
                    "루임트 배송방식이 확정되지 않았습니다. 주소행의 특이 칸에서 택배/화물을 확인해 주세요.", None))

    # 표형 발주의 수량과 좌·우 숫자는 거래처와 무관하게 교차검증한다.
    # 서로 다르면 프로그램이 남는 창의 방향을 추측하지 않고 반드시 확인시킨다.
    if items:
        total_declared = total_lr = 0
        comparable_rows = 0
        for idx, item in enumerate(items, 1):
            try:
                count = int(float(item.get("창개수") or 0))
                left_n = int(float(item.get("좌개수") or 0))
                right_n = int(float(item.get("우개수") or 0))
            except (TypeError, ValueError):
                continue
            lr = left_n + right_n
            if count > 0 and lr > 0:
                comparable_rows += 1
                total_declared += count
                total_lr += lr
                if count != lr:
                    out.append(("red", "좌우수량불일치",
                                f"표 {idx}행: 수량 {count}인데 좌 {left_n} + 우 {right_n} = {lr}입니다. 원본 표를 확인해 주세요.", idx))
        if comparable_rows >= 2 and total_declared != total_lr:
            out.append(("red", "표전체수량불일치",
                        f"표 전체 수량 합계 {total_declared}와 좌·우 합계 {total_lr}가 다릅니다.", None))

    # 발주서 메모의 총 창 수와 대조. 표 한 행에 수량 2처럼 여러 창이
    # 들어갈 수 있으므로 항목 행 개수가 아니라 실제 창개수 합계를 사용한다.
    memo = str(order.get("전체원문") or "") + str(order.get("전체기재사항") or "")
    m = re.search(r"총\s*(\d+)\s*창", memo)
    if m:
        actual_windows = 0
        for it in items:
            try:
                n = int(float(it.get("창개수") or it.get("수량") or 1))
            except (TypeError, ValueError):
                n = 1
            actual_windows += max(1, n)
        if actual_windows != int(m.group(1)):
            out.append(("red", "창개수불일치",
                        f"발주서 메모는 {m.group(1)}창인데 실제 창수 합계가 {actual_windows}창으로 추출되었습니다", None))

    # 인천)트루 추가부속의 경영박사 단가/품목 존재 여부를 사전점검한다.
    if client == "인천)트루":
        for acc in order.get("_true_accessories") or []:
            edi_name = str(acc.get("품명") or "").strip()
            if not edi_name:
                continue
            hit = M.find_named_item(edi_name, client)
            if not hit:
                out.append(("yellow", "트루부속단가확인",
                            f"{edi_name}: 현재 마스터에 품목/단가가 없습니다. 경영박사에는 수량 1, 단가 0으로 출력되므로 단가를 확인해 주세요.", None))
            elif not hit.get("단가"):
                out.append(("yellow", "트루부속단가확인",
                            f"{edi_name}: {hit.get('단가등급')} 단가가 0입니다. 단가를 확인해 주세요.", None))

    for i, it in enumerate(items, 1):
        w, h = it.get("가로"), it.get("세로")
        kind, type_ = it.get("종류"), it.get("타입")
        code = it.get("품목코드")
        is_holding = it.get("_product_group") == "holding" or it.get("_holding_accessory")

        if it.get("예외품목"):
            out.append(("info", "예외품목",
                        f"{it['예외품목']} — 장부 제외, 담당자 확인 필요", i))
            continue

        if not it.get("_holding_accessory"):
            if w in (None, ""):
                out.append(("red", "가로누락", "가로 치수가 없습니다", i))
            if h in (None, ""):
                out.append(("red", "세로누락", "세로 치수가 없습니다", i))
        if w and h:
            # 가로·세로의 대소관계만으로는 순서 오독을 판단하지 않는다.
            # 홀딩도어는 세로 270cm부터 제작 확인을 주고, 일반 블라인드는
            # 기존 대형치수 기준(가로 250 / 세로 400)을 유지한다.
            if is_holding and not it.get("_holding_accessory"):
                if float(h) >= 270:
                    out.append(("yellow", "홀딩세로길이확인",
                                f"홀딩도어 세로 {h}cm — 270cm 이상입니다. 원본/제작 가능 여부를 확인해 주세요.", i))
            elif w >= 250 or h >= 400:
                out.append(("yellow", "대형치수확인",
                            f"대형 치수 {w}x{h} — "
                            + ("가로 250cm 이상" if w >= 250 else "")
                            + (" / " if w >= 250 and h >= 400 else "")
                            + ("세로 400cm 이상" if h >= 400 else ""), i))
            lo = None if is_holding else MIN_WIDTH.get(kind)
            if lo and w < lo:
                out.append(("red", "최소사이즈미달",
                            f"{kind} 최소 가로 {lo}cm 미만 ({w}cm)", i))

        if not code:
            out.append(("red", "품목미등록", "색상 코드를 읽지 못했습니다", i))
        else:
            hit = M.find_order_item(it, client)
            if not hit:
                if it.get("_mix_name"):
                    desc = f"{it.get('_mix_name')}({it.get('_mix_codes')})/{kind}/{type_}"
                else:
                    desc = f"{code}/{kind}/{type_}"
                out.append(("red", "품목미등록",
                            f"마스터에 없는 조합: {desc}", i))
            elif not hit["단가"]:
                # DI 믹스는 색상표/마스터 품명까지는 정확히 복원했지만
                # 출고L가가 비어 있는 경우가 많다. 임의 단가를 만들지 않고
                # 사용자가 확인할 수 있도록 '확인' 상태로 남긴다.
                if client == "DI" and it.get("_mix_name"):
                    out.append(("yellow", "DI믹스단가확인",
                                f"{hit['품명']} — {hit['단가등급']} 단가가 0입니다. "
                                "품명은 복원했으며 단가만 확인해 주세요", i))
                else:
                    out.append(("red", "단가없음",
                                f"{hit['품명']} — {hit['단가등급']} 단가가 0입니다", i))

        if is_holding and not it.get("_holding_accessory"):
            if it.get("_holding_upper_roller"):
                ex = M.find_holding_extra("H상하로라(가로m당)", client)
                if not ex or not ex.get("단가"):
                    out.append(("yellow", "홀딩상하로라단가확인",
                                f"상하로라 — {CLIENT_INFO.get(client, ('','출고I가'))[1]} 단가를 확인해 주세요", i))
            magnet_names = holding_magnet_extra_names(
                it.get("_holding_operation") or it.get("수량"), h)
            tier = CLIENT_INFO.get(client, ('','출고I가'))[1]
            for magnet_name in magnet_names:
                ex = M.find_holding_extra(magnet_name, client)
                if not ex or not ex.get("단가"):
                    out.append(("yellow", "홀딩자석바단가확인",
                                f"{magnet_name} — {tier} 단가를 확인해 주세요", i))

        if it.get("_blind_mix"):
            out.append(("yellow", "MIX확인",
                        f"MIX 자동 처리: {mix_memo_text(it)} — 섞인 코드와 길이를 확인해 주세요", i))
        elif it.get("_mix_unresolved"):
            out.append(("yellow", "MIX조합확인",
                        "MIX 표시가 있지만 섞인 코드 조합(예: 102+500)을 찾지 못했습니다. 품목을 확인해 주세요", i))

        handle_length = it.get("손잡이길이")
        if not is_holding and h and handle_length is not None \
                and abs(float(h) - float(handle_length)) >= 100:
            out.append(("yellow", "손잡이길이차이",
                        f"세로 {h}cm / 손잡이 {handle_length}cm — "
                        "100cm 이상 차이", i))
        is_roll = (it.get("_product_group") == "roll_combo"
                   or order.get("_product_mode") == "roll_combo")
        if is_holding:
            adj = False
        elif is_roll:
            _, adj = normalize_roll_handle(it.get("손잡이길이"))
        else:
            _, adj = normalize_handle(it.get("손잡이길이"))
        if adj:
            out.append(("yellow", "손잡이길이보정",
                        f"{it['손잡이길이']} → {'5' if is_roll else '10'}단위 내림", i))

        conf = it.get("확신도") or {}
        # 손잡이는 기본값 보정이 가능하고 손글씨에서 확신도가 자주 낮게
        # 나오므로 확인을 남발하지 않는다. 제작에 직접 영향을 주는
        # 품목·가로·세로만 낮은 확신도 경고 대상으로 삼는다.
        low = [k for k, v in conf.items()
               if k in {"가로", "세로", "품목코드"}
               and isinstance(v, (int, float))
               and v < CONFIDENCE_THRESHOLD]
        if low:
            out.append(("yellow", "확신도낮음",
                        "판독 확신도 낮음: " + ", ".join(low), i))

    if order.get("변경요청"):
        txt = (order.get("변경문구") or "").strip()
        out.append(("review", "변경요청",
                    "변경 요청으로 보입니다"
                    + (f" — \"{txt[:60]}\"" if txt else "")
                    + " · 대상 주문을 확인하세요", None))

    if client == "RT" and not (order.get("배송") or {}).get("방식"):
        out.append(("red", "배송판별불가", "RT — 택배/화물 구분 불가", None))

    return out


def worst(issues):
    for lv in ("red", "yellow", "info"):
        if any(i[0] == lv for i in issues):
            return lv
    return "ok"


# ─────────────────────────────────────────────
# 중복 감지 / 주문 변경 탐지
# ─────────────────────────────────────────────
DUP_LOOKBACK_DAYS = 7
CHANGE_LOOKBACK_DAYS = 2
SIZE_TOL = 0.0          # 중복은 치수가 완전히 같을 때만 (±1cm 의심 제거)
CHANGE_TOL = 1.0        # 변경 후보 판정에만 허용 오차를 둔다


def item_key(client, it):
    """중복 판정 키"""
    return (client, it.get("품목코드"), it.get("_mix_name"), it.get("_mix_codes"),
            it.get("종류"), it.get("타입"), it.get("가로"), it.get("세로"), it.get("수량"),
            it.get("손잡이방향"), it.get("손잡이길이"))


def _near(a, b, tol=SIZE_TOL):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return a == b


def find_duplicates(orders, past=None):
    """과거 주문과의 중복만 찾는다.

       같은 발주서/현재 분석 묶음 안의 동일 창은 정상적인 반복 주문으로
       간주한다. 중복 표시는 DB에 저장된 과거 처리 주문과 완전히 같은
       경우에만 만든다.
       past = [(라벨, 거래처, item), ...]
       반환: [(등급, 코드, 메시지, 전역행번호)]"""
    out, rows = [], []
    for oi, o in enumerate(orders):
        c = o.get("거래처")
        for it in o.get("items", []):
            if it.get("예외품목"):
                continue
            rows.append((oi, c, it))

    for n, (oi, c, it) in enumerate(rows, 1):
        k = item_key(c, it)
        for label, pc, pit in (past or []):
            if item_key(pc, pit) == k:
                out.append(("review", "과거중복",
                            f"{label} 주문과 동일합니다", n))
                break
    return out


def change_candidate(client_a, a, client_b, b):
    """변경 대상 후보 판정 — 한쪽 치수만 바뀌는 경우가 많으므로
       거래처·품목이 같고 가로 또는 세로 하나가 비슷하면 후보로 본다."""
    if client_a != client_b:
        return False
    if (a.get("품목코드"), a.get("_mix_name"), a.get("_mix_codes")) != \
            (b.get("품목코드"), b.get("_mix_name"), b.get("_mix_codes")):
        return False
    return (_near(a.get("가로"), b.get("가로"), CHANGE_TOL)
            or _near(a.get("세로"), b.get("세로"), CHANGE_TOL))


def find_changes(orders, past):
    """최근 주문 중 변경 대상 후보를 찾는다.
       past = [(라벨, 거래처, item, 출력여부), ...]"""
    out, n = [], 0
    for o in orders:
        c = o.get("거래처")
        for it in o.get("items", []):
            if it.get("예외품목"):
                continue
            n += 1
            for label, pc, pit, printed in past:
                if not change_candidate(c, it, pc, pit):
                    continue
                diff = []
                for f, nm in (("가로", "가로"), ("세로", "세로"),
                              ("손잡이길이", "손잡이"), ("품목코드", "색상")):
                    a, b = pit.get(f), it.get(f)
                    if a != b:
                        diff.append(f"{nm} {a} → {b}")
                if not diff:
                    continue
                msg = f"{label} 주문의 변경일 수 있습니다 · " + ", ".join(diff)
                suffix = " · 이전 장부가 출력됨" if printed else ""
                out.append(("review", "변경의심", msg + suffix, n))
                break
    return out
