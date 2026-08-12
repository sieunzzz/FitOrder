"""
엑셀 발주 파서 (휴안 · DI) — AI 를 거치지 않고 열 위치로 직접 읽는다.
출력 형태는 extract.extract_order() 와 동일하다.
"""
import re

import openpyxl

from rules import DI_COLOR, DI_COLOR_UNSURE

SIZE = re.compile(r"(\d+(?:\.\d+)?)\s*[*x×X]\s*(\d+(?:\.\d+)?)")


def _blank_item():
    return {"품목코드": None, "색상원문": None, "타입": "C자", "종류": "투코드",
            "가로": None, "세로": None, "수량": None, "손잡이방향": None,
            "손잡이길이": None, "연창": False, "설치장소": None,
            "기재사항": None, "예외품목": None, "원문": None,
            "창개수": 1, "좌개수": None, "우개수": None,
            "확신도": {"가로": 1.0, "세로": 1.0, "품목코드": 1.0, "손잡이": 1.0}}


def _shell(client):
    return {"거래처": client, "발주일": None, "고객명": None, "items": [],
            "배송": {"방식": None, "화물지점": None, "주소": None,
                     "수령인": None, "연락처": None, "선불착불": None,
                     "전달사항": None},
            "전체원문": None, "전체기재사항": None}


# ─────────────────────────────────────────────
# 휴안
# ─────────────────────────────────────────────
def _huan_product(s):
    """'YL500 L원코드' -> (500, L자, 원코드)"""
    s = str(s or "").strip()
    m = re.search(r"([A-Z]{1,2})?\s*(\d{3}[A-Za-z]*)", s)
    code = m.group(2) if m else None
    type_ = "L자" if re.search(r"\bL\s*(?=원|투|셔|18|타입|자)", s) or "L18" in s else "C자"
    if "셔터" in s:
        kind = "셔터"
    elif "원코드" in s or re.search(r"원\s*코드", s):
        kind = "원코드"
    else:
        kind = "투코드"
    return code, type_, kind


SIDO_SHORT = {"서울특별시": "서울시", "부산광역시": "부산시", "대구광역시": "대구시",
              "인천광역시": "인천시", "광주광역시": "광주시", "대전광역시": "대전시",
              "울산광역시": "울산시", "세종특별자치시": "세종시"}
DO = ("경기도", "강원도", "강원특별자치도", "충청북도", "충청남도", "전라북도",
      "전북특별자치도", "전라남도", "경상북도", "경상남도", "제주특별자치도")


def _short_addr(addr):
    """'경상남도 진주시 초전북로 151 (초전동) 인사인광고기획'
         -> '진주시 초전북로 151, 인사인광고기획'
       '서울특별시 영등포구 도신로 161 (도림동, 정성빌) 201호'
         -> '서울시 영등포구 도신로 161, 정성빌 201호'"""
    s = str(addr or "").strip()
    for k, v in SIDO_SHORT.items():          # 광역시/특별시는 줄여서 유지
        if s.startswith(k):
            s = v + s[len(k):]
            break
    else:
        for d in DO:                          # 도는 삭제
            if s.startswith(d):
                s = s[len(d):]
                break

    def keep_building(m):
        """괄호 안에서 '○○동' 은 버리고 건물명만 남긴다"""
        parts = [x.strip() for x in m.group(1).split(",")]
        keep = [x for x in parts if not re.search(r"[동읍면리]$", x)]
        return (", " + " ".join(keep)) if keep else ""

    s = re.sub(r"\s*\(([^)]*)\)", keep_building, s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    if "," not in s:                      # 끝의 상호/건물명 앞에 콤마
        head, _, tail = s.rpartition(" ")
        if head and tail and not re.search(r"\d", tail):
            s = f"{head}, {tail}"
    return s


def parse_huan(path):
    """휴안 발주 엑셀 -> 주문 리스트 (수취인 단위로 분리)"""
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    orders, cur, last_prod = [], None, (None, "C자", "투코드", "")
    for r in range(2, ws.max_row + 1):
        g = lambda c: ws.cell(r, c).value
        size = g(9)
        if not size:
            continue
        name = g(3)
        if name:                                  # 새 주문 시작
            cur = _shell("휴안")
            d = g(1)
            cur["발주일"] = d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else None
            cur["고객명"] = str(name).strip()
            cur["배송"].update({
                "방식": "택배", "주소": _short_addr(g(6)),
                "수령인": str(name).strip(),
                "연락처": str(g(4) or "").strip(),
                "선불착불": "선불",
                "전달사항": (str(g(12)).strip() if g(12) else None),
            })
            cur["전체기재사항"] = "포장비용"       # 휴안은 창당 포장비
            orders.append(cur)
        if cur is None:
            continue
        m = SIZE.search(str(size))
        it = _blank_item()
        if m:
            it["가로"], it["세로"] = float(m.group(1)), float(m.group(2))
        if g(7):
            code, type_, kind = _huan_product(g(7))
            last_prod = (code, type_, kind, str(g(7)).strip())
        elif cur["items"]:                       # 상품명 공란 = 위 행과 동일
            code, type_, kind, _ = last_prod
        else:
            code, type_, kind = None, "C자", "투코드"
        it.update({"품목코드": code,
                   "색상원문": (str(g(7)).strip() if g(7) else last_prod[3]),
                   "타입": type_, "종류": kind,
                   "손잡이방향": (str(g(10)).strip() if g(10) else None),
                   "기재사항": cur["고객명"], "원문": f"{g(7)} {size} {g(10)}"})
        n = g(11)
        it["창개수"] = int(n) if isinstance(n, (int, float)) and n else 1
        cur["items"].append(it)
    return orders


# ─────────────────────────────────────────────
# DI
# ─────────────────────────────────────────────
DI_KIND = {"원": "원코드", "투": "투코드", "셔터": "셔터",
           "원코드": "원코드", "투코드": "투코드"}

# 헤더 이름 -> 내부 필드 (파일마다 열 위치가 달라 헤더로 찾는다)
DI_HEAD = {
    "수령인": "recv", "분류": "cls", "원단/색상": "color",
    "원/투": "kind", "커버/점보": "cover", "사이즈": "size",
    "줄": "line", "피스": "piece", "비고": "note",
}


def _di_header(ws):
    """헤더 행을 찾아 {필드: 열번호} 반환. 좌측(원본) 영역만 사용."""
    for r in range(1, 8):
        vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, 20)]
        if "사이즈" not in vals:
            continue
        col, seen = {}, set()
        for i, v in enumerate(vals, 1):
            f = DI_HEAD.get(v)
            if f and f not in seen:        # 같은 이름이 우측에 또 나오면 무시
                col[f] = i
                seen.add(f)
        if "size" in col and "color" in col:
            return r, col
    return None, {}


def parse_di(path, sheet="원본"):
    """DI 발주 엑셀(좌측 원본 영역) -> 주문 1건"""
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet not in wb.sheetnames:
        sheet = "장부" if "장부" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet]
    hr, col = _di_header(ws)
    if not col:
        raise ValueError("DI 발주서 형식을 인식하지 못했습니다 (헤더를 찾을 수 없음)")

    g = lambda r, f: (ws.cell(r, col[f]).value if f in col else None)
    o = _shell("DI")
    for r in range(hr + 1, ws.max_row + 1):
        size = g(r, "size")
        if not size:
            continue
        m = SIZE.search(str(size))
        if not m:
            continue
        it = _blank_item()
        cname = str(g(r, "color") or "").strip()
        it["색상원문"] = cname
        it["품목코드"] = DI_COLOR.get(cname)
        if cname in DI_COLOR_UNSURE or (cname and not it["품목코드"]):
            it["확신도"]["품목코드"] = 0.4      # 미등록 색상명 -> 담당자 확인
        # 종류/부속 판정 — 헤더 이름이 파일마다 달라 값으로 판단한다.
        # '원'/'투'/'셔터' 이면 종류, 그 밖의 문구면 커버·점보 등 부속.
        ks = ""
        for f in ("kind", "cover"):
            v = str(g(r, f) or "").strip()
            if not v or v in ("0", "-"):
                continue
            if v in DI_KIND:
                ks = v
            else:
                it["예외품목"] = "부속"
        it["종류"] = DI_KIND.get(ks, "투코드")
        it["타입"] = "L자"                       # DI 는 L타입 발주
        it["가로"], it["세로"] = float(m.group(1)), float(m.group(2))
        d = str(g(r, "line") or "").strip()
        if d.startswith(("좌", "우")):
            it["손잡이방향"] = d[0]
            n = re.search(r"(\d+)", d)
            if n:
                it["손잡이길이"] = int(n.group(1))
        recv = str(g(r, "recv") or "").strip()
        note = str(g(r, "note") or "").strip()
        piece = str(g(r, "piece") or "").strip()
        parts = [x for x in (recv, note, piece) if x and x != "None"]
        it["기재사항"] = "/".join(parts) or None
        it["원문"] = f"{cname} {ks} {size} {d} {note}".strip()
        o["items"].append(it)
    o["배송"]["방식"] = None                      # DI/SP 출고
    return o


# ─────────────────────────────────────────────
def parse_excel(path, client):
    if client == "휴안":
        return parse_huan(path)
    if client == "DI":
        return [parse_di(path)]
    raise ValueError(f"엑셀 파서 없음: {client}")


if __name__ == "__main__":
    import json, sys
    p = sys.argv[1]
    c = "DI" if "DI" in p.upper() else "휴안"
    for o in parse_excel(p, c):
        print(json.dumps(o, ensure_ascii=False, indent=2, default=str))
