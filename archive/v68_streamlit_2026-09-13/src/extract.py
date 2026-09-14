"""
FitOrder 추출 엔진

발주서 이미지(또는 여러 장) -> 표준 JSON

사용:
    from extract import extract_order
    result = extract_order(["samples/DU/DU_01.png"], client_hint="DU")
"""
import base64
import io
import json
import os
import re
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
import importlib
import sys
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI
from PIL import Image, ImageOps

# 프로젝트 최상단의 .env 를 읽는다 (src 에서 실행되므로 상위 폴더)
load_dotenv(Path(__file__).resolve().parent.parent / ".env")

from extract_schema import RESPONSE_FORMAT
from prompt import SYSTEM_PROMPT, user_prompt
from holding import (HOLDING_FEATURE_ENABLED, holding_accessory_from_text, holding_product_from_text,
                     looks_like_holding_operation, normalize_holding_operation)

# 저렴한 비전 모델부터 시작. 정확도가 부족하면 상위 모델로 교체.
MODEL = os.getenv("FITORDER_MODEL", "gpt-4o")
MAX_SIDE = 2000          # 이보다 크면 축소
MIN_SIDE = 1600          # 이보다 작으면 확대 (촘촘한 표 판독용)
MAX_RETRY = 3

TIMEOUT = 90.0           # 응답이 없으면 90초 후 중단 (방화벽 무한대기 방지)

_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(timeout=TIMEOUT, max_retries=0) if _key else None


def extract_pdf(pdf_path, client_hint=None, ship_date=None, model=None):
    """PDF를 렌더링한다. SP는 페이지별 추출 후 같은 주문번호를 합친다."""
    try:
        import pymupdf
    except ImportError as e:
        # 기존 PyInstaller exe에 포함되지 않은 새 패키지는
        # exe 옆 포터블 python의 site-packages에서 추가로 찾는다.
        root = Path(__file__).resolve().parent.parent
        candidates = [root / "python" / "Lib" / "site-packages"]
        candidates += list((root / "python" / "lib").glob("python*/site-packages"))
        for site in candidates:
            if not site.exists():
                continue
            if str(site) not in sys.path:
                sys.path.insert(0, str(site))
            if os.name == "nt" and hasattr(os, "add_dll_directory"):
                for dll_dir in (site, site / "pymupdf"):
                    if dll_dir.exists():
                        try:
                            os.add_dll_directory(str(dll_dir))
                        except OSError:
                            pass
        importlib.invalidate_caches()
        try:
            import pymupdf
        except ImportError:
            raise RuntimeError(
                "PDF 처리 모듈이 없습니다. FitOrder.exe 옆에서 "
                "PDF기능_설치.bat를 한 번 실행해 주세요."
            ) from e

    doc = pymupdf.open(str(pdf_path))
    try:
        if doc.needs_pass:
            raise ValueError("암호가 걸린 PDF는 읽을 수 없습니다")
        if doc.page_count == 0:
            raise ValueError("페이지가 없는 PDF입니다")
        if doc.page_count > 50:
            raise ValueError("PDF는 한 번에 50페이지까지 처리할 수 있습니다")

        with tempfile.TemporaryDirectory(prefix="fitorder_pdf_") as td:
            paths = []
            matrix = pymupdf.Matrix(2, 2)  # 약 144 dpi: 표의 작은 글자 판독용
            for i, page in enumerate(doc):
                pix = page.get_pixmap(matrix=matrix, alpha=False)
                p = Path(td) / f"page_{i + 1:03d}.png"
                pix.save(str(p))
                if client_hint == "SP":
                    # 스페이스 팩스 PDF는 표가 옆으로 누워 스캔된다.
                    with Image.open(p) as img:
                        rotated = img.rotate(90, expand=True).convert("RGB")
                        # SP 팩스는 작업일지가 왼쪽에 있고 오른쪽 약 40%가
                        # 빈 스캔 영역이다. 표가 잘리지 않는 범위에서 먼저 제거한다.
                        rotated = rotated.crop(
                            (0, 0, int(rotated.width * 0.62), rotated.height))
                        # 팩스 스캔의 오른쪽 큰 흰 여백을 제거해 표와
                        # 손글씨 숫자가 비전 모델에 더 크게 전달되게 한다.
                        gray = ImageOps.grayscale(rotated)
                        ink = gray.point(lambda v: 255 if v < 220 else 0)
                        box = ink.getbbox()
                        if box:
                            pad = 24
                            box = (max(0, box[0] - pad), max(0, box[1] - pad),
                                   min(rotated.width, box[2] + pad),
                                   min(rotated.height, box[3] + pad))
                            rotated = rotated.crop(box)
                        # 연한 연필 숫자의 대비를 높여 1/7, 3/5, 0/6 혼동을 줄인다.
                        rotated = ImageOps.autocontrast(
                            ImageOps.grayscale(rotated), cutoff=1).convert("RGB")
                        rotated.save(p, format="PNG", optimize=True)
                paths.append(p)

            if client_hint != "SP":
                order = extract_order(paths, client_hint=client_hint,
                                      ship_date=ship_date, model=model)
                order["_source_pages"] = list(range(1, len(paths) + 1))
                return [order]

            # 기본은 1페이지씩 처리한다. 여러 페이지를 동시 호출하면
            # 공장용 API 계정의 RPM/TPM 한도를 넘어 RateLimitError가 발생한다.
            workers = min(int(os.getenv("FITORDER_PDF_WORKERS", "1")), len(paths))
            def one(p):
                return extract_order([p], client_hint="SP",
                                     ship_date=ship_date, model=model)
            if workers <= 1:
                page_orders = []
                delay = float(os.getenv("FITORDER_PDF_PAGE_DELAY", "1.0"))
                for i, p in enumerate(paths):
                    page_orders.append(one(p))
                    if delay > 0 and i < len(paths) - 1:
                        time.sleep(delay)
            else:
                with ThreadPoolExecutor(max_workers=workers) as pool:
                    page_orders = list(pool.map(one, paths))
            empty_pages = [str(i + 1) for i, order in enumerate(page_orders)
                           if not (order.get("items") or [])]
            if empty_pages:
                raise ValueError(
                    "SP PDF에서 품목을 읽지 못한 페이지: "
                    + ", ".join(empty_pages)
                    + "페이지. 해당 페이지를 확인해 주세요."
                )
            for page_no, order in enumerate(page_orders, 1):
                order["_source_pages"] = [page_no]
            return _merge_sp_orders(page_orders)
    finally:
        doc.close()


def _sp_order_key(value):
    """'<17-9>', '17 - 9' 등을 '17-9'로 통일한다."""
    m = re.search(r"(\d{1,2})\s*[-–]\s*(\d+)", str(value or ""))
    return f"{int(m.group(1))}-{int(m.group(2))}" if m else None


def _merge_sp_orders(page_orders):
    """입력 순서를 유지하며 같은 SP 주문번호의 품목을 합친다."""
    merged = {}
    order_keys = []
    for page_i, order in enumerate(page_orders):
        key = _sp_order_key(order.get("주문번호"))
        if key:
            order["주문번호"] = key
        unique_key = key or f"__page_{page_i}"
        if unique_key not in merged:
            merged[unique_key] = order
            order_keys.append(unique_key)
            continue
        dst = merged[unique_key]
        dst.setdefault("items", []).extend(order.get("items") or [])
        dst["_source_pages"] = list(dict.fromkeys(
            (dst.get("_source_pages") or [])
            + (order.get("_source_pages") or [])))
        for field in ("전체기재사항", "전체원문"):
            vals = []
            for value in (dst.get(field), order.get(field)):
                for part in str(value or "").split("/"):
                    part = part.strip()
                    if part and part not in vals:
                        vals.append(part)
            dst[field] = "/".join(vals) or None
    return [merged[k] for k in order_keys]


def diagnose():
    """실패 원인을 한국어로 돌려준다. 정상이면 None."""
    import urllib.request
    if not _key:
        return ("API 키가 없습니다.\n"
                "프로그램 폴더에 .env 파일이 있는지 확인해 주세요.")
    try:
        req = urllib.request.Request("https://api.openai.com/v1/models",
                                     headers={"Authorization": f"Bearer {_key}"})
        urllib.request.urlopen(req, timeout=15)
        return None
    except urllib.error.HTTPError as e:
        if e.code == 401:
            return "API 키가 올바르지 않습니다. 키를 다시 발급받아야 합니다."
        if e.code == 429:
            return ("사용 한도를 초과했거나 잔액이 부족합니다.\n"
                    "OpenAI 계정에서 크레딧을 충전해 주세요.")
        if e.code == 403:
            return ("접속이 차단되었습니다.\n"
                    "회사 네트워크나 보안 프로그램이 막고 있을 수 있습니다.\n"
                    "휴대폰 핫스팟으로 연결해 보시거나\n"
                    "네트워크 담당자에게 api.openai.com 허용을 요청해 주세요.")
        return f"서버 응답 오류 (코드 {e.code}) — 잠시 후 다시 시도해 주세요."
    except Exception:
        return ("인터넷 연결이 되지 않거나 회사 네트워크에서 차단되어 있습니다.\n"
                "다른 인터넷(휴대폰 핫스팟 등)으로 시도해 보시거나\n"
                "네트워크 담당자에게 api.openai.com 허용을 요청해 주세요.")


def _encode(path: str) -> str:
    """긴 변을 MIN_SIDE ~ MAX_SIDE 범위로 맞춘 뒤 base64 PNG 로 반환.
       캡처가 작으면 확대한다 — 타일 수가 늘어 작은 글자 판독이 좋아진다."""
    img = Image.open(path)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    w, h = img.size
    long_side = max(w, h)
    if long_side > MAX_SIDE:
        k = MAX_SIDE / long_side
    elif long_side < MIN_SIDE:
        k = MIN_SIDE / long_side
    else:
        k = 1
    if k != 1:
        img = img.resize((int(w * k), int(h * k)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return base64.b64encode(buf.getvalue()).decode()


def _encode_jo_table_crop(path: str) -> str:
    """제이원 표의 작은 수량/좌/우 숫자를 확대해 같은 호출에 보조 이미지로 전달한다."""
    img = Image.open(path)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    w, h = img.size
    # 제이원 주문서의 품목 표는 보통 하단에 있다. 좌우 여백도 조금 제거한다.
    y0 = int(h * 0.43)
    x0, x1 = int(w * 0.015), int(w * 0.985)
    crop = img.crop((x0, y0, x1, h))
    # 표 선/작은 숫자 대비를 살리되 색상은 유지한다.
    if crop.mode == "RGB":
        crop = ImageOps.autocontrast(crop, cutoff=0.5)
    long_side = max(crop.size)
    if long_side < 2200:
        k = 2200 / long_side
        crop = crop.resize((int(crop.width * k), int(crop.height * k)), Image.LANCZOS)
    elif long_side > 2600:
        k = 2600 / long_side
        crop = crop.resize((int(crop.width * k), int(crop.height * k)), Image.LANCZOS)
    buf = io.BytesIO()
    crop.save(buf, format="PNG", optimize=True)
    return base64.b64encode(buf.getvalue()).decode()


def _clean_jo_note(value):
    """제이원 장부 기재사항에서는 손/줄/봉 길이 표기를 제거한다."""
    text = str(value or "").strip()
    if not text:
        return None
    text = re.sub(r"(?:손잡이(?:길이)?|손|줄|봉)\s*[:=]?\s*\d{2,3}", " ", text)
    text = re.sub(r"\s*[/|,]+\s*", "/", text)
    text = re.sub(r"/{2,}", "/", text).strip(" /,.-")
    return text or None


def _postprocess_jo(data):
    """제이원 주문번호/표 수량/좌우/메모를 결정 규칙으로 한 번 더 고정한다."""
    whole = str(data.get("전체원문") or "")
    raw_no = str(data.get("주문번호") or "").strip()
    if not raw_no:
        m = re.search(r"주문\s*번\s*호\s*[:：]?\s*([A-Za-z0-9-]+)", whole)
        if m:
            raw_no = m.group(1).strip()
    data["주문번호"] = raw_no or None

    # 전체기재사항에는 주문번호를 중복 보관하지 않는다. 장부 기재사항1은 주문번호 필드에서 만든다.
    common = []
    for part in str(data.get("전체기재사항") or "").split("/"):
        part = _clean_jo_note(part)
        if not part:
            continue
        if raw_no and re.sub(r"[()（）\s]", "", part) == re.sub(r"[()（）\s]", "", raw_no):
            continue
        if part not in common:
            common.append(part)
    data["전체기재사항"] = "/".join(common) or None

    for item in data.get("items") or []:
        raw = " ".join(str(item.get(k) or "") for k in ("원문", "기재사항"))
        if item.get("손잡이길이") in (None, ""):
            hm = re.search(r"(?:손잡이(?:길이)?|손|줄|봉)\s*[:=]?\s*(\d{2,3})", raw)
            if hm:
                item["손잡이길이"] = int(hm.group(1))
        item["기재사항"] = _clean_jo_note(item.get("기재사항"))

        try:
            left_n = int(float(item.get("좌개수") or 0))
            right_n = int(float(item.get("우개수") or 0))
        except (TypeError, ValueError):
            left_n = right_n = 0
        total_lr = left_n + right_n
        try:
            count = int(float(item.get("창개수"))) if item.get("창개수") not in (None, "") else 0
        except (TypeError, ValueError):
            count = 0
        if count <= 0 and total_lr > 0:
            item["창개수"] = total_lr
            count = total_lr
        if left_n > 0 and right_n == 0:
            item["손잡이방향"] = "좌"
        elif right_n > 0 and left_n == 0:
            item["손잡이방향"] = "우"
    return data


def _normalize_holding_item(item):
    """모델의 홀딩도어 필드를 공식 품목/내부 필드로 정규화한다."""
    if not HOLDING_FEATURE_ENABLED:
        # 홀딩 자동 판별을 사용하지 않는 동안에는 어떤 H/홀딩/자바라 표기도
        # 별도 제품군으로 바꾸지 않는다. 원래 읽은 일반 주문 필드를 유지한다.
        item["제품군"] = "블라인드"
        item["홀딩방식"] = None
        item["홀딩레일"] = None
        item["홀딩상하로라"] = False
        item["홀딩부속"] = None
        item["홀딩부속색상"] = None
        exc = str(item.get("예외품목") or "").strip()
        if exc and ("홀딩" in exc or "자바라" in exc):
            item["예외품목"] = None
        item.pop("_product_group", None)
        item.pop("_ledger_prefix", None)
        for key in list(item):
            if key.startswith("_holding_"):
                item.pop(key, None)
        return item
    raw = " ".join(str(item.get(k) or "") for k in
                   ("색상원문", "품목코드", "원문", "홀딩부속"))
    accessory_name = item.get("홀딩부속")
    accessory = None
    if accessory_name:
        accessory = holding_accessory_from_text(
            f"{accessory_name} {item.get('홀딩부속색상') or '화이트'}",
            item.get("홀딩부속색상") or "화이트")
    if not accessory:
        accessory = holding_accessory_from_text(raw,
                                                item.get("홀딩부속색상") or "화이트")
    product = None if accessory else (holding_product_from_text(item.get("색상원문"))
                                      or holding_product_from_text(item.get("품목코드"))
                                      or holding_product_from_text(raw))
    explicit = item.get("제품군") == "홀딩도어" or bool(accessory_name)
    if not explicit and not accessory and not product:
        return item

    item["_product_group"] = "holding"
    item["_ledger_prefix"] = "H"
    item["예외품목"] = None
    item["타입"], item["종류"] = "C자", "투코드"
    item["손잡이방향"] = None
    item["손잡이길이"] = None
    item["연창"] = False

    if accessory:
        item["_holding_accessory"] = True
        item["_holding_product_name"] = accessory["품명"]
        item["_holding_accessory_label"] = accessory.get("장부표시")
        item["_holding_accessory_color"] = accessory.get("부속색상") or "화이트"
        item["품목코드"] = accessory.get("코드")
        # 부속 개수는 structured output의 창개수 또는 수량을 그대로 사용한다.
        return item

    item.pop("_holding_accessory", None)
    if product:
        item["_holding_product_name"] = product["품명"]
        item["_holding_label"] = product.get("장부표시")
        item["품목코드"] = product.get("코드")
    operation = item.get("홀딩방식")
    if not operation and looks_like_holding_operation(item.get("수량")):
        operation = item.get("수량")
        item["수량"] = None
    item["_holding_operation"] = normalize_holding_operation(operation)
    item["_holding_rail"] = str(item.get("홀딩레일") or "").strip() or None
    item["_holding_upper_roller"] = bool(item.get("홀딩상하로라")) or \
        "+상하로라" in raw.replace(" ", "")
    return item


def _postprocess_azit(data):
    """아지트의 `82,154,82×220 ... 좌,우,우` 축약을 실제 창 단위로 펼친다.

    모델이 쉼표 폭 목록을 한두 창으로 축약해도 원문 숫자 개수와 방향 개수를
    다시 읽어 같은 높이의 독립 item으로 복원한다.
    """
    whole = str(data.get("전체원문") or "")
    if not whole:
        return data
    lines = [re.sub(r"\s+", " ", x).strip() for x in whole.splitlines() if x.strip()]
    parsed = []
    for i, line in enumerate(lines):
        # 방/위치명 + 쉼표로 나열된 가로들 + X + 세로
        m = re.search(r"([^\d\n]{1,40}?)(\d+(?:\s*,\s*\d+)+)\s*[xX×*]\s*(\d+)(.*)$", line)
        if not m:
            continue
        prefix = m.group(1).strip(" /,:-")
        # 긴 주소/문장이 붙어도 치수 바로 앞의 마지막 토큰을 설치장소로 사용한다.
        room_m = re.search(r"([가-힣A-Za-z][가-힣A-Za-z0-9_-]{0,20})$", prefix)
        room = room_m.group(1) if room_m else prefix
        widths = [float(x) for x in re.findall(r"\d+", m.group(2))]
        height = float(m.group(3))
        tail = m.group(4) or ""
        if i + 1 < len(lines) and not re.search(r"\d+(?:\s*,\s*\d+)+\s*[xX×*]\s*\d+", lines[i + 1]):
            tail += " " + lines[i + 1]
        # 코드, 타입, 종류, 방향은 치수 뒤쪽 문구에서 결정한다.
        code_m = re.search(r"(?:[A-Za-z]{1,3}\s*[-_]?)?(\d{3,4}[A-Za-z]*)", tail)
        code = code_m.group(1) if code_m else None
        type_ = "L자" if re.search(r"L\s*(?:타입|형|자|18|21)", tail, re.I) else "C자"
        kind = "원코드" if "원코드" in tail else ("셔터" if "셔터" in tail else "투코드")
        dirs = re.findall(r"[좌우]", tail)
        # 방향이 폭 개수보다 짧으면 추측하지 않고 None을 채운다.
        dirs = (dirs + [None] * len(widths))[:len(widths)]
        for j, width in enumerate(widths):
            template = dict((data.get("items") or [{}])[0])
            template.update({
                "제품군": "블라인드", "품목코드": code or template.get("품목코드"),
                "색상원문": template.get("색상원문") or (code or None),
                "타입": type_, "종류": kind, "가로": width, "세로": height,
                "수량": None, "손잡이방향": dirs[j], "설치장소": room or None,
                "창개수": 1, "좌개수": 1 if dirs[j] == "좌" else 0,
                "우개수": 1 if dirs[j] == "우" else 0,
                "원문": f"{room}{m.group(2)}x{int(height) if height.is_integer() else height} {tail}".strip(),
            })
            conf = dict(template.get("확신도") or {})
            conf.update({"가로": 1.0, "세로": 1.0, "손잡이": 1.0})
            conf.setdefault("품목코드", 0.9 if code else 0.5)
            template["확신도"] = conf
            parsed.append(template)
    if not parsed:
        return data

    # 쉼표 축약으로 표현되지 않은 기존 단일창은 보존한다.
    kept = []
    for it in data.get("items") or []:
        try:
            w, h = float(it.get("가로")), float(it.get("세로"))
        except (TypeError, ValueError):
            kept.append(it)
            continue
        same_group = any(abs(float(p.get("세로")) - h) < 0.001
                         and (not p.get("품목코드") or not it.get("품목코드")
                              or str(p.get("품목코드")) == str(it.get("품목코드")))
                         and str(p.get("설치장소") or "") == str(it.get("설치장소") or "")
                         for p in parsed)
        if not same_group:
            kept.append(it)
    data["items"] = parsed + kept
    return data


def _clean_rt_note_text(value):
    """RT 제작 기재사항에서 배송 결제 문구를 제거한다."""
    text = str(value or "").strip()
    if not text:
        return None
    text = re.sub(r"(?:택배비\s*)?(?:선불|착불)", " ", text, flags=re.I)
    text = re.sub(r"\s*[/|,]+\s*", "/", text)
    text = re.sub(r"/{2,}", "/", text).strip(" /,-")
    return text or None


def _rt_receiver_from_item(item):
    """루임트 주소 뒤 수령인 표기의 결정적 약칭을 복원한다."""
    raw = " ".join(str(item.get(k) or "") for k in
                   ("기재사항", "원문", "색상원문"))
    if "더커튼" in raw or re.search(r"(?:^|[/,\s])더(?:$|[/,\s])", raw):
        return "더"
    if re.search(r"(?:^|[/,\s])지엘(?:$|[/,\s])", raw):
        return "지엘"
    if re.search(r"(?:^|[/,\s])H(?:$|[/,\s])", raw, re.I):
        return "H"
    return None


def _postprocess_rt(data):
    """루임트는 주소/수령인과 결제 표기를 후처리로 한 번 더 고정한다."""
    delivery = data.setdefault("배송", {})
    # 착불/선불은 배송 정보에만 남기고 제작 메모에서는 제거한다.
    delivery["전달사항"] = _clean_rt_note_text(delivery.get("전달사항"))
    data["전체기재사항"] = _clean_rt_note_text(data.get("전체기재사항"))
    receiver_common = str(delivery.get("수령인") or "").strip()
    if receiver_common == "더커튼":
        receiver_common = "더"
        delivery["수령인"] = "더"
    for item in data.get("items") or []:
        note = _clean_rt_note_text(item.get("기재사항"))
        receiver = _rt_receiver_from_item(item) or receiver_common
        if receiver == "더커튼":
            receiver = "더"
        parts = [x.strip() for x in str(note or "").split("/") if x.strip()]
        # 더커튼은 항상 '더'로 축약한다.
        parts = ["더" if x == "더커튼" else x for x in parts]
        if receiver and receiver not in parts:
            parts.insert(0, receiver)
        item["기재사항"] = "/".join(dict.fromkeys(parts)) or None
    return data


def _clean_true_accessory_note(value):
    """인천)트루 부속은 별도 행으로 만들되 일반 `피스추가` 표시는 보존한다.

    사용자가 요청한 것은 노피스/스냅의 *개수*를 기재사항에 중복하지 않는 것이다.
    따라서 `피스추가` 자체는 장부/EDI의 `피스`로 남긴다.
    """
    text = str(value or "").strip()
    if not text:
        return None
    # 부속 원문과 뒤따르는 수량 표현은 별도 부속행으로 처리한다.
    text = re.sub(r"창틀용\s*무타공\s*(?:\d+\s*(?:세트|개|EA))?", " ", text, flags=re.I)
    text = re.sub(r"커튼박스용\s*무타공\s*(?:\d+\s*(?:세트|개|EA))?", " ", text, flags=re.I)
    text = re.sub(r"(?<![가-힣A-Za-z])스냅\s*(?:\d+\s*(?:세트|개|EA))?(?![가-힣A-Za-z])", " ", text, flags=re.I)
    # 노피스(1)/(2) 같은 부속 개수 표시는 메모에서 제거한다.
    text = re.sub(r"노피스\s*(?:\(\s*\d+\s*\)|\d+)?", " ", text, flags=re.I)
    # 일반 피스 요청은 남긴다. `피스추가` -> `피스`.
    text = re.sub(r"피스\s*추가", "피스", text, flags=re.I)
    # 트루 공통 배송/출고 문구는 제작 기재사항에서 제외한다.
    text = re.sub(r"택배비\s*선불", " ", text)
    text = re.sub(r"빠른\s*출고", " ", text)
    text = re.sub(r"\s*[+|,]+\s*", "/", text)
    text = re.sub(r"\s*/\s*", "/", text)
    text = re.sub(r"/{2,}", "/", text).strip(" /,.-+")
    # 중복 토큰 제거. `피스/피스` 방지.
    parts = []
    for part in text.split('/'):
        part = part.strip()
        if part and part not in parts:
            parts.append(part)
    return "/".join(parts) or None


def _postprocess_true(data):
    """인천)트루 부속/배송 규칙을 코드로 확정한다."""
    whole = " ".join(str(x or "") for x in (
        data.get("전체원문"), data.get("전체기재사항")))
    item_raw = " ".join(" ".join(str(it.get(k) or "") for k in
                                   ("원문", "기재사항", "색상원문"))
                        for it in (data.get("items") or []))
    raw = f"{whole} {item_raw}"
    accessories = []
    if re.search(r"창틀용\s*무타공", raw):
        accessories.append({"표시": "노피스(2)", "품명": "노피스브라켓(2EA)",
                            "관리코드": "노피스브라켓(2EA)", "수량": 1})
    if re.search(r"커튼박스용\s*무타공", raw):
        accessories.append({"표시": "노피스(1)", "품명": "노피스브라켓(1EA)-커튼박스용",
                            "관리코드": "노피스브라켓(1EA)-커튼박스용", "수량": 1})
    if re.search(r"(?<![가-힣A-Za-z])스냅(?![가-힣A-Za-z])", raw):
        accessories.append({"표시": "B 원코드 브라켓", "품명": "B브라켓(25mm원코드/구)",
                            "관리코드": "B브라켓(25mm원코드/구)", "수량": 1})
    if accessories:
        # 같은 부속이 원문 여러 곳에 반복되어도 주문당 한 행만 만든다.
        uniq = []
        seen = set()
        for acc in accessories:
            key = acc["표시"]
            if key not in seen:
                uniq.append(acc); seen.add(key)
        data["_true_accessories"] = uniq

    # 트루 메시지는 배송을 기본 택배로 본다. 모델이 주소/선불을 읽었으면 그대로 보존한다.
    delivery = data.setdefault("배송", {})
    if delivery.get("주소") and not delivery.get("방식"):
        delivery["방식"] = "택배"
    if re.search(r"택배비\s*선불", raw) and not delivery.get("선불착불"):
        delivery["선불착불"] = "선불"

    data["전체기재사항"] = _clean_true_accessory_note(data.get("전체기재사항"))
    for item in data.get("items") or []:
        item["기재사항"] = _clean_true_accessory_note(item.get("기재사항"))
    return data


def _mark_non_di_mix(data):
    """대일 외 거래처는 원문에 MIX가 보이면 표시용 플래그를 보존한다."""
    if data.get("거래처") == "DI":
        return
    for item in data.get("items") or []:
        raw = " ".join(str(item.get(k) or "") for k in
                       ("색상원문", "원문", "기재사항"))
        if re.search(r"(?<![A-Za-z])MIX(?![A-Za-z])", raw, re.I):
            item["_mix_word"] = True


def extract_order(image_paths, client_hint=None, ship_date=None, model=None):
    """여러 장을 한 번의 호출에 함께 넣는다 (스크롤 분할·첨부 사진 대응)"""
    if isinstance(image_paths, (str, Path)):
        image_paths = [image_paths]

    content = [{"type": "text", "text": user_prompt(client_hint, ship_date)}]
    for p in image_paths:
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{_encode(str(p))}",
                          "detail": "high"},
        })
        if client_hint == "JO":
            content.append({"type": "text", "text": "제이원 품목표 확대본입니다. 수량/손잡이 좌/우 숫자를 행별로 다시 대조하세요."})
            content.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{_encode_jo_table_crop(str(p))}",
                              "detail": "high"},
            })

    if client is None:
        raise RuntimeError("API 키가 없습니다. .env 파일을 확인해 주세요.")

    last = None
    for attempt in range(MAX_RETRY):
        try:
            resp = client.chat.completions.create(
                model=model or MODEL,
                messages=[{"role": "system", "content": SYSTEM_PROMPT},
                          {"role": "user", "content": content}],
                response_format=RESPONSE_FORMAT,
                temperature=0,
            )
            data = json.loads(resp.choices[0].message.content)
            # 사용자 선택이나 파일명으로 확정한 거래처는
            # 모델이 다른 상호로 잘못 판독해도 바꾸지 못하게 한다.
            if client_hint:
                data["거래처"] = client_hint
            for item in data.get("items") or []:
                if str(item.get("손잡이방향") or "").strip() == "ㅈ":
                    item["손잡이방향"] = "좌"
                _normalize_holding_item(item)
                if client_hint == "SP" and item.get("_product_group") != "holding":
                    raw = " ".join(str(item.get(k) or "")
                                   for k in ("수량", "원문", "기재사항"))
                    fraction = re.search(r"(?<!\d)1\s*/\s*([2-9])(?!\d)", raw)
                    if fraction:
                        item["수량"] = f"1/{fraction.group(1)}"
                    else:
                        for symbol, value in (("½", "1/2"), ("⅓", "1/3"),
                                              ("¼", "1/4")):
                            if symbol in raw:
                                item["수량"] = value
                                break
                    # SP 양식의 '피스'는 미리 인쇄된 항목명이다.
                    # 빈 칸의 항목명이 모든 창의 메모로 복사되지 않게 한다.
                    note = str(item.get("기재사항") or "")
                    parts = [p.strip() for p in note.split("/")
                             if p.strip() and p.strip() != "피스"]
                    item["기재사항"] = "/".join(parts) or None
                    # 수기 '이타입'은 일반 C자 형상을 뜻한다.
                    if str(item.get("타입") or "").strip().upper() in {
                            "C", "C타입", "C형", "이타입"}:
                        item["타입"] = "C자"
                if data.get("거래처") == "MS":
                    raw = " ".join(str(item.get(k) or "") for k in
                                   ("색상원문", "원문", "기재사항"))
                    # 미성텍스의 자연어 약칭을 정답 장부 값으로 확정한다.
                    if "알루미늄" in raw and "화이트" in raw:
                        item["품목코드"] = "102"
                    if "원코드" in raw:
                        item["종류"] = "원코드"
                    elif "투코드" in raw:
                        item["종류"] = "투코드"
                    item["타입"] = "C자"
                    direction = re.search(r"(?:^|[\s,])([좌우])(?:[\s,]|$)", raw)
                    if direction:
                        item["손잡이방향"] = direction.group(1)
                    handle = re.search(r"(?:조절|줄|손)\s*(\d{2,3})", raw)
                    if handle:
                        item["손잡이길이"] = int(handle.group(1))

                    # 출고 요청은 장부 기재사항이 아니다.
                    note = str(item.get("기재사항") or "")
                    note = re.sub(
                        r"(?:오늘|내일|모레|월요일|화요일|수요일|목요일|금요일|토요일|일요일)?\s*"
                        r"출고\s*(?:부탁드립니다|부탁드려요|요청|희망)?", "", note)
                    item["기재사항"] = note.strip(" /,.") or None
            _mark_non_di_mix(data)
            if data.get("거래처") == "RT":
                _postprocess_rt(data)
            if data.get("거래처") == "JO":
                _postprocess_jo(data)
            if data.get("거래처") == "인천)트루":
                _postprocess_true(data)
            if data.get("거래처") == "아지트":
                _postprocess_azit(data)
            if data.get("거래처") == "MS":
                common = str(data.get("전체기재사항") or "")
                common = re.sub(
                    r"(?:오늘|내일|모레|월요일|화요일|수요일|목요일|금요일|토요일|일요일)?\s*"
                    r"출고\s*(?:부탁드립니다|부탁드려요|요청|희망)?", "", common)
                data["전체기재사항"] = common.strip(" /,.") or None
            if data.get("거래처") == "JL":
                # JL은 모델 판독 뒤에도 결정 가능한 규칙을 한 번 더 강제한다.
                whole = str(data.get("전체원문") or "")
                # 발주번호는 우측 상단 괄호 안 숫자가 정답이다. 모델이 놓치면
                # 명시 라벨 -> 괄호 안 숫자 순으로 보완하고, 내부 값에는 괄호를 제거한다.
                raw_no = str(data.get("주문번호") or "").strip()
                raw_no = re.sub(r"^[（(]\s*|\s*[）)]$", "", raw_no).strip()
                if not raw_no:
                    no = re.search(
                        r"(?:발주|주문)\s*번\s*호\s*[:：]?\s*[（(]?\s*([A-Za-z0-9-]+)\s*[）)]?", whole)
                    if not no:
                        no = re.search(r"[（(]\s*(\d[\d-]*)\s*[）)]", whole)
                    if no:
                        raw_no = no.group(1)
                data["주문번호"] = raw_no or None

                # JL 이름 뒤 내부 표기 DW는 장부/EDI에 필요 없다.
                if data.get("고객명"):
                    data["고객명"] = re.sub(r"\s*DW\s*$", "",
                                             str(data["고객명"]).strip(), flags=re.I).strip() or None
                if not data.get("고객명"):
                    receiver = str((data.get("배송") or {}).get("수령인") or "").strip()
                    receiver = re.sub(r"\s*DW\s*$", "", receiver, flags=re.I).strip()
                    if receiver:
                        data["고객명"] = receiver
                delivery = data.get("배송") or {}
                if delivery.get("수령인"):
                    delivery["수령인"] = re.sub(r"\s*DW\s*$", "",
                                                  str(delivery["수령인"]).strip(),
                                                  flags=re.I).strip() or None

                common_parts = []
                jl_customer = re.sub(r"\s*DW\s*$", "",
                                     str(data.get("고객명") or "").strip(), flags=re.I).strip()
                for part in str(data.get("전체기재사항") or "").split("/"):
                    part = part.strip()
                    if not part or re.search(r"(?:비닐\s*포장|겉\s*비닐|걷\s*비닐)", part):
                        continue
                    part_no_dw = re.sub(r"\s*DW\s*$", "", part, flags=re.I).strip()
                    if jl_customer and part_no_dw == jl_customer:
                        continue
                    if re.fullmatch(r"DW", part, re.I):
                        continue
                    if part_no_dw not in common_parts:
                        common_parts.append(part_no_dw)
                # `원코드`는 JL 발주서에서 고정 칸이 아니다. 각 행 원문/색상원문/
                # 기재사항 어디에서든 읽고, 페이지 전체가 원코드 단일 종류일 때는
                # 전체 품목에 적용한다.
                whole_has_one = bool(re.search(r"원\s*코드", whole))
                whole_has_two = bool(re.search(r"투\s*코드", whole))
                explicit_piece = "피스" in whole
                for item in data.get("items") or []:
                    raw = str(item.get("원문") or "")
                    note = str(item.get("기재사항") or "")
                    kind_text = " ".join(str(item.get(k) or "") for k in
                                         ("색상원문", "원문", "기재사항"))
                    if re.search(r"원\s*코드", kind_text) or (whole_has_one and not whole_has_two):
                        item["종류"] = "원코드"
                    elif re.search(r"투\s*코드", kind_text):
                        item["종류"] = "투코드"
                    if "피스" in raw or "피스" in note:
                        explicit_piece = True
                    # 포장 지시는 출력 대상이 아니다.
                    cleaned = []
                    for part in note.split("/"):
                        part = part.strip()
                        if not part or re.search(
                                r"(?:비닐\s*포장|겉\s*비닐|걷\s*비닐)", part):
                            continue
                        cleaned.append(part)
                    note = "/".join(cleaned)
                    # 해당 행 원문에 '틀안'이 없는데 모델이 만들어낸 경우 제거한다.
                    if raw and "틀안" not in raw:
                        note = re.sub(r"(?:^|/)\s*틀안\s*(?=/|$)", "", note)
                        note = re.sub(r"/{2,}", "/", note).strip("/")
                    item["기재사항"] = note or None
                    # L 표기가 실제 행 원문에 없으면 C타입으로 되돌린다.
                    if item.get("타입") == "L자" and raw:
                        explicit_l = re.search(
                            r"(?:^|[\s/,(])L\s*(?:자|타입|형|18|21)(?:$|[\s/),])", raw, re.I)
                        if not explicit_l:
                            item["타입"] = "C자"
                if explicit_piece and "피스" not in common_parts:
                    common_parts.insert(0, "피스")
                data["전체기재사항"] = "/".join(common_parts) or None
            data["_meta"] = {
                "model": resp.model,
                "files": [os.path.basename(str(p)) for p in image_paths],
                "tokens_in": resp.usage.prompt_tokens,
                "tokens_out": resp.usage.completion_tokens,
            }
            return data
        except Exception as e:
            last = e
            if attempt < MAX_RETRY - 1:
                if type(e).__name__ == "RateLimitError":
                    # API 오류 문구의 'try again in 12.3s'를 우선 사용한다.
                    m = re.search(r"try again in\s*([\d.]+)\s*s", str(e), re.I)
                    wait = min(float(m.group(1)) + 1, 60) if m else 10 * (attempt + 1)
                    time.sleep(wait)
                else:
                    time.sleep(1.5 * (attempt + 1))
    hint = diagnose()
    raise RuntimeError(hint or f"분석에 실패했습니다.\n({type(last).__name__})")


# ─────────────────────────────────────────────
# 후처리: 추출값 -> 장부에 넣을 형태
# ─────────────────────────────────────────────
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
    """10단위 내림. 보정이 있었으면 확인 대상."""
    if length is None:
        return None, False
    r = (int(length) // 10) * 10
    return r, r != int(length)


def color_text(kind, code):
    """장부 색상 열 문자열. 종류 표기는 투코드일 때 생략."""
    part = "" if kind == "투코드" else f"{kind} "
    return f"B {part}{code}".replace("  ", " ")


def ledger_color(item):
    """B [L18-]{종류} {코드}; 믹스는 이름과 전체 조합을 보존."""
    code = item.get("품목코드") or ""
    kind = item.get("종류") or "투코드"
    if item.get("_mix_name") and item.get("_mix_codes"):
        code = f"{item['_mix_name']} ({item['_mix_codes']})"
    if item.get("타입") == "L자":
        return f"B L18-{kind} {code}".strip()
    if kind == "투코드":
        return f"B {code}".strip()
    return f"B {kind} {code}".strip()


if __name__ == "__main__":
    import sys
    paths = sys.argv[1:]
    if not paths:
        print("사용: python extract.py <이미지경로> [추가이미지...]")
        raise SystemExit(1)
    hint = Path(paths[0]).stem.split("_")[0]
    out = extract_order(paths, client_hint=hint if hint in
                        ("DI", "휴안", "M", "DU", "RT", "JO", "JL") else None)
    print(json.dumps(out, ensure_ascii=False, indent=2))
