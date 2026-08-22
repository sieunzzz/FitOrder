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
                if client_hint == "SP":
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
            if data.get("거래처") == "MS":
                common = str(data.get("전체기재사항") or "")
                common = re.sub(
                    r"(?:오늘|내일|모레|월요일|화요일|수요일|목요일|금요일|토요일|일요일)?\s*"
                    r"출고\s*(?:부탁드립니다|부탁드려요|요청|희망)?", "", common)
                data["전체기재사항"] = common.strip(" /,.") or None
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
    """B [L18-]{종류} {코드}"""
    code = item.get("품목코드") or ""
    kind = item.get("종류") or "투코드"
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
