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
import time
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI
from PIL import Image

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