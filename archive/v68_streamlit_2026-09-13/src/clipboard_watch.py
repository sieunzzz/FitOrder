"""
클립보드 감시 — 캡처하면 자동으로 주문에 추가

담당자 동작:  카톡에서 Win+Shift+S 로 캡처  ->  자동 추가
Streamlit 서버와 클라이언트가 같은 PC 여야 동작한다 (로컬 실행 전제).
"""
import hashlib
import io
from pathlib import Path

try:
    from PIL import ImageGrab
    AVAILABLE = True
except Exception:                       # Pillow 없거나 지원 안 되는 OS
    AVAILABLE = False


def grab():
    """클립보드의 이미지를 (PIL Image, 해시) 로 반환. 없으면 (None, None)"""
    if not AVAILABLE:
        return None, None
    try:
        obj = ImageGrab.grabclipboard()
    except Exception:
        return None, None

    img = None
    if obj is None:
        return None, None
    if isinstance(obj, list):           # 파일 복사 -> 경로 목록
        for p in obj:
            if str(p).lower().endswith((".png", ".jpg", ".jpeg")):
                try:
                    from PIL import Image
                    img = Image.open(p)
                    break
                except Exception:
                    continue
    elif hasattr(obj, "size"):
        img = obj
    if img is None:
        return None, None

    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    data = buf.getvalue()
    return img, hashlib.md5(data).hexdigest()


def save(img, out_dir, tag):
    """감시로 잡은 이미지를 파일로 저장하고 경로 반환"""
    d = Path(out_dir) / "_clip"
    d.mkdir(parents=True, exist_ok=True)
    p = d / f"clip_{tag}.png"
    img.save(p, format="PNG")
    return p


def too_small(img, min_side=200):
    """아이콘·작은 조각을 발주서로 오인하지 않도록"""
    return img is None or min(img.size) < min_side
