"""클립보드 감시 — 캡처하면 자동으로 주문에 추가."""
import hashlib, io
from pathlib import Path

try:
    from PIL import ImageGrab
    AVAILABLE = True
except Exception:
    AVAILABLE = False

def grab():
    if not AVAILABLE:
        return None, None
    try:
        obj = ImageGrab.grabclipboard()
    except Exception:
        return None, None
    img = None
    if obj is None:
        return None, None
    if isinstance(obj, list):
        for p in obj:
            if str(p).lower().endswith((".png",".jpg",".jpeg")):
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
    if img.mode not in ("RGB","L"):
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    data = buf.getvalue()
    return img, hashlib.md5(data).hexdigest()

def save(img, out_dir, tag):
    d = Path(out_dir) / "_clip"
    d.mkdir(parents=True, exist_ok=True)
    p = d / f"clip_{tag}.png"
    img.save(p, format="PNG")
    return p

def too_small(img, min_side=200):
    return img is None or min(img.size) < min_side
