"""클립보드 감시 — 캡처하면 자동으로 주문에 추가.

Windows의 Win+Shift+S 캡처와 Ctrl+V 붙여넣기를 모두 지원한다.
Pillow ImageGrab가 특정 Windows/원격/OneDrive 환경에서 클립보드 이미지를
놓치는 경우가 있어 PySide6 QClipboard를 보조 경로로 함께 사용한다.
"""
from __future__ import annotations

import hashlib
import io
from pathlib import Path

try:
    from PIL import Image, ImageGrab
    PIL_AVAILABLE = True
except Exception:
    Image = None
    ImageGrab = None
    PIL_AVAILABLE = False

AVAILABLE = PIL_AVAILABLE


def _from_file_list(paths):
    if not PIL_AVAILABLE:
        return None
    for p in paths or []:
        try:
            path = Path(str(p))
            if path.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp", ".bmp") and path.exists():
                with Image.open(path) as opened:
                    return opened.convert("RGB").copy()
        except Exception:
            continue
    return None


def _grab_pillow():
    if not PIL_AVAILABLE or ImageGrab is None:
        return None
    try:
        obj = ImageGrab.grabclipboard()
    except Exception:
        return None
    if obj is None:
        return None
    if isinstance(obj, list):
        return _from_file_list(obj)
    if hasattr(obj, "size"):
        try:
            return obj.convert("RGB") if obj.mode != "RGB" else obj.copy()
        except Exception:
            return obj
    return None


def _grab_qt():
    """PySide6 QClipboard 보조 경로. QApplication이 떠 있을 때만 사용한다."""
    if not PIL_AVAILABLE:
        return None
    try:
        from PySide6.QtCore import QBuffer, QIODevice
        from PySide6.QtWidgets import QApplication
        app = QApplication.instance()
        if app is None:
            return None
        cb = app.clipboard()
        mime = cb.mimeData()

        # 파일을 복사한 경우
        if mime is not None and mime.hasUrls():
            files = [u.toLocalFile() for u in mime.urls() if u.isLocalFile()]
            img = _from_file_list(files)
            if img is not None:
                return img

        qimg = None
        if mime is not None and mime.hasImage():
            try:
                qimg = mime.imageData()
            except Exception:
                qimg = None
        if qimg is None or not hasattr(qimg, "isNull") or qimg.isNull():
            try:
                qimg = cb.image()
            except Exception:
                qimg = None
        if qimg is None or not hasattr(qimg, "isNull") or qimg.isNull():
            return None

        buf = QBuffer()
        if not buf.open(QIODevice.OpenModeFlag.WriteOnly):
            return None
        try:
            if not qimg.save(buf, "PNG"):
                return None
            data = bytes(buf.data())
        finally:
            buf.close()
        with Image.open(io.BytesIO(data)) as opened:
            return opened.convert("RGB").copy()
    except Exception:
        return None


def grab():
    """클립보드의 이미지를 (PIL Image, 해시)로 반환. 없으면 (None, None)."""
    img = _grab_pillow()
    if img is None:
        img = _grab_qt()
    if img is None:
        return None, None

    try:
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        data = buf.getvalue()
        return img, hashlib.md5(data).hexdigest()
    except Exception:
        return None, None


def save(img, out_dir, tag):
    """감시로 잡은 이미지를 파일로 저장하고 경로 반환."""
    d = Path(out_dir) / "_clip"
    d.mkdir(parents=True, exist_ok=True)
    p = d / f"clip_{tag}.png"
    img.save(p, format="PNG")
    return p


def too_small(img, min_side=100):
    """아이콘·아주 작은 조각을 발주서로 오인하지 않도록 한다."""
    return img is None or min(img.size) < min_side
