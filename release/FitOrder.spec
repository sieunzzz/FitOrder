# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 스펙 — streamlit 은 데이터 파일과 동적 임포트가 많아
# collect_all 로 한 번에 수집한다.
from PyInstaller.utils.hooks import collect_all, copy_metadata

datas, binaries, hiddenimports = [], [], []
for pkg in ("streamlit", "altair", "pyarrow", "pandas", "openpyxl",
            "PIL", "openai", "dotenv", "narwhals", "pydeck", "pymupdf"):
    d, b, h = collect_all(pkg)
    datas += d; binaries += b; hiddenimports += h

for pkg in ("streamlit", "pandas", "openpyxl", "openai", "pillow",
            "python-dotenv", "altair", "pyarrow", "pymupdf"):
    try:
        datas += copy_metadata(pkg)
    except Exception:
        pass

hiddenimports += [
    "streamlit.web.bootstrap", "streamlit.runtime.scriptrunner.magic_funcs",
    "streamlit.web.cli", "pandas._libs.tslibs.base",
]

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=["matplotlib", "scipy", "tkinter", "test"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="FitOrder",
    debug=False,
    console=True,          # 오류 확인용. 숨기려면 False
    icon="fitorder.ico",
)
coll = COLLECT(
    exe, a.binaries, a.datas,
    strip=False, upx=False, name="FitOrder",
)
