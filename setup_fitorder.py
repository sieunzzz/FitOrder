from __future__ import annotations

import argparse
import importlib.util
import subprocess
import sys

# Import name -> pip requirement
REQUIRED = {
    "PySide6": "PySide6>=6.7",
    "pandas": "pandas>=2.0",
    "openpyxl": "openpyxl>=3.1",
    "PIL": "pillow>=10.0",
    "openai": "openai>=1.40",
    "dotenv": "python-dotenv>=1.0",
    "fitz": "pymupdf>=1.24",
    "xlrd": "xlrd==2.0.1",
    "xlwt": "xlwt==1.3.0",
}


def missing_requirements() -> list[str]:
    missing: list[str] = []
    for module_name, requirement in REQUIRED.items():
        if importlib.util.find_spec(module_name) is None:
            missing.append(requirement)
    return missing


def ensure_pip() -> bool:
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "--version"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return True
    except Exception:
        pass
    try:
        subprocess.check_call([sys.executable, "-m", "ensurepip", "--upgrade"])
        return True
    except Exception as exc:
        print(f"[ERROR] pip is not available: {exc}")
        return False


def install(requirements: list[str]) -> int:
    if not requirements:
        print("[OK] All required packages are already installed.")
        return 0
    if not ensure_pip():
        return 2

    print("[INFO] Missing packages:")
    for requirement in requirements:
        print(f"  - {requirement}")
    print("[INFO] Installing packages. This can take several minutes...")

    cmd = [
        sys.executable,
        "-m",
        "pip",
        "install",
        "--disable-pip-version-check",
        *requirements,
    ]
    rc = subprocess.call(cmd)
    if rc != 0:
        print("[INFO] Normal install failed. Retrying with --user...")
        rc = subprocess.call(cmd[:4] + ["--user"] + cmd[4:])
    if rc != 0:
        print(f"[ERROR] pip install failed with error code {rc}.")
        return rc

    still_missing = missing_requirements()
    if still_missing:
        print("[ERROR] These packages are still unavailable after installation:")
        for requirement in still_missing:
            print(f"  - {requirement}")
        return 3

    print("[OK] FitOrder dependencies are ready.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    missing = missing_requirements()
    if args.check:
        if missing:
            print("Missing: " + ", ".join(missing))
            return 1
        print("OK")
        return 0
    return install(missing)


if __name__ == "__main__":
    raise SystemExit(main())
