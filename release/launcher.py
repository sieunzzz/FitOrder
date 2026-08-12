"""
FitOrder 실행 진입점 — PyInstaller 로 exe 로 묶는 대상.

streamlit 을 하위 프로세스로 띄우고 브라우저를 앱 모드로 연다.
콘솔 창을 숨기려면 build.bat 의 --noconsole 옵션을 사용한다.
"""
import os
import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

PORT = 8501
APP = "app.py"


def base_dir() -> Path:
    """exe 로 묶인 경우와 스크립트 실행을 모두 처리"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def free_port(start: int) -> int:
    for p in range(start, start + 20):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start


def wait_ready(port: int, timeout: float = 60.0) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", port)) == 0:
                return True
        time.sleep(0.4)
    return False


def open_browser(url: str) -> None:
    """크롬·엣지가 있으면 앱 모드로, 없으면 기본 브라우저로"""
    pf = [os.environ.get("PROGRAMFILES", r"C:\Program Files"),
          os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
          os.environ.get("LOCALAPPDATA", "")]
    rel = [r"Google\Chrome\Application\chrome.exe",
           r"Microsoft\Edge\Application\msedge.exe"]
    for root in pf:
        for r in rel:
            exe = Path(root) / r
            if exe.exists():
                subprocess.Popen([str(exe), f"--app={url}",
                                  "--window-size=1600,950"])
                return
    webbrowser.open(url)


def main() -> int:
    root = base_dir()
    src = root / "src"
    app = src / APP
    if not app.exists():
        print(f"[오류] {app} 을 찾을 수 없습니다.")
        input("엔터를 누르면 종료합니다.")
        return 1

    port = free_port(PORT)
    url = f"http://localhost:{port}"

    # 환경변수로도 걸어둔다 — set_option 이 막히는 버전에 대한 보험
    for k, v in {
        "STREAMLIT_SERVER_HEADLESS": "true",
        "STREAMLIT_SERVER_PORT": str(port),
        "STREAMLIT_SERVER_ADDRESS": "127.0.0.1",
        "STREAMLIT_BROWSER_SERVER_PORT": str(port),
        "STREAMLIT_BROWSER_GATHER_USAGE_STATS": "false",
        "STREAMLIT_GLOBAL_DEVELOPMENT_MODE": "false",
    }.items():
        os.environ[k] = v
    env = os.environ.copy()

    if getattr(sys, "frozen", False):
        # exe 안에서는 streamlit 을 같은 프로세스에서 구동한다.
        # 포트·헤드리스 설정은 config 로 강제해야 반영된다.
        import threading

        from streamlit import config as st_config
        from streamlit.web import bootstrap

        os.chdir(src)
        print(f"FitOrder 서버 시작 — {url}")
        for opt, val in (("server.port", port),
                         ("server.address", "127.0.0.1"),
                         ("server.headless", True),
                         ("browser.serverPort", port),
                         ("browser.gatherUsageStats", False),
                         ("global.developmentMode", False)):
            try:
                st_config.set_option(opt, val)
            except Exception:
                pass

        # 서버가 준비된 뒤에 브라우저를 연다
        threading.Thread(
            target=lambda: open_browser(url) if wait_ready(port) else
            print("[오류] 서버가 시간 내에 시작되지 않았습니다."),
            daemon=True,
        ).start()

        bootstrap.run(str(app), False, [], {})
        return 0

    cmd = [sys.executable, "-m", "streamlit", "run", str(app),
           "--server.headless", "true", "--server.port", str(port)]
    proc = subprocess.Popen(cmd, cwd=str(src), env=env)
    if wait_ready(port):
        open_browser(url)
    else:
        print("[오류] 서버가 시간 내에 시작되지 않았습니다.")
    try:
        proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
    return 0


if __name__ == "__main__":
    sys.exit(main())