# 배포

세 가지 방식이 있습니다. 상황에 맞게 고르면 됩니다.

| 방식 | 대상 | 준비 시간 | Python 설치 |
| --- | --- | --- | --- |
| 개발 실행 | 개발자 | 5분 | 필요 |
| 포터블 | 원거리 사용자 | 30분 | **불필요** |
| 실행파일 | 배포·제출 | 30분 | **불필요** |

---

## 1. 개발 실행

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

cd src
streamlit run app.py
```

Windows 에서는 `run.bat` 을 실행하면 브라우저가 앱 모드로 열립니다.

---

## 2. 포터블 배포

Python 설치 없이 폴더 복사만으로 동작합니다.
원거리 사용자에게 전달할 때 가장 안전한 방식입니다.

### 준비 (개발 PC에서 1회)

**포터블 Python 받기**

https://github.com/astral-sh/python-build-standalone/releases

```
cpython-3.12.x-x86_64-pc-windows-msvc-install_only.tar.gz
```

`install_only` 가 붙은 것을 받습니다. pip 이 포함된 완전한 CPython 입니다.
python.org 의 embeddable package 는 pip 이 없고 경로 설정이 까다로워
권장하지 않습니다.

**배치**

압축을 풀면 나오는 `python` 폴더를 프로젝트 최상단에 둡니다.

```
FitOrder/
├─ python/
│   └─ python.exe
├─ src/
└─ setup_portable.bat
```

**패키지 설치**

```
setup_portable.bat
```

streamlit · pandas · openpyxl · pillow · openai · python-dotenv 를
포터블 Python 안에 설치합니다. 수 분 걸립니다.

### 전달

`run.bat` 만 더블클릭하면 실행됩니다.
폴더 전체를 압축해 전달하면 되고, 용량은 250 ~ 350 MB 입니다.

**함께 넣어야 하는 것**

- `.env` — API 키
- `data/master/fitorder_master.xlsx` — 마스터 데이터
- `THEHappyfruit.ttf` — 장부 글꼴

**빼야 하는 것**

- `data/samples/` — 고객 정보 포함
- `db/` `out/` `src/_cache/` — 실행 중 생성물

---

## 3. 실행파일 (exe)

### 빌드

```
cd release
build.bat
```

PyInstaller 설치와 빌드로 5 ~ 15분 걸립니다.
결과는 `release/dist/FitOrder/FitOrder.exe` 입니다.

`build.bat` 이 `src` · `data` · `.env` · 글꼴을 자동으로 복사합니다.

### 구조

```
dist/FitOrder/
├─ FitOrder.exe
├─ _internal/          런타임 · 라이브러리 (자동 생성)
├─ src/
├─ data/master/
├─ .env
└─ THEHappyfruit.ttf
```

**exe 파일만 전달해서는 동작하지 않습니다.** 폴더 전체를 전달해야 합니다.

### 동작 방식

exe 는 내부에서 로컬 웹 서버를 띄우고 브라우저를 앱 모드로 엽니다.

```
FitOrder.exe  →  localhost:8501 서버 시작  →  브라우저 앱 모드
```

클립보드 접근과 파일 시스템 연동이 필요해 로컬 실행 구조를 택했습니다.
거래 데이터가 외부로 나가지 않는 이점도 있습니다.

### 자주 나는 문제

| 증상 | 원인 | 해결 |
| --- | --- | --- |
| `Icon input file not found` | `fitorder.ico` 없음 | 파일을 `release/` 에 두거나 spec 에서 `icon=` 제거 |
| 브라우저가 다른 포트로 열림 | streamlit 이 config 를 무시 | `launcher.py` 가 환경변수와 `set_option` 양쪽으로 강제 |
| 연결 거부 | 서버 준비 전에 브라우저 열림 | `wait_ready()` 로 대기 후 열도록 처리 |
| 빈 화면 | 데이터 파일 수집 누락 | spec 의 `collect_all` 대상에 패키지 추가 |

빌드가 계속 실패하면 포터블 방식으로 전환하는 편이 빠릅니다.
제출 요건상 "실행파일" 은 `run.bat` + 포터블 Python 조합으로도 충족됩니다.

---

## 실행 환경 요구사항

- Windows 10 이상
- 인터넷 연결 — 이미지 추출용. 엑셀 발주는 오프라인 동작
- `THE행복열매` 글꼴 — 장부 서식 재현용

엑셀 파일에는 글꼴이 포함되지 않으므로, 문서를 열 PC 에 글꼴이 설치되어
있어야 서식이 정확히 보입니다. `run.bat` 이 최초 실행 시 설치를 안내합니다.

---

## API 키 관리

`.env` 에 평문으로 저장됩니다. 다음을 권합니다.

- **Monthly budget 설정** — platform.openai.com → Settings → Limits.
  키가 유출되어도 피해가 그 금액으로 제한됩니다
- 자동 충전(Auto recharge) 끄기
- 저장소에 `.env` 를 올리지 않기 (`.gitignore` 에 포함되어 있습니다)

키가 노출된 것으로 의심되면 즉시 폐기하고 새로 발급하세요.
