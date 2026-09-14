# 복구/재구성 상태

## 2026-09-13 (최신) — 실제 사용 중인 PySide 버전 확보
- 사용자가 `Desktop\FitOrder_.zip`(PySide6, APP_VERSION `2026.09.06-108`)을 제공. 실제 운영 프로그램이다.
- `src/`, `data/master/`(블라인드·홀딩·롤콤비 품목장 포함), 실행/설치 bat를 해시 동일하게 복사. `.env`, `out/`은 제외.
- v68 Streamlit(09-11)과 비교: 108에만 롤·콤비/홀딩 전용 거래처/외부 품목장/시작화면/병합 UI가 있다.
  v68에만 있는 블라인드 기재사항 분배 변경(output.py 17개 함수 차이 등)은 블라인드 업체 작업 때 대조한다.
- v68 뼈대 정리본은 `archive/v68_streamlit_2026-09-13/`에 보관.
- 같은 날 v109 수정: README "2026-09-13 변경" 참고.

---

## 2026-09-13 업데이트 — v68 실제 소스 확보
- `Desktop\FitOrder_v68_src_windows_compatible\src`(APP_VERSION `2026.09.11-68`, Streamlit)를 발견해 `src/`에 해시 동일하게 복사했다.
- 포함: app.py, extract.py, extract_schema.py, prompt.py, parsers.py, rules.py, output.py, holding.py, clipboard_watch.py, evaluate.py, xlrd/, xlwt/
- 아직 없음: `data/master/fitorder_master.xlsx`(품목/단가 마스터), `data/logo.png`, 정답셋 샘플(`data/samples`)
- 아래 기록의 재구성본(PySide6)은 `archive/reconstructed_2026-09-13/`로 이동했다.

---
(이하 v68 확보 이전 기록)

## 파일 라이브러리에서 확인된 원본
- `FitOrder_v82_qt_app.py` — PySide6 UI v82
- `FitOrder_v79_output.py` — 출력 엔진 v79
- `FitOrder_clipboard_watch.py` — 클립보드 감시
- README 변경 기록 — v84 UI/규칙 변경사항
- 2026-08-30 인수인계 README — 당시 `src/`, `fitorder_master.json/CSV (원본 legacy는 xlsx)`, support, private_data가 포함된 전체 패키지였음을 확인
- 시연용 발주서 Excel 및 여러 테스트 장부/작업지시서 결과가 파일 라이브러리에 존재

## 제약
현재 세션에서는 과거 FULL ZIP 내부 파일 바이트를 직접 `/mnt/data`로 복사해 새 ZIP에 넣는 인터페이스가 없다.
따라서 이 패키지의 `src/fitorder`는:
1. 확인된 최신 모듈 구조,
2. 남아 있는 코드의 함수/입출력 구조,
3. 지금까지 대화에서 확정된 최신 규칙
을 기준으로 재구성했다.

## Claude에게
원본과 완전히 동일한 동작이 필요한 영역은 우선순위를 아래처럼 둔다.
1. 업체별 parser
2. v79 output의 경영박사 EDI 세부 열/품목코드/단가 계산
3. v84 UI 동기화
4. master 데이터의 실제 품목/단가 전체 복원
5. golden test 추가

현재 코드는 "수정 가능한 프로젝트 기준본"이며, 규칙 문서는 삭제하지 말고 구현을 점진적으로 맞춘다.
