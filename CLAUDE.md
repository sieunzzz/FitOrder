# Claude 작업 지침 — FitOrder

이 저장소를 열면 가장 먼저 아래 순서대로 읽는다.

1. `README.md` (현재 구조·변경·알려진 한계)
2. `docs/RULES_MASTER.md`
   - `docs/RULES_HOLDING.md` (홀딩도어 7개 업체 — 코드에 없으면 "복구 누락"으로 판단)
3. `docs/RECOVERY_STATUS.md`
4. `src/qt_app.py` (OrderWorkspace), `src/desktop_workflow.py`
5. `src/output.py` (to_rows / _write_rows / build_erp / read_ledger)
6. `src/rules.py`, `src/holding.py`, `src/roll_combo.py`
7. `src/extract.py`, `src/parsers.py`

## 목적
FitOrder는 블라인드 제조 공장의 비정형 발주(카카오톡 이미지/메시지, Excel, PDF 등)를 표준 주문 데이터로 바꿔
1) 사전점검,
2) 장부,
3) 작업지시서,
4) 경영박사 ERP(EDI)
까지 이어 주는 발주 자동화 시스템이다.

## 기준 코드
- 기준은 **실제 사용 중인 PySide6 데스크톱 버전**(v108 → 2026-09-13 v109). 실행: `01_RUN_FITORDER.bat` → `src/qt_app.py`
- `archive/`의 v68 Streamlit·재구성본은 참고용이다. 수정·실행 대상이 아니다.

## 작업 원칙
- 업무 규칙은 UI보다 우선한다.
- `docs/RULES_MASTER.md`, `docs/RULES_HOLDING.md`의 확정 규칙을 임의로 삭제/단순화하지 않는다. 코드와 충돌하면 문서를 기준으로 보고 사용자에게 알린다.
- 거래처별 예외를 공통화하거나 옮길 때 출력 결과가 달라지지 않는지 golden 테스트로 확인한다. 비슷해도 업체끼리 합치지 않는다.
- AI는 추출/해석에 사용하고, 금액/수량/최소과금/색상표시/배송판정처럼 결정 가능한 로직은 규칙 기반으로 유지한다.
- 사용자가 사전점검에서 수정한 값이 최우선이다(`_manual_note1/2`, `_manual_handle_length` 등).
- 자동 병합 금지. 사용자가 요청한 병합/해제만 유지한다.
- 주소/부속/전달사항/손잡이길이는 선택행과 연동해 확인 가능해야 한다.
- 실제 개인정보/API 키를 커밋하지 않는다(`.env`, `out/` 제외).

## 사전점검 구조 규칙 (v109)
- **화면 자동 병합은 `output.ledger_merge_ranges`(= 장부 `_write_rows`) 결과만 쓴다.** qt_app에 별도 병합 규칙을 새로 만들지 않는다. 장부 병합 규칙을 바꾸면 화면에도 자동 반영된다. (특수행 가로 병합만 화면 편집용 예외)
- 수동 병합/해제는 행 고유 ID(`_order_uid`, `_uid` → `_row_key`)로 저장한다. 행 위치 번호로 저장하지 않는다.
- 병합된 셀 수정/삭제는 병합된 모든 제품 행에 반영한다.

## 수정 절차
1. `python -m pytest`로 golden 통과 확인.
2. 작은 단위로 수정 → 테스트.
3. 동작을 의도적으로 바꾼 경우에만 golden 갱신(`python tests/_generate_golden.py <항목>`)하고 바뀐 내용을 사용자에게 보고.
4. 원본이 평면 import(`from rules import ...`)와 `src/../data/master` 경로를 쓰므로 파일 위치 변경 시 두 가지를 함께 확인.
