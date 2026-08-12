# 개발 참고

## 새 거래처 추가

거래처를 추가할 때 손대야 하는 곳은 네 군데입니다.

### 1. 마스터 데이터

`data/master/fitorder_master.xlsx` 의 `거래처` 시트에 행을 추가합니다.
상호 · 관리코드 · 단가등급 · 단가컬럼 · 배송방식 · 내부표시 · 발주형태.

### 2. `rules.py`

```python
CLIENT_INFO = {
    "새거래처": ("관리코드", "출고I가", "배송방식", "(K)", "image"),
}
```

특수 규칙이 있으면 해당 집합에 추가합니다.

```python
NO_DITTO        = {"휴안"}          # 세로에 " 표기를 쓰지 않는 거래처
NO_MERGE_PLACE  = {"JO"}            # 설치장소를 기재사항에 합치지 않는 거래처
PACKING_CLIENTS = {"휴안"}          # 포장비용을 전표에 붙이는 거래처
```

### 3. `prompt.py`

섹션 14 에 양식 설명을 추가합니다. 표 구조, 열 의미, 특이 표기를 적습니다.

```
[새거래처] 양식 설명
  - 상단 어느 칸이 거래처인지
  - 표의 각 열이 무엇인지
  - 수량 열이 창 개수를 뜻하는지
  - 하단 비고를 어떻게 해석할지
```

### 4. 엑셀 발주라면 `parsers.py`

`parse_excel()` 에 분기를 추가하고 전용 파서를 만듭니다.
**헤더 이름이 파일마다 다를 수 있으므로 값을 보고 판단**하는 것이 안전합니다.

### 5. 정확도 확인

발주서 3건과 정답 장부를 `data/samples/새거래처/` 에 넣고 측정합니다.

```bash
cd src
python evaluate.py ../data/samples 새거래처
```

---

## 프롬프트를 수정했을 때

캐시를 지우고 다시 측정해야 합니다.

```bash
cd src
rm -rf _cache                # Windows: Remove-Item -Recurse -Force _cache
python evaluate.py ../data/samples
```

거래처를 하나씩 돌리는 것이 원인 파악에 유리합니다.
여러 거래처를 한 번에 돌리면 오류가 겹쳐 해석이 어렵습니다.

---

## 검증 규칙 추가

`rules.py` 의 `validate()` 에 조건을 추가하고 등급을 정합니다.

```python
if 조건:
    out.append(("red", "규칙코드", "담당자에게 보일 메시지", 행번호))
```

| 등급 | 기준 |
| --- | --- |
| `red` | 이 값으로 제작하면 손실이 발생한다 |
| `review` | 중복이거나 기존 주문의 변경으로 보인다 |
| `yellow` | 맞을 수도 있지만 사람 눈이 한 번 닿아야 한다 |
| `info` | 장부 대상이 아니다 |

**`red` 를 남발하지 않는 것이 중요합니다.** 하루 400건에서 차단이 자주 걸리면
담당자가 습관적으로 해제하게 되고, 그때부터 안전장치가 아니게 됩니다.

---

## 문서 서식을 바꿀 때

`output.py` 상단의 상수를 먼저 확인합니다.

```python
FONT_DATA  = "THE행복열매"      # 데이터 글꼴
FONT_UI    = "맑은 고딕"        # 헤더 글꼴
COL_W      = [...]              # 열 너비 12개
ALIGN      = [...]              # 열 정렬 12개
NO_LEFT / NO_RIGHT              # 세로 테두리를 그리지 않는 열
LEDGER_ROWS = 35                # 장부 한 시트 행 수
SHEET_ROWS  = 16                # 작업지시서 블록당 행 수
```

부분 색상은 `_rich()` 를 사용합니다. **모든 조각을 `TextBlock` 으로 만들어야**
색이 없는 부분이 기본 글꼴로 떨어지지 않습니다.

```python
def _tb(text, color=None, sz=14):
    return TextBlock(InlineFont(rFont=FONT_DATA, sz=sz, color=color), text)
```

---

## 코드 스타일

- 표준 라이브러리 → 서드파티 → 로컬 순서로 임포트
- 함수에는 한 줄 독스트링. 왜 그렇게 했는지를 적습니다
- 도메인 규칙은 상수나 딕셔너리로 분리해 코드 본문에 매직 넘버를 두지 않습니다
- 금액 계산에는 반드시 `Decimal` 을 사용합니다

---

## 테스트

자동화된 단위 테스트는 없습니다. 대신 정답셋 대조로 회귀를 확인합니다.

```bash
cd src
python evaluate.py ../data/samples
```

엑셀 파서를 수정했으면 전체 파일을 한 번 돌려 오류가 없는지 봅니다.

```python
from parsers import parse_excel
from rules import Master, validate
from output import to_rows, build_ledger, build_worksheet, build_erp

M = Master("../data/master/fitorder_master.xlsx")
orders = []
for p in 엑셀_파일_목록:
    orders += parse_excel(p, 거래처)

rows = []
for o in orders:
    rows.extend(to_rows(o, "월", M))

red = sum(len([x for x in validate(o, M) if x[0] == "red"]) for o in orders)
print(f"주문 {len(orders)}건 / 장부 {len(rows)}행 / 오류 {red}")

build_ledger(rows, "/tmp/a.xlsx")
build_worksheet(rows, "/tmp/b.xlsx")
build_erp(orders, M, "/tmp/c.xlsx")
```
