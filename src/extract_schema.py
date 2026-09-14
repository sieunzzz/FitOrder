"""
OpenAI Structured Outputs(strict) 규격에 맞춘 추출 스키마.

strict 모드 제약:
  - 모든 object 에 additionalProperties: false
  - 모든 property 를 required 에 포함 (선택 항목은 ["string","null"] 로)
"""

CLIENTS = ["DI", "SP", "휴안", "M", "DU", "RT", "JO", "JL", "DD",
           "유앤", "아지트", "WT", "인천)트루", "미래가공", "보노", "MS",
           "구미)경남", "창문애", "한길"]

_ITEM = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "제품군": {
            "type": "string",
            "enum": ["블라인드"],
            "description": "현재 홀딩도어 자동 판별은 사용하지 않는다. 모든 제작 주문은 블라인드로 추출",
        },
        "홀딩방식": {
            "type": ["string", "null"],
            "description": "현재 사용하지 않음. 반드시 null",
        },
        "홀딩레일": {
            "type": ["string", "null"],
            "description": "현재 사용하지 않음. 반드시 null",
        },
        "홀딩상하로라": {
            "type": "boolean",
            "description": "현재 사용하지 않음. 반드시 false",
        },
        "홀딩부속": {
            "type": ["string", "null"],
            "enum": ["레일연결부속", "라운드부속", None],
            "description": "현재 사용하지 않음. 반드시 null",
        },
        "홀딩부속색상": {
            "type": ["string", "null"],
            "enum": ["화이트", "블랙", None],
            "description": "현재 사용하지 않음. 반드시 null",
        },
        "품목코드": {
            "type": ["string", "null"],
            "description": "슬랫 색상 번호. 'WH102'→'102', 'IV200'→'200', "
                           "'아이보리 B200(IV)'→'200'. FP 등 접미사는 유지",
        },
        "색상원문": {
            "type": ["string", "null"],
            "description": "발주서에 적힌 색상 표기 그대로 (예: 'WH102', '아이보리')",
        },
        "타입": {
            "type": "string",
            "enum": ["C자", "L자"],
            "description": "'L자'/'L타입'/'L18' 표기가 실제로 보이면 L자, 없으면 C자. "
                           "색상 약어(GL, YL)나 '알루미늄'의 L 은 타입 표기가 아님",
        },
        "종류": {
            "type": "string",
            "enum": ["원코드", "투코드", "셔터"],
            "description": "명시 없으면 투코드",
        },
        "가로": {"type": ["number", "null"], "description": "cm. 사이즈 표기의 앞 숫자"},
        "세로": {"type": ["number", "null"], "description": "cm. 사이즈 표기의 뒤 숫자"},
        "수량": {
            "type": ["string", "null"],
            "description": "공지·부속 개수 등 특수 표기만. 일반 창은 null. "
                           "손잡이/모형 칸의 '½','1/3','1/4'는 반드시 "
                           "'1/2','1/3','1/4'로 적음. 창이 1개라는 뜻으로 '1'은 넣지 말 것",
        },
        "손잡이방향": {
            "type": ["string", "null"],
            "enum": ["좌", "우", None],
            "description": "발주서에 단일 '좌'/'우' 표기. 한 치수에 좌/우가 둘 다 있으면 이 필드는 null로 두고 좌개수/우개수/창개수로 표현",
        },
        "손잡이길이": {
            "type": ["integer", "null"],
            "description": "'줄120'/'손120' → 120. 표기가 없으면 반드시 null. 0을 넣지 말 것",
        },
        "연창": {"type": "boolean", "description": "'연창' 이라고 적혀 있으면 true"},
        "설치장소": {
            "type": ["string", "null"],
            "description": "방 위치. 예 '주방','작은방','드레스룸'",
        },
        "기재사항": {
            "type": ["string", "null"],
            "description": "이 행에만 해당하는 고객명·현장명·동호수. 설치장소는 제외",
        },
        "예외품목": {
            "type": ["string", "null"],
            "enum": ["수리", "롤스크린", "BMIX", "부속", "레일", None],
            "description": "일반 블라인드의 예외품목만. 홀딩도어/홀딩 전용 부속은 예외품목이 아님",
        },
        "창개수": {
            "type": ["integer", "null"],
            "description": "표의 수량 열 값 또는 한 치수에 적힌 방향 개수. 예: '77*88 좌 우' -> 창개수 2",
        },
        "좌개수": {
            "type": ["integer", "null"],
            "description": "손잡이 좌 열의 값. 빈 칸이면 0 또는 null. 추측해 채우지 말 것",
        },
        "우개수": {
            "type": ["integer", "null"],
            "description": "손잡이 우 열의 값. 빈 칸이면 0 또는 null. 추측해 채우지 말 것",
        },
        "원문": {"type": ["string", "null"], "description": "이 창에 해당하는 발주서 원문 한 줄"},
        "확신도": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "가로": {"type": "number"},
                "세로": {"type": "number"},
                "품목코드": {"type": "number"},
                "손잡이": {"type": "number"},
            },
            "required": ["가로", "세로", "품목코드", "손잡이"],
        },
    },
    "required": [
        "제품군", "홀딩방식", "홀딩레일", "홀딩상하로라", "홀딩부속", "홀딩부속색상",
        "품목코드", "색상원문", "타입", "종류", "가로", "세로", "수량",
        "손잡이방향", "손잡이길이", "연창", "설치장소", "기재사항",
        "예외품목", "창개수", "좌개수", "우개수", "원문", "확신도",
    ],
}

_DELIVERY = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "방식": {
            "type": ["string", "null"],
            "enum": ["택배", "화물", "내사", None],
            "description": "'내사픽업'→내사, 택배 주소 있으면 택배, 화물 지점명 있으면 화물",
        },
        "화물지점": {"type": ["string", "null"]},
        "주소": {"type": ["string", "null"]},
        "수령인": {"type": ["string", "null"]},
        "연락처": {"type": ["string", "null"]},
        "선불착불": {"type": ["string", "null"], "enum": ["선불", "착불", None]},
        "전달사항": {"type": ["string", "null"]},
    },
    "required": ["방식", "화물지점", "주소", "수령인", "연락처", "선불착불", "전달사항"],
}

EXTRACTION_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "거래처": {
            "type": ["string", "null"],
            "enum": CLIENTS + [None],
            "description": "발주서 제목/로고/발신처로 판별. 두창·두창블라인드는 DU. 불확실하면 null",
        },
        "발주일": {"type": ["string", "null"], "description": "YYYY-MM-DD"},
        "주문번호": {
            "type": ["string", "null"],
            "description": "거래처 주문번호. 표기가 없으면 null",
        },
        "고객명": {"type": ["string", "null"], "description": "발주서 상단 '고객명' 값"},
        "items": {"type": "array", "items": _ITEM},
        "배송": _DELIVERY,
        "전체원문": {
            "type": ["string", "null"],
            "description": "캡처에서 읽은 주문 관련 텍스트 전문",
        },
        "전체기재사항": {
            "type": ["string", "null"],
            "description": "표 하단 비고의 공통 지시와 주문번호. "
                           "'피스 동봉해주세요'→'피스', '( 기재 : 임지애 DW )'→JL은 '임지애'(DW 제거), "
                           "'주문 번 호 : 226095'→'226095'. 여러 개면 / 로 연결. "
                           "포장 관련 지시(비닐포장·걷비닐)는 제외",
        },
        "변경요청": {
            "type": "boolean",
            "description": "기존 주문의 수정을 요청하는 내용이면 true. "
                           "'주문 변경되나요', '사이즈 이걸로 변경', '수정 부탁', "
                           "'취소', '아까 그거', '어제 주문한 것' 같은 표현이 있으면 true",
        },
        "변경문구": {
            "type": ["string", "null"],
            "description": "변경을 요청한 문장 원문. 변경요청이 false 면 null",
        },
    },
    "required": [
        "거래처", "발주일", "주문번호", "고객명", "items", "배송",
        "전체원문", "전체기재사항", "변경요청", "변경문구",
    ],
}

RESPONSE_FORMAT = {
    "type": "json_schema",
    "json_schema": {
        "name": "blind_order",
        "strict": True,
        "schema": EXTRACTION_SCHEMA,
    },
}
