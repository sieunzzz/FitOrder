"""사람이 작성한 기존 장부 / FitOrder 장부 -> 주문 데이터 역변환.

2026-09-13 뼈대 정리: v68 output.py 에서 코드 변경 없이 분리.
"""
import re
from pathlib import Path
import openpyxl
from rules import CLIENT_INFO, di_mix_info
from holding import (
    HOLDING_FEATURE_ENABLED, accessory_qty, holding_accessory_from_text,
    holding_product_from_text, looks_like_holding_operation, normalize_holding_operation)

from ledger_rows import _true_accessory_catalog, expand_same_size_directions
from order_text import _jl_customer_name, _jo_clean_note_text, _split_note_parts, JL_PACKAGING


def _ledger_client(value):
    """장부 상호 셀에서 내부표시를 제외한 FitOrder 거래처를 찾는다."""
    text = str(value or "").strip()
    aliases = {"두창": "DU", "두창블라인드": "DU",
               "대일": "DI", "대일산업": "DI", "루임트": "RT",
               "대동산업": "DD", "대동": "DD", "유앤아이티엔에스": "유앤",
               "윈도우투모로우": "WT", "트루갤러리": "인천)트루",
               "미성텍스": "MS", "이끌림": "보노"}
    for name in sorted(CLIENT_INFO, key=len, reverse=True):
        if re.match(rf"^{re.escape(name)}(?:\s|\(|$)", text):
            return name
    for name, code in aliases.items():
        if text.startswith(name):
            return code
    return None


def _ledger_number(value):
    if value in (None, ""):
        return None
    try:
        number = float(str(value).replace(",", "").strip())
        return int(number) if number.is_integer() else number
    except (TypeError, ValueError):
        return None


def read_ledger(source, M=None, filename=None):
    """사람이 작성한 기존 장부 또는 FitOrder 장부를 작업지시서/EDI 데이터로 읽는다.

    `장부`, `장부 (2)` 시트를 우선한다. 그런 이름이 없으면 상호/색상/규격 헤더가
    있는 첫 실제 장부 시트를 자동으로 찾는다. 홀딩도어는 H 표기를 공식 품목으로 복원한다.
    """
    suffix = Path(filename or getattr(source, "name", "")).suffix.lower()

    def find_header(get_cell, max_row):
        for r in range(1, min(max_row, 12) + 1):
            vals = [re.sub(r"\s+", "", str(get_cell(r, c) or "")) for c in range(1, 14)]
            if "상호" in vals and any(v.startswith("색") for v in vals) \
                    and any(v.startswith("규격") for v in vals):
                return r
        return None

    if suffix == ".xls":
        try:
            import xlrd
        except ImportError as e:
            raise RuntimeError("기존 .xls 장부 처리 모듈이 없습니다.") from e
        contents = source.getvalue() if hasattr(source, "getvalue") else None
        book = xlrd.open_workbook(filename=None if contents is not None else str(source),
                                  file_contents=contents)

        def make_info(ws):
            def get_cell(row, col):
                if row < 1 or col < 1 or row > ws.nrows or col > ws.ncols:
                    return None
                return ws.cell_value(row - 1, col - 1)
            hr = find_header(get_cell, ws.nrows)
            return (ws.name, ws.nrows, get_cell, hr) if hr else None

        preferred = [ws for ws in book.sheets()
                     if ws.name == "장부" or re.fullmatch(r"장부 \(\d+\)", ws.name)]
        infos = [make_info(ws) for ws in preferred]
        sheets = [x for x in infos if x]
        if not sheets:
            # FitOrder DI 장부는 C 원코드/C 투코드/...처럼 여러 실제 시트로
            # 나뉘므로 헤더가 있는 모든 장부 시트를 읽는다. 견본/샘플/현장은 제외한다.
            for ws in book.sheets():
                if re.search(r"견본|샘플|현장", ws.name):
                    continue
                info = make_info(ws)
                if info:
                    sheets.append(info)
    else:
        wb = openpyxl.load_workbook(source, data_only=True)

        def make_info(ws):
            get_cell = lambda row, col: ws.cell(row, col).value
            hr = find_header(get_cell, ws.max_row)
            return (ws.title, ws.max_row, get_cell, hr) if hr else None

        preferred = [ws for ws in wb.worksheets
                     if ws.title == "장부" or re.fullmatch(r"장부 \(\d+\)", ws.title)]
        infos = [make_info(ws) for ws in preferred]
        sheets = [x for x in infos if x]
        if not sheets:
            for ws in wb.worksheets:
                if re.search(r"견본|샘플|현장", ws.title):
                    continue
                info = make_info(ws)
                if info:
                    sheets.append(info)
    if not sheets:
        raise ValueError("상호/색상/규격이 있는 장부 시트를 찾지 못했습니다.")

    rows, orders, errors = [], [], []
    current = None
    prev_color = prev_height = None
    prev_di_recipient_ctx = None
    for sheet_name, max_row, get_cell, header_row in sheets:
        header = [str(get_cell(header_row, c) or "").strip() for c in range(1, 14)]
        detailed = "특이" in header
        for excel_row in range(header_row + 1, max_row + 1):
            values = [get_cell(excel_row, c) for c in range(1, 14)]
            client_cell, color_cell = values[1], values[2]
            if str(client_cell or "").replace(" ", "") in ("상호", "현장용"):
                current = None
                prev_color = prev_height = None
                prev_di_recipient_ctx = None
                continue
            width, height = values[3], values[5]
            qty, direction, length = values[6], values[7], values[8]
            if detailed:
                special, note1, note2 = values[9], values[10], values[11]
            else:
                special = "#" if "#" in str(length or "") else None
                note1, note2 = values[9], values[10]
            meaningful = [values[i] for i in (1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12)]
            if not any(v not in (None, "") for v in meaningful):
                continue

            # 인천)트루의 별도 부속행을 사람이 작성한 장부 -> EDI에서도 복원한다.
            # 장부 표시명은 현장용이고, EDI는 공식 품명/관리코드로 변환한다.
            if current is not None and current.get("거래처") == "인천)트루" \
                    and height in (None, "") and width in (None, ""):
                acc = _true_accessory_catalog(color_cell)
                if acc:
                    acc["수량"] = 1
                    current.setdefault("_true_accessories", [])
                    if not any(x.get("표시") == acc.get("표시") for x in current["_true_accessories"]):
                        current["_true_accessories"].append(acc)
                    rows.append({"_특수": "부속", "문구": acc.get("표시"), "수량": 1})
                    continue

            # 루임트 MIX 조합의 길이 보조행(예: 030 / )128cm, 990 / )32cm)은
            # 별도 제품이 아니다. 바로 위 `B 원코드 030+990`의 구성 정보이므로 건너뛴다.
            if current is not None and current.get("거래처") == "RT" \
                    and height in (None, "") and color_cell in (None, ""):
                width_token = re.sub(r"\.0$", "", str(width or "").strip())
                qty_token = str(qty or "").strip()
                if re.fullmatch(r"\d{2,3}", width_token) and re.search(r"cm", qty_token, re.I):
                    continue

            # 주소/수령인/전달 등 장부 특수행. 사람이 작성한 장부를 EDI로
            # 되돌릴 때 배송 정보도 같이 복원한다. ☆행은 주소가 비어 있고
            # 우측에 선불만 적힌 경우도 있으므로 width가 문자열일 때만으로 제한하지 않는다.
            marker = str(color_cell or "").strip()
            is_special_row = (height in (None, "") and marker in ("#", "☆", "") and
                              (isinstance(width, str) or marker == "☆"))
            if is_special_row:
                text = str(width or "").strip()
                rows.append({"_특수": marker, "문구": width})
                if current is not None:
                    # 사람이 작성한 장부의 `# 포장비용`도 장부→EDI에서 잃지 않는다.
                    if marker == "#" and re.sub(r"\s+", "", text) == "포장비용":
                        parts = [x.strip() for x in str(current.get("전체기재사항") or "").split("/") if x.strip()]
                        if "포장비용" not in parts:
                            parts.append("포장비용")
                        current["전체기재사항"] = "/".join(parts)
                    delivery = current.setdefault("배송", {})
                    prepaid = str(note2 or note1 or "").strip()
                    if prepaid in {"선불", "착불"}:
                        delivery["선불착불"] = prepaid
                    if marker == "☆":
                        if text:
                            delivery["주소"] = text
                            default_delivery = CLIENT_INFO.get(current.get("거래처"), (None, None, None))[2]
                            if default_delivery in {"택배", "화물", "배달"}:
                                delivery["방식"] = default_delivery
                            elif not delivery.get("방식"):
                                delivery["방식"] = "택배"
                            if current.get("거래처") == "보노" and not delivery.get("선불착불"):
                                delivery["선불착불"] = "착불"
                            # 이 주소 행까지가 현재 수령인의 주문 묶음. 다음 제품행은
                            # 상호가 공란이어도 같은 거래처의 새 주문으로 시작한다.
                            current["_split_next_product"] = True
                    elif marker == "":
                        # 전달/발신 행에 전화번호가 있어도 수령인으로 오인하지 않는다.
                        if text.startswith("전달"):
                            delivery["전달사항"] = text.split(":", 1)[-1].strip()
                        elif text.startswith("발신"):
                            delivery["발신"] = text.split(":", 1)[-1].strip()
                        else:
                            phone = re.search(r"(01\d[- ]?\d{3,4}[- ]?\d{4})", text)
                            if phone:
                                number = phone.group(1).replace(" ", "")
                                receiver = text[:phone.start()].strip(" /,-")
                                delivery["연락처"] = number
                                if receiver:
                                    delivery["수령인"] = receiver
                                    current["고객명"] = receiver
                                    if current.get("거래처") == "보노" and "보노" in receiver.replace(" ", ""):
                                        current["_bono_receiver_is_bono"] = True
                continue

            # ☆ 주소행을 지난 뒤 상호가 공란인 다음 제품은 같은 거래처의
            # 새 수령인 주문으로 분리한다. (휴안 정답 장부 형태)
            if client_cell in (None, "") and current is not None \
                    and current.pop("_split_next_product", False):
                client = current.get("거래처")
                current = {"거래처": client, "주문번호": None, "고객명": None,
                           "전체기재사항": None, "배송": {}, "items": []}
                orders.append(current)
                prev_color = prev_height = None
                prev_di_recipient_ctx = None

            if client_cell not in (None, "") and str(client_cell).strip():
                raw_client_mark = str(client_cell).strip()
                client = _ledger_client(client_cell)
                marker_only = False
                marker_order_no = None
                if not client and current is not None:
                    # 같은 업체를 연속 작성할 때 상호 대신 (K)/(F) 같은 내부표시만
                    # 쓰거나, SP/JL은 주문번호만 쓰는 장부를 다시 읽을 수 있게 한다.
                    paren_mark = re.fullmatch(r"\(([A-Za-z0-9-]+)\)", raw_client_mark)
                    if paren_mark:
                        client = current.get("거래처")
                        marker_only = True
                        if current.get("거래처") == "JL" and paren_mark.group(1).upper() != "K":
                            marker_order_no = paren_mark.group(1)
                    elif current.get("거래처") in {"SP", "JL"} and \
                            re.fullmatch(r"[A-Za-z0-9-]+", raw_client_mark):
                        client = current.get("거래처")
                        marker_only = True
                        marker_order_no = raw_client_mark
                if not client:
                    errors.append(f"{sheet_name} {excel_row}행: 거래처 '{client_cell}'를 찾을 수 없습니다.")
                    current = None
                    prev_color = prev_height = None
                    continue
                current = {"거래처": client, "주문번호": marker_order_no, "고객명": None,
                           "전체기재사항": None, "배송": {}, "items": []}
                orders.append(current)
                prev_color = prev_height = None
                prev_di_recipient_ctx = None
            if current is None:
                errors.append(f"{sheet_name} {excel_row}행: 주문의 첫 행에 상호가 없습니다.")
                continue

            raw_color = str(color_cell or "").strip()
            color_text = raw_color or str(prev_color or "").strip()
            if HOLDING_FEATURE_ENABLED:
                accessory_hit = holding_accessory_from_text(color_text)
                product_hit = None if accessory_hit else holding_product_from_text(color_text)
                is_holding = bool(accessory_hit or product_hit)
            else:
                # 홀딩 자동 판별 OFF: H/홀딩처럼 보여도 일반 장부 행으로 계속 처리한다.
                accessory_hit = product_hit = None
                is_holding = False

            if str(height or "").strip() in ('"', '”'):
                height = prev_height
            elif height not in (None, ""):
                prev_height = height

            if is_holding:
                if accessory_hit:
                    # 부속은 앞 본품의 색상 반복 기준을 바꾸지 않는다.
                    aq = accessory_qty(qty)
                    item = {
                        "품목코드": accessory_hit.get("코드"), "색상원문": color_text,
                        "종류": "투코드", "타입": "C자", "가로": None, "세로": None,
                        "수량": qty if qty not in (None, "") else aq,
                        "손잡이방향": None, "손잡이길이": None, "연창": False,
                        "설치장소": note2, "기재사항": note1, "예외품목": None,
                        "창개수": None, "좌개수": None, "우개수": None,
                        "_product_group": "holding", "_holding_accessory": True,
                        "_holding_product_name": accessory_hit["품명"],
                        "_holding_accessory_label": accessory_hit.get("장부표시"),
                        "_holding_accessory_color": accessory_hit.get("부속색상") or "화이트",
                        "_ledger_prefix": "H",
                    }
                    if current.get("거래처") == "DI":
                        item["_DIrecipient_context"] = prev_di_recipient_ctx
                    if M is not None and not M.find_order_item(item, current["거래처"]):
                        errors.append(f"{sheet_name} {excel_row}행: 등록되지 않은 홀딩 부속 {color_text}")
                    current["items"].append(item)
                    rows.append({"거래처": current["거래처"] if len(current["items"]) == 1 else None,
                                 "내부표시": None, "색상": color_text,
                                 "가로": None, "세로": None, "수량": qty,
                                 "모형1": None, "모형2": None, "특이": special,
                                 "기재사항": note1, "기재사항2": note2, "출고일": None,
                                 "_같은색": False, "_특수": None,
                                 "_product_group": "holding"})
                    continue

                prev_color = raw_color or prev_color
                w, h = _ledger_number(width), _ledger_number(height)
                if w is None or h is None:
                    errors.append(f"{sheet_name} {excel_row}행: 홀딩도어 가로·세로 규격을 확인해 주세요.")
                op_source = direction if looks_like_holding_operation(direction) else \
                    (qty if looks_like_holding_operation(qty) else None)
                operation = normalize_holding_operation(op_source)
                rail = str(length or "").strip() if "레일" in str(length or "") else None
                raw_qty = None if looks_like_holding_operation(qty) else qty
                try:
                    count = int(float(raw_qty)) if raw_qty not in (None, "") else 1
                except (TypeError, ValueError):
                    count = 1
                di_ctx = None
                if current.get("거래처") == "DI":
                    visible_recipient = str(note1 or "").strip()
                    if visible_recipient:
                        prev_di_recipient_ctx = visible_recipient
                    di_ctx = prev_di_recipient_ctx
                item = {
                    "품목코드": product_hit.get("코드") if product_hit else None,
                    "색상원문": color_text, "종류": "투코드", "타입": "C자",
                    "가로": w, "세로": h, "수량": raw_qty,
                    "손잡이방향": None, "손잡이길이": None, "연창": False,
                    "설치장소": note2, "기재사항": note1, "예외품목": None,
                    "창개수": count, "좌개수": None, "우개수": None,
                    "_수동특이": re.sub(r"(?:#|틀안)", " ", str(special or "")).strip() or None,
                    "_product_group": "holding", "_ledger_prefix": "H",
                    "_DIrecipient_context": di_ctx,
                    "_holding_product_name": product_hit.get("품명") if product_hit else None,
                    "_holding_label": product_hit.get("장부표시") if product_hit else None,
                    "_holding_operation": operation,
                    "_holding_rail": rail,
                    # +상하로라는 현재 행에 실제로 적혀 있을 때만 적용한다.
                    "_holding_upper_roller": "+상하로라" in raw_color.replace(" ", ""),
                }
                if not product_hit:
                    errors.append(f"{sheet_name} {excel_row}행: 홀딩도어 품목 '{color_text}'를 찾지 못했습니다.")
                elif M is not None and not M.find_order_item(item, current["거래처"]):
                    errors.append(f"{sheet_name} {excel_row}행: 등록되지 않은 홀딩도어 품목 {color_text}")
                current["items"].append(item)
                rows.append({"거래처": current["거래처"] if len(current["items"]) == 1 else None,
                             "내부표시": None, "색상": raw_color or None,
                             "가로": w, "세로": ('"' if str(values[5] or "").strip() in ('"','”') else h),
                             "수량": raw_qty, "모형1": operation, "모형2": rail,
                             "특이": special, "기재사항": note1, "기재사항2": note2,
                             "출고일": None, "_같은색": not bool(raw_color), "_특수": None,
                             "_product_group": "holding"})
                continue

            # ── 기존 블라인드 장부 ──
            if current.get("거래처") == "JO":
                # 사람이 작성한 JO 장부: 기재사항1=주문번호, 기재사항2=나머지 메모.
                if note1 not in (None, ""):
                    jo_no = str(note1).strip()
                    if jo_no:
                        current["주문번호"] = jo_no
                note1 = _jo_clean_note_text(note2)
                note2 = None
            elif current.get("거래처") == "보노":
                # 보노 장부 역변환: 기재사항1=받는사람/화물지점,
                # 기재사항2=오더명/시공위치.
                if note1 not in (None, ""):
                    current["_bono_note1"] = str(note1).strip()
                note2_parts = [x.strip() for x in str(note2 or "").split("/") if x.strip()]
                if note2_parts:
                    if not current.get("주문번호"):
                        current["주문번호"] = note2_parts[0]
                    note2 = "/".join(note2_parts[1:]) or None
                note1 = None
            if raw_color:
                prev_color = raw_color
            code, kind, type_ = parse_ledger_color(color_text, note1)
            mix_name, mix_codes = parse_ledger_mix(color_text, note1)
            if mix_name == "MIX" and mix_codes:
                # 사람이 작성한 장부의 B MIX/B 원코드 MIX는 조합을 기재사항에서 읽고,
                # EDI 수령인/메모에는 조합 문자열을 중복시키지 않는다.
                note_parts = [x.strip() for x in str(note1 or "").split("/") if x.strip()]
                note1 = "/".join(x for x in note_parts if x.replace(" ", "") != mix_codes) or None
            w, h = _ledger_number(width), _ledger_number(height)
            if not code:
                errors.append(f"{sheet_name} {excel_row}행: 품목 '{color_text}'를 해석할 수 없습니다.")
            if w is None or h is None:
                errors.append(f"{sheet_name} {excel_row}행: 가로·세로 규격을 확인해 주세요.")

            if current.get("거래처") == "JL":
                # JL 장부의 `피스/강종민` 같은 병합 기재사항에서 고객명을 복원한다.
                # 두 번째 이후 행의 기재사항이 병합으로 비어 있어도 첫 행에서 저장한
                # 고객명을 현재 주문 전체에 사용한다.
                for part in _split_note_parts(note1):
                    candidate = _jl_customer_name(part)
                    if not candidate or "피스" in candidate or JL_PACKAGING.search(candidate):
                        continue
                    if re.fullmatch(r"손\d{2,3}", candidate):
                        continue
                    if candidate not in {"틀안", "포장X", "브라켓"}:
                        current["고객명"] = candidate
                        break

            item = {"품목코드": code, "종류": kind, "타입": type_,
                    "가로": w, "세로": h, "수량": qty or 1,
                    "손잡이방향": (str(direction).strip()
                                    if re.fullmatch(r"[좌우]{1,10}", str(direction or "").strip())
                                    else None),
                    "손잡이길이": parse_handle(length)[0],
                    "연창": "#" in str(special or ""),
                    "_수동특이": re.sub(r"(?:#|틀안)", " ",
                                      str(special or "")).strip() or None,
                    "설치장소": note2, "기재사항": note1,
                    "예외품목": None}
            if mix_name and mix_codes:
                item["_mix_name"], item["_mix_codes"] = mix_name, mix_codes
                if mix_name == "MIX":
                    item["_generic_mix"] = True
            if code and M is not None and not M.find_order_item(item, current["거래처"]):
                errors.append(f"{sheet_name} {excel_row}행: 마스터에 없는 품목 {color_text}")
            current["items"].append(item)
            rows.append({"거래처": current["거래처"] if len(current["items"]) == 1 else None,
                         "내부표시": None, "색상": color_text,
                         "가로": w, "세로": h, "수량": qty,
                         "모형1": item["손잡이방향"], "모형2": length,
                         "특이": special, "기재사항": note1,
                         "기재사항2": note2, "출고일": None,
                         "_같은색": color_cell in (None, ""), "_특수": None})
    for _order in orders:
        expand_same_size_directions(_order)
    return rows, orders, errors


# ─────────────────────────────────────────────
# 편집 반영: 장부 표시값 -> 주문 항목
# ─────────────────────────────────────────────
def _ledger_mix_codes(*values):
    for value in values:
        m = re.search(r"(?<!\d)(\d{3}(?:\s*\+\s*\d{3})+)(?!\d)", str(value or ""))
        if m:
            return re.sub(r"\s+", "", m.group(1))
    return None


def parse_ledger_mix(text, note=None):
    """정식 MIX 또는 사람이 적은 `B MIX` + 기재사항 코드 조합을 복원."""
    s = str(text or "").strip()
    m = re.search(r"([가-힣A-Za-z0-9]+)\s*\(\s*(\d{3}(?:\s*\+\s*\d{3})+)\s*\)", s)
    if m:
        return m.group(1), re.sub(r"\s+", "", m.group(2))
    if re.search(r"\bMIX\b", s, re.I):
        return "MIX", _ledger_mix_codes(s, note)
    # 사용자가 코드 괄호를 생략해도 제공된 색상표에 있는 이름이면 복원한다.
    name_only = re.sub(r"^\s*[BHRC]\s+", "", s, flags=re.I)
    name_only = re.sub(r"^L(?:18|21)-(?:원코드|투코드|셔터)\s+", "", name_only, flags=re.I)
    name_only = re.sub(r"^(?:원코드|투코드|셔터)\s+", "", name_only)
    known = di_mix_info(name_only)
    return known if known else (None, None)


def parse_ledger_color(text, note=None):
    """장부 색상 문자열 -> (코드/믹스조합, 종류, 타입)."""
    s = str(text or "").strip()
    if not s:
        return None, None, None
    type_ = "L자" if re.search(r"L(18|21)-", s) else "C자"
    kind = "투코드"
    for k in ("원코드", "셔터", "투코드"):
        if k in s:
            kind = k
            break
    mix_name, mix_codes = parse_ledger_mix(s, note)
    if mix_name and mix_codes:
        return mix_codes, kind, type_
    m = re.search(r"(\d{3,4}[A-Za-z]*)\s*$", s)
    return (m.group(1) if m else None), kind, type_


def parse_handle(text):
    """'#손140' -> (140, True) / '손120' -> (120, False)"""
    s = str(text or "").strip()
    yeon = s.startswith("#")
    m = re.search(r"(\d+)", s)
    return (int(m.group(1)) if m else None), yeon
