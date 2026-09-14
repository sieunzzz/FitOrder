"""경영박사 EDI(.xls) 전표 생성. 업체별 적요, 포장비, 배송행.

2026-09-13 뼈대 정리: v68 output.py 에서 코드 변경 없이 분리.
"""
import re
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from types import SimpleNamespace

from rules import ERP_CLIENT_NAME, clean_delivery_notice, calc_erp
from holding import (
    accessory_qty, holding_calc, holding_magnet_charge_qty, looks_like_holding_operation,
    normalize_holding_operation)

from clients import get_client_rules
from ledger_rows import (
    _blind_counts, _effective_delivery_mode, _holding_count, _needs_window_packing,
    _true_accessory_catalog, handle_split, synchronize_order_for_outputs, to_rows)
from ledger_book import sort_di_items_in_ledger_order


# ─────────────────────────────────────────────
# 4. 경영박사 전표
# ─────────────────────────────────────────────
ERP_HEAD = ["날짜", "전표번호", "계정코드", "계정", "거래처관리코드", "상호",
            "대체_코드", "대체_상호", "품목관리코드", "품명", "규격", "수량",
            "단가", "금액", "부가세", "전표적요", "사원코드", "사원"]


def _erp_text_width(text):
    """경영박사 인쇄 폭: 한글 2칸, 영문·숫자·기호 1칸."""
    return sum(2 if "\uac00" <= ch <= "\ud7a3" else 1 for ch in str(text))


def _split_erp_token(token, limit=30):
    """공백 없이 30칸을 넘는 토큰을 문자 경계에서 나눈다."""
    parts, cur, width = [], "", 0
    for ch in str(token):
        cw = 2 if "\uac00" <= ch <= "\ud7a3" else 1
        if cur and width + cw > limit:
            parts.append(cur)
            cur, width = "", 0
        cur += ch
        width += cw
    if cur:
        parts.append(cur)
    return parts


def wrap_erp_text(text, limit=30):
    """EDI ** 행의 적요를 인쇄 폭에 맞게 나눈다.

    `명지로 106`, `도원로 45`, `산호대로25길 44`처럼 도로명과
    바로 뒤의 건물번호는 하나의 단위로 취급해 서로 떨어지지 않게 한다.
    """
    tokens = str(text or "").strip().split()
    units, i = [], 0
    while i < len(tokens):
        token = tokens[i]
        road = token.rstrip(",")
        if (re.search(r"(?:로|길)$", road) and i + 1 < len(tokens)
                and re.match(r"^\d", tokens[i + 1])):
            units.append(token + " " + tokens[i + 1])
            i += 2
        else:
            units.append(token)
            i += 1

    expanded = []
    for unit in units:
        expanded.extend([unit] if _erp_text_width(unit) <= limit
                        else _split_erp_token(unit, limit))

    # 30칸을 그냥 앞에서부터 채우면 마지막 줄에 한 단어만
    # 남을 수 있다. 각 줄의 남는 폭 제곱합이 작아지도록
    # 동적 계획법으로 균형 있게 나눈다.
    n = len(expanded)
    best = [(float("inf"), None)] * (n + 1)
    best[n] = (0, n)
    for i in range(n - 1, -1, -1):
        line = ""
        for j in range(i, n):
            line = expanded[j] if j == i else f"{line} {expanded[j]}"
            width = _erp_text_width(line)
            if width > limit:
                break
            score = (limit - width) ** 2 + best[j + 1][0]
            if score < best[i][0]:
                best[i] = (score, j + 1)

    lines, i = [], 0
    while i < n:
        j = best[i][1] or i + 1
        lines.append(" ".join(expanded[i:j]))
        i = j
    return lines


def _erp_address_text(value):
    """경영박사 EDI 전용 택배/화물 주소 정리.

    - 번지/도로번호 뒤 쉼표 제거
    - 괄호에는 아파트/건물명만 남기고 상세 전달문구는 제거
    - 101동202호 -> 101-202
    장부 원문은 변경하지 않는다.
    """
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    if not text:
        return ""
    dongho = None
    m = re.search(r"(\d+)\s*동\s*(\d+)\s*호", text)
    if m:
        dongho = f"{m.group(1)}-{m.group(2)}"
        text = text[:m.start()] + text[m.end():]

    def clean_paren(m):
        inside = re.sub(r"\s+", " ", m.group(1)).strip()
        # 동/호, 쉼표, 배송메모 앞은 버리고 건물/아파트 이름만 유지한다.
        inside = re.split(r"\d+\s*동|\d+\s*호|[,;/]", inside, maxsplit=1)[0].strip()
        inside = re.sub(r"\s+(?:경비실|문앞|배송|연락|부재|세대).*", "", inside).strip()
        return f"({inside})" if inside else ""

    text = re.sub(r"\(([^()]*)\)", clean_paren, text)
    text = re.sub(r"(\d+(?:-\d+)?(?:번지)?)\s*,\s*", r"\1 ", text, count=1)
    text = re.sub(r"\s{2,}", " ", text).strip(" ,")
    if dongho and dongho not in text:
        text = f"{text} {dongho}".strip()
    return text


def build_erp(orders, M, path, order_date=None):
    """경영박사 EDI용 Excel 97-2003 파일을 만든다.

    블라인드는 기존 계산식을 유지한다. 홀딩도어는 사용자 제공 정답 전표에 맞춰
    세로 최소 150cm, 최종 최소 2.5㎡를 적용하고 작동방식/부속을 별도 품목으로 쓴다.
    """
    try:
        import xlwt
    except ImportError as e:
        raise RuntimeError(
            "EDI .xls 생성 모듈이 없습니다. requirements.txt의 xlwt를 설치해 주세요."
        ) from e

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("전표")
    out_row = 0

    def write_row(values):
        nonlocal out_row
        for col, value in enumerate(values):
            ws.write(out_row, col, value)
        out_row += 1

    def money(unit, qty, vat_zero=False):
        amount = (Decimal(str(unit)) * Decimal(str(qty))).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP)
        vat = Decimal("0") if vat_zero else (amount / 10).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP)
        return int(amount), int(vat)

    for voucher_no, o in enumerate(orders, 1):
        synchronize_order_for_outputs(o)
        d = o.get("_ship_date") or order_date or date.today()
        if isinstance(d, str):
            try:
                d = date.fromisoformat(d[:10])
            except ValueError:
                d = order_date or date.today()
        dstr = f"{d.year}.{d.month:02d}.{d.day:02d}"
        client = o.get("거래처")
        client_rules = get_client_rules(client)
        erp_client = ERP_CLIENT_NAME.get(client, client)
        delivery = o.get("배송") or {}
        _common_parts = [x.strip() for x in str(o.get("전체기재사항") or "").split("/")
                         if x.strip() and x.strip() != "포장비용"]
        extra_accessory = str(o.get("_추가부속") or "").strip()
        if extra_accessory and extra_accessory not in _common_parts:
            _common_parts.append(extra_accessory)
        order_common_note = "/".join(_common_parts)
        # 장부에서 최종 확정되는 기재사항1/2를 EDI도 그대로 재사용한다.
        # 이렇게 하면 같은 수정값이 장부에는 보이고 EDI에서는 빠지는 문제를 막는다.
        ledger_note_by_index = {}
        try:
            for lr in to_rows(o, "", M, sort_di=False):
                idx = lr.get("_source_item_index")
                if isinstance(idx, int) and lr.get("_특수") is None:
                    ledger_note_by_index[idx] = (lr.get("기재사항"), lr.get("기재사항2"))
        except Exception:
            ledger_note_by_index = {}
        blind_valid_count = 0
        blind_window_count = 0
        total_valid_count = 0
        holding_pack_qty = Decimal("0")
        holding_pack_recipient = None
        di_pack_recipient = None
        di_pack_qty = Decimal("0")

        def holding_recipient(it=None):
            it = it or {}
            # 업체별 홀딩 적요명(DI: 수령인 묶음, 휴안: 공백 제거) — clients/*.holding_recipient
            return client_rules.holding_recipient(o, it, delivery, di_pack_recipient)

        def write_holding_pack(recipient, qty):
            if not client_rules.holding_pack_enabled or not qty or Decimal(str(qty)) <= 0:
                return
            q = Decimal(str(qty)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            unit = client_rules.holding_pack_unit
            amount, vat = money(unit, q, vat_zero=client_rules.erp_vat_zero)
            name = "포장비용(H)"
            values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                      "", "", name, name, "", float(q), unit,
                      amount, vat, str(recipient or "").strip(), "", ""]
            write_row(values)

        source_items = list(o.get("items", []))
        source_idx_by_id = {id(x): i for i, x in enumerate(source_items)}
        items = source_items
        if client_rules.erp_sort_in_ledger_order:
            items = sort_di_items_in_ledger_order(items, M)

        for it in items:
            if it.get("예외품목"):
                continue
            is_holding = it.get("_product_group") == "holding" or it.get("_holding_accessory")
            is_accessory = bool(it.get("_holding_accessory"))

            # ── 홀딩도어 ────────────────────────────
            if is_holding:
                hit = M.find_order_item(it, client)
                if not hit or not hit.get("단가"):
                    continue

                recipient = holding_recipient(it)
                if client_rules.holding_pack_by_recipient and not is_accessory:
                    if di_pack_recipient is not None and recipient != di_pack_recipient:
                        write_holding_pack(di_pack_recipient, di_pack_qty)
                        di_pack_qty = Decimal("0")
                    di_pack_recipient = recipient

                if is_accessory:
                    q = Decimal(str(accessory_qty(
                        it.get("수량") if it.get("수량") not in (None, "")
                        else it.get("창개수")))).quantize(Decimal("0.01"))
                    amount, vat = money(hit["단가"], q, vat_zero=client_rules.erp_vat_zero)
                    q_text = f"{float(q):g}EA"
                    memo_parts = [q_text]
                    for value in (it.get("_수동특이"), it.get("기재사항"),
                                  it.get("설치장소"), recipient):
                        for part in str(value or "").split("/"):
                            part = part.strip()
                            if part and part not in memo_parts:
                                memo_parts.append(part)
                    memo = " ".join(memo_parts)
                    values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                              "", "", hit.get("관리코드") or hit["품명"], hit["품명"],
                              hit.get("규격") or "", float(q), int(hit["단가"]),
                              amount, vat, memo, "", ""]
                    write_row(values)
                    total_valid_count += 1
                    continue

                w, h = it.get("가로"), it.get("세로")
                if not (w and h):
                    continue
                count = _holding_count(it)
                base = holding_calc(w, h, hit["단가"])
                q = (Decimal(base["수량"]) * count).quantize(
                    Decimal("0.01"), rounding=ROUND_HALF_UP)
                amount, vat = money(hit["단가"], q, vat_zero=client_rules.erp_vat_zero)
                op = normalize_holding_operation(
                    it.get("_holding_operation") or
                    (it.get("수량") if looks_like_holding_operation(it.get("수량")) else None))
                rail = str(it.get("_holding_rail") or "").strip()
                memo_base = f"{float(w):.1f}*{float(h):.1f}/{op}"
                if rail:
                    memo_base += f"/{rail}"

                note_parts = client_rules.holding_note_parts(it, recipient)
                if note_parts:
                    memo_base += "/" + "/".join(note_parts)
                memo = ("상하로라>" if it.get("_holding_upper_roller") else "") + memo_base
                values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                          "", "", hit.get("관리코드") or hit["품명"], hit["품명"],
                          hit.get("규격") or "", float(q), int(hit["단가"]),
                          amount, vat, memo, "", ""]
                write_row(values)
                total_valid_count += 1

                # 포장비용은 부속이 아닌 본품의 청구 ㎡ 합계로 계산한다.
                if client_rules.holding_pack_by_recipient:
                    di_pack_qty += q
                elif client_rules.holding_pack_enabled:
                    holding_pack_qty += q
                    if recipient:
                        holding_pack_recipient = recipient

                if it.get("_holding_upper_roller"):
                    ex = M.find_holding_extra("H상하로라(가로m당)", client)
                    if ex and ex.get("단가"):
                        # 제공 휴안 정답: 가로 88cm도 1.0m, DI 270cm는 2.7m.
                        uq = max(Decimal(str(w)) / 100, Decimal("1.0")) * count
                        uq = uq.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                        uamount, uvat = money(ex["단가"], uq, vat_zero=client_rules.erp_vat_zero)
                        values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                                  "", "", ex.get("관리코드") or ex["품명"], ex["품명"],
                                  ex.get("규격") or "", float(uq), int(ex["단가"]),
                                  uamount, uvat, recipient, "", ""]
                        write_row(values)

                charge_each = holding_magnet_charge_qty(op)
                if charge_each:
                    ex = M.find_holding_extra("H추가비용(+자석바1)", client)
                    if ex and ex.get("단가"):
                        mq = Decimal(charge_each * count)
                        mamount, mvat = money(ex["단가"], mq, vat_zero=client_rules.erp_vat_zero)
                        values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                                  "", "", ex.get("관리코드") or ex["품명"], ex["품명"],
                                  ex.get("규격") or "", float(mq), int(ex["단가"]),
                                  mamount, mvat, recipient, "", ""]
                        write_row(values)
                continue

            # ── 기존 블라인드 ───────────────────────
            hit = M.find_order_item(it, client)
            if not hit:
                continue
            w, h = it.get("가로"), it.get("세로")
            if not (w and h):
                continue
            calc = calc_erp(w, h, hit["단가"])
            memo = f"{float(w):.1f}*{float(h):.1f}/1EA"
            # 업체별 적요 — clients/*.erp_blind_memo_parts (장부 기재사항1/2는 아래에서 맨 뒤에 붙는다)
            extra_parts = client_rules.erp_blind_memo_parts(SimpleNamespace(
                o=o, it=it, client=client, delivery=delivery,
                order_common_note=order_common_note))
            # 기재사항1 -> 기재사항2 순서를 EDI 맨 뒤에 강제한다.
            src_idx = source_idx_by_id.get(id(it))
            ledger_notes = ledger_note_by_index.get(src_idx, ()) if src_idx is not None else ()
            ledger_tail = []
            for value in ledger_notes:
                for part in str(value or "").split("/"):
                    part = part.strip()
                    if part and part not in ledger_tail:
                        ledger_tail.append(part)
            if ledger_tail:
                extra_parts = [x for x in extra_parts if x not in ledger_tail]
                extra_parts.extend(ledger_tail)
            extra = "/".join(extra_parts)
            if extra:
                memo += f" {extra}"
            split_n = handle_split(it.get("수량"))
            if not split_n:
                left_n, right_n, window_count = _blind_counts(it)
                erp_parts = []
                dir_counts = []
                if left_n or right_n:
                    if left_n:
                        dir_counts.append(("좌", left_n))
                    if right_n:
                        dir_counts.append(("우", right_n))
                else:
                    dir_counts.append((it.get("손잡이방향") or "우", window_count))
                for direction, count_n in dir_counts:
                    qty = (Decimal(str(calc["수량"])) * count_n).quantize(Decimal("0.01"))
                    amount = int((Decimal(str(calc["단가"])) * qty).quantize(
                        Decimal("1"), rounding=ROUND_HALF_UP))
                    vat = int((Decimal(amount) / 10).quantize(
                        Decimal("1"), rounding=ROUND_HALF_UP))
                    erp_parts.append((direction, qty, amount, vat))
            else:
                left_n = split_n // 2
                directions = ["좌"] * left_n + ["우"] * (split_n - left_n)
                qty_cents = int((Decimal(calc["수량"]) * 100).to_integral_value())
                q_base, q_rem = divmod(qty_cents, split_n)
                amount_total, vat_total = int(calc["금액"]), int(calc["부가세"])
                a_base, a_rem = divmod(amount_total, split_n)
                v_base, v_rem = divmod(vat_total, split_n)
                erp_parts = []
                for part_i, direction in enumerate(directions):
                    qty = Decimal(q_base + (1 if part_i < q_rem else 0)) / 100
                    amount = a_base + (1 if part_i < a_rem else 0)
                    vat = v_base + (1 if part_i < v_rem else 0)
                    erp_parts.append((direction, qty, amount, vat))
            for dir_, qty, amount, vat in erp_parts:
                if client_rules.erp_vat_zero:
                    vat = 0
                values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                          "", "", hit.get("관리코드") or hit["품명"], hit["품명"],
                          hit["규격"], float(qty), int(calc["단가"]),
                          int(amount), int(vat), memo,
                          1 if dir_ == "좌" else 2, dir_]
                write_row(values)
            blind_valid_count += 1
            try:
                _lpack, _rpack, _npack = _blind_counts(it)
                blind_window_count += max(1, int(_npack or 1))
            except Exception:
                blind_window_count += 1
            total_valid_count += 1

        # 인천)트루 추가부속은 장부/작업지시서와 같은 주문에서 경영박사에도 1개씩 출력한다.
        # 창틀용/커튼박스용 무타공은 사용자 제공 품목코드표의 공식 관리코드로 연결하며,
        # 로컬 마스터에 단가가 없으면 안전하게 0원 + 사전점검 확인으로 남긴다.
        if client == "인천)트루":
            for accessory in o.get("_true_accessories") or []:
                catalog = _true_accessory_catalog(accessory.get("표시")) or accessory
                edi_name = str(catalog.get("품명") or catalog.get("관리코드") or catalog.get("표시") or "").strip()
                if not edi_name:
                    continue
                hit = M.find_named_item(edi_name, client)
                qty = 1.0
                if hit:
                    unit = int(hit.get("단가") or 0)
                    spec = hit.get("규격") or ""
                    code = hit.get("관리코드") or edi_name
                    pname = hit.get("품명") or edi_name
                else:
                    unit = 0
                    spec = ""
                    code = str(catalog.get("관리코드") or edi_name)
                    pname = edi_name
                amount = unit
                vat = int((Decimal(amount) / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                values = [dstr, voucher_no, 3, "외출", erp_client, erp_client, "", "",
                          code, pname, spec, qty, unit, amount, vat, "", "", ""]
                write_row(values)

        # DI는 수령인별 홀딩도어 본품 합계 바로 뒤에 포장비를 붙인다.
        if client_rules.holding_pack_by_recipient and di_pack_recipient is not None:
            write_holding_pack(di_pack_recipient, di_pack_qty)
        elif (not client_rules.holding_pack_by_recipient and client_rules.holding_pack_enabled
              and holding_pack_qty):
            write_holding_pack(holding_pack_recipient or holding_recipient(), holding_pack_qty)

        # 전 업체 공통 포장비용. 택배/화물 주문이면 실제 출력되는 창 수만큼
        # 한 줄로 합산한다. 보노는 사용자 지정 품목/단가(700원)를 유지하고,
        # 그 외 거래처는 기존 공통 포장비용(25mm, 500원)을 사용한다.
        # 기존 장부에서 `# 포장비용`을 직접 넣은 경우도 역변환 시 보존한다.
        explicit_pack = "포장비용" in str(o.get("전체기재사항") or "")
        if blind_window_count and (_needs_window_packing(o) or explicit_pack):
            pack = client_rules.blind_packing_item
            amount = int(pack["단가"] * blind_window_count)
            vat = int((Decimal(amount) / 10).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
            values = [dstr, voucher_no, 3, "외출", erp_client, erp_client, "", "",
                      pack["품명"], pack["품명"], pack["규격"], float(blind_window_count),
                      pack["단가"], amount, vat, "", "", ""]
            write_row(values)

        # 휴안·유앤·보노·인천)트루 배송/주소 행.
        # 트루는 제품만 EDI에 나오고 주소가 빠지는 문제가 있었으므로 같은 배송행 구조를 사용한다.
        has_true_accessory = client == "인천)트루" and bool(o.get("_true_accessories"))
        cleaned_notice = clean_delivery_notice(delivery.get("전달사항"))
        has_shipping_info = any(str(delivery.get(k) or "").strip() for k in ("주소", "수령인", "연락처", "발신")) or bool(cleaned_notice)
        if (total_valid_count or has_true_accessory) and has_shipping_info and _effective_delivery_mode(o) in {"택배", "화물"}:
            address = _erp_address_text(delivery.get("주소"))
            method = str(delivery.get("방식") or "택배").strip()
            prepaid = str(delivery.get("선불착불") or "선불").strip()
            is_freight = "화물" in method
            is_collect = "착불" in prepaid
            if is_freight:
                ship_code = ("#화물(*파손무책)" if is_collect else
                             "#화물(선불★/파손무책/청구分)")
                ship_spec = "화물/착불" if is_collect else "화물/선불"
            else:
                ship_code = ("#택배(*파손무책/대신)" if is_collect else
                             "#택배(선불★/파손무책/청구分)")
                ship_spec = "택배/착불" if is_collect else "택배/선불"

            address_lines = wrap_erp_text(address) if address else [""]
            values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                      "", "", ship_code, ship_code, ship_spec, 0, 0, 0, 0,
                      address_lines[0], "", ""]
            write_row(values)
            for memo_line in address_lines[1:]:
                values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                          "", "", "**", "**", "", 0, 0, 0, 0,
                          memo_line, "", ""]
                write_row(values)

            if client == "보노" and o.get("_bono_receiver_is_bono"):
                # 보노 수령 건은 주소에 -보노가 들어가므로 다음 줄에는 전화만 반복한다.
                receiver = str(delivery.get("연락처") or "").strip()
            else:
                receiver = " ".join(x for x in (
                    str(delivery.get("수령인") or "").strip(),
                    str(delivery.get("연락처") or "").strip(),
                ) if x)
            if receiver:
                for memo_line in wrap_erp_text(receiver):
                    values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                              "", "", "**", "**", "", 0, 0, 0, 0,
                              memo_line, "", ""]
                    write_row(values)

            sender = str(delivery.get("발신") or "").strip()
            if sender:
                for memo_line in wrap_erp_text(sender):
                    values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                              "", "", "발신", "발신", "", 0, 0, 0, 0,
                              memo_line, "", ""]
                    write_row(values)

            notice = cleaned_notice or ""
            if notice:
                notice_code = f"전달사항({'화물' if is_freight else '택배'})"
                for memo_line in wrap_erp_text(notice):
                    values = [dstr, voucher_no, 3, "외출", erp_client, erp_client,
                              "", "", notice_code, notice_code, "", 0, 0, 0, 0,
                              memo_line, "", ""]
                    write_row(values)

    wb.save(path)
    return path
