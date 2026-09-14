"""
FitOrder 메인 화면

    streamlit run app.py
"""
import base64
import copy
import io
import json
import re
import shutil
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import openpyxl
import streamlit as st
from PIL import Image

from output import (apply_edit, build_erp, build_ledger, build_worksheet,
                    expand_same_size_directions, read_ledger, synchronize_order_for_outputs, to_rows)
import clipboard_watch as clip
from parsers import parse_excel
from holding import HOLDING_FEATURE_ENABLED, apply_holding_feature_gate
from rules import (CHANGE_LOOKBACK_DAYS, CLIENT_INFO, DUP_LOOKBACK_DAYS,
                   Master, clean_delivery_notice, default_handle_length,
                   find_changes, find_duplicates, validate)

ROOT = Path(__file__).parent.parent
MASTER = ROOT / "data" / "master" / "fitorder_master.xlsx"
LOGO = ROOT / "data" / "logo.png"
DB = ROOT / "db" / "fitorder.db"
OUT = ROOT / "out"
HISTORY = OUT / "history"
CLIENTS = list(CLIENT_INFO)
APP_VERSION = "2026.09.11-68"

st.set_page_config(page_title="FitOrder", layout="wide")
st.markdown("""<style>
[data-testid="stHeader"] {display: none;}
.block-container {padding-top: 2rem !important; max-width: 100%;
                 padding-left: 1.5rem; padding-right: 1.5rem;}
[data-testid="stImage"] img, [data-testid="stImage"], .stImage img
  {border-radius: 0 !important;}
.stButton button, .stDownloadButton button,
[data-testid="stFileUploader"] section,
[data-baseweb="input"], [data-baseweb="select"] > div,
[data-testid="stAlert"], [data-testid="stMetric"],
[data-testid="stDataFrame"], [data-testid="stExpander"]
  {border-radius: 0 !important;}
</style>""", unsafe_allow_html=True)

EDIT_COLS = ["상호", "색상", "가로", "세로", "수량", "방향", "길이", "특이",
             "기재사항", "기재사항2", "출고일"]


def normalize_editor_color(value, prefix="B"):
    """편집 셀의 숫자 품목코드를 화면용 ` B 200` 형태로 바꾼다."""
    text = str(value or "").strip()
    if not re.fullmatch(r"\d{3,4}[A-Za-z]*", text):
        return value
    prefix = str(prefix or "B").strip().upper()
    if prefix not in {"B", "H", "R", "C"}:
        prefix = "B"
    return f" {prefix} {text}"


def client_label(client):
    """편집 표에 보여줄 상호 + 내부표시."""
    if not client:
        return ""
    mark = CLIENT_INFO.get(client, (None, None, None, ""))[3]
    return f"{client} {mark}".strip()


CLIENT_LABELS = [client_label(c) for c in CLIENTS]
FILE_CLIENT_ALIASES = {
    "두창블라인드": "DU", "두창": "DU",
    "대일산업": "DI", "대일": "DI",
    "루임트": "RT", "미더스": "M", "제이원": "JO",
    "대동산업": "DD", "대동": "DD",
    "유앤아이티엔에스": "유앤", "유앤아이": "유앤", "UNITNS": "유앤",
    "아지트": "아지트", "윈도우투모로우": "WT",
    "트루갤러리": "인천)트루", "인천)트루": "인천)트루",
    "미래가공": "미래가공", "이끌림": "보노", "보노": "보노",
    "미성텍스": "MS",
}


def client_from_filename(path):
    """파일명의 업체명/코드로 안전하게 거래처를 추론한다."""
    stem = path.stem.strip()
    upper = stem.upper()
    for name, code in FILE_CLIENT_ALIASES.items():
        if name in stem:
            return code
    for code in CLIENTS:
        if re.search(rf"(^|[^A-Z]){re.escape(code.upper())}([^A-Z]|$)", upper):
            return code
    return None


# ─────────────────────────────────────────────
# DB
# ─────────────────────────────────────────────
def db():
    DB.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB)
    con.execute("""CREATE TABLE IF NOT EXISTS batch(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ship_date TEXT, created TEXT, payload TEXT,
        printed TEXT, print_seq INTEGER DEFAULT 0)""")
    con.execute("""CREATE TABLE IF NOT EXISTS output_history(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        created TEXT, day_key TEXT, day_seq INTEGER,
        first_client TEXT, ledger_file TEXT)""")
    return con


def save_batch(ship_date, orders):
    con = db()
    cur = con.execute(
        "INSERT INTO batch(ship_date, created, payload) VALUES(?,?,?)",
        (str(ship_date), datetime.now().isoformat(),
         json.dumps(orders, ensure_ascii=False, default=str)))
    con.commit()
    bid = cur.lastrowid
    con.close()
    return bid


def mark_printed(bid, orders):
    con = db()
    con.execute("""UPDATE batch SET printed=?, print_seq=print_seq+1,
                   payload=? WHERE id=?""",
                (datetime.now().isoformat(),
                 json.dumps(orders, ensure_ascii=False, default=str), bid))
    con.commit()
    con.close()


def reserve_output_names(orders):
    """작성일-첫 업체-오늘 출력 순번으로 다운로드 파일명을 만든다."""
    now = datetime.now()
    day_key = now.date().isoformat()
    first = str((orders[0].get("거래처") if orders else None) or "미확인")
    first = re.sub(r'[\\/:*?"<>|\s]+', "", first) or "미확인"
    con = db()
    con.execute("BEGIN IMMEDIATE")
    seq = con.execute(
        "SELECT COALESCE(MAX(day_seq), 0) + 1 FROM output_history WHERE day_key=?",
        (day_key,)).fetchone()[0]
    tag = f"{now:%m%d}-{first}{seq:02d}"
    ledger_name = f"장부_{tag}.xlsx"
    con.execute("""INSERT INTO output_history
                   (created, day_key, day_seq, first_client, ledger_file)
                   VALUES(?,?,?,?,?)""",
                (now.isoformat(), day_key, seq, first, ledger_name))
    con.commit()
    con.close()
    return {
        "장부": ledger_name,
        "작업지시서": f"작업지시서_{tag}.xlsx",
        "경영박사": f"경영박사EDI_{tag}.xls",
    }


def archive_ledger(filename):
    HISTORY.mkdir(parents=True, exist_ok=True)
    shutil.copy2(OUT / "장부.xlsx", HISTORY / filename)


def ledger_history(limit=30):
    con = db()
    rows = con.execute("""SELECT created, first_client, ledger_file
                          FROM output_history ORDER BY id DESC LIMIT ?""",
                       (limit,)).fetchall()
    con.close()
    return [(created, client, name) for created, client, name in rows
            if name and (HISTORY / name).exists()]


def show_ledger_history():
    """이전에 생성한 장부만 화면에서 확인한다."""
    history = ledger_history()
    with st.expander("이전 장부 보기"):
        if not history:
            st.caption("이 기능을 적용한 후 생성한 장부가 없습니다.")
            return
        labels = {
            name: f"{datetime.fromisoformat(created):%Y-%m-%d %H:%M} · {client} · {name}"
            for created, client, name in history
        }
        selected = st.selectbox("장부 선택", list(labels),
                                format_func=lambda name: labels[name])
        path = HISTORY / selected
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        headers = ["상호", "색상", "가로", "X", "세로", "수량", "방향", "길이",
                   "특이", "기재사항", "기재사항2", "출고일"]
        for ws in wb.worksheets:
            data = []
            for values in ws.iter_rows(min_row=3, max_col=13, values_only=True):
                row = list(values[1:13])
                if any(v not in (None, "") for i, v in enumerate(row) if i != 3):
                    data.append(row)
            if len(wb.worksheets) > 1:
                st.caption(ws.title)
            st.dataframe(pd.DataFrame(data, columns=headers),
                         hide_index=True, width="stretch")
        st.download_button("이 장부 다운로드", path.read_bytes(),
                           file_name=selected,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width="stretch")


def past_items(days):
    """최근 days 일 배치에서 (라벨, 거래처, item, 출력여부) 목록"""
    con = db()
    since = (datetime.now() - timedelta(days=days)).isoformat()
    rows = con.execute("""SELECT id, ship_date, created, payload, printed
                          FROM batch WHERE created >= ? ORDER BY created DESC""",
                       (since,)).fetchall()
    con.close()
    out = []
    for bid, ship, created, payload, printed in rows:
        if bid == st.session_state.get("batch_id"):
            continue                       # 현재 배치는 제외
        try:
            orders = json.loads(payload)
        except Exception:
            continue
        label = f"{str(ship)[5:]} 출고"
        for o in orders:
            c = o.get("거래처")
            for it in o.get("items", []):
                if not it.get("예외품목"):
                    out.append((label, c, it, bool(printed)))
    return out


# ─────────────────────────────────────────────
# 출고일
# ─────────────────────────────────────────────
WD = "월화수목금토일"


def default_ship():
    now = datetime.now()
    d = now.date() + timedelta(days=1 if now.hour < 15 else 2)
    if d.weekday() == 6:
        d += timedelta(days=1)
    return d


def ship_label(d, mode=None):
    lab = WD[d.weekday()]
    return f"{lab}({mode})" if mode in ("택배", "화물") else lab


# ─────────────────────────────────────────────
# 장부 행 계산 (주문 -> 행 + 역참조)
# ─────────────────────────────────────────────
def rebuild_rows(orders, ship, M=None, sort_di=False):
    rows = []
    sp_notices = []
    for order in orders:
        if order.get("거래처") != "SP":
            continue
        for part in str(order.get("전체기재사항") or "").split("/"):
            part = part.strip()
            if "공지" in part and part not in sp_notices:
                sp_notices.append(part)
    sp_notice_written = False
    for oi, o in enumerate(orders):
        # RT는 항상 실제 배송 방식을 표시한다. 그 외 업체는 실제 배송이
        # 업체 기본 배송과 다를 때만 (택배)/(화물)/(배달)을 표시한다.
        client = o.get("거래처")
        actual = (o.get("배송") or {}).get("방식")
        default = CLIENT_INFO.get(client, (None, None, None))[2]
        mode = actual if client == "RT" or (
            actual in ("택배", "화물", "배달") and actual != default) else None
        order_ship = o.get("_ship_date") or ship
        if isinstance(order_ship, str):
            try:
                order_ship = date.fromisoformat(order_ship[:10])
            except ValueError:
                order_ship = ship
        got = to_rows(o, ship_label(order_ship, mode), M, sort_di=sort_di)
        if o.get("거래처") == "SP":
            first_product = None
            for row in got:
                if row.get("_특수") is not None:
                    continue
                if first_product is None:
                    first_product = row
                parts = [x.strip() for x in
                         str(row.get("기재사항") or "").split("/")
                         if x.strip() and "공지" not in x]
                row["기재사항"] = "/".join(parts) or None
            if first_product is not None and sp_notices and not sp_notice_written:
                old = first_product.get("기재사항")
                first_product["기재사항"] = "/".join(
                    sp_notices + ([old] if old else []))
                sp_notice_written = True
        visible_i = 0
        for r in got:
            # 특수행까지 주문 ID를 부여해야 출고일이 한 사람의 상품부터
            # 주소·연락처·전달사항 끝까지 하나의 셀로 병합된다.
            r["_oi"] = oi
            if r.get("_특수") is None:
                # 편집/삭제는 `예외품목` 필터나 DI 정렬과 무관하게 원 주문 items의
                # 실제 인덱스를 사용한다. 그래야 화면에서 수정한 특이/품명 등이
                # 다른 행에 적용되거나 파일 생성 때 원래 값으로 돌아가지 않는다.
                source_i = r.get("_source_item_index")
                r["_ii"] = source_i if isinstance(source_i, int) else visible_i
                r["_출고일날짜"] = order_ship.isoformat() if visible_i == 0 else None
                visible_i += 1
            rows.append(r)
    return rows


def ingest(path, client_hint, label):
    """이미지 또는 엑셀 파일 1개 -> 주문 목록에 추가"""
    from pathlib import Path as _P
    path = _P(path)
    if path.suffix.lower() in (".xlsx", ".xls"):
        got = parse_excel(path, client_hint)
    elif path.suffix.lower() == ".pdf":
        from extract import extract_pdf
        # 사용자가 선택한 거래처 또는 파일명에서 확인된 거래처만 전달한다.
        c = client_hint or client_from_filename(path)
        got = extract_pdf(path,
                          client_hint=c if c in CLIENTS else None)
    else:
        from extract import extract_order
        c = client_hint or client_from_filename(path)
        got = [extract_order([path],
                             client_hint=c if c in CLIENTS else None)]
    # 홀딩 자동판별 OFF 상태에서는 실수로 붙은 홀딩 분류를 일반 주문으로 되돌린다.
    for _order in got or []:
        apply_holding_feature_gate(_order)

    if not got or not any((o.get("items") or []) for o in got):
        raise ValueError(
            "발주서에서 유효한 품목을 읽지 못했습니다. "
            "거래처 선택과 발주서 형식을 확인해 주세요."
        )
    for o in got:
        # 좌/우 열에 적힌 개수를 방향으로 확정한다. LLM이 열의 숫자는 읽었지만
        # 손잡이방향을 비운 경우에도 코드에서 결정적으로 보정한다.
        for it in o.get("items") or []:
            try:
                left_count = int(float(it.get("좌개수") or 0))
                right_count = int(float(it.get("우개수") or 0))
            except (TypeError, ValueError):
                left_count = right_count = 0
            if left_count > 0 and right_count == 0:
                it["손잡이방향"] = "좌"
            elif right_count > 0 and left_count == 0:
                it["손잡이방향"] = "우"
        if o.get("거래처") == "DD":
            for it in o.get("items") or []:
                # 대동산업은 1700, 1745처럼 mm 단위로 적으므로 cm로 변환한다.
                for field in ("가로", "세로"):
                    value = it.get(field)
                    if isinstance(value, (int, float)) and value >= 500:
                        converted = value / 10
                        it[field] = int(converted) if converted.is_integer() else converted
                # NA029FP, GR870FP의 재질 접두사는 제외하고 색상코드는 유지한다.
                code = str(it.get("품목코드") or "").strip().upper()
                it["품목코드"] = re.sub(r"^(?:NA|GR)", "", code) or None
        if o.get("거래처") == "아지트":
            raw = " ".join(str(x or "") for x in (
                o.get("전체원문"), o.get("전체기재사항")))
            o["_az_place"] = "금빛커텐" if "에어캡+포장" in raw else "시온가공소"
            common = str(o.get("전체기재사항") or "")
            common = re.sub(r"(?:^|/)☆?(?:금빛커텐|시온가공소)(?=/|$)", "", common)
            o["전체기재사항"] = common.strip("/") or None
            o.setdefault("배송", {})["방식"] = "배달"
        # 사이즈가 한 행뿐이어도 방향이 `좌우`/`우우`처럼 여러 개면
        # 같은 사이즈의 여러 창으로 확정한다. 세 출력물이 모두 이 주문 데이터를 사용한다.
        expand_same_size_directions(o)
        synchronize_order_for_outputs(o)
        o["_file"] = label
        o["_source_path"] = str(path)
        st.session_state.orders.append(o)
    return len(got)


def show_source(order, key_prefix):
    """주문과 연결된 업로드 이미지/PDF 페이지를 보여준다."""
    label = str(order.get("_file") or "")
    path = Path(order.get("_source_path") or (OUT / "_upload" / label))
    if not path.exists():
        st.caption("원본 파일을 찾을 수 없습니다. (out/_upload을 지운 경우 재업로드 필요)")
        return
    suffix = path.suffix.lower()
    if suffix in (".png", ".jpg", ".jpeg"):
        st.image(str(path), caption=label, width="stretch")
        return
    if suffix != ".pdf":
        st.caption("이 형식은 원본 미리보기를 제공하지 않습니다.")
        return
    try:
        import pymupdf
        pages = order.get("_source_pages") or [1]
        pages = sorted({int(p) for p in pages if int(p) > 0}) or [1]
        page_no = pages[0] if len(pages) == 1 else st.selectbox(
            "원본 PDF 페이지", pages,
            format_func=lambda p: f"{p}페이지", key=f"src_page_{key_prefix}")
        with pymupdf.open(str(path)) as doc:
            page_no = min(page_no, doc.page_count)
            pix = doc[page_no - 1].get_pixmap(
                matrix=pymupdf.Matrix(1.5, 1.5), alpha=False)
            image = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            if order.get("거래처") == "SP":
                image = image.rotate(90, expand=True)
            st.image(image, caption=f"{label} · {page_no}페이지",
                     width="stretch")
    except Exception as e:
        st.warning(f"PDF 미리보기를 열지 못했습니다. ({type(e).__name__})")


ss = st.session_state
ss.setdefault("orders", [])
ss.setdefault("ship", default_ship())
ss.setdefault("batch_id", None)
ss.setdefault("dismissed", set())
ss.setdefault("seen", set())
ss.setdefault("upl", 0)
ss.setdefault("watch", False)
ss.setdefault("clip_seen", set())
ss.setdefault("clip_log", [])
ss.setdefault("clip_last", None)
ss.setdefault("net_msg", None)
ss.setdefault("net_at", 0)
ss.setdefault("done", 0)
ss.setdefault("issue_view", None)
ss.setdefault("ledger_editor_version", 0)
ss.setdefault("issue_selected_row", None)
ss.setdefault("manual_item_prefix", None)
ss.setdefault("ledger_convert", None)
ss.setdefault("output_names", None)
ss.setdefault("undo_stack", [])
ss.setdefault("detail_row_idx", 0)
ss.setdefault("editor_selected_rows", [])
# 사전점검 열 너비는 Excel 양식과 완전히 별개다. 이전 세션의 과도하게 넓어진
# 값을 끌고 오지 않도록 UI 규격 버전이 바뀌면 한 번만 기준폭으로 초기화한다.
EDITOR_WIDTH_PROFILE = "v66-compact"
if ss.get("editor_width_profile") != EDITOR_WIDTH_PROFILE:
    ss.editor_col_widths = {
        "색상": 150, "특이": 85, "기재사항": 145, "기재사항2": 145
    }
    ss.editor_width_profile = EDITOR_WIDTH_PROFILE
else:
    ss.setdefault("editor_col_widths", {
        "색상": 150, "특이": 85, "기재사항": 145, "기재사항2": 145
    })

def _push_undo(label):
    """사전점검의 파괴적 변경을 최대 10단계까지 복구할 수 있게 보관한다."""
    ss.undo_stack.append((label, copy.deepcopy(ss.orders)))
    ss.undo_stack = ss.undo_stack[-10:]


def _reset_editor_selection():
    ss.ledger_editor_version += 1
    ss.issue_selected_row = None
    ss.issue_view = None
    ss.editor_selected_rows = []
    ss.batch_id = None


def _blank_row_like(base):
    """선택 행 아래 추가용. 품목 종류는 복사하고 행별 제작값만 비운다."""
    item = copy.deepcopy(base)
    for key in ("가로", "세로", "수량", "손잡이방향", "손잡이길이",
                "설치장소", "기재사항", "예외품목", "원문",
                "좌개수", "우개수"):
        item[key] = None
    item["창개수"] = 1
    item["연창"] = False
    item["_수동특이"] = None
    item["_manual_added"] = True
    item["확신도"] = {"가로": 1.0, "세로": 1.0,
                      "품목코드": 1.0, "손잡이": 1.0}
    return item


def _move_order_block(order_indices, direction):
    """선택한 연속 주문 묶음을 한 묶음 그대로 위/아래로 이동한다."""
    idx = sorted(set(int(x) for x in order_indices))
    if not idx:
        return False
    if idx != list(range(idx[0], idx[-1] + 1)):
        return False
    start, end = idx[0], idx[-1]
    if direction < 0:
        if start == 0:
            return False
        block = ss.orders[start:end + 1]
        prev = ss.orders[start - 1]
        ss.orders[start - 1:end + 1] = block + [prev]
    else:
        if end >= len(ss.orders) - 1:
            return False
        block = ss.orders[start:end + 1]
        nxt = ss.orders[end + 1]
        ss.orders[start:end + 2] = [nxt] + block
    return True


def _note_parts(value):
    return [x.strip() for x in str(value or "").split("/") if x.strip()]


def _detail_address(order):
    """사전점검 부가정보에 보여줄 주소. 배송 객체를 우선하고 구형 필드도 보조한다."""
    delivery = order.get("배송") or {}
    for value in (delivery.get("주소"), order.get("주소"), order.get("배송주소"),
                  order.get("수신주소")):
        text = str(value or "").strip()
        if text:
            return text
    return ""


def _accessory_summary(order):
    """사전점검 하단에 보여줄 추가부속 요약.

    구조화된 유앤/트루 부속뿐 아니라 휴안 등에서 기재사항에 이미 들어간
    노피스·콘크리트·석고·앙카 계열도 함께 보여준다.
    """
    parts = []

    def add(text):
        text = str(text or "").strip()
        if text and text not in parts:
            parts.append(text)

    add(order.get("_추가부속"))
    for a in order.get("_unit_accessories") or []:
        name = str(a.get("품명") or "").strip()
        if not name:
            continue
        sets = a.get("세트")
        piece_count = a.get("피스수")
        suffix = ""
        if piece_count:
            suffix += f" ({piece_count})"
        if sets:
            suffix += f" X{sets}set"
        add(name + suffix)
    for a in order.get("_true_accessories") or []:
        text = str(a.get("표시") or "").strip()
        if text:
            add(text + " X1")

    accessory_rx = re.compile(
        r"(?:노피스|콘크리트|석고(?:앙카|날개)?|앙카|브라켓|브라캣|스냅|무타공|피스\s*\()",
        re.I)
    for it in order.get("items") or []:
        # 파서별 저장 위치가 달라도 사전점검에서는 한 군데에서 보이게 한다.
        for field in ("기재사항", "설치장소", "특이", "_수동특이", "원문"):
            for token in _note_parts(it.get(field)):
                if accessory_rx.search(token):
                    add(token)
    for token in _note_parts(order.get("전체기재사항")):
        if accessory_rx.search(token):
            add(token)
    return " · ".join(parts)


def _set_piece_flag(order, enabled):
    parts = _note_parts(order.get("전체기재사항"))
    parts = [x for x in parts if x != "피스"]
    if enabled:
        parts.insert(0, "피스")
    order["전체기재사항"] = "/".join(parts) or None


M = Master(MASTER) if MASTER.exists() else None
left, right = st.columns([3, 7], gap="medium")

# ─────────────────────────────────────────────
# 좌측
# ─────────────────────────────────────────────
with left:
    issue_view = ss.get("issue_view")
    if issue_view and 0 <= issue_view.get("order", -1) < len(ss.orders):
        st.subheader(issue_view.get("제목") or "상태 원인")
        ledger_row = issue_view.get("장부행") or {}
        if ledger_row:
            st.markdown(
                f"**{ledger_row.get('상호') or '-'} · "
                f"{ledger_row.get('색상') or '-'}**  \n"
                f"규격: {ledger_row.get('가로') or '-'} X "
                f"{ledger_row.get('세로') or '-'} · "
                f"수량: {ledger_row.get('수량') or '-'}  \n"
                f"손잡이: {ledger_row.get('방향') or '-'} "
                f"{ledger_row.get('길이') or ''}  \n"
                f"특이: {ledger_row.get('특이') or '-'}  \n"
                f"기재사항: {ledger_row.get('기재사항') or '-'} "
                f"{ledger_row.get('기재사항2') or ''}")
        issue_details = issue_view.get("원인") or []
        if not issue_details:
            st.success("검증 문제 없음")
        for detail in issue_details:
            text = f"{detail.get('코드')} · {detail.get('내용')}"
            if detail.get("등급") == "red":
                st.error(text)
            elif detail.get("등급") == "review":
                st.info(text, icon="🟪")
            else:
                st.warning(text)
        show_source(ss.orders[issue_view["order"]], "left_issue")
        if st.button("상태 상세 닫기", width="stretch"):
            ss.issue_view = None
            ss.issue_selected_row = None
            st.rerun()
        st.divider()

    if LOGO.exists():
        st.image(str(LOGO), width=110)
    st.caption(f"코드 버전 {APP_VERSION}")

    st.caption("출고일")
    new_ship = st.date_input("출고일", ss.ship, label_visibility="collapsed")
    if new_ship != ss.ship:
        ss.ship = new_ship
    st.caption(f"{WD[ss.ship.weekday()]}요일"
               + ("  ·  3시 전 익일" if datetime.now().hour < 15
                  else "  ·  3시 이후 익익일"))

    # 시작할 때 API 연결을 미리 검사하지 않는다.
    # 공장망이 느리거나 차단된 경우 화면만 여는 데도
    # 최대 15초가 더 걸리던 문제를 방지한다. 실패 시 extract.py가 진단한다.
    if ss.get("net_msg"):
        st.error("**연결 문제**\n\n" + ss.net_msg)
        if st.button("다시 확인", width="stretch"):
            ss.net_at = 0
            st.rerun()

    st.divider()
    hint = st.selectbox("거래처", ["자동 판별"] + CLIENTS)

    if clip.AVAILABLE:
        def take_clip():
            """클립보드의 이미지를 1건 가져와 주문에 추가"""
            img, h = clip.grab()
            if not h or clip.too_small(img):
                st.warning("클립보드에 발주서 이미지가 없습니다.")
                return
            if h in ss.clip_seen:
                st.info("이미 추가한 캡처입니다.")
                return
            ss.clip_seen.add(h)
            ss.clip_last = h
            c = None if hint == "자동 판별" else hint
            with st.spinner("캡처 분석 중"):
                try:
                    fp = clip.save(img, OUT, h[:8])
                    ingest(fp, c, f"캡처 {len(ss.clip_log) + 1}")
                    ss.clip_log.append(h)
                except Exception as e:
                    st.error(f"분석 실패 — {e}")
                    return
            st.rerun()

        if st.button("붙여넣기 (Ctrl+V)", width="stretch"):
            take_clip()

        # Ctrl+V 키를 위 버튼 클릭으로 연결
        st.iframe("""<div style="height:0;overflow:hidden"></div><script>
const doc = window.parent.document;
if (!doc.__fitorderPaste) {
  doc.__fitorderPaste = true;
  doc.addEventListener('keydown', function (e) {
    if (!(e.ctrlKey || e.metaKey) || e.key.toLowerCase() !== 'v') return;
    const tag = (doc.activeElement && doc.activeElement.tagName) || '';
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;
    const btn = Array.from(doc.querySelectorAll('button'))
      .find(b => b.innerText.indexOf('붙여넣기') !== -1);
    if (btn) { e.preventDefault(); btn.click(); }
  });
}
</script>""", height=1)

        prev_watch = ss.watch
        ss.watch = st.toggle(
            "자동 캡처 감시", value=ss.watch,
            help="켜두면 화면을 캡처하는 즉시 장부에 추가됩니다. "
                 "다른 용도로 캡처할 때는 꺼두세요.")

        if ss.watch and not prev_watch:
            # 감시를 켠 시점에 클립보드에 남아 있던 이미지는 무시한다
            _, h0 = clip.grab()
            if h0:
                ss.clip_seen.add(h0)
                ss.clip_last = h0

        @st.fragment(run_every=2 if ss.watch else None)
        def _watch():
            if not ss.watch:
                return
            st.caption("감시 중 · 캡처하면 자동 추가")
            img, h = clip.grab()
            if not h or h == ss.clip_last or h in ss.clip_seen \
                    or clip.too_small(img):
                return
            ss.clip_seen.add(h)
            ss.clip_last = h
            c = None if hint == "자동 판별" else hint
            with st.spinner("캡처 분석 중"):
                try:
                    fp = clip.save(img, OUT, h[:8])
                    ingest(fp, c, f"캡처 {len(ss.clip_log) + 1}")
                    ss.clip_log.append(h)
                except Exception as e:
                    st.error(f"캡처 분석 실패 — {e}")
                    return
            st.rerun(scope="app")

        _watch()
    files = st.file_uploader(
        "발주서를 끌어다 놓거나 선택하세요",
        type=["png", "jpg", "jpeg", "pdf", "xlsx", "xls"],
        accept_multiple_files=True,
        key=f"upl_{ss.upl}")
    if not HOLDING_FEATURE_ENABLED:
        st.caption("홀딩도어 자동 판별: 사용 안 함 · 모든 품목을 일반 주문으로 처리합니다.")

    analyze_col, ledger_col = st.columns(2)
    analyze_clicked = analyze_col.button(
        "분석 시작", type="primary", width="stretch", disabled=not files)
    ledger_clicked = ledger_col.button(
        "장부 → 경영박사 EDI", width="stretch", disabled=not files,
        help="사람이 직접 작성한 기존 장부.xls/xlsx 또는 FitOrder 장부를 읽어 경영박사 EDI와 작업지시서를 생성합니다. 일반 블라인드·MIX를 지원하며 현재 홀딩도어 별도 판별은 하지 않습니다.")

    if analyze_clicked:
        client = None if hint == "자동 판별" else hint
        todo = [f for f in files
                if (f.name, f.size) not in ss.seen]      # 이미 처리한 건 제외
        skipped = len(files) - len(todo)
        if not todo:
            st.info("새로 추가된 발주서가 없습니다.")
        bar = st.progress(0.0, "준비 중")
        succeeded = 0
        for i, f in enumerate(todo, 1):
            bar.progress((i - 1) / len(todo), f"{f.name} 분석 중")
            tmp = OUT / "_upload" / f.name
            tmp.parent.mkdir(parents=True, exist_ok=True)
            tmp.write_bytes(f.getbuffer())
            try:
                ingest(tmp, client, f.name)
                # 분석과 주문 추가가 성공한 파일만 완료 처리한다.
                # 실패한 파일을 미리 seen에 넣으면 재시도할 때
                # '이미 처리한 파일'로 건너뛰는 문제가 발생한다.
                ss.seen.add((f.name, f.size))
                succeeded += 1
            except Exception as e:
                st.error(f"**{f.name} 분석 실패**\n\n{e}")
                ss.seen.discard((f.name, f.size))
                ss.net_at = 0
            bar.progress(i / len(todo))
        bar.empty()
        if succeeded:
            ss.upl += 1
        if skipped:
            st.caption(f"이미 처리한 {skipped}건은 건너뛰었습니다.")
        if succeeded:
            st.rerun()

    if ledger_clicked:
        if len(files) != 1:
            st.error("장부 변환에는 장부 파일 하나만 넣어 주세요.")
        elif Path(files[0].name).suffix.lower() not in (".xlsx", ".xls"):
            st.error("기존 장부.xls 또는 FitOrder 장부.xlsx 파일을 넣어 주세요.")
        elif M is None:
            st.error(f"마스터 파일이 없습니다: {MASTER}")
        else:
            try:
                ledger_rows, ledger_orders, ledger_errors = read_ledger(
                    io.BytesIO(files[0].getvalue()), M, files[0].name)
                item_count = sum(1 for o in ledger_orders for it in o.get("items", []) if not it.get("예외품목"))
                ss.ledger_convert = {
                    "file": files[0].name, "orders": len(ledger_orders),
                    "items": item_count, "errors": ledger_errors,
                }
                if ledger_errors:
                    st.error("장부에서 확인이 필요한 부분이 있습니다.")
                    for message in ledger_errors:
                        st.warning(message)
                elif not ledger_orders or not item_count:
                    st.error("장부에서 변환할 주문을 찾지 못했습니다.")
                else:
                    OUT.mkdir(parents=True, exist_ok=True)
                    build_worksheet(ledger_rows, OUT / "작업지시서.xlsx")
                    build_erp(ledger_orders, M, OUT / "경영박사_EDI.xls", ss.ship)
                    st.success(f"주문 {len(ledger_orders)}건 · 품목 {item_count}행 변환 완료")
            except Exception as e:
                ss.ledger_convert = None
                st.error(f"장부 변환 실패 — {e}")

    converted = ss.get("ledger_convert")
    if converted and not converted.get("errors") \
            and (OUT / "작업지시서.xlsx").exists() \
            and (OUT / "경영박사_EDI.xls").exists():
        st.caption(f"{converted['file']} · 주문 {converted['orders']}건 · "
                   f"품목 {converted['items']}행")
        cv1, cv2 = st.columns(2)
        cv1.download_button(
            "작업지시서 다운로드", (OUT / "작업지시서.xlsx").read_bytes(),
            file_name="작업지시서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch")
        cv2.download_button(
            "경영박사 EDI 다운로드", (OUT / "경영박사_EDI.xls").read_bytes(),
            file_name="경영박사_EDI.xls", mime="application/vnd.ms-excel",
            width="stretch")

    show_ledger_history()

    if ss.orders:
        st.divider()
        st.caption(f"주문 {len(ss.orders)}건 대기 중")
        if st.button("방금 넣은 것 취소", width="stretch"):
            ss.orders.pop()
            if ss.clip_log:
                ss.clip_seen.discard(ss.clip_log.pop())
            ss.batch_id = None
            st.rerun()
        if st.button("전체 비우기", width="stretch"):
            ss.orders, ss.batch_id = [], None
            ss.issue_view = None
            ss.dismissed, ss.seen = set(), set()
            ss.clip_log = []
            # 감시 중이면 현재 클립보드를 본 것으로 남겨 재분석을 막는다
            ss.clip_seen, ss.clip_last = set(), None
            if ss.get("watch"):
                try:
                    _, _h = clip.grab()
                    if _h:
                        ss.clip_seen.add(_h)
                        ss.clip_last = _h
                except Exception:
                    pass
            st.rerun()

# ─────────────────────────────────────────────
# 우측
# ─────────────────────────────────────────────
with right:
    if not ss.orders:
        if ss.get("done"):
            st.success(f"장부 {ss.done}행을 파일로 만들었습니다. "
                       "아래에서 내려받으세요.")
            MIME = ("application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet")
            d1, d2, d3 = st.columns(3)
            names = ss.get("output_names") or {}
            downloads = ((d1, OUT / "장부.xlsx", MIME, names.get("장부")),
                         (d2, OUT / "작업지시서.xlsx", MIME,
                          names.get("작업지시서")),
                         (d3, OUT / "경영박사_EDI.xls",
                          "application/vnd.ms-excel", names.get("경영박사")))
            for col, fp, mime, download_name in downloads:
                if fp.exists():
                    col.download_button(fp.name, fp.read_bytes(),
                                        file_name=download_name or fp.name, mime=mime,
                                        width="stretch")
            st.caption("새 발주서를 올리면 다음 장부가 시작됩니다.")
        else:
            st.info("발주서를 업로드하면 장부가 여기에 쌓입니다.")
        st.stop()
    if M is None:
        st.error(f"마스터 파일이 없습니다: {MASTER}")
        st.stop()

    rows = rebuild_rows(ss.orders, ss.ship, M)
    prod = [r for r in rows if r.get("_특수") is None]

    # 장부가 생겼을 때만 상단에 H/R/C 직접 입력 접두사 스위치를 표시한다.
    # 마지막으로 켠 항목 하나만 유지되며 모두 끄면 기본 B다.
    def choose_item_prefix(prefix):
        key = f"item_prefix_{prefix}"
        if ss.get(key):
            for other in ("H", "R", "C"):
                if other != prefix:
                    ss[f"item_prefix_{other}"] = False
            ss.manual_item_prefix = prefix
        elif ss.manual_item_prefix == prefix:
            ss.manual_item_prefix = None

    p1, p2, p3, _ = st.columns([1, 1, 1, 7])
    for col, prefix in ((p1, "H"), (p2, "R"), (p3, "C")):
        col.toggle(prefix, key=f"item_prefix_{prefix}",
                   on_change=choose_item_prefix, args=(prefix,))

    # 검증
    issues, exceptions = [], []
    for oi, o in enumerate(ss.orders):
        for lv, code, msg, ri in validate(o, M):
            # 예전 rules.py가 함께 남아 있더라도 폐기된 규칙은 표시하지 않는다.
            if code == "치수순서의심":
                continue
            rec = {"주문": oi + 1, "파일": o.get("_file"), "등급": lv,
                   "코드": code, "내용": msg, "행": ri,
                   "_key": f"{oi}|{code}|{ri}"}
            (exceptions if lv == "info" else issues).append(rec)

    # 중복 감지 (7일) / 변경 탐지 (2일)
    try:
        p7 = past_items(DUP_LOOKBACK_DAYS)
        p2 = [x for x in past_items(CHANGE_LOOKBACK_DAYS)]
        extra = find_duplicates(ss.orders, [(l, c, i) for l, c, i, _ in p7]) \
                + find_changes(ss.orders, p2)
    except Exception:
        extra = []
    for lv, code, msg, ri in extra:
        issues.append({"주문": "-", "파일": "장부", "등급": lv, "코드": code,
                       "내용": msg, "행": ri, "_key": f"x|{code}|{ri}|{msg[:20]}"})

    live = [i for i in issues if i["_key"] not in ss.dismissed]
    n_red = sum(1 for i in live if i["등급"] == "red")
    n_rev = sum(1 for i in live if i["등급"] == "review")
    n_yel = sum(1 for i in live if i["등급"] == "yellow")

    # 장부 행별 상태: 오류 > 중복·변경 > 확인 순으로 한 개만 표시
    status_rank = {"yellow": 1, "review": 2, "red": 3}
    status_text = {
        "red": "🟥 오류",
        "yellow": "🟨 확인",
        "review": "🟪 중복·변경",
    }
    row_levels = {}
    row_issue_details = {}

    def mark_status(row_index, level, issue):
        if not (0 <= row_index < len(prod)) or level not in status_rank:
            return
        row_issue_details.setdefault(row_index, []).append(issue)
        old = row_levels.get(row_index)
        if old is None or status_rank[level] > status_rank[old]:
            row_levels[row_index] = level

    for issue in live:
        level, ri = issue["등급"], issue.get("행")
        order_num = issue.get("주문")
        if isinstance(order_num, int):
            oi = order_num - 1
            if ri is None:  # 거래처·배송 등 주문 전체 문제
                for k, row in enumerate(prod):
                    if row.get("_oi") == oi:
                        mark_status(k, level, issue)
                continue
            items = ss.orders[oi].get("items") or []
            source_i = int(ri) - 1
            if not (0 <= source_i < len(items)) or items[source_i].get("예외품목"):
                continue
            for k, row in enumerate(prod):
                if row.get("_oi") == oi and row.get("_ii") == source_i:
                    mark_status(k, level, issue)
                    break
        elif isinstance(ri, int):  # 중복·변경의 전역 장부 행번호
            mark_status(ri - 1, level, issue)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("장부 행", len(prod))
    c2.metric("오류", n_red)
    c3.metric("중복·변경", n_rev)
    c4.metric("확인", n_yel)
    c5.metric("예외품목", len(exceptions))

    if exceptions:
        with st.expander(f"예외 사항이 발생한 주문 건 — {len(exceptions)}건", True):
            st.dataframe(pd.DataFrame(exceptions)[["파일", "내용"]],
                         hide_index=True, width="stretch")

    tabs = st.tabs([f"장부 ({len(prod)})", f"오류 {n_red}",
                    f"중복·변경 {n_rev}", f"확인 {n_yel}",
                    f"주문 관리 ({len(ss.orders)})"])

    # ── 장부 (편집 가능) ──
    with tabs[0]:
        st.caption("🟥 오류  ·  🟨 확인  ·  🟪 중복·변경  "
                   "(우선순위: 오류 > 중복·변경 > 확인)")
        df = pd.DataFrame([{
            "상호": (f"{r['거래처']} {r['내부표시']}".strip()
                     if r["거래처"] else ""),
            "색상": r["색상"] or "", "가로": r["가로"], "X": "X",
            "세로": r["세로"], "수량": r["수량"] or "",
            "방향": r["모형1"] or "", "길이": r["모형2"] or "",
            "특이": r.get("특이") or "",
            "기재사항": r["기재사항"] or "", "기재사항2": r["기재사항2"] or "",
            "출고일": r.get("_출고일날짜") or "",
        } for r in prod]).astype(str).replace("None", "")

        df.insert(0, "선택", False)
        df.insert(1, "삭제", False)
        df.insert(2, "원본 보기", False)
        df.insert(3, "상태", [status_text.get(row_levels.get(k), "")
                            for k in range(len(prod))])
        editor_client_labels = list(dict.fromkeys(
            [""] + CLIENT_LABELS + [x for x in df["상호"].tolist() if x]))

        editor_key = f"ledger_editor_{ss.ledger_editor_version}"
        detail_picker_key = f"detail_picker_{editor_key}"

        def select_original_row():
            state = ss.get(editor_key) or {}
            changes = state.get("edited_rows", {}) if isinstance(state, dict) else {}
            # 숫자만 입력한 품목은 widget 상태 안에서 바로 ` B 200`으로
            # 바꾼다. 새 editor를 만들지 않으므로 현재 행·스크롤은 유지된다.
            for value in changes.values():
                if "색상" in value:
                    value["색상"] = normalize_editor_color(
                        value["색상"], ss.manual_item_prefix or "B")

            # v66에서는 detail_row_idx만 바꿨고 selectbox의 widget key는 예전 값을
            # 계속 기억해서 안내문과 달리 자동 연결이 되지 않았다.
            # 현재 체크된 행과 직전 체크 목록을 비교해 '새로 체크한 행'을 활성 행으로
            # 잡고, 아래 부가정보 selectbox의 실제 session-state key도 함께 갱신한다.
            checked_rows = sorted(
                int(k) for k, value in changes.items()
                if isinstance(value, dict) and value.get("선택") is True)
            previous = set(int(x) for x in (ss.get("editor_selected_rows") or []))
            newly_checked = [k for k in checked_rows if k not in previous]
            target = newly_checked[-1] if newly_checked else None
            if target is None and len(checked_rows) == 1:
                target = checked_rows[0]
            if (target is None and checked_rows and
                    ss.get("detail_row_idx") not in checked_rows):
                target = checked_rows[-1]
            ss.editor_selected_rows = checked_rows
            if target is not None and 0 <= target < len(prod):
                ss.detail_row_idx = target
                ss[detail_picker_key] = target

            candidates = [int(k) for k, value in changes.items()
                          if value.get("원본 보기") is True]
            if not candidates:
                return
            k = next((x for x in reversed(candidates)
                      if x != ss.get("issue_selected_row")), candidates[-1])
            if not (0 <= k < len(prod)):
                return
            ss.issue_selected_row = k
            selected_ledger = df.iloc[k].drop(labels=["선택", "삭제", "원본 보기"]).to_dict()
            row_details = row_issue_details.get(k, [])
            row_status = status_text.get(row_levels.get(k), "정상")
            ss.issue_view = {
                "order": prod[k]["_oi"],
                "제목": f"{k + 1}행 · {row_status}",
                "장부행": selected_ledger,
                "원인": [{"등급": x["등급"], "코드": x["코드"],
                         "내용": x["내용"]} for x in row_details],
            }
            # 원본 보기 클릭만으로 data_editor key를 바꾸면 표가 새로 생성되면서
            # 사전 점검 화면의 스크롤이 맨 위로 튄다. widget key는 유지하고,
            # 현재 editor 상태에서 선택 체크만 되돌려 사용자가 보던 위치를 보존한다.
            try:
                checked = changes.get(k) if k in changes else changes.get(str(k))
                if isinstance(checked, dict):
                    checked["원본 보기"] = False
            except Exception:
                pass

        # Streamlit 표에서 마우스로 늘린 열 너비는 셀 편집 재실행 시 브라우저가
        # 초기화할 수 있으므로, 자주 조정하는 열은 세션에 저장되는 고정값으로 관리한다.
        with st.expander("사전점검 열 너비 고정", expanded=False):
            w1, w2, w3, w4 = st.columns(4)
            ss.editor_col_widths["색상"] = int(w1.slider(
                "품목/색상", 110, 240, int(ss.editor_col_widths.get("색상", 150)), 5,
                key="editor_width_color"))
            ss.editor_col_widths["특이"] = int(w2.slider(
                "특이", 65, 180, int(ss.editor_col_widths.get("특이", 85)), 5,
                key="editor_width_special"))
            ss.editor_col_widths["기재사항"] = int(w3.slider(
                "기재사항1", 95, 240, int(ss.editor_col_widths.get("기재사항", 145)), 5,
                key="editor_width_note1"))
            ss.editor_col_widths["기재사항2"] = int(w4.slider(
                "기재사항2", 95, 240, int(ss.editor_col_widths.get("기재사항2", 145)), 5,
                key="editor_width_note2"))

        selected = ss.get("issue_selected_row")
        styled = df.style.apply(
            lambda row: ["background-color: rgba(255, 80, 80, 0.18)"
                         if row.name == selected and col == "상태" else ""
                         for col in row.index], axis=1)
        edited = st.data_editor(
            styled, hide_index=True, width=1260, num_rows="fixed",
            key=editor_key, on_change=select_original_row,
            disabled=["상태", "X"],
            column_config={
                "선택": st.column_config.CheckboxColumn(
                    "선택", help="행 추가·묶음 이동·부가정보 편집 대상을 선택",
                    width="small"),
                "삭제": st.column_config.CheckboxColumn(
                    "삭제", help="체크 후 아래 버튼을 누르면 그 행이 빠집니다",
                    width="small"),
                "원본 보기": st.column_config.CheckboxColumn(
                    "원본", help="선택한 행과 발주서 원본을 왼쪽에 표시",
                    width="small"),
                "상태": st.column_config.TextColumn("상태", width=70),
                "상호": st.column_config.SelectboxColumn(
                    "상호", options=editor_client_labels,
                    required=False, width=110),
                "출고일": st.column_config.TextColumn(
                    "출고일", help="주문 첫 행에 YYYY-MM-DD 형식으로 입력",
                    width=95),
                "색상": st.column_config.TextColumn(
                    "색상", width=int(ss.editor_col_widths["색상"])),
                "가로": st.column_config.TextColumn("가로", width=68),
                "세로": st.column_config.TextColumn("세로", width=68),
                "수량": st.column_config.TextColumn("수량", width=60),
                "방향": st.column_config.TextColumn("방향", width=60),
                "길이": st.column_config.TextColumn("길이", width=68),
                "특이": st.column_config.TextColumn(
                    "특이", width=int(ss.editor_col_widths["특이"])),
                "기재사항": st.column_config.TextColumn(
                    "기재사항", width=int(ss.editor_col_widths["기재사항"])),
                "기재사항2": st.column_config.TextColumn(
                    "기재사항2", width=int(ss.editor_col_widths["기재사항2"])),
            })
        selected_rows = [k for k in range(len(prod))
                         if edited.iloc[k]["선택"] in (True, "True")]
        picked = [k for k in range(len(prod))
                  if edited.iloc[k]["삭제"] in (True, "True")]

        # 행 추가 / 주문 묶음 이동 / 최근 변경 복구
        ctl1, ctl2, ctl3, ctl4 = st.columns(4)
        if ctl1.button("+ 선택 행 아래 추가", width="stretch"):
            if len(selected_rows) != 1:
                st.warning("행 하나만 선택해 주세요.")
            else:
                k = selected_rows[0]
                r = prod[k]
                o = ss.orders[r["_oi"]]
                ii = r["_ii"]
                if not isinstance(ii, int) or not (0 <= ii < len(o.get("items", []))):
                    st.warning("이 행에는 새 품목을 추가할 수 없습니다.")
                else:
                    _push_undo("행 추가")
                    o["items"].insert(ii + 1, _blank_row_like(o["items"][ii]))
                    _reset_editor_selection()
                    st.rerun()

        order_selection = sorted({prod[k]["_oi"] for k in selected_rows})
        if ctl2.button("묶음 ↑", width="stretch"):
            if not order_selection:
                st.warning("옮길 주문의 행을 선택해 주세요.")
            elif order_selection != list(range(order_selection[0], order_selection[-1] + 1)):
                st.warning("서로 붙어 있는 주문 묶음만 함께 이동할 수 있습니다.")
            elif order_selection[0] == 0:
                st.info("이미 맨 위 묶음입니다.")
            else:
                _push_undo("묶음 위로 이동")
                _move_order_block(order_selection, -1)
                _reset_editor_selection()
                st.rerun()
        if ctl3.button("묶음 ↓", width="stretch"):
            if not order_selection:
                st.warning("옮길 주문의 행을 선택해 주세요.")
            elif order_selection != list(range(order_selection[0], order_selection[-1] + 1)):
                st.warning("서로 붙어 있는 주문 묶음만 함께 이동할 수 있습니다.")
            elif order_selection[-1] >= len(ss.orders) - 1:
                st.info("이미 맨 아래 묶음입니다.")
            else:
                _push_undo("묶음 아래로 이동")
                _move_order_block(order_selection, 1)
                _reset_editor_selection()
                st.rerun()
        undo_label = ss.undo_stack[-1][0] if ss.undo_stack else None
        if ctl4.button("최근 변경 복구" + (f" · {undo_label}" if undo_label else ""),
                       disabled=not ss.undo_stack, width="stretch"):
            _, previous = ss.undo_stack.pop()
            ss.orders = previous
            _reset_editor_selection()
            st.rerun()

        if picked and st.button(f"선택한 {len(picked)}행 삭제", type="primary"):
            _push_undo("행 삭제")
            # 같은 주문에서 여러 행을 지울 때 실제 item index가 밀리지 않도록 주문별 역순 삭제.
            grouped = {}
            for k in picked:
                r = prod[k]
                grouped.setdefault(r["_oi"], set()).add(r["_ii"])
            for oi, item_indices in sorted(grouped.items(), reverse=True):
                o = ss.orders[oi]
                for source_i in sorted((x for x in item_indices if isinstance(x, int)), reverse=True):
                    if 0 <= source_i < len(o.get("items", [])):
                        o["items"].pop(source_i)
            ss.orders = [o for o in ss.orders if o.get("items")]
            _reset_editor_selection()
            st.rerun()

        compare_cols = [c for c in df.columns if c not in {"선택", "원본 보기", "삭제"}]
        if not edited[compare_cols].equals(df[compare_cols]):
            _push_undo("셀 수정")
            for k, r in enumerate(prod):
                for col in EDIT_COLS:
                    if col in edited.columns:
                        new = edited.iloc[k][col]
                        if str(new) != str(df.iloc[k][col]):
                            apply_edit(ss.orders[r["_oi"]], r["_ii"], col, new,
                                       ss.manual_item_prefix or "B")
                            synchronize_order_for_outputs(ss.orders[r["_oi"]])
            # 편집값은 이미 주문 상태에 반영됐다. 추가 st.rerun을 호출하면
            # data_editor의 스크롤이 위로 튀므로 현재 표 위치를 유지한다.

        # 주소/추가부속은 사전점검에서 항상 보이도록 한다.
        # 체크박스를 하나 선택하면 그 행으로 자동 이동하고, 선택하지 않아도 아래 선택기로
        # 원하는 주문을 골라 확인/수정할 수 있다.
        if prod:
            if len(selected_rows) == 1:
                target = selected_rows[0]
                ss.detail_row_idx = target
                # keyed selectbox는 index보다 기존 session state를 우선하므로
                # 실제 widget key도 동기화해야 즉시 해당 주문으로 바뀐다.
                ss[detail_picker_key] = target
            if not isinstance(ss.get("detail_row_idx"), int) or not (0 <= ss.detail_row_idx < len(prod)):
                ss.detail_row_idx = 0
                ss[detail_picker_key] = 0

            def _detail_label(idx):
                rr = prod[idx]
                vendor = (f"{rr.get('거래처') or ''} {rr.get('내부표시') or ''}").strip() or "-"
                color = str(rr.get("색상") or "-")
                return f"{idx + 1}행 · {vendor} · {color}"

            st.markdown("**사전점검 부가정보 — 주소 · 추가부속 · 전달사항 · 손잡이길이**")
            chosen_k = st.selectbox(
                "확인할 주문 행", list(range(len(prod))),
                index=ss.detail_row_idx, format_func=_detail_label,
                key=detail_picker_key,
                help="위 표에서 '선택'을 체크하면 그 주문으로 즉시 자동 연결됩니다.")
            ss.detail_row_idx = int(chosen_k)
            r = prod[ss.detail_row_idx]
            order = ss.orders[r["_oi"]]
            delivery = order.setdefault("배송", {})
            item = (order.get("items") or [])[r["_ii"]] if isinstance(r.get("_ii"), int) and 0 <= r["_ii"] < len(order.get("items") or []) else None
            with st.container(border=True):
                address_value = st.text_input(
                    "주소", value=_detail_address(order),
                    key=f"detail_addr_{editor_key}_{r['_oi']}_{r['_ii']}",
                    help="장부에 쓰이는 주소 원문입니다. 경영박사 출력 때만 전산용 주소 후처리가 적용됩니다.")
                accessory_value = st.text_input(
                    "추가부속", value=_accessory_summary(order),
                    help="노피스·콘크리트·석고·앙카·트루 부속 등을 한 줄에서 확인/수정",
                    key=f"detail_accessory_{editor_key}_{r['_oi']}_{r['_ii']}")
                d1, d2, d3 = st.columns([2, 2, 1])
                name_value = d1.text_input(
                    "이름/수령인", value=str(order.get("고객명") or delivery.get("수령인") or ""),
                    key=f"detail_name_{editor_key}_{r['_oi']}_{r['_ii']}")
                phone_value = d2.text_input(
                    "연락처", value=str(delivery.get("연락처") or ""),
                    key=f"detail_phone_{editor_key}_{r['_oi']}_{r['_ii']}")
                order_no_value = d3.text_input(
                    "발주번호", value=str(order.get("주문번호") or ""),
                    key=f"detail_no_{editor_key}_{r['_oi']}_{r['_ii']}")
                common_value = st.text_input(
                    "공통 기재사항", value=str(order.get("전체기재사항") or ""),
                    key=f"detail_common_{editor_key}_{r['_oi']}_{r['_ii']}")
                n1, n2 = st.columns([4, 1])
                delivery_notice_value = n1.text_input(
                    "전달사항", value=clean_delivery_notice(delivery.get("전달사항")) or "",
                    key=f"detail_notice_{editor_key}_{r['_oi']}_{r['_ii']}",
                    help="기본 문구인 배송전연락/던지지마세요는 자동 제외하고 실제 전달사항만 세 문서에 반영합니다.")
                _explicit_handle = (item or {}).get("손잡이길이")
                _effective_handle = _explicit_handle
                if _effective_handle in (None, "") and item is not None:
                    _effective_handle = default_handle_length(item.get("종류"), item.get("세로"))
                handle_value = n2.text_input(
                    "손잡이길이", value=str(_effective_handle or ""),
                    key=f"detail_handle_{editor_key}_{r['_oi']}_{r['_ii']}",
                    help="명시값이 없으면 계산 가능한 기본 길이도 보여줍니다. 수정한 값은 장부·작업지시서·EDI에 같이 반영됩니다.")
                p1, p2 = st.columns([1, 5])
                piece_value = p1.checkbox(
                    "피스", value="피스" in _note_parts(order.get("전체기재사항")),
                    key=f"detail_piece_{editor_key}_{r['_oi']}_{r['_ii']}")
                pay_options = ["", "선불", "착불"]
                current_pay = str(delivery.get("선불착불") or "")
                if current_pay not in pay_options:
                    current_pay = ""
                pay_value = p2.selectbox(
                    "결제표기", pay_options, index=pay_options.index(current_pay),
                    key=f"detail_pay_{editor_key}_{r['_oi']}_{r['_ii']}")
                if st.button("부가정보 적용", key=f"detail_apply_{editor_key}_{r['_oi']}_{r['_ii']}",
                             type="secondary", width="stretch"):
                    _push_undo("부가정보 수정")
                    delivery["주소"] = address_value.strip() or None
                    order["_추가부속"] = accessory_value.strip() or None
                    name_clean = name_value.strip() or None
                    order["고객명"] = name_clean
                    delivery["수령인"] = name_clean
                    delivery["연락처"] = phone_value.strip() or None
                    order["주문번호"] = order_no_value.strip() or None
                    order["전체기재사항"] = common_value.strip(" /") or None
                    delivery["전달사항"] = clean_delivery_notice(delivery_notice_value)
                    if item is not None:
                        hm = re.search(r"(\d{2,3})", str(handle_value or ""))
                        item["손잡이길이"] = int(hm.group(1)) if hm else None
                    _set_piece_flag(order, piece_value)
                    delivery["선불착불"] = pay_value or None
                    synchronize_order_for_outputs(order)
                    _reset_editor_selection()
                    st.rerun()

        st.caption("특수 행(#포장비용, ☆주소, 연락처, 전달사항)은 다운로드 시 자동 삽입됩니다. "
                   "사전점검 열 너비는 화면에만 적용되고 Excel 셀 폭에는 영향을 주지 않습니다. "
                   "행 추가·묶음 이동·삭제 복구와 부가정보 수정값도 장부·작업지시서·경영박사에 모두 반영됩니다.")

    # ── 오류 / 확인 ──
    def issue_panel(level, label):
        sel = [i for i in live if i["등급"] == level]
        if not sel:
            st.success(f"{label} 항목이 없습니다.")
            return
        cols = ["파일", "행", "코드", "내용"]
        st.dataframe(pd.DataFrame(sel)[cols], hide_index=True,
                     width="stretch")
        with st.form(f"dismiss_{level}"):
            pick = st.multiselect(
                "해제할 항목", [i["_key"] for i in sel],
                format_func=lambda k: next(
                    f"{i['파일']} · {i['코드']}" for i in sel if i["_key"] == k))
            why = st.selectbox("사유", ["신규 품목", "특수 주문",
                                        "거래처 확인함", "기타"])
            if st.form_submit_button("선택 해제") and pick:
                ss.dismissed |= set(pick)
                st.rerun()

    with tabs[1]:
        issue_panel("red", "오류")
    with tabs[2]:
        issue_panel("review", "중복·변경")
    with tabs[3]:
        issue_panel("yellow", "확인")
    with tabs[4]:
        st.caption("잘못 들어온 주문을 삭제할 수 있습니다.")
        for oi, o in enumerate(ss.orders):
            n = len([x for x in o.get("items", []) if not x.get("예외품목")])
            c1, c2 = st.columns([5, 1])
            with c1:
                with st.expander(
                        f"{oi + 1}. {o.get('_file')} — "
                        f"{o.get('거래처') or '거래처 불명'} · {n}행"):
                    show_source(o, f"order_{oi}")
                    st.json(o, expanded=False)
            if c2.button("삭제", key=f"del_{oi}", width="stretch"):
                ss.orders.pop(oi)
                ss.issue_view = None
                ss.batch_id = None
                st.rerun()

    if ss.dismissed:
        st.caption(f"수동 해제 {len(ss.dismissed)}건")

    st.divider()
    if n_red:
        st.warning(f"오류 {n_red}건이 있습니다. 내용 확인 후 그대로 출력할 수 있습니다.")
    if st.button("파일 생성 (장부 · 작업지시서 · 경영박사)",
                 type="primary", width="stretch"):
        OUT.mkdir(parents=True, exist_ok=True)
        for _order in ss.orders:
            synchronize_order_for_outputs(_order)
        bid = ss.batch_id or save_batch(ss.ship, ss.orders)
        ss.batch_id = bid
        # 편집 중에는 행 위치를 유지하고, 실제 파일을 만들 때만 DI의
        # 수령인·실색·품목 순서 규칙을 적용한다.
        output_rows = rebuild_rows(ss.orders, ss.ship, M, sort_di=True)
        build_ledger(output_rows, OUT / "장부.xlsx")
        build_worksheet(output_rows, OUT / "작업지시서.xlsx")
        build_erp(ss.orders, M, OUT / "경영박사_EDI.xls", ss.ship)
        ss.output_names = reserve_output_names(ss.orders)
        archive_ledger(ss.output_names["장부"])
        mark_printed(bid, ss.orders)
        # 감시 중이라면 현재 클립보드를 본 것으로 등록해 재분석을 막는다
        if ss.get("watch"):
            try:
                _, _h = clip.grab()
                if _h:
                    ss.clip_seen.add(_h)
                    ss.clip_last = _h
            except Exception:
                pass
        ss.done = len(prod)
        ss.orders = []
        ss.issue_view = None
        ss.dismissed, ss.seen = set(), set()
        # clip_seen · clip_last 는 유지한다.
        # 초기화하면 클립보드에 남아 있는 캡처가 다시 분석된다.
        ss.batch_id = None
        st.rerun()

    if (OUT / "장부.xlsx").exists():
        MIME = ("application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet")
        d1, d2, d3 = st.columns(3)
        names = ss.get("output_names") or {}
        files = ((d1, "장부", OUT / "장부.xlsx", MIME, names.get("장부")),
                 (d2, "작업지시서", OUT / "작업지시서.xlsx", MIME,
                  names.get("작업지시서")),
                 (d3, "경영박사 EDI", OUT / "경영박사_EDI.xls",
                  "application/vnd.ms-excel", names.get("경영박사")))
        for col, name, p, mime, download_name in files:
            if p.exists():
                col.download_button(p.name, p.read_bytes(),
                                    file_name=download_name or p.name, mime=mime,
                                    width="stretch")
