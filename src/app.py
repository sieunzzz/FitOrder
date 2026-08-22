"""
FitOrder 메인 화면

    streamlit run app.py
"""
import base64
import io
import json
import re
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image

from output import (apply_edit, build_erp, build_ledger, build_worksheet,
                    read_ledger, to_rows)
import clipboard_watch as clip
from parsers import parse_excel
from rules import (CHANGE_LOOKBACK_DAYS, CLIENT_INFO, DUP_LOOKBACK_DAYS,
                   Master, find_changes, find_duplicates, validate)

ROOT = Path(__file__).parent.parent
MASTER = ROOT / "data" / "master" / "fitorder_master.xlsx"
LOGO = ROOT / "data" / "logo.png"
DB = ROOT / "db" / "fitorder.db"
OUT = ROOT / "out"
CLIENTS = list(CLIENT_INFO)
APP_VERSION = "2026.08.22-26"

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
             "기재사항", "기재사항2"]


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
    "스페이스": "SP",
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
def rebuild_rows(orders, ship, M=None):
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
        got = to_rows(o, ship_label(order_ship, mode), M)
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
        ii = 0
        for r in got:
            if r.get("_특수") is None:
                r["_oi"], r["_ii"] = oi, ii
                ii += 1
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
        # 현재 PDF 전용 거래처는 SP다. IMG_0001.pdf처럼
        # 업체명이 없는 스캔 파일도 자동으로 SP 파서로 보낸다.
        # 사용자가 거래처를 직접 선택했으면 그 값을 우선한다.
        c = client_hint or client_from_filename(path) or "SP"
        got = extract_pdf(path,
                          client_hint=c if c in CLIENTS else None)
    else:
        from extract import extract_order
        c = client_hint or client_from_filename(path)
        got = [extract_order([path],
                             client_hint=c if c in CLIENTS else None)]
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
        if o.get("거래처") == "아지트":
            raw = " ".join(str(x or "") for x in (
                o.get("전체원문"), o.get("전체기재사항")))
            o["_az_place"] = "금빛커텐" if "에어캡+포장" in raw else "시온가공소"
            common = str(o.get("전체기재사항") or "")
            common = re.sub(r"(?:^|/)☆?(?:금빛커텐|시온가공소)(?=/|$)", "", common)
            o["전체기재사항"] = common.strip("/") or None
            o.setdefault("배송", {})["방식"] = "배달"
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

    analyze_col, ledger_col = st.columns(2)
    analyze_clicked = analyze_col.button(
        "분석 시작", type="primary", width="stretch", disabled=not files)
    ledger_clicked = ledger_col.button(
        "장부 변환", width="stretch", disabled=not files,
        help="FitOrder 장부.xlsx를 작업지시서와 경영박사 EDI로 변환합니다.")

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
                item_count = sum(len(o.get("items", [])) for o in ledger_orders)
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
                    build_worksheet(ledger_rows, OUT / "작업지시서.xlsx", ss.ship)
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
            downloads = ((d1, OUT / "장부.xlsx", MIME),
                         (d2, OUT / "작업지시서.xlsx", MIME),
                         (d3, OUT / "경영박사_EDI.xls",
                          "application/vnd.ms-excel"))
            for col, fp, mime in downloads:
                if fp.exists():
                    col.download_button(fp.name, fp.read_bytes(),
                                        file_name=fp.name, mime=mime,
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
            visible_i = sum(1 for it in items[:source_i + 1]
                            if not it.get("예외품목")) - 1
            for k, row in enumerate(prod):
                if row.get("_oi") == oi and row.get("_ii") == visible_i:
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
            "출고일": r["출고일"] or "",
        } for r in prod]).astype(str).replace("None", "")

        df.insert(0, "삭제", False)
        df.insert(1, "원본 보기", False)
        df.insert(2, "상태", [status_text.get(row_levels.get(k), "")
                            for k in range(len(prod))])
        editor_client_labels = list(dict.fromkeys(
            [""] + CLIENT_LABELS + [x for x in df["상호"].tolist() if x]))

        editor_key = f"ledger_editor_{ss.ledger_editor_version}"

        def select_original_row():
            state = ss.get(editor_key) or {}
            changes = state.get("edited_rows", {}) if isinstance(state, dict) else {}
            candidates = [int(k) for k, value in changes.items()
                          if value.get("원본 보기") is True]
            if not candidates:
                return
            k = next((x for x in reversed(candidates)
                      if x != ss.get("issue_selected_row")), candidates[-1])
            if not (0 <= k < len(prod)):
                return
            ss.issue_selected_row = k
            selected_ledger = df.iloc[k].drop(labels=["삭제", "원본 보기"]).to_dict()
            row_details = row_issue_details.get(k, [])
            row_status = status_text.get(row_levels.get(k), "정상")
            ss.issue_view = {
                "order": prod[k]["_oi"],
                "제목": f"{k + 1}행 · {row_status}",
                "장부행": selected_ledger,
                "원인": [{"등급": x["등급"], "코드": x["코드"],
                         "내용": x["내용"]} for x in row_details],
            }
            # 원본은 옆에 계속 표시하되 선택용 체크는 즉시 해제한다.
            # 삭제 체크박스에는 영향을 주지 않는다.
            row_change = changes.get(str(k), changes.get(k))
            if isinstance(row_change, dict):
                row_change["원본 보기"] = False
            # data_editor를 새 키로 다시 만들어 화면의 체크 표시도 즉시 없앤다.
            ss.ledger_editor_version += 1

        selected = ss.get("issue_selected_row")
        styled = df.style.apply(
            lambda row: ["background-color: rgba(255, 80, 80, 0.18)"
                         if row.name == selected and col == "상태" else ""
                         for col in row.index], axis=1)
        edited = st.data_editor(
            styled, hide_index=True, width="stretch", num_rows="fixed",
            key=editor_key, on_change=select_original_row,
            disabled=["상태", "X", "출고일"],
            column_config={
                "삭제": st.column_config.CheckboxColumn(
                    "삭제", help="체크 후 아래 버튼을 누르면 그 행이 빠집니다",
                    width="small"),
                "원본 보기": st.column_config.CheckboxColumn(
                    "원본", help="선택한 행과 발주서 원본을 왼쪽에 표시",
                    width="small"),
                "상태": st.column_config.TextColumn("상태", width="small"),
                "상호": st.column_config.SelectboxColumn(
                    "상호", options=editor_client_labels,
                    required=False, width="small"),
                **{c: st.column_config.TextColumn(width="small")
                   for c in ("가로", "세로", "수량", "방향", "길이", "특이")}})
        picked = [k for k in range(len(prod))
                  if edited.iloc[k]["삭제"] in (True, "True")]
        if picked and st.button(f"선택한 {len(picked)}행 삭제", type="primary"):
            for k in sorted(picked, reverse=True):
                r = prod[k]
                o = ss.orders[r["_oi"]]
                visible = [i for i in o["items"] if not i.get("예외품목")]
                if r["_ii"] < len(visible):
                    o["items"].remove(visible[r["_ii"]])
            ss.orders = [o for o in ss.orders if o.get("items")]
            ss.issue_view = None
            ss.issue_selected_row = None
            ss.batch_id = None
            st.rerun()
        compare_cols = [c for c in df.columns if c != "원본 보기"]
        if not edited[compare_cols].equals(df[compare_cols]):
            for k, r in enumerate(prod):
                for col in EDIT_COLS:
                    if col in edited.columns:
                        new = edited.iloc[k][col]
                        if str(new) != str(df.iloc[k][col]):
                            apply_edit(ss.orders[r["_oi"]], r["_ii"], col, new,
                                       ss.manual_item_prefix or "B")
            st.rerun()

        st.caption("특수 행(#포장비용, ☆주소, 연락처)은 다운로드 시 자동 삽입됩니다. "
                   "값을 고치면 장부·작업지시서·경영박사에 모두 반영됩니다.")

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
        st.warning(f"오류 {n_red}건을 수정하거나 해제해야 파일을 생성할 수 있습니다.")
    if st.button("파일 생성 (장부 · 작업지시서 · 경영박사)",
                 type="primary", width="stretch",
                 disabled=bool(n_red)):
        OUT.mkdir(parents=True, exist_ok=True)
        bid = ss.batch_id or save_batch(ss.ship, ss.orders)
        ss.batch_id = bid
        build_ledger(rows, OUT / "장부.xlsx", ss.ship)
        build_worksheet(rows, OUT / "작업지시서.xlsx", ss.ship)
        build_erp(ss.orders, M, OUT / "경영박사_EDI.xls", ss.ship)
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
        files = ((d1, "장부", OUT / "장부.xlsx", MIME),
                 (d2, "작업지시서", OUT / "작업지시서.xlsx", MIME),
                 (d3, "경영박사 EDI", OUT / "경영박사_EDI.xls",
                  "application/vnd.ms-excel"))
        for col, name, p, mime in files:
            if p.exists():
                col.download_button(p.name, p.read_bytes(),
                                    file_name=p.name, mime=mime,
                                    width="stretch")
