"""
FitOrder 메인 화면

    streamlit run app.py
"""
import base64
import json
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

from output import (apply_edit, build_erp, build_ledger, build_worksheet,
                    to_rows)
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

EDIT_COLS = ["색상", "가로", "세로", "수량", "모형1", "모형2",
             "기재사항", "기재사항2"]


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
    return f"{lab}({mode})" if mode in ("택배", "화물", "내사") else lab


# ─────────────────────────────────────────────
# 장부 행 계산 (주문 -> 행 + 역참조)
# ─────────────────────────────────────────────
def rebuild_rows(orders, ship, M=None):
    rows = []
    for oi, o in enumerate(orders):
        mode = (o.get("배송") or {}).get("방식") \
               or CLIENT_INFO.get(o.get("거래처"), (None, None, None))[2]
        got = to_rows(o, ship_label(ship, mode), M)
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
        c = client_hint or ("DI" if "DI" in path.name.upper() else "휴안")
        got = parse_excel(path, c)
    else:
        from extract import extract_order
        c = client_hint or path.stem.split("_")[0]
        got = [extract_order([path],
                             client_hint=c if c in CLIENTS else None)]
    for o in got:
        o["_file"] = label
        st.session_state.orders.append(o)
    return len(got)


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

M = Master(MASTER) if MASTER.exists() else None
left, right = st.columns([3, 7], gap="medium")

# ─────────────────────────────────────────────
# 좌측
# ─────────────────────────────────────────────
with left:
    if LOGO.exists():
        st.image(str(LOGO), width=110)

    st.caption("출고일")
    new_ship = st.date_input("출고일", ss.ship, label_visibility="collapsed")
    if new_ship != ss.ship:
        ss.ship = new_ship
    st.caption(f"{WD[ss.ship.weekday()]}요일"
               + ("  ·  3시 전 익일" if datetime.now().hour < 15
                  else "  ·  3시 이후 익익일"))

    # 연결 상태 자가 진단 (10분마다 재확인)
    import time as _t
    if ss.get("net_at", 0) < _t.time() - 600:
        try:
            from extract import diagnose
            ss.net_msg = diagnose()
        except Exception as e:
            ss.net_msg = f"프로그램 오류: {e}"
        ss.net_at = _t.time()
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
        type=["png", "jpg", "jpeg", "xlsx", "xls"],
        accept_multiple_files=True,
        key=f"upl_{ss.upl}")

    if st.button("분석 시작", type="primary", width="stretch",
                 disabled=not files):
        client = None if hint == "자동 판별" else hint
        todo = [f for f in files
                if (f.name, f.size) not in ss.seen]      # 이미 처리한 건 제외
        skipped = len(files) - len(todo)
        if not todo:
            st.info("새로 추가된 발주서가 없습니다.")
        bar = st.progress(0.0, "준비 중")
        for i, f in enumerate(todo, 1):
            bar.progress((i - 1) / len(todo), f"{f.name} 분석 중")
            ss.seen.add((f.name, f.size))
            tmp = OUT / "_upload" / f.name
            tmp.parent.mkdir(parents=True, exist_ok=True)
            tmp.write_bytes(f.getbuffer())
            try:
                ingest(tmp, client, f.name)
            except Exception as e:
                st.error(f"**{f.name} 분석 실패**\n\n{e}")
                ss.net_at = 0
            bar.progress(i / len(todo))
        bar.empty()
        if todo:
            ss.upl += 1
        if skipped:
            st.caption(f"이미 처리한 {skipped}건은 건너뛰었습니다.")
        if todo:
            st.rerun()

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
            ss.dismissed, ss.seen = set(), set()
            ss.clip_seen, ss.clip_log, ss.clip_last = set(), [], None
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
            for col, name in ((d1, "장부"), (d2, "작업지시서"), (d3, "경영박사")):
                fp = OUT / f"{name}.xlsx"
                if fp.exists():
                    col.download_button(f"{name}.xlsx", fp.read_bytes(),
                                        file_name=fp.name, mime=MIME,
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

    # 검증
    issues, exceptions = [], []
    for oi, o in enumerate(ss.orders):
        for lv, code, msg, ri in validate(o, M):
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
        df = pd.DataFrame([{
            "상호": (f"{r['거래처']} {r['내부표시']}".strip()
                     if r["거래처"] else ""),
            "색상": r["색상"] or "", "가로": r["가로"], "X": "X",
            "세로": r["세로"], "수량": r["수량"] or "",
            "모형1": r["모형1"] or "", "모형2": r["모형2"] or "",
            "기재사항": r["기재사항"] or "", "기재사항2": r["기재사항2"] or "",
            "출고일": r["출고일"] or "",
        } for r in prod]).astype(str).replace("None", "")

        df.insert(0, "삭제", False)
        edited = st.data_editor(
            df, hide_index=True, width="stretch", num_rows="fixed",
            key="ledger_editor",
            disabled=["상호", "X", "출고일"],
            column_config={
                "삭제": st.column_config.CheckboxColumn(
                    "삭제", help="체크 후 아래 버튼을 누르면 그 행이 빠집니다",
                    width="small"),
                **{c: st.column_config.TextColumn(width="small")
                   for c in ("가로", "세로", "수량", "모형1", "모형2")}})

        picked = [k for k in range(len(prod)) if edited.iloc[k]["삭제"] in (True, "True")]
        if picked:
            if st.button(f"선택한 {len(picked)}행 삭제", type="primary"):
                for k in sorted(picked, reverse=True):
                    r = prod[k]
                    o = ss.orders[r["_oi"]]
                    live = [i for i in o["items"] if not i.get("예외품목")]
                    if r["_ii"] < len(live):
                        o["items"].remove(live[r["_ii"]])
                ss.orders = [o for o in ss.orders if o.get("items")]
                ss.batch_id = None
                st.rerun()

        if not edited.equals(df):
            for k, r in enumerate(prod):
                for col in EDIT_COLS:
                    if col == "삭제":
                        continue
                    if col not in edited.columns:
                        continue
                    new = edited.iloc[k][col]
                    if str(new) != str(df.iloc[k][col]):
                        apply_edit(ss.orders[r["_oi"]], r["_ii"], col, new)
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
                    st.json(o, expanded=False)
            if c2.button("삭제", key=f"del_{oi}", width="stretch"):
                ss.orders.pop(oi)
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
        build_erp(ss.orders, M, OUT / "경영박사.xlsx", ss.ship)
        mark_printed(bid, ss.orders)
        ss.done = len(prod)
        ss.orders = []
        ss.dismissed, ss.seen, ss.clip_seen = set(), set(), set()
        ss.clip_log, ss.clip_last = [], None
        ss.batch_id = None
        st.rerun()

    if (OUT / "장부.xlsx").exists():
        MIME = ("application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet")
        d1, d2, d3 = st.columns(3)
        for col, name in ((d1, "장부"), (d2, "작업지시서"), (d3, "경영박사")):
            p = OUT / f"{name}.xlsx"
            if p.exists():
                col.download_button(f"{name}.xlsx", p.read_bytes(),
                                    file_name=p.name, mime=MIME,
                                    width="stretch")