"""FitOrder PySide6 데스크톱 UI — v100.

핵심 원칙
- 기존 AI/파서/규칙/openpyxl 출력 엔진은 유지한다.
- 기존 Streamlit 작업 흐름(3:7, 붙여넣기/자동 캡처/파일 업로드)을 PySide6로 유지한다.
- 사전점검을 '장부 미리보기 + 직접 수정'으로 보여준다.
- 행번호/상태를 장부 앞에 붙이고 상태 셀 클릭으로 원본을 연다.
- 주소/수령인/연락처/전달사항/포장 같은 장부 특수행을 표 안에 항상 보이게 한다.
- Alt+1 병합 / Alt+2 해제 / Ctrl+Z / Delete / F2·더블클릭 편집을 지원한다.
"""
from __future__ import annotations

import copy
import os
import sys
from datetime import date
from pathlib import Path
from typing import Any

from PySide6.QtCore import QDate, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import (
    QAction,
    QColor,
    QDesktopServices,
    QFont,
    QFontMetrics,
    QImage,
    QKeySequence,
    QPainter,
    QPixmap,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QCalendarWidget,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSplitter,
    QStackedWidget,
    QStyle,
    QStyleOptionViewItem,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import clipboard_watch as clip
from output import ledger_merge_ranges, read_ledger
from rules import CLIENT_INFO, validate
from roll_combo import (ROLL_CLIENTS, ROLL_CLIENT_INFO, client_label as roll_client_label)
from master_registry import get_master_for_mode

from holding import HOLDING_CLIENTS, holding_client_label
from desktop_workflow import (
    BLIND_CLIENTS,
    CLIENTS,
    LEDGER_COLUMNS,
    Master,
    apply_excel_merge_overrides,
    apply_preview_cell_edit,
    build_preview_rows,
    client_label,
    default_ship,
    editable_columns_for_row,
    generate_outputs,
    ingest_file,
    ledger_orders_for_precheck,
    preview_statuses,
    resolve_stable_merge_specs,
    row_display_values,
)

APP_VERSION = "2026.09.13-109"
HEADERS = ["No", "상태"] + LEDGER_COLUMNS
COL = {name: i for i, name in enumerate(HEADERS)}


def project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


ROOT = project_root()
OUT_DIR = ROOT / "out"


class NoWheelDateEdit(QDateEdit):
    """출고일은 달력으로만 변경한다. 화면 스크롤/키보드로 날짜가 바뀌지 않는다."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCalendarPopup(True)
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        line = self.lineEdit()
        if line is not None:
            line.setReadOnly(True)

    def wheelEvent(self, event):
        event.ignore()

    def keyPressEvent(self, event):
        # Tab 계열은 포커스 이동을 위해 기본 처리하고, 날짜를 증감할 수 있는
        # 방향키/숫자 입력 등은 막는다. 날짜 변경은 달력 팝업에서만 한다.
        if event.key() in (Qt.Key.Key_Tab, Qt.Key.Key_Backtab, Qt.Key.Key_Escape):
            super().keyPressEvent(event)
            return
        event.ignore()


class NoWheelComboBox(QComboBox):
    """콤보박스가 닫힌 상태에서는 마우스휠로 값이 바뀌지 않게 한다."""

    def wheelEvent(self, event):
        if self.view().isVisible():
            super().wheelEvent(event)
        else:
            event.ignore()


class LedgerColorDelegate(QStyledItemDelegate):
    """색상 셀의 부분 글자색을 그린다.

    예: B 원코드_029FP -> B(검정), 원코드_(초록), 029FP(파랑)
    일반 숫자 코드는 검정으로 유지하고 P/FP가 붙은 코드는 전체를 파랑으로 그린다.
    """

    GREEN = QColor("#17823b")
    BLUE = QColor("#1769d2")
    RED = QColor("#d1242f")
    BLACK = QColor("#171717")

    @staticmethod
    def _segments(text: str):
        import re
        spans = []
        # 이전 장부 규칙과 동일: 원코드는 뒤에 '_'가 없어도 초록색.
        for m in re.finditer(r"원코드_?", text):
            spans.append((m.start(), m.end(), LedgerColorDelegate.GREEN))
        for m in re.finditer(r"셔터", text):
            spans.append((m.start(), m.end(), LedgerColorDelegate.RED))
        # P/FP 코드는 숫자와 접미사를 포함한 코드 전체를 파랑으로 표시한다.
        for m in re.finditer(r"\d{1,4}(?:FP|P)(?![A-Za-z0-9])", text, re.I):
            spans.append((m.start(), m.end(), LedgerColorDelegate.BLUE))
        for m in re.finditer(r"방염|<필증>", text):
            spans.append((m.start(), m.end(), LedgerColorDelegate.RED))
        spans.sort(key=lambda x: (x[0], x[1]))
        out = []
        pos = 0
        for start, end, color in spans:
            if start < pos:
                continue
            if start > pos:
                out.append((text[pos:start], LedgerColorDelegate.BLACK))
            out.append((text[start:end], color))
            pos = end
        if pos < len(text):
            out.append((text[pos:], LedgerColorDelegate.BLACK))
        return out or [(text, LedgerColorDelegate.BLACK)]

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index) -> None:
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        text = opt.text
        opt.text = ""
        style = opt.widget.style() if opt.widget else QApplication.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, opt, painter, opt.widget)
        if not text:
            return

        painter.save()
        painter.setFont(opt.font)
        fm = QFontMetrics(opt.font)
        rect = option.rect.adjusted(5, 1, -5, -1)
        lines = text.splitlines() or [text]
        line_h = fm.height()
        first_baseline = rect.center().y() - (len(lines) * line_h) // 2 + fm.ascent()
        for li, line in enumerate(lines):
            segments = self._segments(line)
            total = sum(fm.horizontalAdvance(t) for t, _ in segments)
            x = rect.left() + max(0, (rect.width() - total) // 2)
            baseline = first_baseline + li * line_h
            for part, color in segments:
                painter.setPen(color)
                painter.drawText(x, baseline, part)
                x += fm.horizontalAdvance(part)
        painter.restore()


class SourceViewer(QWidget):
    """상태 셀을 클릭했을 때 원본 이미지/PDF를 선명하게 표시한다."""

    close_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._path: Path | None = None
        self._pages: list[int] = []
        self._page_index = 0
        self._original_pixmap: QPixmap | None = None
        self._view_mode = "fit_width"

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)
        title = QLabel("원본 발주서")
        title.setObjectName("sourceTitle")
        layout.addWidget(title)
        self.caption = QLabel("상태 셀을 클릭하면 해당 주문의 원본이 표시됩니다.")
        self.caption.setWordWrap(True)
        layout.addWidget(self.caption)

        nav = QHBoxLayout()
        nav.setSpacing(5)
        self.prev_btn = QPushButton("◀")
        self.next_btn = QPushButton("▶")
        self.fit_btn = QPushButton("폭 맞춤")
        self.actual_btn = QPushButton("100%")
        self.open_btn = QPushButton("원본 파일 열기")
        self.large_btn = QPushButton("크게 보기")
        self.close_btn = QPushButton("✕ 원본 닫기")
        self.prev_btn.clicked.connect(self.previous_page)
        self.next_btn.clicked.connect(self.next_page)
        self.fit_btn.clicked.connect(self.fit_width)
        self.actual_btn.clicked.connect(self.actual_size)
        self.open_btn.clicked.connect(self.open_external)
        self.large_btn.clicked.connect(self.open_large)
        self.close_btn.clicked.connect(self.close_requested.emit)
        nav.addWidget(self.prev_btn)
        nav.addWidget(self.next_btn)
        nav.addWidget(self.fit_btn)
        nav.addWidget(self.actual_btn)
        nav.addStretch(1)
        nav.addWidget(self.large_btn)
        nav.addWidget(self.open_btn)
        nav.addWidget(self.close_btn)
        layout.addLayout(nav)

        self.scroll = QScrollArea()
        # 폭만 맞추고 세로는 스크롤해 작은 글자가 뭉개지지 않게 한다.
        self.scroll.setWidgetResizable(False)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.image_label = QLabel("원본 없음")
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        self.image_label.setMinimumSize(260, 420)
        self.scroll.setWidget(self.image_label)
        layout.addWidget(self.scroll, 1)
        self._refresh_nav()

    def _refresh_nav(self):
        multi = len(self._pages) > 1
        self.prev_btn.setVisible(multi)
        self.next_btn.setVisible(multi)
        self.prev_btn.setEnabled(multi and self._page_index > 0)
        self.next_btn.setEnabled(multi and self._page_index < len(self._pages) - 1)
        self.open_btn.setEnabled(bool(self._path and self._path.exists()))
        ok = self._original_pixmap is not None and not self._original_pixmap.isNull()
        self.large_btn.setEnabled(ok)
        self.fit_btn.setEnabled(ok)
        self.actual_btn.setEnabled(ok)

    def show_order(self, order: dict):
        label = str(order.get("_file") or "")
        path = Path(str(order.get("_source_path") or ""))
        self._path = path if str(path) else None
        self._pages = []
        self._page_index = 0
        self._original_pixmap = None
        self._view_mode = "fit_width"

        if not self._path or not self._path.exists():
            self.caption.setText(f"{label}\n원본 파일을 찾을 수 없습니다.")
            self.image_label.setText("원본 파일 없음")
            self.image_label.setPixmap(QPixmap())
            self._refresh_nav()
            return

        suffix = self._path.suffix.lower()
        if suffix in {".png", ".jpg", ".jpeg", ".bmp", ".webp"}:
            pix = QPixmap(str(self._path))
            self._set_pixmap(pix)
            self.caption.setText(label or self._path.name)
        elif suffix == ".pdf":
            pages = order.get("_source_pages") or [1]
            try:
                self._pages = sorted({max(1, int(x)) for x in pages}) or [1]
            except Exception:
                self._pages = [1]
            self._render_pdf_page()
        else:
            self.caption.setText(f"{label or self._path.name}\n이 형식은 프로그램 내 미리보기를 제공하지 않습니다.")
            self.image_label.setText("[원본 파일 열기]로 확인하세요.")
            self.image_label.setPixmap(QPixmap())
        self._refresh_nav()

    def _render_pdf_page(self):
        if not self._path:
            return
        try:
            import pymupdf
            page_no = self._pages[self._page_index] if self._pages else 1
            with pymupdf.open(str(self._path)) as doc:
                page_no = min(max(1, page_no), doc.page_count)
                pix = doc[page_no - 1].get_pixmap(matrix=pymupdf.Matrix(3.2, 3.2), alpha=False)
                img = QImage.fromData(pix.tobytes("png"), "PNG")
                self._set_pixmap(QPixmap.fromImage(img))
            self.caption.setText(f"{self._path.name} · {page_no}페이지")
        except Exception as exc:
            self._original_pixmap = None
            self.image_label.setPixmap(QPixmap())
            self.image_label.setText(f"PDF 미리보기를 열지 못했습니다.\n{type(exc).__name__}: {exc}")

    def _set_pixmap(self, pixmap: QPixmap):
        self._original_pixmap = pixmap
        self._apply_view_mode()

    def _apply_view_mode(self):
        if not self._original_pixmap or self._original_pixmap.isNull():
            return
        if self._view_mode == "actual":
            shown = self._original_pixmap
        else:
            target_w = max(260, self.scroll.viewport().width() - 20)
            if self._original_pixmap.width() > target_w:
                shown = self._original_pixmap.scaledToWidth(target_w, Qt.TransformationMode.SmoothTransformation)
            else:
                shown = self._original_pixmap
        self.image_label.setPixmap(shown)
        self.image_label.setText("")
        self.image_label.resize(max(shown.width(), 260), max(shown.height(), 420))

    def fit_width(self):
        self._view_mode = "fit_width"
        self._apply_view_mode()

    def actual_size(self):
        self._view_mode = "actual"
        self._apply_view_mode()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._view_mode == "fit_width":
            self._apply_view_mode()

    def previous_page(self):
        if self._page_index > 0:
            self._page_index -= 1
            self._render_pdf_page()
            self._refresh_nav()

    def next_page(self):
        if self._page_index + 1 < len(self._pages):
            self._page_index += 1
            self._render_pdf_page()
            self._refresh_nav()

    def open_external(self):
        if self._path and self._path.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._path)))

    def open_large(self):
        if not self._original_pixmap or self._original_pixmap.isNull():
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("원본 크게 보기")
        dlg.resize(1200, 900)
        lay = QVBoxLayout(dlg)
        scroll = QScrollArea()
        scroll.setWidgetResizable(False)
        lab = QLabel()
        lab.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        lab.setPixmap(self._original_pixmap)
        lab.resize(self._original_pixmap.size())
        scroll.setWidget(lab)
        lay.addWidget(scroll)
        dlg.exec()


class FileDropZone(QFrame):
    """Streamlit 파일 업로더와 비슷한 클릭 + 드래그앤드롭 입력 영역."""

    files_dropped = Signal(object)
    clicked = Signal()
    ALLOWED = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".pdf", ".xlsx", ".xls"}

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("dropZone")
        self.setAcceptDrops(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(72)
        self.setMaximumHeight(82)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(2)
        title = QLabel("발주서를 끌어다 놓거나 선택하세요")
        title.setObjectName("dropTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub = QLabel("PNG · JPG · PDF · Excel")
        sub.setObjectName("dropSub")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title)
        lay.addWidget(sub)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)

    def dragEnterEvent(self, event):
        urls = event.mimeData().urls() if event.mimeData().hasUrls() else []
        if any(Path(u.toLocalFile()).suffix.lower() in self.ALLOWED for u in urls):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        paths = []
        for url in event.mimeData().urls():
            path = Path(url.toLocalFile())
            if path.is_file() and path.suffix.lower() in self.ALLOWED:
                paths.append(str(path))
        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()


class ImportWorker(QThread):
    progress = Signal(str, int, int)
    completed = Signal(object, object, object)

    def __init__(self, paths: list[str], client_hint: str | None, product_mode: str = "blind"):
        super().__init__()
        self.paths = paths
        self.client_hint = client_hint
        self.product_mode = product_mode

    def run(self):
        orders = []
        errors = []
        succeeded = []
        total = len(self.paths)
        for i, raw in enumerate(self.paths, 1):
            path = Path(raw)
            self.progress.emit(path.name, i, total)
            try:
                orders.extend(ingest_file(path, self.client_hint, self.product_mode))
                succeeded.append(str(path))
            except Exception as exc:
                errors.append(f"{path.name}: {type(exc).__name__} - {exc}")
        self.completed.emit(orders, errors, succeeded)


class StartPage(QWidget):
    selected = Signal(str)

    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.addStretch(1)
        wrap = QFrame()
        wrap.setObjectName("startCard")
        wrap.setMaximumWidth(900)
        box = QVBoxLayout(wrap)
        box.setContentsMargins(50, 45, 50, 45)
        title = QLabel("FitOrder")
        title.setObjectName("startTitle")
        subtitle = QLabel("처리할 제품군을 선택해 주세요")
        subtitle.setObjectName("startSubtitle")
        box.addWidget(title)
        box.addWidget(subtitle)
        box.addSpacing(28)
        buttons = QHBoxLayout()
        blind = QPushButton("블라인드")
        blind.setObjectName("modeButton")
        blind.setMinimumHeight(150)
        roll = QPushButton("롤/콤비")
        roll.setObjectName("modeButton")
        roll.setMinimumHeight(150)
        holding = QPushButton("홀딩도어")
        holding.setObjectName("modeButton")
        holding.setMinimumHeight(150)
        blind.clicked.connect(lambda: self.selected.emit("blind"))
        roll.clicked.connect(lambda: self.selected.emit("roll_combo"))
        holding.clicked.connect(lambda: self.selected.emit("holding"))
        buttons.addWidget(blind)
        buttons.addWidget(roll)
        buttons.addWidget(holding)
        box.addLayout(buttons)
        box.addSpacing(16)
        note = QLabel("블라인드 / 롤·콤비 / 홀딩도어를 제품군별로 분리해 처리합니다.\n"
                      "롤/콤비: 안산)보노 Excel + 나머지 업체 카카오톡 이미지/PDF.\n"
                      "홀딩도어: 기존 H 품목·계산·장부·작업지시서·경영박사 엔진을 전용 모드에서 사용합니다.")
        note.setObjectName("startNote")
        note.setWordWrap(True)
        box.addWidget(note)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(wrap)
        row.addStretch(1)
        outer.addLayout(row)
        outer.addStretch(1)


class OrderWorkspace(QWidget):
    back_requested = Signal()

    def __init__(self, product_mode: str = "blind"):
        super().__init__()
        self.product_mode = product_mode
        self.orders: list[dict] = []
        self.preview_rows: list[dict] = []
        self.output_rows: list[dict] = []
        self.statuses: list[tuple[str, str, str]] = []
        self.undo_stack: list[tuple[list[dict], list[tuple], list[tuple]]] = []
        self.manual_merges: list[tuple] = []
        self.manual_unmerges: list[tuple] = []
        self._loading = False
        self.worker: ImportWorker | None = None
        self._worker_origin = ""
        self._active_import_paths: list[str] = []
        self.pending_files: list[str] = []
        self.clip_seen: set[str] = set()
        self.clip_last: str | None = None
        self._active_clip_hash: str | None = None
        self.last_ledger_path: Path | None = self._load_last_ledger_path()

        self.M = get_master_for_mode(self.product_mode)
        if self.product_mode == "roll_combo":
            self.mode_clients = list(ROLL_CLIENTS)
            self.mode_client_label = roll_client_label
        elif self.product_mode == "holding":
            self.mode_clients = list(HOLDING_CLIENTS)
            self.mode_client_label = holding_client_label
        else:
            # 미더스와 홀딩 전용 거래처는 블라인드 목록에 섞이지 않는다.
            self.mode_clients = list(BLIND_CLIENTS)
            self.mode_client_label = client_label
        OUT_DIR.mkdir(parents=True, exist_ok=True)

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 8, 14, 10)
        root.setSpacing(6)

        nav = QHBoxLayout()
        self.back_btn = QPushButton("← 시작화면")
        self.back_btn.setObjectName("navButton")
        self.back_btn.clicked.connect(self.back_requested)
        nav.addWidget(self.back_btn)
        nav_title = QLabel("① 사전점검  →  ② 장부  →  ③ 작업지시서·경영박사")
        nav_title.setObjectName("navCaption")
        nav.addWidget(nav_title)
        nav.addStretch(1)
        version = QLabel(f"코드 버전 {APP_VERSION}")
        version.setObjectName("versionLabel")
        nav.addWidget(version)
        root.addLayout(nav)

        # 예전 버전처럼 왼쪽 입력은 항상 유지하고 오른쪽만 사전점검으로 바뀐다.
        body = QHBoxLayout()
        body.setSpacing(12)

        left = QFrame()
        left.setObjectName("leftPanel")
        left_lay = QVBoxLayout(left)
        left_lay.setContentsMargins(14, 10, 14, 10)
        left_lay.setSpacing(5)

        self.detail_panel = QFrame()
        self.detail_panel.setObjectName("precheckDetailPanel")
        detail_lay = QVBoxLayout(self.detail_panel)
        detail_lay.setContentsMargins(10, 9, 10, 9)
        detail_lay.setSpacing(6)
        detail_title = QLabel("상태 상세")
        detail_title.setObjectName("sectionTitle")
        detail_lay.addWidget(detail_title)
        self.issue_head = QLabel("상태 셀을 클릭하면 오류/확인 이유와 원본을 함께 볼 수 있습니다.")
        self.issue_head.setObjectName("issueHead")
        self.issue_head.setWordWrap(True)
        detail_lay.addWidget(self.issue_head)
        reason_frame = QFrame()
        reason_frame.setObjectName("issueReasonBox")
        reason_lay = QVBoxLayout(reason_frame)
        reason_lay.setContentsMargins(9, 7, 9, 7)
        self.issue_reason = QLabel("검증 항목을 선택해 주세요.")
        self.issue_reason.setWordWrap(True)
        self.issue_reason.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        reason_lay.addWidget(self.issue_reason)
        reason_scroll = QScrollArea()
        reason_scroll.setWidgetResizable(True)
        reason_scroll.setFrameShape(QFrame.Shape.NoFrame)
        reason_scroll.setMinimumHeight(74)
        reason_scroll.setMaximumHeight(120)
        reason_scroll.setWidget(reason_frame)
        detail_lay.addWidget(reason_scroll)
        self.source_wrap = QFrame()
        self.source_wrap.setObjectName("sourcePanel")
        source_lay = QVBoxLayout(self.source_wrap)
        source_lay.setContentsMargins(0, 0, 0, 0)
        self.source = SourceViewer()
        self.source.setMinimumHeight(470)
        self.source.close_requested.connect(self.hide_source)
        source_lay.addWidget(self.source)
        detail_lay.addWidget(self.source_wrap)
        self.detail_panel.hide()
        left_lay.addWidget(self.detail_panel)

        brand = QLabel("FitOrder")
        brand.setObjectName("brandTitle")
        left_lay.addWidget(brand)
        brand_sub = QLabel(
            "롤/콤비 발주서 입력" if self.product_mode == "roll_combo"
            else ("홀딩도어 발주서 입력" if self.product_mode == "holding" else "블라인드 발주서 입력")
        )
        brand_sub.setObjectName("sectionCaption")
        left_lay.addWidget(brand_sub)
        left_lay.addWidget(self._divider())

        ship_label = QLabel("출고일")
        ship_label.setObjectName("fieldLabel")
        left_lay.addWidget(ship_label)
        self.ship_edit = NoWheelDateEdit()
        self.ship_edit.setDisplayFormat("yyyy-MM-dd")
        d = default_ship()
        self.ship_edit.setDate(QDate(d.year, d.month, d.day))
        self.ship_edit.dateChanged.connect(self._on_ship_date_changed)
        self.ship_edit.setToolTip("달력 버튼을 눌러 출고일을 선택합니다. 마우스휠/직접입력으로는 변경되지 않습니다.")
        left_lay.addWidget(self.ship_edit)
        self.ship_hint = QLabel("")
        self.ship_hint.setObjectName("mutedText")
        left_lay.addWidget(self.ship_hint)
        self._update_ship_hint()

        client_label_widget = QLabel("거래처")
        client_label_widget.setObjectName("fieldLabel")
        left_lay.addWidget(client_label_widget)
        self.client_combo = NoWheelComboBox()
        self.client_combo.addItem("자동 판별", None)
        for c in self.mode_clients:
            self.client_combo.addItem(self.mode_client_label(c), c)
        self.client_combo.setToolTip("목록을 열어 거래처를 선택합니다. 닫힌 상태에서는 마우스휠로 바뀌지 않습니다.")
        left_lay.addWidget(self.client_combo)

        left_lay.addWidget(self._divider())
        input_title = QLabel("발주서 입력")
        input_title.setObjectName("sectionTitle")
        left_lay.addWidget(input_title)

        self.paste_btn = QPushButton("붙여넣기 (Ctrl+V)")
        self.paste_btn.clicked.connect(self.paste_clipboard)
        left_lay.addWidget(self.paste_btn)
        self.watch_check = QCheckBox("자동 캡처 감시")
        self.watch_check.toggled.connect(self._toggle_watch)
        left_lay.addWidget(self.watch_check)
        self.watch_hint = QLabel("")
        self.watch_hint.setWordWrap(True)
        self.watch_hint.setObjectName("mutedText")
        left_lay.addWidget(self.watch_hint)
        self.watch_timer = QTimer(self)
        self.watch_timer.setInterval(2000)
        self.watch_timer.timeout.connect(self._watch_clipboard)

        self.drop_zone = FileDropZone()
        self.drop_zone.clicked.connect(self.choose_files)
        self.drop_zone.files_dropped.connect(self.queue_files)
        left_lay.addWidget(self.drop_zone)

        pending_head = QHBoxLayout()
        pending_head.setSpacing(6)
        self.pending_count = QLabel("선택된 발주서 0개")
        self.pending_count.setObjectName("pendingCount")
        pending_head.addWidget(self.pending_count)
        pending_head.addStretch(1)
        self.remove_pending_btn = QPushButton("선택 삭제")
        self.remove_pending_btn.setObjectName("tinyButton")
        self.remove_pending_btn.clicked.connect(self.remove_selected_pending_files)
        self.clear_pending_btn = QPushButton("전체 삭제")
        self.clear_pending_btn.setObjectName("tinyButton")
        self.clear_pending_btn.clicked.connect(self.clear_pending_files)
        pending_head.addWidget(self.remove_pending_btn)
        pending_head.addWidget(self.clear_pending_btn)
        left_lay.addLayout(pending_head)

        self.pending_list = QListWidget()
        self.pending_list.setObjectName("pendingFileList")
        self.pending_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.pending_list.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        self.pending_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.pending_list.setMinimumHeight(54)
        self.pending_list.setMaximumHeight(104)
        left_lay.addWidget(self.pending_list)

        file_buttons = QHBoxLayout()
        self.add_btn = QPushButton("파일 선택")
        self.add_btn.clicked.connect(self.choose_files)
        self.analyze_btn = QPushButton("분석 시작")
        self.analyze_btn.setObjectName("primaryButton")
        self.analyze_btn.clicked.connect(self.analyze_pending)
        file_buttons.addWidget(self.add_btn, 1)
        file_buttons.addWidget(self.analyze_btn, 1)
        left_lay.addLayout(file_buttons)

        ledger_tools = QHBoxLayout()
        ledger_tools.setSpacing(6)
        self.ledger_convert_btn = QPushButton("장부 → 사전점검")
        self.ledger_convert_btn.setToolTip(
            "직접 작성한 기존 장부.xls/xlsx를 사전점검으로 불러옵니다.\n"
            "확인·수정한 뒤 [파일 생성]을 누르면 장부·작업지시서·경영박사 EDI를 함께 만듭니다.")
        self.ledger_convert_btn.clicked.connect(self.convert_selected_ledger)
        if self.product_mode == "roll_combo":
            self.ledger_convert_btn.setEnabled(False)
            self.ledger_convert_btn.setToolTip("롤/콤비는 발주서 분석 → 3문서 출력 흐름을 사용합니다.")
        self.previous_ledger_btn = QPushButton("이전 장부 보기")
        self.previous_ledger_btn.clicked.connect(self.open_previous_ledger)
        ledger_tools.addWidget(self.ledger_convert_btn, 1)
        ledger_tools.addWidget(self.previous_ledger_btn, 1)
        left_lay.addLayout(ledger_tools)

        self.progress_label = QLabel("")
        self.progress_label.setWordWrap(True)
        self.progress_label.setObjectName("progressLabel")
        left_lay.addWidget(self.progress_label)
        self.progress = QProgressBar()
        self.progress.hide()
        left_lay.addWidget(self.progress)
        left_lay.addStretch(1)

        order_tools = QHBoxLayout()
        self.undo_last_btn = QPushButton("방금 넣은 것 취소")
        self.undo_last_btn.clicked.connect(self.remove_last_order)
        self.clear_btn = QPushButton("전체 비우기")
        self.clear_btn.clicked.connect(self.clear_all)
        order_tools.addWidget(self.undo_last_btn)
        order_tools.addWidget(self.clear_btn)
        left_lay.addLayout(order_tools)

        self.left_scroll = QScrollArea()
        self.left_scroll.setObjectName("leftScroll")
        self.left_scroll.setWidgetResizable(True)
        self.left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.left_scroll.setWidget(left)
        body.addWidget(self.left_scroll, 3)

        self.right_stack = QStackedWidget()
        right = QFrame()
        right.setObjectName("homeRightPanel")
        right_lay = QVBoxLayout(right)
        right_lay.setContentsMargins(24, 20, 24, 20)
        right_lay.setSpacing(8)
        home_title = QLabel("사전점검")
        home_title.setObjectName("homeTitle")
        right_lay.addWidget(home_title)
        self.main_summary = QLabel("주문 0건 · 장부 0행")
        self.main_summary.setObjectName("homeSummary")
        right_lay.addWidget(self.main_summary)
        self.main_status_title = QLabel("발주서를 추가해 주세요.")
        self.main_status_title.setObjectName("mutedText")
        right_lay.addWidget(self.main_status_title)
        self.main_status_text = QLabel("")
        self.main_status_text.hide()
        self.preview_open_btn = QPushButton("사전점검 보기")
        self.preview_open_btn.setObjectName("primaryButton")
        self.preview_open_btn.clicked.connect(self.show_precheck)
        self.preview_open_btn.setEnabled(False)
        right_lay.addWidget(self.preview_open_btn)
        right_lay.addStretch(1)
        self.home_right = right
        self.right_stack.addWidget(self.home_right)

        self._build_precheck_window()
        self.right_stack.setCurrentWidget(self.home_right)
        body.addWidget(self.right_stack, 7)
        root.addLayout(body, 1)

        self._set_column_widths()
        self._install_actions()
        self._update_pending_label()
        self._update_order_buttons()
        self.rebuild_table(preserve=False)

    def _build_precheck_window(self):
        self.precheck = QFrame()
        self.precheck.setObjectName("ledgerPanel")
        outer = QVBoxLayout(self.precheck)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(7)

        top = QHBoxLayout()
        title = QLabel("사전점검")
        title.setObjectName("precheckTitle")
        top.addWidget(title)
        top.addStretch(1)
        self.summary = QLabel("주문 0건")
        self.summary.setObjectName("summary")
        top.addWidget(self.summary)
        hide_btn = QPushButton("접기")
        hide_btn.setObjectName("smallButton")
        hide_btn.clicked.connect(lambda: self.right_stack.setCurrentWidget(self.home_right))
        top.addWidget(hide_btn)
        outer.addLayout(top)

        row_tools = QHBoxLayout()
        self.row_add_btn = QPushButton("+ 선택 행 아래 추가")
        self.row_add_btn.setObjectName("smallButton")
        self.row_add_btn.clicked.connect(self.add_row_below)
        row_tools.addWidget(self.row_add_btn)
        self.merge_btn = QPushButton("셀 병합")
        self.merge_btn.setObjectName("smallButton")
        self.merge_btn.setToolTip("선택한 사각형 셀을 병합합니다. 단축키 Alt+1")
        self.merge_btn.clicked.connect(self.merge_selected)
        row_tools.addWidget(self.merge_btn)
        self.unmerge_btn = QPushButton("병합 해제")
        self.unmerge_btn.setObjectName("smallButton")
        self.unmerge_btn.setToolTip("선택한 병합 셀을 해제합니다. 단축키 Alt+2")
        self.unmerge_btn.clicked.connect(self.unmerge_selected)
        row_tools.addWidget(self.unmerge_btn)
        self.row_up_btn = QPushButton("선택 행 ↑")
        self.row_up_btn.setObjectName("smallButton")
        self.row_up_btn.clicked.connect(lambda: self.move_selected_rows(-1))
        self.row_down_btn = QPushButton("선택 행 ↓")
        self.row_down_btn.setObjectName("smallButton")
        self.row_down_btn.clicked.connect(lambda: self.move_selected_rows(1))
        self.row_delete_btn = QPushButton("선택 행 삭제")
        self.row_delete_btn.setObjectName("dangerSmallButton")
        self.row_delete_btn.clicked.connect(self.delete_selected_rows)
        row_tools.addWidget(self.row_up_btn)
        row_tools.addWidget(self.row_down_btn)
        row_tools.addWidget(self.row_delete_btn)
        row_tools.addStretch(1)
        outer.addLayout(row_tools)

        self.table = QTableWidget(0, len(HEADERS))
        self.table.setHorizontalHeaderLabels(HEADERS)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.EditKeyPressed)
        self.table.setAlternatingRowColors(False)
        self.table.setWordWrap(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(31)
        self.table.horizontalHeader().setMinimumHeight(36)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setItemDelegateForColumn(COL["색상"], LedgerColorDelegate(self.table))
        self.table.itemChanged.connect(self.on_item_changed)
        self.table.cellClicked.connect(self.on_cell_clicked)
        outer.addWidget(self.table, 1)

        out_row = QHBoxLayout()
        out_row.addStretch(1)
        self.output_btn = QPushButton("파일 생성")
        self.output_btn.setObjectName("primaryButton")
        self.output_btn.clicked.connect(self.export_outputs)
        out_row.addWidget(self.output_btn)
        outer.addLayout(out_row)
        self.right_stack.addWidget(self.precheck)

    def show_precheck(self):
        if not self.orders:
            QMessageBox.information(self, "사전점검", "사전점검할 주문이 없습니다.")
            return
        self.rebuild_table(preserve=False)
        if self.preview_rows:
            self.right_stack.setCurrentWidget(self.precheck)

    def hide_source(self):
        """원본/상태 상세만 접고 입력창과 장부는 그대로 유지한다."""
        if hasattr(self, "detail_panel"):
            self.detail_panel.hide()

    @staticmethod
    def _divider() -> QFrame:
        line = QFrame()
        line.setObjectName("divider")
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFixedHeight(1)
        return line

    def _on_ship_date_changed(self, _):
        # 좌측 출고일 선택값을 이 배치의 단일 기준일로 사용한다.
        # 원본/이전 편집값에 다른 날짜가 있어도 선택한 날짜가 장부/작업지시서/EDI에 우선한다.
        d = self._ship_date().isoformat()
        for order in self.orders:
            order["_ship_date"] = d
            order.pop("_ship_date_manual", None)
        self._update_ship_hint()
        self.rebuild_table()

    def _update_ship_hint(self):
        d = self._ship_date()
        weekdays = ["월", "화", "수", "목", "금", "토", "일"]
        self.ship_hint.setText(f"{weekdays[d.weekday()]}요일 출고")

    def _update_order_buttons(self):
        has_orders = bool(self.orders)
        self.undo_last_btn.setEnabled(has_orders)
        self.clear_btn.setEnabled(has_orders)
        busy = bool(self.worker and self.worker.isRunning())
        if hasattr(self, "previous_ledger_btn"):
            self.previous_ledger_btn.setEnabled(not busy)
        self.output_btn.setEnabled(has_orders and not busy)
        if hasattr(self, "preview_open_btn"):
            self.preview_open_btn.setEnabled(has_orders and not busy)
        if hasattr(self, "row_up_btn"):
            self.row_add_btn.setEnabled(has_orders and not busy)
            self.merge_btn.setEnabled(has_orders and not busy)
            self.unmerge_btn.setEnabled(has_orders and not busy)
            self.row_up_btn.setEnabled(has_orders and not busy)
            self.row_down_btn.setEnabled(has_orders and not busy)
            self.row_delete_btn.setEnabled(has_orders and not busy)

    def _set_column_widths(self):
        """초기에는 전체 열이 보이게 하고, 헤더 경계선 드래그로 직접 폭을 조절한다."""
        widths = {
            "No": 38, "상태": 58, "상호": 78, "색상": 142, "가로": 54, "X": 26,
            "세로": 54, "수량": 46, "방향": 46, "길이": 56, "특이": 58,
            "기재사항1": 150, "기재사항2": 150, "출고일": 72,
        }
        header = self.table.horizontalHeader()
        for name in HEADERS:
            header.setSectionResizeMode(COL[name], QHeaderView.ResizeMode.Interactive)
            self.table.setColumnWidth(COL[name], widths[name])
        header.setStretchLastSection(False)
        self.table.setMinimumWidth(760)

    def _install_actions(self):
        def action(text, shortcut, slot):
            a = QAction(text, self)
            a.setShortcut(QKeySequence(shortcut))
            a.setShortcutContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
            a.triggered.connect(slot)
            self.addAction(a)
            return a
        action("선택 셀 병합", "Alt+1", self.merge_selected)
        action("병합 해제", "Alt+2", self.unmerge_selected)
        action("실행 취소", "Ctrl+Z", self.undo)
        action("셀 내용 삭제", "Delete", self.delete_selected_cells)
        action("셀 수정", "F2", self.edit_current_cell)
        action("셀 복사", "Ctrl+C", self.copy_selected_cells)
        action("원본 닫기", "Esc", self.hide_source)
        action("발주서 붙여넣기", "Ctrl+V", self._paste_shortcut)

    def _ship_date(self) -> date:
        q = self.ship_edit.date()
        return date(q.year(), q.month(), q.day())

    @staticmethod
    def _last_ledger_marker() -> Path:
        return ROOT / ".fitorder_last_ledger.txt"

    def _load_last_ledger_path(self) -> Path | None:
        try:
            marker = self._last_ledger_marker()
            if marker.exists():
                p = Path(marker.read_text(encoding="utf-8").strip())
                if p.is_file():
                    return p
        except Exception:
            pass
        return None

    def _remember_last_ledger(self, path: str | Path) -> None:
        try:
            p = Path(path).resolve()
            self.last_ledger_path = p
            self._last_ledger_marker().write_text(str(p), encoding="utf-8")
        except Exception:
            self.last_ledger_path = Path(path)

    def open_previous_ledger(self):
        p = self.last_ledger_path
        if p and Path(p).is_file():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(p))))
            return
        start = OUT_DIR if OUT_DIR.exists() else ROOT
        chosen, _ = QFileDialog.getOpenFileName(
            self, "이전 장부 선택", str(start), "장부 Excel (*.xlsx *.xls);;모든 파일 (*.*)")
        if not chosen:
            return
        self._remember_last_ledger(chosen)
        QDesktopServices.openUrl(QUrl.fromLocalFile(chosen))

    def choose_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "발주서 선택",
            str(Path.home()),
            "발주서 (*.png *.jpg *.jpeg *.webp *.bmp *.pdf *.xlsx *.xls);;모든 파일 (*.*)",
        )
        if paths:
            self.queue_files(paths)

    def queue_files(self, paths: object):
        allowed = FileDropZone.ALLOWED
        added = 0
        for raw in list(paths or []):
            path = Path(str(raw))
            if not path.is_file() or path.suffix.lower() not in allowed:
                continue
            key = os.path.normcase(str(path.resolve()))
            existing = {os.path.normcase(str(Path(p).resolve())) for p in self.pending_files}
            if key in existing:
                continue
            self.pending_files.append(str(path))
            added += 1
        self._update_pending_label()
        if added:
            self.progress_label.setText(f"발주서 {added}개 선택")

    def _update_pending_label(self):
        # 파일명을 한 줄씩 분리해서 보여 주고 전체 경로는 툴팁으로 제공한다.
        self.pending_list.clear()
        count = len(self.pending_files)
        self.pending_count.setText(f"선택된 발주서 {count}개")
        # 첫 화면에서 하단 버튼까지 한 번에 보이도록 파일 목록 높이를 자동 조절한다.
        # 파일이 많을 때만 목록 내부 스크롤을 사용한다.
        list_h = 54 if count <= 1 else min(104, 38 + count * 22)
        self.pending_list.setFixedHeight(list_h)
        if not self.pending_files:
            item = QListWidgetItem("선택된 발주서 없음")
            item.setFlags(Qt.ItemFlag.NoItemFlags)
            self.pending_list.addItem(item)
            self.analyze_btn.setEnabled(False)
            self.ledger_convert_btn.setEnabled(False)
            self.remove_pending_btn.setEnabled(False)
            self.clear_pending_btn.setEnabled(False)
            return
        for raw in self.pending_files:
            path = Path(raw)
            item = QListWidgetItem(path.name)
            item.setData(Qt.ItemDataRole.UserRole, str(path))
            item.setToolTip(str(path))
            self.pending_list.addItem(item)
        self.remove_pending_btn.setEnabled(True)
        self.clear_pending_btn.setEnabled(True)
        busy = bool(self.worker and self.worker.isRunning())
        self.analyze_btn.setEnabled(not busy)
        self.ledger_convert_btn.setEnabled(
            self.product_mode != "roll_combo" and not busy and len(self.pending_files) == 1
            and Path(self.pending_files[0]).suffix.lower() in {".xls", ".xlsx"}
        )

    def remove_selected_pending_files(self):
        selected = self.pending_list.selectedItems()
        if not selected:
            return
        remove = {str(it.data(Qt.ItemDataRole.UserRole)) for it in selected
                  if it.data(Qt.ItemDataRole.UserRole)}
        if not remove:
            return
        norm_remove = {os.path.normcase(str(Path(p).resolve())) for p in remove}
        self.pending_files = [
            p for p in self.pending_files
            if os.path.normcase(str(Path(p).resolve())) not in norm_remove
        ]
        self._update_pending_label()
        self.progress_label.setText(f"선택 파일 {len(remove)}개를 목록에서 삭제했습니다.")

    def clear_pending_files(self):
        if not self.pending_files:
            return
        self.pending_files.clear()
        self._update_pending_label()
        self.progress_label.setText("선택 파일 목록을 비웠습니다.")

    def analyze_pending(self):
        if not self.pending_files:
            QMessageBox.information(self, "분석 시작", "먼저 발주서를 선택해 주세요.")
            return
        self._start_import(list(self.pending_files), "files")

    def _start_import(self, paths: list[str], origin: str):
        if not paths or (self.worker and self.worker.isRunning()):
            return
        hint = self.client_combo.currentData()
        self._worker_origin = origin
        self._active_import_paths = list(paths)
        self._set_busy(True)
        self.worker = ImportWorker(paths, hint, self.product_mode)
        self.worker.progress.connect(self._on_import_progress)
        self.worker.completed.connect(self._on_import_done)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.start()

    def paste_clipboard(self):
        if self.worker and self.worker.isRunning():
            return
        img, h = clip.grab()
        if not h or clip.too_small(img):
            QMessageBox.information(
                self, "붙여넣기",
                "클립보드에서 발주서 이미지를 찾지 못했습니다.\n"
                "Win+Shift+S로 다시 캡처한 뒤 Ctrl+V를 눌러 주세요.")
            return
        if h in self.clip_seen:
            self.progress_label.setText("이미 정상 처리한 캡처입니다.")
            return
        # 분석 성공 전에는 seen에 넣지 않는다. 실패한 동일 캡처를 Ctrl+V로 재시도할 수 있다.
        self.clip_last = h
        self._active_clip_hash = h
        path = clip.save(img, OUT_DIR, h[:8])
        self._start_import([str(path)], "clipboard")

    def _toggle_watch(self, checked: bool):
        if checked:
            _, h = clip.grab()
            if h:
                self.clip_seen.add(h)
                self.clip_last = h
            self.watch_timer.start()
            self.watch_hint.setText("감시 중 · 캡처하면 자동으로 주문에 추가됩니다.")
        else:
            self.watch_timer.stop()
            self.watch_hint.setText("")

    def _watch_clipboard(self):
        if not self.watch_check.isChecked() or (self.worker and self.worker.isRunning()):
            return
        img, h = clip.grab()
        if not h or h == self.clip_last or h in self.clip_seen or clip.too_small(img):
            return
        # 자동 감시도 성공한 뒤에만 seen 처리한다.
        self.clip_last = h
        self._active_clip_hash = h
        path = clip.save(img, OUT_DIR, h[:8])
        self._start_import([str(path)], "watch")

    def _paste_shortcut(self):
        # 표 셀 편집 중에는 일반 Ctrl+V 입력을 방해하지 않는다.
        if self.table.state() == QAbstractItemView.State.EditingState:
            return
        # 사전점검 표에 초점이 있고 클립보드에 글자가 있으면 엑셀처럼 셀에 붙여넣는다.
        if self._table_has_focus() and self._clipboard_has_cell_text():
            self.paste_cells_from_clipboard()
            return
        self.paste_clipboard()

    def _table_has_focus(self) -> bool:
        return (self.right_stack.currentWidget() is self.precheck
                and (self.table.hasFocus() or self.table.viewport().hasFocus()))

    @staticmethod
    def _clipboard_has_cell_text() -> bool:
        mime = QApplication.clipboard().mimeData()
        return bool(mime is not None and mime.hasText() and not mime.hasImage()
                    and not mime.hasUrls() and mime.text().strip())

    def copy_selected_cells(self):
        """선택한 셀 글자를 엑셀과 같은 탭/줄바꿈 형식으로 복사한다."""
        indexes = [i for i in self.table.selectedIndexes() if i.column() >= COL["상호"]]
        if not indexes:
            return
        selected = {(i.row(), i.column()) for i in indexes}
        top, bottom = min(r for r, _ in selected), max(r for r, _ in selected)
        left, right = min(c for _, c in selected), max(c for _, c in selected)
        lines = []
        for r in range(top, bottom + 1):
            cells = []
            for c in range(left, right + 1):
                item = self.table.item(r, c)
                text = item.text() if item and (r, c) in selected else ""
                cells.append(text.replace("\t", " ").replace("\n", " "))
            lines.append("\t".join(cells))
        QApplication.clipboard().setText("\n".join(lines))

    def paste_cells_from_clipboard(self):
        """클립보드의 탭/줄바꿈 글자를 선택한 셀부터 붙여넣는다. 수정할 수 없는 칸은 건너뛴다."""
        text = QApplication.clipboard().text()
        if not text or not self.preview_rows:
            return
        ranges = self.table.selectedRanges()
        if ranges:
            start_row = min(r.topRow() for r in ranges)
            start_col = min(r.leftColumn() for r in ranges)
        else:
            start_row, start_col = self.table.currentRow(), self.table.currentColumn()
        if start_row < 0 or start_col < COL["상호"]:
            return
        grid = [line.split("\t") for line in text.replace("\r\n", "\n").rstrip("\n").split("\n")]
        product_edits, special_edits = [], []
        for dr, values in enumerate(grid):
            r = start_row + dr
            if r >= len(self.preview_rows):
                break
            preview_row = self.preview_rows[r]
            for dc, value in enumerate(values):
                c = start_col + dc
                if c >= len(HEADERS):
                    break
                name = HEADERS[c]
                if name not in editable_columns_for_row(preview_row):
                    continue
                if preview_row.get("_특수") is None:
                    product_edits.append((preview_row.get("_row_key"), name, value))
                else:
                    special_edits.append((r, name, value))
        if not product_edits and not special_edits:
            return
        self._push_undo()
        for r, name, value in special_edits:
            apply_preview_cell_edit(self.orders, self.preview_rows[r], name, value)
        self._apply_edits_by_key(product_edits)
        self.rebuild_table(preserve=True)

    def convert_selected_ledger(self):
        """직접 작성한 기존 장부를 사전점검으로 불러온다.

        읽은 주문은 발주서 분석 결과와 같은 사전점검 주문이 되고, 확인·수정 후
        [파일 생성]에서 장부·작업지시서·경영박사 EDI를 함께 만든다.
        """
        if len(self.pending_files) != 1:
            QMessageBox.information(self, "장부 불러오기", "장부 파일 하나만 선택해 주세요.")
            return
        source = Path(self.pending_files[0])
        if source.suffix.lower() not in {".xls", ".xlsx"}:
            QMessageBox.warning(self, "장부 불러오기", "기존 장부.xls 또는 장부.xlsx 파일을 선택해 주세요.")
            return
        try:
            self._set_busy(True)
            _rows, orders, errors = read_ledger(source, self.M, source.name)
            if self.product_mode == "holding":
                # 홀딩 품목은 불러온 뒤 홀딩 품목장으로 다시 매칭하므로 블라인드 해석 실패 문구는 제외한다.
                errors = [e for e in errors if "해석할 수 없습니다" not in str(e)]
            orders = ledger_orders_for_precheck(orders, source, self.product_mode, self._ship_date())
            if self.product_mode == "holding":
                errors += [f"홀딩 품목을 찾지 못했습니다: {it.get('_holding_raw') or it.get('색상원문') or ''}"
                           for o in orders for it in o.get("items") or [] if it.get("_holding_unmatched")]
            item_count = sum(1 for o in orders for it in o.get("items", []) if not it.get("예외품목"))
            if not orders or not item_count:
                QMessageBox.warning(self, "장부 불러오기", "장부에서 불러올 주문을 찾지 못했습니다.")
                return
            if errors:
                shown = "\n".join(str(x) for x in errors[:15])
                if len(errors) > 15:
                    shown += f"\n… 외 {len(errors) - 15}건"
                ans = QMessageBox.question(
                    self, "장부에서 확인 필요",
                    f"확인이 필요한 부분이 {len(errors)}건 있습니다.\n\n{shown}\n\n"
                    "그래도 사전점검으로 불러올까요? 문제 행은 사전점검 상태 칸에서 확인할 수 있습니다.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if ans != QMessageBox.StandardButton.Yes:
                    return
            self._push_undo()
            self.orders.extend(orders)
            self.rebuild_table(preserve=False)
            self._update_order_buttons()
            self.progress_label.setText(
                f"장부 불러오기 완료 · 주문 {len(orders)}건 · 품목 {item_count}행 · 사전점검 확인 후 [파일 생성]")
            QTimer.singleShot(0, self.show_precheck)
        except Exception as exc:
            QMessageBox.critical(self, "장부 불러오기 실패", f"{type(exc).__name__}: {exc}")
        finally:
            # 불러오기를 시도한 파일은 성공/실패 모두 입력 목록에서 제거한다.
            target = os.path.normcase(str(source.resolve()))
            self.pending_files = [p for p in self.pending_files
                                  if os.path.normcase(str(Path(p).resolve())) != target]
            self._update_pending_label()
            self._set_busy(False)

    def remove_last_order(self):
        if not self.orders:
            return
        self._push_undo()
        self.orders.pop()
        self.rebuild_table(preserve=False)
        self._update_order_buttons()

    def _set_busy(self, busy: bool):
        self.add_btn.setEnabled(not busy)
        self.paste_btn.setEnabled(not busy)
        self.analyze_btn.setEnabled(not busy and bool(self.pending_files))
        self.ledger_convert_btn.setEnabled(
            not busy and len(self.pending_files) == 1
            and Path(self.pending_files[0]).suffix.lower() in {".xls", ".xlsx"})
        self.output_btn.setEnabled(not busy and bool(self.orders))
        if hasattr(self, "row_up_btn"):
            self.row_add_btn.setEnabled(not busy and bool(self.orders))
            self.row_up_btn.setEnabled(not busy and bool(self.orders))
            self.row_down_btn.setEnabled(not busy and bool(self.orders))
            self.row_delete_btn.setEnabled(not busy and bool(self.orders))
        if hasattr(self, "preview_open_btn"):
            self.preview_open_btn.setEnabled(not busy and bool(self.orders))
        self.clear_btn.setEnabled(not busy and bool(self.orders))
        self.undo_last_btn.setEnabled(not busy and bool(self.orders))
        self.back_btn.setEnabled(not busy)
        self.drop_zone.setEnabled(not busy)
        self.progress.setVisible(busy)
        if not busy and not self.progress_label.text().startswith(("감시", "발주서")):
            pass

    def _on_import_progress(self, name: str, current: int, total: int):
        self.progress.setRange(0, total)
        self.progress.setValue(current - 1)
        self.progress_label.setText(f"분석 중 · {name} ({current}/{total})")

    def _on_import_done(self, new_orders: object, errors: object, succeeded: object):
        new_orders = list(new_orders or [])
        errors = list(errors or [])

        if new_orders and self._worker_origin in {"clipboard", "watch"} and self._active_clip_hash:
            self.clip_seen.add(self._active_clip_hash)

        # 파일 선택 방식은 성공/실패와 관계없이 이번에 분석을 시도한 파일을 목록에서 제거한다.
        # 실패 파일이 계속 남아 재실행 때 또 분석되는 혼란을 막는다.
        if self._worker_origin == "files":
            attempted = {os.path.normcase(str(Path(p).resolve())) for p in self._active_import_paths}
            self.pending_files = [
                p for p in self.pending_files
                if os.path.normcase(str(Path(p).resolve())) not in attempted
            ]
            self._update_pending_label()

        if new_orders:
            selected_ship = self._ship_date().isoformat()
            for order in new_orders:
                order["_ship_date"] = selected_ship
            self._push_undo()
            self.orders.extend(new_orders)
            self.rebuild_table(preserve=False)

        self._set_busy(False)
        self.progress.setValue(self.progress.maximum())
        self._update_order_buttons()

        if errors:
            title = "발주서 분석 실패" if not new_orders else "일부 발주서 분석 실패"
            detail = "\n\n".join(errors)
            if not new_orders:
                detail = (
                    "분석에 실패하여 실제 장부 데이터가 0행입니다. "
                    "실패한 파일은 입력 목록에서 제거했습니다.\n\n" + detail
                )
            QMessageBox.warning(self, title, detail)
            self.progress_label.setText(
                f"분석 결과 · 주문 {len(new_orders)}건 · 실패 {len(errors)}건")
            if not new_orders:
                self.main_status_title.setText("분석 실패 · 장부 생성 0행")
                self.main_status_text.setText("발주서에서 유효한 주문을 읽지 못했습니다. 파일을 다시 선택해 재시도할 수 있습니다.")
        elif new_orders:
            product_rows = sum(1 for r in self.preview_rows if r.get("_특수") is None)
            self.progress_label.setText(f"{len(new_orders)}건 추가 완료 · 장부 {product_rows}행 생성")

        # 성공한 주문이 하나라도 있으면 분석 직후 사전점검 화면으로 전환한다.
        # 자동 캡처 감시는 반복 팝업을 피하고 메인 화면의 [사전점검 다시 열기]로 확인한다.
        if new_orders and self._worker_origin != "watch":
            QTimer.singleShot(0, self.show_precheck)

    def _on_worker_finished(self):
        self.worker = None
        self._active_import_paths = []
        self._active_clip_hash = None
        self._set_busy(False)
        self._update_pending_label()
        self._update_order_buttons()

    def _push_undo(self):
        self.undo_stack.append((
            copy.deepcopy(self.orders),
            list(self.manual_merges),
            list(self.manual_unmerges),
        ))
        self.undo_stack = self.undo_stack[-20:]

    def undo(self):
        if not self.undo_stack:
            return
        orders, merges, unmerges = self.undo_stack.pop()
        self.orders = orders
        self.manual_merges = merges
        self.manual_unmerges = unmerges
        self.rebuild_table(preserve=False)

    def clear_all(self):
        if not self.orders:
            return
        ans = QMessageBox.question(
            self, "전체 비우기", "현재 사전점검 주문을 모두 비울까요?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if ans != QMessageBox.StandardButton.Yes:
            return
        self._push_undo()
        self.orders = []
        self.manual_merges.clear()
        self.manual_unmerges.clear()
        self.rebuild_table(preserve=False)
        self.source.image_label.setText("원본 없음")
        self.source.image_label.setPixmap(QPixmap())
        self.detail_panel.hide()
        self.issue_head.setText("상태 셀을 클릭하면 오류/확인 이유와 원본을 함께 볼 수 있습니다.")
        self.issue_reason.setText("검증 항목을 선택해 주세요.")
        self.right_stack.setCurrentWidget(self.home_right)
        self._update_order_buttons()

    def _status_brush(self, level: str) -> QColor:
        return {
            "red": QColor("#ffd9d9"),
            "review": QColor("#eadbff"),
            "yellow": QColor("#fff1b8"),
            "info": QColor("#e2ebf6"),
            "ok": QColor("#f5f7f9"),
        }.get(level, QColor("#f5f7f9"))

    def rebuild_table(self, preserve: bool = True):
        if self._loading:
            return
        scroll = self.table.verticalScrollBar().value() if preserve else 0
        current_row = self.table.currentRow() if preserve else -1
        self._loading = True
        self.table.blockSignals(True)
        try:
            self.table.clearSpans()
            self.table.clearContents()
            if not self.orders:
                self.preview_rows = []
                self.output_rows = []
                self.statuses = []
                self.table.setRowCount(0)
                self.summary.setText("주문 0건 · 장부 0행")
                if hasattr(self, "main_summary"):
                    self.main_summary.setText("현재 주문 0건 · 장부 0행")
                    self.preview_open_btn.setEnabled(False)
                self._update_order_buttons()
                return

            self.preview_rows, self.output_rows = build_preview_rows(
                self.orders, self._ship_date(), self.M, sort_di=False)
            self.statuses = preview_statuses(self.preview_rows, self.orders, self.M)
            self.table.setRowCount(len(self.preview_rows))

            for r, row in enumerate(self.preview_rows):
                values = row_display_values(row, r + 1)
                level, status, tooltip = self.statuses[r]
                no_item = QTableWidgetItem(str(r + 1))
                no_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                no_item.setFlags(no_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(r, COL["No"], no_item)

                status_item = QTableWidgetItem(status)
                status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                status_item.setToolTip(tooltip + "\n클릭: 원본 발주서 보기")
                status_item.setBackground(self._status_brush(level))
                status_item.setFlags(status_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(r, COL["상태"], status_item)

                editable = editable_columns_for_row(row)
                for name in LEDGER_COLUMNS:
                    text = values.get(name, "")
                    item = QTableWidgetItem("" if text is None else str(text))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    item.setData(Qt.ItemDataRole.UserRole, (r, name))
                    if name not in editable or name == "X":
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    if name == "출고일":
                        item.setToolTip("클릭하여 달력에서 출고일 변경")
                    elif name == "상호" and row.get("_특수") is None:
                        item.setToolTip("클릭하여 거래처 변경")
                    if name == "가로" and row.get("_blue_width"):
                        item.setForeground(QColor("#1769d2"))
                    elif name == "세로" and row.get("_blue_height"):
                        item.setForeground(QColor("#1769d2"))
                    self.table.setItem(r, COL[name], item)

                if row.get("_특수") is not None:
                    for c in range(COL["상호"], COL["출고일"] + 1):
                        it = self.table.item(r, c)
                        if it:
                            it.setBackground(QColor("#fbfbfb"))

            self._apply_spans()
            self._update_summary()
        finally:
            self.table.blockSignals(False)
            self._loading = False
            if preserve:
                self.table.verticalScrollBar().setValue(scroll)
                if 0 <= current_row < self.table.rowCount():
                    self.table.setCurrentCell(current_row, max(COL["상태"], self.table.currentColumn()))

    def _update_summary(self):
        counts = {"red": 0, "review": 0, "yellow": 0}
        for level, _, _ in self.statuses:
            if level in counts:
                counts[level] += 1
        product_count = sum(1 for r in self.preview_rows if r.get("_특수") is None)
        summary_text = (
            f"주문 {len(self.orders)}건 · 장부 {product_count}행 · 오류 {counts['red']} · 확인 {counts['yellow']} · 중복/변경 {counts['review']}"
        )
        self.summary.setText(summary_text)
        if hasattr(self, "main_summary"):
            self.main_summary.setText(summary_text)
            self.main_status_title.setText("사전점검할 주문이 있습니다." if self.orders else "발주서를 업로드하면 장부가 여기에 쌓입니다.")
            self.preview_open_btn.setEnabled(bool(self.orders))
        self._update_order_buttons()

    def _auto_spans(self) -> list[tuple[int, int, int, int]]:
        """화면 자동 병합 (table 좌표).

        제품행 병합은 장부 엑셀과 같은 규칙(output.ledger_merge_ranges → _write_rows)을 그대로 쓴다.
        화면이 장부와 다른 병합 규칙을 가지면 병합/해제 결과가 엑셀과 어긋나기 때문이다.
        특수행(주소·포장·연락처 등)의 가로 병합만 화면 편집용 규칙을 따로 쓴다.
        """
        spans: list[tuple[int, int, int, int]] = []
        # 주소/전달/발신/포장 등 특수행의 가로 병합(화면 편집용).
        for r, row in enumerate(self.preview_rows):
            if row.get("_특수") is None:
                continue
            if row.get("_특수") in ("부속", "믹스"):
                continue
            if row.get("_detail_kind") == "packing":
                spans.append((r, COL["가로"], 1, 3))  # 가로-X-세로만 묶고 수량은 수정 가능
            elif row.get("_detail_kind") == "address":
                # 주소는 넓게 보이되 특이 칸은 배송방식(택배/화물/배달/내사) 수정용으로 남긴다.
                spans.append((r, COL["가로"], 1, COL["길이"] - COL["가로"] + 1))
            else:
                end_col = COL["기재사항1"] if row.get("선불") else COL["기재사항2"]
                spans.append((r, COL["가로"], 1, end_col - COL["가로"] + 1))

        out_to_table = {row.get("_output_index"): r for r, row in enumerate(self.preview_rows)
                        if isinstance(row.get("_output_index"), int)}
        for o1, o2, c1, c2 in ledger_merge_ranges(self.output_rows):
            t1, t2 = out_to_table.get(o1), out_to_table.get(o2)
            if t1 is None or t2 is None or t2 < t1:
                continue
            if (self.preview_rows[t1].get("_특수") is not None
                    and self.preview_rows[t1].get("_detail_kind") != "mix_part"):
                continue  # 특수행 가로 병합은 위의 화면 편집용 규칙을 쓴다(MIX 구성행은 장부와 동일).
            if c1 == COL["출고일"]:
                # 화면에만 있는 빈 주소행까지 같은 주문의 출고일 칸으로 묶는다.
                oi = self.preview_rows[t1].get("_oi")
                while t2 + 1 < len(self.preview_rows) and self.preview_rows[t2 + 1].get("_oi") == oi:
                    t2 += 1
            spans.append((t1, c1, t2 - t1 + 1, c2 - c1 + 1))
        return spans

    def _span_stable_spec(self, span: tuple[int, int, int, int]) -> tuple | None:
        r, c, rs, cs = span
        rows = self.preview_rows[r:r + rs]
        keys = [x.get("_row_key") for x in rows if x.get("_row_key") is not None]
        if not keys:
            return None
        return (tuple(keys), c, c + cs - 1)

    @staticmethod
    def _spec_overlap(a: tuple, b: tuple) -> bool:
        akeys, ac1, ac2 = a
        bkeys, bc1, bc2 = b
        return bool(set(akeys) & set(bkeys)) and not (ac2 < bc1 or bc2 < ac1)

    def _manual_spec_to_table_span(self, spec: tuple) -> tuple[int, int, int, int] | None:
        key_to_table = {r.get("_row_key"): i for i, r in enumerate(self.preview_rows)
                        if r.get("_row_key") is not None}
        keys, c1, c2 = spec
        table_rows = [key_to_table.get(k) for k in keys]
        if not table_rows or any(x is None for x in table_rows):
            return None
        if table_rows != list(range(min(table_rows), max(table_rows) + 1)):
            return None
        return (min(table_rows), min(c1, c2),
                max(table_rows) - min(table_rows) + 1,
                max(c1, c2) - min(c1, c2) + 1)

    @staticmethod
    def _table_span_overlap(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
        ar, ac, ars, acs = a
        br, bc, brs, bcs = b
        return not (ar + ars - 1 < br or br + brs - 1 < ar or ac + acs - 1 < bc or bc + bcs - 1 < ac)

    def _normalize_merge_specs(self):
        """수동 병합/해제 기록을 현재 행 구조에 맞춘다.

        행 고유 ID(_row_key)로 기억하므로 행을 추가·이동·삭제해도 같은 행을 다시 찾는다.
        - 병합: 같은 주문의 새 행이 범위 안에 끼어들면 엑셀의 행 삽입처럼 범위를 넓힌다.
          삭제된 행은 빼고, 남은 행이 연속하지 않거나 한 칸만 남으면 병합 기록을 지운다.
        - 해제: 남아 있는 행만 유지한다(그 행들이 다시 자동 병합되지 않게 막는 기록).
        """
        key_to_row = {r.get("_row_key"): i for i, r in enumerate(self.preview_rows)
                      if r.get("_row_key") is not None}

        def fix_merge(spec):
            keys, c1, c2 = spec
            rows = sorted(key_to_row[k] for k in keys if k in key_to_row)
            if not rows:
                return None
            top, bottom = rows[0], rows[-1]
            block = self.preview_rows[top:bottom + 1]
            if len(block) != len(rows):
                order_ids = {self.preview_rows[i].get("_oi") for i in rows}
                is_product = self.preview_rows[top].get("_특수") is None
                if not (len(order_ids) == 1 and all(
                        r.get("_row_key") is not None and r.get("_oi") in order_ids
                        and (r.get("_특수") is None) == is_product for r in block)):
                    return None
            if top == bottom and c1 == c2:
                return None
            return (tuple(r.get("_row_key") for r in block), c1, c2)

        def fix_unmerge(spec):
            keys, c1, c2 = spec
            kept = tuple(k for k in keys if k in key_to_row)
            return (kept, c1, c2) if kept else None

        merges, unmerges = [], []
        for spec in (fix_merge(x) for x in self.manual_merges):
            if spec and spec not in merges:
                merges.append(spec)
        for spec in (fix_unmerge(x) for x in self.manual_unmerges):
            if spec and spec not in unmerges:
                unmerges.append(spec)
        self.manual_merges, self.manual_unmerges = merges, unmerges

    def _apply_spans(self):
        self.table.clearSpans()
        self._normalize_merge_specs()
        manual_spans = []
        for spec in self.manual_merges:
            span = self._manual_spec_to_table_span(spec)
            if span and (span[2] > 1 or span[3] > 1):
                manual_spans.append(span)
        for span in self._auto_spans():
            stable = self._span_stable_spec(span)
            if stable and any(self._spec_overlap(stable, u) for u in self.manual_unmerges):
                continue
            # 사용자가 직접 병합한 범위와 겹치는 자동 병합은 장부와 같이 수동 병합이 우선한다.
            if any(self._table_span_overlap(span, m) for m in manual_spans):
                continue
            r, c, rs, cs = span
            if rs > 1 or cs > 1:
                self.table.setSpan(r, c, rs, cs)
        for r, c, rs, cs in manual_spans:
            self.table.setSpan(r, c, rs, cs)

    def _choose_calendar_date(self, current: date, title: str = "출고일 선택") -> date | None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        lay = QVBoxLayout(dialog)
        lay.setContentsMargins(12, 12, 12, 12)
        cal = QCalendarWidget(dialog)
        cal.setGridVisible(True)
        cal.setSelectedDate(QDate(current.year, current.month, current.day))
        lay.addWidget(cal)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            parent=dialog,
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        cal.activated.connect(lambda _qdate: dialog.accept())
        lay.addWidget(buttons)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None
        qd = cal.selectedDate()
        return date(qd.year(), qd.month(), qd.day())

    def _edit_order_ship_date(self, table_row: int) -> None:
        """출고 날짜는 달력으로만, 배송방식은 선택값으로 수정한다."""
        if not (0 <= table_row < len(self.preview_rows)):
            return
        oi = self.preview_rows[table_row].get("_oi")
        if not isinstance(oi, int) or not (0 <= oi < len(self.orders)):
            return
        order = self.orders[oi]
        raw = order.get("_ship_date")
        current = self._ship_date()
        if isinstance(raw, date):
            current = raw
        elif raw:
            try:
                current = date.fromisoformat(str(raw)[:10])
            except ValueError:
                pass

        dialog = QDialog(self)
        dialog.setWindowTitle("출고일 / 배송방식 변경")
        lay = QVBoxLayout(dialog)
        lay.setContentsMargins(12, 12, 12, 12)
        cal = QCalendarWidget(dialog)
        cal.setGridVisible(True)
        cal.setSelectedDate(QDate(current.year, current.month, current.day))
        lay.addWidget(cal)
        method_label = QLabel("배송표시")
        method_label.setObjectName("fieldLabel")
        lay.addWidget(method_label)
        method = NoWheelComboBox(dialog)
        method.addItem("없음", "")
        for name in ("배달", "택배", "화물", "내사"):
            method.addItem(name, name)
        actual = str((order.get("배송") or {}).get("방식") or "")
        idx = method.findData(actual)
        method.setCurrentIndex(idx if idx >= 0 else 0)
        lay.addWidget(method)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, parent=dialog)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        lay.addWidget(buttons)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        qd = cal.selectedDate()
        chosen = date(qd.year(), qd.month(), qd.day())
        selected_mode = str(method.currentData() or "")
        if chosen == current and selected_mode == actual:
            return
        self._push_undo()
        order["_ship_date"] = chosen.isoformat()
        if chosen != current:
            order["_ship_date_manual"] = True
        delivery = order.setdefault("배송", {})
        delivery["방식"] = selected_mode or None
        order["_delivery_mode_manual"] = True
        self.rebuild_table(preserve=True)

    def _edit_order_vendor(self, table_row: int) -> None:
        if not (0 <= table_row < len(self.preview_rows)):
            return
        row = self.preview_rows[table_row]
        if row.get("_특수") is not None:
            return
        oi = row.get("_oi")
        if not isinstance(oi, int) or not (0 <= oi < len(self.orders)):
            return
        order = self.orders[oi]
        current = order.get("거래처")

        dialog = QDialog(self)
        dialog.setWindowTitle("거래처 변경")
        lay = QVBoxLayout(dialog)
        lay.setContentsMargins(12, 12, 12, 12)
        combo = NoWheelComboBox(dialog)
        for code in self.mode_clients:
            combo.addItem(self.mode_client_label(code), code)
        idx = combo.findData(current)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        lay.addWidget(combo)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            parent=dialog,
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        lay.addWidget(buttons)
        combo.setFocus()
        combo.showPopup()
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        selected = combo.currentData()
        if not selected or selected == current:
            return

        self._push_undo()
        if self.product_mode == "roll_combo":
            old_default = (ROLL_CLIENT_INFO.get(current) or {}).get("delivery") if current else None
            new_default = (ROLL_CLIENT_INFO.get(selected) or {}).get("delivery")
        else:
            old_default = CLIENT_INFO.get(current, (None, None, None))[2] if current else None
            new_default = CLIENT_INFO.get(selected, (None, None, None))[2]
        delivery = order.setdefault("배송", {})
        if not order.get("_delivery_mode_manual") and (not delivery.get("방식") or delivery.get("방식") == old_default):
            delivery["방식"] = new_default
        order["거래처"] = selected
        for item in order.get("items") or []:
            item.pop("_manual_client", None)
        self.rebuild_table(preserve=True)

    def on_cell_clicked(self, row: int, col: int):
        if not (0 <= row < len(self.preview_rows)):
            return
        if col == COL["출고일"]:
            self._edit_order_ship_date(row)
            return
        if col == COL["상호"] and self.preview_rows[row].get("_특수") is None:
            self._edit_order_vendor(row)
            return
        if col != COL["상태"]:
            return
        oi = self.preview_rows[row].get("_oi")
        if not isinstance(oi, int) or not (0 <= oi < len(self.orders)):
            return

        level, status, tooltip = self.statuses[row] if row < len(self.statuses) else ("ok", "정상", "")
        values = row_display_values(self.preview_rows[row], row + 1)
        vendor = values.get("상호") or self.orders[oi].get("거래처") or "-"
        size = f"{values.get('가로') or '-'} X {values.get('세로') or '-'}"
        self.issue_head.setText(f"{status} · {vendor} · {size} · 수량 {values.get('수량') or '-'}")
        if level == "ok":
            reason = "검증 문제 없음 · 원본 확인용"
        else:
            reason = tooltip.strip() or "상태 원인이 기록되어 있지 않습니다."
        self.issue_reason.setText(reason)
        self.issue_reason.setProperty("level", level)
        self.issue_reason.style().unpolish(self.issue_reason)
        self.issue_reason.style().polish(self.issue_reason)
        self.detail_panel.show()
        self.source_wrap.show()
        self.source.show_order(self.orders[oi])
        self.left_scroll.verticalScrollBar().setValue(0)

    def on_item_changed(self, item: QTableWidgetItem):
        if self._loading:
            return
        row, col = item.row(), item.column()
        if not (0 <= row < len(self.preview_rows)):
            return
        column_name = HEADERS[col]
        if column_name not in editable_columns_for_row(self.preview_rows[row]):
            return
        self._push_undo()
        targets = self._merged_cell_rows(row, col)
        if len(targets) == 1:
            apply_preview_cell_edit(self.orders, self.preview_rows[row], column_name, item.text())
        else:
            # 병합된 셀은 엑셀처럼 하나의 값이다. 병합으로 가려진 아래 행에도 같은 값을 넣어
            # 장부에 보이는 값과 작업지시서·EDI에 들어가는 각 행의 값이 어긋나지 않게 한다.
            self._apply_edits_by_key([(self.preview_rows[r].get("_row_key"), column_name, item.text())
                                      for r in targets])
        self.rebuild_table(preserve=True)

    def _merged_cell_rows(self, row: int, col: int) -> list[int]:
        """(row, col)이 세로 병합의 첫 칸이면 병합에 포함된 제품 행 번호들, 아니면 [row]."""
        span = self.table.rowSpan(row, col)
        if span <= 1 or self.preview_rows[row].get("_특수") is not None:
            return [row]
        rows = [r for r in range(row, min(row + span, len(self.preview_rows)))
                if self.preview_rows[r].get("_특수") is None
                and self.preview_rows[r].get("_row_key") is not None]
        return rows or [row]

    STRUCTURE_EDIT_COLUMNS = {"수량", "방향", "색상"}

    def _apply_edits_by_key(self, edits):
        """제품행 수정들을 행 고유 ID로 대상을 찾아 순서대로 반영한다.

        수량/방향/색상 수정은 창 분할 등으로 행 위치를 바꿀 수 있으므로 그 다음 수정 전에 대상을 다시 찾는다.
        """
        index = {r.get("_row_key"): r for r in self.preview_rows if r.get("_row_key") is not None}
        for key, column_name, value in edits:
            if key is None:
                continue
            if index is None:
                preview, _ = build_preview_rows(self.orders, self._ship_date(), self.M, sort_di=False)
                index = {r.get("_row_key"): r for r in preview if r.get("_row_key") is not None}
            target = index.get(key)
            if target is not None and column_name in editable_columns_for_row(target):
                apply_preview_cell_edit(self.orders, target, column_name, value)
            if column_name in self.STRUCTURE_EDIT_COLUMNS:
                index = None

    def edit_current_cell(self):
        item = self.table.currentItem()
        if not item:
            return
        row, col = item.row(), item.column()
        if not (0 <= row < len(self.preview_rows)):
            return
        name = HEADERS[col]
        if name in editable_columns_for_row(self.preview_rows[row]):
            self.table.editItem(item)

    def delete_selected_cells(self):
        if hasattr(self, "pending_list") and self.pending_list.hasFocus():
            self.remove_selected_pending_files()
            return
        targets = []
        seen = set()
        for idx in self.table.selectedIndexes():
            row, col = idx.row(), idx.column()
            if not (0 <= row < len(self.preview_rows)):
                continue
            name = HEADERS[col]
            # 병합된 셀을 지우면 병합에 포함된 모든 행의 값을 함께 지운다.
            for r in self._merged_cell_rows(row, col):
                preview_row = self.preview_rows[r]
                marker = (preview_row.get("_row_key") or preview_row.get("_preview_id") or r, name)
                if marker in seen or name not in editable_columns_for_row(preview_row):
                    continue
                seen.add(marker)
                targets.append((r, name))
        if not targets:
            return
        self._push_undo()
        for r, name in targets:
            if self.preview_rows[r].get("_특수") is not None:
                apply_preview_cell_edit(self.orders, self.preview_rows[r], name, "")
        self._apply_edits_by_key([(self.preview_rows[r].get("_row_key"), name, "")
                                  for r, name in targets if self.preview_rows[r].get("_특수") is None])
        self.rebuild_table(preserve=True)

    def _selected_preview_row_numbers(self) -> list[int]:
        return sorted({idx.row() for idx in self.table.selectedIndexes()
                       if 0 <= idx.row() < len(self.preview_rows)})

    def _clear_manual_merge_overrides_after_structure_change(self):
        # 병합/해제는 행 고유 ID로 기억하므로 행을 추가·이동·삭제해도 지우지 않는다.
        # 다음 rebuild_table 에서 _normalize_merge_specs 가 현재 행 구조에 맞게 정리한다.
        return

    @staticmethod
    def _blank_row_like(base: dict) -> dict:
        item = copy.deepcopy(base)
        for key in ("가로", "세로", "수량", "손잡이방향", "손잡이길이",
                    "설치장소", "기재사항", "예외품목", "원문", "좌개수", "우개수"):
            item[key] = None
        # 기존 행의 '사람이 수정했다/긴급/분할됐다' 같은 숨은 상태가 새 행에 따라오지 않게 한다.
        for key in list(item):
            if key.startswith("_manual_") or key in {
                "_DI긴급", "_expanded_window", "_count_lr_mismatch", "_source_item_index"
            }:
                item.pop(key, None)
        item["창개수"] = 1
        item["연창"] = False
        item["_수동특이"] = None
        item["_manual_added"] = True
        item["확신도"] = {"가로": 1.0, "세로": 1.0, "품목코드": 1.0, "손잡이": 1.0}
        return item

    def add_row_below(self):
        rows = self._selected_preview_row_numbers()
        if len(rows) != 1:
            QMessageBox.information(self, "행 추가", "제품 행 하나를 선택해 주세요.")
            return
        row = self.preview_rows[rows[0]]
        if row.get("_특수") is not None:
            QMessageBox.information(self, "행 추가", "제품 행을 선택해 주세요.")
            return
        oi, ii = row.get("_oi"), row.get("_ii")
        if not isinstance(oi, int) or not isinstance(ii, int) or not (0 <= oi < len(self.orders)):
            return
        items = self.orders[oi].get("items") or []
        if not (0 <= ii < len(items)):
            return
        self._push_undo()
        items.insert(ii + 1, self._blank_row_like(items[ii]))
        self._clear_manual_merge_overrides_after_structure_change()
        self.rebuild_table(preserve=True)
        for tr, pr in enumerate(self.preview_rows):
            if pr.get("_oi") == oi and pr.get("_ii") == ii + 1 and pr.get("_특수") is None:
                self.table.setCurrentCell(tr, COL["가로"])
                break

    def move_selected_rows(self, direction: int):
        rows = self._selected_preview_row_numbers()
        if not rows:
            QMessageBox.information(self, "행 이동", "옮길 행을 먼저 선택해 주세요.")
            return
        selected = [self.preview_rows[r] for r in rows]
        if any(r.get("_특수") is not None for r in selected):
            QMessageBox.information(self, "행 이동", "주소·부속·포장 같은 특수행은 주문에 붙어 움직이므로 개별 이동할 수 없습니다.")
            return

        grouped: dict[int, list[int]] = {}
        for r in selected:
            oi, ii = r.get("_oi"), r.get("_ii")
            if isinstance(oi, int) and isinstance(ii, int):
                grouped.setdefault(oi, []).append(ii)
        if not grouped:
            return

        order_ids = sorted(grouped)
        # 여러 주문을 선택했으면 각 주문의 모든 보이는 품목이 선택된 경우 주문 묶음 자체를 이동한다.
        if len(order_ids) > 1:
            if order_ids != list(range(order_ids[0], order_ids[-1] + 1)):
                QMessageBox.information(self, "행 이동", "서로 붙어 있는 주문 묶음만 함께 이동할 수 있습니다.")
                return
            for oi in order_ids:
                visible = sorted({r.get("_ii") for r in self.preview_rows
                                  if r.get("_특수") is None and r.get("_oi") == oi and isinstance(r.get("_ii"), int)})
                if sorted(set(grouped[oi])) != visible:
                    QMessageBox.information(self, "행 이동", "여러 주문을 이동할 때는 각 주문의 품목 행 전체를 선택해 주세요.")
                    return
            if direction < 0 and order_ids[0] == 0:
                return
            if direction > 0 and order_ids[-1] >= len(self.orders) - 1:
                return
            self._push_undo()
            block = self.orders[order_ids[0]:order_ids[-1] + 1]
            del self.orders[order_ids[0]:order_ids[-1] + 1]
            insert_at = order_ids[0] - 1 if direction < 0 else order_ids[0] + 1
            self.orders[insert_at:insert_at] = block
            self._clear_manual_merge_overrides_after_structure_change()
            self.rebuild_table(preserve=False)
            return

        # 한 주문 안에서는 선택한 품목 행 블록을 한 칸 위/아래로 이동한다.
        oi = order_ids[0]
        indices = sorted(set(grouped[oi]))
        if indices != list(range(indices[0], indices[-1] + 1)):
            QMessageBox.information(self, "행 이동", "서로 붙어 있는 품목 행만 함께 이동할 수 있습니다.")
            return
        items = self.orders[oi].get("items") or []
        if direction < 0 and indices[0] == 0:
            # 주문 전체를 선택한 경우에는 위 주문과 위치를 바꾼다.
            visible = sorted({r.get("_ii") for r in self.preview_rows
                              if r.get("_특수") is None and r.get("_oi") == oi and isinstance(r.get("_ii"), int)})
            if indices == visible and oi > 0:
                self._push_undo()
                self.orders[oi - 1], self.orders[oi] = self.orders[oi], self.orders[oi - 1]
                self._clear_manual_merge_overrides_after_structure_change()
                self.rebuild_table(preserve=False)
            return
        if direction > 0 and indices[-1] >= len(items) - 1:
            visible = sorted({r.get("_ii") for r in self.preview_rows
                              if r.get("_특수") is None and r.get("_oi") == oi and isinstance(r.get("_ii"), int)})
            if indices == visible and oi < len(self.orders) - 1:
                self._push_undo()
                self.orders[oi], self.orders[oi + 1] = self.orders[oi + 1], self.orders[oi]
                self._clear_manual_merge_overrides_after_structure_change()
                self.rebuild_table(preserve=False)
            return

        self._push_undo()
        block = items[indices[0]:indices[-1] + 1]
        del items[indices[0]:indices[-1] + 1]
        insert_at = indices[0] - 1 if direction < 0 else indices[0] + 1
        items[insert_at:insert_at] = block
        self._clear_manual_merge_overrides_after_structure_change()
        self.rebuild_table(preserve=False)

    def _clear_special_row(self, row: dict) -> bool:
        kind = row.get("_detail_kind")
        column = {
            "address": "가로",
            "az_place": "가로",
            "recipient_contact": "가로",
            "delivery_notice": "가로",
            "sender": "가로",
            "packing": "수량",
            "manual_accessory": "색상",
        }.get(kind)
        if not column:
            return False
        apply_preview_cell_edit(self.orders, row, column, "")
        return True

    def delete_selected_rows(self):
        rows = self._selected_preview_row_numbers()
        if not rows:
            QMessageBox.information(self, "행 삭제", "삭제할 행을 먼저 선택해 주세요.")
            return
        ans = QMessageBox.question(
            self, "행 삭제", f"선택한 {len(rows)}개 행을 삭제할까요?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if ans != QMessageBox.StandardButton.Yes:
            return

        self._push_undo()
        grouped: dict[int, set[int]] = {}
        special_rows = []
        for table_row in rows:
            r = self.preview_rows[table_row]
            if r.get("_특수") is None:
                oi, ii = r.get("_oi"), r.get("_ii")
                if isinstance(oi, int) and isinstance(ii, int):
                    grouped.setdefault(oi, set()).add(ii)
            else:
                special_rows.append(r)

        for r in special_rows:
            self._clear_special_row(r)
        for oi, item_indices in sorted(grouped.items(), reverse=True):
            if not (0 <= oi < len(self.orders)):
                continue
            items = self.orders[oi].get("items") or []
            for ii in sorted(item_indices, reverse=True):
                if 0 <= ii < len(items):
                    items.pop(ii)
        self.orders = [o for o in self.orders if o.get("items")]
        self._clear_manual_merge_overrides_after_structure_change()
        self.rebuild_table(preserve=False)

    def _range_to_stable_spec(self, top: int, bottom: int, left: int, right: int,
                              allow_missing: bool = False):
        if left < COL["상호"]:
            return None
        keys = [self.preview_rows[r].get("_row_key") for r in range(top, bottom + 1)]
        if allow_missing:
            keys = [k for k in keys if k is not None]
            if not keys:
                return None
        elif any(k is None for k in keys):
            return None
        return (tuple(keys), left, right)

    def _expand_to_spans(self, top: int, bottom: int, left: int, right: int):
        """선택 사각형을 겹치는 기존 병합 셀까지 넓힌다(엑셀에서 병합 셀이 걸친 범위를 병합할 때와 같음)."""
        spans = self._current_spans()
        changed = True
        while changed:
            changed = False
            for r, c, rs, cs in spans:
                if r + rs - 1 < top or bottom < r or c + cs - 1 < left or right < c:
                    continue
                grown = (min(top, r), max(bottom, r + rs - 1), min(left, c), max(right, c + cs - 1))
                if grown != (top, bottom, left, right):
                    top, bottom, left, right = grown
                    changed = True
        return top, bottom, left, right

    def merge_selected(self):
        ranges = self.table.selectedRanges()
        if not ranges:
            return
        specs = []
        for rng in ranges:
            top, bottom, left, right = self._expand_to_spans(
                rng.topRow(), rng.bottomRow(), max(COL["상호"], rng.leftColumn()), rng.rightColumn())
            if top == bottom and left == right:
                continue
            spec = self._range_to_stable_spec(top, bottom, left, right)
            if not spec:
                QMessageBox.information(
                    self, "셀 병합",
                    "행번호/상태 열이나 아직 출력행이 아닌 빈 주소행은 병합할 수 없습니다.\n"
                    "주소를 먼저 입력한 뒤 다시 병합해 주세요."
                )
                return
            values = []
            for r in range(top, bottom + 1):
                for c in range(left, right + 1):
                    it = self.table.item(r, c)
                    if it and it.text().strip() and it.text().strip() not in values:
                        values.append(it.text().strip())
            if len(values) > 1:
                ans = QMessageBox.question(
                    self, "서로 다른 값 병합",
                    "선택 영역에 서로 다른 값이 있습니다. 병합하면 장부에서는 왼쪽 위 값만 보입니다.\n"
                    "내부 주문 데이터는 유지됩니다. 계속할까요?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if ans != QMessageBox.StandardButton.Yes:
                    return
            specs.append(spec)
        if not specs:
            return
        self._push_undo()
        for spec in specs:
            self.manual_unmerges = [u for u in self.manual_unmerges if not self._spec_overlap(u, spec)]
            self.manual_merges = [m for m in self.manual_merges if not self._spec_overlap(m, spec)]
            self.manual_merges.append(spec)
        self.rebuild_table(preserve=True)

    def _current_spans(self) -> list[tuple[int, int, int, int]]:
        spans = []
        seen = set()
        for r in range(self.table.rowCount()):
            for c in range(COL["상호"], self.table.columnCount()):
                rs = self.table.rowSpan(r, c)
                cs = self.table.columnSpan(r, c)
                if (rs > 1 or cs > 1) and (r, c) not in seen:
                    spans.append((r, c, rs, cs))
                    for rr in range(r, r + rs):
                        for cc in range(c, c + cs):
                            seen.add((rr, cc))
        return spans

    def unmerge_selected(self):
        ranges = self.table.selectedRanges()
        if not ranges:
            return
        candidates = []
        current_spans = self._current_spans()
        for rng in ranges:
            sel = (rng.topRow(), rng.bottomRow(), rng.leftColumn(), rng.rightColumn())
            for r, c, rs, cs in current_spans:
                span_sel = (r, r + rs - 1, c, c + cs - 1)
                table_overlap = not (span_sel[1] < sel[0] or sel[1] < span_sel[0]
                                     or span_sel[3] < sel[2] or sel[3] < span_sel[2])
                if table_overlap:
                    spec = self._range_to_stable_spec(span_sel[0], span_sel[1], span_sel[2], span_sel[3],
                                                      allow_missing=True)
                    if spec:
                        candidates.append(spec)
            # 화면 병합이 장부 병합과 같은 규칙이므로 실제로 보이는 병합만 해제한다.
        if not candidates:
            return
        self._push_undo()
        for spec in candidates:
            self.manual_merges = [m for m in self.manual_merges if not self._spec_overlap(m, spec)]
            if spec not in self.manual_unmerges:
                self.manual_unmerges.append(spec)
        self.rebuild_table(preserve=True)

    def export_outputs(self):
        if not self.orders:
            QMessageBox.information(self, "Excel 출력", "출력할 주문이 없습니다.")
            return
        # 수량과 좌/우 합계가 모순된 주문은 제작 창 수를 임의 추정할 수 없으므로 출력 전 반드시 수정한다.
        critical = []
        for oi, order in enumerate(self.orders, 1):
            for level, code, message, item_no in validate(order, self.M):
                if code in {"좌우수량불일치", "표전체수량불일치"}:
                    where = f"주문 {oi}" + (f" · {item_no}행" if item_no else "")
                    critical.append(f"{where}: {message}")
        if critical:
            QMessageBox.warning(
                self, "수량 확인 필요",
                "수량과 좌/우 개수가 맞지 않아 파일을 생성하지 않았습니다.\n"
                "사전점검에서 원본을 확인해 수정해 주세요.\n\n" + "\n".join(critical[:8])
            )
            return
        default_dir = ROOT / "out"
        default_dir.mkdir(parents=True, exist_ok=True)
        chosen = QFileDialog.getExistingDirectory(self, "출력 폴더 선택", str(default_dir))
        if not chosen:
            return
        try:
            self._set_busy(True)
            paths, rows = generate_outputs(self.orders, self._ship_date(), self.M, chosen)
            resolved_merges, merge_notes = resolve_stable_merge_specs(rows, self.manual_merges)
            resolved_unmerges, unmerge_notes = resolve_stable_merge_specs(rows, self.manual_unmerges)
            notes = merge_notes + unmerge_notes
            notes += apply_excel_merge_overrides(
                paths["장부"], rows, resolved_merges, resolved_unmerges)
            msg = "\n".join(f"{k}: {p.name}" for k, p in paths.items())
            if notes:
                msg += "\n\n병합 적용 안내:\n" + "\n".join(notes)
            box = QMessageBox(self)
            box.setWindowTitle("출력 완료")
            box.setText("장부·작업지시서·경영박사 파일을 생성했습니다.")
            box.setInformativeText(msg)
            open_btn = box.addButton("폴더 열기", QMessageBox.ButtonRole.ActionRole)
            box.addButton(QMessageBox.StandardButton.Ok)
            box.exec()
            if box.clickedButton() == open_btn:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(chosen))))
            self._remember_last_ledger(paths["장부"])

            # 파일 생성이 성공했으면 이번 배치의 화면 데이터는 종료 처리한다.
            # 생성한 Excel 파일은 그대로 두고, FitOrder 화면만 새 주문 입력 상태로 초기화한다.
            self.orders = []
            self.preview_rows = []
            self.output_rows = []
            self.statuses = []
            self.undo_stack.clear()
            self.manual_merges.clear()
            self.manual_unmerges.clear()
            self.pending_files.clear()
            self._update_pending_label()
            self.rebuild_table(preserve=False)
            self.source.image_label.setText("원본 없음")
            self.source.image_label.setPixmap(QPixmap())
            self.detail_panel.hide()
            self.issue_head.setText("상태 셀을 클릭하면 오류/확인 이유와 원본을 함께 볼 수 있습니다.")
            self.issue_reason.setText("검증 항목을 선택해 주세요.")
            self.progress_label.setText("파일 생성 완료")
            self.main_status_title.setText("발주서를 추가해 주세요.")
            self.main_status_text.setText("")
            self.right_stack.setCurrentWidget(self.home_right)
            self._update_order_buttons()
        except Exception as exc:
            QMessageBox.critical(self, "출력 실패", f"{type(exc).__name__}: {exc}")
        finally:
            self._set_busy(False)


class FitOrderWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"FitOrder v109 · {APP_VERSION}")
        self.resize(1680, 980)
        self.setMinimumSize(1250, 760)
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)
        self.start_page = StartPage()
        self.blind_page = OrderWorkspace("blind")
        self.roll_page = OrderWorkspace("roll_combo")
        self.holding_page = OrderWorkspace("holding")
        self.stack.addWidget(self.start_page)
        self.stack.addWidget(self.blind_page)
        self.stack.addWidget(self.roll_page)
        self.stack.addWidget(self.holding_page)
        self.start_page.selected.connect(self._select_mode)
        self.blind_page.back_requested.connect(lambda: self.stack.setCurrentWidget(self.start_page))
        self.roll_page.back_requested.connect(lambda: self.stack.setCurrentWidget(self.start_page))
        self.holding_page.back_requested.connect(lambda: self.stack.setCurrentWidget(self.start_page))
        self.stack.setCurrentWidget(self.start_page)

    def _select_mode(self, mode: str):
        if mode == "blind":
            self.stack.setCurrentWidget(self.blind_page)
        elif mode == "roll_combo":
            self.stack.setCurrentWidget(self.roll_page)
        else:
            self.stack.setCurrentWidget(self.holding_page)


def apply_style(app: QApplication):
    """Streamlit 버전의 흰색·각진 업무 화면을 PySide6에서 최대한 비슷하게 재현한다."""
    app.setStyle("Fusion")
    app.setFont(QFont("Malgun Gothic", 10))
    app.setStyleSheet("""
        QMainWindow, QDialog, QWidget { background: #ffffff; color: #262730; }
        QFrame#startCard, QFrame#infoPanel, QFrame#homeInfoCard {
            background: #ffffff; border: 1px solid #e1e5ea; border-radius: 0px;
        }
        QFrame#leftPanel { background: #ffffff; border: none; }
        QFrame#homeRightPanel { background: #ffffff; border-left: 1px solid #f0f2f6; }
        QFrame#ledgerPanel { background: #ffffff; border: none; }
        QFrame#precheckDetailPanel { background: #ffffff; border: 1px solid #e6e9ef; }
        QFrame#sourcePanel { background: #ffffff; border: none; }
        QFrame#divider { background: #e5e7eb; border: none; }
        QFrame#issueReasonBox { background: #ffffff; border: 1px solid #e0e4e9; }

        QLabel#startTitle { font-size: 34px; font-weight: 800; }
        QLabel#startSubtitle { font-size: 16px; color: #68707b; }
        QLabel#startNote, QLabel#mutedText { color: #737b86; font-size: 9.5pt; }
        QLabel#brandTitle { font-size: 23px; font-weight: 700; color: #262730; }
        QLabel#navCaption, QLabel#sectionCaption, QLabel#versionLabel { color: #6b7280; font-size: 9.5pt; }
        QLabel#sectionTitle { font-size: 14px; font-weight: 700; color: #20242a; }
        QLabel#fieldLabel { font-weight: 700; margin-top: 0px; }
        QLabel#pageTitle { font-size: 21px; font-weight: 800; }
        QLabel#precheckTitle { font-size: 22px; font-weight: 700; color: #262730; }
        QLabel#homeTitle { font-size: 24px; font-weight: 700; color: #262730; margin-top: 3px; }
        QLabel#homeSubtitle { font-size: 11pt; color: #69717c; margin-bottom: 8px; }
        QLabel#homeStatusTitle { font-size: 16px; font-weight: 700; color: #253042; }
        QLabel#homeSummary { font-size: 12pt; font-weight: 700; color: #374151; padding: 8px 0px; }
        QLabel#sourceTitle { font-size: 14px; font-weight: 700; }
        QLabel#summary { font-weight: 700; color: #374151; font-size: 10pt; }
        QLabel#shortcutHint { color: #6b7280; font-size: 9pt; padding: 0px 0px 4px 0px; }
        QLabel#pendingCount { color: #4b5563; font-weight: 700; margin-top: 0px; font-size: 9.5pt; }
        QLabel#progressLabel { color: #4b5563; min-height: 22px; }
        QLabel#issueHead { font-size: 10.5pt; font-weight: 700; color: #273244; }
        QLabel[level="red"] { color: #9b1c1c; }
        QLabel[level="yellow"] { color: #8a5a00; }
        QLabel[level="review"] { color: #62349a; }

        QListWidget#pendingFileList {
            background: #ffffff; border: 1px solid #d8dde5; padding: 2px; outline: 0; font-size: 10.5pt;
        }
        QListWidget#pendingFileList::item { min-height: 24px; padding: 3px 7px; border-bottom: 1px solid #eef1f4; }
        QListWidget#pendingFileList::item:selected { background: #e7f0ff; color: #1f2937; }

        QFrame#dropZone { background: #fafafa; border: 1px dashed #b9c1cc; }
        QFrame#dropZone:hover { background: #f4f7fb; border-color: #6b8fcf; }
        QLabel#dropTitle { font-weight: 700; color: #374151; font-size: 10.5pt; }
        QLabel#dropSub { color: #7b8490; font-size: 9pt; }

        QPushButton {
            background: #ffffff; border: 1px solid #cfd5dc; padding: 5px 9px;
            border-radius: 0px; min-height: 18px; font-size: 10pt;
        }
        QPushButton:hover { background: #f5f7fa; border-color: #aeb7c2; }
        QPushButton:disabled { color: #a9afb7; background: #f7f7f7; border-color: #e1e4e8; }
        QPushButton#modeButton { font-size: 23px; font-weight: 700; padding: 28px; }
        QPushButton#primaryButton { background: #ff4b4b; color: white; border-color: #ff4b4b; font-weight: 700; }
        QPushButton#primaryButton:hover { background: #e63f3f; }
        QPushButton#smallButton { padding: 3px 8px; min-height: 16px; font-size: 9.3pt; }
        QPushButton#tinyButton { padding: 2px 7px; min-height: 15px; font-size: 9pt; }
        QPushButton#dangerSmallButton { padding: 4px 9px; min-height: 18px; font-size: 9.5pt; color: #9b1c1c; border-color: #e4b7b7; }
        QPushButton#dangerSmallButton:hover { background: #fff3f3; }
        QPushButton#navButton { padding: 5px 10px; min-height: 18px; }

        QComboBox, QDateEdit {
            background: white; border: 1px solid #cfd5dc; padding: 4px 7px;
            border-radius: 0px; min-height: 20px; font-size: 10pt;
        }
        QCheckBox { spacing: 7px; font-size: 10pt; }

        QTableWidget {
            background: white; gridline-color: #cfd4da; selection-background-color: #dbeafe;
            selection-color: #111111; border: 1px solid #cfd4da; font-size: 10pt;
        }
        QTableWidget::item { padding: 3px 4px; }
        QHeaderView::section {
            background: #f2f3f5; border: 1px solid #cfd4da; padding: 7px 5px;
            font-weight: 700; color: #252b33;
        }
        QScrollArea { background: white; border: 1px solid #d9dde3; }
        QScrollArea#leftScroll { border: none; background: white; }
        QProgressBar { border: 1px solid #d5dae0; background: #f5f6f7; text-align: center; border-radius: 0px; }
        QProgressBar::chunk { background: #ff4b4b; }
    """)

def main() -> int:
    app = QApplication(sys.argv)
    apply_style(app)
    try:
        win = FitOrderWindow()
    except Exception as exc:
        QMessageBox.critical(None, "FitOrder 시작 실패", f"{type(exc).__name__}: {exc}")
        return 1
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
