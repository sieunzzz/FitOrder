from __future__ import annotations
import copy, sys
from datetime import date
from pathlib import Path
from typing import Any

from PySide6.QtCore import QDate, Qt
from PySide6.QtGui import QAction, QColor, QKeySequence
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDateEdit, QFileDialog, QHBoxLayout, QLabel,
    QMainWindow, QMessageBox, QPushButton, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget, QHeaderView, QAbstractItemView
)

from . import APP_VERSION
from .desktop_workflow import (
    CLIENTS, LEDGER_COLUMNS, Master, apply_preview_cell_edit, build_preview_rows,
    client_label, default_ship, editable_columns_for_row, generate_outputs,
    ingest_file, preview_statuses, row_display_values
)

HEADERS = ["No","상태"] + LEDGER_COLUMNS
COL = {name:i for i,name in enumerate(HEADERS)}

class FitOrderWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"FitOrder · {APP_VERSION}")
        self.resize(1600, 920)
        self.orders = []
        self.preview_rows = []
        self.statuses = []
        self.undo_stack = []
        self.M = Master()

        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)

        top = QHBoxLayout()
        top.addWidget(QLabel("FitOrder"))
        self.ship = QDateEdit()
        self.ship.setCalendarPopup(True)
        d = default_ship()
        self.ship.setDate(QDate(d.year,d.month,d.day))
        top.addWidget(QLabel("출고일"))
        top.addWidget(self.ship)
        self.client = QComboBox()
        self.client.addItem("자동 판별", None)
        for c in CLIENTS:
            self.client.addItem(client_label(c), c)
        top.addWidget(self.client)
        self.add_btn = QPushButton("발주서 추가")
        self.add_btn.clicked.connect(self.choose_files)
        top.addWidget(self.add_btn)
        self.output_btn = QPushButton("장부·작업지시서·EDI 생성")
        self.output_btn.clicked.connect(self.output_files)
        top.addWidget(self.output_btn)
        top.addStretch(1)
        outer.addLayout(top)

        self.summary = QLabel("주문 0건 · 장부 0행")
        outer.addWidget(self.summary)

        self.table = QTableWidget(0,len(HEADERS))
        self.table.setHorizontalHeaderLabels(HEADERS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.EditKeyPressed)
        self.table.itemChanged.connect(self.on_item_changed)
        outer.addWidget(self.table,1)

        bottom = QHBoxLayout()
        self.delete_btn = QPushButton("선택 행 삭제")
        self.delete_btn.clicked.connect(self.delete_rows)
        self.clear_btn = QPushButton("전체 비우기")
        self.clear_btn.clicked.connect(self.clear_all)
        bottom.addWidget(self.delete_btn)
        bottom.addWidget(self.clear_btn)
        bottom.addStretch(1)
        outer.addLayout(bottom)

        self._loading = False
        self._install_actions()
        self.rebuild()

    def _install_actions(self):
        a = QAction("삭제",self); a.setShortcut(QKeySequence("Delete")); a.triggered.connect(self.delete_cells); self.addAction(a)
        u = QAction("실행취소",self); u.setShortcut(QKeySequence("Ctrl+Z")); u.triggered.connect(self.undo); self.addAction(u)
        e = QAction("수정",self); e.setShortcut(QKeySequence("F2")); e.triggered.connect(self.edit_current); self.addAction(e)

    def _ship_date(self):
        q = self.ship.date()
        return date(q.year(),q.month(),q.day())

    def choose_files(self):
        paths,_ = QFileDialog.getOpenFileNames(self,"발주서 선택","", "Orders (*.json *.xlsx *.xls *.png *.jpg *.jpeg *.webp)")
        if not paths: return
        self.undo_stack.append(copy.deepcopy(self.orders))
        forced = self.client.currentData()
        errors=[]
        for p in paths:
            try:
                self.orders.extend(ingest_file(p,self.M,forced))
            except Exception as exc:
                errors.append(f"{Path(p).name}: {exc}")
        self.rebuild()
        if errors:
            QMessageBox.warning(self,"일부 파일 처리 실패","\n".join(errors))

    def rebuild(self):
        self._loading=True
        try:
            self.preview_rows,_ = build_preview_rows(self.orders,self._ship_date(),self.M)
            self.statuses = preview_statuses(self.preview_rows,self.orders,self.M)
            self.table.blockSignals(True)
            self.table.setRowCount(len(self.preview_rows))
            for r,row in enumerate(self.preview_rows):
                level,status,tip = self.statuses[r]
                vals = row_display_values(row,r+1)
                ni=QTableWidgetItem(str(r+1)); ni.setFlags(ni.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(r,COL["No"],ni)
                si=QTableWidgetItem(status); si.setToolTip(tip); si.setFlags(si.flags() & ~Qt.ItemFlag.ItemIsEditable)
                si.setBackground({"red":QColor("#ffd9d9"),"yellow":QColor("#fff1b8"),"info":QColor("#e2ebf6")}.get(level,QColor("#f5f7f9")))
                self.table.setItem(r,COL["상태"],si)
                editable=editable_columns_for_row(row)
                for name in LEDGER_COLUMNS:
                    it=QTableWidgetItem(str(vals.get(name,"") or ""))
                    if name not in editable or name=="X": it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self.table.setItem(r,COL[name],it)
            self.table.blockSignals(False)
            self.summary.setText(f"주문 {len(self.orders)}건 · 장부 {len(self.preview_rows)}행")
        finally:
            self._loading=False

    def on_item_changed(self,item):
        if self._loading: return
        r,c=item.row(),item.column()
        if not (0<=r<len(self.preview_rows)): return
        name=HEADERS[c]
        if name not in editable_columns_for_row(self.preview_rows[r]): return
        self.undo_stack.append(copy.deepcopy(self.orders))
        apply_preview_cell_edit(self.orders,self.preview_rows[r],name,item.text())
        self.rebuild()

    def edit_current(self):
        it=self.table.currentItem()
        if it and HEADERS[it.column()] in editable_columns_for_row(self.preview_rows[it.row()]):
            self.table.editItem(it)

    def delete_cells(self):
        targets=[]
        for idx in self.table.selectedIndexes():
            if idx.column()<2 or idx.row()>=len(self.preview_rows): continue
            name=HEADERS[idx.column()]
            if name in editable_columns_for_row(self.preview_rows[idx.row()]): targets.append((idx.row(),name))
        if not targets: return
        self.undo_stack.append(copy.deepcopy(self.orders))
        for r,name in targets:
            apply_preview_cell_edit(self.orders,self.preview_rows[r],name,"")
        self.rebuild()

    def delete_rows(self):
        selected=sorted({i.row() for i in self.table.selectedIndexes()},reverse=True)
        if not selected: return
        self.undo_stack.append(copy.deepcopy(self.orders))
        refs=[]
        for r in selected:
            row=self.preview_rows[r]
            if row.get("_특수") is None and isinstance(row.get("_oi"),int) and isinstance(row.get("_ii"),int):
                refs.append((row["_oi"],row["_ii"]))
        for oi,ii in sorted(refs,reverse=True):
            if 0<=oi<len(self.orders) and 0<=ii<len(self.orders[oi].get("items",[])):
                self.orders[oi]["items"].pop(ii)
        self.orders=[o for o in self.orders if o.get("items")]
        self.rebuild()

    def undo(self):
        if self.undo_stack:
            self.orders=self.undo_stack.pop()
            self.rebuild()

    def clear_all(self):
        if self.orders:
            self.undo_stack.append(copy.deepcopy(self.orders))
            self.orders=[]
            self.rebuild()

    def output_files(self):
        if not self.orders:
            QMessageBox.information(self,"출력","출력할 주문이 없습니다."); return
        out=QFileDialog.getExistingDirectory(self,"출력 폴더")
        if not out: return
        try:
            result=generate_outputs(self.orders,self._ship_date(),out,self.M)
            QMessageBox.information(self,"완료","\n".join(f"{k}: {v}" for k,v in result.items()))
        except Exception as exc:
            QMessageBox.critical(self,"출력 실패",f"{type(exc).__name__}: {exc}")

def main():
    app=QApplication(sys.argv)
    app.setStyle("Fusion")
    win=FitOrderWindow()
    win.show()
    return app.exec()

if __name__=="__main__":
    raise SystemExit(main())
