"""Screen 6: Edit History (audit log) - filters, old (strikethrough) -> new, Restore, Export to Excel."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (QAbstractItemView, QComboBox, QFileDialog, QHBoxLayout, QHeaderView, QLineEdit,
                               QTableWidget, QTableWidgetItem)

from ...services import auth
from ...services.auth import PermissionDenied
from ...excel.sheets import SHEETS
from ...services import views
from ...services.edits import EditError
from ..dialogs import confirm, error, info
from ..i18n import t
from ..widgets import DatePicker, button, label, wrap
from .base import Screen


class HistoryScreen(Screen):
    key = "history"

    def __init__(self, ctx):
        super().__init__(ctx, scroll=False)
        self.item_id, self.rows = None, []
        top = QHBoxLayout()
        self.title = label("", "h1")
        top.addWidget(self.title)
        top.addStretch(1)
        self.item_lab = label("", "muted")
        self.clear_item = button("")
        self.clear_item.clicked.connect(lambda: self.show_item(None))
        self.export_btn = button("")
        self.export_btn.clicked.connect(self._export)
        for w in (self.item_lab, self.clear_item, self.export_btn):
            top.addWidget(w)
        self.lay.addLayout(top)
        f = QHBoxLayout()
        self.text = QLineEdit()
        self.text.returnPressed.connect(self.refresh)
        self.page, self.editor, self.kind = QComboBox(), QComboBox(), QComboBox()
        self.dfrom, self.dto = DatePicker(), DatePicker()
        for w in (self.page, self.editor, self.kind):
            w.setMinimumWidth(170)
            w.currentIndexChanged.connect(self._filters_changed)
        self.text.setMinimumWidth(220)
        self.dfrom.changed.connect(self.refresh)
        self.dto.changed.connect(self.refresh)
        self.from_lab, self.to_lab = label(""), label("")
        for w in (self.text, self.page, self.editor, self.kind, self.from_lab, self.dfrom, self.to_lab, self.dto):
            f.addWidget(w)
        self.lay.addLayout(f)
        self.table = QTableWidget()
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(True)
        self.lay.addWidget(self.table, 1)
        self._loading = False

    def show_item(self, item_id=None, **_):
        self.item_id = item_id
        self.refresh()

    def _filters_changed(self):
        if not self._loading:
            self.refresh()

    def _fill_combo(self, cb, first, values, labels=None):
        cur = cb.currentData()
        cb.clear()
        cb.addItem(first, None)
        for i, v in enumerate(values):
            cb.addItem(labels[i] if labels else v, v)
        cb.setCurrentIndex(max(cb.findData(cur), 0))

    def refresh(self):
        self._loading = True
        self.title.setText(t("nav.history"))
        self.export_btn.setText(t("hist.export"))
        self.clear_item.setText(t("hist.clear_item"))
        self.item_lab.setText(t("hist.item_filter", id=self.item_id) if self.item_id else "")
        self.item_lab.setVisible(bool(self.item_id))
        self.clear_item.setVisible(bool(self.item_id))
        self.text.setPlaceholderText(t("search"))
        self.from_lab.setText(t("hist.from"))
        self.to_lab.setText(t("hist.to"))
        self._fill_combo(self.page, t("hist.page_all"), list(SHEETS), [t(f"sheet.{s}") for s in SHEETS])
        self._fill_combo(self.editor, t("hist.editor_all"), self.ctx.repo.audit_editors())
        self._fill_combo(self.kind, t("hist.type_all"), list(views.HISTORY_TYPES),
                         [t(f"type.{k}") for k in views.HISTORY_TYPES])
        self._loading = False
        rows = views.history(self.ctx.repo, text=(self.text.text().strip() or self.ctx.search or None),
                             sheet=self.page.currentData(), editor=self.editor.currentData(),
                             date_from=self.dfrom.value(), date_to=self.dto.value(), kind=self.kind.currentData(),
                             item_id=self.item_id)
        self.rows = rows
        cols = [t("col.date"), t("col.editor"), t("col.type"), t("col.page"), t("col.field"), t("col.old"),
                t("col.new"), t("col.reason"), ""]
        tb = self.table
        tb.clearContents()
        tb.setColumnCount(len(cols))
        tb.setHorizontalHeaderLabels(cols)
        tb.setRowCount(len(rows))
        strike = QFont()
        strike.setStrikeOut(True)
        for i, e in enumerate(rows):
            vals = [e["ts"][:16].replace("T", " "), e["editor"], t(f"type.{e['type']}"),
                    f"{e['item_id'] or ''}\n{t('sheet.' + e['sheet']) if e['sheet'] in SHEETS else (e['sheet'] or '')}",
                    e["field"] or "", e["old_value"] or "", e["new_value"] or "", e["reason"] or ""]
            for j, v in enumerate(vals):
                it = QTableWidgetItem(v if len(v) <= 120 else v[:117] + "…")
                if len(v) > 120:
                    it.setToolTip(v[:2000])
                if j == 5 and v:
                    it.setFont(strike)
                    it.setForeground(Qt.darkGray)
                tb.setItem(i, j, it)
            if e["restorable"] and self.ctx.can(auth.RESTORE):
                b = button(t("hist.restore"))
                b.clicked.connect(lambda _=False, x=e: self._restore(x))
                tb.setCellWidget(i, 8, wrap(b))
        hh = tb.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)
        for j, w in enumerate((130, 170, 110, 120, 130, 200, 200, 260, 120)):
            tb.setColumnWidth(j, w)
        hh.setStretchLastSection(False)
        tb.resizeRowsToContents()

    def _restore(self, e):
        if not confirm(self, t("hist.restore_q", field=e["field"], item=e["item_id"], old=e["old_value"] or t("none"))):
            return
        try:
            views.restore(self.ctx.cfg, self.ctx.repo, e["id"], actor=self.ctx.actor)
        except (EditError, ValueError, PermissionDenied) as exc:
            error(self, str(exc))
            return
        self.ctx.refresh_all()

    def _export(self):
        path, _ = QFileDialog.getSaveFileName(self, t("hist.export"), "Edit History.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        views.export_history(self.rows, path)
        info(self, path)
