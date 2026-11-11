"""Screen 4: Test Progress - summary, filters, table, update results, upload result file."""

from __future__ import annotations

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (QAbstractItemView, QFileDialog, QHBoxLayout, QHeaderView, QTableWidget,
                               QTableWidgetItem, QVBoxLayout, QWidget)

from ...services import auth
from ...services.auth import PermissionDenied
from ...excel.sheets import TEST_PROGRESS
from ...parsing.dates import fmt
from ...services import views
from ...services.attachments import UploadError, upload
from ..dialogs import AddApplicationDialog, UpdateDialog, error, info
from ..i18n import t
from ..widgets import bar, button, card, label, tag, wrap
from .base import ChipRow, Screen, clear_layout


class TestProgressScreen(Screen):
    key = "tp"

    def __init__(self, ctx):
        super().__init__(ctx, scroll=False)
        self.flt, self.selected = "all", None
        top = QHBoxLayout()
        self.title = label("", "h1")
        top.addWidget(self.title)
        top.addStretch(1)
        self.add_btn = button("")                     # multi-project: + Add Application
        self.add_btn.clicked.connect(self._add)
        top.addWidget(self.add_btn)
        self.up_btn = button("")
        self.up_btn.clicked.connect(self._upload)
        self.upd_btn = button("", primary=True)
        self.upd_btn.clicked.connect(self._update)
        top.addWidget(self.up_btn)
        top.addWidget(self.upd_btn)
        self.lay.addLayout(top)
        self.summary = QHBoxLayout()
        self.summary.setSpacing(16)
        self.lay.addLayout(self.summary)
        self.chips = ChipRow(self._pick)
        self.lay.addWidget(self.chips)
        self.table = QTableWidget()
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.itemSelectionChanged.connect(self._select)
        self.table.cellDoubleClicked.connect(lambda r, c: self._update())
        self.lay.addWidget(self.table, 1)
        self.rows = []

    def show_item(self, item_id=None, filter=None):
        if filter:
            self.flt = filter
        if item_id:
            self.flt, self.selected = "all", item_id
        self.refresh()

    def refresh(self):
        self.title.setText(t("nav.tp"))
        self.add_btn.setText(t("tp.add"))
        self.add_btn.setVisible(self.ctx.can(auth.ADD_APP))
        self.up_btn.setText(t("tp.upload"))
        self.upd_btn.setText(t("tp.update"))
        s, rows = views.test_progress(self.ctx.cfg, self.ctx.repo, self.ctx.as_of, self.flt, self.ctx.search)
        self._items = self.ctx.repo.load_items(TEST_PROGRESS)
        self.rows = rows
        clear_layout(self.summary)
        for title, value in ((t("tp.executed"), f"{s['executed']:,} / {s['total']:,}"), (t("tp.pass"), f"{s['pass']:,}"),
                             (t("tp.fail"), f"{s['fail']:,}"), (t("tp.amber"), str(s["amber"]))):
            c, cl = card(margins=(16, 12, 16, 12), spacing=4)
            cl.addWidget(label(title, "muted"))
            cl.addWidget(label(value, "kpi"))
            self.summary.addWidget(c)
        self.chips.set_chips([(k, t(f"f.{k}")) for k in views.TP_FILTERS], self.flt)
        cols = [t("col.app"), t("col.reporter"), t("col.progress"), t("col.passfail"), t("col.pstart"),
                t("col.pfinish"), t("col.astart"), t("col.afinish"), t("col.upload_tc"), t("col.light")]
        tb = self.table
        tb.clearContents()
        tb.setColumnCount(len(cols))
        tb.setHorizontalHeaderLabels(cols)
        tb.setRowCount(len(rows))
        sel_row = None
        for i, r in enumerate(rows):
            app = QTableWidgetItem(f"{r['app']}\n{r['id']}")
            app.setData(Qt.UserRole, r["id"])
            tb.setItem(i, 0, app)
            tb.setItem(i, 1, QTableWidgetItem(r["reporter"]))
            pw = QWidget()
            pl = QVBoxLayout(pw)
            pl.setContentsMargins(8, 6, 8, 6)
            pl.setSpacing(4)
            pct = r["executed"] / r["total"] * 100 if r["total"] else 0
            pl.addWidget(label(f"{r['executed']} / {r['total']} · {pct:.0f}%"))
            pl.addWidget(bar(r["executed"], r["total"]))
            tb.setCellWidget(i, 2, pw)
            pf = QWidget()
            pfl = QHBoxLayout(pf)
            pfl.setContentsMargins(8, 2, 8, 2)
            pfl.addWidget(label(f"{r['pass']} / {r['fail']}"))
            if r["mismatch"]:
                pfl.addWidget(tag("mismatch", t("tag.mismatch")))
            pfl.addStretch(1)
            tb.setCellWidget(i, 3, pf)
            # Planned Start | Planned Finish | Actual Start | Actual Finish (same order as the Excel sheet)
            tb.setItem(i, 4, QTableWidgetItem(fmt(r["planned_start"]) if r["planned_start"] else ""))
            tb.setItem(i, 5, QTableWidgetItem(fmt(r["planned_finish"]) if r["planned_finish"] else ""))
            for col, key in ((6, "actual_start"), (7, "actual_finish")):
                if r[key]:
                    tb.setItem(i, col, QTableWidgetItem(fmt(r[key])))
                else:
                    tb.setCellWidget(i, col, wrap(tag("not_filled", t("tag.not_filled"))))
            tb.setCellWidget(i, 8, self._upload_cell(r))
            tb.setCellWidget(i, 9, wrap(tag(r["color"], t(f"tag.{r['color']}"))))
            tb.cellWidget(i, 9).setToolTip(r["reason"])
            tb.setRowHeight(i, 60)
            if r["id"] == self.selected:
                sel_row = i
        hh = tb.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)
        hh.setStretchLastSection(False)
        tb.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        tb.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        # Every column at least as wide as its header + content (incl. tag/button cell widgets);
        # a horizontal scrollbar appears when the window is narrower than the total.
        tb.resizeColumnsToContents()
        self._col_min = []
        for col in range(tb.columnCount()):
            w = max(tb.columnWidth(col), self.MIN_WIDTHS.get(col, 0)) + 16
            self._col_min.append(w)
            tb.setColumnWidth(col, w)
        self._fit_columns()
        if sel_row is not None:
            tb.selectRow(sel_row)

    MIN_WIDTHS = {0: 220, 1: 120, 2: 170}

    def _fit_columns(self):
        """Give spare width (window wider than content) to the Application column."""
        mins = getattr(self, "_col_min", None)
        if not mins:
            return
        spare = self.table.viewport().width() - sum(mins)
        self.table.setColumnWidth(0, mins[0] + max(0, spare))

    def resizeEvent(self, e):
        super().resizeEvent(e)
        QTimer.singleShot(0, self._fit_columns)       # after the table itself has been resized

    def _upload_cell(self, r) -> QWidget:
        """TC/Result Complete (column 'Upload TC/Result'): Complete, or not yet + an Upload button."""
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(6, 2, 6, 2)
        lay.setSpacing(6)
        if r["tc_result"] == "Complete":
            lay.addWidget(tag("Complete", t("tag.Complete")))
            if r["uploaded"]:
                w.setToolTip(t("tp.has_file"))
        else:
            lay.addWidget(tag("not_filled", t("tag.not_uploaded")))
            b = button(t("tp.upload_short"))
            b.setAccessibleName(f"{t('tp.upload_short')} {r['app']}")
            b.clicked.connect(lambda _=False, item=r["id"]: self._upload_for(item))
            if not self.ctx.can(auth.UPLOAD, TEST_PROGRESS, self._items.get(r["id"])):
                b.setEnabled(False)              # Part 5: QA - not their row
                b.setToolTip(t("auth.not_yours"))
            lay.addWidget(b)
        lay.addStretch(1)
        return w

    def _upload_for(self, item_id: str):
        self.selected = item_id
        self._upload()

    def _pick(self, k):
        self.flt = k
        self.refresh()

    def _select(self):
        items = self.table.selectedItems()
        if items:
            self.selected = self.table.item(items[0].row(), 0).data(Qt.UserRole)

    def _update(self):
        if not self.selected:
            info(self, t("tp.select"))
            return
        if UpdateDialog(self.ctx, TEST_PROGRESS, self.selected, self).exec():
            self.ctx.refresh_all()

    def _upload(self):
        if not self.selected:
            info(self, t("tp.select"))
            return
        if not self.ctx.can(auth.UPLOAD, TEST_PROGRESS, self.ctx.repo.get_item(TEST_PROGRESS, self.selected)):
            error(self, t("auth.not_yours"))
            return
        types = " ".join(f"*.{x}" for x in self.ctx.cfg.alerts.attachment_types)
        path, _ = QFileDialog.getOpenFileName(self, t("tp.upload"), "", f"Files ({types})")
        if not path:
            return
        try:
            upload(self.ctx.cfg, self.ctx.repo, self.selected, path, actor=self.ctx.actor)
        except (UploadError, PermissionDenied) as exc:
            error(self, str(exc))
            return
        info(self, t("tp.uploaded"))
        self.ctx.refresh_all()

    def _add(self):
        dlg = AddApplicationDialog(self.ctx, self)
        if dlg.exec():
            self.selected = dlg.new_id
            self.ctx.refresh_all()
