"""Part 6 UI: Generate Weekly Report (Dashboard) -> preview -> Send / Cancel, and the past-reports screen."""

from __future__ import annotations

import json

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (QAbstractItemView, QDialog, QHBoxLayout, QHeaderView, QProgressDialog, QSplitter,
                               QTableWidget, QTableWidgetItem, QTextBrowser, QVBoxLayout, QWidget)

from ..mail.sender import OutlookSender, SendError
from ..services import auth, weekly_report
from ..services.auth import PermissionDenied
from .dialogs import error, info
from .i18n import t
from .screens.base import Screen
from .widgets import button, label


class GenerateWorker(QThread):
    """Re-reads Excel and builds the report on its own DB connection (the AI call can take ~30 s)."""
    done = Signal(int)
    failed = Signal(str)

    def __init__(self, cfg, actor, as_of):
        super().__init__()
        self.cfg, self.actor, self.as_of = cfg, actor, as_of

    def run(self):
        from ..services.sync import SyncService
        from ..storage.sqlite_repo import SqliteRepository
        repo = SqliteRepository(self.cfg.db_path, shared=self.cfg.shared_db)
        try:
            try:
                SyncService(self.cfg, repo).pull()              # latest Excel data first
            except Exception as exc:                            # keep going on the data we have, but say so
                self.note = f"Excel could not be re-read ({exc}); the report uses the data last read."
            g = weekly_report.generate(self.cfg, repo, actor=self.actor, as_of=self.as_of, trigger="manual")
            self.done.emit(g.report_id)
        except PermissionDenied as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            self.failed.emit(f"Weekly report failed: {exc}")
        finally:
            repo.close()


def start_generate(ctx, parent) -> None:
    """Dashboard button: busy window while the report is built, then the preview."""
    if not ctx.can(auth.WEEKLY):
        error(parent, t("auth.not_allowed"))
        return
    prog = QProgressDialog(t("wk.generating"), None, 0, 0, parent)
    prog.setWindowTitle(t("dash.weekly"))
    prog.setWindowModality(Qt.WindowModal)
    prog.setMinimumDuration(0)
    w = GenerateWorker(ctx.cfg, ctx.actor, ctx.as_of)
    parent._weekly_worker = w                                    # keep a reference while it runs

    def ok(report_id):
        prog.close()
        WeeklyPreviewDialog(ctx, report_id, parent, note=getattr(w, "note", "")).exec()
        if ctx.refresh_all:
            ctx.refresh_all()

    def bad(msg):
        prog.close()
        error(parent, msg)
    w.done.connect(ok)
    w.failed.connect(bad)
    w.start()
    prog.show()


def _ai_line(r: dict) -> str:
    if r["ai_used"]:
        cost = f", ${r['cost_usd']:.4f}" if r.get("cost_usd") is not None else ""
        return t("wk.ai_yes", m=r["ai_model"], i=r["input_tokens"] or 0, o=r["output_tokens"] or 0) + cost
    return t("wk.ai_no", e=r["ai_error"] or "-")


class WeeklyPreviewDialog(QDialog):
    def __init__(self, ctx, report_id: int, parent=None, note: str = ""):
        super().__init__(parent)
        self.ctx, self.report_id = ctx, report_id
        r = ctx.repo.get_weekly(report_id)
        self.setWindowTitle(t("wk.preview_title", id=report_id))
        self.resize(980, 820)
        lay = QVBoxLayout(self)
        to = json.loads(r["recipients"] or "[]")
        dry = ctx.cfg.alerts.dry_run
        who = (t("wk.to_dry", p=ctx.cfg.alerts.pmo_email or t("wk.no_pmo")) if dry else
               (", ".join(to) if to else t("wk.to_pmo_only", p=ctx.cfg.alerts.pmo_email or t("wk.no_pmo"))))
        lay.addWidget(label(t("wk.recipients", n=len(to), w=who), wrap_text=True))
        lay.addWidget(label(_ai_line(r), "muted", wrap_text=True))
        if note:
            lay.addWidget(label(note, wrap_text=True))
        self.view = QTextBrowser()
        self.view.setHtml(r["html"])
        lay.addWidget(self.view, 1)
        row = QHBoxLayout()
        row.addStretch(1)
        self.cancel_btn, self.send_btn = button(t("cancel")), button(t("wk.send"), primary=True)
        self.cancel_btn.clicked.connect(self.on_cancel)
        self.send_btn.clicked.connect(self.on_send)
        row.addWidget(self.cancel_btn)
        row.addWidget(self.send_btn)
        lay.addLayout(row)
        self.result_text = ""

    def on_send(self):
        try:
            res = weekly_report.send(self.ctx.cfg, self.ctx.repo, OutlookSender(self.ctx.cfg.alerts.send_account),
                                     self.report_id, actor=self.ctx.actor)
        except (PermissionDenied, ValueError, SendError, RuntimeError) as exc:
            error(self, str(exc))
            return
        self.result_text = res
        info(self, t("wk.sent", r=res))
        self.accept()

    def on_cancel(self):
        self.reject()

    def reject(self):
        """Cancel button or the window's X: a report still waiting is marked cancelled (kept in the list)."""
        r = self.ctx.repo.get_weekly(self.report_id)
        if r and r["status"] == "preview":
            try:
                weekly_report.cancel(self.ctx.repo, self.report_id, actor=self.ctx.actor)
            except PermissionDenied:
                pass
        super().reject()


class WeeklyScreen(Screen):
    """Past weekly reports: list + the report as it was sent."""
    key = "weekly"

    def __init__(self, ctx):
        super().__init__(ctx, scroll=False)
        top = QHBoxLayout()
        self.title = label("", "h1")
        top.addWidget(self.title)
        top.addStretch(1)
        self.gen_btn = button("", primary=True)
        self.gen_btn.clicked.connect(lambda: start_generate(self.ctx, self))
        top.addWidget(self.gen_btn)
        self.lay.addLayout(top)
        split = QSplitter(Qt.Horizontal)
        self.table = QTableWidget()
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.itemSelectionChanged.connect(self._show)
        split.addWidget(self.table)
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(12, 0, 0, 0)
        self.meta = label("", "muted", wrap_text=True)
        rl.addWidget(self.meta)
        self.view = QTextBrowser()
        rl.addWidget(self.view, 1)
        split.addWidget(right)
        split.setSizes([560, 800])
        self.lay.addWidget(split, 1)
        self.rows = []

    def refresh(self):
        self.title.setText(t("nav.weekly"))
        self.gen_btn.setText(t("dash.weekly"))
        self.gen_btn.setVisible(self.ctx.can(auth.WEEKLY))
        self.rows = self.ctx.repo.weekly_reports(200)
        cols = [t("col.date"), t("wk.as_of"), t("wk.trigger"), t("wk.status"), "AI", t("wk.cost"), t("col.editor")]
        tb = self.table
        tb.clear()
        tb.setColumnCount(len(cols))
        tb.setHorizontalHeaderLabels(cols)
        tb.setRowCount(len(self.rows))
        for i, r in enumerate(self.rows):
            vals = [r["created_at"][:16].replace("T", " "), r["as_of"], t(f"wk.t_{r['trigger']}"),
                    t(f"wk.s_{r['status']}"), t("yes") if r["ai_used"] else t("no"),
                    f"${r['cost_usd']:.4f}" if r["cost_usd"] is not None else "-", r["created_by"]]
            for j, v in enumerate(vals):
                it = QTableWidgetItem(str(v))
                it.setData(Qt.UserRole, r["id"])
                tb.setItem(i, j, it)
        tb.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        tb.horizontalHeader().setStretchLastSection(True)
        if self.rows:
            tb.selectRow(0)
        else:
            self.meta.setText(t("wk.none"))
            self.view.clear()

    def _show(self):
        items = self.table.selectedItems()
        if not items:
            return
        r = self.ctx.repo.get_weekly(items[0].data(Qt.UserRole))
        to = json.loads(r["recipients"] or "[]")
        self.meta.setText(f"WR-{r['id']} · {t('wk.s_' + r['status'])} · "
                          + t("wk.recipients_n", n=len(to)) + (f": {', '.join(to)}" if to else "") + " · " + _ai_line(r))
        self.view.setHtml(r["html"])

    def set_search(self, text):
        self.refresh()
