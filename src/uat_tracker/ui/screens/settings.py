"""Screen 8: Settings. Values are stored in the app DB and override config.toml (CLI + scheduled job too)."""

from __future__ import annotations

import logging
from PySide6.QtCore import QTime
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QDoubleSpinBox, QFileDialog, QFormLayout, QHBoxLayout,
                               QHeaderView, QLineEdit, QTableWidget, QTableWidgetItem, QTimeEdit)

from ...alerts.templates import SUBJECTS
from ...config import email_time
from ...parsing.dates import fmt
from ...services import auth
from ...services.auth import PermissionDenied
from ...services import clear_project, projects, weekly_ai
from ...services.settings import fetch_holidays, install_schedule, save_ai_settings, save_settings
from ...services.sync import SyncService
from ..dialogs import confirm, error, info
from ..i18n import t
from ..widgets import DatePicker, button, card, label
from .base import Screen, clear_layout

log = logging.getLogger(__name__)


def _table(headers, rows, height=220) -> QTableWidget:
    tb = QTableWidget(len(rows), len(headers))
    tb.setHorizontalHeaderLabels(headers)
    tb.verticalHeader().setVisible(False)
    tb.setEditTriggers(QAbstractItemView.NoEditTriggers)
    tb.setSelectionBehavior(QAbstractItemView.SelectRows)
    tb.setSelectionMode(QAbstractItemView.SingleSelection)
    tb.setAlternatingRowColors(True)
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            tb.setItem(i, j, QTableWidgetItem("" if v is None else str(v)))
    tb.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    tb.horizontalHeader().setStretchLastSection(True)
    tb.setMinimumHeight(height)
    return tb


class SettingsScreen(Screen):
    key = "settings"

    def refresh(self):
        ctx, repo, cfg = self.ctx, self.ctx.repo, self.ctx.cfg
        clear_layout(self.lay)
        head = QHBoxLayout()
        head.addWidget(label(t("set.title"), "h1"))
        head.addStretch(1)
        save = button(t("save"), primary=True)
        save.clicked.connect(self._save)
        head.addWidget(save)
        self.lay.addLayout(head)
        if clear_project.live_warning(cfg):              # TEMPORARY clear-project feature (see README)
            self.lay.addWidget(self._live_warning())

        app = ctx.app
        self.manager = bool(app) and projects.can_manage(app.base, app.appdb, app.person.email)
        if app and cfg.project_id:
            self.lay.addWidget(self._project_card())
        c, cl = card()
        cl.addWidget(label(t("proj.this") + (f": {cfg.project_name}" if cfg.project_name else ""), "h2"))
        form = QFormLayout()
        form.setSpacing(10)
        self.excel = None
        if not cfg.project_id:                               # single-file mode (no projects): old field
            self.excel = QLineEdit(str(cfg.excel_path))
            browse = button(t("set.browse"))
            browse.clicked.connect(self._browse)
            row = QHBoxLayout()
            row.addWidget(self.excel, 1)
            row.addWidget(browse)
            form.addRow(t("set.excel"), row)
            form.addRow("", label(t("set.restart"), "muted"))
        self.fail = QDoubleSpinBox()
        self.fail.setRange(0, 50)
        self.fail.setDecimals(1)
        self.fail.setSuffix(" %")
        self.fail.setValue(cfg.alerts.fail_threshold * 100)
        form.addRow(t("set.fail"), self.fail)
        self.time = QTimeEdit(QTime.fromString(email_time(repo), "HH:mm"))
        self.time.setDisplayFormat("HH:mm")
        sched = button(t("set.apply_sched"))
        sched.clicked.connect(self._schedule)
        row = QHBoxLayout()
        row.addWidget(self.time)
        row.addWidget(sched)
        row.addStretch(1)
        form.addRow(t("set.email_time"), row)
        rt = QLineEdit(t("set.report_note"))
        rt.setReadOnly(True)
        form.addRow(t("set.report_time"), rt)
        self.dry = QCheckBox(t("set.dry"))
        self.dry.setChecked(cfg.alerts.dry_run)
        form.addRow("", self.dry)
        cl.addLayout(form)
        self.lay.addWidget(c)

        c, cl = card()                                       # shared by all projects (Admin / PMO only)
        cl.addWidget(label(t("proj.shared"), "h2"))
        cl.addWidget(label(t("proj.shared_note"), "muted"))
        form = QFormLayout()
        form.setSpacing(10)
        self.pmo = QLineEdit(cfg.alerts.pmo_email)
        form.addRow(t("set.pmo"), self.pmo)
        self.api_key = QLineEdit()                    # Part 6: kept in Windows Credential Manager only
        self.api_key.setEchoMode(QLineEdit.Password)
        self.has_key = bool(weekly_ai.KEYS.get())
        self.api_key.setPlaceholderText(t("set.api_saved") if self.has_key else t("set.api_note"))
        form.addRow(t("set.api_key"), self.api_key)
        self.ai_model = QLineEdit(weekly_ai.model_of(repo, cfg))
        form.addRow(t("set.ai_model"), self.ai_model)
        form.addRow("", label(t("set.ai_policy"), "muted", wrap_text=True))
        cl.addLayout(form)
        for w in (self.pmo, self.api_key, self.ai_model):
            w.setEnabled(self.manager)
        self.lay.addWidget(c)

        if ctx.can(auth.USERS):                  # Part 5: Admin only
            from ..users import users_card
            self.lay.addWidget(users_card(self))

        c, cl = card()
        cl.addWidget(label(t("set.holidays"), "h2"))
        fetches = repo.holiday_fetches(1)
        last = f"{fetches[0]['ts'][:16]} {fetches[0]['result']}" if fetches else t("none")
        cl.addWidget(label(t("set.hol_status", y=", ".join(map(str, sorted(repo.holiday_years()))) or t("none"),
                             f=last), wrap_text=True))
        fb = button(t("set.fetch"))
        fb.clicked.connect(self._fetch)
        fb.setEnabled(self.manager)
        cl.addWidget(fb)
        self.lay.addWidget(c)

        c, cl = card()
        cl.addWidget(label(t("set.backups"), "h2"))
        cl.addWidget(label(t("set.backup_dir", p=cfg.backup_dir), "muted", wrap_text=True))
        self.backups = repo.backups(20)
        self.btable = _table(["#", t("col.date"), "KB", t("col.reason")],
                             [(b["id"], b["ts"][:16].replace("T", " "), b["size_bytes"] // 1024, b["reason"])
                              for b in self.backups], 180)
        cl.addWidget(self.btable)
        rb = button(t("set.restore"))
        rb.clicked.connect(self._restore)
        cl.addWidget(rb)
        self.lay.addWidget(c)

        c, cl = card()
        cl.addWidget(label(t("set.email_log"), "h2"))
        log = repo.email_log(limit=200)
        cl.addWidget(_table([t("col.date"), "Rule", t("col.item"), "To", "Mode", "Result", "Subject"],
                            [(r["ts"][:16].replace("T", " "), r["rule"], r["item_id"] or "", r["to_addr"] or "",
                              r["mode"], r["result"], r["subject"]) for r in log]))
        self.lay.addWidget(c)

        c, cl = card()
        cl.addWidget(label(t("set.templates"), "h2"))
        cl.addWidget(label(t("set.templates_note"), "muted"))
        cl.addWidget(_table(["Rule", "Variant", "Subject"],
                            [(k[0], k[1] or "", v.template) for k, v in SUBJECTS.items()], 260))
        self.lay.addWidget(c)
        if clear_project.enabled(cfg) and clear_project.allowed(repo, ctx.actor):
            self.lay.addWidget(self._danger_zone())
        self.lay.addStretch(1)

    # ------------------------------------------------------------------ multi-project
    def _project_card(self):
        app, cfg = self.ctx.app, self.ctx.cfg
        project = app.appdb.get_project(cfg.project_id)
        c, cl = card()
        cl.addWidget(label(t("proj.card") + f": {project['name']}", "h2"))
        form = QFormLayout()
        form.setSpacing(10)
        row = QHBoxLayout()
        row.addWidget(label(project["excel_path"], wrap_text=True), 1)
        rl = button(t("proj.relink"))
        rl.clicked.connect(self._relink)
        rl.setVisible(self.manager)
        row.addWidget(rl)
        form.addRow(t("proj.file"), row)
        self.win_from = DatePicker(cfg.window_start.isoformat(), clearable=False)
        self.win_to = DatePicker(cfg.window_end.isoformat(), clearable=False)
        row = QHBoxLayout()
        row.addWidget(self.win_from)
        row.addWidget(label("-"))
        row.addWidget(self.win_to)
        row.addStretch(1)
        form.addRow(t("proj.window"), row)
        form.addRow("", label(t("proj.window_note"), "muted"))
        cl.addLayout(form)
        if self.manager:
            ab = button(t("proj.archive"))
            ab.clicked.connect(self._archive)
            cl.addWidget(ab)
            arch = [p for p in app.appdb.projects(include_archived=True) if p["archived"]]
            if arch:
                cl.addWidget(label(t("proj.archived"), "h2"))
                self.arch_rows = arch
                self.arch_table = _table([t("proj.name"), t("proj.file")],
                                         [(p["name"], p["excel_path"]) for p in arch], 120)
                cl.addWidget(self.arch_table)
                ub = button(t("proj.unarchive"))
                ub.clicked.connect(self._unarchive)
                cl.addWidget(ub)
        return c

    def _relink(self):
        from ..projects import relink_flow
        app = self.ctx.app
        pid = self.ctx.cfg.project_id
        relink_flow(self, app.base, app.appdb, app.person, app.appdb.get_project(pid),
                    on_done=lambda: app.open_project(pid))

    def _archive(self):
        app, cfg = self.ctx.app, self.ctx.cfg
        if not confirm(self, t("proj.archive_q", n=cfg.project_name)):
            return
        try:
            projects.set_archived(app.base, app.appdb, app.person, cfg.project_id, True)
        except projects.ProjectError as exc:
            error(self, str(exc))
            return
        app.open_best()

    def _unarchive(self):
        row = self.arch_table.currentRow()
        if row < 0:
            return
        app = self.ctx.app
        try:
            projects.set_archived(app.base, app.appdb, app.person, self.arch_rows[row]["id"], False)
        except projects.ProjectError as exc:
            error(self, str(exc))
            return
        info(self, t("done"))
        if app.win is not None:
            app.win.retranslate()
        self.ctx.refresh_all()

    def _live_warning(self):
        lab = label(t("danger.live_warning"), wrap_text=True)
        lab.setObjectName("liveClearWarning")
        lab.setStyleSheet("background:#C62828;color:white;font-weight:bold;padding:10px;border-radius:6px;")
        return lab

    def _danger_zone(self):
        from ..projects import DANGER_QSS, ClearProjectDialog
        c, cl = card()
        c.setStyleSheet("QFrame#card{border:2px solid #C62828;}")
        cl.addWidget(label(t("danger.title"), "h2"))
        cl.addWidget(label(t("danger.note"), "muted", wrap_text=True))
        b = button(t("danger.button"))
        b.setStyleSheet(DANGER_QSS)

        def go():
            if ClearProjectDialog(self.ctx, self).exec():
                if self.ctx.app:
                    self.ctx.app.restart_sync()          # stops the watch: reading Excel is paused
                self.ctx.refresh_all()
        b.clicked.connect(go)
        row = QHBoxLayout()
        row.addWidget(b)
        row.addStretch(1)
        cl.addLayout(row)
        return c

    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(self, t("set.excel"), str(self.ctx.cfg.excel_path.parent), "Excel (*.xlsx)")
        if path:
            self.excel.setText(path)

    def _save(self):
        repo, cfg, app = self.ctx.repo, self.ctx.cfg, self.ctx.app
        if cfg.alerts.dry_run and not self.dry.isChecked() and not confirm(self, t("set.dry_off_q")):
            self.dry.setChecked(True)
            return
        excel = self.excel.text().strip() if self.excel is not None else ""
        try:
            if app and cfg.project_id:
                import datetime as dt
                ws, we = dt.date.fromisoformat(self.win_from.value()), dt.date.fromisoformat(self.win_to.value())
                if (ws, we) != (cfg.window_start, cfg.window_end):
                    projects.set_window(app.base, app.appdb, app.person, cfg.project_id, ws, we)
                    self.ctx.base_cfg = projects.project_config(app.base, app.appdb.get_project(cfg.project_id))
            save_settings(repo, self.ctx.actor, fail_pct=self.fail.value(), dry_run=self.dry.isChecked(),
                          pmo_email=self.pmo.text() if self.manager else None,
                          email_time=self.time.time().toString("HH:mm"),
                          excel_path=excel if excel and excel != str(cfg.excel_path) else None)
            key = self.api_key.text().strip()
            typed = self.ai_model.text().strip()
            if self.manager and (key or typed != weekly_ai.model_of(repo, cfg)):
                # unchanged model: keep what is stored (blank = follow config [weekly_ai] model)
                model = typed if typed != weekly_ai.model_of(repo, cfg) else (repo.get_setting(weekly_ai.S_MODEL) or "")
                save_ai_settings(repo, self.ctx.actor, model=model, api_key=key or None)
        except (PermissionDenied, projects.ProjectError) as exc:
            error(self, str(exc))
            return
        self.api_key.clear()
        self.ctx.reload_cfg()
        info(self, t("set.saved"))
        self.ctx.refresh_all()

    def _schedule(self):
        try:
            daily, weekly = install_schedule(self.ctx.repo, self.ctx.actor, self.ctx.config_path,
                                             self.time.time().toString("HH:mm"), self.ctx.cfg.project_id,
                                             self.ctx.cfg.project_name or None)
        except (ValueError, PermissionDenied) as exc:
            error(self, str(exc))
            return
        ok = daily.returncode == 0 and weekly.returncode == 0
        msgs = [(p.stdout or p.stderr or "").strip() for p in (daily, weekly)]
        (info if ok else error)(self, "\n".join(m for m in msgs if m))

    def _fetch(self):
        try:
            msgs = fetch_holidays(self.ctx.cfg, self.ctx.repo, self.ctx.actor)
        except PermissionDenied as exc:
            error(self, str(exc))
            return
        info(self, "\n".join(msgs))
        self.refresh()

    def _restore(self):
        row = self.btable.currentRow()
        if row < 0:
            return
        b = self.backups[row]
        if not confirm(self, t("set.restore_q", id=b["id"])):
            return
        from ..worker import project_repo, run_in_worker
        cfg, actor = self.ctx.cfg, self.ctx.actor

        def job():
            with project_repo(cfg) as repo:
                return SyncService(cfg, repo).restore(b["id"], actor)

        def done(_):
            info(self, t("done"))
            self.ctx.refresh_all()
        return run_in_worker(self, t("work.restoring"), job, done, name="restore-backup")
