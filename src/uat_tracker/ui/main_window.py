"""Shell: left nav, top bar, screens. Sync runs in a background thread while the app is open
(approved): Excel edits are read in automatically and the open screen refreshes without a restart."""

from __future__ import annotations

import datetime as dt
import threading

from PySide6.QtCore import Qt, QThread, QTimer
from PySide6.QtWidgets import (QApplication, QButtonGroup, QComboBox, QFrame, QHBoxLayout, QLineEdit, QMainWindow,
                               QPushButton, QStackedWidget, QVBoxLayout, QWidget)

from ..mail.sender import OutlookSender
from ..services import auth
from ..services.sync import SyncService
from ..storage.sqlite_repo import SqliteRepository
from .i18n import get_lang, t
from .screens.conflicts import ConflictsScreen
from .screens.dashboard import DashboardScreen
from .screens.defects import DefectsScreen
from .screens.history import HistoryScreen
from .screens.settings import SettingsScreen
from .screens.test_progress import TestProgressScreen
from .screens.uat_plan import UatPlanScreen
from .weekly import WeeklyScreen
from .widgets import DatePicker, button, label

NAV = (("dashboard", "nav.dashboard"), ("uat", "nav.uat"), ("tp", "nav.tp"), ("defects", "nav.defects"),
       ("history", "nav.history"), ("conflicts", "nav.conflicts"), ("weekly", "nav.weekly"), ("settings", "nav.settings"))


class SyncThread(QThread):
    """Owns its own DB connection (SQLite connections are per thread)."""

    def __init__(self, cfg):
        super().__init__()
        self.cfg = cfg
        self.stop_event = threading.Event()

    def run(self):
        repo = SqliteRepository(self.cfg.db_path, shared=self.cfg.shared_db)
        try:
            SyncService(self.cfg, repo, sender=OutlookSender(self.cfg.alerts.send_account)).watch(
                self.stop_event, on_event=lambda _m: None)
        finally:
            repo.close()

    def stop(self):
        self.stop_event.set()
        self.wait(10000)


class MainWindow(QMainWindow):
    def __init__(self, ctx, start_sync: bool = True, sync_thread=None, on_closed=None):
        """sync_thread: owned by the app controller (Part 5). on_closed: called when the user closes the window."""
        super().__init__()
        self.on_closed = on_closed
        self.ctx = ctx
        ctx.navigate = self.navigate
        ctx.refresh_all = self.refresh_current
        self.setWindowTitle(f"Planning Progress - {ctx.cfg.project_name}" if ctx.cfg.project_name
                            else "Planning Progress")
        self.resize(1440, 900)
        root = QWidget()
        h = QHBoxLayout(root)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)
        # nav
        nav = QFrame()
        nav.setObjectName("nav")
        nav.setFixedWidth(240)
        nl = QVBoxLayout(nav)
        nl.setContentsMargins(12, 20, 12, 20)
        nl.setSpacing(4)
        self.app_title = label("", "h2")
        self.app_sub = label("", "sub")
        nl.addWidget(self.app_title)
        nl.addWidget(self.app_sub)
        nl.addSpacing(10)
        # Multi-project: project drop-down under the app name (+ New / Import for Admin / PMO)
        self.proj_lab = label("", "sub")
        nl.addWidget(self.proj_lab)
        self.proj_cb = QComboBox()
        self.proj_cb.setStyleSheet("QComboBox{background:#22345C;color:white;border:1px solid #3A5084;"
                                   "border-radius:8px;padding:6px 8px;}"
                                   " QComboBox QAbstractItemView{background:white;color:#14213D;"
                                   "border:1px solid #3A5084;outline:0;"
                                   "selection-background-color:#1E4FD8;selection-color:white;}"
                                   " QComboBox QAbstractItemView::item{min-height:32px;padding:4px 8px;"
                                   "color:#14213D;background:white;}"
                                   " QComboBox QAbstractItemView::item:hover,"
                                   " QComboBox QAbstractItemView::item:selected{background:#1E4FD8;color:white;}")
        self.proj_cb.activated.connect(self._pick_project)
        nl.addWidget(self.proj_cb)
        prow = QHBoxLayout()
        prow.setSpacing(4)
        self.new_proj_btn, self.imp_proj_btn = QPushButton(), QPushButton()
        for b in (self.new_proj_btn, self.imp_proj_btn):
            b.setStyleSheet("min-height:28px; font-size:9pt; padding:0 6px;")
            prow.addWidget(b)
        self.new_proj_btn.clicked.connect(lambda: ctx.app.new_project() if ctx.app else None)
        self.imp_proj_btn.clicked.connect(lambda: ctx.app.import_project() if ctx.app else None)
        nl.addLayout(prow)
        nl.addSpacing(12)
        self.nav_buttons: dict[str, QPushButton] = {}
        group = QButtonGroup(self)
        group.setExclusive(True)
        for key, _ in NAV:
            b = QPushButton()
            b.setCheckable(True)
            b.clicked.connect(lambda _=False, k=key: self.navigate(k))
            group.addButton(b)
            nl.addWidget(b)
            self.nav_buttons[key] = b
        nl.addStretch(1)
        h.addWidget(nav)
        # right side
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(0)
        top = QFrame()
        top.setObjectName("topbar")
        tl = QHBoxLayout(top)
        tl.setContentsMargins(24, 10, 24, 10)
        tl.setSpacing(14)
        proj = QVBoxLayout()
        self.proj_name = label(ctx.cfg.project_name or "", "h2")
        proj.addWidget(self.proj_name)
        self.file_lab = label("", "muted")
        proj.addWidget(self.file_lab)
        tl.addLayout(proj)
        self.search = QLineEdit()
        self.search.setClearButtonEnabled(True)
        self.search.returnPressed.connect(self._search)
        self.search.textChanged.connect(lambda s: self._search() if not s else None)
        tl.addWidget(self.search, 1)
        self.asof_lab = label("")
        tl.addWidget(self.asof_lab)
        self.asof = DatePicker(ctx.as_of.isoformat(), clearable=False)
        self.asof.setFixedWidth(200)
        self.asof.changed.connect(self._asof)
        tl.addWidget(self.asof)
        self.sync_chip = label("")
        self.mode_chip = label("")
        tl.addWidget(self.sync_chip)
        tl.addWidget(self.mode_chip)
        self.user_lab = label("")                # Part 5: who is logged in + role
        self.user_lab.setStyleSheet("color:#1E4E8C; background:#E6EFFB; border-radius:14px; padding:6px 12px;"
                                    " font-weight:600;")
        tl.addWidget(self.user_lab)
        self.switch_btn = button("")
        self.switch_btn.clicked.connect(self._switch)
        tl.addWidget(self.switch_btn)
        self.th_btn, self.en_btn = button("TH"), button("EN")
        for b, lang in ((self.th_btn, "th"), (self.en_btn, "en")):
            b.setCheckable(True)
            b.setFixedWidth(56)
            b.clicked.connect(lambda _=False, l=lang: self.set_lang(l))
            tl.addWidget(b)
        rl.addWidget(top)
        # banners: Excel file not found (Relink) / data cleared, reading Excel paused (Import from Excel)
        self.banner = QFrame()
        self.banner.setObjectName("warn")
        bl = QHBoxLayout(self.banner)
        bl.setContentsMargins(24, 8, 24, 8)
        self.banner_lab = label("", wrap_text=True)
        bl.addWidget(self.banner_lab, 1)
        self.banner_btn = button("")
        self.banner_btn.clicked.connect(self._banner_action)
        bl.addWidget(self.banner_btn)
        rl.addWidget(self.banner)
        self.stack = QStackedWidget()
        self.screens = {
            "dashboard": DashboardScreen(ctx), "uat": UatPlanScreen(ctx), "tp": TestProgressScreen(ctx),
            "defects": DefectsScreen(ctx), "history": HistoryScreen(ctx), "conflicts": ConflictsScreen(ctx),
            "weekly": WeeklyScreen(ctx), "settings": SettingsScreen(ctx),
        }
        for s in self.screens.values():
            self.stack.addWidget(s)
        rl.addWidget(self.stack, 1)
        h.addWidget(right, 1)
        self.setCentralWidget(root)
        self.current = "dashboard"
        self._audit_seen = ctx.repo.max_audit_id()
        self.sync_thread, self._owns_sync = sync_thread, False
        if sync_thread is None and start_sync:
            self.sync_thread, self._owns_sync = SyncThread(ctx.cfg), True
            self.sync_thread.start()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick)
        self.timer.start(2000)
        self.retranslate()
        self.navigate("dashboard")
        self.search.setFocus()

    # ------------------------------------------------------------------ navigation
    def navigate(self, key: str, **kw):
        self.current = key
        self.nav_buttons[key].setChecked(True)
        screen = self.screens[key]
        self.stack.setCurrentWidget(screen)
        if kw and hasattr(screen, "show_item"):
            screen.show_item(kw.get("item"), **({"filter": kw["filter"]} if "filter" in kw else {}))
        else:
            screen.refresh()
        self._status()

    def refresh_current(self):
        self._audit_seen = self.ctx.repo.max_audit_id()
        self.screens[self.current].refresh()
        self._status()

    def retranslate(self):
        self.app_title.setText(t("app"))
        self.app_sub.setText(t("app_sub"))
        self.proj_lab.setText(t("proj.label"))
        self.new_proj_btn.setText(t("proj.new"))
        self.imp_proj_btn.setText(t("proj.import"))
        self._fill_projects()
        self.search.setPlaceholderText(t("search"))
        self.search.setAccessibleName(t("search"))
        self.asof_lab.setText(t("as_of"))
        self.th_btn.setChecked(get_lang() == "th")
        self.en_btn.setChecked(get_lang() == "en")
        self.file_lab.setText(self.ctx.cfg.excel_path.name)
        self.switch_btn.setText(t("top.switch"))
        self.switch_btn.setVisible(self.ctx.switch_user is not None)
        self._status()

    def set_lang(self, lang: str):
        self.ctx.set_lang(lang)
        self.retranslate()
        self.screens[self.current].refresh()

    def _search(self):
        self.ctx.search = self.search.text().strip()
        self.screens[self.current].set_search(self.ctx.search)

    def _asof(self):
        self.ctx.as_of = dt.date.fromisoformat(self.asof.value())
        self.screens[self.current].refresh()

    # ------------------------------------------------------------------ status / auto refresh
    def _status(self):
        bs = self._banner_state()
        self.banner.setVisible(bs is not None)
        if bs is not None:
            self.banner_lab.setText(bs[1])
            self.banner_btn.setText(bs[2])
            self.banner_btn.setVisible(bs[3])
        st = SyncService(self.ctx.cfg, self.ctx.repo).status()
        if self.sync_thread is None:
            text, fg, bg = t("sync.off"), "#3D4656", "#E9EDF3"
        elif st["state"] == "synced":
            text, fg, bg = "✓ " + t("sync.synced"), "#1E6B3E", "#E3F2E8"
        elif st["state"] == "pending":
            text, fg, bg = t("sync.pending", n=st["pending"]), "#7A4E00", "#FFF1D6"
        elif st["state"] == "paused":
            text, fg, bg = t("sync.off"), "#7A4E00", "#FFF1D6"
        elif st["state"] == "blocked":
            text, fg, bg = t("sync.blocked"), "#A3260F", "#FDE7E3"
        else:
            text, fg, bg = t("sync.error"), "#A3260F", "#FDE7E3"
        self.sync_chip.setText(text)
        self.sync_chip.setToolTip(st.get("blocked") or st.get("last_error") or "")
        self.sync_chip.setStyleSheet(f"color:{fg}; background:{bg}; border-radius:14px; padding:6px 12px; font-weight:600;")
        dry = self.ctx.cfg.alerts.dry_run
        self.mode_chip.setText(t("dry_run") if dry else t("live"))
        self.mode_chip.setStyleSheet(("color:#7A4E00; background:#FFF1D6;" if dry else "color:#A3260F; background:#FDE7E3;")
                                     + " border-radius:14px; padding:6px 12px; font-weight:700;")
        a = self.ctx.actor
        role = self.ctx.role()
        self.user_lab.setText(f"{a.display_name or a.email} · {role or '-'}" if a else "")
        self.user_lab.setToolTip(a.email if a else "")
        self.nav_buttons["settings"].setVisible(self.ctx.can(auth.SETTINGS))   # QA: no Settings
        conflicts = len(self.ctx.repo.open_conflicts())
        for key, tk in NAV:
            txt = t(tk)
            if key == "conflicts" and conflicts:
                txt += f"   ({conflicts})"
            self.nav_buttons[key].setText(txt)
            self.nav_buttons[key].setAccessibleName(txt)

    # ------------------------------------------------------------------ projects
    def _fill_projects(self):
        app = self.ctx.app
        self.proj_cb.blockSignals(True)
        self.proj_cb.clear()
        mine = app.visible_projects() if app else []
        for p in mine:
            self.proj_cb.addItem(p["name"], p["id"])
        if not mine and self.ctx.cfg.project_id:
            self.proj_cb.addItem(self.ctx.cfg.project_name, self.ctx.cfg.project_id)
        self.proj_cb.setCurrentIndex(max(self.proj_cb.findData(self.ctx.cfg.project_id), 0))
        self.proj_cb.blockSignals(False)
        manager = bool(app) and self._is_manager()
        self.new_proj_btn.setVisible(manager)
        self.imp_proj_btn.setVisible(manager)

    def _is_manager(self) -> bool:
        from ..services import projects
        app = self.ctx.app
        return projects.can_manage(app.base, app.appdb, app.person.email)

    def _pick_project(self, index: int):
        pid = self.proj_cb.itemData(index)
        if pid and pid != self.ctx.cfg.project_id and self.ctx.app:
            self.ctx.app.open_project(pid)

    def _banner_state(self):
        from ..services.importer import S_PAUSED
        if self.ctx.repo.get_setting(S_PAUSED):
            return "paused", t("proj.paused"), t("proj.import_excel"), self.ctx.can(auth.SETTINGS)
        if self.ctx.cfg.project_id and not self.ctx.cfg.excel_path.exists():
            return ("missing", t("proj.missing", p=self.ctx.cfg.excel_path), t("proj.relink"),
                    bool(self.ctx.app) and self._is_manager())
        return None

    def _banner_action(self):
        st = self._banner_state()
        if st is None:
            return
        if st[0] == "paused":
            from ..services.importer import run_import
            from .dialogs import error, info
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                res = run_import(self.ctx.cfg, self.ctx.repo, editor=self.ctx.editor, force=True)
            except Exception as exc:
                QApplication.restoreOverrideCursor()
                error(self, str(exc))
                return
            QApplication.restoreOverrideCursor()
            info(self, t("uat.imported", n=sum(x.created + x.updated for x in res.sheets.values())))
            if self.ctx.app:
                self.ctx.app.restart_sync()
            self.refresh_current()
        elif st[0] == "missing" and self.ctx.app:
            from .projects import relink_flow
            app = self.ctx.app
            if relink_flow(self, app.base, app.appdb, app.person, app.appdb.get_project(self.ctx.cfg.project_id)):
                app.open_project(self.ctx.cfg.project_id)

    def _tick(self):
        if QApplication.activeModalWidget() is not None:   # never rebuild a screen under an open dialog
            return
        latest = self.ctx.repo.max_audit_id()
        if latest != self._audit_seen:                 # something changed (Excel edit read in, write done...)
            self._audit_seen = latest
            self.screens[self.current].refresh()
        self._status()

    def _switch(self):
        if self.ctx.switch_user is not None:
            self.ctx.switch_user()

    def closeEvent(self, e):
        self.timer.stop()
        if self.sync_thread is not None and self._owns_sync:
            self.sync_thread.stop()
        super().closeEvent(e)
        if self.on_closed is not None:
            self.on_closed()
