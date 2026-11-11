"""Desktop app entry: python -m uat_tracker ui

Part 5 v2 (approved 10-Nov-2026): no login. The first start on a PC asks First name / Last name / Email
(profile saved per Windows user); later starts open straight in. "Switch user" in the top bar opens the same
form.

Multi-project (approved 10-Nov-2026): the drop-down under the app name picks the project; every page shows only
that project. The last project is remembered (profile file). Switching closes the project (its Excel watch
stops after trying to write pending changes; anything left stays queued in that project) and opens the next.
A user without any project sees a page to create / import one (Admin or PMO) or to ask the Admin.
"""

from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import QObject, Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication

from ..services import projects
from ..services.importer import S_PAUSED
from ..services.profile import open_profile, read_profile_data
from .context import Ctx
from .i18n import set_lang
from .main_window import MainWindow, SyncThread
from .widgets import QSS

APP_NAME = "Planning Progress"


class AppController(QObject):
    def __init__(self, base, appdb, config_path: Path, *, start_sync: bool = True):
        super().__init__()
        self.base, self.appdb, self.config_path, self.start_sync = base, appdb, config_path, start_sync
        self.person = None
        self.win = None
        self.ctx = None
        self.sync_thread = None

    # ------------------------------------------------------------------ profile
    def _form(self, previous=None, parent=None):
        from .profile import ProfileDialog
        while True:
            dlg = ProfileDialog(self.base, self.appdb, previous=previous, parent=parent)
            r = dlg.exec()
            if r != 2:                                   # 2 = language switched: open again
                return dlg.person if r else None

    def start(self) -> bool:
        p = projects.current_person(self.base)
        if p is not None and projects.is_deactivated_everywhere(self.base, self.appdb, p.email):
            p = None
        if p is not None:
            stored = projects.find_person(self.base, self.appdb, p.email)
            p = stored or p
            projects.auto_join(self.base, self.appdb, p)
        self.person = p or self._form()
        if self.person is None:
            self.quit()
            return False
        self.open_best()
        return True

    def switch_user(self) -> bool:
        person = self._form(previous=self.person, parent=self.win)
        if person is None:
            return False
        if self.ctx is not None and person.email != self.person.email:      # Part 5: logged in the project
            from ..services.auth import audit_auth
            audit_auth(self.ctx.repo, editor=person.editor, field="switch_user", item_id=person.email,
                       old=self.person.editor, new=person.editor, reason="Switch user")
        self.person = person
        self.open_best(keep_current=True)
        return True

    # ------------------------------------------------------------------ projects
    def visible_projects(self) -> list[dict]:
        return projects.projects_for(self.base, self.appdb, self.person.email)

    def open_best(self, preferred: str | None = None, keep_current: bool = False) -> None:
        mine = self.visible_projects()
        if not mine:
            self._show_no_project()
            return
        want = preferred or (self.ctx.cfg.project_id if keep_current and self.ctx else None) \
            or read_profile_data(self.base.profile_path).get("last_project")
        p = next((x for x in mine if x["id"] == want), mine[0])
        self.open_project(p["id"])

    def open_project(self, project_id: str) -> None:
        project = self.appdb.get_project(project_id)
        self._close_current()
        base_cfg = projects.project_config(self.base, project)
        cfg, repo = projects.open_project(self.base, project)
        actor = open_profile(cfg, repo)
        if actor is None:                                # not (or no longer) a member: back to the list
            repo.close()
            self.open_best()
            return
        projects.remember_project(self.base, project_id)
        ctx = Ctx(base_cfg, repo)
        ctx.actor = actor
        ctx.config_path = self.config_path
        ctx.switch_user = self.switch_user
        ctx.app = self
        self.ctx = ctx
        if self.start_sync and not projects.file_missing(project) and not repo.get_setting(S_PAUSED):
            self.sync_thread = SyncThread(ctx.cfg)
            self.sync_thread.start()
        self.win = MainWindow(ctx, start_sync=False, sync_thread=self.sync_thread, on_closed=self.quit)
        self.win.showMaximized()

    def restart_sync(self) -> None:
        """After Import from Excel / Relink: (re)start watching this project's file."""
        if self.sync_thread is not None:
            self.sync_thread.stop()
            self.sync_thread = None
        project = self.appdb.get_project(self.ctx.cfg.project_id)
        if self.start_sync and not projects.file_missing(project) and not self.ctx.repo.get_setting(S_PAUSED):
            self.sync_thread = SyncThread(self.ctx.cfg)
            self.sync_thread.start()
        if self.win is not None:
            self.win.sync_thread = self.sync_thread

    def _close_current(self) -> None:
        if self.sync_thread is not None:
            self.sync_thread.stop()
            self.sync_thread = None
            if self.ctx is not None:                     # approved: try to write pending changes first
                from ..services.sync import SyncService
                QApplication.setOverrideCursor(Qt.WaitCursor)
                try:
                    if self.ctx.repo.all_pending():
                        SyncService(self.ctx.cfg, self.ctx.repo).flush(force=True)
                except Exception:
                    pass                                 # stays queued; retried when the project is opened again
                finally:
                    QApplication.restoreOverrideCursor()
        if self.win is not None:
            self.win.on_closed = None
            self.win.close()
            self.win = None
        if self.ctx is not None:
            self.ctx.repo.close()
            self.ctx = None

    def _show_no_project(self) -> None:
        from .projects import NoProjectWindow
        self._close_current()
        self.win = NoProjectWindow(self)
        self.win.on_closed = self.quit
        self.win.show()

    def new_project(self) -> None:
        from .projects import NewProjectDialog
        dlg = NewProjectDialog(self.base, self.appdb, self.person, parent=self.win)
        if dlg.exec() and dlg.project:
            self.open_project(dlg.project["id"])

    def import_project(self) -> None:
        from .projects import import_project_flow
        p = import_project_flow(self.win, self.base, self.appdb, self.person)
        if p:
            self.open_project(p["id"])

    def quit(self) -> None:
        if self.sync_thread is not None:
            self.sync_thread.stop()
            self.sync_thread = None
        QApplication.instance().quit()


def run(base, config_path: Path, start_sync: bool = True) -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setFont(QFont("IBM Plex Sans Thai", 10) if "IBM Plex Sans Thai" in _families() else QFont("Tahoma", 10))
    app.setStyleSheet(QSS)
    appdb = projects.open_app_db(base)
    set_lang(appdb.get_setting("ui.lang") or "en")
    ctl = AppController(base, appdb, config_path, start_sync=start_sync)
    if not ctl.start():
        return 0
    return app.exec()


def _families() -> set[str]:
    from PySide6.QtGui import QFontDatabase
    return set(QFontDatabase.families())


def create_shortcut(config_path: Path) -> Path:
    """Desktop shortcut that starts the app without a console window (approved). The old
    'UAT Tracker' shortcut is removed (renamed app)."""
    import win32com.client
    shell = win32com.client.Dispatch("WScript.Shell")
    desktop = Path(shell.SpecialFolders("Desktop"))
    lnk = desktop / f"{APP_NAME}.lnk"
    exe = Path(sys.executable)
    pyw = exe.with_name("pythonw.exe")
    sc = shell.CreateShortCut(str(lnk))
    sc.TargetPath = str(pyw if pyw.exists() else exe)
    sc.Arguments = f'-m uat_tracker --config "{config_path.resolve()}" ui'
    sc.WorkingDirectory = str(config_path.resolve().parent)
    sc.Description = APP_NAME
    sc.Save()
    old = desktop / "UAT Tracker.lnk"
    if old.exists():
        old.unlink()
    return lnk
