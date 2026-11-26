"""Multi-project UI pieces: New project form, Import Excel (choose a file or paste a SharePoint link), Relink,
Clear project data (temporary), Delete an archived project for good (backup zip first, the Excel file is never
touched), and the window shown when the user has no project yet."""

from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtWidgets import (QDialog, QFileDialog, QFormLayout, QHBoxLayout, QInputDialog, QLineEdit,
                               QMainWindow, QMessageBox, QPushButton, QVBoxLayout, QWidget)

from ..services import projects
from ..services.importer import run_import
from .dialogs import confirm, error, info
from .i18n import t
from .widgets import button, card, label

log = logging.getLogger(__name__)

DANGER_QSS = ("QPushButton{background:#C62828;color:white;border:none;border-radius:8px;padding:8px 16px;"
              "font-weight:700;} QPushButton:disabled{background:#E8B4B4;}")


class NewProjectDialog(QDialog):
    """Project Name only; folder asked every time; final file name shown before Create (approved D)."""

    def __init__(self, base, appdb, person, parent=None):
        super().__init__(parent)
        self.base, self.appdb, self.person, self.project, self.folder = base, appdb, person, None, None
        self.setWindowTitle(t("proj.new_title"))
        self.setMinimumWidth(560)
        outer = QVBoxLayout(self)
        form = QFormLayout()
        self.name = QLineEdit()
        self.name.setAccessibleName(t("proj.name"))
        self.name.textChanged.connect(self._preview)
        form.addRow(t("proj.name"), self.name)
        self.folder_lab = label("-", "muted", wrap_text=True)
        pick = button(t("proj.pick_folder"))
        pick.clicked.connect(self._pick)
        row = QHBoxLayout()
        row.addWidget(self.folder_lab, 1)
        row.addWidget(pick)
        form.addRow(t("proj.folder"), row)
        outer.addLayout(form)
        self.file_lab = label("", wrap_text=True)
        self.file_lab.setStyleSheet("font-weight:600;")
        outer.addWidget(self.file_lab)
        self.problem = label("", wrap_text=True)
        self.problem.setStyleSheet("color:#A3260F; font-weight:600;")
        outer.addWidget(self.problem)
        btns = QHBoxLayout()
        btns.addStretch(1)
        c = button(t("cancel"))
        c.clicked.connect(self.reject)
        self.ok = button(t("proj.create"), primary=True)
        self.ok.clicked.connect(self._create)
        btns.addWidget(c)
        btns.addWidget(self.ok)
        outer.addLayout(btns)
        self._preview()

    def _pick(self):
        d = QFileDialog.getExistingDirectory(self, t("proj.folder"), str(self.folder or Path.home()))
        if d:
            self.folder = Path(d)
            self.folder_lab.setText(str(self.folder))
            self._preview()

    def _preview(self):
        self.problem.setText("")
        self.file_lab.setText("")
        self.ok.setEnabled(False)
        name = self.name.text().strip()
        if not name:
            return
        try:
            fname = projects.file_name_for(name)
        except projects.ProjectError as exc:
            self.problem.setText(str(exc))
            return
        self.file_lab.setText(t("proj.file_will_be", f=fname))
        if self.folder is None:
            self.problem.setText(t("proj.need_folder"))
            return
        why = projects.duplicate_reason(self.appdb, self.folder, fname)
        if why:
            self.problem.setText(why)
            return
        self.ok.setEnabled(True)

    def _create(self):
        try:
            self.project = projects.create_project(self.base, self.appdb, self.person, name=self.name.text(),
                                                   folder=self.folder)
        except (projects.ProjectError, OSError) as exc:
            self.problem.setText(str(exc))
            return
        info(self, t("proj.created", n=self.project["name"], p=self.project["excel_path"]))
        self.accept()


def _import_choice_box(parent) -> tuple[QMessageBox, QPushButton, QPushButton]:
    """Box text proj.import_how; buttons proj.import_file, proj.import_link (+ Cancel). Not shown."""
    box = QMessageBox(QMessageBox.Question, t("app"), t("proj.import_how"), parent=parent)
    by_file = box.addButton(t("proj.import_file"), QMessageBox.AcceptRole)
    by_link = box.addButton(t("proj.import_link"), QMessageBox.AcceptRole)
    box.addButton(QMessageBox.Cancel)
    return box, by_file, by_link


def import_choice(parent) -> str | None:
    """'file' / 'link' / None (Cancel / closed)."""
    box, by_file, by_link = _import_choice_box(parent)
    box.exec()
    clicked = box.clickedButton()
    if clicked is by_file:
        return "file"
    if clicked is by_link:
        return "link"
    return None


def link_error_text(exc: projects.ExcelLinkError) -> str:
    return t(f"proj.xl_{exc.code}", p=exc.path)


def import_project_flow(parent, base, appdb, person, on_done=None) -> dict | None:
    """The project is created here; reading its workbook runs in a worker, then on_done(project)."""
    path, _ = QFileDialog.getOpenFileName(parent, t("proj.import_title"), str(Path.home()), "Excel (*.xlsx)")
    if not path:
        return None
    problems = projects.validate_workbook(Path(path), base.header_overrides)
    if problems:
        error(parent, "\n".join(problems))
        return None
    if not confirm(parent, t("proj.import_q", f=Path(path).name, n=projects.name_from_file(path))):
        return None
    try:
        p = projects.import_project(base, appdb, person, path=Path(path))
    except projects.ProjectError as exc:
        error(parent, str(exc))
        return None
    _read_new_project(parent, base, person, p, on_done)
    return p


def import_link_flow(parent, base, appdb, person, on_done=None) -> dict | None:
    """Import from a pasted SharePoint .xlsx link (file synced by OneDrive); confirm shows link + local path.
    Reading the new project's workbook runs in a worker, then on_done(project)."""
    link, ok = QInputDialog.getText(parent, t("proj.xl_link_title"), t("proj.xl_link_prompt"))
    if not ok or not link.strip():
        return None
    try:
        url, path = projects.resolve_excel_link(link)
        problems = projects.validate_link_file(base, path)
    except projects.ExcelLinkError as exc:
        error(parent, link_error_text(exc))
        return None
    if problems:
        error(parent, "\n".join(problems))
        return None
    if not confirm(parent, t("proj.import_link_q", u=url, p=path, n=projects.name_from_file(path))):
        return None
    try:
        p = projects.import_project_link(base, appdb, person, link=url)
    except projects.ExcelLinkError as exc:
        error(parent, link_error_text(exc))
        return None
    except projects.ProjectError as exc:
        error(parent, str(exc))
        return None
    _read_new_project(parent, base, person, p, on_done)
    return p


def _read_new_project(parent, base, person, p: dict, on_done=None) -> None:
    """First read of a just-created project in a worker, then info(proj.imported) + on_done(project)."""
    from .worker import run_in_worker

    def job():                                   # own connections: runs in a worker thread
        cfg, repo = projects.open_project(base, p)
        try:
            res = run_import(cfg, repo, editor=person.editor)
            return ", ".join(f"{s} {st.rows}" for s, st in res.sheets.items())
        except Exception as exc:                 # file readable but import failed: project kept, shown as error
            log.error("Import of the new project failed", exc_info=True)
            return str(exc)
        finally:
            repo.close()

    def done(summary):
        info(parent, t("proj.imported", n=p["name"], r=summary))
        if on_done:
            on_done(p)
    run_in_worker(parent, t("work.reading"), job, done, name="import-project")


def relink_flow(parent, base, appdb, person, project, on_done=None):
    """Pick the moved / renamed file; the check (reads the workbook) runs in a worker. on_done() after success."""
    start = Path(project["excel_path"]).parent
    path, _ = QFileDialog.getOpenFileName(parent, t("proj.relink"), str(start if start.exists() else Path.home()),
                                          "Excel (*.xlsx)")
    if not path:
        return None
    from .worker import run_in_worker

    def job():                                   # app.db connection of this thread
        db = projects.open_app_db(base)
        try:
            return projects.relink(base, db, person, project["id"], Path(path))
        finally:
            db.close()

    def done(summary):
        info(parent, t("proj.relinked", s=summary))
        if on_done:
            on_done()
    return run_in_worker(parent, t("work.relinking"), job, done, name="relink")


def relink_link_flow(parent, base, appdb, person, project, on_done=None):
    """Relink to a pasted SharePoint .xlsx link (file synced by OneDrive); confirm shows link + local path.
    The check (reads the workbook) runs in a worker. on_done() after success."""
    link, ok = QInputDialog.getText(parent, t("proj.xl_link_title"), t("proj.xl_link_prompt"))
    if not ok or not link.strip():
        return None
    try:
        url, path = projects.resolve_excel_link(link)      # no file content read: UI thread is fine
    except projects.ExcelLinkError as exc:
        error(parent, link_error_text(exc))
        return None
    if not confirm(parent, t("proj.relink_link_q", u=url, p=path, n=project["name"])):
        return None
    from .worker import friendly, run_in_worker

    def job():                                   # app.db connection of this thread
        db = projects.open_app_db(base)
        try:
            return projects.relink_link(base, db, person, project["id"], url)
        finally:
            db.close()

    def done(summary):
        info(parent, t("proj.relinked", s=summary))
        if on_done:
            on_done()

    def failed(exc):
        error(parent, link_error_text(exc) if isinstance(exc, projects.ExcelLinkError) else friendly(exc))
    return run_in_worker(parent, t("work.relinking"), job, done, failed, name="relink-link")


class ClearProjectDialog(QDialog):
    """TEMPORARY (section I): what is cleared / kept; the project name must be typed to enable the button."""

    def __init__(self, ctx, parent=None):
        super().__init__(parent)
        self.ctx, self.result_zip = ctx, None
        name = ctx.cfg.project_name
        self.setWindowTitle(t("danger.dialog", n=name))
        self.setMinimumWidth(620)
        outer = QVBoxLayout(self)
        outer.addWidget(label(t("danger.dialog", n=name), "h2"))
        outer.addWidget(label(t("danger.cleared"), "h2"))
        outer.addWidget(label(t("danger.cleared_list"), wrap_text=True))
        outer.addWidget(label(t("danger.kept"), "h2"))
        outer.addWidget(label(t("danger.kept_list"), wrap_text=True))
        outer.addWidget(label(t("danger.type") + f":  {name}", "muted"))
        self.typed = QLineEdit()
        self.typed.setAccessibleName(t("danger.type"))
        self.typed.textChanged.connect(lambda s: self.go.setEnabled(s.strip() == name.strip()))
        outer.addWidget(self.typed)
        row = QHBoxLayout()
        row.addStretch(1)
        c = button(t("cancel"))
        c.clicked.connect(self.reject)
        self.go = button(t("danger.go"))
        self.go.setStyleSheet(DANGER_QSS)
        self.go.setEnabled(False)
        self.go.clicked.connect(self._go)
        row.addWidget(c)
        row.addWidget(self.go)
        outer.addLayout(row)

    def _go(self):
        from ..services.clear_project import clear_project
        from .worker import project_repo, run_in_worker
        cfg, actor, typed = self.ctx.cfg, self.ctx.actor, self.typed.text()
        self.go.setEnabled(False)

        def job():
            with project_repo(cfg) as repo:
                return clear_project(cfg, repo, actor, typed_name=typed)

        def done(res):
            self.result_zip = res.backup_zip
            info(self, t("danger.done", p=res.backup_zip))
            self.accept()

        def failed(exc):
            self.go.setEnabled(True)
            from .worker import friendly
            error(self, friendly(exc))
        return run_in_worker(self, t("work.clearing"), job, done, failed, name="clear-project")


class DeleteProjectDialog(QDialog):
    """Delete an ARCHIVED project for good: what is deleted / not touched (the Excel file) / where the backup
    goes; the project name must be typed to enable the red button. Runs in a worker with its own app.db."""

    def __init__(self, base, person, project: dict, parent=None) -> None:
        super().__init__(parent)
        from ..services.delete_project import what_is_deleted
        self.base, self.person, self.project, self.result = base, person, project, None
        name = project["name"]
        paths = what_is_deleted(base, project)
        self.setWindowTitle(t("proj.delete_title", n=name))
        self.setMinimumWidth(620)
        outer = QVBoxLayout(self)
        outer.addWidget(label(t("proj.delete_title", n=name), "h2"))
        outer.addWidget(label(t("proj.delete_will"), "h2"))
        outer.addWidget(label(t("proj.delete_list", p=paths["folder"]), wrap_text=True))
        outer.addWidget(label(t("proj.delete_kept"), "h2"))
        outer.addWidget(label(t("proj.delete_kept_list", p=paths["excel"]), wrap_text=True))
        outer.addWidget(label(t("proj.delete_backup", p=paths["backup_dir"]), wrap_text=True))
        outer.addWidget(label(t("proj.delete_type") + f":  {name}", "muted"))
        self.typed = QLineEdit()
        self.typed.setAccessibleName(t("proj.delete_type"))
        self.typed.textChanged.connect(lambda s: self.go.setEnabled(s.strip() == name.strip()))
        outer.addWidget(self.typed)
        row = QHBoxLayout()
        row.addStretch(1)
        c = button(t("cancel"))
        c.clicked.connect(self.reject)
        self.go = button(t("proj.delete_go"))
        self.go.setStyleSheet(DANGER_QSS)
        self.go.setEnabled(False)
        self.go.clicked.connect(self._go)
        row.addWidget(c)
        row.addWidget(self.go)
        outer.addLayout(row)

    def _go(self):
        from ..services.delete_project import delete_project
        from .worker import friendly, run_in_worker
        base, person, pid, typed = self.base, self.person, self.project["id"], self.typed.text()
        self.go.setEnabled(False)

        def job():                                   # app.db connection of this thread
            db = projects.open_app_db(base)
            try:
                return delete_project(base, db, person, pid, typed_name=typed)
            finally:
                db.close()

        def done(res):
            self.result = res
            text = t("proj.delete_done", p=res.backup_zip)
            if res.leftover:
                text += "\n" + t("proj.delete_left", p=res.leftover)
            info(self, text)
            self.accept()

        def failed(exc):
            self.go.setEnabled(True)
            error(self, friendly(exc))
        return run_in_worker(self, t("work.deleting"), job, done, failed, name="delete-project")


class NoProjectWindow(QMainWindow):
    """Shown when the user has no project: managers create / import; others are told to ask the Admin."""

    def __init__(self, controller):
        super().__init__()
        self.ctl = controller
        self.setWindowTitle(t("app"))
        self.resize(720, 420)
        root = QWidget()
        root.setObjectName("page")
        lay = QVBoxLayout(root)
        lay.setContentsMargins(40, 40, 40, 40)
        c, cl = card()
        cl.addWidget(label(t("app"), "h1"))
        cl.addWidget(label(t("proj.none_title"), "h2"))
        person = controller.person
        manager = projects.can_manage(controller.base, controller.appdb, person.email)
        cl.addWidget(label(t("proj.none_manager") if manager else t("proj.none_user", e=person.email),
                           wrap_text=True))
        row = QHBoxLayout()
        if manager:
            n = button(t("proj.new"), primary=True)
            n.clicked.connect(controller.new_project)
            i = button(t("proj.import"))
            i.clicked.connect(controller.import_project)
            row.addWidget(n)
            row.addWidget(i)
        sw = button(t("top.switch"))
        sw.clicked.connect(controller.switch_user)
        row.addWidget(sw)
        row.addStretch(1)
        cl.addLayout(row)
        lay.addWidget(c)
        lay.addStretch(1)
        self.setCentralWidget(root)

    def closeEvent(self, e):
        super().closeEvent(e)
        if getattr(self, "on_closed", None):
            self.on_closed()
