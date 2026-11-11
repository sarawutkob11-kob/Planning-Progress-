"""Multi-project UI pieces: New project form, Import Excel, Relink, Clear project data (temporary), and the
window shown when the user has no project yet."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QApplication, QDialog, QFileDialog, QFormLayout, QHBoxLayout, QLineEdit,
                               QMainWindow, QVBoxLayout, QWidget)

from ..services import projects
from ..services.importer import run_import
from .dialogs import confirm, error, info
from .i18n import t
from .widgets import button, card, label

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


def import_project_flow(parent, base, appdb, person) -> dict | None:
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
    cfg, repo = projects.open_project(base, p)
    QApplication.setOverrideCursor(Qt.WaitCursor)
    try:
        res = run_import(cfg, repo, editor=person.editor)
        summary = ", ".join(f"{s} {st.rows}" for s, st in res.sheets.items())
    except Exception as exc:                     # file readable but import failed: project kept, shown as error
        summary = str(exc)
    finally:
        QApplication.restoreOverrideCursor()
        repo.close()
    info(parent, t("proj.imported", n=p["name"], r=summary))
    return p


def relink_flow(parent, base, appdb, person, project) -> bool:
    start = Path(project["excel_path"]).parent
    path, _ = QFileDialog.getOpenFileName(parent, t("proj.relink"), str(start if start.exists() else Path.home()),
                                          "Excel (*.xlsx)")
    if not path:
        return False
    try:
        summary = projects.relink(base, appdb, person, project["id"], Path(path))
    except projects.ProjectError as exc:
        error(parent, str(exc))
        return False
    info(parent, t("proj.relinked", s=summary))
    return True


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
        from ..services.clear_project import ClearError, clear_project
        try:
            res = clear_project(self.ctx.cfg, self.ctx.repo, self.ctx.actor, typed_name=self.typed.text())
        except ClearError as exc:
            error(self, str(exc))
            return
        self.result_zip = res.backup_zip
        info(self, t("danger.done", p=res.backup_zip))
        self.accept()


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
