"""Dialogs. Every save goes through the services (edit_item, add_task, add_defect, send_reminder):
the rules live there; the dialogs only show live warnings so the PMO sees why Save is blocked."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (QComboBox, QDialog, QFormLayout, QFrame, QHBoxLayout, QLabel, QLineEdit,
                               QMessageBox, QPlainTextEdit, QScrollArea, QSpinBox, QVBoxLayout, QWidget)

from ..excel.sheets import DEFECT_LIST, TEST_PROGRESS, UAT_PLAN
from ..mail.sender import OutlookSender, SendError
from ..parsing.dates import fmt
from ..parsing.status import DEFECT_STATUSES, TASK_STATUSES
from ..services.edits import ROOT_CAUSES, EditError, edit_item
from ..services.items import add_application, add_defect, add_task
from ..services import auth
from ..services.auth import PermissionDenied
from ..services.reminders import ReminderError, lookup_contact, remember_contact, send_reminder
from .i18n import t
from .widgets import DatePicker, button, label


def error(parent, msg: str):
    QMessageBox.warning(parent, t("error"), msg)


def info(parent, msg: str):
    QMessageBox.information(parent, t("done"), msg)


def confirm(parent, msg: str) -> bool:
    return QMessageBox.question(parent, t("app"), msg, QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.No) == QMessageBox.Yes


class _Base(QDialog):
    def __init__(self, ctx, title: str, parent=None, scrollable: bool = False):
        super().__init__(parent)
        self.ctx = ctx
        self.setWindowTitle(title)
        self.setMinimumWidth(560)
        self.outer = QVBoxLayout(self)
        self.outer.setSpacing(12)
        self.form = QFormLayout()
        self.form.setSpacing(10)
        self._scroll = None
        if scrollable:
            # long forms: only the fields scroll; warning + Save/Cancel stay pinned at the bottom
            self._inner = QWidget()
            self._inner.setObjectName("dlgBody")
            self._inner.setLayout(self.form)
            self.form.setContentsMargins(0, 0, 8, 0)      # room for the scrollbar
            self._scroll = QScrollArea()
            self._scroll.setWidgetResizable(True)
            self._scroll.setFrameShape(QFrame.NoFrame)
            self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self._scroll.setStyleSheet("QScrollArea, QWidget#dlgBody { background: transparent; }")
            self._scroll.setWidget(self._inner)
            self.outer.addWidget(self._scroll, 1)
            self._sized = False
        else:
            self.outer.addLayout(self.form)
        self.warn = label("", wrap_text=True)
        self.warn.setStyleSheet("color:#A3260F; font-weight:600;")
        self.outer.addWidget(self.warn)
        self.read_only = False

    def showEvent(self, e):
        super().showEvent(e)
        if self._scroll is None or self._sized:
            return
        # open tall enough to show the whole form, but never taller than ~85% of the screen
        self._sized = True
        scr = self.screen() or QGuiApplication.primaryScreen()
        cap = int(scr.availableGeometry().height() * 0.85)
        chrome = self.sizeHint().height() - self._scroll.sizeHint().height()
        want = chrome + self._inner.sizeHint().height() + 4
        self.resize(self.width(), min(want, cap))
        geo = self.frameGeometry()
        geo.moveCenter(scr.availableGeometry().center() if self.parent() is None
                       else self.parentWidget().window().frameGeometry().center())
        geo.moveTop(max(geo.top(), scr.availableGeometry().top()))
        self.move(geo.topLeft())

    def lock_unless(self, allowed: bool):
        """Part 5: a row the user may not change opens read-only (the service would refuse anyway)."""
        if allowed:
            return
        self.read_only = True
        self.save_btn.setEnabled(False)
        self.save_btn.setToolTip(t("auth.not_yours"))
        self.warn.setText(t("auth.read_only"))

    def finish(self, save_text: str | None = None):
        note = label(t("all_logged"), "muted")
        self.outer.addWidget(note)
        row = QHBoxLayout()
        row.addStretch(1)
        c = button(t("cancel"))
        c.clicked.connect(self.reject)
        self.save_btn = button(save_text or t("save"), primary=True)
        self.save_btn.clicked.connect(self.on_save)
        row.addWidget(c)
        row.addWidget(self.save_btn)
        self.outer.addLayout(row)

    def status_combo(self, values, current):
        cb = QComboBox()
        cb.addItem(t("none"), None)
        for v in values:
            cb.addItem(v, v)
        i = cb.findData(current)
        cb.setCurrentIndex(i if i >= 0 else 0)
        return cb


class UpdateDialog(_Base):
    """Update a UAT Plan task or a Test Progress app (status, dates, remark, results)."""

    def __init__(self, ctx, sheet: str, item_id: str, parent=None):
        super().__init__(ctx, t("upd.title", id=item_id), parent, scrollable=True)
        self.sheet, self.item_id = sheet, item_id
        self.rec = ctx.repo.get_item(sheet, item_id)
        r = self.rec
        self.inputs = {}
        if sheet == UAT_PLAN:
            self._uat_fields(r)
        else:
            self.inputs["status"] = self.status_combo(TASK_STATUSES, r.get("status"))
            self.form.addRow(t("upd.status"), self.inputs["status"])
            for f, key in (("planned_start", "upd.p_start"), ("planned_finish", "upd.p_finish"),
                           ("actual_start", "upd.a_start"), ("actual_finish", "upd.a_finish")):
                dp = DatePicker(r.get(f))
                dp.changed.connect(self.check)
                self.inputs[f] = dp
                self.form.addRow(t(key), dp)
        if sheet == TEST_PROGRESS:
            for f, key in (("total_tc", "upd.total"), ("executed", "upd.executed"), ("pass_count", "upd.pass"),
                           ("fail_count", "upd.fail")):
                sb = QSpinBox()
                sb.setRange(0, 100000)
                sb.setValue(int(r.get(f) or 0))
                self.inputs[f] = sb
                self.form.addRow(t(key), sb)
            for f, key in (("report_by_email", "upd.email"), ("cc_emails", "upd.cc")):
                le = QLineEdit(r.get(f) or "")
                self.inputs[f] = le
                self.form.addRow(t(key), le)
        self.inputs["remark"] = QPlainTextEdit(r.get("remark") or "")
        self.inputs["remark"].setFixedHeight(80)
        self.form.addRow(t("upd.remark"), self.inputs["remark"])
        self.inputs["status"].currentIndexChanged.connect(self.check)
        # reschedule block (shown when a planned date moves)
        self.moved = label("", wrap_text=True)
        self.moved.setStyleSheet("color:#7A4E00; background:#FFF1D6; border-radius:8px; padding:8px;")
        self.reason = QLineEdit()
        self.cause = QComboBox()
        self.cause.addItems(list(ROOT_CAUSES))
        self.reason_lab, self.cause_lab = label(t("upd.reason")), label(t("upd.cause"))
        self.form.addRow(self.moved)
        self.form.addRow(self.reason_lab, self.reason)
        self.form.addRow(self.cause_lab, self.cause)
        self.reason.textChanged.connect(self.check)
        self.finish()
        self.check()
        self.lock_unless(ctx.can(auth.EDIT, sheet, self.rec))

    def _uat_fields(self, r):
        """UAT Plan: every field, in the same order as New task. Group is shown, not edited (moving a
        task to another group means moving its Excel row)."""
        grp = QLineEdit(f"{r.get('group_code') or ''}: {r.get('group_title') or ''}")
        grp.setReadOnly(True)
        grp.setFocusPolicy(Qt.NoFocus)
        grp.setToolTip(t("task.group_fixed"))
        self.form.addRow(t("task.group"), grp)

        def text(f, key, multi=False):
            w = QPlainTextEdit(r.get(f) or "") if multi else QLineEdit(r.get(f) or "")
            if multi:
                w.setFixedHeight(64)
                w.textChanged.connect(self.check)
            else:
                w.textChanged.connect(self.check)
            self.inputs[f] = w
            self.form.addRow(t(key), w)

        def date(f, key):
            dp = DatePicker(r.get(f))
            dp.changed.connect(self.check)
            self.inputs[f] = dp
            self.form.addRow(t(key), dp)

        text("task_no", "task.no")
        text("task_name", "task.name", multi=True)
        text("responsible", "task.owner", multi=True)
        self.inputs["status"] = self.status_combo(TASK_STATUSES, r.get("status"))
        self.form.addRow(t("upd.status"), self.inputs["status"])
        date("start_date", "upd.start")
        text("start_time", "task.start_time")
        date("finish_date", "upd.finish")
        text("finish_time", "task.finish_time")
        text("note", "uat.note", multi=True)
        text("sub_task", "task.sub_task")
        text("file_path", "task.file_path")
        for f in ("start_time", "finish_time"):
            self.inputs[f].setPlaceholderText("HH:MM")

    def _changes(self) -> dict:
        ch = {}
        for f, w in self.inputs.items():
            if isinstance(w, DatePicker):
                v = w.value()
            elif isinstance(w, QComboBox):
                v = w.currentData()
            elif isinstance(w, QSpinBox):
                v = w.value() if (w.value() or self.rec.get(f) is not None) else None
            elif isinstance(w, QPlainTextEdit):
                v = w.toPlainText().strip() or None
            else:
                v = w.text().strip() or None
            ch[f] = v
        return ch

    def _moved(self, ch) -> list[str]:
        planned = ("start_date", "finish_date") if self.sheet == UAT_PLAN else ("planned_start", "planned_finish")
        return [f for f in planned if self.rec.get(f) and ch.get(f) and ch[f] != self.rec.get(f)]

    def check(self):
        ch = self._changes()
        moved = self._moved(ch)
        lines = []
        for f in moved:
            n = len([h for h in self.ctx.repo.date_history(self.sheet, self.item_id, f)]) + 1
            what = t("upd.finish" if "finish" in f else "upd.start") if self.sheet == UAT_PLAN else \
                t("upd.p_finish" if "finish" in f else "upd.p_start")
            lines.append(t("upd.moved", what=what, a=fmt(self.rec.get(f)), b=fmt(ch[f]), n=n))
        show = bool(moved)
        self.moved.setText("\n".join(lines))
        for w in (self.moved, self.reason, self.cause, self.reason_lab, self.cause_lab):
            w.setVisible(show)
        problems = []
        if show and not self.reason.text().strip():
            problems.append(t("upd.need_reason"))
        if self.sheet == UAT_PLAN:
            if not ch.get("task_name"):
                problems.append(t("task.need_name"))
            if not ch.get("responsible"):
                problems.append(t("task.need_owner"))
            s, f = ch.get("start_date"), ch.get("finish_date")
            if s and f and f < s:
                problems.append(t("task.finish_before_start"))
        if self.sheet == TEST_PROGRESS:
            af, pf = ch.get("actual_finish"), ch.get("planned_finish")
            if ch.get("status") == "Complete" and not af:
                problems.append(t("upd.need_af"))
            if af and pf and af > pf:
                problems.append(t("upd.af_late"))
        if self.read_only:
            return
        self.warn.setText("\n".join(problems))
        self.save_btn.setEnabled(not problems) if hasattr(self, "save_btn") else None

    def on_save(self):
        ch = self._changes()
        try:
            changed = edit_item(self.ctx.cfg, self.ctx.repo, self.sheet, self.item_id, ch,
                                actor=self.ctx.actor, reason=self.reason.text().strip() or None,
                                root_cause=self.cause.currentText() if self._moved(ch) else None)
        except (EditError, PermissionDenied) as exc:
            error(self, str(exc))
            return
        self.changed = changed
        self.accept()


NEW_GROUP = "__new__"


class AddTaskDialog(_Base):
    def __init__(self, ctx, groups: list[tuple[str, str]], parent=None):
        super().__init__(ctx, t("task.new"), parent)
        self.group = QComboBox()
        for code, title in groups:
            self.group.addItem(f"{code}: {title}", code)
        self.group.addItem(t("task.new_group"), NEW_GROUP)      # multi-project: new projects start empty
        self.new_group = QLineEdit()
        self.new_group.setAccessibleName(t("task.new_group_title"))
        self.task_no, self.name, self.owner = QLineEdit(), QLineEdit(), QLineEdit()
        self.status = self.status_combo(TASK_STATUSES, None)
        self.start, self.finish_d = DatePicker(), DatePicker()
        self.note = QLineEdit()
        for lab, w in ((t("task.group"), self.group), (t("task.new_group_title"), self.new_group),
                       (t("task.no"), self.task_no), (t("task.name"), self.name),
                       (t("task.owner"), self.owner), (t("upd.status"), self.status), (t("upd.start"), self.start),
                       (t("upd.finish"), self.finish_d), (t("uat.note"), self.note)):
            self.form.addRow(lab, w)
        self.group.currentIndexChanged.connect(self._group_changed)
        self._group_changed()
        self.finish()

    def _group_changed(self, *_):
        new = self.group.currentData() == NEW_GROUP
        self.new_group.setVisible(new)
        self.form.labelForField(self.new_group).setVisible(new)

    def on_save(self):
        new = self.group.currentData() == NEW_GROUP
        try:
            self.new_id = add_task(self.ctx.cfg, self.ctx.repo, group=None if new else self.group.currentData(),
                                   task_no=self.task_no.text(), task_name=self.name.text(),
                                   responsible=self.owner.text(), status=self.status.currentData(),
                                   start_date=self.start.value(), finish_date=self.finish_d.value(),
                                   note=self.note.text(), actor=self.ctx.actor,
                                   new_group_title=self.new_group.text() if new else None)
        except (EditError, PermissionDenied) as exc:
            error(self, str(exc))
            return
        info(self, t("task.queued"))
        self.accept()


class AddApplicationDialog(_Base):
    """+ Add Application (multi-project): a new Test Progress row, written to this project's file only."""

    def __init__(self, ctx, parent=None):
        super().__init__(ctx, t("app.new"), parent)
        self.app, self.report_by, self.email, self.cc = QLineEdit(), QLineEdit(), QLineEdit(), QLineEdit()
        self.total = QSpinBox()
        self.total.setRange(0, 100000)
        self.pstart, self.pfinish = DatePicker(), DatePicker()
        self.remark = QLineEdit()
        for lab, w in ((t("app.name"), self.app), (t("app.report_by"), self.report_by), (t("app.email"), self.email),
                       (t("app.cc"), self.cc), (t("app.total"), self.total), (t("app.pstart"), self.pstart),
                       (t("app.pfinish"), self.pfinish), (t("upd.remark"), self.remark)):
            self.form.addRow(lab, w)
        self.finish()
        self.lock_unless(ctx.can(auth.ADD_APP))

    def on_save(self):
        try:
            self.new_id = add_application(self.ctx.cfg, self.ctx.repo, application=self.app.text(),
                                          report_by=self.report_by.text(), report_by_email=self.email.text(),
                                          cc_emails=self.cc.text(), total_tc=self.total.value() or None,
                                          planned_start=self.pstart.value() or "",
                                          planned_finish=self.pfinish.value() or "", remark=self.remark.text(),
                                          actor=self.ctx.actor)
        except (EditError, PermissionDenied) as exc:
            error(self, str(exc))
            return
        info(self, t("task.queued"))
        self.accept()


class DefectDialog(_Base):
    """Add a defect, or edit one (owner + fix target date required)."""

    def __init__(self, ctx, item_id: str | None = None, parent=None):
        super().__init__(ctx, t("def.edit", id=item_id) if item_id else t("def.new"), parent)
        self.item_id = item_id
        r = ctx.repo.get_item(DEFECT_LIST, item_id) if item_id else {}
        self.rec = r
        self.system, self.owner = QLineEdit(r.get("system") or ""), QLineEdit(r.get("assigned_to") or "")
        self.detail = QPlainTextEdit(r.get("detail") or "")
        self.detail.setFixedHeight(90)
        self.severity, self.raised = QLineEdit(r.get("severity") or ""), QLineEdit(r.get("raised_by") or "")
        self.target = DatePicker(r.get("fix_target_date"))
        self.found = DatePicker(r.get("defect_date") or ctx.as_of.isoformat(), clearable=False)
        self.status = self.status_combo(DEFECT_STATUSES, r.get("status") or "Open")
        self.remark = QLineEdit(r.get("remark") or "")
        rows = [(t("def.system"), self.system), (t("def.detail"), self.detail), (t("def.owner"), self.owner),
                (t("def.target"), self.target), (t("def.severity"), self.severity), (t("def.raised"), self.raised),
                (t("def.date"), self.found)]
        if item_id:
            rows.insert(3, (t("upd.status"), self.status))
        rows.append((t("upd.remark"), self.remark))
        for lab, w in rows:
            self.form.addRow(lab, w)
        self.reason = QLineEdit()
        self.form.addRow(t("upd.reason"), self.reason)
        self.reason.setVisible(bool(item_id))
        self.form.labelForField(self.reason).setVisible(bool(item_id))
        self.finish()
        self.lock_unless(ctx.can(auth.EDIT, DEFECT_LIST, r) if item_id else ctx.can(auth.ADD_DEFECT))

    def on_save(self):
        try:
            if self.item_id is None:
                self.new_id = add_defect(self.ctx.cfg, self.ctx.repo, system=self.system.text(),
                                         detail=self.detail.toPlainText(), assigned_to=self.owner.text(),
                                         fix_target_date=self.target.value() or "", severity=self.severity.text(),
                                         raised_by=self.raised.text(), remark=self.remark.text(),
                                         defect_date=self.found.value(), actor=self.ctx.actor)
                info(self, t("task.queued"))
            else:
                if not self.owner.text().strip() or not self.target.value():
                    raise EditError(t("def.owner") + " / " + t("def.target"))
                edit_item(self.ctx.cfg, self.ctx.repo, DEFECT_LIST, self.item_id,
                          {"system": self.system.text(), "detail": self.detail.toPlainText(),
                           "assigned_to": self.owner.text(), "fix_target_date": self.target.value(),
                           "severity": self.severity.text(), "raised_by": self.raised.text(),
                           "defect_date": self.found.value(), "status": self.status.currentData(),
                           "remark": self.remark.text()},
                          actor=self.ctx.actor, reason=self.reason.text().strip() or None)
        except (EditError, PermissionDenied) as exc:
            error(self, str(exc))
            return
        self.accept()


class ReminderDialog(_Base):
    """Manual reminder (UAT Plan / Defect). Recipient remembered per owner name. Respects dry-run."""

    def __init__(self, ctx, item_id: str, owner: str, parent=None):
        super().__init__(ctx, t("rem.title", id=item_id), parent)
        self.item_id, self.owner_name = item_id, owner
        self.to = QLineEdit(lookup_contact(ctx.repo, owner) or "")
        self.cc = QLineEdit()
        self.note = QPlainTextEdit()
        self.note.setFixedHeight(80)
        self.form.addRow(t("col.owner"), QLabel(owner or t("none")))
        self.form.addRow(t("rem.to"), self.to)
        self.form.addRow(t("rem.cc"), self.cc)
        self.form.addRow(t("rem.note"), self.note)
        if ctx.cfg.alerts.dry_run:
            d = label(t("rem.dry"), wrap_text=True)
            d.setStyleSheet("color:#7A4E00; background:#FFF1D6; border-radius:8px; padding:8px;")
            self.outer.addWidget(d)
        self.finish(t("rem.send"))

    def on_save(self):
        sender = OutlookSender(self.ctx.cfg.alerts.send_account)
        try:
            r = send_reminder(self.ctx.cfg, self.ctx.repo, sender, self.item_id, to=self.to.text().strip(),
                              cc=[c for c in self.cc.text().split(";") if c.strip()],
                              note=self.note.toPlainText().strip() or None, actor=self.ctx.actor)
        except (ReminderError, SendError, PermissionDenied) as exc:
            error(self, str(exc))
            return
        remember_contact(self.ctx.repo, self.owner_name, self.to.text().strip())
        info(self, t("rem.result", r=r))
        self.accept()
