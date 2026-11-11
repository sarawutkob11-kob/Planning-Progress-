"""Settings -> Users (Part 5, Admin only): list, add, edit role / names in Excel, deactivate, import."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QDialog, QHBoxLayout, QHeaderView,
                               QLineEdit, QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QVBoxLayout)

from ..services.auth import QA, ROLES, PermissionDenied, name_key
from ..services.users import UserError, add_user, excel_names, import_candidates, list_users, update_user
from .dialogs import _Base, confirm, error, info
from .i18n import t
from .widgets import button, card, label


def _ts(v) -> str:
    return (v or "")[:16].replace("T", " ")


def users_card(screen):
    """Card for the Settings screen. screen.refresh() is called after every change."""
    ctx = screen.ctx
    c, cl = card()
    head = QHBoxLayout()
    head.addWidget(label(t("users.title"), "h2"))
    head.addStretch(1)
    add, edit, act, imp = (button(t("users.add"), primary=True), button(t("users.edit")),
                           button(t("users.toggle")), button(t("users.import")))
    for b in (imp, act, edit, add):
        head.addWidget(b)
    cl.addLayout(head)
    cl.addWidget(label(t("users.note"), "muted", wrap_text=True))
    users = list_users(ctx.repo)
    tb = QTableWidget(len(users), 7)
    tb.setHorizontalHeaderLabels([t("users.first"), t("users.lastname"), t("users.email"), t("users.role"),
                                  t("users.active"), t("users.last"), t("users.aliases")])
    tb.verticalHeader().setVisible(False)
    tb.setEditTriggers(QAbstractItemView.NoEditTriggers)
    tb.setSelectionBehavior(QAbstractItemView.SelectRows)
    tb.setSelectionMode(QAbstractItemView.SingleSelection)
    for i, u in enumerate(users):
        vals = (u["first_name"] or "", u["last_name"] or "", u["email"], u["role"],
                t("yes") if u["active"] else t("users.off"), _ts(u["last_used_at"]), ", ".join(u["aliases"]))
        for j, v in enumerate(vals):
            it = QTableWidgetItem(v)
            if not u["active"]:
                it.setForeground(Qt.darkGray)
            tb.setItem(i, j, it)
    tb.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    tb.horizontalHeader().setStretchLastSection(True)
    tb.setMinimumHeight(200)
    cl.addWidget(tb)
    screen.users_table, screen.users = tb, users

    def selected():
        r = tb.currentRow()
        return users[r] if 0 <= r < len(users) else None

    def do_add():
        if UserDialog(ctx, None, screen).exec():
            screen.refresh()

    def do_edit():
        u = selected()
        if u and UserDialog(ctx, u, screen).exec():
            screen.refresh()

    def do_toggle():
        u = selected()
        if not u:
            return
        q = t("users.deactivate_q" if u["active"] else "users.reactivate_q", e=u["email"])
        if not confirm(screen, q):
            return
        try:
            update_user(ctx.repo, ctx.actor, u["email"], active=not u["active"])
        except (UserError, PermissionDenied) as exc:
            error(screen, str(exc))
            return
        screen.refresh()

    def do_import():
        if ImportUsersDialog(ctx, screen).exec():
            screen.refresh()

    add.clicked.connect(do_add)
    edit.clicked.connect(do_edit)
    act.clicked.connect(do_toggle)
    imp.clicked.connect(do_import)
    tb.cellDoubleClicked.connect(lambda r, _c: do_edit())
    return c


class UserDialog(_Base):
    """Add a user, or edit name / role / names as written in Excel."""

    def __init__(self, ctx, user: dict | None, parent=None):
        super().__init__(ctx, t("users.edit_title" if user else "users.add_title"), parent)
        self.user = user
        self.email = QLineEdit(user["email"] if user else "")
        self.email.setReadOnly(bool(user))
        self.first = QLineEdit((user or {}).get("first_name") or "")
        self.last = QLineEdit((user or {}).get("last_name") or "")
        self.role = QComboBox()
        for r in ROLES:
            self.role.addItem(r, r)
        self.role.setCurrentIndex(max(0, self.role.findData(user["role"] if user else QA)))
        self.form.addRow(t("users.email"), self.email)
        self.form.addRow(t("users.first"), self.first)
        self.form.addRow(t("users.lastname"), self.last)
        self.form.addRow(t("users.role"), self.role)
        self.form.addRow(label(t("users.aliases_help"), "muted", wrap_text=True))
        self.filter = QLineEdit()
        self.filter.setPlaceholderText(t("users.alias_filter"))
        self.filter.textChanged.connect(self._filter)
        self.names = QListWidget()
        self.names.setMinimumHeight(220)
        mine = {name_key(a) for a in (user or {}).get("aliases", [])}
        known = excel_names(ctx.repo)
        extra = [a for a in (user or {}).get("aliases", []) if name_key(a) not in {name_key(k) for k in known}]
        for n in sorted(set(known) | set(extra), key=str.casefold):
            it = QListWidgetItem(n)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if name_key(n) in mine else Qt.Unchecked)
            self.names.addItem(it)
        self.other = QLineEdit()
        self.other.setPlaceholderText(t("users.alias_other"))
        self.form.addRow(t("users.aliases"), self.filter)
        self.form.addRow("", self.names)
        self.form.addRow("", self.other)
        self.finish()

    def _filter(self, text):
        k = text.strip().casefold()
        for i in range(self.names.count()):
            it = self.names.item(i)
            it.setHidden(bool(k) and k not in it.text().casefold())

    def aliases(self) -> list[str]:
        out = [self.names.item(i).text() for i in range(self.names.count())
               if self.names.item(i).checkState() == Qt.Checked]
        out += [x.strip() for x in self.other.text().split(";") if x.strip()]
        return out

    def on_save(self):
        try:
            if self.user is None:
                add_user(self.ctx.repo, self.ctx.actor, self.email.text(), role=self.role.currentData(),
                         first_name=self.first.text(), last_name=self.last.text(), aliases=self.aliases())
            else:
                update_user(self.ctx.repo, self.ctx.actor, self.user["email"], role=self.role.currentData(),
                            first_name=self.first.text(), last_name=self.last.text(), aliases=self.aliases())
        except (UserError, PermissionDenied) as exc:
            error(self, str(exc))
            return
        self.accept()


class ImportUsersDialog(_Base):
    """Addresses from Test Progress Report By Email + CC Email that are not users yet. Nothing is saved
    until the Admin ticks them and clicks Save."""

    def __init__(self, ctx, parent=None):
        super().__init__(ctx, t("users.import_title"), parent)
        self.setMinimumWidth(820)
        self.cands = import_candidates(ctx.repo)
        lay = QVBoxLayout()
        lay.addWidget(label(t("users.import_note") if self.cands else t("users.import_none"), "muted",
                            wrap_text=True))
        self.tb = QTableWidget(len(self.cands), 5)
        self.tb.setHorizontalHeaderLabels(["", t("users.email"), t("users.role"), t("users.source"), t("users.apps")])
        self.tb.verticalHeader().setVisible(False)
        self.checks, self.roles = [], []
        for i, c in enumerate(self.cands):
            cb = QCheckBox()
            self.checks.append(cb)
            self.tb.setCellWidget(i, 0, cb)
            self.tb.setItem(i, 1, QTableWidgetItem(c["email"]))
            rc = QComboBox()
            for r in ROLES:
                rc.addItem(r, r)
            rc.setCurrentIndex(rc.findData(c["role"]))
            self.roles.append(rc)
            self.tb.setCellWidget(i, 2, rc)
            self.tb.setItem(i, 3, QTableWidgetItem(", ".join(c["sources"])))
            self.tb.setItem(i, 4, QTableWidgetItem(", ".join(c["apps"])))
        self.tb.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tb.horizontalHeader().setStretchLastSection(True)
        self.tb.setMinimumHeight(260)
        lay.addWidget(self.tb)
        self.outer.insertLayout(0, lay)
        self.finish()
        self.save_btn.setEnabled(bool(self.cands))

    def on_save(self):
        done, errs = 0, []
        for c, cb, rc in zip(self.cands, self.checks, self.roles):
            if not cb.isChecked():
                continue
            try:
                add_user(self.ctx.repo, self.ctx.actor, c["email"], role=rc.currentData(),
                         first_name=c["names"][0] if c["names"] else "", aliases=c["names"])
                done += 1
            except (UserError, PermissionDenied) as exc:
                errs.append(str(exc))
        if errs:
            error(self, "\n".join(errs))
        info(self, t("users.imported", n=done))
        self.accept()
