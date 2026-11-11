"""Profile form (Part 5 v2 + multi-project): First name / Last name / Email. First start on a PC, and
"Switch user". The person is the same in every project; the role comes from each project."""

from __future__ import annotations

from PySide6.QtWidgets import QDialog, QFormLayout, QHBoxLayout, QLineEdit, QVBoxLayout

from ..services import projects
from ..services.profile import ProfileError, valid_email
from .i18n import get_lang, set_lang, t
from .widgets import button, card, label


class ProfileDialog(QDialog):
    def __init__(self, base, appdb, *, previous=None, message: str = "", parent=None):
        super().__init__(parent)
        self.base, self.appdb, self.previous, self.person = base, appdb, previous, None
        self.setWindowTitle(t("profile.title"))
        self.setMinimumWidth(480)
        self.setObjectName("page")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        c, cl = card()
        top = QHBoxLayout()
        top.addWidget(label(t("profile.title"), "h1"))
        top.addStretch(1)
        self.lang_btn = button("EN" if get_lang() == "th" else "TH")
        self.lang_btn.setFixedWidth(56)
        self.lang_btn.clicked.connect(self._toggle_lang)
        top.addWidget(self.lang_btn)
        cl.addLayout(top)
        cl.addWidget(label(t("profile.sub"), "muted", wrap_text=True))
        self.notice = label(message, wrap_text=True)
        self.notice.setStyleSheet("color:#7A4E00; background:#FFF1D6; border-radius:8px; padding:8px;")
        self.notice.setVisible(bool(message))
        cl.addWidget(self.notice)
        self.boot = label(t("proj.first_admin"), wrap_text=True)
        self.boot.setStyleSheet("color:#1E4E8C; background:#E6EFFB; border-radius:8px; padding:8px;")
        self.boot.setVisible(not appdb.projects(include_archived=True))
        cl.addWidget(self.boot)
        form = QFormLayout()
        form.setSpacing(10)
        self.email, self.first, self.last = QLineEdit(), QLineEdit(), QLineEdit()
        self.email.setPlaceholderText("name@company.com")
        for w, key in ((self.email, "profile.email"), (self.first, "profile.first"), (self.last, "profile.last")):
            w.setAccessibleName(t(key))
            form.addRow(t(key), w)
        self.email.textChanged.connect(self._email_changed)
        cl.addLayout(form)
        self.hint = label("", "muted", wrap_text=True)
        cl.addWidget(self.hint)
        self.status = label("", wrap_text=True)
        self.status.setStyleSheet("color:#A3260F; font-weight:600;")
        cl.addWidget(self.status)
        row = QHBoxLayout()
        row.addStretch(1)
        if previous is not None:
            cancel = button(t("cancel"))
            cancel.clicked.connect(self.reject)
            row.addWidget(cancel)
        self.save_btn = button(t("profile.save"), primary=True)
        self.save_btn.clicked.connect(self.on_save)
        row.addWidget(self.save_btn)
        cl.addLayout(row)
        cl.addWidget(label(t("profile.limit"), "muted", wrap_text=True))
        outer.addWidget(c)
        self.email.setFocus()

    def _email_changed(self, text):
        """An existing email keeps its stored name (approved): fill it in and lock it."""
        p = projects.find_person(self.base, self.appdb, text) if valid_email(text) else None
        if p is not None and projects.is_deactivated_everywhere(self.base, self.appdb, p.email):
            self.hint.setText(t("profile.inactive", e=p.email))
        elif p is not None:
            self.first.setText(p.first_name)
            self.last.setText(p.last_name)
            self.hint.setText(t("proj.profile_known", n=p.full_name))
        else:
            self.hint.setText(t("proj.profile_new") if valid_email(text) and self.appdb.projects() else "")
        for w in (self.first, self.last):
            w.setReadOnly(p is not None and bool(w.text()))
        self.status.setText("")

    def on_save(self):
        email = self.email.text()
        if not valid_email(email):
            self.status.setText(t("profile.bad_email"))
            return
        try:
            self.person = projects.identify_person(self.base, self.appdb, first_name=self.first.text(),
                                                   last_name=self.last.text(), email=email)
        except ProfileError as exc:
            self.status.setText(str(exc))
            return
        self.accept()

    def _toggle_lang(self):
        lang = "en" if get_lang() == "th" else "th"
        set_lang(lang)
        with self.appdb.transaction():
            self.appdb.set_setting("ui.lang", lang)
        self.done(2)                       # the caller opens the form again in the new language
