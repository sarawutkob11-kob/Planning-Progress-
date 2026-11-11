"""Screen 7: Conflict Center (approved mockup) - Conflicts tab + PMO review queue tab."""

from __future__ import annotations

import logging
import re

from PySide6.QtWidgets import (QComboBox, QDialog, QFormLayout, QGridLayout, QHBoxLayout, QLineEdit, QVBoxLayout,
                               QWidget)

from ...services import auth
from ...services.auth import PermissionDenied
from ...domain.models import DATE_FIELDS, HISTORY_SUFFIX
from ...excel.sheets import DEFECT_LIST
from ...parsing.dates import fmt
from ...parsing.status import DEFECT_STATUSES, TASK_STATUSES
from ...services.resolve import ResolveError, resolve_conflict, resolve_review
from ...services.review_suggest import suggestions
from ...services.sync import SyncService
from ..dialogs import error, info
from ..i18n import t
from ..widgets import DatePicker, button, card, label, tag
from .base import ChipRow, Screen, clear_layout

log = logging.getLogger(__name__)

PROB_TAG = {"ambiguous_date": "Amber", "out_of_window": "neutral", "bad_date": "Red",
            "date_in_time_column": "Today", "unknown_status": "Amber", "bad_number": "Red"}


FIELD_KEYS = {"start_date": "upd.start", "finish_date": "upd.finish", "planned_start": "upd.p_start",
              "planned_finish": "upd.p_finish", "actual_start": "upd.a_start", "actual_finish": "upd.a_finish",
              "fix_target_date": "col.target", "defect_date": "def.date", "status": "upd.status",
              "remark": "upd.remark", "start_time": "fld.start_time", "finish_time": "fld.finish_time"}


def field_label(field: str) -> str:
    hist = field.endswith(HISTORY_SUFFIX)
    base = field[: -len(HISTORY_SUFFIX)] if hist else field
    name = t(FIELD_KEYS[base]) if base in FIELD_KEYS else base
    return name + (f" ({t('fld.earlier')})" if hist else "")


def raw_text(raw: str | None) -> str:
    raw = (raw or "").replace("\n", " | ")
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", raw)
    if m:                                               # a real Excel date: show it as Excel does
        y, mo, d = m.groups()
        return f"{d}-{fmt('2000-' + mo + '-01')[3:6]}-{y} ({t('fld.excel_date')})"
    return raw


def _show(field: str, v) -> str:
    if v in (None, ""):
        return t("none")
    return fmt(v) if field in DATE_FIELDS and len(str(v)) == 10 else str(v)


class ValueDialog(QDialog):
    """Enter a value for a field: date picker for dates, list for status, text otherwise."""

    def __init__(self, title: str, field: str, sheet: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(420)
        self.field = field
        lay = QVBoxLayout(self)
        form = QFormLayout()
        if field in DATE_FIELDS:
            self.w = DatePicker(None)
        elif field == "status":
            self.w = QComboBox()
            for s in (DEFECT_STATUSES if sheet == DEFECT_LIST else TASK_STATUSES):
                self.w.addItem(s, s)
        else:
            self.w = QLineEdit()
        form.addRow(label(title), self.w)
        lay.addLayout(form)
        row = QHBoxLayout()
        row.addStretch(1)
        c, ok = button(t("cancel")), button(t("ok"), primary=True)
        c.clicked.connect(self.reject)
        ok.clicked.connect(self.accept)
        row.addWidget(c)
        row.addWidget(ok)
        lay.addLayout(row)

    def value(self):
        if isinstance(self.w, DatePicker):
            return self.w.value()
        if isinstance(self.w, QComboBox):
            return self.w.currentData()
        return self.w.text().strip()


class ConflictsScreen(Screen):
    key = "conflicts"

    def __init__(self, ctx):
        super().__init__(ctx)
        self.tab = "conf"

    def show_item(self, item_id=None, **_):
        self.refresh()

    def refresh(self):
        repo = self.ctx.repo
        clear_layout(self.lay)
        self.lay.addWidget(label(t("nav.conflicts"), "h1"))
        conflicts, reviews = repo.open_conflicts(), repo.open_reviews()
        q = self.ctx.search.casefold()
        if q:
            conflicts = [c for c in conflicts if q in f"{c['item_id']} {c['field']}".casefold()]
            reviews = [r for r in reviews if q in f"{r['item_id']} {r['field']} {r['raw_value']}".casefold()]
        chips = ChipRow(self._pick)
        chips.set_chips([("conf", f"{t('conf.tab_conf')} · {len(conflicts)}"),
                         ("rev", f"{t('conf.tab_rev')} · {len(reviews)}")], self.tab)
        self.lay.addWidget(chips)
        if self.tab == "conf":
            if not conflicts:
                c, cl = card()
                cl.addWidget(label(t("conf.none"), "muted"))
                self.lay.addWidget(c)
            for c in conflicts:
                self.lay.addWidget(self._conflict_card(c))
        else:
            self.lay.addWidget(label(t("rev.intro"), "muted", wrap_text=True))
            if not reviews:
                c, cl = card()
                cl.addWidget(label(t("rev.none"), "muted"))
                self.lay.addWidget(c)
            for r in reviews:
                self.lay.addWidget(self._review_row(r))
        self.lay.addStretch(1)

    def _pick(self, k):
        self.tab = k
        self.refresh()

    # ------------------------------------------------------------------ conflicts
    def _name(self, sheet, item_id) -> str:
        rec = self.ctx.repo.get_item(sheet, item_id) or {}
        n = rec.get("application") or rec.get("task_name") or rec.get("system") or ""
        return str(n).split("\n")[0]

    def _conflict_card(self, c) -> QWidget:
        w, cl = card()
        head = QHBoxLayout()
        head.addWidget(tag("Red", t(f"conf.kind.{c['kind']}")))
        head.addWidget(label(f"{c['item_id']} · {self._name(c['sheet'], c['item_id'])}", "h2"))
        head.addWidget(label(f"{t('sheet.' + c['sheet'])} · {t('conf.field')}: {field_label(c['field'])}", "muted"))
        head.addStretch(1)
        cl.addLayout(head)
        grid = QGridLayout()
        grid.setSpacing(14)
        for col, (title, value, when, who) in enumerate((
                (t("conf.excel"), c["excel_value"], c["excel_time"], None),
                (t("conf.app"), c["app_value"], c["app_time"], c.get("app_editor")))):
            box, bl = card(margins=(14, 12, 14, 12), spacing=4)
            bl.addWidget(label(title, "muted"))
            v = label(_show(c["field"], value), wrap_text=True)
            v.setStyleSheet("font-size: 12pt; font-weight: 600;")
            bl.addWidget(v)
            bl.addWidget(label(" · ".join(x for x in ((when or "")[:16].replace("T", " "), who or "") if x), "muted"))
            grid.addWidget(box, 0, col)
        cl.addLayout(grid)
        reason = QLineEdit()
        reason.setPlaceholderText(t("conf.reason_ph"))
        cl.addWidget(label(t("conf.reason")))
        cl.addWidget(reason)
        row = QHBoxLayout()
        for text, choice, primary in ((t("conf.use_excel"), "excel", False), (t("conf.use_app"), "app", False),
                                      (t("conf.enter_new"), "new", True)):
            b = button(text, primary=primary)
            b.clicked.connect(lambda _=False, ch=choice, cc=c, r=reason: self._resolve(cc, ch, r.text()))
            b.setEnabled(self.ctx.can(auth.RESOLVE))       # Part 5: QA reads only
            row.addWidget(b)
        row.addStretch(1)
        cl.addLayout(row)
        cl.addWidget(label(t("conf.rule"), "muted", wrap_text=True))
        return w

    def _resolve(self, c, choice, reason):
        if not reason.strip():
            error(self, t("conf.need_reason"))
            return
        new_value = None
        if choice == "new":
            dlg = ValueDialog(t("conf.new_value", field=c["field"]), c["field"], c["sheet"], self)
            if not dlg.exec():
                return
            new_value = dlg.value()
        try:
            resolve_conflict(self.ctx.cfg, self.ctx.repo, c["id"], choice, actor=self.ctx.actor,
                             reason=reason, new_value=new_value)
        except (ResolveError, PermissionDenied) as exc:
            error(self, str(exc))
            return
        from ..worker import friendly, project_repo, run_in_worker, sync_locked
        cfg = self.ctx.cfg

        def job():                                                       # apply / re-check against Excel
            with project_repo(cfg) as repo, sync_locked(cfg):
                SyncService(cfg, repo).pull(force=True)

        def failed(exc):                                                 # Excel unreadable: decision is saved
            error(self, friendly(exc))
            self.ctx.refresh_all()
        return run_in_worker(self, t("work.reading"), job, lambda _: self.ctx.refresh_all(), failed,
                             name="resolve-pull")

    # ------------------------------------------------------------------ review queue
    def _review_row(self, r) -> QWidget:
        w, cl = card(margins=(16, 12, 16, 12), spacing=6)
        top = QHBoxLayout()
        top.addWidget(tag(PROB_TAG.get(r["problem"], "neutral"), t(f"prob.{r['problem']}")))
        top.addWidget(label(f"{self._name(r['sheet'], r['item_id'])}", "h2"))
        top.addWidget(label(f"{r['item_id']} · {t('sheet.' + r['sheet'])} · {field_label(r['field'])}", "muted"))
        top.addStretch(1)
        cl.addLayout(top)
        raw = label(raw_text(r["raw_value"]), wrap_text=True)
        raw.setStyleSheet("background:#F4F6F9; border-radius:6px; padding:4px 8px; font-family: Consolas, monospace;")
        cl.addWidget(raw)
        tok = label(t("rev.token", t=raw_text(r["token"])))
        tok.setStyleSheet("color:#A3260F;")
        cl.addWidget(tok)
        row = QHBoxLayout()
        sugg = suggestions(self.ctx.cfg, r)
        for s in sugg:
            text = (t("rev.drop") if s["why"] == "drop" else t("rev.clear") if s["why"] == "clear"
                    else fmt(s["value"]) if len(str(s["value"])) == 10 and str(s["value"])[4] == "-" else s["value"])
            b = button(text, obj="chip")
            b.setToolTip(t(f"why.{s['why']}"))
            b.setAccessibleName(f"{text} - {t('why.' + s['why'])}")
            b.clicked.connect(lambda _=False, rr=r, v=s["value"]: self._apply(rr, v))
            b.setEnabled(self.ctx.can(auth.RESOLVE))
            row.addWidget(b)
        if sugg:
            row.addWidget(label(" / ".join(dict.fromkeys(t(f"why.{s['why']}") for s in sugg)), "muted"))
        row.addStretch(1)
        base = r["field"][: -len(HISTORY_SUFFIX)] if r["field"].endswith(HISTORY_SUFFIX) else r["field"]
        other = button(t("rev.pick") if base in DATE_FIELDS else t("rev.enter"))
        other.clicked.connect(lambda _=False, rr=r, b=base: self._other(rr, b))
        other.setEnabled(self.ctx.can(auth.RESOLVE))
        row.addWidget(other)
        cl.addLayout(row)
        return w

    def _other(self, r, base):
        field = base if base in DATE_FIELDS or base == "status" else "text"
        dlg = ValueDialog(t("rev.enter"), field, r["sheet"], self)
        if dlg.exec():
            self._apply(r, dlg.value())

    def _apply(self, r, value):
        try:
            resolve_review(self.ctx.cfg, self.ctx.repo, r["id"], value or None, actor=self.ctx.actor,
                           reason="Conflict Center")
        except (ResolveError, PermissionDenied) as exc:
            error(self, str(exc))
            return
        info(self, t("rev.done"))
        self.ctx.refresh_all()
