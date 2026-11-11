"""Screen 1: Dashboard (PMO) - as approved in the mockup."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QAbstractItemView, QFrame, QGridLayout, QHBoxLayout, QHeaderView, QSizePolicy, QTableWidget,
                               QTableWidgetItem, QVBoxLayout, QWidget)

from ...services import auth
from ...excel.sheets import TEST_PROGRESS
from ...parsing.dates import fmt
from ...services import dashboard
from ..dialogs import ReminderDialog
from ..i18n import t
from ..widgets import PRIMARY, bar, button, card, label, tag, wrap
from .base import ChipRow, Screen, clear_layout

TABS = ("all", "Overdue", "Today", "No target", "Red", "Amber", "Missing data")


def _pct(x) -> str:
    return "-" if x is None else f"{x * 100:.1f}%"


def detail_text(d: dict) -> str:
    k = d["kind"]
    if k == "finish_late":
        return t("d.finish_late", date=fmt(d["date"]), days=d["days"])
    if k == "fix_late":
        return t("d.fix_late", date=fmt(d["date"]), days=d["days"])
    if k == "due_today":
        return t("d.due_today", date=fmt(d["date"]))
    if k == "past_finish":
        return t("d.past_finish", date=fmt(d["date"]))
    if k == "fail_high":
        return t("d.fail_high", pct=_pct(d["pct"]))
    if k == "behind":
        s = t("d.behind", exec=_pct(d["exec"]), time=_pct(d["time"]))
        if d["exec"] is not None and d["time"] is not None and d["exec"] >= d["time"] and d.get("fail"):
            s = t("d.behind_fail", n=d["fail"])
        return s
    return d.get("text") or ""


class DashboardScreen(Screen):
    key = "dashboard"

    def __init__(self, ctx):
        super().__init__(ctx)
        self.tab = "all"
        self.data = None

    def refresh(self):
        self.data = dashboard.build(self.ctx.cfg, self.ctx.repo, self.ctx.as_of)
        clear_layout(self.lay)
        d = self.data
        # header
        head = QHBoxLayout()
        head.addWidget(label(t("dash.title"), "h1"))
        head.addStretch(1)
        wk = button(t("dash.weekly"), primary=True)    # Part 6: Admin, PMO, User BU/PM
        wk.setVisible(self.ctx.can(auth.WEEKLY))
        wk.clicked.connect(self._weekly)
        self.weekly_btn = wk
        head.addWidget(wk)
        self.lay.addLayout(head)
        # KPI cards
        grid = QGridLayout()
        grid.setSpacing(16)
        c1, l1 = card()
        l1.addWidget(label(t("dash.kpi_exec"), "muted"))
        r = QHBoxLayout()
        r.addWidget(label(f"{d.pct_executed * 100:.0f}%", "kpi"))
        r.addWidget(label(f"{d.executed:,} / {d.total_tc:,}", "muted"), 0, Qt.AlignBottom)
        r.addStretch(1)
        l1.addLayout(r)
        l1.addWidget(bar(d.executed, d.total_tc))
        c2, l2 = card()
        l2.addWidget(label(t("dash.kpi_open"), "muted"))
        r = QHBoxLayout()
        r.addWidget(label(str(d.open_tasks), "kpi"))
        r.addWidget(label(t("dash.of", n=d.all_tasks), "muted"), 0, Qt.AlignBottom)
        r.addStretch(1)
        l2.addLayout(r)
        if d.overdue_tasks:
            l2.addWidget(tag("Overdue", t("dash.overdue_n", n=d.overdue_tasks)))
        c3, l3 = card()
        l3.addWidget(label(t("dash.kpi_today"), "muted"))
        l3.addWidget(label(str(d.due_today), "kpi"))
        if d.next_due:
            l3.addWidget(label(t("dash.next", n=d.next_due[1], d=fmt(d.next_due[0])), "muted"))
        c4, l4 = card()
        l4.addWidget(label(t("dash.kpi_defects"), "muted"))
        l4.addWidget(label(str(d.open_defects), "kpi"))
        l4.addWidget(tag("Overdue" if d.overdue_defects else "neutral", t("dash.def_overdue", n=d.overdue_defects)))
        l4.addWidget(tag("No target" if d.no_target_defects else "neutral", t("dash.def_notarget", n=d.no_target_defects)))
        for i, c in enumerate((c1, c2, c3, c4)):
            c.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)   # cards share the width, never widen the page
            c.setMinimumWidth(200)
            grid.addWidget(c, 0, i)
            grid.setColumnStretch(i, 1)
        self.lay.addLayout(grid)
        # missing email warning
        if d.missing_email:
            w = QFrame()
            w.setObjectName("warn")
            wl = QHBoxLayout(w)
            wl.setContentsMargins(16, 10, 16, 10)
            msg = label(t("dash.mail_warn", n=len(d.missing_email), m=d.apps), wrap_text=True)
            msg.setStyleSheet("font-weight:600;")
            wl.addWidget(msg, 1)
            vb = button(t("dash.view_list"))
            vb.clicked.connect(lambda: self.ctx.navigate("tp", filter="incomplete"))
            wl.addWidget(vb)
            w.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
            self.lay.addWidget(w)
        # follow-up (full width) + incomplete apps / due soon below
        fc, fl = card()
        h = QHBoxLayout()
        h.addWidget(label(t("dash.follow"), "h2"))
        rows = self._rows()
        h.addWidget(label(t("dash.items", n=len(rows)), "muted"))
        h.addStretch(1)
        fl.addLayout(h)
        chips = ChipRow(self._pick)
        counts = {k: (len(d.follow) if k == "all" else sum(1 for f in d.follow if f.tag == k)) for k in TABS}
        chips.set_chips([(k, f"{t('tag.' + k)} · {counts[k]}") for k in TABS if k == "all" or counts[k]],
                        self.tab)
        fl.addWidget(chips)
        fl.addWidget(self._table(rows))
        fc.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.lay.addWidget(fc)
        right = QHBoxLayout()
        right.setSpacing(20)
        ic, il = card()
        il.addWidget(label(t("dash.incomplete"), "h2"))
        for a in d.incomplete_apps or []:
            line = QHBoxLayout()
            line.addWidget(label(a["name"]), 1)
            pct = a["executed"] / a["total"] * 100 if a["total"] else 0
            line.addWidget(label(f"{a['executed']} / {a['total']} · {pct:.0f}%", "muted"))
            if a["color"] in ("Red", "Amber", "Green", "Missing"):
                line.addWidget(tag(a["color"]))
            il.addLayout(line)
            il.addWidget(bar(a["executed"], a["total"], {"Red": "#A3260F", "Amber": "#7A4E00"}.get(a["color"], PRIMARY)))
        if not d.incomplete_apps:
            il.addWidget(label(t("dash.nothing"), "muted"))
        ic.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        right.addWidget(ic, 1)
        dc, dl = card()
        dl.addWidget(label(t("dash.due2"), "h2"))
        last = None
        for x in d.due_soon:
            if x["date"] != last:
                dh = label(fmt(x["date"]))
                dh.setStyleSheet("color:#1F4FBF; font-weight:700;")
                dl.addWidget(dh)
                last = x["date"]
            line = QHBoxLayout()
            idl = label(x["item_id"], "muted")
            idl.setFixedWidth(64)
            line.addWidget(idl)
            line.addWidget(label(x["name"], wrap_text=True), 1)
            line.addWidget(label(x["owner"], "muted"))
            dl.addLayout(line)
        if not d.due_soon:
            dl.addWidget(label(t("dash.nothing"), "muted"))
        dc.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        right.addWidget(dc, 1)
        self.lay.addLayout(right)
        self.lay.addStretch(1)

    def _pick(self, key):
        self.tab = key
        self.refresh()

    def _rows(self):
        q = self.ctx.search.casefold()
        return [f for f in self.data.follow
                if (self.tab == "all" or f.tag == self.tab)
                and (not q or q in f"{f.item_id} {f.name} {f.owner}".casefold())]

    def _table(self, rows) -> QTableWidget:
        tb = QTableWidget(len(rows), 5)
        tb.setHorizontalHeaderLabels([t("col.tag"), t("col.item"), t("col.detail"), t("col.owner"), t("col.action")])
        tb.verticalHeader().setVisible(False)
        tb.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tb.setSelectionBehavior(QAbstractItemView.SelectRows)
        tb.setAlternatingRowColors(True)
        tb.setWordWrap(True)
        hh = tb.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)
        hh.setSectionResizeMode(2, QHeaderView.Stretch)
        for col, width in ((0, 140), (1, 300), (3, 200), (4, 190)):
            tb.setColumnWidth(col, width)
        for i, f in enumerate(rows):
            tb.setCellWidget(i, 0, wrap(tag(f.tag)))
            it = QTableWidgetItem(f"{f.name}\n{f.item_id} · {t('sheet.' + f.sheet)}")
            tb.setItem(i, 1, it)
            tb.setItem(i, 2, QTableWidgetItem(detail_text(f.detail)))
            tb.setItem(i, 3, QTableWidgetItem(f.owner))
            if f.manual and self.ctx.can(auth.REMIND):
                b = button(t("dash.remind"), primary=True)
                b.clicked.connect(lambda _=False, x=f: self._remind(x))
                tb.setCellWidget(i, 4, wrap(b))
            elif f.manual:
                tb.setItem(i, 4, QTableWidgetItem(t("none")))
            else:
                tb.setItem(i, 4, QTableWidgetItem(t("dash.auto")))
        tb.resizeRowsToContents()
        for i in range(len(rows)):
            tb.setRowHeight(i, max(tb.rowHeight(i), 56))
        tb.setMinimumHeight(min(520, 60 + 60 * max(len(rows), 1)))
        tb.cellDoubleClicked.connect(lambda r, c: self._open(rows[r]))
        return tb

    def _remind(self, f):
        ReminderDialog(self.ctx, f.item_id, f.owner, self).exec()

    def _open(self, f):
        self.ctx.navigate({"UAT Plan": "uat", TEST_PROGRESS: "tp", "Defect List": "defects"}[f.sheet],
                          item=f.item_id)

    def _weekly(self):
        from ..weekly import start_generate
        start_generate(self.ctx, self)
