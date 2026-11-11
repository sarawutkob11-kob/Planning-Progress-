"""Screen 2: UAT Plan - groups A-E, filters, detail panel."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import (QComboBox, QHBoxLayout, QHeaderView, QScrollArea, QSplitter, QTreeWidget,
                               QTreeWidgetItem, QVBoxLayout, QWidget)

from ...services import auth
from ...excel.sheets import UAT_PLAN
from ...parsing.dates import fmt
from ...services import views
from ...services.sync import SyncService
from ..dialogs import AddTaskDialog, UpdateDialog, error, info
from ..i18n import t
from ..widgets import bar, button, card, hline, label, tag
from .base import ChipRow, Screen, clear_layout


class UatPlanScreen(Screen):
    key = "uat"

    def __init__(self, ctx):
        super().__init__(ctx, scroll=False)
        self.flt, self.owner, self.selected = "all", None, None
        top = QHBoxLayout()
        self.title = label("", "h1")
        top.addWidget(self.title)
        top.addStretch(1)
        self.import_btn = button("")
        self.import_btn.clicked.connect(self._import)
        self.add_btn = button("", primary=True)
        self.add_btn.clicked.connect(self._add)
        top.addWidget(self.import_btn)
        top.addWidget(self.add_btn)
        self.lay.addLayout(top)
        frow = QHBoxLayout()
        self.chips = ChipRow(self._pick)
        frow.addWidget(self.chips, 1)
        self.owner_cb = QComboBox()
        self.owner_cb.setMinimumWidth(220)
        self.owner_cb.currentIndexChanged.connect(self._owner)
        frow.addWidget(self.owner_cb)
        self.lay.addLayout(frow)
        split = QSplitter(Qt.Horizontal)
        self.tree = QTreeWidget()
        self.tree.setAlternatingRowColors(True)
        self.tree.itemSelectionChanged.connect(self._select)
        self.tree.itemDoubleClicked.connect(lambda it, c: self._update())
        split.addWidget(self.tree)
        self.panel = QWidget()
        self.panel_lay = QVBoxLayout(self.panel)
        self.panel_lay.setContentsMargins(16, 0, 0, 0)
        sa = QScrollArea()
        sa.setWidgetResizable(True)
        sa.setWidget(self.panel)
        sa.setMinimumWidth(320)
        split.addWidget(sa)
        split.setSizes([1000, 340])
        self.lay.addWidget(split, 1)

    def show_item(self, item_id: str):
        self.flt, self.owner, self.selected = "all", None, item_id
        self.refresh()

    def refresh(self):
        self.title.setText(t("nav.uat"))
        self.import_btn.setText(t("uat.import"))
        self.add_btn.setText(t("uat.add"))
        self.add_btn.setVisible(self.ctx.can(auth.ADD_TASK))
        self.chips.set_chips([(k, t(f"f.{k}")) for k in views.UAT_FILTERS], self.flt)
        owners = views.uat_owners(self.ctx.repo)
        self.owner_cb.blockSignals(True)
        self.owner_cb.clear()
        self.owner_cb.addItem(t("uat.owner_all"), None)
        for o in owners:
            self.owner_cb.addItem(o, o)
        self.owner_cb.setCurrentIndex(max(self.owner_cb.findData(self.owner), 0))
        self.owner_cb.blockSignals(False)
        groups = views.uat_plan(self.ctx.repo, self.ctx.as_of, self.flt, self.owner, self.ctx.search)
        self.tree.blockSignals(True)
        self.tree.clear()
        cols = [t("col.id"), t("col.task"), t("col.owner"), t("col.status"), t("col.planned_first"),
                t("col.current"), t("col.resched")]
        self.tree.setHeaderLabels(cols)
        bold = QFont()
        bold.setBold(True)
        select_item = None
        for g in groups:
            gi = QTreeWidgetItem([f"{g['code']}: {g['title']}  ({len(g['items'])})"])
            gi.setFont(0, bold)
            self.tree.addTopLevelItem(gi)
            gi.setFirstColumnSpanned(True)          # must be set after the item is in the tree
            for it in g["items"]:
                row = QTreeWidgetItem([it["id"], it["task"], it["owner"], it["status"] or t("f.no_status"),
                                       fmt(it["planned_first"]), fmt(it["current"]), str(it["reschedules"])])
                row.setData(0, Qt.UserRole, it["id"])
                if it["reschedules"] >= 2:
                    row.setForeground(6, QBrush(QColor("#A3260F")))
                    row.setFont(6, bold)
                if it["overdue"]:
                    row.setForeground(5, QBrush(QColor("#A3260F")))
                    row.setText(5, f"{fmt(it['current'])}  · {t('tag.Overdue')}")
                gi.addChild(row)
                if it["id"] == self.selected:
                    select_item = row
            gi.setExpanded(not g["complete"] or self.flt != "all" or bool(self.ctx.search)
                           or (select_item is not None and select_item.parent() is gi))
        hh = self.tree.header()
        hh.setSectionResizeMode(QHeaderView.Interactive)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        hh.setStretchLastSection(False)
        for col, width in ((0, 92), (2, 160), (3, 100), (4, 130), (5, 150), (6, 90)):
            self.tree.setColumnWidth(col, width)
        self.tree.blockSignals(False)
        if select_item is not None:
            self.tree.setCurrentItem(select_item)
        self._detail()

    def _pick(self, k):
        self.flt = k
        self.refresh()

    def _owner(self):
        self.owner = self.owner_cb.currentData()
        self.refresh()

    def _select(self):
        it = self.tree.currentItem()
        self.selected = it.data(0, Qt.UserRole) if it else None
        self._detail()

    def _detail(self):
        clear_layout(self.panel_lay)
        if not self.selected or self.ctx.repo.get_item(UAT_PLAN, self.selected) is None:
            self.panel_lay.addWidget(label(t("uat.pick"), "muted"))
            self.panel_lay.addStretch(1)
            return
        d = views.uat_detail(self.ctx.cfg, self.ctx.repo, self.selected, self.ctx.as_of)
        r = d["item"]
        c, cl = card()
        cl.addWidget(label(f"{r['id']} · {r.get('task_name') or ''}", "h2", wrap_text=True))
        for k, v in ((t("upd.status"), r.get("status") or t("f.no_status")), (t("col.owner"), r.get("responsible")),
                     (t("uat.start"), fmt(r.get("start_date"))), (t("uat.finish"), fmt(r.get("finish_date"))),
                     (t("uat.note"), r.get("note"))):
            row = QHBoxLayout()
            kl = label(k, "muted")
            kl.setFixedWidth(130)
            row.addWidget(kl, 0, Qt.AlignTop)
            row.addWidget(label(v or t("none"), wrap_text=True), 1)
            cl.addLayout(row)
        cl.addWidget(hline())
        cl.addWidget(label(t("uat.linked_tp"), "h2"))
        for a in d["linked_apps"]:
            line = QHBoxLayout()
            line.addWidget(label(f"{a['name']}  {a['executed']} / {a['total']}"), 1)
            if a["color"] in ("Red", "Amber", "Green", "Complete", "Missing"):
                line.addWidget(tag(a["color"]))
            cl.addLayout(line)
            cl.addWidget(bar(a["executed"], a["total"]))
        if not d["linked_apps"]:
            cl.addWidget(label(t("none"), "muted"))
        cl.addWidget(hline())
        cl.addWidget(label(t("uat.timeline"), "h2"))
        for h in d["timeline"]:
            src = t("uat.from_excel") if h["source"] == "import" else (h["reason"] or "")
            if h.get("root_cause"):
                src += f" [{h['root_cause']}]"
            what = t("upd.finish") if h["field"] == "finish_date" else t("upd.start")
            cl.addWidget(label(f"• {what}: {fmt(h['value'])} — {src}", wrap_text=True))
        if not d["timeline"]:
            cl.addWidget(label(t("none"), "muted"))
        cl.addWidget(hline())
        cl.addWidget(label(t("uat.defects"), "h2"))
        for x in d["defects"]:
            cl.addWidget(label(f"• {x['id']} {x['system']} — {x['detail']} ({x['status']})", wrap_text=True))
        if not d["defects"]:
            cl.addWidget(label(t("none"), "muted"))
        row = QHBoxLayout()
        hb = button(t("uat.history"))
        hb.clicked.connect(lambda: self.ctx.navigate("history", item=self.selected))
        ub = button(t("uat.update"), primary=True)
        ub.clicked.connect(self._update)
        if not self.ctx.can(auth.EDIT, UAT_PLAN, self.ctx.repo.get_item(UAT_PLAN, self.selected)):
            ub.setEnabled(False)                 # Part 5: QA - not their task
            ub.setToolTip(t("auth.not_yours"))
        row.addWidget(hb)
        row.addStretch(1)
        row.addWidget(ub)
        cl.addLayout(row)
        self.panel_lay.addWidget(c)
        self.panel_lay.addStretch(1)

    def _update(self):
        if not self.selected:
            return
        if UpdateDialog(self.ctx, UAT_PLAN, self.selected, self).exec():
            self.ctx.refresh_all()

    def _add(self):
        items = self.ctx.repo.load_items(UAT_PLAN).values()
        groups = sorted({(r["group_code"], r.get("group_title") or "") for r in items if r.get("active") and r.get("group_code")})
        dlg = AddTaskDialog(self.ctx, groups, self)
        if dlg.exec():
            self.selected = dlg.new_id
            self.ctx.refresh_all()

    def _import(self):
        try:
            res = SyncService(self.ctx.cfg, self.ctx.repo).pull(force=True)
        except Exception as exc:
            error(self, str(exc))
            return
        n = sum(s.updated + s.created + s.removed for s in res.sheets.values()) if res else 0
        info(self, t("uat.imported", n=n))
        self.ctx.refresh_all()
