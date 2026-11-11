"""Screen 5: Defect List - Board (drag a card = status change, audit-logged) / Table."""

from __future__ import annotations

from PySide6.QtCore import QSize, Qt
from PySide6.QtWidgets import (QAbstractItemView, QHBoxLayout, QHeaderView, QInputDialog, QListWidget,
                               QListWidgetItem, QMenu, QStackedWidget, QTableWidget, QTableWidgetItem, QVBoxLayout,
                               QWidget)

from ...services import auth
from ...services.auth import PermissionDenied
from ...excel.sheets import DEFECT_LIST
from ...parsing.dates import fmt
from ...services import views
from ...services.edits import EditError, edit_item
from ..dialogs import DefectDialog, error
from ..i18n import t
from ..widgets import button, card, label, tag, wrap
from .base import ChipRow, Screen, clear_layout


class _Column(QListWidget):
    """One board column. Cards dropped here change the defect's status."""

    def __init__(self, screen, column: str):
        super().__init__()
        self.screen, self.column = screen, column
        self.setDragDropMode(QAbstractItemView.DragDrop)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSpacing(6)
        self.setMinimumWidth(200)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.itemDoubleClicked.connect(lambda it: screen.edit(it.data(Qt.UserRole)))

    def relayout(self):
        width = max(self.viewport().width() - 14, 160)
        for i in range(self.count()):
            it = self.item(i)
            w = self.itemWidget(it)
            if w is None:
                continue
            w.setFixedWidth(width)
            lay = w.layout()
            h = lay.totalHeightForWidth(width) if lay.hasHeightForWidth() else w.sizeHint().height()
            it.setSizeHint(QSize(width, max(h, w.minimumSizeHint().height()) + 4))

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.relayout()

    def dropEvent(self, e):
        src = e.source()
        if isinstance(src, _Column) and src is not self and src.currentItem() is not None:
            item_id = src.currentItem().data(Qt.UserRole)
            e.ignore()                                   # the refresh redraws the board from the DB
            self.screen.move(item_id, self.column)
            return
        e.ignore()


class DefectsScreen(Screen):
    key = "defects"

    def __init__(self, ctx):
        super().__init__(ctx, scroll=False)
        self.mode, self.selected = "board", None
        top = QHBoxLayout()
        self.title = label("", "h1")
        top.addWidget(self.title)
        top.addStretch(1)
        self.toggle = ChipRow(self._mode)
        top.addWidget(self.toggle)
        self.add_btn = button("", primary=True)
        self.add_btn.clicked.connect(self._add)
        top.addWidget(self.add_btn)
        self.lay.addLayout(top)
        self.stack = QStackedWidget()
        self.board = QWidget()
        self.board_lay = QHBoxLayout(self.board)
        self.board_lay.setContentsMargins(0, 0, 0, 0)
        self.board_lay.setSpacing(12)
        self.table = QTableWidget()
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.cellDoubleClicked.connect(lambda r, c: self.edit(self.table.item(r, 0).data(Qt.UserRole)))
        self.stack.addWidget(self.board)
        self.stack.addWidget(self.table)
        self.lay.addWidget(self.stack, 1)

    def show_item(self, item_id=None, **_):
        self.selected = item_id
        self.refresh()
        if item_id:
            self.edit(item_id)

    def refresh(self):
        self.title.setText(t("nav.defects"))
        self.add_btn.setText(t("def.add"))
        self.add_btn.setVisible(self.ctx.can(auth.ADD_DEFECT))
        self._items = self.ctx.repo.load_items(DEFECT_LIST)
        self.toggle.set_chips([("board", t("def.board")), ("table", t("def.table"))], self.mode)
        rows = views.defects(self.ctx.repo, self.ctx.as_of, self.ctx.search)
        self._board(rows)
        self._table(rows)
        self.stack.setCurrentIndex(0 if self.mode == "board" else 1)

    def _card(self, r) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(10, 8, 10, 8)
        v.setSpacing(4)
        v.addWidget(label(f"#{r['no']} · {r['system']}", "h2", wrap_text=True))
        tags = QHBoxLayout()
        tags.setSpacing(6)
        if r["severity"]:
            tags.addWidget(tag("Red" if r["severe"] else "neutral", f"Sev {r['severity']}"))
        if r["overdue"]:
            tags.addWidget(tag("Overdue"))
        if r["no_target"]:
            tags.addWidget(tag("No target"))
        tags.addStretch(1)
        v.addLayout(tags)
        detail = r["detail"] if len(r["detail"]) <= 90 else r["detail"][:87] + "…"
        v.addWidget(label(detail, wrap_text=True))
        v.addWidget(label(r["owner"] or t("none"), "muted", wrap_text=True))
        if r["days_open"] is not None:
            v.addWidget(label(t("def.days_open", n=r["days_open"]), "muted"))
        mv = button(t("def.move"))
        mv.setAccessibleName(f"{t('def.move')} {r['id']}")
        menu = QMenu(mv)
        for col in views.DEFECT_COLUMNS:
            if col != r["column"]:
                menu.addAction(col, lambda c=col, i=r["id"]: self.move(i, c))
        mv.setMenu(menu)
        if not self.ctx.can(auth.EDIT, DEFECT_LIST, self._items.get(r["id"])):
            mv.setEnabled(False)                 # Part 5: QA - not assigned to them
            mv.setToolTip(t("auth.not_yours"))
        v.addWidget(mv)
        w.setToolTip(r["detail"])
        return w

    def _board(self, rows):
        clear_layout(self.board_lay)
        for col in views.DEFECT_COLUMNS:
            items = [r for r in rows if r["column"] == col]
            c, cl = card(margins=(10, 10, 10, 10), spacing=8)
            cl.addWidget(label(f"{col} · {len(items)}", "h2"))
            lst = _Column(self, col)
            for r in items:
                li = QListWidgetItem()
                li.setData(Qt.UserRole, r["id"])
                w = self._card(r)
                lst.addItem(li)
                lst.setItemWidget(li, w)
            lst.relayout()
            cl.addWidget(lst, 1)
            self.board_lay.addWidget(c, 1)

    def _table(self, rows):
        cols = ["#", t("def.system"), t("def.detail"), t("upd.status"), t("col.sev"), t("col.owner"),
                t("col.target"), t("col.days"), t("col.tag")]
        tb = self.table
        tb.clearContents()
        tb.setColumnCount(len(cols))
        tb.setHorizontalHeaderLabels(cols)
        tb.setRowCount(len(rows))
        for i, r in enumerate(rows):
            it = QTableWidgetItem(str(r["no"]))
            it.setData(Qt.UserRole, r["id"])
            tb.setItem(i, 0, it)
            for j, v in enumerate((r["system"], r["detail"], r["status"], r["severity"], r["owner"],
                                   fmt(r["target"]), "" if r["days_open"] is None else str(r["days_open"])), start=1):
                tb.setItem(i, j, QTableWidgetItem(v))
            if r["overdue"] or r["no_target"]:
                tb.setCellWidget(i, 8, wrap(tag("Overdue" if r["overdue"] else "No target")))
            tb.setRowHeight(i, 48)
        hh = tb.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.Stretch)

    def _mode(self, m):
        self.mode = m
        self.refresh()

    def move(self, item_id: str, column: str):
        status = column
        if column == "Close/Rejected":
            choice, ok = QInputDialog.getItem(self, t("def.move"), t("def.close_or_reject"), ["Close", "Rejected"], 0, False)
            if not ok:
                self.refresh()
                return
            status = choice
        try:
            edit_item(self.ctx.cfg, self.ctx.repo, DEFECT_LIST, item_id, {"status": status}, actor=self.ctx.actor)
        except (EditError, PermissionDenied) as exc:
            error(self, str(exc))
        self.ctx.refresh_all()

    def edit(self, item_id):
        if item_id and DefectDialog(self.ctx, item_id, self).exec():
            self.ctx.refresh_all()

    def _add(self):
        if DefectDialog(self.ctx, None, self).exec():
            self.ctx.refresh_all()
