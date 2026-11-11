from __future__ import annotations

from PySide6.QtCore import QPoint, QRect, QSize, Qt
from PySide6.QtWidgets import QFrame, QLayout, QPushButton, QScrollArea, QVBoxLayout, QWidget

from ..widgets import button


def clear_layout(lay: QLayout) -> None:
    while lay.count():
        it = lay.takeAt(0)
        w = it.widget()
        if w is not None:
            w.hide()
            w.setParent(None)
            w.deleteLater()
        elif it.layout() is not None:
            clear_layout(it.layout())
            it.layout().setParent(None)
            it.layout().deleteLater()


class Screen(QWidget):
    """A page. refresh() rebuilds texts + data (also on TH/EN switch); set_search() filters."""

    key = ""

    def __init__(self, ctx, scroll: bool = True):
        super().__init__()
        self.ctx = ctx
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.body = QWidget()
        self.body.setObjectName("page")
        self.lay = QVBoxLayout(self.body)
        self.lay.setContentsMargins(24, 24, 24, 24)
        self.lay.setSpacing(18)
        if scroll:
            sa = QScrollArea()
            sa.setWidgetResizable(True)
            sa.setFrameShape(QFrame.NoFrame)
            sa.setWidget(self.body)
            outer.addWidget(sa)
        else:
            outer.addWidget(self.body)

    def refresh(self):
        pass

    def set_search(self, text: str):
        self.refresh()


class FlowLayout(QLayout):
    """Left-to-right layout that wraps to the next line (chips never force the page wider)."""

    def __init__(self, parent=None, spacing=8):
        super().__init__(parent)
        self._items = []
        self.setContentsMargins(0, 0, 0, 0)
        self.setSpacing(spacing)

    def addItem(self, item):
        self._items.append(item)

    def count(self):
        return len(self._items)

    def itemAt(self, i):
        return self._items[i] if 0 <= i < len(self._items) else None

    def takeAt(self, i):
        return self._items.pop(i) if 0 <= i < len(self._items) else None

    def expandingDirections(self):
        return Qt.Orientations(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        return self._do(QRect(0, 0, width, 0), True)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._do(rect, False)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()
        for it in self._items:
            size = size.expandedTo(it.minimumSize())
        return size

    def _do(self, rect, test):
        x, y, line = rect.x(), rect.y(), 0
        sp = self.spacing()
        for it in self._items:
            w = it.sizeHint()
            if x + w.width() > rect.right() + 1 and line > 0:
                x, y, line = rect.x(), y + line + sp, 0
            if not test:
                it.setGeometry(QRect(QPoint(x, y), w))
            x += w.width() + sp
            line = max(line, w.height())
        return y + line - rect.y()


class ChipRow(QWidget):
    """Filter chips (one checked), wrapping. on_pick(key)."""

    def __init__(self, on_pick):
        super().__init__()
        self.on_pick = on_pick
        self.row = FlowLayout(self)
        self.buttons: dict[str, QPushButton] = {}

    def set_chips(self, chips: list[tuple[str, str]], current: str):
        clear_layout(self.row)
        self.buttons = {}
        for key, text in chips:
            b = button(text, obj="chip")
            b.setCheckable(True)
            b.setChecked(key == current)
            b.setAccessibleName(text)
            b.clicked.connect(lambda _=False, k=key: self.on_pick(k))
            self.row.addWidget(b)
            self.buttons[key] = b
        self.updateGeometry()
