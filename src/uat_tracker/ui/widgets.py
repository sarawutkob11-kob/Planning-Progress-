"""Theme (spec §9.1) and small shared widgets."""

from __future__ import annotations

import datetime as dt

from PySide6.QtCore import QDate, QLocale, Qt, Signal
from PySide6.QtWidgets import (QCalendarWidget, QDialog, QFrame, QHBoxLayout, QLabel, QLineEdit, QProgressBar,
                               QPushButton, QSizePolicy, QToolButton, QVBoxLayout, QWidget)

from ..parsing.dates import fmt
from .i18n import t

NAVY, BG, BORDER, PRIMARY = "#14213D", "#F4F6F9", "#E1E6EE", "#1F4FBF"
MUTED = "#5A6577"
TAG_COLORS = {   # (text, background) - always used with words, never color alone
    "Red": ("#A3260F", "#FDE7E3"), "Overdue": ("#A3260F", "#FDE7E3"), "mismatch": ("#A3260F", "#FDE7E3"),
    "not_filled": ("#A3260F", "#FDE7E3"),
    "Amber": ("#7A4E00", "#FFF1D6"), "No target": ("#7A4E00", "#FFF1D6"),
    "Green": ("#1E6B3E", "#E3F2E8"), "Complete": ("#1E6B3E", "#E3F2E8"),
    "Today": ("#1F4FBF", "#E5ECFA"), "Missing": ("#3D4656", "#E9EDF3"), "Missing data": ("#3D4656", "#E9EDF3"),
    "neutral": ("#3D4656", "#E9EDF3"),
}

QSS = f"""
* {{ font-family: "IBM Plex Sans Thai", Tahoma, sans-serif; font-size: 10pt; color: {NAVY}; }}
QMainWindow, QWidget#page, QScrollArea, QScrollArea > QWidget > QWidget#page {{ background: {BG}; }}
QFrame#nav {{ background: {NAVY}; }}
QFrame#nav QLabel {{ color: #FFFFFF; }}
QFrame#nav QLabel#sub {{ color: #B8C3D9; font-size: 9pt; }}
QFrame#nav QPushButton {{ color: #D5DCEA; background: transparent; border: 0; border-radius: 10px;
    text-align: left; padding: 0 12px; min-height: 44px; }}
QFrame#nav QPushButton:hover {{ background: #22345C; }}
QFrame#nav QPushButton:checked {{ background: {PRIMARY}; color: #FFFFFF; font-weight: 600; }}
QFrame#nav QPushButton:focus {{ border: 2px solid #9DB6F2; }}
QFrame#topbar {{ background: #FFFFFF; border-bottom: 1px solid {BORDER}; }}
QFrame#card {{ background: #FFFFFF; border: 1px solid {BORDER}; border-radius: 12px; }}
QFrame#warn {{ background: #FFF1D6; border: 1px solid #F0D9A8; border-radius: 12px; }}
QFrame#warn QLabel {{ color: #7A4E00; }}
QLabel#h1 {{ font-size: 17pt; font-weight: 700; }}
QLabel#h2 {{ font-size: 12pt; font-weight: 700; }}
QLabel#muted {{ color: {MUTED}; }}
QLabel#kpi {{ font-size: 22pt; font-weight: 700; }}
QPushButton {{ min-height: 44px; padding: 0 16px; border-radius: 10px; border: 1px solid #C9D1DE;
    background: #FFFFFF; }}
QPushButton:hover {{ background: #F1F4F9; }}
QPushButton#primary {{ background: {PRIMARY}; color: #FFFFFF; border: 0; font-weight: 600; }}
QPushButton#primary:hover {{ background: #1A43A3; }}
QPushButton#chip {{ border-radius: 22px; font-weight: 600; }}
QPushButton#chip:checked {{ background: {PRIMARY}; color: #FFFFFF; border-color: {PRIMARY}; }}
QPushButton:disabled, QPushButton#primary:disabled {{ background: #EEF1F6; color: {MUTED};
    border: 1px solid #D5DBE5; }}
QPushButton:focus, QToolButton:focus {{ border: 2px solid {PRIMARY}; }}
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QTimeEdit, QPlainTextEdit {{ min-height: 44px;
    border: 1px solid {BORDER}; border-radius: 10px; padding: 0 10px; background: #F9FAFC; }}
QPlainTextEdit {{ padding: 6px 10px; }}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus, QTimeEdit:focus,
QPlainTextEdit:focus {{ border: 2px solid {PRIMARY}; }}
QLineEdit[readOnly="true"] {{ background: #FFFFFF; }}
QToolButton {{ min-height: 44px; min-width: 44px; border: 1px solid #C9D1DE; border-radius: 10px; background: #FFFFFF; }}
QTableWidget, QTreeWidget, QListWidget {{ background: #FFFFFF; border: 1px solid {BORDER}; border-radius: 10px;
    gridline-color: #EEF1F6; alternate-background-color: #FAFBFD; }}
QTableWidget::item, QTreeWidget::item {{ padding: 6px; }}
QTableWidget::item:selected, QTreeWidget::item:selected, QListWidget::item:selected {{
    background: #E5ECFA; color: {NAVY}; }}
QHeaderView::section {{ background: {BG}; border: 0; border-bottom: 1px solid {BORDER}; padding: 8px;
    font-weight: 600; color: #3D4656; }}
QProgressBar {{ border: 0; border-radius: 4px; background: #E9EDF3; max-height: 8px; min-height: 8px; }}
QProgressBar::chunk {{ border-radius: 4px; background: {PRIMARY}; }}
QCheckBox {{ min-height: 44px; }}
QTabBar::tab {{ min-height: 40px; padding: 0 16px; }}
"""


def tag(kind: str, text: str | None = None) -> QLabel:
    fg, bg = TAG_COLORS.get(kind, TAG_COLORS["neutral"])
    lab = QLabel(text if text is not None else t(f"tag.{kind}"))
    lab.setStyleSheet(f"color:{fg}; background:{bg}; border-radius:10px; padding:3px 10px; font-weight:700;"
                      " font-size: 9pt;")
    lab.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Maximum)
    return lab


def wrap(widget: QWidget, align=Qt.AlignLeft | Qt.AlignVCenter) -> QWidget:
    """Put a small widget in a table cell without stretching it."""
    w = QWidget()
    lay = QHBoxLayout(w)
    lay.setContentsMargins(6, 2, 6, 2)
    lay.addWidget(widget, 0, align)
    lay.addStretch(1)
    return w


def card(*, margins=(20, 18, 20, 18), spacing=12) -> tuple[QFrame, QVBoxLayout]:
    f = QFrame()
    f.setObjectName("card")
    lay = QVBoxLayout(f)
    lay.setContentsMargins(*margins)
    lay.setSpacing(spacing)
    return f, lay


def label(text: str = "", obj: str | None = None, wrap_text: bool = False) -> QLabel:
    lab = QLabel(text)
    if obj:
        lab.setObjectName(obj)
    lab.setWordWrap(wrap_text)
    return lab


def bar(value: int, total: int, color: str = PRIMARY) -> QProgressBar:
    b = QProgressBar()
    b.setRange(0, max(total, 1))
    b.setValue(min(value, max(total, 1)))
    b.setTextVisible(False)
    b.setStyleSheet(f"QProgressBar::chunk {{ background: {color}; border-radius: 4px; }}")
    b.setAccessibleName(f"{value} / {total}")
    return b


def button(text: str, primary: bool = False, obj: str | None = None) -> QPushButton:
    b = QPushButton(text)
    if primary:
        b.setObjectName("primary")
    elif obj:
        b.setObjectName(obj)
    b.setCursor(Qt.PointingHandCursor)
    return b


class DatePicker(QWidget):
    """Date input by picker ONLY: the text box is read-only, so typed text is impossible.
    Shows DD-Mon-YYYY; value() is ISO or None."""

    changed = Signal()

    def __init__(self, value: str | None = None, clearable: bool = True, parent=None):
        super().__init__(parent)
        self._value: str | None = None
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
        self.edit = QLineEdit()
        self.edit.setReadOnly(True)
        self.edit.setFocusPolicy(Qt.ClickFocus)
        self.btn = QToolButton()
        self.btn.setText("📅")
        self.btn.setToolTip(t("pick_date"))
        self.btn.setAccessibleName(t("pick_date"))
        self.btn.clicked.connect(self.open_calendar)
        lay.addWidget(self.edit, 1)
        lay.addWidget(self.btn)
        self.clear_btn = None
        if clearable:
            self.clear_btn = QToolButton()
            self.clear_btn.setText("✕")
            self.clear_btn.setToolTip(t("clear"))
            self.clear_btn.setAccessibleName(t("clear"))
            self.clear_btn.clicked.connect(lambda: self.setValue(None, emit=True))
            lay.addWidget(self.clear_btn)
        self.setValue(value)

    def mousePressEvent(self, e):
        self.open_calendar()

    def open_calendar(self):
        dlg = QDialog(self)
        dlg.setWindowTitle(t("pick_date"))
        cal = QCalendarWidget(dlg)
        cal.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
        cal.setGridVisible(True)
        if self._value:
            d = dt.date.fromisoformat(self._value)
            cal.setSelectedDate(QDate(d.year, d.month, d.day))
        lay = QVBoxLayout(dlg)
        lay.addWidget(cal)
        okb = button(t("ok"), primary=True)
        lay.addWidget(okb)

        def accept():
            q = cal.selectedDate()
            self.setValue(dt.date(q.year(), q.month(), q.day()).isoformat(), emit=True)
            dlg.accept()

        okb.clicked.connect(accept)
        cal.activated.connect(lambda _d: accept())
        dlg.exec()

    def value(self) -> str | None:
        return self._value

    def setValue(self, iso: str | None, emit: bool = False):
        self._value = iso or None
        self.edit.setText(fmt(iso) if iso else t("none"))
        if emit:
            self.changed.emit()


def hline() -> QFrame:
    f = QFrame()
    f.setFrameShape(QFrame.HLine)
    f.setStyleSheet(f"color: {BORDER};")
    return f
