"""App -> Excel write-back (Part 3). Touches only the 3 sheets and only the cells it must change.

Real writes go through Excel itself (COM): images, sensitivity label, formulas, formatting and the
other sheets stay exactly as they are. If the workbook is already open in Excel on this PC the
write goes into that open copy (only when it has no unsaved edits); otherwise a hidden Excel
instance opens it, saves and closes. If the workbook is open in Excel but Excel does not let the
app in, the write waits (WorkbookBusy -> retry) - it never saves underneath an open session.

`OpenpyxlWorkbook` exists for unit tests only: openpyxl drops images, so it must never be used on
the real workbook.

Each sheet gets a hidden 'App ID' column right after the last header of its table (approved);
rows are then found by App ID and every write is checked against it.
"""

from __future__ import annotations

import datetime as dt
import functools
import json
import logging
import os
import re
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import unquote

from openpyxl.utils import get_column_letter

from ..parsing.headers import HeaderError, header_label, map_headers
from .sheets import APP_ID_HEADER, APP_ID_KEY, DB_TO_SHEET_KEY, SheetDef

log = logging.getLogger(__name__)

EXCEL_DATE_FORMAT = "[$-409]dd-mmm-yyyy"      # English month names whatever the PC locale
_EXCEL_EPOCH = dt.date(1899, 12, 30)
MAX_HEADER_COLS = 80
ROW_FIELD = "(row)"        # write-queue entry meaning "create this item's row"


class WriterError(Exception):
    pass


class WorkbookLocked(WriterError):
    """File opened read-only / locked by another program. Retry later."""


class WorkbookBusy(WriterError):
    """Workbook open in Excel with unsaved edits. Retry later."""


@dataclass(frozen=True)
class CellWrite:
    sheet: str
    item_id: str
    field: str            # DB field name
    excel_row: int        # row the item had at the last import (checked against App ID)
    excel_value: str      # JSON {"kind": date|text|number|empty, "value": ...}


@dataclass
class WriteReport:
    written: list[tuple[str, str, str]] = field(default_factory=list)
    failed: list[tuple[tuple[str, str, str], str]] = field(default_factory=list)
    ids_written: int = 0
    id_column_added: list[str] = field(default_factory=list)


# --------------------------------------------------------------------------- core logic

# New-row formulas: field -> (template, columns used). Same shape as the PMO's existing rows (=E2/D2).
ROW_FORMULAS = {"pct_progress": ("={0}{r}/{1}{r}", "executed", "total_tc")}


def _table_end(labels: list[str], blank_first: bool) -> int:
    """Number of header columns in the table (stops at the first blank header)."""
    for i, lab in enumerate(labels):
        if not lab and not (i == 0 and blank_first) and i > 0:
            return i
    return len(labels)


def apply_writes(book, sheet_defs: dict[str, SheetDef], writes: list[CellWrite],
                 rows: dict[str, dict[str, int]]) -> WriteReport:
    """rows: sheet -> {item_id: excel_row} for all active items (from the last import)."""
    report = WriteReport()
    for name, sd in sheet_defs.items():
        sheet_writes = [w for w in writes if w.sheet == name]
        ws = book.sheet(name)
        labels = [header_label(v) for v in ws.header_values()]
        hmap = map_headers(labels, list(sd.fields), name, sd.blank_first_as)
        if APP_ID_KEY not in hmap.columns:
            col = _table_end(labels, bool(sd.blank_first_as)) + 1          # 1-based, right after table
            ws.insert_column(col)
            ws.set_text(1, col, APP_ID_HEADER)
            ws.hide_column(col)
            report.id_column_added.append(name)
            labels = [header_label(v) for v in ws.header_values()]
            hmap = map_headers(labels, list(sd.fields), name, sd.blank_first_as)
        id_col = hmap.columns[APP_ID_KEY] + 1

        # App IDs: every active item's row carries its ID (fixes blanks and copied rows).
        for item_id, row in sorted(rows.get(name, {}).items(), key=lambda kv: kv[1]):
            if str(ws.get(row, id_col) or "").strip() != item_id:
                ws.set_text(row, id_col, item_id)
                report.ids_written += 1

        cell_writes = [w for w in sheet_writes if w.field != ROW_FIELD]
        row_writes = sorted((w for w in sheet_writes if w.field == ROW_FIELD), key=lambda w: -w.excel_row)
        for w in cell_writes:
            key3 = (w.sheet, w.item_id, w.field)
            if str(ws.get(w.excel_row, id_col) or "").strip() != w.item_id:
                report.failed.append((key3, f"Row {w.excel_row} is not {w.item_id} any more"))
                continue
            sheet_key = DB_TO_SHEET_KEY.get(w.field, w.field)
            if sheet_key not in hmap.columns:
                report.failed.append((key3, f"Column for '{w.field}' not found in sheet {name}"))
                continue
            col = hmap.columns[sheet_key] + 1
            if ws.has_formula(w.excel_row, col):
                report.failed.append((key3, "Cell holds a formula; the app never overwrites formulas"))
                continue
            v = json.loads(w.excel_value)
            kind, value = v["kind"], v.get("value")
            if kind == "date":
                ws.set_date(w.excel_row, col, dt.date.fromisoformat(value))
            elif kind == "number":
                ws.set_number(w.excel_row, col, value)
            elif kind == "text":
                ws.set_text(w.excel_row, col, str(value))
            elif kind == "empty":
                ws.clear(w.excel_row, col)
            else:
                report.failed.append((key3, f"Unknown value kind {kind!r}"))
                continue
            report.written.append(key3)

        # New rows last, bottom-up, so the rows checked above are not shifted under us.
        for w in row_writes:
            key3 = (w.sheet, w.item_id, w.field)
            spec = json.loads(w.excel_value)
            target = w.excel_row + 1
            if spec.get("insert"):
                gh = spec.get("group_header")            # multi-project: first task of a NEW group
                ws.insert_row(target, below=target == 2)
                if gh:
                    ws.insert_row(target, below=target == 2)
                    for key, text in (("group", gh["group"]), ("task_name", gh["title"])):
                        if key in hmap.columns:
                            ws.set_text(target, hmap.columns[key] + 1, text)
                    target += 1
            else:
                skip = {APP_ID_KEY, "no"}
                data_cols = [c + 1 for k, c in hmap.columns.items() if k not in skip]
                last = ws.last_row()
                while target <= last and any(str(ws.get(target, c) or "").strip() for c in data_cols):
                    target += 1
            ws.set_text(target, id_col, w.item_id)
            for key, ev in spec.get("values", {}).items():
                if key not in hmap.columns:
                    continue
                v = json.loads(ev)
                col = hmap.columns[key] + 1
                if v["kind"] == "date":
                    ws.set_date(target, col, dt.date.fromisoformat(v["value"]))
                elif v["kind"] == "number":
                    ws.set_number(target, col, v["value"])
                elif v["kind"] == "text":
                    ws.set_text(target, col, str(v["value"]))
            for key in spec.get("formulas", ()):         # e.g. %Progress = Executed / Total, like the old rows
                f = ROW_FORMULAS.get(key)
                if f and key in hmap.columns and all(k in hmap.columns for k in f[1:]):
                    letters = [get_column_letter(hmap.columns[k] + 1) for k in f[1:]]
                    ws.set_formula(target, hmap.columns[key] + 1, f[0].format(*letters, r=target))
            report.written.append(key3)
    return report


def write_workbook(opener, path: Path, sheet_defs, writes, rows) -> WriteReport:
    book = opener(path)
    try:
        report = apply_writes(book, sheet_defs, writes, rows)
        if report.written or report.ids_written or report.id_column_added:
            book.save()
        return report
    except (WriterError, HeaderError):
        raise
    except Exception as exc:
        log.warning("Excel refused the write (retried later)", exc_info=True)
        # COM error from the Excel the user has open (busy uploading, a cell being typed in, a dialog):
        # retried later. Cells already written are harmless: the next read sees the app's value there.
        raise WorkbookBusy(f"Excel refused the change for now: {exc}") from exc
    finally:
        book.close()


# --------------------------------------------------------------------------- Excel (COM)

def _same_workbook(wb, path: Path) -> bool:
    full = str(wb.FullName)
    if not full.lower().startswith("http"):
        return os.path.normcase(os.path.abspath(full)) == os.path.normcase(os.path.abspath(path))
    if str(wb.Name).lower() != path.name.lower():
        return False
    for env in ("OneDriveCommercial", "OneDrive"):
        root = os.environ.get(env)
        if not root:
            continue
        try:
            rel = path.resolve().relative_to(Path(root).resolve()).as_posix()
        except ValueError:
            continue
        return unquote(full).replace("\\", "/").lower().endswith("/" + rel.lower())
    return False


def title_stem(title: str) -> str:
    """'UAT Planning.xlsx [Read-Only] - Saved - Excel' -> 'uat planning'."""
    t = re.sub(r"\s+-\s+Excel$", "", title.strip(), flags=re.I)
    t = t.split(" - ")[0]
    t = re.sub(r"\[[^\]]*\]", "", t).strip()
    t = re.sub(r"\.(xlsx|xlsm|xls)$", "", t, flags=re.I).strip()
    return t.casefold()


def excel_windows_for(path: Path) -> list[int]:
    """Top-level Excel windows (class XLMAIN) showing this workbook."""
    try:
        import win32gui
    except ImportError:                                  # pragma: no cover
        return []
    stem = Path(path).stem.casefold()
    found: list[int] = []

    def visit(h, _):
        try:
            if win32gui.GetClassName(h) == "XLMAIN" and title_stem(win32gui.GetWindowText(h)) == stem:
                found.append(h)
        except Exception:
            log.info("Could not list Excel windows", exc_info=True)
        return True

    win32gui.EnumWindows(visit, None)
    return found


def _attach_via_window(hwnds: list[int], path: Path):
    """Workbook object behind an Excel window, or None if Excel does not answer."""
    try:
        import pythoncom
        import win32com.client
        import win32con
        import win32gui
        for h in hwnds:
            desk = win32gui.FindWindowEx(h, 0, "XLDESK", None)
            w7 = win32gui.FindWindowEx(desk, 0, "EXCEL7", None) if desk else 0
            if not w7:
                continue
            lres = win32gui.SendMessage(w7, win32con.WM_GETOBJECT, 0, 0xFFFFFFF0)   # OBJID_NATIVEOM
            win = win32com.client.Dispatch(pythoncom.ObjectFromLresult(lres, pythoncom.IID_IDispatch, 0))
            wb = win.Parent
            if _same_workbook(wb, path):
                return wb
    except Exception:
        log.info("Could not attach to the open Excel window", exc_info=True)
        return None
    return None


# --------------------------------------------------------------------------- Excel busy (Hardening D)
# Excel answers COM calls with these while it is busy (a cell being edited, a dialog open, AutoSave upload).
# The call is retried with a growing wait, up to COM_RETRY_BUDGET_S seconds; then WorkbookBusy -> the write
# goes back to the normal retry queue (sync retry_minutes).
RPC_E_CALL_REJECTED = 0x80010001
RPC_E_SERVERCALL_RETRYLATER = 0x8001010A
BUSY_HRESULTS = frozenset({RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER})
COM_RETRY_BUDGET_S = 30.0
COM_RETRY_DELAYS = (0.5, 1.0, 2.0, 4.0, 8.0)       # then 8 s each time until the budget is used
_sleep = time.sleep                                  # tests replace this
_budget = threading.local()


def _hresult(exc: BaseException):
    code = getattr(exc, "hresult", None)
    if code is None and getattr(exc, "args", None):
        code = exc.args[0]
    return code & 0xFFFFFFFF if isinstance(code, int) else None


def is_busy_error(exc: BaseException) -> bool:
    return _hresult(exc) in BUSY_HRESULTS


@contextmanager
def com_retry_budget(seconds: float):
    """Shorter (or longer) busy-retry budget for the COM calls made inside, on this thread only."""
    old = getattr(_budget, "s", None)
    _budget.s = seconds
    try:
        yield
    finally:
        _budget.s = old


def com_call(fn, *args, **kwargs):
    """Run one COM call; retry while Excel says it is busy, then raise WorkbookBusy."""
    budget = getattr(_budget, "s", None)
    budget = COM_RETRY_BUDGET_S if budget is None else budget
    waited, attempt = 0.0, 0
    while True:
        try:
            return fn(*args, **kwargs)
        except Exception as exc:
            if not is_busy_error(exc):
                raise
            delay = COM_RETRY_DELAYS[min(attempt, len(COM_RETRY_DELAYS) - 1)]
            if waited + delay > budget:
                log.warning("Excel stayed busy for %.1f s; the change is queued for later", waited, exc_info=True)
                raise WorkbookBusy(f"Excel is busy (tried for {waited:.0f} s); the change is written later") \
                    from exc
            log.info("Excel busy (%#x); retry in %.1f s", _hresult(exc), delay)
            _sleep(delay)
            waited += delay
            attempt += 1


def _busy_retry(method):
    @functools.wraps(method)
    def wrapper(*args, **kwargs):
        return com_call(method, *args, **kwargs)
    return wrapper


class ComSheet:
    """Each method is one or a few COM calls; it is retried while Excel is busy."""

    def __init__(self, ws):
        self.ws = ws

    @_busy_retry
    def header_values(self):
        vals = self.ws.Range(self.ws.Cells(1, 1), self.ws.Cells(1, MAX_HEADER_COLS)).Value
        return list(vals[0])

    @_busy_retry
    def insert_column(self, col):
        self.ws.Columns(col).Insert()

    @_busy_retry
    def insert_row(self, row, below: bool = False):
        # below=True: take the format of the row underneath (row 2 of a new sheet sits under the header,
        # whose look must not be copied). -4121 = xlShiftDown, 1 = xlFormatFromRightOrBelow.
        if below:
            self.ws.Rows(row).Insert(-4121, 1)
        else:
            self.ws.Rows(row).Insert()

    @_busy_retry
    def set_formula(self, row, col, formula):
        self.ws.Cells(row, col).Formula = formula

    @_busy_retry
    def last_row(self):
        ur = self.ws.UsedRange
        return int(ur.Row) + int(ur.Rows.Count) - 1

    @_busy_retry
    def hide_column(self, col):
        self.ws.Columns(col).Hidden = True

    @_busy_retry
    def get(self, row, col):
        return self.ws.Cells(row, col).Value

    @_busy_retry
    def has_formula(self, row, col):
        return bool(self.ws.Cells(row, col).HasFormula)

    @_busy_retry
    def set_date(self, row, col, d):
        c = self.ws.Cells(row, col)
        c.NumberFormat = EXCEL_DATE_FORMAT
        c.Value2 = float((d - _EXCEL_EPOCH).days)

    @_busy_retry
    def set_number(self, row, col, n):
        self.ws.Cells(row, col).Value2 = n

    @_busy_retry
    def set_text(self, row, col, text):
        self.ws.Cells(row, col).Value = "'" + text if text else ""     # ' = keep as text, not a date

    @_busy_retry
    def clear(self, row, col):
        # Value = None, not ClearContents(): the Excel the PMO keeps open (AutoSave on) refuses
        # ClearContents through COM (0x800A03EC) but accepts this. Same result; formatting is kept.
        self.ws.Cells(row, col).Value = None


def _autosave_on(wb) -> bool:
    try:
        return bool(wb.AutoSaveOn)
    except Exception:                                    # older Excel / local file: no AutoSave
        log.info("AutoSave state not available (older Excel / local file)", exc_info=True)
        return False


class ComWorkbook:
    def __init__(self, path: Path):
        try:
            import pythoncom
            import pywintypes
            import win32com.client
        except ImportError as exc:                     # pragma: no cover
            raise WriterError("pywin32 is not installed") from exc
        self._com_error = pywintypes.com_error
        pythoncom.CoInitialize()
        self.path = Path(path)
        self.app, self.wb, self.ours = None, None, False
        # Is the workbook open in Excel on this PC? A OneDrive file opened in Excel is NOT locked on
        # disk, so a second Excel could save over the open session. The window title tells us it is
        # open; we then write into that open copy, or wait (busy) if Excel does not let us in.
        windows = excel_windows_for(self.path)
        if windows:
            self.wb = _attach_via_window(windows, self.path)
            if self.wb is None:
                raise WorkbookBusy("Workbook is open in Excel on this PC; changes are written after it "
                                   "is closed (or when Excel lets the app in)")
            self.app = com_call(lambda: self.wb.Application)
        if self.wb is not None:
            if com_call(lambda: self.wb.ReadOnly):
                raise WorkbookLocked("Workbook is open read-only in Excel")
            # Unsaved changes: wait, unless AutoSave is on (approved 09-Nov-2026). With AutoSave Excel uploads
            # every edit anyway, and volatile formulas (TODAY()) leave Saved=False forever after opening.
            # A cell being typed in still blocks us: Excel refuses COM calls then, and the write is retried.
            if not com_call(lambda: self.wb.Saved) and not _autosave_on(self.wb):
                raise WorkbookBusy("Workbook is open in Excel with unsaved changes (AutoSave off)")
            return
        try:
            self.app = win32com.client.DispatchEx("Excel.Application")
            self.ours = True
            self.app.Visible = False
            self.app.DisplayAlerts = False
            self.app.ScreenUpdating = False
            self.wb = com_call(self.app.Workbooks.Open, str(self.path), 0, False)   # UpdateLinks=0, ReadOnly=False
        except WorkbookBusy:
            self.close()
            raise
        except self._com_error as exc:
            self.close()
            raise WorkbookLocked(f"Excel could not open the workbook: {exc}") from exc
        if com_call(lambda: self.wb.ReadOnly):
            self.close()
            raise WorkbookLocked("Workbook is locked by another program (opened read-only)")

    def sheet(self, name):
        try:
            return ComSheet(com_call(self.wb.Worksheets, name))
        except self._com_error as exc:
            raise HeaderError(f"Sheet '{name}' not found") from exc

    def save(self):
        try:
            com_call(self.wb.Save)
        except self._com_error as exc:
            raise WorkbookLocked(f"Save failed: {exc}") from exc

    def close(self):
        if self.ours:
            try:
                if self.wb is not None:
                    self.wb.Close(False)
            except Exception:
                log.warning("Closing the workbook failed", exc_info=True)
            try:
                if self.app is not None:
                    self.app.Quit()
            except Exception:
                log.warning("Quitting Excel failed", exc_info=True)
        self.wb = self.app = None


# --------------------------------------------------------------------------- openpyxl (tests only)

class OpenpyxlSheet:
    def __init__(self, ws):
        self.ws = ws

    def header_values(self):
        return [self.ws.cell(1, c).value for c in range(1, MAX_HEADER_COLS + 1)]

    def insert_column(self, col):
        self.ws.insert_cols(col)

    def insert_row(self, row, below: bool = False):
        self.ws.insert_rows(row)

    def set_formula(self, row, col, formula):
        self.ws.cell(row, col).value = formula

    def last_row(self):
        return self.ws.max_row

    def hide_column(self, col):
        from openpyxl.utils import get_column_letter
        self.ws.column_dimensions[get_column_letter(col)].hidden = True

    def get(self, row, col):
        return self.ws.cell(row, col).value

    def has_formula(self, row, col):
        v = self.ws.cell(row, col).value
        return isinstance(v, str) and v.startswith("=")

    def set_date(self, row, col, d):
        c = self.ws.cell(row, col)
        c.value = dt.datetime(d.year, d.month, d.day)
        c.number_format = "dd-mmm-yyyy"

    def set_number(self, row, col, n):
        self.ws.cell(row, col).value = n

    def set_text(self, row, col, text):
        c = self.ws.cell(row, col)
        c.value = text
        c.number_format = "@"

    def clear(self, row, col):
        self.ws.cell(row, col).value = None


class OpenpyxlWorkbook:
    """TESTS ONLY - loses images/drawings. Never use on the real workbook."""

    def __init__(self, path: Path):
        import openpyxl
        self.path = Path(path)
        self.wb = openpyxl.load_workbook(self.path)

    def sheet(self, name):
        return OpenpyxlSheet(self.wb[name])

    def save(self):
        self.wb.save(self.path)

    def close(self):
        self.wb.close()
