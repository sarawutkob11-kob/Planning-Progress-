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
from collections.abc import Sequence
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol
from urllib.parse import unquote

from openpyxl.utils import get_column_letter

from ..parsing.headers import HeaderError, header_label, map_headers
from .sheets import APP_ID_HEADER, APP_ID_KEY, DB_TO_SHEET_KEY, SheetDef, is_item_row

log = logging.getLogger(__name__)

EXCEL_DATE_FORMAT = "[$-409]dd-mmm-yyyy"      # English month names whatever the PC locale
_EXCEL_EPOCH = dt.date(1899, 12, 30)
MAX_HEADER_COLS = 80
# Range.Find arguments (Excel constants)
# LookIn: xlFormulas. Finds cells in rows hidden by hand, collapsed groups and hidden columns (xlValues does
# not), but is NOT reliable under a filter (sheet AutoFilter or a filtered Table, both set ws.FilterMode)
# nor on a protected sheet with FormulaHidden cells -> last_data_row falls back to UsedRange there.
XL_FORMULAS = -4123
XL_PART = 2                # LookAt: xlPart
XL_BY_ROWS = 1             # SearchOrder: xlByRows
XL_PREVIOUS = 2            # SearchDirection: xlPrevious
ROW_FIELD = "(row)"        # write-queue entry meaning "create this item's row"


class WriterError(Exception):
    pass


class WorkbookLocked(WriterError):
    """File opened read-only / locked by another program. Retry later."""


class WorkbookBusy(WriterError):
    """Workbook open in Excel with unsaved edits. Retry later."""


class RowsMoved(WriterError):
    """Rows in the workbook moved since the last read (a row inserted / deleted / sorted in the open Excel).
    Nothing is written; the change is retried after the file is read again."""


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

class SheetLike(Protocol):
    """What the core logic needs from one sheet (ComSheet in the app, OpenpyxlSheet in tests)."""

    def header_values(self) -> list[object]: ...
    def insert_column(self, col: int) -> None: ...
    def insert_row(self, row: int, below: bool = False) -> None: ...
    def set_formula(self, row: int, col: int, formula: str) -> None: ...
    def last_row(self) -> int: ...
    def last_data_row(self, col: int) -> int: ...
    def column_values(self, col: int, first_row: int, last_row: int) -> list[object]: ...
    def row_values(self, row: int, first_col: int, last_col: int) -> list[object]: ...
    def hide_column(self, col: int) -> None: ...
    def get(self, row: int, col: int) -> object: ...
    def has_formula(self, row: int, col: int) -> bool: ...
    def set_date(self, row: int, col: int, d: dt.date) -> None: ...
    def set_number(self, row: int, col: int, n: object) -> None: ...
    def set_text(self, row: int, col: int, text: str) -> None: ...
    def clear(self, row: int, col: int) -> None: ...


class BookLike(Protocol):
    """What the core logic needs from a workbook (ComWorkbook in the app, OpenpyxlWorkbook in tests)."""

    def sheet(self, name: str) -> SheetLike: ...
    def save(self) -> None: ...
    def close(self) -> None: ...


# New-row formulas: field -> (template, columns used). Same shape as the PMO's existing rows (=E2/D2).
ROW_FORMULAS = {"pct_progress": ("={0}{r}/{1}{r}", "executed", "total_tc")}


def _table_end(labels: list[str], blank_first: bool) -> int:
    """Number of header columns in the table (stops at the first blank header)."""
    for i, lab in enumerate(labels):
        if not lab and not (i == 0 and blank_first) and i > 0:
            return i
    return len(labels)


def _id_text(ws: SheetLike, row: int, id_col: int) -> str:
    return str(ws.get(row, id_col) or "").strip()


def _holds_item(ws: SheetLike, sheet: str, row: int, cols: dict[str, int]) -> bool:
    """Whether the row is an item row by the importer's rules (excel.sheets.is_item_row, App ID never counted):
    not an empty row, a UAT Plan group header, a Test Progress Grand Total or a pre-numbered empty defect row.
    Such a row in an item's place with no App ID means the item's row was deleted (e.g. the last item)."""
    return is_item_row(sheet, {k: ws.get(row, c) for k, c in cols.items() if k != APP_ID_KEY})


def _row_holds_item(ws: SheetLike, sheet: str, row: int, cols: dict[str, int]) -> bool:
    """Same rule as _holds_item, the row's table cells read with one call (one Range through COM)."""
    data = {k: c for k, c in cols.items() if k != APP_ID_KEY}
    if not data:
        return False
    lo = min(data.values())
    values = ws.row_values(row, lo, max(data.values()))
    return is_item_row(sheet, {k: values[c - lo] for k, c in data.items()})


def _moved_rows(ws: SheetLike, sheet: str, id_col: int, expected: dict[str, int], cols: dict[str, int],
                written: frozenset[str] = frozenset()) -> list[str]:
    """Rows that no longer hold the item they held at the last read. An App ID cell may be filled only when
    it is blank, unknown, or a copy of an ID that is still in its own row (a copied row) - and only on a row
    that is still an item row; an ID that sits somewhere else, another known item's ID found in the wrong
    row, or a non-item row (empty, group header, Grand Total, pre-numbered) in an item's place means the
    rows moved. written: items that get a cell written now - their row must still be an item row even with
    the right App ID (its content may have been cleared after the read, the hidden App ID left behind)."""
    if not expected:
        return []
    last = max(ws.last_data_row(id_col), max(expected.values()))
    found: dict[str, list[int]] = {}
    for r, value in enumerate(ws.column_values(id_col, 2, last), start=2):   # one read, not one per cell
        if text := str(value or "").strip():
            found.setdefault(text, []).append(r)
    in_row = {r: i for i, rs in found.items() for r in rs}
    problems = []
    for item_id, row in sorted(expected.items(), key=lambda kv: kv[1]):
        cur = in_row.get(row, "")
        if cur == item_id:
            if item_id in written and not _row_holds_item(ws, sheet, row, cols):
                problems.append(f"row {row} no longer holds an item (only its App ID {item_id} is left)")
            continue
        if item_id in found:
            problems.append(f"{item_id} expected in row {row} but found in row {found[item_id][0]}")
        elif cur in expected and expected[cur] not in found[cur]:
            problems.append(f"row {row} holds {cur}, expected {item_id}")
        elif cur not in expected and not _holds_item(ws, sheet, row, cols):
            problems.append(f"row {row} no longer holds an item, expected {item_id}")
    return problems


def _check_rows_unmoved(book: BookLike, sheet_defs: dict[str, SheetDef], rows: dict[str, dict[str, int]],
                        writes: Sequence[CellWrite] = ()) -> None:
    """Before anything is written: every sheet that already has the App ID column must still have its items in
    the rows of the last read, and every row about to get a cell written must still be an item row. Else raise
    RowsMoved (nothing touched; the sync reads the file again and retries)."""
    problems = []
    for name, sd in sheet_defs.items():
        ws = book.sheet(name)
        labels = [header_label(v) for v in ws.header_values()]
        hmap = map_headers(labels, list(sd.fields), name, sd.blank_first_as)
        if APP_ID_KEY not in hmap.columns:
            continue
        cols = {k: c + 1 for k, c in hmap.columns.items()}
        written = frozenset(w.item_id for w in writes if w.sheet == name and w.field != ROW_FIELD)
        moved = _moved_rows(ws, name, cols[APP_ID_KEY], rows.get(name, {}), cols, written)
        if moved:
            more = f" (+{len(moved) - 3} more)" if len(moved) > 3 else ""
            problems.append(f"{name}: " + "; ".join(moved[:3]) + more)
    if problems:
        raise RowsMoved("Rows in Excel moved since the last read; nothing written, retried after reading the "
                        "file again. " + " | ".join(problems))


def apply_writes(book: BookLike, sheet_defs: dict[str, SheetDef], writes: list[CellWrite],
                 rows: dict[str, dict[str, int]]) -> WriteReport:
    """rows: sheet -> {item_id: excel_row} for all active items (from the last import).
    Raises RowsMoved (before writing anything) when the rows no longer match the last read."""
    _check_rows_unmoved(book, sheet_defs, rows, writes)
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

        # App IDs: every active item's row carries its ID (fixes blanks and copied rows). Safe here: the
        # rows were checked above, so a cell that differs is blank, unknown or a copied row's duplicate.
        for item_id, row in sorted(rows.get(name, {}).items(), key=lambda kv: kv[1]):
            if _id_text(ws, row, id_col) != item_id:
                ws.set_text(row, id_col, item_id)
                report.ids_written += 1

        cell_writes = [w for w in sheet_writes if w.field != ROW_FIELD]
        row_writes = sorted((w for w in sheet_writes if w.field == ROW_FIELD), key=lambda w: -w.excel_row)
        for w in cell_writes:
            key3 = (w.sheet, w.item_id, w.field)
            if _id_text(ws, w.excel_row, id_col) != w.item_id:
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
    def last_data_row(self, col: int) -> int:
        """Last row with something in this column, hidden rows included (UsedRange can reach row 1,048,576
        after formatting). End(xlUp) skips every hidden row; Find with LookIn=xlFormulas sees rows hidden by
        hand, collapsed groups and a hidden column (checked in real Excel). Searching backwards after the first
        cell wraps to the bottom of the sheet first, so row 1,048,576 is included.
        Find is NOT reliable (checked in real Excel) under a filter - a sheet AutoFilter or a filtered Table,
        both set ws.FilterMode - where it can return e.g. row 1, nor on a protected sheet whose cells are
        FormulaHidden, where it returns None. Then UsedRange is used: it always reaches the last filled row.
        Do not remove this fallback."""
        if self.ws.FilterMode or self.ws.ProtectContents:
            ur = self.ws.UsedRange
            return int(ur.Row) + int(ur.Rows.Count) - 1
        # positional: What, After, LookIn, LookAt, SearchOrder, SearchDirection
        hit = self.ws.Columns(col).Find("*", self.ws.Cells(1, col), XL_FORMULAS, XL_PART, XL_BY_ROWS,
                                        XL_PREVIOUS)
        return 1 if hit is None else int(hit.Row)

    @_busy_retry
    def column_values(self, col: int, first_row: int, last_row: int) -> list[object]:
        """Values of rows first_row..last_row of one column, read with a single COM call."""
        if last_row < first_row:
            return []
        vals = self.ws.Range(self.ws.Cells(first_row, col), self.ws.Cells(last_row, col)).Value
        if first_row == last_row:                         # one cell: COM returns the value, not a tuple
            return [vals]
        return [row[0] for row in vals]

    @_busy_retry
    def row_values(self, row: int, first_col: int, last_col: int) -> list[object]:
        """Values of columns first_col..last_col of one row, read with a single COM call."""
        if last_col < first_col:
            return []
        vals = self.ws.Range(self.ws.Cells(row, first_col), self.ws.Cells(row, last_col)).Value
        if first_col == last_col:                         # one cell: COM returns the value, not a tuple
            return [vals]
        return list(vals[0])

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

    def last_data_row(self, col: int) -> int:
        return self.ws.max_row

    def column_values(self, col: int, first_row: int, last_row: int) -> list[object]:
        return [self.ws.cell(r, col).value for r in range(first_row, last_row + 1)]

    def row_values(self, row: int, first_col: int, last_col: int) -> list[object]:
        return [self.ws.cell(row, c).value for c in range(first_col, last_col + 1)]

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
