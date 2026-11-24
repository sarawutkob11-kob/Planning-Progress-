"""Bug fixes 12-Nov-2026 (docs/tasks.md T1-T6).

T1  writer: rows moved in the open Excel after the last read -> RowsMoved, nothing written, sync retries
T2  sync: interrupted 'writing' entries with attempts > 0 are retried; failed new row fails its field edits
T3  app: a project that cannot be opened is not tried again (no endless recursion)
T4  restore_backup: zip entries outside the attachment folder are refused
T5  sync alerts use the PMO email changed in Settings after the SyncService was created
T6  unused imports removed
"""

from __future__ import annotations

import ast
import dataclasses
import datetime as dt
import hashlib
import itertools
import json
import re
import sqlite3
import zipfile
from pathlib import Path

import openpyxl
import pytest

from uat_tracker.config import AlertsConfig
from uat_tracker.excel.reader import RawRow
from uat_tracker.excel.sheets import (APP_ID_KEY, SHEETS, is_defect_item, is_group_header, is_item_row,
                                      is_test_progress_item, with_overrides)
from uat_tracker.excel.writer import (CellWrite, OpenpyxlWorkbook, RowsMoved, WorkbookLocked, WriterError,
                                      apply_writes)
from uat_tracker.mail.sender import FakeSender
from uat_tracker.parsing.text import cell_text, collapse, task_no_text
from uat_tracker.services.clear_project import ClearError, restore_backup
from uat_tracker.services.edits import edit_item
from uat_tracker.services.importer import run_import
from uat_tracker.services.items import add_application, add_task
from uat_tracker.services.normalize import normalize_defects, normalize_test_progress, normalize_uat_plan
from uat_tracker.services.settings import save_settings
from uat_tracker.services.sync import SyncService

from conftest import PMO_ACTOR, WINDOW
# fixtures shared with the multi-project tests (T3)
from test_projects import PERSON, appdb, base, qapp, two  # noqa: F401

SRC = Path(__file__).resolve().parents[1] / "src" / "uat_tracker"
TP, UAT, DEF = "Test Progress", "UAT Plan", "Defect List"
PMO = "pmo@example.test"
D = dt.datetime


# =========================================================================== helpers

@pytest.fixture
def scfg(cfg):
    return dataclasses.replace(cfg, alerts=AlertsConfig(pmo_email=PMO))


@pytest.fixture
def svc(scfg, repo):
    run_import(scfg, repo, editor="t")
    s = SyncService(scfg, repo, opener=OpenpyxlWorkbook, sender=FakeSender())
    assert s.init_ids().result == "written"                 # App ID column present (as in production)
    return s


def sha(path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def cols_of(ws) -> dict[str, int]:
    """Header name (first line) -> column. Careful: a name can be used by more than one column (UAT Plan:
    'Status' and 'Status (Optional)'; Test Progress: 'Status' of the table and of the legend on the right).
    The FIRST column with the name is kept (the table's own column)."""
    cols: dict[str, int] = {}
    for c in ws[1]:
        if c.value:
            cols.setdefault(str(c.value).split("\n")[0].strip(), c.column)
    return cols


def snapshot(wb) -> dict:
    """Every cell value of every sheet (the in-memory workbook)."""
    return {(ws.title, c.row, c.column): c.value for ws in wb.worksheets for row in ws.iter_rows() for c in row
            if c.value is not None}


def defs():
    return {name: with_overrides(sd, {}) for name, sd in SHEETS.items()}


def rows_of(repo) -> dict[str, dict[str, int]]:
    return {name: {i: r["excel_row"] for i, r in repo.load_items(name).items()
                   if r.get("active") and r.get("excel_row")}
            for name in SHEETS}


def text(v) -> str:
    return json.dumps({"kind": "text", "value": v})


def row_by_id(ws, cols, item_id) -> list[int]:
    return [r for r in range(2, ws.max_row + 1) if ws.cell(r, cols["App ID"]).value == item_id]


def assert_ids_consistent(path, repo):
    """Each active item's DB row holds its own App ID in Excel and no App ID appears twice."""
    wb = openpyxl.load_workbook(path)
    for name in SHEETS:
        ws = wb[name]
        c = cols_of(ws)
        ids = [ws.cell(r, c["App ID"]).value for r in range(2, ws.max_row + 1) if ws.cell(r, c["App ID"]).value]
        assert len(ids) == len(set(ids)), f"{name}: duplicate App IDs {ids}"
        for item_id, row in rows_of(repo)[name].items():
            assert ws.cell(row, c["App ID"]).value == item_id, f"{name} row {row}: expected {item_id}"


def insert_row(ws, at: int, values: list | None = None):
    ws.insert_rows(at)
    for i, v in enumerate(values or [], start=1):
        ws.cell(at, i).value = v


def row_values(ws, r) -> list:
    return [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]


def reverse_rows(ws, first: int, last: int):
    """Sort descending in Excel: whole rows (incl. the hidden App ID) change places."""
    vals = [row_values(ws, r) for r in range(first, last + 1)]
    for r, v in zip(range(first, last + 1), reversed(vals)):
        for c, x in enumerate(v, start=1):
            ws.cell(r, c).value = x


def tp_row_of(ws, app_name) -> int:
    c = cols_of(ws)
    return next(r for r in range(2, ws.max_row + 1) if ws.cell(r, c["Application"]).value == app_name)


# mutations made in the OPEN Excel after the app's last read (the disk file is not changed yet)
MUTATIONS = {
    "tp_insert_top": lambda wb: insert_row(wb[TP], 2),
    "tp_insert_middle": lambda wb: insert_row(wb[TP], 4, [99, "NEW APP"]),
    "tp_delete_first": lambda wb: wb[TP].delete_rows(2),
    "tp_delete_middle": lambda wb: wb[TP].delete_rows(4),
    "tp_sort_reverse": lambda wb: reverse_rows(wb[TP], 2, 6),
    "tp_swap_two": lambda wb: reverse_rows(wb[TP], 4, 5),
    "tp_copy_above_original": lambda wb: insert_row(wb[TP], 5, row_values(wb[TP], 5)),
    "uat_insert_top": lambda wb: insert_row(wb[UAT], 2),
    "uat_delete_middle": lambda wb: wb[UAT].delete_rows(7),
    "def_insert_top": lambda wb: insert_row(wb[DEF], 2),
}


def mutating_opener(mutate, calls: dict | None = None):
    """Opener whose workbook is the open Excel copy: `mutate` applied in memory, after the sync read the file."""
    calls = calls if calls is not None else {}

    class Book(OpenpyxlWorkbook):
        def __init__(self, path):
            super().__init__(path)
            mutate(self.wb)
            calls["book"] = self
            calls["saved"] = False

        def save(self):
            calls["saved"] = True
            super().save()
    return Book


# =========================================================================== T1: writer, direct

@pytest.mark.parametrize("name", sorted(MUTATIONS))
def test_t1_moved_rows_raise_and_nothing_changes(scfg, repo, svc, name):
    book = OpenpyxlWorkbook(scfg.excel_path)
    MUTATIONS[name](book.wb)
    before = snapshot(book.wb)
    writes = [CellWrite(TP, "TP-38", "remark", 5, text("tp38 note")),
              CellWrite(UAT, "B-6", "remark", 7, text("b6 note")),
              CellWrite(DEF, "DEF-5", "remark", 3, text("def5 note"))]
    with pytest.raises(RowsMoved) as ei:
        apply_writes(book, defs(), writes, rows_of(repo))
    assert isinstance(ei.value, WriterError)
    assert snapshot(book.wb) == before                        # no App ID, no value, on any sheet
    book.close()


def test_t1_insert_top_reported_case_no_wrong_row_no_id_swap(scfg, repo, svc):
    """The reported case: row inserted on top of Test Progress, then TP-38's remark is written."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    insert_row(book.wb[TP], 2)
    with pytest.raises(RowsMoved, match="moved"):
        apply_writes(book, defs(), [CellWrite(TP, "TP-38", "remark", 5, text("x"))], rows_of(repo))
    ws = book.wb[TP]
    c = cols_of(ws)
    collins = tp_row_of(ws, "COLLINS")
    assert ws.cell(collins, c["Remark"]).value is None
    assert ws.cell(collins, c["App ID"]).value == "TP-12"
    for app, item in (("Nleads", "TP-1"), ("ATM", "TP-6"), ("DMS-RAMC", "TP-38"), ("SCBPDM", "TP-35")):
        assert ws.cell(tp_row_of(ws, app), c["App ID"]).value == item
    book.close()


def test_t1_unmoved_rows_still_written(scfg, repo, svc):
    book = OpenpyxlWorkbook(scfg.excel_path)
    rep = apply_writes(book, defs(), [CellWrite(TP, "TP-38", "remark", 5, text("ok"))], rows_of(repo))
    assert rep.written == [(TP, "TP-38", "remark")] and not rep.failed and rep.ids_written == 0
    assert book.wb[TP].cell(5, cols_of(book.wb[TP])["Remark"]).value == "ok"
    book.close()


def test_t1_rows_added_below_last_item_are_not_a_move(scfg, repo, svc):
    """Adding rows under the table (or a copy of the last row under it) shifts no known row."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    ws = book.wb[TP]
    insert_row(ws, 7, row_values(ws, 6))                      # copy of TP-35 under it, above Grand Total
    rep = apply_writes(book, defs(), [CellWrite(TP, "TP-35", "remark", 6, text("ok"))], rows_of(repo))
    assert rep.written == [(TP, "TP-35", "remark")] and not rep.failed
    assert ws.cell(6, cols_of(ws)["Remark"]).value == "ok" and ws.cell(7, cols_of(ws)["Remark"]).value is None
    book.close()


def test_t1_blank_and_unknown_id_cells_are_filled(scfg, repo, svc):
    book = OpenpyxlWorkbook(scfg.excel_path)
    ws = book.wb[TP]
    c = cols_of(ws)
    ws.cell(4, c["App ID"]).value = None                      # TP-12 blank
    ws.cell(3, c["App ID"]).value = "ZZ-999"                  # TP-6 unknown
    rep = apply_writes(book, defs(), [CellWrite(TP, "TP-12", "remark", 4, text("collins"))], rows_of(repo))
    assert rep.written == [(TP, "TP-12", "remark")] and rep.ids_written == 2
    assert ws.cell(4, c["App ID"]).value == "TP-12" and ws.cell(3, c["App ID"]).value == "TP-6"
    assert ws.cell(4, c["Remark"]).value == "collins"
    book.close()


def test_t1_copied_row_duplicate_id_is_replaced(scfg, repo, svc):
    """A copied row whose App ID duplicates an item still in its own row (as bound after the next read):
    the copy's row gets its own ID, nothing is reported as moved."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    rows = rows_of(repo)
    ws = book.wb[TP]
    c = cols_of(ws)
    insert_row(ws, 6, row_values(ws, 5))                      # copy of TP-38 under it (rows 5 and 6 = TP-38)
    rows[TP] = {"TP-1": 2, "TP-6": 3, "TP-12": 4, "TP-38": 5, "TP-38b": 6, "TP-35": 7}   # what the next read binds
    rep = apply_writes(book, defs(), [CellWrite(TP, "TP-38b", "remark", 6, text("copy"))], rows)
    assert rep.written == [(TP, "TP-38b", "remark")] and rep.ids_written == 1
    assert [ws.cell(r, c["App ID"]).value for r in range(2, 8)] == ["TP-1", "TP-6", "TP-12", "TP-38", "TP-38b",
                                                                    "TP-35"]
    book.close()


def test_t1_first_write_adds_id_column_without_check(scfg, repo):
    run_import(scfg, repo, editor="t")
    book = OpenpyxlWorkbook(scfg.excel_path)
    rep = apply_writes(book, defs(), [CellWrite(TP, "TP-38", "remark", 5, text("first"))], rows_of(repo))
    assert sorted(rep.id_column_added) == sorted(SHEETS) and rep.written == [(TP, "TP-38", "remark")]
    ws = book.wb[TP]
    assert ws.cell(tp_row_of(ws, "DMS-RAMC"), cols_of(ws)["Remark"]).value == "first"
    book.close()


@pytest.mark.parametrize("sheet,item,row", [(TP, "TP-35", 6), (UAT, "E-x1", 13)])
def test_t1_deleted_last_item_row_is_not_written_elsewhere(scfg, repo, svc, sheet, item, row):
    """Deleting the LAST item row of a sheet in the open Excel is also a row change since the read: the
    item's App ID / value must not land in whatever row now has that number (Grand Total / an empty row)."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    ws = book.wb[sheet]
    ws.delete_rows(row)
    before = snapshot(book.wb)
    with pytest.raises(RowsMoved):
        apply_writes(book, defs(), [CellWrite(sheet, item, "remark", row, text("lost"))], rows_of(repo))
    c = cols_of(ws)
    assert ws.cell(row, c["App ID"]).value != item, f"{item}'s App ID written into row {row} after it was deleted"
    assert ws.cell(row, c["Remark"]).value != "lost"
    assert snapshot(book.wb) == before
    book.close()


def test_t1_deleted_last_defect_above_prenumbered_rows_raises(scfg, repo, svc):
    """Defect List: rows 5/6 carry only a pre-filled 'No.' (23, 24), no data. Deleting DEF-21 (row 4, the last
    item) moves the pre-numbered row 23 into row 4: that is a row change, not a row to put DEF-21's ID on."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    ws = book.wb[DEF]
    c = cols_of(ws)
    assert ws.cell(5, 1).value == 23 and ws.cell(4, c["App ID"]).value == "DEF-21"
    ws.delete_rows(4)
    assert ws.cell(4, 1).value == 23                          # the pre-numbered empty row is now row 4
    before = snapshot(book.wb)
    with pytest.raises(RowsMoved):
        apply_writes(book, defs(), [CellWrite(DEF, "DEF-21", "remark", 4, text("lost"))], rows_of(repo))
    assert snapshot(book.wb) == before
    assert ws.cell(4, c["App ID"]).value is None and ws.cell(4, c["Remark"]).value is None
    book.close()


def test_t1_deleted_last_defect_other_item_write_raises(scfg, repo, svc):
    """Same deletion, but the write is for another (unmoved) defect: still RowsMoved for the sheet, because
    the App ID pass would otherwise stamp DEF-21 onto the pre-numbered row."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    book.wb[DEF].delete_rows(4)
    before = snapshot(book.wb)
    with pytest.raises(RowsMoved):
        apply_writes(book, defs(), [CellWrite(DEF, "DEF-5", "remark", 3, text("x"))], rows_of(repo))
    assert snapshot(book.wb) == before
    book.close()


def test_t1_deleted_last_task_of_group_followed_by_group_header(scfg, repo, svc):
    """UAT Plan: delete A-2 (row 5, last task of group A) -> the group B header moves into row 5."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    ws = book.wb[UAT]
    c = cols_of(ws)
    ws.delete_rows(5)
    assert ws.cell(5, c["Group"]).value == "B" and ws.cell(5, c["Task No"]).value is None
    before = snapshot(book.wb)
    with pytest.raises(RowsMoved):
        apply_writes(book, defs(), [CellWrite(UAT, "A-2", "remark", 5, text("lost"))], rows_of(repo))
    assert ws.cell(5, c["App ID"]).value is None              # no App ID on the group header row
    assert ws.cell(5, c["Remark"]).value is None
    assert snapshot(book.wb) == before
    book.close()


@pytest.fixture
def trailing_group(tmp_path, scfg, repo):
    """Workbook (as read) whose LAST group header 'F' has no task yet and sits right under E-x1 (row 14).
    Deleting E-x1 then shifts no other known item: only the header lands in E-x1's row."""
    from conftest import UAT_ROWS, build_workbook, uat_row
    build_workbook(scfg.excel_path, uat=UAT_ROWS + [uat_row("F", None, "Go-Live Phase")])
    run_import(scfg, repo, editor="t")
    s = SyncService(scfg, repo, opener=OpenpyxlWorkbook, sender=FakeSender())
    assert s.init_ids().result == "written"
    rows = rows_of(repo)
    assert rows[UAT]["E-x1"] == 13 and max(rows[UAT].values()) == 13
    ws = openpyxl.load_workbook(scfg.excel_path)[UAT]
    assert ws.cell(14, cols_of(ws)["Group"]).value == "F"
    return s


def test_t1_deleted_last_task_last_group_header_below_not_stamped(scfg, repo, trailing_group):
    """The case that only a header (no other item) moves into the deleted item's row: the header has a Group
    letter + title but no Task No. E-x1's App ID / value must not be written onto it."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    ws = book.wb[UAT]
    c = cols_of(ws)
    ws.delete_rows(13)
    assert ws.cell(13, c["Group"]).value == "F" and ws.cell(13, c["Task No"]).value is None
    before = snapshot(book.wb)
    with pytest.raises(RowsMoved):
        apply_writes(book, defs(), [CellWrite(UAT, "E-x1", "remark", 13, text("lost"))], rows_of(repo))
    assert ws.cell(13, c["App ID"]).value != "E-x1", "E-x1's App ID written onto the group F header row"
    assert ws.cell(13, c["Remark"]).value != "lost"
    assert snapshot(book.wb) == before
    book.close()


def test_t1_deleted_last_task_last_group_header_below_raises(scfg, repo, trailing_group):
    """Stricter: the case must be reported as RowsMoved (so the sync re-reads), whichever item is written."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    book.wb[UAT].delete_rows(13)
    with pytest.raises(RowsMoved):
        apply_writes(book, defs(), [CellWrite(UAT, "E-11", "remark", 12, text("x"))], rows_of(repo))
    book.close()


# =========================================================================== T1: through the sync


@pytest.mark.parametrize("sheet,row,other,other_row", [
    (DEF, 4, "DEF-5", 3),                # last defect above the pre-numbered rows 23/24
    (TP, 6, "TP-38", 5),                 # last application above Grand Total
])
def test_t1_sync_deleted_last_row_is_retry_then_written_right(scfg, repo, svc, sheet, row, other, other_row):
    edit_item(scfg, repo, sheet, other, {"remark": "kept"}, actor=PMO_ACTOR)
    disk = sha(scfg.excel_path)
    calls: dict = {}
    moving = SyncService(scfg, repo, opener=mutating_opener(lambda wb: wb[sheet].delete_rows(row), calls),
                         sender=FakeSender())
    fr = moving.flush(force=True)
    assert fr.result == "retry", fr
    assert calls["saved"] is False and sha(scfg.excel_path) == disk
    pend = repo.all_pending()
    assert len(pend) == 1 and pend[0]["status"] == "pending" and pend[0]["next_attempt_at"]
    assert moving.status()["failed"] == 0
    assert not moving.sender.sent
    # the file catches up: the deleted row is saved, the next read drops the item and binds the rest
    wb = openpyxl.load_workbook(scfg.excel_path)
    wb[sheet].delete_rows(row)
    wb.save(scfg.excel_path)
    svc.retry_now()
    fr2 = svc.flush(force=True)
    assert fr2.result == "written", fr2
    assert repo.all_pending() == []
    ws = openpyxl.load_workbook(scfg.excel_path)[sheet]
    c = cols_of(ws)
    assert ws.cell(other_row, c["App ID"]).value == other and ws.cell(other_row, c["Remark"]).value == "kept"
    assert ws.cell(row, c["App ID"]).value is None             # the row now in the deleted item's place
    assert ws.cell(row, c["Remark"]).value is None
    assert_ids_consistent(scfg.excel_path, repo)


def _edit_three_sheets(scfg, repo):
    edit_item(scfg, repo, TP, "TP-38", {"remark": "tp38 note"}, actor=PMO_ACTOR)
    edit_item(scfg, repo, UAT, "B-6", {"remark": "b6 note"}, actor=PMO_ACTOR)
    edit_item(scfg, repo, DEF, "DEF-5", {"remark": "def5 note"}, actor=PMO_ACTOR)


def test_t1_sync_rows_moved_is_retry_then_written_to_right_row(scfg, repo, svc):
    _edit_three_sheets(scfg, repo)
    disk = sha(scfg.excel_path)
    calls: dict = {}
    moving = SyncService(scfg, repo, opener=mutating_opener(MUTATIONS["tp_insert_top"], calls), sender=FakeSender())
    fr = moving.flush(force=True)
    assert fr.result == "retry", fr
    assert "moved" in fr.message.lower()
    assert calls["saved"] is False and sha(scfg.excel_path) == disk          # workbook not changed (no sheet)
    pend = repo.all_pending()
    assert len(pend) == 3
    assert all(p["status"] == "pending" and p["attempts"] == 1 and p["next_attempt_at"] for p in pend)
    assert not [p for p in pend if p["status"] == "failed"]
    assert moving.status()["failed"] == 0
    # the file catches up (Excel saves the inserted row), the watcher sees the change -> tried again
    wb = openpyxl.load_workbook(scfg.excel_path)
    insert_row(wb[TP], 2)
    wb.save(scfg.excel_path)
    svc.retry_now()
    fr2 = svc.flush(force=True)
    assert fr2.result == "written", fr2
    assert repo.all_pending() == []
    wb = openpyxl.load_workbook(scfg.excel_path)
    ws = wb[TP]
    c = cols_of(ws)
    assert ws.cell(tp_row_of(ws, "DMS-RAMC"), c["Remark"]).value == "tp38 note"
    assert ws.cell(tp_row_of(ws, "COLLINS"), c["Remark"]).value is None
    for app, item in (("Nleads", "TP-1"), ("ATM", "TP-6"), ("COLLINS", "TP-12"), ("DMS-RAMC", "TP-38"),
                      ("SCBPDM", "TP-35")):
        assert ws.cell(tp_row_of(ws, app), c["App ID"]).value == item
    ws = wb[UAT]
    c = cols_of(ws)
    assert [ws.cell(r, c["Remark"]).value for r in row_by_id(ws, c, "B-6")] == ["b6 note"]
    ws = wb[DEF]
    c = cols_of(ws)
    assert [ws.cell(r, c["Remark"]).value for r in row_by_id(ws, c, "DEF-5")] == ["def5 note"]
    assert_ids_consistent(scfg.excel_path, repo)


@pytest.mark.parametrize("name", ["tp_delete_middle", "tp_sort_reverse", "uat_insert_top", "def_insert_top"])
def test_t1_sync_other_moves_are_retry_and_nothing_saved(scfg, repo, svc, name):
    _edit_three_sheets(scfg, repo)
    disk = sha(scfg.excel_path)
    calls: dict = {}
    moving = SyncService(scfg, repo, opener=mutating_opener(MUTATIONS[name], calls), sender=FakeSender())
    fr = moving.flush(force=True)
    assert fr.result == "retry", fr
    assert calls["saved"] is False and sha(scfg.excel_path) == disk
    assert all(p["status"] == "pending" for p in repo.all_pending())
    assert not moving.sender.sent                                             # a retry, not an error email


def test_t1_sync_one_sheet_moved_other_sheets_not_written(scfg, repo, svc):
    """Only Test Progress moved: the UAT Plan / Defect List edits are not written either (in memory too)."""
    _edit_three_sheets(scfg, repo)
    calls: dict = {}
    moving = SyncService(scfg, repo, opener=mutating_opener(MUTATIONS["tp_insert_top"], calls), sender=FakeSender())
    assert moving.flush(force=True).result == "retry"
    wb = calls["book"].wb                                                      # the open Excel copy
    for sheet, item in ((UAT, "B-6"), (DEF, "DEF-5")):
        ws = wb[sheet]
        c = cols_of(ws)
        assert [ws.cell(r, c["Remark"]).value for r in row_by_id(ws, c, item)] == [None]


def test_t1_sync_copy_row_above_and_below_saved(scfg, repo, svc):
    """Copied rows (saved in Excel, App ID copied too) still work: the copy gets its own ID, edits land right."""
    wb = openpyxl.load_workbook(scfg.excel_path)
    ws = wb[TP]
    insert_row(ws, 5, row_values(ws, 5))                      # copy of TP-38 above the original
    insert_row(ws, 4, row_values(ws, 3))                      # copy of TP-6 below the original (rows 3, 4)
    wb.save(scfg.excel_path)
    edit_item(scfg, repo, TP, "TP-12", {"remark": "collins"}, actor=PMO_ACTOR)
    fr = svc.flush(force=True)
    assert fr.result == "written", fr
    assert repo.all_pending() == []
    assert_ids_consistent(scfg.excel_path, repo)
    wb = openpyxl.load_workbook(scfg.excel_path)
    ws = wb[TP]
    c = cols_of(ws)
    assert ws.cell(tp_row_of(ws, "COLLINS"), c["Remark"]).value == "collins"
    assert ws.cell(tp_row_of(ws, "COLLINS"), c["App ID"]).value == "TP-12"
    ids = [ws.cell(r, c["App ID"]).value for r in range(2, ws.max_row + 1) if ws.cell(r, c["App ID"]).value]
    assert len(ids) == 7 and ids.count("TP-38") == 1 and ids.count("TP-6") == 1


def test_t1_sync_blank_and_unknown_app_id_saved(scfg, repo, svc):
    wb = openpyxl.load_workbook(scfg.excel_path)
    ws = wb[TP]
    c = cols_of(ws)
    ws.cell(tp_row_of(ws, "COLLINS"), c["App ID"]).value = None
    ws.cell(tp_row_of(ws, "ATM"), c["App ID"]).value = "ZZ-999"
    wb.save(scfg.excel_path)
    edit_item(scfg, repo, TP, "TP-12", {"remark": "collins"}, actor=PMO_ACTOR)
    edit_item(scfg, repo, TP, "TP-6", {"remark": "atm"}, actor=PMO_ACTOR)
    fr = svc.flush(force=True)
    assert fr.result == "written", fr
    assert repo.all_pending() == []
    ws = openpyxl.load_workbook(scfg.excel_path)[TP]
    c = cols_of(ws)
    assert ws.cell(tp_row_of(ws, "COLLINS"), c["App ID"]).value == "TP-12"
    assert ws.cell(tp_row_of(ws, "ATM"), c["App ID"]).value == "TP-6"
    assert ws.cell(tp_row_of(ws, "COLLINS"), c["Remark"]).value == "collins"
    assert ws.cell(tp_row_of(ws, "ATM"), c["Remark"]).value == "atm"
    assert_ids_consistent(scfg.excel_path, repo)


def test_t1_sync_first_write_adds_app_id_column(scfg, repo):
    run_import(scfg, repo, editor="t")
    s = SyncService(scfg, repo, opener=OpenpyxlWorkbook, sender=FakeSender())
    edit_item(scfg, repo, TP, "TP-38", {"remark": "first"}, actor=PMO_ACTOR)
    fr = s.flush(force=True)
    assert fr.result == "written" and "column added" in fr.message
    ws = openpyxl.load_workbook(scfg.excel_path)[TP]
    assert ws.cell(tp_row_of(ws, "DMS-RAMC"), cols_of(ws)["Remark"]).value == "first"
    assert_ids_consistent(scfg.excel_path, repo)


def test_t1_sync_new_rows_and_new_group_with_edits(scfg, repo, svc):
    tp_id = add_application(scfg, repo, application="NewApp", report_by="QA X", planned_start="2026-11-10",
                            planned_finish="2026-11-20", total_tc=10, actor=PMO_ACTOR)
    t1 = add_task(scfg, repo, group=None, new_group_title="New Phase", task_name="First new", responsible="QA X",
                  actor=PMO_ACTOR)
    letter = repo.get_item(UAT, t1)["group_code"]
    t2 = add_task(scfg, repo, group=letter, task_name="Second new", responsible="QA X", actor=PMO_ACTOR)
    t3 = add_task(scfg, repo, group="B", task_name="In group B", responsible="QA Y", actor=PMO_ACTOR)
    edit_item(scfg, repo, TP, "TP-38", {"remark": "tp38"}, actor=PMO_ACTOR)
    edit_item(scfg, repo, UAT, "E-11", {"remark": "e11"}, actor=PMO_ACTOR)
    fr = svc.flush(force=True)
    assert fr.result == "written", fr
    fr2 = svc.flush(force=True)                                # 2nd task of the new group: next round
    assert fr2.result == "written", fr2
    assert repo.all_pending() == []
    assert_ids_consistent(scfg.excel_path, repo)
    wb = openpyxl.load_workbook(scfg.excel_path)
    ws = wb[TP]
    c = cols_of(ws)
    assert ws.cell(row_by_id(ws, c, tp_id)[0], c["Application"]).value == "NewApp"
    assert ws.cell(tp_row_of(ws, "DMS-RAMC"), c["Remark"]).value == "tp38"
    ws = wb[UAT]
    c = cols_of(ws)
    r1, r2, r3 = (row_by_id(ws, c, t)[0] for t in (t1, t2, t3))
    assert ws.cell(r1 - 1, c["Group"]).value == letter and ws.cell(r1 - 1, c["Task name"]).value == "New Phase"
    assert r2 == r1 + 1
    assert ws.cell(r3, c["Task name"]).value == "In group B" and r3 == row_by_id(ws, c, "B-7")[0] + 1
    assert ws.cell(row_by_id(ws, c, "E-11")[0], c["Remark"]).value == "e11"


# =========================================================================== T1: shared "is this an item row" rules

# The importer's conditions before the rules moved to excel.sheets (HEAD normalize.py, copied verbatim).
_OLD_GROUP_LETTER = re.compile(r"[A-Za-z]")


def _old_is_group_header(v: dict) -> bool:
    return bool(_OLD_GROUP_LETTER.fullmatch(cell_text(v.get("group")))) and not task_no_text(v.get("task_no"))


def _old_tp_skipped(v: dict) -> bool:
    app = cell_text(v.get("application"))
    no = task_no_text(v.get("no"))
    return collapse(app) == "grand total" or (not app and not no)


def _old_defect_skipped(v: dict) -> bool:
    return all(cell_text(val) == "" for k, val in v.items() if k != "no")


_CELLS = [None, "", " ", "\xa0", "\n", "a", "A", "z", "B ", " c", "AB", "A1", "1", 1, 0, 2.5, 14.1, "x",
          "Grand Total", " grand  TOTAL ", "Grand\nTotal", "GRAND TOTAL", "Grand Totals", True, False,
          dt.datetime(2026, 9, 1), "ก"]


def test_t1_rules_group_header_same_as_old_importer():
    for g, t, name in itertools.product(_CELLS, _CELLS, [None, "SETUP", ""]):
        v = {"group": g, "task_no": t, "task_name": name}
        assert is_group_header(v) == _old_is_group_header(v), v


def test_t1_rules_test_progress_same_as_old_importer():
    for a, no, rb in itertools.product(_CELLS, _CELLS, [None, "QA"]):
        v = {"no": no, "application": a, "report_by": rb}
        assert is_test_progress_item(v) == (not _old_tp_skipped(v)), v


def test_t1_rules_defect_same_as_old_importer():
    for no, d, rem in itertools.product(_CELLS, _CELLS, [None, "", "r", 0]):
        v = {"no": no, "detail": d, "remark": rem}
        assert is_defect_item(v) == (not _old_defect_skipped(v)), v
    assert is_defect_item({}) is False and is_defect_item({"no": 23}) is False


def test_t1_rules_is_item_row_per_sheet():
    assert not is_item_row(UAT, {}) and not is_item_row(TP, {}) and not is_item_row(DEF, {})
    assert not is_item_row(UAT, {"group": None, "task_no": " ", "task_name": "\xa0"})        # all blank
    assert not is_item_row(UAT, {"group": "f", "task_no": None, "task_name": "Go-Live", "note": "n"})
    assert is_item_row(UAT, {"group": "F", "task_no": 3, "task_name": "x"})
    assert is_item_row(UAT, {"group": None, "task_no": None, "task_name": "no number"})
    assert not is_item_row(TP, {"no": None, "application": "Grand Total", "total_tc": 238})
    assert is_item_row(TP, {"no": 7, "application": None})
    assert not is_item_row(TP, {"no": None, "application": None, "report_by": "QA"})
    assert not is_item_row(DEF, {"no": 23})
    assert is_item_row(DEF, {"no": None, "detail": "d"})
    assert is_item_row("Other sheet", {"x": 1})


def _old_filter_rows(sheet: str, rows: list[RawRow]) -> list[int]:
    """Excel rows the old importer turned into items."""
    if sheet == UAT:
        kept, group = [], None
        for r in rows:
            if _old_is_group_header(r.values):
                group = "seen"
                continue
            if group is not None:
                kept.append(r.excel_row)
        return kept
    skip = _old_tp_skipped if sheet == TP else _old_defect_skipped
    return [r.excel_row for r in rows if not skip(r.values)]


def test_t1_normalize_keeps_same_rows_as_old_rules():
    cells = [None, "", "a", "A", "AB", 1, "Grand Total", " grand total "]
    uat_rows, tp_rows, def_rows = [], [], []
    for i, (g, t, n) in enumerate(itertools.product(cells, cells, [None, "name"]), start=2):
        uat_rows.append(RawRow(i, {"group": g, "task_no": t, "task_name": n}))
    for i, (a, no) in enumerate(itertools.product(cells, cells), start=2):
        tp_rows.append(RawRow(i, {"no": no, "application": a}))
        def_rows.append(RawRow(i, {"no": no, "detail": a}))
    assert [it.excel_row for it in normalize_uat_plan(uat_rows, WINDOW)] == _old_filter_rows(UAT, uat_rows)
    assert [it.excel_row for it in normalize_test_progress(tp_rows, WINDOW)] == _old_filter_rows(TP, tp_rows)
    assert [it.excel_row for it in normalize_defects(def_rows, WINDOW)] == _old_filter_rows(DEF, def_rows)


@pytest.fixture
def edge_import(scfg, repo):
    """Workbook with the edge rows of all 3 sheets, imported."""
    from conftest import UAT_ROWS, build_workbook, uat_row
    uat = [
        UAT_ROWS[0],                                                                    # 2 above group A
        uat_row("a", None, "Lower-case group", owner="QA X", status="Done", note="hdr note",
                remark="hdr remark", start=D(2026, 9, 1)),                              # 3 header 'a'
        uat_row(None, 1, "Task a1", "QA A"),                                            # 4 A-1
        uat_row("b", None, "Group b", note="n"),                                        # 5 header 'b'
        uat_row(None, 2, "Task b2", "QA B"),                                            # 6 B-2
        uat_row("C", 3, "Letter AND task no"),                                          # 7 item of group B
        uat_row("AB", None, "Two letters"),                                             # 8 item of group B
        uat_row(" d ", None, "Spaced group"),                                           # 9 header 'D'
        uat_row(None, 4, "Task d4", "QA D"),                                            # 10 D-4
    ]
    tp = [
        [1, "Nleads", "QA K", 17, 17, 1, 17, 0],                                        # 2 item
        [7, None, "QA N", 4],                                                           # 3 No. only -> item
        [None, None, "QA only", 3, 1],                                                  # 4 neither -> not
        [None, "Grand Total", None, 238, 178, 0.75, 174, 2],                            # 5 not
    ]
    defects = [
        [1, D(2026, 9, 17), "Salesforce", "Branch", "Close", 3, "W", None, None, "C", None],   # 2 item
        [2],                                                                            # 3 pre-numbered
        [None, None, None, "Only detail"],                                              # 4 item
        [3] + [None] * 9 + ["only remark"],                                             # 5 item
        [4],                                                                            # 6 pre-numbered
    ]
    build_workbook(scfg.excel_path, uat=uat, tp=tp, defects=defects)
    run_import(scfg, repo, editor="t")
    return {name: {i: r for i, r in repo.load_items(name).items() if r.get("active")} for name in SHEETS}


def test_t1_import_uat_plan_group_headers(edge_import):
    items = edge_import[UAT]
    assert sorted(r["excel_row"] for r in items.values()) == [4, 6, 7, 8, 10]
    by_row = {r["excel_row"]: r for r in items.values()}
    assert by_row[4]["group_code"] == "A" and by_row[4]["group_title"] == "Lower-case group"
    assert by_row[4]["task_name"] == "Task a1"
    assert by_row[6]["group_code"] == "B" and by_row[7]["group_code"] == "B" and by_row[8]["group_code"] == "B"
    assert by_row[7]["task_name"] == "Letter AND task no" and by_row[8]["task_name"] == "Two letters"
    assert by_row[10]["group_code"] == "D" and by_row[10]["group_title"] == "Spaced group"
    names = {r["task_name"] for r in items.values()}
    assert not names & {"Lower-case group", "Group b", "Spaced group", "Pre-Implement Activity"}
    assert {"A-1", "B-2", "D-4"} <= set(items)


def test_t1_import_test_progress_rows(edge_import):
    items = edge_import[TP]
    assert sorted(r["excel_row"] for r in items.values()) == [2, 3]
    assert not any(str(r.get("application") or "").lower() == "grand total" for r in items.values())
    no_only = next(r for r in items.values() if r["excel_row"] == 3)
    assert not no_only.get("application")


def test_t1_import_defect_rows(edge_import):
    assert sorted(r["excel_row"] for r in edge_import[DEF].values()) == [2, 4, 5]


def test_t1_default_workbook_import_unchanged(scfg, repo):
    """The standard fixture workbook still imports the same items in the same rows."""
    run_import(scfg, repo, editor="t")
    rows = rows_of(repo)
    assert rows[TP] == {"TP-1": 2, "TP-6": 3, "TP-12": 4, "TP-38": 5, "TP-35": 6}
    assert sorted(rows[DEF].values()) == [2, 3, 4] and set(rows[DEF]) == {"DEF-1", "DEF-5", "DEF-21"}
    assert sorted(rows[UAT].values()) == [4, 5, 7, 8, 10, 11, 12, 13]
    assert {"A-1", "A-2", "B-6", "B-7", "E-11", "E-x1"} <= set(rows[UAT])


@pytest.mark.parametrize("write_item,write_row", [("TP-35", 6), ("TP-38", 5)])
def test_t1_tp_deleted_last_row_above_grand_total_raises(scfg, repo, svc, write_item, write_row):
    """Test Progress: delete the last application (TP-35, row 6) -> Grand Total moves into row 6.
    RowsMoved whichever item is written, and the workbook is not changed at all."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    ws = book.wb[TP]
    c = cols_of(ws)
    ws.delete_rows(6)
    assert ws.cell(6, c["Application"]).value == "Grand Total"
    before = snapshot(book.wb)
    with pytest.raises(RowsMoved, match="row 6"):
        apply_writes(book, defs(), [CellWrite(TP, write_item, "remark", write_row, text("x"))], rows_of(repo))
    assert snapshot(book.wb) == before
    assert ws.cell(6, c["App ID"]).value is None and ws.cell(6, c["Remark"]).value is None
    book.close()


def test_t1_uat_lowercase_group_header_in_deleted_row_raises(scfg, repo):
    """The writer uses the same rule: a lower-case group header moving into a deleted task's row is a move."""
    from conftest import UAT_ROWS, build_workbook, uat_row
    build_workbook(scfg.excel_path, uat=UAT_ROWS + [uat_row("f", None, "Lower header", note="n")])
    run_import(scfg, repo, editor="t")
    s = SyncService(scfg, repo, opener=OpenpyxlWorkbook, sender=FakeSender())
    assert s.init_ids().result == "written"
    book = OpenpyxlWorkbook(scfg.excel_path)
    ws = book.wb[UAT]
    ws.delete_rows(13)                                        # E-x1, the last task
    assert ws.cell(13, cols_of(ws)["Group"]).value == "f"
    before = snapshot(book.wb)
    with pytest.raises(RowsMoved):
        apply_writes(book, defs(), [CellWrite(UAT, "E-11", "remark", 12, text("x"))], rows_of(repo))
    assert snapshot(book.wb) == before
    book.close()


# =========================================================================== T7: App ID never makes a row an item

def test_t7_rules_ignore_app_id():
    assert not is_item_row(UAT, {"group": None, "task_no": None, "task_name": None, "app_id": "OLD-9"})
    assert not is_item_row(DEF, {"no": 25, "detail": None, "app_id": "OLD-DEF-9"})
    assert not is_item_row(TP, {"no": None, "application": None, "app_id": "OLD-TP-9"})
    assert not is_defect_item({"no": 25, "app_id": "OLD-DEF-9"}) and not is_defect_item({"app_id": "X"})
    assert is_defect_item({"no": 25, "detail": "d", "app_id": "DEF-25"})
    assert is_item_row(UAT, {"group": None, "task_no": 1, "task_name": "t", "app_id": "A-1"})


def test_t7_normalize_skips_rows_with_only_app_id():
    uat = [RawRow(2, {"group": "A", "task_no": None, "task_name": "SETUP"}),
           RawRow(3, {"group": None, "task_no": 1, "task_name": "t1", "app_id": "A-1"}),
           RawRow(4, {"group": None, "task_no": None, "task_name": None, "app_id": "OLD-9"}),
           RawRow(5, {"group": None, "task_no": None, "task_name": "no number", "app_id": None})]
    assert [it.excel_row for it in normalize_uat_plan(uat, WINDOW)] == [3, 5]
    defects = [RawRow(2, {"no": 1, "detail": "d", "app_id": "DEF-1"}),
               RawRow(3, {"no": 25, "detail": None, "app_id": "OLD-DEF-9"}),
               RawRow(4, {"no": None, "detail": None, "app_id": "OLD-DEF-8"})]
    assert [it.excel_row for it in normalize_defects(defects, WINDOW)] == [2]


def _leave_only_unknown_app_ids(path) -> tuple[int, int]:
    """In the saved file: a UAT Plan row under the last group with only an unknown App ID, and the pre-numbered
    Defect row (No. 23) given an unknown App ID - e.g. visible cells cleared, the hidden App ID left behind."""
    wb = openpyxl.load_workbook(path)
    ws = wb[UAT]
    uat_row = ws.max_row + 1
    ws.cell(uat_row, cols_of(ws)["App ID"]).value = "OLD-UAT-9"
    ws = wb[DEF]
    assert ws.cell(5, 1).value == 23
    ws.cell(5, cols_of(ws)["App ID"]).value = "OLD-DEF-9"
    wb.save(path)
    return uat_row, 5


def test_t7_rows_with_only_unknown_app_id_not_imported(scfg, repo, svc):
    uat_row, def_row = _leave_only_unknown_app_ids(scfg.excel_path)
    run_import(scfg, repo, editor="t")
    rows = rows_of(repo)
    assert uat_row not in rows[UAT].values() and def_row not in rows[DEF].values()
    assert sorted(rows[UAT].values()) == [4, 5, 7, 8, 10, 11, 12, 13]
    assert sorted(rows[DEF].values()) == [2, 3, 4]


def test_t7_sync_writes_with_only_unknown_app_id_rows(scfg, repo, svc):
    """Before T7 such rows were imported as items but the writer saw no item there: RowsMoved on every flush."""
    uat_row, def_row = _leave_only_unknown_app_ids(scfg.excel_path)
    _edit_three_sheets(scfg, repo)
    fr = svc.flush(force=True)
    assert fr.result == "written", fr
    assert repo.all_pending() == []
    fr2 = svc.flush(force=True)                                # stays healthy, no RowsMoved later either
    assert fr2.result != "retry", fr2
    wb = openpyxl.load_workbook(scfg.excel_path)
    for sheet, item, note in ((TP, "TP-38", "tp38 note"), (UAT, "B-6", "b6 note"), (DEF, "DEF-5", "def5 note")):
        ws = wb[sheet]
        c = cols_of(ws)
        assert [ws.cell(r, c["Remark"]).value for r in row_by_id(ws, c, item)] == [note]
    ws = wb[UAT]
    assert ws.cell(uat_row, cols_of(ws)["App ID"]).value == "OLD-UAT-9"     # not an item: left alone
    ws = wb[DEF]
    assert ws.cell(def_row, cols_of(ws)["App ID"]).value == "OLD-DEF-9"
    assert ws.cell(def_row, cols_of(ws)["Remark"]).value is None
    assert_ids_consistent(scfg.excel_path, repo)


# --------------------------------------------------------------------------- T7 (QA): extra cases

ORPHAN_UAT = 8          # a row with only an unknown App ID, inside group B (between B-6 and B-7)


def _orphan_in_group_b(path):
    """Saved file: a row inserted between B-6 (row 7) and B-7 with only an unknown App ID in it."""
    wb = openpyxl.load_workbook(path)
    ws = wb[UAT]
    insert_row(ws, ORPHAN_UAT)
    ws.cell(ORPHAN_UAT, cols_of(ws)["App ID"]).value = "OLD-UAT-5"
    wb.save(path)


# mutations made in the open Excel after the read, around the orphan row (rows after _orphan_in_group_b)
ORPHAN_MUTATIONS = {
    "delete_item_above_orphan": lambda wb: wb[UAT].delete_rows(7),          # orphan slides into B-6's place
    "delete_orphan": lambda wb: wb[UAT].delete_rows(ORPHAN_UAT),            # B-7 and below move up
    "insert_above_orphan": lambda wb: insert_row(wb[UAT], ORPHAN_UAT),      # orphan, B-7 and below move down
    "insert_top": lambda wb: insert_row(wb[UAT], 2),
}


def test_t7_qa_orphan_in_middle_of_group_not_imported(scfg, repo, svc):
    _orphan_in_group_b(scfg.excel_path)
    res = run_import(scfg, repo, editor="t")
    rows = rows_of(repo)[UAT]
    assert ORPHAN_UAT not in rows.values()
    assert rows == {"A-1": 4, "A-2": 5, "B-6": 7, "B-7": 9, "E-8": 11, "E-8b": 12, "E-11": 13, "E-x1": 14}
    assert res.sheets[UAT].created == 0 and res.sheets[UAT].removed == 0
    assert len(repo.load_items(UAT)) == 8


@pytest.mark.parametrize("name", sorted(ORPHAN_MUTATIONS))
def test_t7_qa_orphan_in_group_then_rows_moved_raises_nothing_written(scfg, repo, svc, name):
    """T1 still guards the writer when an App-ID-only row sits between items and rows move after the read."""
    _orphan_in_group_b(scfg.excel_path)
    run_import(scfg, repo, editor="t")
    book = OpenpyxlWorkbook(scfg.excel_path)
    ORPHAN_MUTATIONS[name](book.wb)
    before = snapshot(book.wb)
    with pytest.raises(RowsMoved):
        apply_writes(book, defs(), [CellWrite(UAT, "B-7", "remark", 9, text("lost"))], rows_of(repo))
    assert snapshot(book.wb) == before
    book.close()


def test_t7_qa_orphan_in_group_unmoved_writes_normally(scfg, repo, svc):
    _orphan_in_group_b(scfg.excel_path)
    run_import(scfg, repo, editor="t")
    book = OpenpyxlWorkbook(scfg.excel_path)
    rep = apply_writes(book, defs(), [CellWrite(UAT, "B-7", "remark", 9, text("b7"))], rows_of(repo))
    assert rep.written == [(UAT, "B-7", "remark")] and not rep.failed and rep.ids_written == 0
    ws = book.wb[UAT]
    c = cols_of(ws)
    assert ws.cell(9, c["Remark"]).value == "b7" and ws.cell(ORPHAN_UAT, c["Remark"]).value is None
    assert ws.cell(ORPHAN_UAT, c["App ID"]).value == "OLD-UAT-5"
    book.close()


@pytest.mark.parametrize("name", sorted(ORPHAN_MUTATIONS))
def test_t7_qa_sync_orphan_in_group_moved_retry_then_written(scfg, repo, svc, name):
    _orphan_in_group_b(scfg.excel_path)
    edit_item(scfg, repo, UAT, "B-7", {"remark": "b7 note"}, actor=PMO_ACTOR)
    edit_item(scfg, repo, UAT, "E-11", {"remark": "e11 note"}, actor=PMO_ACTOR)
    edit_item(scfg, repo, TP, "TP-38", {"remark": "tp38 note"}, actor=PMO_ACTOR)
    disk = sha(scfg.excel_path)
    calls: dict = {}
    moving = SyncService(scfg, repo, opener=mutating_opener(ORPHAN_MUTATIONS[name], calls), sender=FakeSender())
    fr = moving.flush(force=True)
    assert fr.result == "retry", fr
    assert calls["saved"] is False and sha(scfg.excel_path) == disk
    assert all(p["status"] == "pending" for p in repo.all_pending()) and len(repo.all_pending()) == 3
    assert not moving.sender.sent
    # the file catches up; after the next read everything is written to the right rows
    wb = openpyxl.load_workbook(scfg.excel_path)
    ORPHAN_MUTATIONS[name](wb)
    wb.save(scfg.excel_path)
    svc.retry_now()
    fr2 = svc.flush(force=True)
    assert fr2.result == "written", fr2
    assert repo.all_pending() == []
    assert svc.flush(force=True).result != "retry"
    wb = openpyxl.load_workbook(scfg.excel_path)
    for sheet, item, note in ((UAT, "B-7", "b7 note"), (UAT, "E-11", "e11 note"), (TP, "TP-38", "tp38 note")):
        ws = wb[sheet]
        c = cols_of(ws)
        assert [ws.cell(r, c["Remark"]).value for r in row_by_id(ws, c, item)] == [note]
    ws = wb[UAT]
    c = cols_of(ws)
    orphan = row_by_id(ws, c, "OLD-UAT-5")
    if name == "delete_orphan":
        assert orphan == []
    else:
        assert len(orphan) == 1
        assert all(ws.cell(orphan[0], k).value is None for k in range(1, c["App ID"]))   # still only an App ID
    assert_ids_consistent(scfg.excel_path, repo)


# cleared visible cells, the hidden (known) App ID left behind -> same as deleting the row
CLEARED = [(TP, "TP-12", 4, "TP-38"), (UAT, "B-7", 8, "B-6"), (DEF, "DEF-5", 3, "DEF-21")]


def _clear_keep_id(path, sheet: str, row: int, mode: str):
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    if mode == "delete":
        ws.delete_rows(row)
    else:
        id_col = cols_of(ws)["App ID"]
        first = 2 if sheet == DEF else 1                     # Defect List: the pre-filled 'No.' stays too
        for k in range(first, id_col):
            ws.cell(row, k).value = None
    wb.save(path)


@pytest.mark.parametrize("mode", ["clear", "delete"])
@pytest.mark.parametrize("sheet,item,row,other", CLEARED)
def test_t7_qa_cleared_row_with_known_id_removed_like_deleted(scfg, repo, svc, sheet, item, row, other, mode):
    edit_item(scfg, repo, sheet, item, {"remark": "gone"}, actor=PMO_ACTOR)
    edit_item(scfg, repo, sheet, other, {"remark": "kept"}, actor=PMO_ACTOR)
    before_ids = set(rows_of(repo)[sheet])
    _clear_keep_id(scfg.excel_path, sheet, row, mode)
    res = run_import(scfg, repo, editor="t")
    assert res.sheets[sheet].removed == 1 and res.sheets[sheet].created == 0
    assert repo.get_item(sheet, item)["active"] in (False, 0)
    assert set(rows_of(repo)[sheet]) == before_ids - {item}
    assert any(a["item_id"] == item and a["new_value"] == "removed" for a in repo.audit_entries(item))
    # the writer: no RowsMoved now, nor on later flushes
    book = OpenpyxlWorkbook(scfg.excel_path)
    apply_writes(book, defs(), [], rows_of(repo))
    book.close()
    fr = svc.flush(force=True)
    assert fr.result == "written", fr
    pend = repo.all_pending()
    assert [(p["item_id"], p["status"]) for p in pend] == [(item, "failed")]
    assert svc.flush(force=True).result != "retry"
    ws = openpyxl.load_workbook(scfg.excel_path)[sheet]
    c = cols_of(ws)
    assert [ws.cell(r, c["Remark"]).value for r in row_by_id(ws, c, other)] == ["kept"]
    if mode == "clear":
        assert row_by_id(ws, c, item) == [row]                       # left alone, nothing written into it
        assert ws.cell(row, c["Remark"]).value is None
    else:
        assert row_by_id(ws, c, item) == []
    assert_ids_consistent(scfg.excel_path, repo)


@pytest.mark.parametrize("at", [4, 7, 9])          # between items / right above Grand Total / under Grand Total
def test_t7_qa_test_progress_row_with_only_app_id(scfg, repo, svc, at):
    wb = openpyxl.load_workbook(scfg.excel_path)
    ws = wb[TP]
    if at < 9:
        insert_row(ws, at)
    ws.cell(at, cols_of(ws)["App ID"]).value = "OLD-TP-9"
    wb.save(scfg.excel_path)
    res = run_import(scfg, repo, editor="t")
    assert res.sheets[TP].created == 0 and res.sheets[TP].removed == 0
    assert len(repo.load_items(TP)) == 5
    rows = rows_of(repo)[TP]
    assert at not in rows.values()
    edit_item(scfg, repo, TP, "TP-38", {"remark": "tp38 note"}, actor=PMO_ACTOR)
    edit_item(scfg, repo, TP, "TP-1", {"remark": "tp1 note"}, actor=PMO_ACTOR)
    fr = svc.flush(force=True)
    assert fr.result == "written", fr
    assert repo.all_pending() == []
    assert svc.flush(force=True).result != "retry"
    ws = openpyxl.load_workbook(scfg.excel_path)[TP]
    c = cols_of(ws)
    assert [ws.cell(r, c["Remark"]).value for r in row_by_id(ws, c, "TP-38")] == ["tp38 note"]
    assert [ws.cell(r, c["Remark"]).value for r in row_by_id(ws, c, "TP-1")] == ["tp1 note"]
    assert row_by_id(ws, c, "OLD-TP-9") == [at] and ws.cell(at, c["Remark"]).value is None
    assert_ids_consistent(scfg.excel_path, repo)


def test_t7_qa_test_progress_only_app_id_row_moved_after_read_retry(scfg, repo, svc):
    wb = openpyxl.load_workbook(scfg.excel_path)
    insert_row(wb[TP], 4)
    wb[TP].cell(4, cols_of(wb[TP])["App ID"]).value = "OLD-TP-9"
    wb.save(scfg.excel_path)
    edit_item(scfg, repo, TP, "TP-38", {"remark": "tp38 note"}, actor=PMO_ACTOR)
    moving = SyncService(scfg, repo, opener=mutating_opener(lambda b: b[TP].delete_rows(4)), sender=FakeSender())
    assert moving.flush(force=True).result == "retry"
    wb = openpyxl.load_workbook(scfg.excel_path)
    wb[TP].delete_rows(4)
    wb.save(scfg.excel_path)
    svc.retry_now()
    assert svc.flush(force=True).result == "written"
    ws = openpyxl.load_workbook(scfg.excel_path)[TP]
    c = cols_of(ws)
    assert [ws.cell(r, c["Remark"]).value for r in row_by_id(ws, c, "TP-38")] == ["tp38 note"]
    assert_ids_consistent(scfg.excel_path, repo)


# =========================================================================== T2

FIRST_FAIL = "2026-11-09T10:00:00+07:00"


def test_t2_interrupted_writing_with_attempts_is_written(scfg, repo, svc):
    edit_item(scfg, repo, TP, "TP-38", {"remark": "after crash"}, actor=PMO_ACTOR)
    with repo.transaction():
        repo.set_pending_state(TP, "TP-38", "remark", status="writing", attempts=2, first_failed_at=FIRST_FAIL,
                               next_attempt_at=None, last_error="locked")
    fr = svc.flush(force=True)
    assert fr.result == "written", fr
    assert repo.all_pending() == []
    ws = openpyxl.load_workbook(scfg.excel_path)[TP]
    assert ws.cell(tp_row_of(ws, "DMS-RAMC"), cols_of(ws)["Remark"]).value == "after crash"


def test_t2_interrupted_writing_keeps_first_failed_at(scfg, repo, svc):
    class Locked:
        def __init__(self, path):
            raise WorkbookLocked("locked")
    edit_item(scfg, repo, TP, "TP-38", {"remark": "x"}, actor=PMO_ACTOR)
    with repo.transaction():
        repo.set_pending_state(TP, "TP-38", "remark", status="writing", attempts=2, first_failed_at=FIRST_FAIL,
                               next_attempt_at=None, last_error="locked")
    locked = SyncService(scfg, repo, opener=Locked, sender=FakeSender())
    fr = locked.flush(force=True)
    assert fr.result in ("retry", "error"), fr                # it was due (tried), not "nothing"
    p = repo.pending(TP)[("TP-38", "remark")]
    assert p["attempts"] == 3 and p["first_failed_at"] == FIRST_FAIL and p["status"] == "pending"


def test_t2_interrupted_writing_reset_is_due_now(scfg, repo, svc):
    """After the reset the entry is due immediately (next_attempt_at = now), also when the flush then waits."""
    edit_item(scfg, repo, TP, "TP-38", {"remark": "x"}, actor=PMO_ACTOR)
    with repo.transaction():
        repo.set_pending_state(TP, "TP-38", "remark", status="writing", attempts=4, first_failed_at=FIRST_FAIL,
                               next_attempt_at="2099-01-01T00:00:00+07:00")
    assert svc.flush(force=False).result == "waiting"          # just edited: batching, but reset already done
    p = repo.pending(TP)[("TP-38", "remark")]
    assert p["status"] == "pending" and p["attempts"] == 4 and p["first_failed_at"] == FIRST_FAIL
    assert p["next_attempt_at"] and p["next_attempt_at"] <= svc._now()
    assert p in svc._due() or any(d["item_id"] == "TP-38" for d in svc._due())


def _new_app(scfg, repo):
    return add_application(scfg, repo, application="NewApp", report_by="QA X", planned_start="2026-11-10",
                           planned_finish="2026-11-20", total_tc=10, actor=PMO_ACTOR)


def _queue_field(repo, item_id, field, value):
    with repo.transaction():
        repo.upsert_pending(TP, item_id, field, value=value, excel_value=text(value),
                            now="2026-11-09T09:00:00+07:00", reason="t", editor="t", channel="app")


def test_t2_failed_new_row_fails_its_field_edits(scfg, repo, svc):
    tp_id = _new_app(scfg, repo)
    with repo.transaction():
        repo.set_pending_state(TP, tp_id, "(row)", status="failed", attempts=0, last_error="row failed")
    _queue_field(repo, tp_id, "remark", "late note")
    fr = svc.flush(force=True)
    assert fr.result != "nothing"
    p = repo.pending(TP)[(tp_id, "remark")]
    assert p["status"] == "failed"
    assert p["last_error"] and "row" in p["last_error"].lower()
    assert repo.pending(TP)[(tp_id, "(row)")]["status"] == "failed"     # the row entry itself unchanged
    assert svc.status()["failed"] >= 2


def test_t2_queued_new_row_still_carries_field_edits(scfg, repo, svc):
    tp_id = _new_app(scfg, repo)
    edit_item(scfg, repo, TP, tp_id, {"remark": "with the row"}, actor=PMO_ACTOR)
    fr = svc.flush(force=True)
    assert fr.result == "written", fr
    assert not [p for p in repo.all_pending() if p["status"] == "failed"]
    ws = openpyxl.load_workbook(scfg.excel_path)[TP]
    c = cols_of(ws)
    r = row_by_id(ws, c, tp_id)
    assert len(r) == 1 and ws.cell(r[0], c["Application"]).value == "NewApp"


class _Crash(Exception):
    """The process stops right after the new row was saved to Excel, before the read-back."""


def test_t2_interrupted_row_writing_already_in_file_not_inserted_twice(scfg, repo, svc, monkeypatch):
    """A ROW entry left 'writing' with attempts > 0 whose row is already in the file (the run stopped after
    saving, before pull(force=True)): the next flush reads the file first and must not insert the row again."""
    tp_id = _new_app(scfg, repo)
    real_pull = svc.pull

    def crash_on_read_back(force=False):
        if force:
            raise _Crash()
        return real_pull(force)
    monkeypatch.setattr(svc, "pull", crash_on_read_back)
    with pytest.raises(_Crash):
        svc.flush(force=True)
    monkeypatch.setattr(svc, "pull", real_pull)
    p = repo.pending(TP)[(tp_id, "(row)")]
    assert p["status"] == "writing"                                        # left behind by the stopped run
    with repo.transaction():
        repo.set_pending_state(TP, tp_id, "(row)", status="writing", attempts=2, first_failed_at=FIRST_FAIL,
                               next_attempt_at=None, last_error="locked")
    ws = openpyxl.load_workbook(scfg.excel_path)[TP]
    assert len(row_by_id(ws, cols_of(ws), tp_id)) == 1                     # the row is already in the file

    fr = svc.flush(force=True)
    assert fr.result in ("nothing", "written"), fr
    assert [q for q in repo.all_pending() if q["item_id"] == tp_id] == []
    ws = openpyxl.load_workbook(scfg.excel_path)[TP]
    c = cols_of(ws)
    assert len(row_by_id(ws, c, tp_id)) == 1                               # not inserted a second time
    apps = [ws.cell(r, c["Application"]).value for r in range(2, ws.max_row + 1)]
    assert apps.count("NewApp") == 1
    assert repo.get_item(TP, tp_id)["active"]
    assert_ids_consistent(scfg.excel_path, repo)


# =========================================================================== T3

def test_t3_no_membership_anywhere_shows_no_project_page_once(qapp, base, appdb, two, monkeypatch):  # noqa: F811
    from uat_tracker.services.profile import write_profile
    from uat_tracker.ui import app as app_mod
    from uat_tracker.ui.projects import NoProjectWindow
    a, b = two
    write_profile(base.profile_path, PERSON.email, first_name="Pim", last_name="Pmo", last_project=a["id"])
    tried: list[str] = []

    def no_member(cfg, repo):
        tried.append(cfg.project_id)
        if len(tried) > 10:
            raise RuntimeError("open_project keeps trying (endless loop)")
        return None
    monkeypatch.setattr(app_mod, "open_profile", no_member)
    monkeypatch.setattr(app_mod.AppController, "_form", lambda self, **k: None)
    ctl = app_mod.AppController(base, appdb, Path("config.toml"), start_sync=False)
    assert ctl.start()
    assert isinstance(ctl.win, NoProjectWindow)
    assert sorted(tried) == sorted([a["id"], b["id"]])         # each project tried exactly once
    assert ctl.ctx is None
    ctl.win.on_closed = None
    ctl.win.close()


def test_t3_failed_project_skipped_other_opened(qapp, base, appdb, two, monkeypatch):  # noqa: F811
    from uat_tracker.services.profile import open_profile, write_profile
    from uat_tracker.ui import app as app_mod
    a, b = two
    write_profile(base.profile_path, PERSON.email, first_name="Pim", last_name="Pmo", last_project=a["id"])
    tried: list[str] = []

    def only_b(cfg, repo):
        tried.append(cfg.project_id)
        if len(tried) > 10:
            raise RuntimeError("endless loop")
        return open_profile(cfg, repo) if cfg.project_id == b["id"] else None
    monkeypatch.setattr(app_mod, "open_profile", only_b)
    ctl = app_mod.AppController(base, appdb, Path("config.toml"), start_sync=False)
    assert ctl.start()
    assert tried == [a["id"], b["id"]]
    assert ctl.ctx is not None and ctl.ctx.cfg.project_id == b["id"]
    # later opening A directly again is still allowed (skip only lasts for that round) and ends on B again
    tried.clear()
    ctl.open_project(a["id"])
    assert tried == [a["id"], b["id"]] and ctl.ctx.cfg.project_id == b["id"]
    ctl.win.on_closed = None
    ctl._close_current()


# =========================================================================== T4

@pytest.fixture
def pcfg(cfg, repo):
    """A project config whose DB exists (repo fixture) with some data."""
    c = dataclasses.replace(cfg, project_id="p-test")
    run_import(c, repo, editor="t")
    return c


def _db_bytes(path: Path) -> bytes:
    con = sqlite3.connect(path)
    try:
        return "\n".join(con.iterdump()).encode()
    finally:
        con.close()


def _backup_zip(tmp_path, project_id, entries: dict[str, bytes]) -> Path:
    dbf = tmp_path / "src.db"
    con = sqlite3.connect(dbf)
    con.execute("CREATE TABLE marker (x)")
    con.commit()
    con.close()
    z = tmp_path / "backup.zip"
    with zipfile.ZipFile(z, "w") as zf:
        zf.writestr("manifest.json", json.dumps({"project_id": project_id, "created_at": "2026-11-12T10:00"}))
        zf.writestr("tracker.db", dbf.read_bytes())
        for name, data in entries.items():
            zf.writestr(zipfile.ZipInfo(name), data)
    with zipfile.ZipFile(z) as zf:                              # names kept exactly as written
        assert {n.replace("\\", "/") for n in entries} <= set(zf.namelist())
    return z


def _evil_names(tmp_path) -> list[str]:
    drive_abs = (tmp_path / "evil_abs.txt").as_posix()          # e.g. C:/Users/.../evil_abs.txt
    return [
        "attachments/../evil_up.txt",
        "attachments/../../evil_up2.txt",
        "attachments/TP-1/../../evil_up3.txt",
        "attachments/" + drive_abs,
        "attachments//evil_root.txt",
        "attachments/..\\evil_bs.txt",
    ]


@pytest.mark.parametrize("idx", range(6))
def test_t4_restore_refuses_entry_outside_attachment_dir(tmp_path, pcfg, repo, idx):
    evil = _evil_names(tmp_path)[idx]
    pcfg.attachment_dir.mkdir(parents=True, exist_ok=True)
    (pcfg.attachment_dir / "keep.txt").write_bytes(b"keep")
    z = _backup_zip(tmp_path, pcfg.project_id, {"attachments/TP-1/ok.pdf": b"ok", evil: b"EVIL"})
    db_before = _db_bytes(pcfg.db_path)
    files_before = sorted(p for p in tmp_path.rglob("*") if p.is_file() and p != z)
    with pytest.raises(ClearError, match="outside"):
        restore_backup(pcfg, z, editor="t")
    assert _db_bytes(pcfg.db_path) == db_before                               # DB not touched
    assert repo.count_items()[TP]["active"] == 5
    files_after = sorted(p for p in tmp_path.rglob("*") if p.is_file() and p != z)
    new = [p for p in files_after if p not in files_before and not p.name.startswith("tracker.db-")]
    assert new == [], new                                                     # no before-restore copy, no file
    assert not (pcfg.attachment_dir / "TP-1" / "ok.pdf").exists()            # nothing restored, not even good ones


def test_t4_restore_valid_nested_attachments_ok(tmp_path, pcfg, repo):
    repo.close()
    z = _backup_zip(tmp_path, pcfg.project_id, {"attachments/TP-1/sub/result.pdf": b"pdf",
                                                "attachments/TP-1/a/../b.pdf": b"b"})
    man = restore_backup(pcfg, z, editor="t")
    assert man["project_id"] == pcfg.project_id
    assert (pcfg.attachment_dir / "TP-1" / "sub" / "result.pdf").read_bytes() == b"pdf"
    assert (pcfg.attachment_dir / "TP-1" / "b.pdf").read_bytes() == b"b"


def test_t4_wrong_project_still_refused(tmp_path, pcfg, repo):
    z = _backup_zip(tmp_path, "other", {"attachments/../evil.txt": b"x"})
    with pytest.raises(ClearError, match="not 'p-test'"):
        restore_backup(pcfg, z, editor="t")


# =========================================================================== T5

NEW_PMO = "new.pmo@example.test"


def _conflict_copy(scfg, monkeypatch):
    monkeypatch.setenv("COMPUTERNAME", "PMO-PC")
    copy = scfg.excel_path.with_name(f"{scfg.excel_path.stem}-PMO-PC.xlsx")
    copy.write_bytes(b"x")
    return copy


def test_t5_alert_goes_to_pmo_email_changed_after_sync_started(scfg, repo, svc, monkeypatch):
    sender = svc.sender
    save_settings(repo, PMO_ACTOR, fail_pct=5, dry_run=True, pmo_email=NEW_PMO, email_time="08:30")
    copy = _conflict_copy(scfg, monkeypatch)
    edit_item(scfg, repo, TP, "TP-38", {"remark": "x"}, actor=PMO_ACTOR)
    assert svc.flush(force=True).result == "blocked"
    alerts = [m for m in sender.sent if "conflict copy" in m.subject]
    assert [m.to for m in alerts] == [NEW_PMO]
    assert svc.cfg.alerts.pmo_email == PMO                                    # the service's cfg is not mutated
    # changed again while still running -> the next alert follows
    copy.unlink()
    svc.unblock()
    with repo.transaction():
        repo.set_setting("settings.pmo_email", "third@example.test")
    _conflict_copy(scfg, monkeypatch)
    assert svc.flush(force=True).result == "blocked"
    assert [m.to for m in sender.sent if "conflict copy" in m.subject] == [NEW_PMO, "third@example.test"]


def test_t5_error_alert_uses_new_pmo_email(scfg, repo, svc):
    class Locked:
        def __init__(self, path):
            raise WorkbookLocked("locked")
    t = [dt.datetime(2026, 11, 9, 10, 0)]
    locked = SyncService(scfg, repo, opener=Locked, sender=FakeSender(), clock=lambda: t[0])
    edit_item(scfg, repo, TP, "TP-38", {"remark": "x"}, actor=PMO_ACTOR, now="2026-11-09T09:00:00+07:00")
    with repo.transaction():
        repo.set_setting("settings.pmo_email", NEW_PMO)
    for minutes in (0, 1, 2, 5, 10, 15, 30):
        t[0] = dt.datetime(2026, 11, 9, 10, 0) + dt.timedelta(minutes=minutes, seconds=1)
        locked.flush(force=True)
    errs = [m for m in locked.sender.sent if "Excel sync error" in m.subject]
    assert [m.to for m in errs] == [NEW_PMO]


def test_t5_no_override_uses_config_email(scfg, repo, svc, monkeypatch):
    _conflict_copy(scfg, monkeypatch)
    edit_item(scfg, repo, TP, "TP-38", {"remark": "x"}, actor=PMO_ACTOR)
    svc.flush(force=True)
    assert [m.to for m in svc.sender.sent if "conflict copy" in m.subject] == [PMO]


# =========================================================================== T6

UNUSED = [
    ("services/daily_job.py", "DRY"),
    ("services/views.py", "dataclass"),
    ("ui/main_window.py", "Qt"),
    ("ui/users.py", "QDialog"),
    ("ui/screens/dashboard.py", "QVBoxLayout"),
    ("ui/screens/dashboard.py", "QWidget"),
    ("ui/screens/settings.py", "fmt"),
]


def _imported_and_used(path: Path) -> tuple[set[str], set[str]]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    imported = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Import, ast.ImportFrom)) and node.col_offset == 0:
            for a in node.names:
                imported.add((a.asname or a.name).split(".")[0])
    used = {n.id for n in ast.walk(tree) if isinstance(n, ast.Name)}
    used |= {n.value.id for n in ast.walk(tree) if isinstance(n, ast.Attribute) and isinstance(n.value, ast.Name)}
    # names used in string annotations / __all__
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str) and node.value.isidentifier():
            used.add(node.value)
    return imported, used


@pytest.mark.parametrize("rel,name", UNUSED)
def test_t6_unused_import_removed_and_name_not_used(rel, name):
    imported, used = _imported_and_used(SRC / rel)
    assert name not in imported, f"{rel}: '{name}' still imported at module level"
    assert name not in used, f"{rel}: '{name}' is used but no longer imported (NameError)"


@pytest.mark.parametrize("module", ["uat_tracker.services.daily_job", "uat_tracker.services.views",
                                    "uat_tracker.ui.main_window", "uat_tracker.ui.users",
                                    "uat_tracker.ui.screens.dashboard", "uat_tracker.ui.screens.settings"])
def test_t6_modules_still_import(module):
    import importlib
    if module.startswith("uat_tracker.ui"):
        pytest.importorskip("PySide6")
    importlib.import_module(module)


@pytest.mark.parametrize("rel", sorted({r for r, _ in UNUSED}))
def test_t6_every_remaining_import_in_touched_files_is_used(rel):
    """No other unused module-level import left in the files of T6 (a later clean-up must not regress)."""
    imported, used = _imported_and_used(SRC / rel)
    unused = sorted(n for n in imported - used if n not in ("annotations", "__future__"))
    assert unused == [], f"{rel}: unused imports {unused}"


# =========================================================================== T8: App ID column read in one go

XL_ROWS = 1_048_576


class _ComCounter:
    def __init__(self):
        self.range_reads = 0          # Range(...).Value
        self.cell_reads = 0           # Cells(r, c).Value
        self.end_calls = 0
        self.find_calls = 0           # Columns(c).Find(...)
        self.used_range = 0


class _FakeComCell:
    def __init__(self, ws, row, col):
        self._ws, self.Row, self.Column = ws, row, col

    @property
    def Value(self):
        self._ws.n.cell_reads += 1
        return self._ws.value(self.Row, self.Column)


def _fake_find(ws, col, args, hidden=frozenset()):
    """Columns(col).Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection) as Excel does it:
    after the first cell, backwards -> wraps to row 1,048,576 and goes up. LookIn=xlFormulas sees hidden
    rows, xlValues does not. None when nothing is found."""
    what, after, look_in, look_at, order, direction = args
    assert (what, look_at, order, direction) == ("*", 2, 1, 2)       # xlPart, xlByRows, xlPrevious
    assert (after.Row, after.Column) == (1, col)                    # start after the first cell
    assert look_in in (-4123, -4163)                                # xlFormulas / xlValues
    rows = sorted({r for r, c in ws.filled() if c == col}, reverse=True)
    r = next((r for r in rows if look_in == -4123 or r not in hidden), None)
    return None if r is None else ws.Cells(r, col)


class _FakeComColumn:
    def __init__(self, ws, col):
        self._ws, self._col = ws, col

    def Find(self, *args):
        self._ws.n.find_calls += 1
        if self._ws.FilterMode:                 # real Excel (T10): Find under a filter is unreliable (row 1)
            return self._ws.Cells(1, self._col)
        return _fake_find(self._ws, self._col, args)


class _FakeComRange:
    def __init__(self, ws, a, b):
        self._ws, self._a, self._b = ws, a, b

    @property
    def Value(self):
        self._ws.n.range_reads += 1
        rows = range(self._a.Row, self._b.Row + 1)
        cols = range(self._a.Column, self._b.Column + 1)
        assert len(rows) <= 10_000, "the scan must not read up to the end of UsedRange"
        if len(rows) == 1 and len(cols) == 1:
            return self._ws.value(self._a.Row, self._a.Column)     # COM: a single value, not a tuple
        return tuple(tuple(self._ws.value(r, c) for c in cols) for r in rows)


class FakeComWs:
    """Excel worksheet as seen through COM, backed by an openpyxl sheet. UsedRange reaches row 1,000,000
    (formatting applied to whole columns) while the data stops after a few rows."""

    def __init__(self, xl_ws, counter: _ComCounter, filter_mode: bool = False, used_rows: int = 1_000_000):
        self.xl, self.n, self.FilterMode, self._used = xl_ws, counter, filter_mode, used_rows
        self.ProtectContents = False
        self.Rows = type("Rows", (), {"Count": XL_ROWS})()

    def value(self, r, c):
        return self.xl.cell(r, c).value if r <= self.xl.max_row else None

    def filled(self):
        return [(c.row, c.column) for row in self.xl.iter_rows() for c in row if c.value not in (None, "")]

    def Cells(self, r, c):
        return _FakeComCell(self, r, c)

    def Columns(self, c):
        return _FakeComColumn(self, c)

    def Range(self, a, b):
        return _FakeComRange(self, a, b)

    @property
    def UsedRange(self):
        self.n.used_range += 1
        return type("UR", (), {"Row": 1, "Rows": type("R", (), {"Count": self._used})()})()


def fake_com_book(book: OpenpyxlWorkbook, counter: _ComCounter):
    from uat_tracker.excel.writer import ComSheet

    class Book:
        def sheet(self, name):
            return ComSheet(FakeComWs(book.wb[name], counter))
    return Book()


@pytest.mark.parametrize("sheet", [TP, UAT, DEF])
def test_t8_column_values_same_as_cell_by_cell_openpyxl(scfg, repo, svc, sheet):
    book = OpenpyxlWorkbook(scfg.excel_path)
    ws = book.sheet(sheet)
    for col in range(1, book.wb[sheet].max_column + 1):
        last = ws.last_data_row(col)
        assert ws.column_values(col, 2, last) == [ws.get(r, col) for r in range(2, last + 1)]
        assert ws.column_values(col, 3, 3) == [ws.get(3, col)]
        assert ws.column_values(col, 5, 4) == []
    book.close()


@pytest.mark.parametrize("sheet", [TP, UAT, DEF])
def test_t8_column_values_same_as_cell_by_cell_com(scfg, repo, svc, sheet):
    from uat_tracker.excel.writer import ComSheet
    book = OpenpyxlWorkbook(scfg.excel_path)
    xl = book.wb[sheet]
    ws = ComSheet(FakeComWs(xl, _ComCounter()))
    for col in range(1, xl.max_column + 1):
        assert ws.column_values(col, 2, xl.max_row) == [ws.get(r, col) for r in range(2, xl.max_row + 1)]
        assert ws.column_values(col, 4, 4) == [ws.get(4, col)]              # single cell: no tuple from COM
        assert ws.column_values(col, 5, 4) == []
    book.close()


def test_t8_last_data_row_com(scfg, repo, svc):
    from uat_tracker.excel.writer import ComSheet
    book = OpenpyxlWorkbook(scfg.excel_path)
    xl = book.wb[TP]
    id_col = cols_of(xl)["App ID"]
    last_id = max(r for r in range(1, xl.max_row + 1) if xl.cell(r, id_col).value not in (None, ""))
    n = _ComCounter()
    assert ComSheet(FakeComWs(xl, n)).last_data_row(id_col) == last_id
    assert n.used_range == 0                                    # not UsedRange (1,000,000 rows)
    empty_col = xl.max_column + 5
    assert ComSheet(FakeComWs(xl, n)).last_data_row(empty_col) == 1
    # AutoFilter on: End(xlUp) would stop at the last visible row -> UsedRange instead
    assert ComSheet(FakeComWs(xl, n, filter_mode=True, used_rows=40)).last_data_row(id_col) == 40
    # the very last row of the sheet holds a value: End(xlUp) from it would skip it
    big = FakeComWs(xl, n)
    big.value = lambda r, c: "X" if r == XL_ROWS else None
    big.filled = lambda: [(XL_ROWS, id_col)]
    assert ComSheet(big).last_data_row(id_col) == XL_ROWS
    book.close()


def _check(book, repo):
    from uat_tracker.excel.writer import _check_rows_unmoved
    try:
        _check_rows_unmoved(book, defs(), rows_of(repo))
    except RowsMoved as exc:
        return str(exc)
    return None


@pytest.mark.parametrize("name", [None] + sorted(MUTATIONS))
def test_t8_com_scan_same_result_as_before_and_few_com_calls(scfg, repo, svc, name, monkeypatch):
    """UsedRange 1,000,000 rows, little data: same RowsMoved result as the openpyxl (cell by cell) path,
    the App ID column read with one Range call per sheet and no cell-by-cell get for the scan."""
    from uat_tracker.excel.writer import ComSheet
    book = OpenpyxlWorkbook(scfg.excel_path)
    if name:
        MUTATIONS[name](book.wb)
    expected = _check(book, repo)
    assert (expected is None) == (name is None)
    gets = {"n": 0}
    real_get = ComSheet.get

    def counting_get(self, row, col):
        gets["n"] += 1
        return real_get(self, row, col)
    monkeypatch.setattr(ComSheet, "get", counting_get)
    n = _ComCounter()
    assert _check(fake_com_book(book, n), repo) == expected
    sheets = len(SHEETS)
    assert n.used_range == 0
    assert n.find_calls == sheets and n.end_calls == 0           # one Find per sheet (T9), no End(xlUp)
    assert n.range_reads == 2 * sheets                          # header row + App ID column, per sheet
    if name is None:
        assert gets["n"] == 0 and n.cell_reads == 0              # no cell-by-cell read at all
    else:                                                       # _holds_item on the few rows that look wrong
        assert gets["n"] <= 3 * 40
    book.close()


# --------------------------------------------------------------------------- T8 QA: simulated COM only (no real Excel)

QA_BUSY = -2147418111                       # RPC_E_CALL_REJECTED (signed, as pywin32 reports it)


class _QaComError(Exception):
    """Same shape as pywintypes.com_error: args = (hresult, text, excepinfo, argerror)."""

    def __init__(self, hresult):
        super().__init__(hresult, "Call was rejected by callee.", None, None)


class _QaCell:
    def __init__(self, ws, row, col):
        self._ws, self.Row, self.Column = ws, row, col

    @property
    def Value(self):
        self._ws.tick("cell")
        return self._ws.value(self.Row, self.Column)

    def End(self, direction):
        assert direction == -4162
        self._ws.tick("end")
        # Like Excel: End(xlUp) does not see any hidden row (filter, hidden by hand, collapsed group)
        r = next((r for r in range(self.Row - 1, 0, -1)
                  if r not in self._ws.hidden
                  and self._ws.value(r, self.Column) not in (None, "")), 1)
        return _QaCell(self._ws, r, self.Column)


class _QaColumn:
    def __init__(self, ws, col):
        self._ws, self._col = ws, col

    def Find(self, *args):
        self._ws.tick("find")
        if self._ws.filter_on:                  # real Excel (T10): Find under a filter is unreliable (row 1)
            return _QaCell(self._ws, 1, self._col)
        if self._ws.protect_on and self._ws.formula_hidden:
            return None                         # real Excel (T10): protected + FormulaHidden -> Find sees nothing
        return _fake_find(self._ws, self._col, args, self._ws.hidden)


class _QaRange:
    def __init__(self, ws, a, b):
        self._ws, self._a, self._b = ws, a, b

    @property
    def Value(self):
        self._ws.tick("range")
        self._ws.range_calls.append((self._a.Row, self._a.Column, self._b.Row, self._b.Column))
        rows = range(self._a.Row, self._b.Row + 1)
        cols = range(self._a.Column, self._b.Column + 1)
        if len(rows) == 1 and len(cols) == 1:
            return self._ws.value(self._a.Row, self._a.Column)      # COM: a scalar for one cell
        return tuple(tuple(self._ws.value(r, c) for c in cols) for r in rows)


class QaComWs:
    """Worksheet through COM: values from {(row, col): value}; hidden rows (by the filter when filter_mode,
    else hidden by hand / collapsed group), UsedRange (first row, row count) and busy failures:
    busy = {"what": n} makes the next n such calls busy.
    As in real Excel (T10 smoke test): under a filter Find returns an unreliable row (row 1); on a protected
    sheet (protect) whose cells are FormulaHidden (formula_hidden) Find returns None."""

    def __init__(self, data: dict, filter_mode=False, hidden=(), used=(1, 1_000_000), busy=None,
                 protect=False, formula_hidden=False):
        self.data, self.filter_on, self.hidden, self.used = dict(data), filter_mode, set(hidden), used
        self.protect_on, self.formula_hidden = protect, formula_hidden
        self.busy = dict(busy or {})
        self.calls: dict[str, int] = {}
        self.range_calls: list[tuple] = []
        self.Rows = type("Rows", (), {"Count": XL_ROWS})()

    def tick(self, what):
        self.calls[what] = self.calls.get(what, 0) + 1
        if self.busy.get(what, 0) > 0:
            self.busy[what] -= 1
            raise _QaComError(QA_BUSY)

    def value(self, r, c):
        return self.data.get((r, c))

    def filled(self):
        return [k for k, v in self.data.items() if v not in (None, "")]

    def Columns(self, c):
        return _QaColumn(self, c)

    @property
    def FilterMode(self):
        self.tick("filter")
        return self.filter_on

    @property
    def ProtectContents(self):
        self.tick("protect")
        return self.protect_on

    def Cells(self, r, c):
        return _QaCell(self, r, c)

    def Range(self, a, b):
        return _QaRange(self, a, b)

    @property
    def UsedRange(self):
        self.tick("used")
        first, count = self.used
        return type("UR", (), {"Row": first, "Rows": type("R", (), {"Count": count})()})()


def _qa_from_openpyxl(xl) -> dict:
    return {(c.row, c.column): c.value for row in xl.iter_rows() for c in row if c.value is not None}


@pytest.fixture
def qa_no_sleep(monkeypatch):
    import uat_tracker.excel.writer as W
    slept = []
    monkeypatch.setattr(W, "_sleep", slept.append)
    return slept


# small hand-made Test Progress sheet: No.=1, Application=2, App ID=3
_QA_TP_COLS = {"no": 1, "application": 2, APP_ID_KEY: 3}


def _qa_moved(ws, expected):
    from uat_tracker.excel.writer import ComSheet, _moved_rows
    return _moved_rows(ComSheet(ws), TP, 3, expected, _QA_TP_COLS)


# 1. filter on + rows moved -> RowsMoved still right

@pytest.mark.parametrize("name", [None] + sorted(MUTATIONS))
def test_t8_qa_filter_on_same_rows_moved_result(scfg, repo, svc, name):
    """AutoFilter hides the bottom rows of every sheet: same RowsMoved result as the cell-by-cell openpyxl path."""
    from uat_tracker.excel.writer import ComSheet
    book = OpenpyxlWorkbook(scfg.excel_path)
    if name:
        MUTATIONS[name](book.wb)
    expected = _check(book, repo)
    assert (expected is None) == (name is None)

    class Book:
        def sheet(self, sheet_name):
            xl = book.wb[sheet_name]
            hidden = set(range(max(2, xl.max_row - 3), xl.max_row + 1))
            return ComSheet(QaComWs(_qa_from_openpyxl(xl), filter_mode=True, hidden=hidden,
                                    used=(1, xl.max_row + 2)))
    assert _check(Book(), repo) == expected
    book.close()


def test_t8_qa_filter_hides_moved_row_still_detected():
    """Row 2 now holds a new item (no ID yet); X was sorted to row 10, which the filter hides.
    End(xlUp) would stop at row 2 and miss X; with FilterMode the scan must reach row 10."""
    data = {(1, 1): "No.", (1, 2): "Application", (1, 3): "App ID",
            (2, 1): 1, (2, 2): "NEW APP",
            (10, 1): 2, (10, 2): "Nleads", (10, 3): "X"}
    ws = QaComWs(data, filter_mode=True, hidden={10}, used=(1, 12))
    assert _qa_moved(ws, {"X": 2}) == ["X expected in row 2 but found in row 10"]
    assert ws.calls.get("end", 0) == 0 and ws.calls["used"] == 1
    assert ws.range_calls == [(2, 3, 12, 3)]                                 # one read, rows 2..12


def test_t8_qa_filter_used_range_not_from_row_1():
    """UsedRange starting below row 1: last row = Row + Rows.Count - 1."""
    from uat_tracker.excel.writer import ComSheet
    assert ComSheet(QaComWs({(7, 3): "X"}, filter_mode=True, used=(3, 5))).last_data_row(3) == 7


def test_t8_qa_filter_off_does_not_use_used_range():
    ws = QaComWs({(2, 1): 1, (2, 2): "A", (2, 3): "X"}, filter_mode=False, used=(1, 1_000_000))
    assert _qa_moved(ws, {"X": 2}) == []
    assert ws.calls.get("used", 0) == 0
    assert ws.range_calls == [(2, 3, 2, 3)]


# 2. App ID column with a single data row (COM returns a scalar)

def test_t8_qa_single_row_in_place():
    ws = QaComWs({(1, 3): "App ID", (2, 1): 1, (2, 2): "A", (2, 3): "X"})
    assert _qa_moved(ws, {"X": 2}) == []
    assert ws.calls["range"] == 1 and ws.calls.get("cell", 0) == 0 and ws.calls["find"] == 1


def test_t8_qa_single_row_id_cleared_but_row_still_item():
    assert _qa_moved(QaComWs({(2, 1): 1, (2, 2): "A"}), {"X": 2}) == []      # blank ID on an item row


def test_t8_qa_single_row_emptied():
    assert _qa_moved(QaComWs({}), {"X": 2}) == ["row 2 no longer holds an item, expected X"]


def test_t8_qa_single_row_unknown_id_and_spaces():
    assert _qa_moved(QaComWs({(2, 1): 1, (2, 2): "A", (2, 3): "ZZZ"}), {"X": 2}) == []
    assert _qa_moved(QaComWs({(2, 1): 1, (2, 2): "A", (2, 3): "  X  "}), {"X": 2}) == []


def test_t8_qa_single_cell_column_values_scalar_types():
    from uat_tracker.excel.writer import ComSheet
    for v in (None, "", "X", 42, 3.5):
        ws = QaComWs({(5, 3): v} if v is not None else {})
        assert ComSheet(ws).column_values(3, 5, 5) == [v]
    assert ComSheet(QaComWs({})).column_values(3, 6, 5) == []


def test_t8_qa_single_row_moved_down():
    """The only item; a row inserted above it: its ID is now in row 3 -> RowsMoved."""
    ws = QaComWs({(3, 1): 1, (3, 2): "A", (3, 3): "X"})
    assert _qa_moved(ws, {"X": 2}) == ["X expected in row 2 but found in row 3"]
    assert ws.range_calls == [(2, 3, 3, 3)]


# 3. sparse data / blank rows before the highest expected row

def test_t8_qa_read_always_covers_max_expected_row():
    """Data stops at row 3 but an item is expected at row 800 (its row was cleared): the read reaches 800."""
    ws = QaComWs({(2, 1): 1, (2, 2): "A", (2, 3): "X", (3, 1): 2, (3, 2): "B", (3, 3): "Y"})
    assert _qa_moved(ws, {"X": 2, "Y": 3, "Z": 800}) == ["row 800 no longer holds an item, expected Z"]
    assert ws.range_calls == [(2, 3, 800, 3)]


def test_t8_qa_sparse_far_apart_ids_in_place():
    ws = QaComWs({(2, 1): 1, (2, 2): "A", (2, 3): "X", (5000, 1): 2, (5000, 2): "B", (5000, 3): "Y"})
    assert _qa_moved(ws, {"X": 2, "Y": 5000}) == []
    assert ws.range_calls == [(2, 3, 5000, 3)] and ws.calls["range"] == 1


def test_t8_qa_id_below_expected_rows_is_found():
    """The ID sits far below every expected row: the read reaches the real last data row."""
    ws = QaComWs({(2, 1): 1, (2, 2): "A", (9000, 1): 2, (9000, 2): "B", (9000, 3): "X"})
    assert _qa_moved(ws, {"X": 2}) == ["X expected in row 2 but found in row 9000"]
    assert ws.range_calls == [(2, 3, 9000, 3)]


def test_t8_qa_swap_across_long_blank_gap():
    ws = QaComWs({(2, 1): 1, (2, 2): "A", (2, 3): "Y", (700, 1): 2, (700, 2): "B", (700, 3): "X"})
    assert _qa_moved(ws, {"X": 2, "Y": 700}) == ["X expected in row 2 but found in row 700",
                                                 "Y expected in row 700 but found in row 2"]


def test_t8_qa_last_row_of_sheet_filled():
    """App ID in row 1,048,576: End(xlUp) from that cell would skip it."""
    from uat_tracker.excel.writer import ComSheet
    ws = QaComWs({(XL_ROWS, 3): "X"})
    assert ComSheet(ws).last_data_row(3) == XL_ROWS
    assert ws.calls.get("end", 0) == 0 and ws.calls["find"] == 1


def test_t8_qa_empty_expected_reads_nothing():
    ws = QaComWs({(2, 3): "X"})
    assert _qa_moved(ws, {}) == []
    assert ws.calls == {}


# 4. Excel busy during column_values / last_data_row

@pytest.mark.parametrize("what", ["range", "find", "filter"])
def test_t8_qa_busy_once_then_ok(qa_no_sleep, what):
    ws = QaComWs({(2, 1): 1, (2, 2): "A", (2, 3): "X", (3, 1): 2, (3, 2): "B", (3, 3): "Y"}, busy={what: 1})
    assert _qa_moved(ws, {"X": 2, "Y": 3}) == []
    assert qa_no_sleep == [0.5]


def test_t8_qa_busy_once_used_range_under_filter(qa_no_sleep):
    ws = QaComWs({(2, 1): 1, (2, 2): "A", (2, 3): "X"}, filter_mode=True, used=(1, 4), busy={"used": 1})
    assert _qa_moved(ws, {"X": 2}) == []
    assert qa_no_sleep == [0.5]


@pytest.mark.parametrize("what", ["range", "find", "filter", "used"])
def test_t8_qa_busy_forever_becomes_workbook_busy(qa_no_sleep, what):
    from uat_tracker.excel.writer import WorkbookBusy
    ws = QaComWs({(2, 1): 1, (2, 2): "A", (2, 3): "X", (3, 1): 2, (3, 2): "B", (3, 3): "Y"},
                 filter_mode=(what == "used"), used=(1, 4), busy={what: 10_000})
    with pytest.raises(WorkbookBusy, match="Excel is busy"):
        _qa_moved(ws, {"X": 2, "Y": 3})
    assert 20 <= sum(qa_no_sleep) <= 30


def test_t8_qa_busy_short_budget_column_values_and_last_data_row(qa_no_sleep):
    from uat_tracker.excel.writer import ComSheet, WorkbookBusy, com_retry_budget
    sh = ComSheet(QaComWs({(2, 3): "X"}, busy={"range": 10_000, "find": 10_000}))
    with com_retry_budget(5):
        with pytest.raises(WorkbookBusy):
            sh.column_values(3, 2, 4)
        with pytest.raises(WorkbookBusy):
            sh.last_data_row(3)
    assert sum(qa_no_sleep) <= 10


def test_t8_qa_other_com_error_not_retried(qa_no_sleep):
    from uat_tracker.excel.writer import ComSheet

    class Broken(QaComWs):
        def Range(self, a, b):
            raise _QaComError(-2146827284)                                   # 0x800A03EC: a real error
    with pytest.raises(_QaComError):
        ComSheet(Broken({(2, 3): "X"})).column_values(3, 2, 3)
    assert qa_no_sleep == []


def test_t8_qa_busy_in_check_rows_unmoved_is_workbook_busy_not_rows_moved(scfg, repo, svc, qa_no_sleep):
    """Excel turns busy while the App ID column of Defect List is read: WorkbookBusy (queued), not RowsMoved."""
    from uat_tracker.excel.writer import ComSheet, WorkbookBusy, _check_rows_unmoved
    book = OpenpyxlWorkbook(scfg.excel_path)

    class Book:
        def sheet(self, sheet_name):
            xl = book.wb[sheet_name]
            ws = QaComWs(_qa_from_openpyxl(xl), used=(1, xl.max_row))
            if sheet_name == DEF:
                real_tick, ranges = ws.tick, {"n": 0}

                def tick(what):
                    if what == "range":
                        ranges["n"] += 1
                        if ranges["n"] > 1:                  # header read ok, App ID read stays busy
                            raise _QaComError(QA_BUSY)
                    real_tick(what)
                ws.tick = tick
            return ComSheet(ws)
    with pytest.raises(WorkbookBusy):
        _check_rows_unmoved(Book(), defs(), rows_of(repo))
    book.close()


# =========================================================================== T9: last data row incl. hidden rows

# Rows hidden without a sheet filter: hidden by hand, or a collapsed group (Outline). Through COM both are just
# Rows(r).Hidden = True with FilterMode False; End(xlUp) skips them, Find with LookIn=xlFormulas does not.
T9_KINDS = ["hidden_by_hand", "collapsed_group"]


def _t9_hidden(kind: str, rows: set[int]) -> set[int]:
    return set(rows) if kind == "hidden_by_hand" else set(rows) | {max(rows) + 1}   # group: a few more rows


@pytest.mark.parametrize("kind", T9_KINDS)
def test_t9_last_data_row_counts_hidden_rows(kind):
    from uat_tracker.excel.writer import ComSheet
    data = {(1, 3): "App ID", (2, 3): "A", (3, 3): "B", (4, 3): "C", (5, 3): "D"}
    ws = QaComWs(data, filter_mode=False, hidden=_t9_hidden(kind, {4, 5}))
    assert ws.Cells(XL_ROWS, 3).End(-4162).Row == 3                 # the fake: End(xlUp) misses rows 4-5
    ws.calls.clear()
    assert ComSheet(ws).last_data_row(3) == 5
    assert ws.calls == {"filter": 1, "protect": 1, "find": 1}       # three COM calls, no UsedRange, no End


@pytest.mark.parametrize("kind", T9_KINDS)
def test_t9_last_data_row_all_rows_hidden_and_empty_column(kind):
    from uat_tracker.excel.writer import ComSheet
    assert ComSheet(QaComWs({(1, 3): "App ID", (7, 3): "X"}, hidden=_t9_hidden(kind, {7}))).last_data_row(3) == 7
    assert ComSheet(QaComWs({(1, 3): "App ID"})).last_data_row(3) == 1      # header only
    assert ComSheet(QaComWs({(5, 2): "other column"})).last_data_row(3) == 1  # nothing found -> 1
    assert ComSheet(QaComWs({(XL_ROWS, 3): "X"}, hidden={XL_ROWS})).last_data_row(3) == XL_ROWS


@pytest.mark.parametrize("kind", T9_KINDS)
def test_t9_sorted_into_hidden_row_detected(kind):
    """Row 2 now holds a new item (no ID yet); X was sorted to row 10, which is hidden (no filter).
    End(xlUp) would stop at row 2 and miss X -> IDs written in the wrong rows (T1). Must be RowsMoved."""
    data = {(1, 1): "No.", (1, 2): "Application", (1, 3): "App ID",
            (2, 1): 1, (2, 2): "NEW APP",
            (10, 1): 2, (10, 2): "Nleads", (10, 3): "X"}
    ws = QaComWs(data, filter_mode=False, hidden=_t9_hidden(kind, {10}))
    assert _qa_moved(ws, {"X": 2}) == ["X expected in row 2 but found in row 10"]
    assert ws.range_calls == [(2, 3, 10, 3)]
    assert ws.calls.get("used", 0) == 0 and ws.calls.get("end", 0) == 0


@pytest.mark.parametrize("kind", T9_KINDS)
def test_t9_row_inserted_pushes_item_into_hidden_row(kind):
    """A row inserted above: X 2->3, Y 3->4 and row 4 is hidden."""
    data = {(2, 1): 1, (2, 2): "NEW",
            (3, 1): 2, (3, 2): "A", (3, 3): "X",
            (4, 1): 3, (4, 2): "B", (4, 3): "Y"}
    ws = QaComWs(data, hidden=_t9_hidden(kind, {4}))
    assert _qa_moved(ws, {"X": 2, "Y": 3}) == ["X expected in row 2 but found in row 3",
                                               "Y expected in row 3 but found in row 4"]


@pytest.mark.parametrize("kind", T9_KINDS)
def test_t9_hidden_rows_in_place_no_false_alarm(kind):
    data = {(2, 1): 1, (2, 2): "A", (2, 3): "X", (3, 1): 2, (3, 2): "B", (3, 3): "Y"}
    ws = QaComWs(data, hidden=_t9_hidden(kind, {3}))
    assert _qa_moved(ws, {"X": 2, "Y": 3}) == []
    assert ws.range_calls == [(2, 3, 3, 3)]


@pytest.mark.parametrize("kind", T9_KINDS)
@pytest.mark.parametrize("name", [None] + sorted(MUTATIONS))
def test_t9_hidden_rows_same_rows_moved_result(scfg, repo, svc, name, kind):
    """Bottom rows of every sheet hidden without a filter: same RowsMoved result as the openpyxl path."""
    from uat_tracker.excel.writer import ComSheet
    book = OpenpyxlWorkbook(scfg.excel_path)
    if name:
        MUTATIONS[name](book.wb)
    expected = _check(book, repo)
    assert (expected is None) == (name is None)

    class Book:
        def sheet(self, sheet_name):
            xl = book.wb[sheet_name]
            hidden = _t9_hidden(kind, set(range(max(2, xl.max_row - 3), xl.max_row + 1)))
            return ComSheet(QaComWs(_qa_from_openpyxl(xl), filter_mode=False, hidden=hidden))
    assert _check(Book(), repo) == expected
    book.close()


def test_t9_find_busy_retried_and_wrapped(qa_no_sleep):
    from uat_tracker.excel.writer import ComSheet, WorkbookBusy, com_retry_budget
    ws = QaComWs({(2, 3): "X", (6, 3): "Y"}, hidden={6}, busy={"find": 1})
    assert ComSheet(ws).last_data_row(3) == 6
    assert qa_no_sleep == [0.5]
    with com_retry_budget(5):
        with pytest.raises(WorkbookBusy):
            ComSheet(QaComWs({(2, 3): "X"}, busy={"find": 10_000})).last_data_row(3)


# --------------------------------------------------------------------------- T9 QA: simulated COM only (no real Excel)

XL_FORMULAS_DOC, XL_VALUES_DOC, XL_PART_DOC, XL_BY_ROWS_DOC, XL_PREVIOUS_DOC = -4123, -4163, 2, 1, 2


class _T9Column:
    def __init__(self, ws, col):
        self._ws, self._col = ws, col

    def Find(self, *args, **kwargs):
        self._ws.tick("find")
        self._ws.find_calls.append((self._col, args, kwargs))
        return self._ws.find(self._col, *args, **kwargs)


class T9QaWs(QaComWs):
    """QaComWs plus: whole hidden columns, Table (ListObject) filter rows, formula cells (formula text, cached
    value) and a Find that follows Excel's documented signature
    Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat):
    LookIn=xlFormulas sees rows hidden by hand / collapsed groups, hidden columns and formulas that return "",
    xlValues sees none of them. As in real Excel (T10 smoke test): a Table that filters rows (table_hidden)
    sets Worksheet.FilterMode True, and under any filter Find returns an unreliable row (row 1)."""

    def __init__(self, data: dict, *, hidden_cols=(), table_hidden=(), formulas=None, **kw):
        formulas = dict(formulas or {})                       # (row, col) -> (formula, cached value)
        data = dict(data)
        data.update({k: v for k, (_, v) in formulas.items()})
        super().__init__(data, **kw)
        self.hidden_cols, self.formulas = set(hidden_cols), formulas
        self.hidden |= set(table_hidden)
        if table_hidden:
            self.filter_on = True                             # a filtering Table sets ws.FilterMode (real Excel)
        self.find_calls: list[tuple] = []

    def Columns(self, c):
        return _T9Column(self, c)

    def find(self, col, What=None, After=None, LookIn=None, LookAt=None, SearchOrder=None, SearchDirection=None,
             MatchCase=None, MatchByte=None, SearchFormat=None):
        assert What == "*" and LookAt == XL_PART_DOC and SearchOrder == XL_BY_ROWS_DOC
        assert SearchDirection == XL_PREVIOUS_DOC and (After.Row, After.Column) == (1, col)
        assert LookIn in (XL_FORMULAS_DOC, XL_VALUES_DOC)
        if self.filter_on:
            return _QaCell(self, 1, col)                      # unreliable under a filter (real Excel)
        if self.protect_on and self.formula_hidden:
            return None                                       # protected + FormulaHidden (real Excel)

        def seen(r):
            if LookIn == XL_FORMULAS_DOC:
                return (r, col) in self.formulas or self.data.get((r, col)) not in (None, "")
            return (r not in self.hidden and col not in self.hidden_cols
                    and self.data.get((r, col)) not in (None, ""))
        rows = sorted({r for (r, c) in self.data if c == col} | {r for (r, c) in self.formulas if c == col},
                      reverse=True)
        r = next((r for r in rows if seen(r)), None)
        return None if r is None else _QaCell(self, r, col)


def _t9_sheet(**kw) -> T9QaWs:
    """Row 2 now holds a new item (no ID yet); X was sorted down to row 10."""
    data = {(1, 1): "No.", (1, 2): "Application", (1, 3): "App ID",
            (2, 1): 1, (2, 2): "NEW APP",
            (10, 1): 2, (10, 2): "Nleads", (10, 3): "X"}
    return T9QaWs(data, **kw)


# 4. positional arguments of Range.Find = Excel docs order

def test_t9_qa_find_positional_args_follow_excel_docs():
    """Excel: Range.Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte,
    SearchFormat). xlFormulas=-4123, xlPart=2, xlByRows=1, xlPrevious=2 (xlValues=-4163, xlNext=1)."""
    import uat_tracker.excel.writer as W
    assert (W.XL_FORMULAS, W.XL_PART, W.XL_BY_ROWS, W.XL_PREVIOUS) == (-4123, 2, 1, 2)
    ws = T9QaWs({(1, 3): "App ID", (4, 3): "X"})
    assert W.ComSheet(ws).last_data_row(3) == 4
    [(col, args, kwargs)] = ws.find_calls
    assert col == 3 and kwargs == {} and len(args) == 6
    what, after, look_in, look_at, order, direction = args
    assert what == "*"
    assert isinstance(after, _QaCell) and (after.Row, after.Column) == (1, 3)
    assert (look_in, look_at, order, direction) == (-4123, 2, 1, 2)


def test_t9_qa_find_with_xlvalues_would_miss_hidden_rows_so_fake_is_meaningful():
    """Guard for the fake itself: with xlValues the hidden row is not found, with xlFormulas it is."""
    ws = T9QaWs({(2, 3): "A", (9, 3): "X"}, hidden={9})
    after = ws.Cells(1, 3)
    assert ws.Columns(3).Find("*", after, XL_VALUES_DOC, 2, 1, 2).Row == 2
    assert ws.Columns(3).Find("*", after, XL_FORMULAS_DOC, 2, 1, 2).Row == 9


# 2. whole App ID column hidden / Table (ListObject) filter (a filtering Table sets FilterMode True)

def test_t9_qa_whole_id_column_hidden_last_row_and_rows_moved():
    """The app hides the App ID column itself (hide_column): Find xlFormulas still finds the last ID."""
    from uat_tracker.excel.writer import ComSheet
    ws = _t9_sheet(hidden_cols={3})
    assert ComSheet(ws).last_data_row(3) == 10
    ws = _t9_sheet(hidden_cols={3})
    assert _qa_moved(ws, {"X": 2}) == ["X expected in row 2 but found in row 10"]
    assert ws.range_calls == [(2, 3, 10, 3)]
    assert ws.calls.get("used", 0) == 0 and ws.calls.get("end", 0) == 0 and ws.calls["find"] == 1


def test_t9_qa_whole_id_column_hidden_and_bottom_rows_hidden():
    ws = _t9_sheet(hidden_cols={3}, hidden={9, 10, 11})
    assert _qa_moved(ws, {"X": 2}) == ["X expected in row 2 but found in row 10"]


def test_t9_qa_whole_id_column_hidden_in_place_no_false_alarm():
    data = {(2, 1): 1, (2, 2): "A", (2, 3): "X", (3, 1): 2, (3, 2): "B", (3, 3): "Y"}
    ws = T9QaWs(data, hidden_cols={3}, hidden={3})
    assert _qa_moved(ws, {"X": 2, "Y": 3}) == []
    assert ws.range_calls == [(2, 3, 3, 3)]


def test_t9_qa_table_filter_buttons_without_criteria_uses_find_only():
    """A Table with its filter buttons shown but no criteria filters nothing: Worksheet.FilterMode is False
    (real Excel). Row 10 is hidden by a collapsed group: only Find is used, it still reaches row 10 ->
    RowsMoved (no UsedRange, no End)."""
    ws = _t9_sheet(filter_mode=False, hidden={10})
    assert _qa_moved(ws, {"X": 2}) == ["X expected in row 2 but found in row 10"]
    assert ws.calls.get("used", 0) == 0 and ws.calls.get("end", 0) == 0
    assert ws.calls["filter"] == 1 and ws.calls["find"] == 1
    assert ws.range_calls == [(2, 3, 10, 3)]


def test_t9_qa_filtering_table_sets_filter_mode_uses_used_range():
    """A Table that filters row 10 out: Excel sets Worksheet.FilterMode True and Find there is unreliable
    (returns row 1): UsedRange is used, the read reaches row 10 -> RowsMoved (no Find, no End)."""
    ws = _t9_sheet(table_hidden={10}, used=(1, 12))
    assert ws.Columns(3).Find("*", ws.Cells(1, 3), XL_FORMULAS_DOC, 2, 1, 2).Row == 1    # the fake: unreliable
    ws.calls.clear()
    ws.find_calls.clear()
    assert _qa_moved(ws, {"X": 2}) == ["X expected in row 2 but found in row 10"]
    assert ws.calls["used"] == 1 and ws.calls.get("find", 0) == 0 and ws.calls.get("end", 0) == 0
    assert ws.range_calls == [(2, 3, 12, 3)]


def test_t9_qa_table_filter_and_hidden_id_column_insert_above():
    """Table filter hides rows 4-6, App ID column hidden, a row inserted above: every item moved one down."""
    data = {(2, 1): 1, (2, 2): "NEW",
            (3, 1): 2, (3, 2): "A", (3, 3): "X",
            (4, 1): 3, (4, 2): "B", (4, 3): "Y",
            (6, 1): 4, (6, 2): "C", (6, 3): "Z"}
    ws = T9QaWs(data, hidden_cols={3}, table_hidden={4, 5, 6}, used=(1, 6))
    assert _qa_moved(ws, {"X": 2, "Y": 3, "Z": 5}) == ["X expected in row 2 but found in row 3",
                                                       "Y expected in row 3 but found in row 4",
                                                       "Z expected in row 5 but found in row 6"]
    assert ws.range_calls == [(2, 3, 6, 3)]
    assert ws.calls["used"] == 1 and ws.calls.get("find", 0) == 0          # filtering Table: UsedRange


@pytest.mark.parametrize("variant", ["id_column_hidden", "table_filter", "both"])
@pytest.mark.parametrize("name", [None] + sorted(MUTATIONS))
def test_t9_qa_hidden_column_and_table_filter_same_rows_moved_result(scfg, repo, svc, name, variant):
    """Whole workbook: App ID column hidden and/or a Table filter (FilterMode True, as in real Excel) hiding the
    bottom rows: same RowsMoved result as the openpyxl (cell by cell) path; never End; UsedRange only when a
    Table filters (Find is unreliable there)."""
    from uat_tracker.excel.writer import ComSheet
    book = OpenpyxlWorkbook(scfg.excel_path)
    if name:
        MUTATIONS[name](book.wb)
    expected = _check(book, repo)
    assert (expected is None) == (name is None)
    made: list[T9QaWs] = []

    class Book:
        def sheet(self, sheet_name):
            xl = book.wb[sheet_name]
            id_col = cols_of(xl).get("App ID")
            hide_col = {id_col} if id_col and variant != "table_filter" else set()
            table = (set(range(max(2, xl.max_row - 3), xl.max_row + 1)) if variant != "id_column_hidden"
                     else set())
            ws = T9QaWs(_qa_from_openpyxl(xl), hidden_cols=hide_col, table_hidden=table, filter_mode=False,
                        used=(1, xl.max_row))
            made.append(ws)
            return ComSheet(ws)
    assert _check(Book(), repo) == expected
    assert made and all(ws.calls.get("end", 0) == 0 for ws in made)
    filtering = variant != "id_column_hidden"
    assert all((ws.calls.get("used", 0) == 1) is filtering and (ws.calls.get("find", 0) == 0) is filtering
               for ws in made)
    book.close()


# 3. a formula returning "" in the App ID column

def test_t9_qa_formula_empty_string_counts_as_data_row():
    """=IF(...,"") in row 50 of the ID column: Find xlFormulas sees it (last row 50); its value "" is no ID."""
    from uat_tracker.excel.writer import ComSheet
    ws = T9QaWs({(2, 3): "X"}, formulas={(50, 3): ('=IF(A50="","",A50)', "")})
    assert ComSheet(ws).last_data_row(3) == 50
    assert ComSheet(T9QaWs({}, formulas={(7, 3): ('=""', "")})).last_data_row(3) == 7


@pytest.mark.parametrize("case", ["in_place", "moved_down", "sorted_into_hidden", "row_emptied", "id_blank"])
def test_t9_qa_formula_empty_string_same_rows_moved_result(case):
    """Formulas returning "" below the data (and in the item's own ID cell for id_blank): RowsMoved result
    is the same as without the formulas."""
    cases = {"in_place": ({(2, 1): 1, (2, 2): "A", (2, 3): "X", (3, 1): 2, (3, 2): "B", (3, 3): "Y"},
                          {"X": 2, "Y": 3}, set()),
             "moved_down": ({(2, 1): 1, (2, 2): "NEW", (3, 1): 2, (3, 2): "A", (3, 3): "X"}, {"X": 2}, set()),
             "sorted_into_hidden": ({(2, 1): 1, (2, 2): "NEW", (10, 1): 2, (10, 2): "B", (10, 3): "X"},
                                    {"X": 2}, {10}),
             "row_emptied": ({(3, 1): 2, (3, 2): "B", (3, 3): "Y"}, {"X": 2, "Y": 3}, set()),
             "id_blank": ({(2, 1): 1, (2, 2): "A", (3, 1): 2, (3, 2): "B", (3, 3): "Y"}, {"X": 2, "Y": 3}, set())}
    data, expected, hidden = cases[case]
    want = _qa_moved(T9QaWs(data, hidden=hidden), expected)
    formulas = {(r, 3): (f'=IF(A{r}="","",A{r})', "") for r in range(20, 31)}
    if case == "id_blank":
        formulas[(2, 3)] = ('=""', "")
    ws = T9QaWs(data, hidden=hidden | {25, 30}, formulas=formulas)
    got = _qa_moved(ws, expected)
    assert got == want
    assert ws.range_calls == [(2, 3, 30, 3)]                   # read reaches the last formula row, one call
    if case == "in_place":
        assert got == []
    elif case in ("moved_down", "sorted_into_hidden"):
        assert got and got[0].startswith("X expected in row 2 but found in row")


# =========================================================================== T11
# A row's content cleared in the open Excel *after* the read, its hidden App ID left behind: a write to that
# item must not land in the now-empty row -> RowsMoved, then the next read makes the item inactive.

def _clear_row_keep_id(wb, sheet: str, row: int) -> None:
    """In memory: clear the visible cells of the row (Defect List keeps its pre-filled 'No.'), keep the App ID."""
    ws = wb[sheet]
    first = 2 if sheet == DEF else 1
    for k in range(first, cols_of(ws)["App ID"]):
        ws.cell(row, k).value = None


@pytest.mark.parametrize("sheet,item,row,other", CLEARED)
def test_t11_cleared_after_read_raises_nothing_written(scfg, repo, svc, sheet, item, row, other):
    book = OpenpyxlWorkbook(scfg.excel_path)
    _clear_row_keep_id(book.wb, sheet, row)
    before = snapshot(book.wb)
    writes = [CellWrite(sheet, item, "remark", row, text("lost")),
              CellWrite(sheet, other, "remark", rows_of(repo)[sheet][other], text("kept"))]
    with pytest.raises(RowsMoved, match=f"row {row} no longer holds an item"):
        apply_writes(book, defs(), writes, rows_of(repo))
    assert snapshot(book.wb) == before
    book.close()


@pytest.mark.parametrize("sheet,item,row,other", CLEARED)
def test_t11_cleared_after_read_not_written_is_not_checked(scfg, repo, svc, sheet, item, row, other):
    """Only rows written in this round are checked: a cleared row of an item with no write stays as before
    (no RowsMoved; the next read makes it inactive) and the other item's write goes through."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    _clear_row_keep_id(book.wb, sheet, row)
    other_row = rows_of(repo)[sheet][other]
    rep = apply_writes(book, defs(), [CellWrite(sheet, other, "remark", other_row, text("kept"))], rows_of(repo))
    assert rep.written == [(sheet, other, "remark")] and not rep.failed
    ws = book.wb[sheet]
    assert ws.cell(row, cols_of(ws)["Remark"]).value is None
    book.close()


@pytest.mark.parametrize("sheet,item,row,other", CLEARED)
def test_t11_sync_cleared_after_read_retry_then_inactive(scfg, repo, svc, sheet, item, row, other):
    edit_item(scfg, repo, sheet, item, {"remark": "gone"}, actor=PMO_ACTOR)
    edit_item(scfg, repo, sheet, other, {"remark": "kept"}, actor=PMO_ACTOR)
    disk = sha(scfg.excel_path)
    calls: dict = {}
    clear = lambda wb: _clear_row_keep_id(wb, sheet, row)               # noqa: E731
    moving = SyncService(scfg, repo, opener=mutating_opener(clear, calls), sender=FakeSender())
    fr = moving.flush(force=True)
    assert fr.result == "retry", fr
    assert calls["saved"] is False and sha(scfg.excel_path) == disk
    assert sorted(p["status"] for p in repo.all_pending()) == ["pending", "pending"]
    # the open Excel is saved with the cleared row; the next read makes the item inactive
    wb = openpyxl.load_workbook(scfg.excel_path)
    clear(wb)
    wb.save(scfg.excel_path)
    svc.retry_now()
    fr2 = svc.flush(force=True)
    assert fr2.result == "written", fr2
    assert repo.get_item(sheet, item)["active"] in (False, 0)
    assert [(p["item_id"], p["status"]) for p in repo.all_pending()] == [(item, "failed")]
    assert svc.flush(force=True).result != "retry"                      # no RowsMoved left behind
    ws = openpyxl.load_workbook(scfg.excel_path)[sheet]
    c = cols_of(ws)
    assert [ws.cell(r, c["Remark"]).value for r in row_by_id(ws, c, other)] == ["kept"]
    assert row_by_id(ws, c, item) == [row] and ws.cell(row, c["Remark"]).value is None
    assert_ids_consistent(scfg.excel_path, repo)


def test_t11_com_one_row_read_per_written_item(scfg, repo, svc):
    """Through COM: one extra Range read per item written (its whole row at once), none for the other rows,
    and no cell-by-cell read."""
    from uat_tracker.excel.writer import _check_rows_unmoved
    book = OpenpyxlWorkbook(scfg.excel_path)
    rows = rows_of(repo)
    writes = [CellWrite(TP, "TP-38", "remark", rows[TP]["TP-38"], text("a")),
              CellWrite(TP, "TP-38", "status", rows[TP]["TP-38"], text("b")),          # same row: read once
              CellWrite(UAT, "B-6", "remark", rows[UAT]["B-6"], text("c"))]
    n = _ComCounter()
    _check_rows_unmoved(fake_com_book(book, n), defs(), rows, writes)
    assert n.range_reads == 2 * len(SHEETS) + 2 and n.cell_reads == 0
    _clear_row_keep_id(book.wb, TP, rows[TP]["TP-38"])
    n = _ComCounter()
    with pytest.raises(RowsMoved, match="only its App ID TP-38 is left"):
        _check_rows_unmoved(fake_com_book(book, n), defs(), rows, writes)
    assert n.range_reads == 2 * len(SHEETS) + 2 and n.cell_reads == 0
    book.close()


def test_t11_qa_com_row_values_single_and_many_cells():
    from uat_tracker.excel.writer import ComSheet, _moved_rows
    ws = QaComWs({(2, 1): 1, (2, 2): "A", (2, 3): "X"})
    assert ComSheet(ws).row_values(2, 1, 2) == [1, "A"]
    assert ComSheet(ws).row_values(2, 2, 2) == ["A"]                    # one cell: COM returns a scalar
    assert ComSheet(ws).row_values(2, 3, 2) == []
    assert _moved_rows(ComSheet(ws), TP, 3, {"X": 2}, _QA_TP_COLS, frozenset({"X"})) == []
    emptied = QaComWs({(2, 3): "X"})
    assert _moved_rows(ComSheet(emptied), TP, 3, {"X": 2}, _QA_TP_COLS, frozenset({"X"})) == \
        ["row 2 no longer holds an item (only its App ID X is left)"]
    assert emptied.range_calls == [(2, 3, 2, 3), (2, 1, 2, 2)] and emptied.calls.get("cell", 0) == 0
    assert _moved_rows(ComSheet(QaComWs({(2, 3): "X"})), TP, 3, {"X": 2}, _QA_TP_COLS) == []   # not written


# --------------------------------------------------------------------------- T11 QA
# Through simulated COM: a cell holding a formula is seen as its value ("" for =IF(...,"",...), an int error
# code such as #DIV/0! = -2146826281); stale values left in some columns. The writer decides with
# is_item_row, the same rule the importer uses.

XL_DIV0 = -2146826281                       # what Range.Value returns for a #DIV/0! cell


class _T11RowBusyWs(QaComWs):
    """QaComWs whose one-row reads below the header (row_values) are busy the first `row_busy` times."""

    def __init__(self, data: dict, row_busy: int = 0):
        super().__init__(data)
        self.row_busy, self.row_reads = row_busy, 0

    def Range(self, a, b):
        rng = _QaRange(self, a, b)
        if not (a.Row == b.Row and a.Row > 1):
            return rng
        ws = self

        class _Row:
            @property
            def Value(self):
                ws.row_reads += 1
                if ws.row_busy > 0:
                    ws.row_busy -= 1
                    raise _QaComError(QA_BUSY)
                return rng.Value
        return _Row()


def _t11_qa_book(book: OpenpyxlWorkbook, sheet: str, row: int, leftover: dict, busy_row_reads: int = 0):
    """Fake COM book over the openpyxl workbook; `row` of `sheet` cleared to "" (formulas returning "") in every
    column before App ID, then `leftover` {header: value} put back (Defect List 'No.' = column 1)."""
    from uat_tracker.excel.writer import ComSheet
    data = {name: _qa_from_openpyxl(book.wb[name]) for name in SHEETS}
    c = cols_of(book.wb[sheet])
    for k in range(1, c["App ID"]):
        data[sheet][(row, k)] = ""
    for header, v in leftover.items():
        data[sheet][(row, 1 if (sheet == DEF and header == "No.") else c[header])] = v
    sheets = {name: _T11RowBusyWs(data[name], busy_row_reads if name == sheet else 0) for name in SHEETS}

    class Book:
        def sheet(self, name):
            return ComSheet(sheets[name])
    return Book(), sheets


UAT_TASK_NAME = "Task name"
# (sheet, item, row, leftover after the clear, still an item row?)
T11_LEFTOVERS = [
    (TP, "TP-12", 4, {}, False),                                         # every cell a formula returning ""
    (TP, "TP-12", 4, {"%Progress": XL_DIV0}, False),                     # =E/D -> #DIV/0! after the clear
    (TP, "TP-12", 4, {"Report By": "QA Ratana", "Status": "Complete"}, False),   # no Application, no No.
    (TP, "TP-12", 4, {"Application": " \xa0 "}, False),                  # nbsp only = blank
    (TP, "TP-12", 4, {"Application": "Grand Total", "Total TC": 1}, False),
    (TP, "TP-12", 4, {"No.": 12}, True),                                 # No. alone is an item (importer rule)
    (TP, "TP-12", 4, {"Application": "COLLINS"}, True),
    (UAT, "B-7", 8, {}, False),
    (UAT, "B-7", 8, {"Group": "C", UAT_TASK_NAME: "New group"}, False),  # turned into a group header
    (UAT, "B-7", 8, {"Group": "C"}, False),                              # group letter only = group header
    (UAT, "B-7", 8, {"Group": "C", "Task No": 7}, True),                 # has a Task No -> not a header
    (UAT, "B-7", 8, {"Responsible name": "QA X"}, True),                 # stale value -> still an item row
    (UAT, "B-7", 8, {"Remark": XL_DIV0}, True),                          # error value counts as content
    (DEF, "DEF-5", 3, {"No.": 5}, False),                                # pre-numbered: only No. left
    (DEF, "DEF-5", 3, {}, False),
    (DEF, "DEF-5", 3, {"No.": 5, "Status": "Close"}, True),
    (DEF, "DEF-5", 3, {"No.": 5, "Remark": XL_DIV0}, True),
]


@pytest.mark.parametrize("sheet,item,row,leftover,item_row", T11_LEFTOVERS)
def test_t11_qa_com_leftovers_follow_is_item_row(scfg, repo, svc, sheet, item, row, leftover, item_row):
    from uat_tracker.excel.writer import _check_rows_unmoved
    book = OpenpyxlWorkbook(scfg.excel_path)
    com, sheets = _t11_qa_book(book, sheet, row, leftover)
    rows = rows_of(repo)
    writes = [CellWrite(sheet, item, "remark", row, text("x"))]
    if item_row:
        _check_rows_unmoved(com, defs(), rows, writes)
    else:
        with pytest.raises(RowsMoved, match=f"row {row} no longer holds an item \\(only its App ID {item} is left"):
            _check_rows_unmoved(com, defs(), rows, writes)
    assert all(ws.calls.get("cell", 0) == 0 for ws in sheets.values())   # no cell-by-cell reads
    assert sheets[sheet].row_reads == 1
    book.close()


@pytest.mark.parametrize("sheet,item,row,leftover,item_row",
                         [c for c in T11_LEFTOVERS if XL_DIV0 not in c[3].values()])
def test_t11_qa_writer_decision_matches_importer(scfg, repo, svc, sheet, item, row, leftover, item_row):
    """The same leftovers saved to disk: the importer keeps the item active exactly when the writer lets the
    write through."""
    wb = openpyxl.load_workbook(scfg.excel_path)
    ws = wb[sheet]
    c = cols_of(ws)
    for k in range(1, c["App ID"]):
        ws.cell(row, k).value = None
    for header, v in leftover.items():
        ws.cell(row, 1 if (sheet == DEF and header == "No.") else c[header]).value = v
    wb.save(scfg.excel_path)
    book = OpenpyxlWorkbook(scfg.excel_path)
    try:
        apply_writes(book, defs(), [CellWrite(sheet, item, "remark", row, text("x"))], rows_of(repo))
        writer_ok = True
    except RowsMoved:
        writer_ok = False
    book.close()
    assert writer_ok is item_row
    run_import(scfg, repo, editor="t")
    assert bool(repo.get_item(sheet, item)["active"]) is item_row


# 2. a cell edit and a ROW write in the same round

def _row_write(item_id: str, anchor: int) -> CellWrite:
    spec = {"insert": True, "values": {"application": text("NewApp"),
                                       "no": json.dumps({"kind": "number", "value": 99})}}
    return CellWrite(TP, item_id, "(row)", anchor, json.dumps(spec))


def test_t11_qa_edit_and_row_write_anchor_cleared_nothing_written(scfg, repo, svc):
    """TP-38 has a remark edit and is the anchor of a new row; its content was cleared after the read:
    RowsMoved before anything, so the new row is not inserted either."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    rows = rows_of(repo)
    r = rows[TP]["TP-38"]
    _clear_row_keep_id(book.wb, TP, r)
    before = snapshot(book.wb)
    max_row = book.wb[TP].max_row
    writes = [CellWrite(TP, "TP-38", "remark", r, text("x")), _row_write("TP-99", r)]
    with pytest.raises(RowsMoved, match="only its App ID TP-38 is left"):
        apply_writes(book, defs(), writes, rows)
    assert snapshot(book.wb) == before and book.wb[TP].max_row == max_row
    book.close()


def test_t11_qa_edit_and_row_write_same_round_ok(scfg, repo, svc):
    book = OpenpyxlWorkbook(scfg.excel_path)
    rows = rows_of(repo)
    r = rows[TP]["TP-38"]
    writes = [CellWrite(TP, "TP-38", "remark", r, text("x")), _row_write("TP-99", r)]
    rep = apply_writes(book, defs(), writes, rows)
    assert sorted(rep.written) == [(TP, "TP-38", "remark"), (TP, "TP-99", "(row)")] and not rep.failed
    ws = book.wb[TP]
    c = cols_of(ws)
    assert row_by_id(ws, c, "TP-38") == [r] and ws.cell(r, c["Remark"]).value == "x"
    assert row_by_id(ws, c, "TP-99") == [r + 1] and ws.cell(r + 1, c["Application"]).value == "NewApp"
    book.close()


def test_t11_qa_row_write_only_is_not_a_cell_write(scfg, repo, svc):
    """A ROW write alone does not make its anchor (cleared, not edited) a written item: no RowsMoved."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    rows = rows_of(repo)
    r = rows[TP]["TP-38"]
    _clear_row_keep_id(book.wb, TP, r)
    rep = apply_writes(book, defs(), [_row_write("TP-99", r)], rows)
    assert rep.written == [(TP, "TP-99", "(row)")]
    ws = book.wb[TP]
    assert row_by_id(ws, cols_of(ws), "TP-99") == [r + 1]
    book.close()


def test_t11_qa_same_item_row_and_cell_write_checked(scfg, repo, svc):
    """Same item with a ROW entry and a cell edit in one round (the cell edit counts): cleared row -> RowsMoved."""
    book = OpenpyxlWorkbook(scfg.excel_path)
    rows = rows_of(repo)
    r = rows[TP]["TP-12"]
    _clear_row_keep_id(book.wb, TP, r)
    before = snapshot(book.wb)
    writes = [_row_write("TP-12", rows[TP]["TP-35"]), CellWrite(TP, "TP-12", "remark", r, text("x"))]
    with pytest.raises(RowsMoved, match="only its App ID TP-12 is left"):
        apply_writes(book, defs(), writes, rows)
    assert snapshot(book.wb) == before
    book.close()


def test_t11_qa_sync_new_app_and_edit_on_cleared_anchor(scfg, repo, svc):
    """Sync: a new application + an edit of TP-35 (the new row's anchor). TP-35 cleared in the open Excel ->
    retry, nothing saved. Once the file is saved that way: TP-35 inactive and its edit failed, the new row
    is written."""
    tp_id = _new_app(scfg, repo)
    edit_item(scfg, repo, TP, "TP-35", {"remark": "gone"}, actor=PMO_ACTOR)
    r = rows_of(repo)[TP]["TP-35"]
    disk = sha(scfg.excel_path)
    calls: dict = {}
    clear = lambda wb: _clear_row_keep_id(wb, TP, r)                    # noqa: E731
    fr = SyncService(scfg, repo, opener=mutating_opener(clear, calls), sender=FakeSender()).flush(force=True)
    assert fr.result == "retry", fr
    assert calls["saved"] is False and sha(scfg.excel_path) == disk
    wb = openpyxl.load_workbook(scfg.excel_path)
    clear(wb)
    wb.save(scfg.excel_path)
    svc.retry_now()
    fr2 = svc.flush(force=True)
    assert fr2.result == "written", fr2
    assert repo.get_item(TP, "TP-35")["active"] in (False, 0)
    assert [(p["item_id"], p["field"], p["status"]) for p in repo.all_pending()] == [("TP-35", "remark", "failed")]
    ws = openpyxl.load_workbook(scfg.excel_path)[TP]
    c = cols_of(ws)
    new = row_by_id(ws, c, tp_id)
    assert len(new) == 1 and ws.cell(new[0], c["Application"]).value == "NewApp"
    assert ws.cell(r, c["Remark"]).value is None
    assert_ids_consistent(scfg.excel_path, repo)


# 3. Excel busy during row_values

def test_t11_qa_row_values_busy_retried(qa_no_sleep):
    from uat_tracker.excel.writer import ComSheet
    ws = QaComWs({(2, 1): 1, (2, 2): "A"}, busy={"range": 2})
    assert ComSheet(ws).row_values(2, 1, 2) == [1, "A"]
    assert qa_no_sleep == [0.5, 1.0] and ws.calls["range"] == 3


def test_t11_qa_row_values_busy_too_long_workbook_busy(qa_no_sleep):
    from uat_tracker.excel.writer import ComSheet, WorkbookBusy, com_retry_budget
    ws = QaComWs({(2, 1): 1, (2, 2): "A"}, busy={"range": 100})
    with com_retry_budget(3.0), pytest.raises(WorkbookBusy, match="Excel is busy"):
        ComSheet(ws).row_values(2, 1, 2)
    assert qa_no_sleep == [0.5, 1.0]                      # 0.5 + 1.0 + 2.0 > 3.0


def test_t11_qa_row_values_non_busy_error_not_retried(qa_no_sleep):
    from uat_tracker.excel.writer import ComSheet

    class Ws(QaComWs):
        def Range(self, a, b):
            raise _QaComError(-2146827284)                # 0x800A03EC: not a busy code
    with pytest.raises(_QaComError):
        ComSheet(Ws({})).row_values(2, 1, 2)
    assert qa_no_sleep == []


def test_t11_qa_busy_in_row_check_retried_then_same_result(scfg, repo, svc, qa_no_sleep):
    from uat_tracker.excel.writer import _check_rows_unmoved
    book = OpenpyxlWorkbook(scfg.excel_path)
    com, sheets = _t11_qa_book(book, TP, 4, {}, busy_row_reads=2)
    with pytest.raises(RowsMoved, match="only its App ID TP-12 is left"):
        _check_rows_unmoved(com, defs(), rows_of(repo), [CellWrite(TP, "TP-12", "remark", 4, text("x"))])
    assert qa_no_sleep == [0.5, 1.0] and sheets[TP].row_reads == 3
    book.close()


def test_t11_qa_busy_in_row_check_write_workbook_busy_nothing_saved(scfg, repo, svc, qa_no_sleep):
    """Excel stays busy on the row read: write_workbook raises WorkbookBusy (the sync retries), nothing saved."""
    from uat_tracker.excel.writer import WorkbookBusy, write_workbook
    book = OpenpyxlWorkbook(scfg.excel_path)
    com, _ = _t11_qa_book(book, TP, 4, {"Application": "COLLINS"}, busy_row_reads=1000)
    state = {"saved": False, "closed": False}

    class Book:
        def sheet(self, name):
            return com.sheet(name)

        def save(self):
            state["saved"] = True

        def close(self):
            state["closed"] = True
    with pytest.raises(WorkbookBusy, match="Excel is busy"):
        write_workbook(lambda p: Book(), scfg.excel_path, defs(), [CellWrite(TP, "TP-12", "remark", 4, text("x"))],
                       rows_of(repo))
    assert state == {"saved": False, "closed": True}
    assert qa_no_sleep and sum(qa_no_sleep) <= 30.0
    book.close()


def test_t11_qa_sync_busy_on_row_read_retry(scfg, repo, svc, monkeypatch):
    """Through the sync: a busy row read -> retry result, the edit stays pending, the file is unchanged."""
    import uat_tracker.excel.writer as W
    edit_item(scfg, repo, TP, "TP-12", {"remark": "later"}, actor=PMO_ACTOR)
    disk = sha(scfg.excel_path)

    def busy(self, row, first_col, last_col):
        raise W.WorkbookBusy("Excel is busy (tried for 30 s); the change is written later")
    monkeypatch.setattr(W.OpenpyxlSheet, "row_values", busy)
    fr = svc.flush(force=True)
    assert fr.result == "retry", fr
    assert sha(scfg.excel_path) == disk
    assert [(p["item_id"], p["status"]) for p in repo.all_pending()] == [("TP-12", "pending")]


# =========================================================================== T10: real Excel smoke test findings
# Real Excel (24-Nov-2026): on a protected sheet whose App ID cells are FormulaHidden, Find returns None; under a
# filter (sheet AutoFilter or a filtering Table, both set FilterMode) Find is unreliable. UsedRange is used then.

def _t10_protected(formula_hidden: bool = True, **kw) -> QaComWs:
    """Protected sheet. Row 2 now holds a new item (no ID yet); X was sorted down to row 10, hidden by hand."""
    data = {(1, 1): "No.", (1, 2): "Application", (1, 3): "App ID",
            (2, 1): 1, (2, 2): "NEW APP",
            (10, 1): 2, (10, 2): "Nleads", (10, 3): "X"}
    return QaComWs(data, hidden={10}, used=(1, 12), protect=True, formula_hidden=formula_hidden, **kw)


def test_t10_protected_formula_hidden_find_none_uses_used_range():
    from uat_tracker.excel.writer import ComSheet
    ws = _t10_protected()
    assert ws.Columns(3).Find("*", ws.Cells(1, 3), -4123, 2, 1, 2) is None       # the fake, as in real Excel
    ws.calls.clear()
    assert ComSheet(ws).last_data_row(3) == 12
    assert ws.calls == {"filter": 1, "protect": 1, "used": 1}                   # no Find, no End


@pytest.mark.parametrize("formula_hidden", [True, False])
def test_t10_protected_sorted_into_hidden_row_is_rows_moved(formula_hidden):
    """Protected sheet (FormulaHidden or not, which cannot be read cheaply): UsedRange, X found in row 10."""
    ws = _t10_protected(formula_hidden)
    assert _qa_moved(ws, {"X": 2}) == ["X expected in row 2 but found in row 10"]
    assert ws.range_calls == [(2, 3, 12, 3)]
    assert ws.calls["used"] == 1 and ws.calls.get("find", 0) == 0 and ws.calls.get("end", 0) == 0


def test_t10_protected_in_place_no_false_alarm():
    data = {(2, 1): 1, (2, 2): "A", (2, 3): "X", (3, 1): 2, (3, 2): "B", (3, 3): "Y"}
    ws = QaComWs(data, hidden={3}, used=(1, 3), protect=True, formula_hidden=True)
    assert _qa_moved(ws, {"X": 2, "Y": 3}) == []


def test_t10_unprotected_sheet_still_uses_find():
    from uat_tracker.excel.writer import ComSheet
    ws = _t10_protected()
    ws.protect_on = False
    assert ComSheet(ws).last_data_row(3) == 10
    assert ws.calls == {"filter": 1, "protect": 1, "find": 1}


def test_t10_filter_find_unreliable_uses_used_range():
    """Sheet AutoFilter on: the fake's Find returns row 1 (as real Excel can); FilterMode -> UsedRange, so the
    unreliable Find is never asked and ProtectContents is not read either."""
    from uat_tracker.excel.writer import ComSheet
    data = {(2, 1): 1, (2, 2): "NEW APP", (10, 1): 2, (10, 2): "Nleads", (10, 3): "X"}
    ws = QaComWs(data, filter_mode=True, hidden={10}, used=(1, 12))
    assert ws.Columns(3).Find("*", ws.Cells(1, 3), -4123, 2, 1, 2).Row == 1      # the fake, as in real Excel
    ws.calls.clear()
    assert ComSheet(ws).last_data_row(3) == 12
    assert ws.calls == {"filter": 1, "used": 1}


def test_t10_protect_contents_busy_retried(qa_no_sleep):
    from uat_tracker.excel.writer import ComSheet, WorkbookBusy, com_retry_budget
    assert ComSheet(_t10_protected(busy={"protect": 1})).last_data_row(3) == 12
    assert qa_no_sleep == [0.5]
    with com_retry_budget(5):
        with pytest.raises(WorkbookBusy):
            ComSheet(_t10_protected(busy={"protect": 10_000})).last_data_row(3)


@pytest.mark.parametrize("name", [None] + sorted(MUTATIONS))
def test_t10_protected_formula_hidden_same_rows_moved_result(scfg, repo, svc, name):
    """Every sheet protected with FormulaHidden App ID cells and its bottom rows hidden: same RowsMoved result
    as the openpyxl (cell by cell) path."""
    from uat_tracker.excel.writer import ComSheet
    book = OpenpyxlWorkbook(scfg.excel_path)
    if name:
        MUTATIONS[name](book.wb)
    expected = _check(book, repo)
    assert (expected is None) == (name is None)

    class Book:
        def sheet(self, sheet_name):
            xl = book.wb[sheet_name]
            hidden = set(range(max(2, xl.max_row - 3), xl.max_row + 1))
            return ComSheet(QaComWs(_qa_from_openpyxl(xl), hidden=hidden, used=(1, xl.max_row),
                                    protect=True, formula_hidden=True))
    assert _check(Book(), repo) == expected
    book.close()


# --------------------------------------------------------------------------- T10/T12 QA

def test_t10_qa_protect_contents_busy_forever_in_moved_rows_is_workbook_busy(qa_no_sleep):
    """ProtectContents stays busy while the App ID column is checked: WorkbookBusy (queued), not RowsMoved."""
    from uat_tracker.excel.writer import WorkbookBusy
    ws = _t10_protected(busy={"protect": 10_000})
    with pytest.raises(WorkbookBusy, match="Excel is busy"):
        _qa_moved(ws, {"X": 2})
    assert 20 <= sum(qa_no_sleep) <= 30
    assert ws.calls.get("used", 0) == 0 and ws.calls.get("range", 0) == 0     # nothing read after the failure


def test_t10_qa_protect_contents_other_com_error_not_retried(qa_no_sleep):
    from uat_tracker.excel.writer import ComSheet

    class Broken(QaComWs):
        @property
        def ProtectContents(self):
            raise _QaComError(-2146827284)                                   # 0x800A03EC: not busy
    with pytest.raises(_QaComError):
        ComSheet(Broken({(2, 3): "X"})).last_data_row(3)
    assert qa_no_sleep == []


def test_t10_qa_protect_busy_through_write_workbook_nothing_saved(scfg, repo, svc, qa_no_sleep):
    """write_workbook: ProtectContents of Defect List stays busy -> WorkbookBusy, no cell written, not saved."""
    from uat_tracker.excel.writer import ComSheet, WorkbookBusy, write_workbook
    book = OpenpyxlWorkbook(scfg.excel_path)
    before = snapshot(book.wb)
    made: dict[str, QaComWs] = {}
    state = {"saved": False, "closed": False}

    class Book:
        def sheet(self, sheet_name):
            xl = book.wb[sheet_name]
            ws = QaComWs(_qa_from_openpyxl(xl), used=(1, xl.max_row), protect=True,
                         busy={"protect": 10_000} if sheet_name == DEF else None)
            made[sheet_name] = ws
            return ComSheet(ws)

        def save(self):
            state["saved"] = True

        def close(self):
            state["closed"] = True
    writes = [CellWrite(TP, "TP-38", "remark", 5, text("x"))]
    with pytest.raises(WorkbookBusy):
        write_workbook(lambda p: Book(), scfg.excel_path, defs(), writes, rows_of(repo))
    assert state == {"saved": False, "closed": True}
    assert snapshot(book.wb) == before
    book.close()


def _crash_after_new_row_saved(scfg, repo, svc, monkeypatch) -> str:
    """Add an application, flush; the process stops after the row was saved, before the read-back. The ROW
    entry is left 'writing' with attempts > 0 (as after a crash on a retried entry)."""
    tp_id = _new_app(scfg, repo)
    real_pull = svc.pull

    def crash_on_read_back(force=False):
        if force:
            raise _Crash()
        return real_pull(force)
    monkeypatch.setattr(svc, "pull", crash_on_read_back)
    with pytest.raises(_Crash):
        svc.flush(force=True)
    monkeypatch.setattr(svc, "pull", real_pull)
    with repo.transaction():
        repo.set_pending_state(TP, tp_id, "(row)", status="writing", attempts=2, first_failed_at=FIRST_FAIL,
                               next_attempt_at=None, last_error="locked")
    return tp_id


def _edit_excel_on_disk(path, how: str, tp_id: str) -> None:
    """The PMO edits the workbook in Excel and saves it between the crash and the next run."""
    wb = openpyxl.load_workbook(path)
    ws = wb[TP]
    c = cols_of(ws)
    [r] = row_by_id(ws, c, tp_id)
    if how == "insert_top":
        insert_row(ws, 2, [98, "TOP APP"])
    elif how == "insert_above_new_row":
        insert_row(ws, r, [98, "ABOVE APP"])
    elif how == "delete_first":
        ws.delete_rows(2)
    elif how == "sort_reverse":
        reverse_rows(ws, 2, r)
    elif how == "blank_row_above":
        ws.insert_rows(r)
    wb.save(path)


def _assert_no_misplaced_ids(path, repo) -> None:
    """Every App ID present in Excel sits in the row the DB has for that item (an ID may still be missing)."""
    wb = openpyxl.load_workbook(path)
    for name in SHEETS:
        ws = wb[name]
        c = cols_of(ws)
        db_rows = rows_of(repo)[name]
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, c["App ID"]).value
            if v and v in db_rows:
                assert db_rows[v] == r, f"{name}: {v} in row {r}, DB row {db_rows[v]}"


@pytest.mark.parametrize("how", ["insert_top", "insert_above_new_row", "delete_first", "sort_reverse",
                                 "blank_row_above"])
def test_t12_qa_row_writing_then_excel_edited_not_inserted_twice(scfg, repo, svc, monkeypatch, how):
    """ROW entry left 'writing' (attempts=2) whose row is already saved; then Excel is edited (rows inserted /
    deleted / sorted) before the next run: the row is not inserted again, nothing lands in a wrong row, and
    every App ID matches the row the DB has for it."""
    tp_id = _crash_after_new_row_saved(scfg, repo, svc, monkeypatch)
    _edit_excel_on_disk(scfg.excel_path, how, tp_id)

    fr = svc.flush(force=True)
    assert fr.result in ("nothing", "written"), fr
    assert [q for q in repo.all_pending() if q["item_id"] == tp_id] == []
    ws = openpyxl.load_workbook(scfg.excel_path)[TP]
    c = cols_of(ws)
    rows = row_by_id(ws, c, tp_id)
    assert len(rows) == 1                                                   # not inserted a second time
    assert ws.cell(rows[0], c["Application"]).value == "NewApp"             # its ID sits on its own row
    apps = [ws.cell(r, c["Application"]).value for r in range(2, ws.max_row + 1)]
    assert apps.count("NewApp") == 1
    rec = repo.get_item(TP, tp_id)
    assert rec["active"] and rec["excel_row"] == rows[0]
    _assert_no_misplaced_ids(scfg.excel_path, repo)
    # a row the PMO typed in Excel (e.g. TOP APP) has no App ID yet: filled on the next write, in its own row
    assert svc.init_ids().result in ("written", "nothing")
    assert_ids_consistent(scfg.excel_path, repo)


@pytest.mark.parametrize("how", ["insert_top", "sort_reverse"])
def test_t12_qa_row_writing_then_excel_edited_later_edit_goes_to_right_row(scfg, repo, svc, monkeypatch, how):
    """Same, and the item gets an edit afterwards: the value lands in the item's (moved) row only."""
    tp_id = _crash_after_new_row_saved(scfg, repo, svc, monkeypatch)
    _edit_excel_on_disk(scfg.excel_path, how, tp_id)
    assert svc.flush(force=True).result in ("nothing", "written")
    edit_item(scfg, repo, TP, tp_id, {"remark": "after restart"}, actor=PMO_ACTOR)
    fr = svc.flush(force=True)
    assert fr.result == "written", fr
    ws = openpyxl.load_workbook(scfg.excel_path)[TP]
    c = cols_of(ws)
    [r] = row_by_id(ws, c, tp_id)
    assert ws.cell(r, c["Application"]).value == "NewApp" and ws.cell(r, c["Remark"]).value == "after restart"
    assert [x for x in range(2, ws.max_row + 1) if ws.cell(x, c["Remark"]).value == "after restart"] == [r]
    assert_ids_consistent(scfg.excel_path, repo)
