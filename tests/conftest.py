"""Test fixtures. Builds synthetic workbooks shaped like the real 'UAT Planning.xlsx'
(header suffixes, duplicate Status columns, legend table, blank Defect 'No.' header, Grand Total)."""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import openpyxl
import pytest

from uat_tracker.config import Config
from uat_tracker.parsing.dates import DateWindow
from uat_tracker.services.auth import Actor, name_key
from uat_tracker.storage.sqlite_repo import SqliteRepository

WINDOW = DateWindow(dt.date(2026, 8, 1), dt.date(2027, 3, 31))
D = dt.datetime

UAT_HEADER = ["Task no.\n(Mandatory)", "Sub Task\n(Optional)", "Group", "Task No",
              "Task name\n(Mandatory)\nSI-4932,SR-56089", "Note", "File Path",
              "Responsible name\n(Mandatory)", "Status", "Start date\n(Mandatory)",
              "Start time\n(Mandatory)", "Finish date\n(Mandatory)", "Finish time\n(Mandatory)",
              "Remark\n(Optional)", "Status\n(Optional)", "Application\n(Optional)"]


def uat_row(group=None, no=None, name=None, owner=None, status=None, start=None, finish=None,
            note=None, finish_time=None, remark=None):
    return [None, None, group, no, name, note, None, owner, status, start, None, finish,
            finish_time, remark, None, None]


UAT_ROWS = [
    [1, None, None, None, "Pre-Implement Activity"] + [None] * 11,
    uat_row("A", None, "SETUP for UAT Testing", note="group note"),
    uat_row(None, 1, "16 Nleads Account", "QA Kittipong", "completed", D(2026, 9, 7), D(2026, 9, 9)),
    uat_row(None, 2, "Setup MPS Profile", "QA Suda", "comlete", "9/14/2569", D(2569, 9, 18)),
    uat_row("B", None, "Change Booking Branch"),
    uat_row(None, 6, "CCMS run retrieve data from EDW", "QA WANNA", "completed",
            "20/10/2026 25/9/2026 => 17/09/2026 14/9/2026",
            "10/10/2026 => 15/10/2026 => 21/10/2026 => 22/10/2026"),
    uat_row(None, 7, "No owner task", None, "In progress", D(2026, 10, 1), "24/09/206"),
    uat_row("E", None, "Downstream Systems/Report"),
    uat_row(None, 8, "CRD-TAX", "QA Anucha", "Not Start", "02/11/2026 (BD 30/09/26)", "29/09/2027"),
    uat_row(None, 8, "EFS - Run Payment File", "QA Prasit", "Inprogress", "10/8/2026", D(2026, 9, 25)),
    uat_row(None, 11, "DMS-RAMC", "QA Somsak", "Waiting", "10-Oct-2026\n20-Oct-2026",
            None, finish_time="06-Now-2026\n13-Nov-2026"),
    uat_row(None, None, "Operation Report from EDW", "Finance", "In progress", "30/10/2026", D(2026, 11, 13)),
]

TP_HEADER = ["No.", "Application", "Report By", "Total TC", "No. of TC Execution", "%Progress",
             "#Pass", "#Fail", "Planned Start", "Planned Finish", "Actual Start", "Actual Finish",
             "Traffic Light", "TC/Result Complete", "Status", "Remark", None, "Status", "Definition"]
TP_ROWS = [
    [1, "Nleads", "QA Kittipong", 17, 17, 1, 17, 0, D(2569, 9, 14), D(2569, 9, 18), "15/09/2569",
     D(2569, 9, 21), "Green", "Complete", "Complete", None, None, "Green", "On Track"],
    [6, "ATM", "QA Nida", 4, 4, 1, 2, 0, D(2026, 9, 21), D(2026, 10, 8), None, None,
     "Green", "Complete", None, None, None, "Amber", "Late"],
    [12, "COLLINS", "QA Ratana", 1, 1, 1, 1, 0, "24/09/206", D(2026, 9, 30), D(2026, 9, 25),
     D(2026, 9, 30), "Green", "Complete", "Complete", None, None, "Red", "Delay"],
    [38, "DMS-RAMC", "QA Somsak", 186, 126, 0.677, 124, 2, "1/4/1900\n22/10/2026",
     "11/06/2026\n13/11/2026", "22/10/2026", None, None, None, "Inprogress"],
    [35, "SCBPDM", "QA Prasit", 30, 30, 1, 30, None, D(2026, 9, 24), D(2026, 10, 30), None, None,
     "Green", None, "Complete"],
    [None, "Grand Total", None, 238, 178, 0.75, 174, 2],
]

DEF_HEADER = [None, "Date", "System/App.", "Issues/detail", "Status", "Severity", "Assigned To",
              "Fix target date", "Data test", "Raised by", "Remark"]
DEF_ROWS = [
    [1, D(2026, 9, 17), "Salesforce", "Branch not displayed", "Close", 3, "WILAI", D(2026, 9, 18),
     None, "CHANIDA", None],
    [5, D(2026, 9, 29), "Collins", "52 accounts", "Close", 3, None, None, None, "Ratana", None],
    [21, D(2026, 11, 3), "DMS-RAMC", "Year check", "Open", 3, "IT-DMS-RAMC", D(2026, 5, 11),
     None, "User-K.Pranee", None],
    [23],
    [24],
]


def build_workbook(path: Path, uat=None, tp=None, defects=None, tp_header=None) -> Path:
    wb = openpyxl.Workbook()
    other = wb.active
    other.title = "NLeadAccount"
    other["A1"] = "do not touch"
    for title, header, rows in (("UAT Plan", UAT_HEADER, uat if uat is not None else UAT_ROWS),
                                ("Defect List", DEF_HEADER, defects if defects is not None else DEF_ROWS),
                                ("Test Progress", tp_header or TP_HEADER, tp if tp is not None else TP_ROWS)):
        ws = wb.create_sheet(title)
        ws.append(header)
        for r in rows:
            ws.append(r)
    wb.save(path)
    return path


@pytest.fixture
def workbook(tmp_path) -> Path:
    return build_workbook(tmp_path / "UAT Planning.xlsx")


@pytest.fixture
def cfg(tmp_path, workbook) -> Config:
    return Config(excel_path=workbook, data_dir=tmp_path / "data",
                  window_start=WINDOW.start, window_end=WINDOW.end,
                  profile_file=tmp_path / "profile" / "profile.json")      # never the real one


@pytest.fixture(autouse=True)
def _no_real_excel_windows(monkeypatch):
    """The test workbook is also called 'UAT Planning.xlsx': a real Excel window on this PC must not
    make tests think it is open."""
    import uat_tracker.services.sync as sync_mod
    monkeypatch.setattr(sync_mod, "excel_windows_for", lambda path: [])


PMO_ACTOR = Actor("pmo@example.test", "PMO", "Pim Pmo")


def add_test_user(repo, email: str, role: str, aliases=(), first: str | None = None, last: str = "Test") -> Actor:
    """Part 5: a user straight in the DB (the services read the role from there)."""
    first = first or email.split("@")[0].capitalize()
    with repo.transaction():
        repo.insert_user(email=email, first_name=first, last_name=last, role=role,
                         created_at="2026-11-09T08:00:00+07:00", created_by="test")
        repo.set_user_aliases(email, {name_key(a): a for a in aliases})
    return Actor(email, role, f"{first} {last}")


@pytest.fixture
def repo(cfg):
    cfg.data_dir.mkdir(parents=True, exist_ok=True)
    r = SqliteRepository(cfg.db_path)
    add_test_user(r, PMO_ACTOR.email, "PMO", first="Pim", last="Pmo")
    yield r
    r.close()


@pytest.fixture(autouse=True)
def _no_project_name_leak():
    """Multi-project: opening a project sets the name used in email footers / PMO alert subjects."""
    from uat_tracker.alerts.templates import set_project
    set_project("")
    yield
    set_project("")
