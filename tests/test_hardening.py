"""Hardening phase 1, section D: logging, SQLite WAL, Excel busy retry, long actions in worker threads."""

from __future__ import annotations

import dataclasses
import logging
import os
import sqlite3
import subprocess
import sys
import textwrap
from pathlib import Path

import pytest

from conftest import PMO_ACTOR
from uat_tracker import logs
from uat_tracker.config import AlertsConfig
from uat_tracker.excel import writer as W
from uat_tracker.excel.writer import ComSheet, OpenpyxlWorkbook, WorkbookBusy, com_call, com_retry_budget
from uat_tracker.mail.sender import FakeSender
from uat_tracker.services.edits import edit_item
from uat_tracker.services.importer import run_import
from uat_tracker.services.sync import SyncBusy, SyncService, file_lock
from uat_tracker.storage.app_db import AppDb
from uat_tracker.storage.sqlite_repo import SqliteRepository

TP = "Test Progress"
BUSY = -2147418111                         # RPC_E_CALL_REJECTED as pywin32 reports it (signed)
RETRY_LATER = -2147417846                  # RPC_E_SERVERCALL_RETRYLATER


class FakeComError(Exception):
    """Same shape as pywintypes.com_error: args = (hresult, text, excepinfo, argerror)."""

    def __init__(self, hresult, text="Call was rejected by callee."):
        super().__init__(hresult, text, None, None)


# --------------------------------------------------------------------------- logging

@pytest.fixture
def logfile(tmp_path):
    path = logs.setup(tmp_path, "app")
    yield path
    lg = logging.getLogger(logs.LOGGER)
    for h in list(lg.handlers):
        lg.removeHandler(h)
        h.close()
    logs._handler = None
    logs.set_project(None)


def test_log_file_location_rotation_and_project_tag(tmp_path, logfile):
    assert logfile == tmp_path / "logs" / "app.log"
    h = logs._handler
    assert h.maxBytes == 5 * 1024 * 1024 and h.backupCount == 5
    logs.set_project("p002", "Test")
    logging.getLogger("uat_tracker.services.sync").warning("something happened")
    h.flush()
    line = logfile.read_text(encoding="utf-8").strip().splitlines()[-1]
    assert "[p002:Test]" in line and "WARNING" in line and "something happened" in line
    jobs = logs.setup(tmp_path, "jobs")                                    # scheduled runs
    assert jobs.name == "jobs.log" and len(logging.getLogger(logs.LOGGER).handlers) == 1


def test_swallowed_errors_are_logged_with_traceback(tmp_path, caplog):
    from uat_tracker.services.projects import validate_workbook
    bad = tmp_path / "broken.xlsx"
    bad.write_bytes(b"not a zip")
    with caplog.at_level(logging.WARNING, logger="uat_tracker"):
        problems = validate_workbook(bad)
    assert problems and problems[0].startswith("Not a readable Excel file")      # behaviour unchanged
    rec = next(r for r in caplog.records if "Not a readable Excel file" in r.getMessage())
    assert rec.exc_info is not None


def test_ai_failure_log_never_contains_the_key(caplog):
    from types import SimpleNamespace
    from uat_tracker.services import weekly_ai
    key = "sk-ant-api03-TESTKEY-never-log-me-123"

    def create(**kw):
        raise RuntimeError(f"boom {key}")
    client = SimpleNamespace(beta=SimpleNamespace(messages=SimpleNamespace(create=create)))
    with caplog.at_level(logging.DEBUG, logger="uat_tracker"):
        r = weekly_ai.summarize({"x": 1}, key=key, client=client)
    assert not r.ok
    text = "\n".join(r.getMessage() + str(r.exc_info or "") for r in caplog.records)
    assert "boom" in text and key not in text


def test_cli_scheduled_run_logs_to_jobs_log(tmp_path, monkeypatch):
    from uat_tracker import cli
    root = tmp_path / "root"
    conf = tmp_path / "config.toml"
    conf.write_text(f'[source]\nexcel_path = "x.xlsx"\n[storage]\ndata_dir = "{root.as_posix()}"\n', encoding="utf-8")
    try:
        cli.main(["--config", str(conf), "run-daily", "--scheduled"])       # no project yet: exits early
        assert logs._handler.baseFilename == str(root / "logs" / "jobs.log")
        cli.main(["--config", str(conf), "project", "list"])
        assert logs._handler.baseFilename == str(root / "logs" / "app.log")
    finally:
        lg = logging.getLogger(logs.LOGGER)
        for h in list(lg.handlers):
            lg.removeHandler(h)
            h.close()
        logs._handler = None


# --------------------------------------------------------------------------- SQLite WAL

def test_project_db_and_app_db_use_wal_and_busy_timeout(tmp_path):
    app = AppDb(tmp_path / "app.db")
    (tmp_path / "p").mkdir()
    repo = SqliteRepository(tmp_path / "p" / "tracker.db", shared=tmp_path / "app.db")
    try:
        for conn in (app.conn, repo.conn):
            assert conn.execute("PRAGMA journal_mode").fetchone()[0] == "wal"
            assert conn.execute("PRAGMA busy_timeout").fetchone()[0] == 10000
        # a plain connection opened later (e.g. the scheduled job) sees WAL too: it is stored in the file
        c = sqlite3.connect(tmp_path / "p" / "tracker.db")
        assert c.execute("PRAGMA journal_mode").fetchone()[0] == "wal"
        c.close()
    finally:
        repo.close()
        app.close()


WRITER = textwrap.dedent("""
    import sys
    sys.path.insert(0, sys.argv[2])
    from uat_tracker.storage.sqlite_repo import SqliteRepository
    repo = SqliteRepository(sys.argv[1])
    for i in range(150):
        with repo.transaction():
            repo.set_setting(f"test.w{i}", str(i))
    repo.close()
    print("writer done")
""")


def test_two_processes_read_and_write_at_the_same_time(cfg, repo):
    """GUI + scheduled job: one process writes in many small transactions while this one keeps reading."""
    src = str(Path(__file__).resolve().parents[1] / "src")
    proc = subprocess.Popen([sys.executable, "-c", WRITER, str(cfg.db_path), src],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    reads = 0
    while proc.poll() is None:
        repo.count_audit()                          # reads never fail with "database is locked"
        with repo.transaction():                    # and short writes from here wait instead of failing
            repo.set_setting("test.reader", str(reads))
        reads += 1
    out, err = proc.communicate(timeout=60)
    assert proc.returncode == 0, err
    assert "writer done" in out and reads > 0
    assert repo.get_setting("test.w149") == "149"


# --------------------------------------------------------------------------- Excel busy

@pytest.fixture
def no_sleep(monkeypatch):
    slept = []
    monkeypatch.setattr(W, "_sleep", slept.append)
    return slept


def test_busy_com_call_retries_with_backoff_then_succeeds(no_sleep):
    calls = []

    def flaky():
        calls.append(1)
        if len(calls) < 4:
            raise FakeComError(BUSY if len(calls) % 2 else RETRY_LATER)
        return "ok"
    assert com_call(flaky) == "ok"
    assert no_sleep == [0.5, 1.0, 2.0] and len(calls) == 4


def test_busy_too_long_raises_workbook_busy_after_about_30s(no_sleep):
    def always_busy():
        raise FakeComError(BUSY)
    with pytest.raises(WorkbookBusy, match="Excel is busy"):
        com_call(always_busy)
    assert sum(no_sleep) <= 30 and sum(no_sleep) >= 20                 # 0.5+1+2+4+8+8 = 23.5 s
    no_sleep.clear()
    with com_retry_budget(5), pytest.raises(WorkbookBusy):           # project switch: shorter
        com_call(always_busy)
    assert sum(no_sleep) <= 5


def test_other_com_errors_are_not_retried(no_sleep):
    def broken():
        raise FakeComError(-2146827284, "Exception occurred.")          # 0x800A03EC: a real error
    with pytest.raises(FakeComError):
        com_call(broken)
    assert no_sleep == []


def test_com_sheet_methods_retry_when_busy(no_sleep):
    class Cell:
        Value = None
    n = {"calls": 0}

    class Ws:
        def Cells(self, r, c):
            n["calls"] += 1
            if n["calls"] == 1:
                raise FakeComError(RETRY_LATER)
            return Cell()
    sh = ComSheet(Ws())
    sh.set_text(2, 3, "hello")
    assert n["calls"] == 2 and no_sleep == [0.5]


class BusyExcel:
    """Opener whose Excel never stops being busy (like a cell being edited in the open workbook)."""

    def __init__(self, path):
        com_call(self._open)

    def _open(self):
        raise FakeComError(BUSY)


def test_busy_excel_goes_to_the_existing_retry_queue(cfg, repo, no_sleep):
    scfg = dataclasses.replace(cfg, alerts=AlertsConfig(pmo_email="pmo@example.test"))
    run_import(scfg, repo, editor="t")
    SyncService(scfg, repo, opener=OpenpyxlWorkbook, sender=FakeSender()).init_ids()
    edit_item(scfg, repo, TP, "TP-38", {"remark": "x"}, actor=PMO_ACTOR)
    res = SyncService(scfg, repo, opener=BusyExcel, sender=FakeSender()).flush(force=True)
    assert res.result == "retry" and "busy" in res.message.lower()
    assert repo.all_pending()                                           # still queued, nothing lost


# --------------------------------------------------------------------------- UI: long actions in a worker

@pytest.fixture
def qapp():
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    pytest.importorskip("PySide6")
    from PySide6.QtWidgets import QApplication
    return QApplication.instance() or QApplication([])


def test_run_in_worker_runs_off_the_ui_thread_and_blocks_clicks(qapp):
    import threading
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QProgressDialog, QWidget
    from uat_tracker.ui.worker import run_in_worker, wait_for_workers
    parent = QWidget()
    seen = {}

    def job():
        seen["thread"] = threading.current_thread() is not threading.main_thread()
        return 42
    run_in_worker(parent, "working", job, lambda r: seen.setdefault("result", r))
    prog = parent.findChild(QProgressDialog)
    assert prog is not None and prog.windowModality() == Qt.ApplicationModal     # no double click
    wait_for_workers()
    assert seen == {"thread": True, "result": 42} and not prog.isVisible()


def test_sync_busy_shows_friendly_message(qapp, cfg, repo, monkeypatch):
    from PySide6.QtWidgets import QWidget
    from uat_tracker.ui import worker
    from uat_tracker.ui.worker import run_in_worker, sync_locked, wait_for_workers
    shown = []
    monkeypatch.setattr(worker, "error", lambda parent, msg: shown.append(msg))
    parent = QWidget()

    def job():
        with sync_locked(cfg):
            return "read"
    with file_lock(cfg.data_dir / "sync.lock"):                         # the sync is writing right now
        run_in_worker(parent, "reading", job)
        wait_for_workers()
    assert shown == ["Sync running, try again in a moment."]
    assert isinstance(SyncBusy("x"), Exception)


# --------------------------------------------------------------------------- E: old Task Scheduler names

def _fake_schtasks(monkeypatch, existing):
    from types import SimpleNamespace
    from uat_tracker.services import scheduler
    calls = []

    def run(args, **kw):
        calls.append(args)
        if args[1] == "/Query":
            return SimpleNamespace(returncode=0 if args[3] in existing else 1, stdout="", stderr="")
        if args[1] == "/Delete":
            existing.discard(args[4])
        return SimpleNamespace(returncode=0, stdout="SUCCESS", stderr="")
    monkeypatch.setattr(scheduler.subprocess, "run", run)
    return calls


def test_install_creates_new_tasks_then_removes_old_named_ones(cfg, repo, monkeypatch):
    from uat_tracker.services.settings import install_schedule
    existing = {"UAT Tracker - Daily Alerts", "UAT Tracker - Weekly Report",
                "Planning Progress - Daily Alerts", "Planning Progress - Weekly Report", "Some other task"}
    calls = _fake_schtasks(monkeypatch, existing)
    daily, weekly = install_schedule(repo, PMO_ACTOR, Path("config.toml"), "08:30", "p002", "Test")
    creates = [c[c.index("/TN") + 1] for c in calls if c[1] == "/Create"]
    deletes = [c[4] for c in calls if c[1] == "/Delete"]
    assert creates == ["Planning Progress - Daily Alerts - Test", "Planning Progress - Weekly Report - Test"]
    assert calls.index(next(c for c in calls if c[1] == "/Delete")) > \
        calls.index(next(c for c in calls if c[1] == "/Create" and "Weekly" in c[c.index("/TN") + 1]))   # new first
    assert set(deletes) == {"UAT Tracker - Daily Alerts", "UAT Tracker - Weekly Report",
                            "Planning Progress - Daily Alerts", "Planning Progress - Weekly Report"}
    assert existing == {"Some other task"} and "Removed old task(s)" in weekly.stdout


def test_old_tasks_kept_when_new_ones_could_not_be_created(cfg, repo, monkeypatch):
    from types import SimpleNamespace
    from uat_tracker.services import scheduler
    from uat_tracker.services.settings import install_schedule
    existing = {"UAT Tracker - Daily Alerts"}
    _fake_schtasks(monkeypatch, existing)
    monkeypatch.setattr(scheduler, "install", lambda *a, **k: SimpleNamespace(returncode=1, stdout="", stderr="no"))
    install_schedule(repo, PMO_ACTOR, Path("config.toml"), "08:30", "p002", "Test")
    assert existing == {"UAT Tracker - Daily Alerts"}
