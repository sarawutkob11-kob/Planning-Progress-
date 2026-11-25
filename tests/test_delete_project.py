"""Delete project (25-Nov-2026) — spec section "2026-11-25 — ลบโปรเจกต์".

T20: migration 002_project_delete.sql + AppDb.delete_project + next_project_id (AC-16, AC-17, AC-19 AppDb part).
"""

from __future__ import annotations

import sqlite3
from importlib import resources
from pathlib import Path

import pytest

from uat_tracker.storage.app_db import AppDb

TRIGGER_MSG = "archive the project before deleting it"


def _sql(name: str) -> str:
    return (resources.files("uat_tracker.storage").joinpath("migrations_app").joinpath(name)
            .read_text(encoding="utf-8"))


def _row(pid: str, name: str, archived: int = 0, **extra) -> dict:
    row = {"id": pid, "name": name, "excel_path": f"C:/x/{name}.xlsx", "created_at": "2026-11-01 09:00",
           "created_by": "Admin", "archived": archived,
           "archived_at": "2026-11-20 10:00" if archived else None,
           "window_start": None, "window_end": None}
    row.update(extra)
    return row


def _make_v1(path: Path) -> list[dict]:
    """An app.db as it was before this change: only 001 applied, user_version = 1, with data."""
    rows = [_row("p001", "Alpha", 0, window_start="2026-11-01", window_end="2026-12-31"),
            _row("p002", "Beta", 1),
            _row("p003", "Gamma ไทย", 0)]
    conn = sqlite3.connect(str(path), isolation_level=None)
    conn.executescript("BEGIN;\n" + _sql("001_app.sql") + "\nPRAGMA user_version = 1;\nCOMMIT;")
    cols = list(rows[0])
    for r in rows:
        conn.execute(f"INSERT INTO project ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                     [r[c] for c in cols])
    conn.execute("INSERT INTO project_event (ts, project_id, editor, action, detail)"
                 " VALUES ('2026-11-01', 'p001', 'Admin', 'created', 'x')")
    conn.execute("INSERT INTO project_event (ts, project_id, editor, action, detail)"
                 " VALUES ('2026-11-20', 'p002', 'Admin', 'archived', '')")
    conn.execute("INSERT INTO settings (key, value) VALUES ('settings.pmo_email', 'pmo@example.com')")
    conn.execute("INSERT INTO holiday (date, name, source, fetched_at) VALUES ('2026-12-05', 'H', 'api', 'now')")
    # v1 really refuses deletes
    with pytest.raises(sqlite3.IntegrityError, match="never deleted"):
        conn.execute("DELETE FROM project WHERE id='p002'")
    conn.close()
    return rows


def _version(db: AppDb) -> int:
    return db.conn.execute("PRAGMA user_version").fetchone()[0]


def _trigger_sql(db: AppDb) -> list[str]:
    return [r[0] for r in db.conn.execute(
        "SELECT sql FROM sqlite_master WHERE type='trigger' AND tbl_name='project'")]


@pytest.fixture
def appdb(tmp_path):
    db = AppDb(tmp_path / "app.db")
    yield db
    db.close()


@pytest.fixture
def filled(appdb):
    with appdb.transaction():
        appdb.insert_project(_row("p001", "Active", 0))
        appdb.insert_project(_row("p002", "Archived", 1))
        appdb.insert_project(_row("p003", "Archived too", 1))
    return appdb


# ---------------------------------------------------------------- AC-16 migration
def test_ac16_v1_db_migrates_to_v2_with_data_intact(tmp_path):
    path = tmp_path / "app.db"
    rows = _make_v1(path)
    db = AppDb(path)
    try:
        assert _version(db) == 2
        got = {p["id"]: p for p in db.projects(include_archived=True)}
        assert set(got) == {"p001", "p002", "p003"}
        for r in rows:
            assert got[r["id"]] == r
        assert [e["action"] for e in db.events("p001")] == ["created"]
        assert [e["action"] for e in db.events("p002")] == ["archived"]
        assert db.get_setting("settings.pmo_email") == "pmo@example.com"
        assert db.conn.execute("SELECT COUNT(*) FROM holiday").fetchone()[0] == 1
        trig = _trigger_sql(db)
        assert len(trig) == 1 and "archived = 0" in trig[0] and TRIGGER_MSG in trig[0]
    finally:
        db.close()


def test_ac16_fresh_db_is_v2(appdb):
    assert _version(appdb) == 2
    assert len(_trigger_sql(appdb)) == 1


def test_ac16_reopen_v2_does_not_rerun_migration(tmp_path):
    path = tmp_path / "app.db"
    _make_v1(path)
    AppDb(path).close()
    # Re-running 002 would fail-or-duplicate; reopening must be a no-op and keep data.
    for _ in range(2):
        db = AppDb(path)
        try:
            assert _version(db) == 2
            assert len(db.projects(include_archived=True)) == 3
            assert len(_trigger_sql(db)) == 1
        finally:
            db.close()


def test_ac16_migration_runs_once_even_if_002_would_fail_on_rerun(tmp_path):
    """002 does DROP TRIGGER without IF EXISTS: proves it is not executed again on a v2 db."""
    path = tmp_path / "app.db"
    AppDb(path).close()
    conn = sqlite3.connect(str(path), isolation_level=None)
    conn.execute("DROP TRIGGER project_no_delete")          # a re-run of 002 would now raise
    conn.close()
    db = AppDb(path)                                        # must not raise
    try:
        assert _version(db) == 2
    finally:
        db.close()


def test_ac16_001_file_unchanged():
    sql = _sql("001_app.sql")
    assert "projects are archived, never deleted" in sql     # 001 left as it was (D6)


# ---------------------------------------------------------------- AC-17 trigger
def test_ac17_raw_delete_of_active_row_aborts(filled):
    with pytest.raises(sqlite3.IntegrityError, match=TRIGGER_MSG):
        filled.conn.execute("DELETE FROM project WHERE id=?", ("p001",))
    assert filled.get_project("p001") is not None


def test_ac17_raw_delete_of_archived_row_works(filled):
    filled.conn.execute("DELETE FROM project WHERE id=?", ("p002",))
    assert filled.get_project("p002") is None
    assert {p["id"] for p in filled.projects(True)} == {"p001", "p003"}


def test_ac17_appdb_delete_active_aborts_in_transaction(filled):
    with pytest.raises(sqlite3.IntegrityError, match=TRIGGER_MSG):
        with filled.transaction():
            filled.delete_project("p001")
    assert filled.get_project("p001") is not None
    # connection still usable (rolled back, no open transaction)
    assert not filled.conn.in_transaction
    with filled.transaction():
        filled.log_event(ts="t", project_id="p001", editor="x", action="noop")


def test_ac17_appdb_delete_archived_row(filled):
    with filled.transaction():
        filled.delete_project("p002")
        filled.log_event(ts="t", project_id="p002", editor="Admin", action="deleted", detail="d")
    assert filled.get_project("p002") is None
    assert "p002" not in {p["id"] for p in filled.projects(include_archived=True)}
    assert filled.events("p002")[0]["action"] == "deleted"


def test_ac17_delete_rolls_back_with_failed_transaction(filled):
    """Archived row deleted, then something else fails in the same transaction -> row is back."""
    with pytest.raises(RuntimeError):
        with filled.transaction():
            filled.delete_project("p002")
            raise RuntimeError("boom")
    assert filled.get_project("p002") is not None


def test_ac17_delete_unknown_id_is_noop(filled):
    with filled.transaction():
        filled.delete_project("p999")
    assert len(filled.projects(True)) == 3


def test_ac17_whole_table_delete_aborts_entirely(filled):
    with pytest.raises(sqlite3.IntegrityError, match=TRIGGER_MSG):
        with filled.transaction():
            filled.conn.execute("DELETE FROM project")
    assert {p["id"] for p in filled.projects(True)} == {"p001", "p002", "p003"}


def test_ac17_whole_table_delete_aborts_without_explicit_transaction(filled):
    with pytest.raises(sqlite3.IntegrityError, match=TRIGGER_MSG):
        filled.conn.execute("DELETE FROM project")           # autocommit statement: atomic
    assert {p["id"] for p in filled.projects(True)} == {"p001", "p002", "p003"}


def test_ac17_unarchived_again_cannot_be_deleted(filled):
    filled.update_project("p002", archived=0, archived_at=None)
    with pytest.raises(sqlite3.IntegrityError, match=TRIGGER_MSG):
        filled.delete_project("p002")
    assert filled.get_project("p002") is not None


def test_ac17_on_migrated_v1_db(tmp_path):
    path = tmp_path / "app.db"
    _make_v1(path)
    db = AppDb(path)
    try:
        with pytest.raises(sqlite3.IntegrityError, match=TRIGGER_MSG):
            db.conn.execute("DELETE FROM project WHERE id='p001'")
        db.conn.execute("DELETE FROM project WHERE id='p002'")
        assert db.get_project("p002") is None
        assert db.get_project("p001") is not None
        assert db.events("p002")                             # old events kept (no FK)
    finally:
        db.close()


# ---------------------------------------------------------------- AC-19 next_project_id
def test_ac19_empty_db_starts_at_p001(appdb):
    assert appdb.next_project_id() == "p001"


def test_ac19_counts_project_table(filled):
    assert filled.next_project_id() == "p004"


def test_ac19_deleted_highest_id_is_not_reused(appdb):
    with appdb.transaction():
        appdb.insert_project(_row("p001", "A"))
        appdb.insert_project(_row("p002", "B", 1))
        appdb.log_event(ts="t", project_id="p001", editor="x", action="created")
        appdb.log_event(ts="t", project_id="p002", editor="x", action="created")
    assert appdb.next_project_id() == "p003"
    with appdb.transaction():
        appdb.delete_project("p002")
        appdb.log_event(ts="t", project_id="p002", editor="x", action="deleted")
    assert appdb.get_project("p002") is None
    new = appdb.next_project_id()
    assert new == "p003"
    assert appdb.events(new) == []


def test_ac19_id_only_in_events_counts(appdb):
    with appdb.transaction():
        appdb.insert_project(_row("p001", "A"))
        appdb.log_event(ts="t", project_id="p007", editor="x", action="deleted")
    assert appdb.next_project_id() == "p008"


def test_ac19_event_id_higher_than_table_and_vice_versa(appdb):
    with appdb.transaction():
        appdb.insert_project(_row("p010", "A"))
        appdb.log_event(ts="t", project_id="p004", editor="x", action="created")
    assert appdb.next_project_id() == "p011"


def test_ac19_null_and_odd_event_ids_do_not_break(appdb):
    with appdb.transaction():
        appdb.insert_project(_row("p002", "A"))
        for pid in (None, "", "p", "pabc", "abc", "p-1", "p 5"):
            appdb.log_event(ts="t", project_id=pid, editor="x", action="setting")
    assert appdb.next_project_id() == "p003"


def test_ac19_only_null_events_and_no_projects(appdb):
    with appdb.transaction():
        appdb.log_event(ts="t", project_id=None, editor="x", action="setting")
    assert appdb.next_project_id() == "p001"


def test_ac19_beyond_three_digits(appdb):
    with appdb.transaction():
        appdb.log_event(ts="t", project_id="p999", editor="x", action="deleted")
    assert appdb.next_project_id() == "p1000"
    with appdb.transaction():
        appdb.insert_project(_row("p1000", "Big"))
    assert appdb.next_project_id() == "p1001"
