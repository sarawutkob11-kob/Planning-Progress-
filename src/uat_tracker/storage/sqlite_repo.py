"""SQLite implementation of Repository (phase 1: local file on the PMO PC)."""

from __future__ import annotations

import hashlib
import json
import sqlite3
from contextlib import contextmanager
from importlib import resources
from pathlib import Path

from ..domain.models import DATE_FIELDS, HISTORY_SUFFIX
from ..excel.sheets import DEFECT_LIST, TEST_PROGRESS, UAT_PLAN

TABLES = {UAT_PLAN: "uat_task", TEST_PROGRESS: "test_progress", DEFECT_LIST: "defect"}
META_COLUMNS = ("id", "excel_row", "fingerprint", "alt_key", "active", "created_at", "updated_at")
AUDIT_HASH_FIELDS = ("ts", "editor", "channel", "sheet", "item_id", "field",
                     "old_value", "new_value", "reason", "import_run_id")


class SqliteRepository:
    def __init__(self, path: Path | str, shared: Path | str | None = None):
        """shared: app.db (multi-project). When given, holidays and SHARED_SETTING_KEYS are read from and
        written to it; everything else stays in this project's own DB."""
        self.path = str(path)
        self.conn = sqlite3.connect(self.path, isolation_level=None)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self._migrate()
        self.shared = str(shared) if shared else None
        self._s = ""
        if self.shared:
            from .app_db import AppDb
            AppDb(self.shared).close()                   # creates / migrates app.db
            self.conn.execute("ATTACH DATABASE ? AS shared", (self.shared,))
            self._s = "shared."
        self._columns = {
            sheet: [r["name"] for r in self.conn.execute(f"PRAGMA table_info({t})")
                    if r["name"] not in META_COLUMNS]
            for sheet, t in TABLES.items()
        }

    # ------------------------------------------------------------------ infra
    def _migrate(self) -> None:
        version = self.conn.execute("PRAGMA user_version").fetchone()[0]
        files = sorted(
            (f for f in resources.files("uat_tracker.storage").joinpath("migrations").iterdir()
             if f.name.endswith(".sql")),
            key=lambda f: f.name,
        )
        for f in files:
            n = int(f.name.split("_", 1)[0])
            if n > version:
                self.conn.executescript("BEGIN;\n" + f.read_text(encoding="utf-8")
                                        + f"\nPRAGMA user_version = {n};\nCOMMIT;")

    @contextmanager
    def transaction(self):
        self.conn.execute("BEGIN IMMEDIATE")
        try:
            yield self
        except BaseException:
            self.conn.execute("ROLLBACK")
            raise
        else:
            self.conn.execute("COMMIT")

    @property
    def in_transaction(self) -> bool:
        return self.conn.in_transaction

    def close(self) -> None:
        self.conn.close()

    # ------------------------------------------------------------------ import runs
    def start_run(self, *, started_at, editor, source_path, source_mtime) -> int:
        cur = self.conn.execute(
            "INSERT INTO import_run (started_at, editor, source_path, source_mtime, result)"
            " VALUES (?, ?, ?, ?, 'running')", (started_at, editor, source_path, source_mtime))
        return cur.lastrowid

    def finish_run(self, run_id, *, finished_at, result, message="", source_sha256=None,
                   snapshot_path=None, stats_json=None) -> None:
        self.conn.execute(
            "UPDATE import_run SET finished_at=?, result=?, message=?, source_sha256=?,"
            " snapshot_path=?, stats_json=? WHERE id=?",
            (finished_at, result, message, source_sha256, snapshot_path, stats_json, run_id))

    def last_run(self) -> dict | None:
        r = self.conn.execute("SELECT * FROM import_run ORDER BY id DESC LIMIT 1").fetchone()
        return dict(r) if r else None

    # ------------------------------------------------------------------ items
    def data_columns(self, sheet: str) -> list[str]:
        return list(self._columns[sheet])

    def load_items(self, sheet: str) -> dict[str, dict]:
        table = TABLES[sheet]
        items = {r["id"]: dict(r) for r in self.conn.execute(f"SELECT * FROM {table}")}
        for rec in items.values():
            for f in DATE_FIELDS:
                if f in rec:
                    rec[f + HISTORY_SUFFIX] = ()
        for r in self.conn.execute(
                "SELECT item_id, field, value FROM date_history WHERE sheet=? AND source='import'"
                " ORDER BY item_id, field, seq", (sheet,)):
            rec = items.get(r["item_id"])
            if rec is not None:
                key = r["field"] + HISTORY_SUFFIX
                rec[key] = rec.get(key, ()) + (r["value"],)
        return items

    def _split(self, sheet: str, fields: dict) -> tuple[dict, dict]:
        cols = self._columns[sheet]
        data = {c: fields.get(c) for c in cols}
        hist = {k[: -len(HISTORY_SUFFIX)]: v for k, v in fields.items() if k.endswith(HISTORY_SUFFIX)}
        return data, hist

    def _write_history(self, sheet, item_id, hist: dict, now: str, only_changed_from: dict | None):
        for field, values in hist.items():
            if only_changed_from is not None and tuple(only_changed_from.get(field + HISTORY_SUFFIX, ())) == tuple(values):
                continue
            self.conn.execute(
                "DELETE FROM date_history WHERE sheet=? AND item_id=? AND field=? AND source='import'",
                (sheet, item_id, field))
            for seq, value in enumerate(values, start=1):
                self.conn.execute(
                    "INSERT INTO date_history (sheet, item_id, field, seq, value, source, recorded_at)"
                    " VALUES (?, ?, ?, ?, ?, 'import', ?)", (sheet, item_id, field, seq, value, now))

    def insert_item(self, sheet, item_id, fields, meta) -> None:
        data, hist = self._split(sheet, fields)
        cols = ["id", "excel_row", "fingerprint", "alt_key", "active", "created_at", "updated_at", *data]
        vals = [item_id, meta["excel_row"], meta["fingerprint"], meta["alt_key"], 1,
                meta["now"], meta["now"], *data.values()]
        self.conn.execute(
            f"INSERT INTO {TABLES[sheet]} ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})", vals)
        self._write_history(sheet, item_id, hist, meta["now"], None)

    def update_item(self, sheet, item_id, fields, meta) -> None:
        """meta: excel_row, fingerprint, alt_key, now, data_changed (bool), previous (old record)."""
        data, hist = self._split(sheet, fields)
        sets = {"excel_row": meta["excel_row"], "fingerprint": meta["fingerprint"],
                "alt_key": meta["alt_key"], **data}
        if meta.get("data_changed"):
            sets["updated_at"] = meta["now"]
        self.conn.execute(
            f"UPDATE {TABLES[sheet]} SET {', '.join(f'{c}=?' for c in sets)} WHERE id=?",
            [*sets.values(), item_id])
        self._write_history(sheet, item_id, hist, meta["now"], meta.get("previous"))

    def set_active(self, sheet, item_id, active, now) -> None:
        self.conn.execute(f"UPDATE {TABLES[sheet]} SET active=?, updated_at=? WHERE id=?",
                          (1 if active else 0, now, item_id))

    def count_items(self) -> dict[str, dict[str, int]]:
        out = {}
        for sheet, t in TABLES.items():
            r = self.conn.execute(
                f"SELECT SUM(active=1) AS active, SUM(active=0) AS inactive FROM {t}").fetchone()
            out[sheet] = {"active": r["active"] or 0, "inactive": r["inactive"] or 0}
        return out

    def date_history(self, sheet: str, item_id: str, field: str) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM date_history WHERE sheet=? AND item_id=? AND field=? ORDER BY seq",
            (sheet, item_id, field))]

    def get_item(self, sheet: str, item_id: str) -> dict | None:
        return self.load_items(sheet).get(item_id)

    # ------------------------------------------------------------------ baseline
    def load_baseline(self, sheet) -> dict[tuple[str, str], str]:
        return {(r["item_id"], r["field"]): r["raw_value"] for r in self.conn.execute(
            "SELECT item_id, field, raw_value FROM cell_baseline WHERE sheet=?", (sheet,))}

    def upsert_baseline(self, sheet, item_id, field, raw, now) -> None:
        self.conn.execute(
            "INSERT INTO cell_baseline (sheet, item_id, field, raw_value, read_at) VALUES (?,?,?,?,?)"
            " ON CONFLICT(sheet, item_id, field) DO UPDATE SET raw_value=excluded.raw_value,"
            " read_at=excluded.read_at", (sheet, item_id, field, raw, now))

    # ------------------------------------------------------------------ audit
    def append_audit(self, entry: dict) -> int:
        e = {k: entry.get(k) for k in AUDIT_HASH_FIELDS}
        r = self.conn.execute("SELECT hash FROM audit_log ORDER BY id DESC LIMIT 1").fetchone()
        prev = r["hash"] if r else ""
        digest = hashlib.sha256(
            (prev + json.dumps(e, ensure_ascii=False, sort_keys=True, default=str)).encode("utf-8")
        ).hexdigest()
        cols = [*AUDIT_HASH_FIELDS, "prev_hash", "hash"]
        cur = self.conn.execute(
            f"INSERT INTO audit_log ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
            [*e.values(), prev, digest])
        return cur.lastrowid

    def audit_entries(self, item_id=None) -> list[dict]:
        if item_id is None:
            rows = self.conn.execute("SELECT * FROM audit_log ORDER BY id")
        else:
            rows = self.conn.execute("SELECT * FROM audit_log WHERE item_id=? ORDER BY id", (item_id,))
        return [dict(r) for r in rows]

    def count_audit(self) -> int:
        return self.conn.execute("SELECT COUNT(*) FROM audit_log").fetchone()[0]

    def verify_audit_chain(self) -> bool:
        prev = ""
        for r in self.conn.execute("SELECT * FROM audit_log ORDER BY id"):
            e = {k: r[k] for k in AUDIT_HASH_FIELDS}
            digest = hashlib.sha256(
                (prev + json.dumps(e, ensure_ascii=False, sort_keys=True, default=str)).encode("utf-8")
            ).hexdigest()
            if r["prev_hash"] != prev or r["hash"] != digest:
                return False
            prev = r["hash"]
        return True

    # ------------------------------------------------------------------ review queue
    def load_reviews(self, sheet) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM review_queue WHERE sheet=?", (sheet,))]

    def insert_review(self, review: dict) -> int:
        cols = ["created_at", "sheet", "item_id", "excel_row", "field", "raw_value", "token",
                "problem", "import_run_id"]
        cur = self.conn.execute(
            f"INSERT INTO review_queue ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
            [review.get(c) for c in cols])
        return cur.lastrowid

    def set_review_status(self, review_id, status) -> None:
        self.conn.execute("UPDATE review_queue SET status=? WHERE id=?", (status, review_id))

    def open_reviews(self) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM review_queue WHERE status='open' ORDER BY sheet, excel_row, field")]

    # ------------------------------------------------------------------ flags
    def load_flags(self, sheet) -> dict[tuple[str, str], dict]:
        return {(r["item_id"], r["code"]): dict(r) for r in self.conn.execute(
            "SELECT * FROM data_flag WHERE sheet=?", (sheet,))}

    def upsert_flag(self, sheet, item_id, code, detail, now) -> bool:
        """Returns True when the flag is new or re-activated."""
        r = self.conn.execute("SELECT id, active, detail FROM data_flag WHERE sheet=? AND item_id=? AND code=?",
                              (sheet, item_id, code)).fetchone()
        if r is None:
            self.conn.execute(
                "INSERT INTO data_flag (sheet, item_id, code, detail, active, first_seen, last_seen)"
                " VALUES (?,?,?,?,1,?,?)", (sheet, item_id, code, detail, now, now))
            return True
        if r["active"] and r["detail"] == detail:
            return False
        self.conn.execute("UPDATE data_flag SET active=1, detail=?, last_seen=? WHERE id=?",
                          (detail, now, r["id"]))
        return not r["active"]

    def deactivate_flag(self, flag_id, now) -> None:
        self.conn.execute("UPDATE data_flag SET active=0, last_seen=? WHERE id=?", (now, flag_id))

    def active_flags(self) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM data_flag WHERE active=1 ORDER BY sheet, item_id, code")]

    # ------------------------------------------------------------------ Part 2: single fields + settings
    def set_field(self, sheet: str, item_id: str, field: str, value, now: str) -> None:
        if field not in self._columns[sheet]:
            raise KeyError(f"{sheet} has no field {field}")
        self.conn.execute(f"UPDATE {TABLES[sheet]} SET {field}=?, updated_at=? WHERE id=?",
                          (value, now, item_id))

    def _st(self, key: str) -> str:
        from ..config import SHARED_SETTING_KEYS
        return f"{self._s}settings" if key in SHARED_SETTING_KEYS else "settings"

    def get_setting(self, key: str) -> str | None:
        r = self.conn.execute(f"SELECT value FROM {self._st(key)} WHERE key=?", (key,)).fetchone()
        return r["value"] if r else None

    def set_setting(self, key: str, value: str | None) -> None:
        self.conn.execute(f"INSERT INTO {self._st(key)} (key, value) VALUES (?, ?) ON CONFLICT(key)"
                          " DO UPDATE SET value=excluded.value", (key, value))

    # ------------------------------------------------------------------ holidays
    def holidays(self, year: int | None = None) -> dict[str, dict]:
        if year is None:
            rows = self.conn.execute(f"SELECT * FROM {self._s}holiday ORDER BY date")
        else:
            rows = self.conn.execute(f"SELECT * FROM {self._s}holiday WHERE date LIKE ? ORDER BY date",
                                     (f"{year}-%",))
        return {r["date"]: dict(r) for r in rows}

    def holiday_years(self) -> set[int]:
        return {int(r[0]) for r in self.conn.execute(f"SELECT DISTINCT substr(date, 1, 4) FROM {self._s}holiday")}

    def upsert_holiday(self, date_iso: str, name: str, source: str, now: str) -> bool:
        """Returns True when the holiday is new or its name/source changed."""
        r = self.conn.execute(f"SELECT name, source FROM {self._s}holiday WHERE date=?", (date_iso,)).fetchone()
        if r is not None and (r["name"], r["source"]) == (name, source):
            return False
        self.conn.execute(
            f"INSERT INTO {self._s}holiday (date, name, source, fetched_at) VALUES (?,?,?,?) ON CONFLICT(date)"
            f" DO UPDATE SET name=excluded.name, source=excluded.source, fetched_at=excluded.fetched_at",
            (date_iso, name, source, now))
        return True

    def delete_holiday(self, date_iso: str) -> dict | None:
        r = self.conn.execute(f"SELECT * FROM {self._s}holiday WHERE date=?", (date_iso,)).fetchone()
        if r is None:
            return None
        self.conn.execute(f"DELETE FROM {self._s}holiday WHERE date=?", (date_iso,))
        return dict(r)

    def log_holiday_fetch(self, *, ts, year, result, count, message) -> None:
        self.conn.execute(f"INSERT INTO {self._s}holiday_fetch (ts, year, result, count, message) VALUES (?,?,?,?,?)",
                          (ts, year, result, count, message))

    def holiday_fetches(self, limit: int = 20) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            f"SELECT * FROM {self._s}holiday_fetch ORDER BY id DESC LIMIT ?", (limit,))]

    # ------------------------------------------------------------------ job runs
    def start_job(self, *, job, run_date, started_at, mode) -> int:
        return self.conn.execute(
            "INSERT INTO job_run (job, run_date, started_at, result, mode) VALUES (?,?,?,'running',?)",
            (job, run_date, started_at, mode)).lastrowid

    def finish_job(self, job_id, *, finished_at, result, message="") -> None:
        self.conn.execute("UPDATE job_run SET finished_at=?, result=?, message=? WHERE id=?",
                          (finished_at, result, message, job_id))

    def record_job(self, *, job, run_date, ts, result, mode, message) -> int:
        return self.conn.execute(
            "INSERT INTO job_run (job, run_date, started_at, finished_at, result, mode, message)"
            " VALUES (?,?,?,?,?,?,?)", (job, run_date, ts, ts, result, mode, message)).lastrowid

    def job_runs(self, job=None, limit=30) -> list[dict]:
        if job is None:
            rows = self.conn.execute("SELECT * FROM job_run ORDER BY id DESC LIMIT ?", (limit,))
        else:
            rows = self.conn.execute("SELECT * FROM job_run WHERE job=? ORDER BY id DESC LIMIT ?",
                                     (job, limit))
        return [dict(r) for r in rows]

    def job_dates(self, job) -> set[str]:
        return {r[0] for r in self.conn.execute("SELECT DISTINCT run_date FROM job_run WHERE job=?", (job,))}

    def job_ok_on(self, job, run_date) -> bool:
        return self.conn.execute("SELECT 1 FROM job_run WHERE job=? AND run_date=? AND result='ok'",
                                 (job, run_date)).fetchone() is not None

    # ------------------------------------------------------------------ traffic light history
    def light_history(self, item_id) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM traffic_light_history WHERE item_id=? ORDER BY id", (item_id,))]

    def latest_lights(self) -> dict[str, dict]:
        return {r["item_id"]: dict(r) for r in self.conn.execute(
            "SELECT h.* FROM traffic_light_history h JOIN (SELECT item_id, MAX(id) AS id"
            " FROM traffic_light_history GROUP BY item_id) m ON h.id = m.id")}

    def insert_light(self, row: dict) -> int:
        cols = ["item_id", "color", "prev_color", "pct_exec", "pct_time", "pct_fail", "reason",
                "calc_date", "ts", "job_run_id"]
        return self.conn.execute(
            f"INSERT INTO traffic_light_history ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
            [row.get(c) for c in cols]).lastrowid

    # ------------------------------------------------------------------ email log
    def insert_email(self, row: dict) -> int:
        cols = ["ts", "send_date", "rule", "variant", "sheet", "item_id", "app", "to_addr", "cc",
                "delivered_to", "subject", "mode", "result", "error", "event_ref", "preview_path",
                "job_run_id"]
        return self.conn.execute(
            f"INSERT INTO email_log ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
            [row.get(c) for c in cols]).lastrowid

    def email_done(self, rule, item_id, send_date, mode) -> bool:
        """Dedupe: a sent / dry-run email already exists for this rule + item + day + mode."""
        return self.conn.execute(
            "SELECT 1 FROM email_log WHERE rule=? AND item_id=? AND send_date=? AND mode=?"
            " AND result IN ('sent','dry_run')", (rule, item_id, send_date, mode)).fetchone() is not None

    def event_notified(self, rule, event_ref, mode) -> bool:
        return self.conn.execute(
            "SELECT 1 FROM email_log WHERE rule=? AND event_ref=? AND mode=? AND result IN ('sent','dry_run')",
            (rule, event_ref, mode)).fetchone() is not None

    def email_log(self, *, send_date=None, item_id=None, limit=200) -> list[dict]:
        q, args = "SELECT * FROM email_log WHERE 1=1", []
        if send_date:
            q += " AND send_date=?"
            args.append(send_date)
        if item_id:
            q += " AND item_id=?"
            args.append(item_id)
        q += " ORDER BY id DESC LIMIT ?"
        args.append(limit)
        return [dict(r) for r in self.conn.execute(q, args)]

    # ------------------------------------------------------------------ attachments
    def insert_attachment(self, row: dict) -> int:
        cols = ["sheet", "item_id", "original_name", "stored_path", "sha256", "size_bytes",
                "uploaded_by", "uploaded_at"]
        return self.conn.execute(
            f"INSERT INTO attachment ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
            [row[c] for c in cols]).lastrowid

    def attachments(self, sheet, item_id=None) -> list[dict]:
        if item_id is None:
            rows = self.conn.execute("SELECT * FROM attachment WHERE sheet=? ORDER BY id", (sheet,))
        else:
            rows = self.conn.execute("SELECT * FROM attachment WHERE sheet=? AND item_id=? ORDER BY id",
                                     (sheet, item_id))
        return [dict(r) for r in rows]

    # ------------------------------------------------------------------ pending write-back (Part 3 drains it)
    def set_pending(self, sheet, item_id, field, value, now, reason) -> None:
        self.conn.execute(
            "INSERT INTO pending_writeback (sheet, item_id, field, value, created_at, reason)"
            " VALUES (?,?,?,?,?,?) ON CONFLICT(sheet, item_id, field) DO UPDATE SET"
            " value=excluded.value, created_at=excluded.created_at, reason=excluded.reason",
            (sheet, item_id, field, value, now, reason))

    def pending(self, sheet) -> dict[tuple[str, str], dict]:
        return {(r["item_id"], r["field"]): dict(r) for r in self.conn.execute(
            "SELECT * FROM pending_writeback WHERE sheet=?", (sheet,))}

    def clear_pending(self, sheet, item_id, field) -> None:
        self.conn.execute("DELETE FROM pending_writeback WHERE sheet=? AND item_id=? AND field=?",
                          (sheet, item_id, field))

    # ------------------------------------------------------------------ Part 3: write queue
    def upsert_pending(self, sheet, item_id, field, *, value, excel_value, now, reason, editor, channel) -> None:
        """Queue (or replace) an app value for Excel. Resets retry state."""
        self.conn.execute(
            "INSERT INTO pending_writeback (sheet, item_id, field, value, created_at, reason, excel_value,"
            " editor, channel, status, attempts) VALUES (?,?,?,?,?,?,?,?,?,'pending',0)"
            " ON CONFLICT(sheet, item_id, field) DO UPDATE SET value=excluded.value,"
            " created_at=excluded.created_at, reason=excluded.reason, excel_value=excluded.excel_value,"
            " editor=excluded.editor, channel=excluded.channel, status='pending', attempts=0,"
            " first_failed_at=NULL, next_attempt_at=NULL, last_error=NULL",
            (sheet, item_id, field, value, now, reason, excel_value, editor, channel))

    def all_pending(self) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM pending_writeback ORDER BY created_at, sheet, item_id, field")]

    def set_pending_state(self, sheet, item_id, field, *, status, attempts=None, first_failed_at=None,
                          next_attempt_at=None, last_error=None) -> None:
        self.conn.execute(
            "UPDATE pending_writeback SET status=?, attempts=COALESCE(?, attempts),"
            " first_failed_at=?, next_attempt_at=?, last_error=? WHERE sheet=? AND item_id=? AND field=?",
            (status, attempts, first_failed_at, next_attempt_at, last_error, sheet, item_id, field))

    # ------------------------------------------------------------------ conflicts
    def insert_conflict(self, row: dict) -> int:
        cols = ["created_at", "sheet", "item_id", "field", "kind", "excel_raw", "excel_value",
                "excel_time", "app_value", "app_time", "app_editor"]
        return self.conn.execute(
            f"INSERT INTO conflict ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
            [row.get(c) for c in cols]).lastrowid

    def open_conflicts(self, sheet: str | None = None) -> list[dict]:
        if sheet is None:
            rows = self.conn.execute("SELECT * FROM conflict WHERE status='open' ORDER BY id")
        else:
            rows = self.conn.execute("SELECT * FROM conflict WHERE status='open' AND sheet=? ORDER BY id",
                                     (sheet,))
        return [dict(r) for r in rows]

    def get_conflict(self, conflict_id: int) -> dict | None:
        r = self.conn.execute("SELECT * FROM conflict WHERE id=?", (conflict_id,)).fetchone()
        return dict(r) if r else None

    def resolve_conflict(self, conflict_id, *, resolution, resolved_value, reason, resolved_by, resolved_at) -> None:
        self.conn.execute(
            "UPDATE conflict SET status='resolved', resolution=?, resolved_value=?, reason=?,"
            " resolved_by=?, resolved_at=? WHERE id=?",
            (resolution, resolved_value, reason, resolved_by, resolved_at, conflict_id))

    # ------------------------------------------------------------------ review queue (resolve)
    def get_review(self, review_id: int) -> dict | None:
        r = self.conn.execute("SELECT * FROM review_queue WHERE id=?", (review_id,)).fetchone()
        return dict(r) if r else None

    def resolve_review(self, review_id, *, value, by, at) -> None:
        self.conn.execute(
            "UPDATE review_queue SET status='resolved', resolved_value=?, resolved_by=?, resolved_at=?"
            " WHERE id=?", (value, by, at, review_id))

    # ------------------------------------------------------------------ backups + sync events
    def insert_backup(self, row: dict) -> int:
        return self.conn.execute(
            "INSERT INTO backup (ts, path, sha256, size_bytes, reason) VALUES (?,?,?,?,?)",
            (row["ts"], row["path"], row["sha256"], row["size_bytes"], row.get("reason"))).lastrowid

    def backups(self, limit: int = 50) -> list[dict]:
        return [dict(r) for r in self.conn.execute("SELECT * FROM backup ORDER BY id DESC LIMIT ?", (limit,))]

    def get_backup(self, backup_id: int) -> dict | None:
        r = self.conn.execute("SELECT * FROM backup WHERE id=?", (backup_id,)).fetchone()
        return dict(r) if r else None

    def log_sync(self, *, ts, kind, result, message="") -> None:
        self.conn.execute("INSERT INTO sync_event (ts, kind, result, message) VALUES (?,?,?,?)",
                          (ts, kind, result, message))

    def sync_events(self, limit: int = 30) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM sync_event ORDER BY id DESC LIMIT ?", (limit,))]

    def last_ok_import(self) -> dict | None:
        r = self.conn.execute(
            "SELECT * FROM import_run WHERE result='ok' ORDER BY id DESC LIMIT 1").fetchone()
        return dict(r) if r else None

    def insert_app_history(self, sheet, item_id, field, value, *, reason, root_cause, now) -> None:
        """Reschedule made in the app: the previous date, with reason + root cause."""
        seq = self.conn.execute("SELECT COALESCE(MAX(seq), 0) + 1 FROM date_history"
                                " WHERE sheet=? AND item_id=? AND field=?", (sheet, item_id, field)).fetchone()[0]
        self.conn.execute(
            "INSERT INTO date_history (sheet, item_id, field, seq, value, source, reason, root_cause,"
            " recorded_at) VALUES (?,?,?,?,?,'app',?,?,?)",
            (sheet, item_id, field, seq, value, reason, root_cause, now))

    def set_import_history(self, sheet, item_id, field, values, now) -> None:
        """History as the Excel cell will show it after an app write (so reading it back is not a change)."""
        self._write_history(sheet, item_id, {field: tuple(values)}, now, None)

    # ------------------------------------------------------------------ Part 4: contacts, queries for screens
    def get_contact(self, name_key: str) -> dict | None:
        r = self.conn.execute("SELECT * FROM contact WHERE name_key=?", (name_key,)).fetchone()
        return dict(r) if r else None

    def set_contact(self, name_key: str, name: str, email: str, now: str) -> None:
        self.conn.execute(
            "INSERT INTO contact (name_key, name, email, updated_at) VALUES (?,?,?,?) ON CONFLICT(name_key)"
            " DO UPDATE SET name=excluded.name, email=excluded.email, updated_at=excluded.updated_at",
            (name_key, name, email, now))

    def all_date_history(self, sheet: str) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM date_history WHERE sheet=? ORDER BY item_id, field, seq", (sheet,))]

    def max_audit_id(self) -> int:
        return self.conn.execute("SELECT COALESCE(MAX(id), 0) FROM audit_log").fetchone()[0]

    def search_audit(self, *, text=None, sheet=None, editor=None, date_from=None, date_to=None,
                     item_id=None, limit=2000) -> list[dict]:
        q, a = "SELECT * FROM audit_log WHERE 1=1", []
        if text:
            q += " AND (item_id LIKE ? OR field LIKE ? OR old_value LIKE ? OR new_value LIKE ? OR reason LIKE ?)"
            a += [f"%{text}%"] * 5
        if sheet:
            q += " AND sheet=?"
            a.append(sheet)
        if editor:
            q += " AND editor=?"
            a.append(editor)
        if date_from:
            q += " AND substr(ts, 1, 10) >= ?"
            a.append(date_from)
        if date_to:
            q += " AND substr(ts, 1, 10) <= ?"
            a.append(date_to)
        if item_id:
            q += " AND item_id=?"
            a.append(item_id)
        q += " ORDER BY id DESC LIMIT ?"
        a.append(limit)
        return [dict(r) for r in self.conn.execute(q, a)]

    def audit_editors(self) -> list[str]:
        return [r[0] for r in self.conn.execute("SELECT DISTINCT editor FROM audit_log ORDER BY editor")]

    def get_audit(self, audit_id: int) -> dict | None:
        r = self.conn.execute("SELECT * FROM audit_log WHERE id=?", (audit_id,)).fetchone()
        return dict(r) if r else None

    def insert_new_item(self, sheet: str, item_id: str, fields: dict, now: str) -> None:
        """Item created in the app: no Excel row yet (the sync writes it)."""
        self.insert_item(sheet, item_id, fields, {"excel_row": None, "fingerprint": None, "alt_key": None,
                                                  "now": now})

    # ------------------------------------------------------------------ Part 5: users, codes, sessions
    def count_users(self, active_only: bool = False) -> int:
        q = "SELECT COUNT(*) FROM app_user" + (" WHERE active=1" if active_only else "")
        return self.conn.execute(q).fetchone()[0]

    def get_user(self, email: str) -> dict | None:
        r = self.conn.execute("SELECT * FROM app_user WHERE email=?", (email,)).fetchone()
        return dict(r) if r else None

    def list_users(self) -> list[dict]:
        return [dict(r) for r in self.conn.execute("SELECT * FROM app_user ORDER BY active DESC, email")]

    def insert_user(self, *, email, first_name, last_name, role, created_at, created_by) -> None:
        self.conn.execute(
            "INSERT INTO app_user (email, first_name, last_name, display_name, role, active, created_at, created_by)"
            " VALUES (?,?,?,?,?,1,?,?)",
            (email, first_name, last_name, f"{first_name} {last_name}".strip(), role, created_at, created_by))

    def update_user(self, email: str, **fields) -> None:
        allowed = {"first_name", "last_name", "display_name", "role", "active", "last_used_at"}
        bad = set(fields) - allowed
        if bad:
            raise KeyError(f"app_user has no field(s) {bad}")
        self.conn.execute(f"UPDATE app_user SET {', '.join(f'{k}=?' for k in fields)} WHERE email=?",
                          [*fields.values(), email])

    def count_active_admins(self) -> int:
        return self.conn.execute("SELECT COUNT(*) FROM app_user WHERE role='Admin' AND active=1").fetchone()[0]

    def user_aliases(self, email: str) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            "SELECT name_key, name FROM user_alias WHERE email=? ORDER BY name", (email,))]

    def set_user_aliases(self, email: str, aliases: dict[str, str]) -> None:
        """aliases: {name_key: name as typed}. Replaces the user's list."""
        self.conn.execute("DELETE FROM user_alias WHERE email=?", (email,))
        for key, name in aliases.items():
            self.conn.execute("INSERT INTO user_alias (email, name_key, name) VALUES (?,?,?)", (email, key, name))

    # ------------------------------------------------------------------ Part 6: weekly reports
    _WEEKLY_COLS = ("created_at", "as_of", "trigger", "created_by", "status", "mode", "recipients",
                    "invalid_recipients", "subject", "html", "data_json", "ai_used", "ai_model", "ai_error",
                    "input_tokens", "output_tokens", "cost_usd", "sent_at", "sent_by", "message", "job_run_id")

    def insert_weekly(self, row: dict) -> int:
        cols = [c for c in self._WEEKLY_COLS if c in row]
        return self.conn.execute(
            f"INSERT INTO weekly_report ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
            [row[c] for c in cols]).lastrowid

    def update_weekly(self, report_id: int, **fields) -> None:
        bad = set(fields) - set(self._WEEKLY_COLS)
        if bad:
            raise KeyError(f"weekly_report has no field(s) {bad}")
        self.conn.execute(f"UPDATE weekly_report SET {', '.join(f'{k}=?' for k in fields)} WHERE id=?",
                          [*fields.values(), report_id])

    def get_weekly(self, report_id: int) -> dict | None:
        r = self.conn.execute("SELECT * FROM weekly_report WHERE id=?", (report_id,)).fetchone()
        return dict(r) if r else None

    def weekly_reports(self, limit: int = 100) -> list[dict]:
        return [dict(r) for r in self.conn.execute(
            "SELECT id, created_at, as_of, trigger, created_by, status, mode, recipients, subject, ai_used,"
            " ai_model, ai_error, input_tokens, output_tokens, cost_usd, sent_at, sent_by, message"
            " FROM weekly_report ORDER BY id DESC LIMIT ?", (limit,))]
