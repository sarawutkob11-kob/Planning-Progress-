-- Part 1 schema. Later parts add tables via 002_*.sql, 003_*.sql ...
-- Dates are stored as ISO text (YYYY-MM-DD) for sorting; always shown as DD-Mon-YYYY.

CREATE TABLE import_run (
    id              INTEGER PRIMARY KEY,
    started_at      TEXT NOT NULL,
    finished_at     TEXT,
    editor          TEXT NOT NULL,
    source_path     TEXT NOT NULL,
    source_mtime    TEXT,
    source_sha256   TEXT,
    snapshot_path   TEXT,
    result          TEXT,            -- running | ok | failed
    message         TEXT,
    stats_json      TEXT
);

CREATE TABLE uat_task (
    id TEXT PRIMARY KEY, excel_row INTEGER, fingerprint TEXT, alt_key TEXT,
    active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
    group_code TEXT, group_title TEXT, task_no TEXT, sub_task TEXT, task_name TEXT, note TEXT,
    file_path TEXT, responsible TEXT, status TEXT,
    start_date TEXT, start_date_note TEXT, start_time TEXT,
    finish_date TEXT, finish_date_note TEXT, finish_time TEXT,
    remark TEXT
);

CREATE TABLE test_progress (
    id TEXT PRIMARY KEY, excel_row INTEGER, fingerprint TEXT, alt_key TEXT,
    active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
    no TEXT, application TEXT, report_by TEXT, report_by_email TEXT, cc_emails TEXT,
    total_tc NUMERIC, executed NUMERIC, pct_progress NUMERIC, pass_count NUMERIC, fail_count NUMERIC,
    planned_start TEXT, planned_start_note TEXT, planned_finish TEXT, planned_finish_note TEXT,
    actual_start TEXT, actual_start_note TEXT, actual_finish TEXT, actual_finish_note TEXT,
    traffic_light TEXT, tc_result_complete TEXT, status TEXT, remark TEXT
);

CREATE TABLE defect (
    id TEXT PRIMARY KEY, excel_row INTEGER, fingerprint TEXT, alt_key TEXT,
    active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
    no TEXT, defect_date TEXT, defect_date_note TEXT, system TEXT, detail TEXT, status TEXT,
    severity TEXT, assigned_to TEXT, fix_target_date TEXT, fix_target_date_note TEXT,
    data_test TEXT, raised_by TEXT, remark TEXT
);

-- Reschedule history: prior values of a date field, in typed order (seq).
-- source = import (parsed from the Excel cell) | app (Part 3, with reason + root cause).
CREATE TABLE date_history (
    id INTEGER PRIMARY KEY, sheet TEXT NOT NULL, item_id TEXT NOT NULL, field TEXT NOT NULL,
    seq INTEGER NOT NULL, value TEXT NOT NULL, source TEXT NOT NULL,
    reason TEXT, root_cause TEXT, recorded_at TEXT NOT NULL
);
CREATE INDEX ix_date_history_item ON date_history(sheet, item_id, field);

-- Raw cell text last read from Excel, per item/field. Basis for conflict detection (Part 4).
CREATE TABLE cell_baseline (
    sheet TEXT NOT NULL, item_id TEXT NOT NULL, field TEXT NOT NULL,
    raw_value TEXT, read_at TEXT NOT NULL,
    PRIMARY KEY (sheet, item_id, field)
);

-- Immutable audit log. Restore = new entry. hash = sha256(prev_hash + entry) for tamper evidence.
CREATE TABLE audit_log (
    id INTEGER PRIMARY KEY, ts TEXT NOT NULL, editor TEXT NOT NULL, channel TEXT NOT NULL,
    sheet TEXT, item_id TEXT, field TEXT, old_value TEXT, new_value TEXT, reason TEXT,
    import_run_id INTEGER, prev_hash TEXT, hash TEXT NOT NULL
);
CREATE INDEX ix_audit_item ON audit_log(sheet, item_id);
CREATE TRIGGER audit_log_no_update BEFORE UPDATE ON audit_log
BEGIN SELECT RAISE(ABORT, 'audit_log is immutable'); END;
CREATE TRIGGER audit_log_no_delete BEFORE DELETE ON audit_log
BEGIN SELECT RAISE(ABORT, 'audit_log is immutable'); END;

-- PMO review queue: values the app refused to guess.
CREATE TABLE review_queue (
    id INTEGER PRIMARY KEY, created_at TEXT NOT NULL, sheet TEXT NOT NULL, item_id TEXT NOT NULL,
    excel_row INTEGER, field TEXT NOT NULL, raw_value TEXT, token TEXT, problem TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'open',      -- open | resolved | superseded
    resolved_value TEXT, resolved_by TEXT, resolved_at TEXT, import_run_id INTEGER
);
CREATE INDEX ix_review_item ON review_queue(sheet, item_id);

-- Data-quality flags (rule 7).
CREATE TABLE data_flag (
    id INTEGER PRIMARY KEY, sheet TEXT NOT NULL, item_id TEXT NOT NULL, code TEXT NOT NULL,
    detail TEXT, active INTEGER NOT NULL DEFAULT 1, first_seen TEXT NOT NULL, last_seen TEXT NOT NULL,
    UNIQUE (sheet, item_id, code)
);

CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT);
