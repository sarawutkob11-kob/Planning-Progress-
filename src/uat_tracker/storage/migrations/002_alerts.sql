-- Part 2: Traffic Light, working days, email alerts, attachments. Part 1 tables are unchanged.

-- Bank holidays (BOT financial-institution holidays, or entered by hand).
CREATE TABLE holiday (
    date TEXT PRIMARY KEY, name TEXT, source TEXT NOT NULL,      -- bot | manual | csv
    fetched_at TEXT NOT NULL
);
CREATE TABLE holiday_fetch (
    id INTEGER PRIMARY KEY, ts TEXT NOT NULL, year INTEGER NOT NULL,
    result TEXT NOT NULL,                                          -- ok | failed
    count INTEGER, message TEXT
);

-- One row per scheduled/manual job run. Missed workdays are recorded as 'skipped'.
CREATE TABLE job_run (
    id INTEGER PRIMARY KEY, job TEXT NOT NULL, run_date TEXT NOT NULL,
    started_at TEXT NOT NULL, finished_at TEXT,
    result TEXT NOT NULL,                  -- running | ok | skipped | failed
    mode TEXT,                             -- dry_run | live
    message TEXT
);
CREATE INDEX ix_job_run_date ON job_run(job, run_date);

-- Computed Traffic Light per Test Progress item. A row is written only when the color changes.
-- color: Complete | Red | Amber | Green | Missing (no color: missing data)
CREATE TABLE traffic_light_history (
    id INTEGER PRIMARY KEY, item_id TEXT NOT NULL, color TEXT NOT NULL, prev_color TEXT,
    pct_exec REAL, pct_time REAL, pct_fail REAL, reason TEXT,
    calc_date TEXT NOT NULL, ts TEXT NOT NULL, job_run_id INTEGER
);
CREATE INDEX ix_light_item ON traffic_light_history(item_id, id);

-- Every email the app sent, previewed, skipped or failed to send.
-- mode: dry_run | live.  result: sent | dry_run | skipped | failed
CREATE TABLE email_log (
    id INTEGER PRIMARY KEY, ts TEXT NOT NULL, send_date TEXT NOT NULL,
    rule TEXT NOT NULL,                    -- R1 R2 R3 R4A R4B MANUAL PMO_ALERT
    variant TEXT,                          -- e.g. today | daily
    sheet TEXT, item_id TEXT, app TEXT,
    to_addr TEXT, cc TEXT, delivered_to TEXT, subject TEXT,
    mode TEXT NOT NULL, result TEXT NOT NULL, error TEXT,
    event_ref INTEGER,                     -- R3: traffic_light_history.id
    preview_path TEXT, job_run_id INTEGER
);
CREATE INDEX ix_email_item ON email_log(item_id, rule);
-- Dedupe: same rule + item + day -> at most one sent (live) and one dry-run email.
CREATE UNIQUE INDEX ux_email_dedupe ON email_log(rule, item_id, send_date, mode)
    WHERE result IN ('sent', 'dry_run') AND rule NOT IN ('MANUAL', 'PMO_ALERT');

-- Test-result documents. Files are never deleted.
CREATE TABLE attachment (
    id INTEGER PRIMARY KEY, sheet TEXT NOT NULL, item_id TEXT NOT NULL,
    original_name TEXT NOT NULL, stored_path TEXT NOT NULL, sha256 TEXT NOT NULL,
    size_bytes INTEGER NOT NULL, uploaded_by TEXT NOT NULL, uploaded_at TEXT NOT NULL
);
CREATE INDEX ix_attachment_item ON attachment(sheet, item_id);

-- App-side values not yet in Excel (Excel write-back comes in Part 3). The import keeps these
-- instead of overwriting them with the older Excel value, and clears the row once Excel matches.
CREATE TABLE pending_writeback (
    sheet TEXT NOT NULL, item_id TEXT NOT NULL, field TEXT NOT NULL,
    value TEXT, created_at TEXT NOT NULL, reason TEXT,
    PRIMARY KEY (sheet, item_id, field)
);
