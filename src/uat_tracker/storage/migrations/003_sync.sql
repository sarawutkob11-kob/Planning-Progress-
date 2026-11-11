-- Part 3: 2-way sync, conflicts, backups. Part 1-2 tables are kept; pending_writeback (Part 2)
-- becomes the write queue App -> Excel.

ALTER TABLE pending_writeback ADD COLUMN excel_value TEXT;      -- JSON {"kind","value","hist"}; NULL = derive
ALTER TABLE pending_writeback ADD COLUMN editor TEXT;
ALTER TABLE pending_writeback ADD COLUMN channel TEXT;
ALTER TABLE pending_writeback ADD COLUMN status TEXT NOT NULL DEFAULT 'pending';   -- pending | conflict | failed
ALTER TABLE pending_writeback ADD COLUMN attempts INTEGER NOT NULL DEFAULT 0;
ALTER TABLE pending_writeback ADD COLUMN first_failed_at TEXT;
ALTER TABLE pending_writeback ADD COLUMN next_attempt_at TEXT;
ALTER TABLE pending_writeback ADD COLUMN last_error TEXT;

-- Same cell changed in Excel and in the app since the last sync, or an Excel edit that breaks a rule
-- (Actual Finish > Planned Finish). While open, neither side is overwritten.
CREATE TABLE conflict (
    id INTEGER PRIMARY KEY, created_at TEXT NOT NULL,
    sheet TEXT NOT NULL, item_id TEXT NOT NULL, field TEXT NOT NULL,
    kind TEXT NOT NULL,                     -- both_changed | actual_after_planned
    excel_raw TEXT, excel_value TEXT, excel_time TEXT,
    app_value TEXT, app_time TEXT, app_editor TEXT,
    status TEXT NOT NULL DEFAULT 'open',    -- open | resolved
    resolution TEXT,                        -- excel | app | new
    resolved_value TEXT, reason TEXT, resolved_by TEXT, resolved_at TEXT
);
CREATE UNIQUE INDEX ux_conflict_open ON conflict(sheet, item_id, field) WHERE status = 'open';

-- Copy of the workbook taken before every write / restore. Never deleted (approved).
CREATE TABLE backup (
    id INTEGER PRIMARY KEY, ts TEXT NOT NULL, path TEXT NOT NULL, sha256 TEXT NOT NULL,
    size_bytes INTEGER NOT NULL, reason TEXT
);

-- Sync activity: pulls (Excel -> app), writes (app -> Excel), errors, blocks.
CREATE TABLE sync_event (
    id INTEGER PRIMARY KEY, ts TEXT NOT NULL, kind TEXT NOT NULL,    -- pull | write | error | block | restore
    result TEXT NOT NULL, message TEXT
);
