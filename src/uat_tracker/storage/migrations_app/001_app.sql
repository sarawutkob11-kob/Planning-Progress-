-- Multi-project add-on (approved 10-Nov-2026): app.db is shared by all projects on this PC.
-- Each project keeps its own tracker.db (items, users + roles, per-project settings, logs).

CREATE TABLE project (
    id           TEXT PRIMARY KEY,              -- p001, p002 ... (folder name under projects/)
    name         TEXT NOT NULL,                 -- display name, as typed (may contain / : * ...)
    excel_path   TEXT NOT NULL,
    created_at   TEXT NOT NULL,
    created_by   TEXT NOT NULL,
    archived     INTEGER NOT NULL DEFAULT 0,    -- archive only, never deleted (approved)
    archived_at  TEXT,
    window_start TEXT,                          -- date window; NULL = config.toml default
    window_end   TEXT
);
CREATE TRIGGER project_no_delete BEFORE DELETE ON project
BEGIN SELECT RAISE(ABORT, 'projects are archived, never deleted'); END;

-- Project-level actions (create, import, relink, archive, rename, migrate).
CREATE TABLE project_event (
    id INTEGER PRIMARY KEY, ts TEXT NOT NULL, project_id TEXT, editor TEXT NOT NULL,
    action TEXT NOT NULL, detail TEXT
);

-- Shared: bank holidays (same shape as the project DB tables, which are no longer used once attached).
CREATE TABLE holiday (
    date TEXT PRIMARY KEY, name TEXT, source TEXT NOT NULL, fetched_at TEXT NOT NULL
);
CREATE TABLE holiday_fetch (
    id INTEGER PRIMARY KEY, ts TEXT NOT NULL, year INTEGER NOT NULL,
    result TEXT NOT NULL, count INTEGER, message TEXT
);

-- Shared settings: settings.pmo_email, ui.lang, settings.ai_model.
CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT);
