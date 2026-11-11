-- Part 5: login (Outlook account + one-time code) and roles. Users are never deleted (deactivate only).

CREATE TABLE app_user (
    email TEXT PRIMARY KEY,                 -- lower case
    display_name TEXT,
    role TEXT NOT NULL CHECK (role IN ('Admin', 'PMO', 'QA', 'User BU/PM')),
    active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL, created_by TEXT NOT NULL,
    last_login_at TEXT
);
CREATE TRIGGER app_user_no_delete BEFORE DELETE ON app_user
BEGIN SELECT RAISE(ABORT, 'users are deactivated, never deleted'); END;

-- "Names as written in Excel" (UAT Plan Responsible, Defect Assigned To) -> user. Approved option A.
CREATE TABLE user_alias (
    email TEXT NOT NULL REFERENCES app_user(email),
    name_key TEXT NOT NULL,                 -- case-folded, spaces collapsed
    name TEXT NOT NULL,
    PRIMARY KEY (email, name_key)
);
CREATE INDEX ix_alias_key ON user_alias(name_key);

-- Wrong-code counter and lockout, per email (also for emails that are not users yet).
CREATE TABLE login_state (
    email TEXT PRIMARY KEY, failed_count INTEGER NOT NULL DEFAULT 0, locked_until TEXT
);

-- One-time codes: only a salted SHA-256 hash is stored. status: active | used | expired | cancelled.
CREATE TABLE otp (
    id INTEGER PRIMARY KEY, email TEXT NOT NULL, purpose TEXT NOT NULL,   -- login | bootstrap
    code_hash TEXT NOT NULL, salt TEXT NOT NULL,
    created_at TEXT NOT NULL, expires_at TEXT NOT NULL,
    attempts INTEGER NOT NULL DEFAULT 0, status TEXT NOT NULL DEFAULT 'active', used_at TEXT
);
CREATE INDEX ix_otp_email ON otp(email, id);

CREATE TABLE session (
    id INTEGER PRIMARY KEY, email TEXT NOT NULL, client TEXT NOT NULL,   -- app | cli
    started_at TEXT NOT NULL, last_seen_at TEXT NOT NULL,
    ended_at TEXT, end_reason TEXT                                       -- logout | idle | expired | closed
);
