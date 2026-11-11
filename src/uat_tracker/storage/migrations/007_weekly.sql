-- Part 6: every weekly report is kept (approved 10-Nov-2026): what was sent, to whom, with or without AI.

CREATE TABLE weekly_report (
    id INTEGER PRIMARY KEY,
    created_at TEXT NOT NULL,
    as_of TEXT NOT NULL,                     -- report date (data as of this day)
    trigger TEXT NOT NULL,                   -- scheduled | manual
    created_by TEXT NOT NULL,
    status TEXT NOT NULL,                    -- preview | sent | dry_run | failed | skipped | cancelled
    mode TEXT,                               -- dry_run | live (when sent)
    recipients TEXT,                         -- JSON list (deduped, valid)
    invalid_recipients TEXT,                 -- JSON list of {"app", "value"}
    subject TEXT,
    html TEXT,
    data_json TEXT,                          -- numbers / lists computed by the tracker
    ai_used INTEGER NOT NULL DEFAULT 0,
    ai_model TEXT, ai_error TEXT,
    input_tokens INTEGER, output_tokens INTEGER, cost_usd REAL,
    sent_at TEXT, sent_by TEXT, message TEXT, job_run_id INTEGER
);
CREATE INDEX ix_weekly_asof ON weekly_report(as_of);
