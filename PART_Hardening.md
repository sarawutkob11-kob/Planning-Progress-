# Planning Progress — Add-on: Hardening, phase 1 (from code review 11-Nov-2026)

## Rules
- Parts 1–6 + Multi-project done. FIRST read existing code, DB schema, tests, settings. Reuse. Mismatch with this prompt → ask, don't rewrite.
- Unspecified → STOP, ask. Never decide alone.
- Step 1: reply with design per section, affected files, open questions. Wait approval before code.
- Backup DB + every project Excel file before any migration. Never delete data. Test on copies first.
- Every change: add/adjust tests. All 165+ existing tests must still pass.
- Do sections in order A → E. Stop after each section, show summary + test result, wait "go".

## A. Repo hygiene (do first)
- Add `.gitignore`: `__pycache__/`, `*.pyc`, `config.toml`, `*.db`, `*.db-wal`, `*.db-shm`, `*.xlsx`, `*.xlsm`, `previews/`, `snapshots/`, `backups/`, `logs/`, `.venv/`, `venv/`.
- Remove from git index only (keep local files): `git rm -r --cached` for all `__pycache__/` and `config.toml`. Keep `config.example.toml`.
- `config.example.toml`: replace company OneDrive path with placeholder `%USERPROFILE%/OneDrive - <Company>/<file>.xlsx`.
- Scan repo for real emails, names, keys. Report findings; don't change test fixtures without asking.
- Show me the git commands before running. Don't push.

## B. Clear-project button (temporary feature)
- `enable_clear_project` default = `false` in `config.example.toml` and in code default.
- Clear must NOT delete audit log. Instead: keep audit rows, add one audit entry "project data cleared by <actor>, backup <path>".
- Startup check: if flag on and `dry_run = false` (live mode) → show red warning banner in Settings + log warning.
- Keep module isolated so it can be deleted later (list exact files/lines to delete in README).

## C. Model name to config
- `services/weekly_ai.py`: model name hardcoded (many places). Move to config `[weekly_ai] model = "..."`, current value as default.
- Remove hardcoded fallback model list from code except one documented default. Unknown/retired model error → clear message + existing fallback (report without AI).
- Do NOT change what data is sent to the API (privacy work is a later part).

## D. Reliability
- Logging: replace `print()` in non-CLI-output code with `logging`. RotatingFileHandler `<data_dir>/logs/app.log` (5 MB × 5). Scheduled jobs log to `<data_dir>/logs/jobs.log`. Every `except Exception` logs `exc_info=True` (keep behavior, add log). Includes `ui/app.py` flush-on-switch `except Exception: pass`: add log only, no UI change (UI change is a later part). CLI user-facing output may keep print.
- SQLite: enable `PRAGMA journal_mode=WAL` on project DB and `app.db` (GUI + scheduled job run at same time). Keep busy_timeout. Test concurrent read/write from 2 processes.
- Excel COM busy: handle `RPC_E_CALL_REJECTED` (0x80010001) and `RPC_E_SERVERCALL_RETRYLATER` (0x8001010A): retry with backoff (max ~30 s), then raise `WorkbookBusy` → existing retry queue. Add test with fake COM error.
- UI freeze: move Import/pull, restore, relink, migrate, clear to worker QThread (same pattern as Weekly). Show progress dialog, block double-click, refresh screen on finish. SyncBusy → friendly message "Sync running, try again in a moment".

## E. Naming + dependencies
- Replace leftover "UAT Tracker" / "Progress Planning" in `services/scheduler.py`, `ui/app.py` (Task Scheduler task names, titles). Package name `uat_tracker` stays (approved).
- Existing Task Scheduler tasks with old names: on next install → create new name, delete old task. Show me old/new names first.
- Pin dependencies: add upper bounds in `pyproject.toml` (next major) + generate `requirements.lock` from current working venv. README: install from lock file.
- State supported Python version (tested one) in README.

## Not in scope (later parts, don't touch)
- Sync for projects not currently open; pending-write UI warning; missed 08:30 alerts (waiting business answer).
- Masking personal data before Weekly Report AI (waiting InfoSec).
- Real login (Windows/AD identity). Phase 2. Write short README note on current limit: name/email login = claimed identity.
- Holiday-service key in SQLite plain text: move to Windows Credential Manager like Claude key — ask me if in scope.

## Done when
- Sections A–E approved + done, all tests pass, README updated (new config keys, logs location, how to remove clear feature).
- Final reply: changed files list, new config keys, manual test checklist for PMO.
