"""TEMPORARY "Clear project data" (multi-project add-on, section I). For testing only.

Off by default ([features] enable_clear_project = false). The full list of what to delete to remove this
feature is in README ("Removing the Clear project feature").

Approved 10-Nov-2026, changed by Hardening phase 1 (11-Nov-2026):
- Admin + PMO of the selected project only. The user types the project name to enable the button.
- Clears the SELECTED project's app-side data: tasks, applications, defects, reschedule history, Traffic Light
  history, sync / conflict / review queue / flags, alert + email log, weekly report history,
  job runs, import runs, Excel snapshots, cached attachments (files too).
- Keeps: the AUDIT LOG (never deleted; one entry "project data cleared by <actor>, backup <path>" is added),
  the project (name, file link), users + roles + names in Excel, per-project settings (X%, dry-run,
  send time, ...), reminder contacts, workbook backups, the Excel file itself (never opened), shared settings,
  other projects.
- Before: zip of the project DB + attachment folder (timestamped), path shown. After: reading Excel is paused
  until "Import from Excel" (approved Q10). Also logged in <root>/clear-project.log.
- Restore: restore_backup(zip) (CLI `project restore-clear <zip>`, approved Q9). The restored DB is from before
  the clear, so an audit entry "project data restored from <zip> ..." is added to it.
- Live mode (dry_run = false) with the feature on: red warning in Settings + a log warning when a project opens.
"""

from __future__ import annotations

import datetime as dt
import json
import logging
import shutil
import sqlite3
import zipfile
from dataclasses import dataclass
from pathlib import Path

from ..domain.models import CHANNEL_APP
from .auth import ADMIN, PMO, current_role, now_iso
from .importer import S_PAUSED

log = logging.getLogger(__name__)

ALLOWED_ROLES = (ADMIN, PMO)
CLEARED_TABLES = ("uat_task", "test_progress", "defect", "date_history", "cell_baseline", "review_queue",
                  "data_flag", "import_run", "job_run", "traffic_light_history", "email_log", "attachment",
                  "pending_writeback", "conflict", "sync_event", "weekly_report")
KEPT = ("audit log (+ one 'cleared' entry)", "project name + Excel file link", "users, roles, names in Excel",
        "settings: Fail X%, dry-run, send time", "reminder contacts", "workbook backups",
        "the Excel file (not touched)", "shared settings", "other projects")
AUDIT_SHEET = "Project"
AUDIT_FIELD_CLEAR = "clear_project_data"
AUDIT_FIELD_RESTORE = "restore_clear_backup"
CLEARED_SETTING_PREFIXES = ("sync.", "alerts.")      # sync / alert state; settings.* and ui.* stay
CLEARED_DIRS = ("attachments", "snapshots", "previews")
LOG_NAME = "clear-project.log"


class ClearError(Exception):
    pass


@dataclass
class ClearResult:
    backup_zip: Path
    counts: dict


def enabled(cfg) -> bool:
    return bool(getattr(cfg, "enable_clear_project", False))


def live_warning(cfg) -> bool:
    """Feature on while emails go to real recipients (dry_run = false): show the red warning."""
    return enabled(cfg) and not cfg.alerts.dry_run


def warn_if_live(cfg) -> bool:
    if live_warning(cfg):
        log.warning("Clear project data is ON in live mode (dry_run = false) for project %s (%s). "
                    "Set [features] enable_clear_project = false.", cfg.project_name or "-", cfg.project_id or "-")
        return True
    return False


def _audit(repo, *, editor: str, field: str, item_id, old="", new="", reason="") -> None:
    repo.append_audit({"ts": now_iso(), "editor": editor, "channel": CHANNEL_APP, "sheet": AUDIT_SHEET,
                       "item_id": item_id, "field": field, "old_value": old, "new_value": new, "reason": reason})


def allowed(repo, actor) -> bool:
    return actor is not None and current_role(repo, actor) in ALLOWED_ROLES


def counts(repo) -> dict:
    return {t: repo.conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] for t in CLEARED_TABLES}


def _backup(cfg, repo, stamp: str) -> Path:
    out_dir = cfg.root_dir / "clear-backups"
    out_dir.mkdir(parents=True, exist_ok=True)
    zpath = out_dir / f"{cfg.project_id or 'project'}_{stamp}.zip"
    tmp_db = out_dir / f".{stamp}.db"
    src = sqlite3.connect(cfg.db_path)                  # consistent copy, even with the app's connection open
    dst = sqlite3.connect(tmp_db)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    try:
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(tmp_db, "tracker.db")
            att = cfg.attachment_dir
            if att.exists():
                for f in att.rglob("*"):
                    if f.is_file():
                        z.write(f, "attachments/" + f.relative_to(att).as_posix())
            z.writestr("manifest.json", json.dumps({
                "project_id": cfg.project_id, "project_name": cfg.project_name, "created_at": now_iso(),
                "db_path": str(cfg.db_path), "attachment_dir": str(att)}, ensure_ascii=False, indent=1))
    finally:
        tmp_db.unlink(missing_ok=True)
    return zpath


def _log(cfg, line: dict) -> None:
    with open(cfg.root_dir / LOG_NAME, "a", encoding="utf-8") as f:
        f.write(json.dumps(line, ensure_ascii=False) + "\n")


def clear_project(cfg, repo, actor, *, typed_name: str) -> ClearResult:
    if not enabled(cfg):
        raise ClearError("Clear project data is turned off ([features] enable_clear_project)")
    if not allowed(repo, actor):
        raise ClearError("Not allowed: only an Admin or PMO of this project can clear its data")
    if typed_name.strip() != (cfg.project_name or "").strip():
        raise ClearError("The typed name does not match the project name")
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    before = counts(repo)
    zpath = _backup(cfg, repo, stamp)
    with repo.transaction():
        for t in CLEARED_TABLES:                         # the audit log is never deleted
            repo.conn.execute(f"DELETE FROM {t}")
        for prefix in CLEARED_SETTING_PREFIXES:
            repo.conn.execute("DELETE FROM main.settings WHERE key LIKE ?", (prefix + "%",))
        repo.set_setting(S_PAUSED, now_iso())
        _audit(repo, editor=actor.editor, field=AUDIT_FIELD_CLEAR, item_id=cfg.project_id,
               old=json.dumps(before, sort_keys=True), new=str(zpath),
               reason=f"project data cleared by {actor.editor}, backup {zpath}")
    for d in CLEARED_DIRS:
        p = cfg.data_dir / d
        if p.exists():
            shutil.rmtree(p)
    _log(cfg, {"ts": now_iso(), "action": "clear_project_data", "project_id": cfg.project_id,
               "project_name": cfg.project_name, "user": actor.editor, "backup": str(zpath), "rows_cleared": before})
    return ClearResult(zpath, before)


def restore_backup(cfg, zip_path: Path, *, editor: str) -> dict:
    """Put a clear-project backup back (DB + attachments). The app must be closed for this project.
    The current DB is kept next to it as tracker.db.before-restore-<stamp>."""
    zip_path = Path(zip_path)
    with zipfile.ZipFile(zip_path) as z:
        man = json.loads(z.read("manifest.json"))
        if man.get("project_id") != cfg.project_id:
            raise ClearError(f"This backup is for project {man.get('project_id')!r}, not {cfg.project_id!r}")
        stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
        if cfg.db_path.exists():
            shutil.copy2(cfg.db_path, cfg.db_path.with_name(f"tracker.db.before-restore-{stamp}"))
        tmp = cfg.db_path.with_name(f".restore-{stamp}.db")
        tmp.write_bytes(z.read("tracker.db"))
        src = sqlite3.connect(tmp)
        dst = sqlite3.connect(cfg.db_path)
        try:
            src.backup(dst)
        finally:
            dst.close()
            src.close()
            tmp.unlink(missing_ok=True)
        for name in z.namelist():
            if name.startswith("attachments/") and not name.endswith("/"):
                target = cfg.attachment_dir / name[len("attachments/"):]
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(z.read(name))
    from ..storage.sqlite_repo import SqliteRepository
    repo = SqliteRepository(cfg.db_path, shared=cfg.shared_db)     # the restored DB: note the restore in it
    try:
        with repo.transaction():
            _audit(repo, editor=editor, field=AUDIT_FIELD_RESTORE, item_id=cfg.project_id, new=str(zip_path),
                   reason=f"project data restored from {zip_path} by {editor}, "
                          f"previous clear at {man.get('created_at')}")
    finally:
        repo.close()
    _log(cfg, {"ts": now_iso(), "action": "restore_clear_backup", "project_id": cfg.project_id,
               "project_name": cfg.project_name, "user": editor, "backup": str(zip_path)})
    return man
