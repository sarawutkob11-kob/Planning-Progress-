"""Delete an ARCHIVED project for good (approved 25-Nov-2026, spec "ลบโปรเจกต์" D1-D12).

- Admin or PMO (of any project, same rule as Archive). Only a project that is archived. The project name must
  be typed (spaces around it ignored, otherwise exact).
- Removed: the project row in app.db and the whole folder <root>/projects/<id>/ (tracker.db, attachments,
  snapshots, previews, workbook backups kept in that folder, everything else in it).
- Never touched: the project's Excel file (not opened, read, moved or deleted), its old project_event rows
  (one "deleted" event is added), <root>/logs, <root>/clear-backups, a [sync] backup_dir outside the
  folder, other projects and shared settings.
- Before: <root>/delete-backups/<id>_<yyyymmdd-HHMMSS>.zip (tracker.db copied with the SQLite backup API, all
  other files, manifest.json with the project row). Never deleted automatically.
- Fail safely: zip -> rename the folder to projects/.deleting-<id>-<stamp> (in use -> stop, nothing deleted)
  -> app.db transaction (fails -> folder renamed back) -> remove the renamed folder (not all removed ->
  still deleted, the folder left is reported and logged).
"""

from __future__ import annotations

import datetime as dt
import json
import logging
import os
import shutil
import sqlite3
import zipfile
from dataclasses import dataclass
from pathlib import Path

from ..config import Config
from ..storage.app_db import PROJECT_COLS, AppDb
from .auth import now_iso
from .projects import Person, ProjectError, require_manager

log = logging.getLogger(__name__)

BACKUP_DIR_NAME = "delete-backups"
TEMP_PREFIX = ".deleting-"
DB_FILES = ("tracker.db", "tracker.db-wal", "tracker.db-shm")
STAMP_FORMAT = "%Y%m%d-%H%M%S"


class DeleteLocked(ProjectError):
    """Folder could not be renamed (file in use). Nothing deleted; backup_zip is kept."""

    def __init__(self, message: str, backup_zip: Path) -> None:
        super().__init__(message)
        self.backup_zip = backup_zip


@dataclass(frozen=True)
class DeleteResult:
    project_id: str
    name: str
    backup_zip: Path
    leftover: Path | None          # temp folder that could not be fully removed (D5.3)


def project_dir(base: Config, project_id: str) -> Path:
    """<root>/projects/<id>"""
    return Path(base.root_dir) / "projects" / project_id


def backup_dir(base: Config) -> Path:
    """<root>/delete-backups"""
    return Path(base.root_dir) / BACKUP_DIR_NAME


def what_is_deleted(base: Config, project: dict) -> dict[str, str]:
    """For CLI / dialog text: {"folder": ..., "excel": ..., "backup_dir": ...}. Reads nothing."""
    return {"folder": str(project_dir(base, project["id"])), "excel": str(project["excel_path"]),
            "backup_dir": str(backup_dir(base))}


# --------------------------------------------------------------------------- file system (monkeypatch points)

def _rename(src: Path, dst: Path) -> None:
    os.rename(src, dst)


def _rmtree(path: Path) -> None:
    shutil.rmtree(path)


# --------------------------------------------------------------------------- checks

def _excel_inside(project: dict, folder: Path) -> bool:
    try:
        return Path(project["excel_path"]).resolve().is_relative_to(folder.resolve())
    except (OSError, ValueError):
        return False


def _checked_project(base: Config, appdb: AppDb, person: Person, project_id: str, typed_name: str) -> dict:
    """Steps 1-5: nothing is changed and no file is written before all of them pass."""
    require_manager(base, appdb, person)
    project = appdb.get_project(project_id)
    if project is None:
        raise ProjectError(f"Unknown project {project_id}")
    if project["archived"] != 1:
        raise ProjectError("Archive the project before deleting it")
    if (typed_name or "").strip() != (project["name"] or "").strip():
        raise ProjectError("The typed name does not match the project name")
    if _excel_inside(project, project_dir(base, project_id)):
        raise ProjectError(f"The Excel file {project['excel_path']} is inside the project's data folder. "
                           "Move it out and Relink the project first: the Excel file is never deleted")
    return project


# --------------------------------------------------------------------------- backup

def _zip_path(out_dir: Path, project_id: str, stamp: str) -> Path:
    zpath = out_dir / f"{project_id}_{stamp}.zip"
    n = 1
    while zpath.exists():                              # never overwrite an older backup
        zpath = out_dir / f"{project_id}_{stamp}-{n}.zip"
        n += 1
    return zpath


def _copy_db(src_path: Path, dst_path: Path) -> None:
    """Consistent copy with the SQLite backup API (includes what is still in the -wal file)."""
    src = sqlite3.connect(src_path)
    try:
        dst = sqlite3.connect(dst_path)
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()


def _manifest(base: Config, project: dict, person: Person) -> str:
    man = {c: project.get(c) for c in PROJECT_COLS}
    man.update({"deleted_at": now_iso(), "deleted_by": person.editor,
                "data_dir": str(project_dir(base, project["id"]))})
    return json.dumps(man, ensure_ascii=False, indent=1)


def _fill_zip(z: zipfile.ZipFile, folder: Path | None, tmp_db: Path) -> None:
    if folder is None:
        return
    db = folder / "tracker.db"
    if db.is_file():
        _copy_db(db, tmp_db)
        z.write(tmp_db, "tracker.db")
    for f in sorted(folder.rglob("*")):
        if not f.is_file() or (f.parent == folder and f.name in DB_FILES):
            continue
        z.write(f, f.relative_to(folder).as_posix())


def _write_backup(base: Config, project: dict, folder: Path | None, person: Person, stamp: str) -> Path:
    """Zip written as <name>.zip.part then os.replace -> <name>.zip. Any error: .part and temp db removed,
    re-raised."""
    out_dir = backup_dir(base)
    out_dir.mkdir(parents=True, exist_ok=True)
    zpath = _zip_path(out_dir, project["id"], stamp)
    part = zpath.with_name(zpath.name + ".part")
    tmp_db = out_dir / f".{zpath.stem}.db"
    try:
        with zipfile.ZipFile(part, "w", zipfile.ZIP_DEFLATED) as z:
            _fill_zip(z, folder, tmp_db)
            z.writestr("manifest.json", _manifest(base, project, person))
        os.replace(part, zpath)
    except BaseException:
        part.unlink(missing_ok=True)
        raise
    finally:
        tmp_db.unlink(missing_ok=True)
    return zpath


# --------------------------------------------------------------------------- delete

def _remove_row(appdb: AppDb, project: dict, person: Person, zpath: Path) -> None:
    with appdb.transaction():
        appdb.delete_project(project["id"])
        appdb.log_event(ts=now_iso(), project_id=project["id"], editor=person.editor, action="deleted",
                        detail=f"name={project['name']}; excel={project['excel_path']}; backup={zpath}")


def _rename_back(tmp: Path | None, folder: Path) -> str:
    """'' when the folder is back (or there was none), else a note for the error message."""
    if tmp is None:
        return ""
    try:
        _rename(tmp, folder)
        return ""
    except OSError:
        log.error("Could not rename %s back to %s", tmp, folder, exc_info=True)
        return f" The project's files are in {tmp}: rename that folder back to {folder.name}."


def _remove_folder(tmp: Path | None) -> Path | None:
    """-> the folder left behind (files in use), or None."""
    if tmp is None:
        return None
    try:
        _rmtree(tmp)
    except OSError:
        log.warning("Project deleted, but not all of its files could be removed. Folder left: %s", tmp,
                    exc_info=True)
    return tmp if tmp.exists() else None


def delete_project(base: Config, appdb: AppDb, person: Person, project_id: str, *,
                   typed_name: str) -> DeleteResult:
    """Delete one ARCHIVED project for good (D1-D9). Raises ProjectError (DeleteLocked when the folder is in
    use). The Excel file is never opened or touched."""
    project = _checked_project(base, appdb, person, project_id, typed_name)
    folder = project_dir(base, project_id)
    has_folder = folder.is_dir()
    stamp = dt.datetime.now().strftime(STAMP_FORMAT)
    try:
        zpath = _write_backup(base, project, folder if has_folder else None, person, stamp)
    except (OSError, sqlite3.Error, zipfile.BadZipFile) as exc:
        log.error("Backup before deleting project %s failed", project_id, exc_info=True)
        raise ProjectError(f"Backup failed, nothing deleted: {exc}") from exc
    tmp = None
    if has_folder:
        tmp = folder.parent / f"{TEMP_PREFIX}{project_id}-{stamp}"
        try:
            _rename(folder, tmp)
        except OSError as exc:
            log.warning("Project folder %s is in use, nothing deleted", folder, exc_info=True)
            raise DeleteLocked("Files of this project are in use: close other app windows / Excel / Explorer "
                               f"and try again. Nothing was deleted. Backup kept: {zpath}", zpath) from exc
    try:
        _remove_row(appdb, project, person, zpath)
    except Exception as exc:
        log.error("app.db update for deleting project %s failed", project_id, exc_info=True)
        note = _rename_back(tmp, folder)
        raise ProjectError(f"Could not update app.db, nothing deleted: {exc}.{note}") from exc
    leftover = _remove_folder(tmp)
    log.info("Project %s '%s' deleted by %s. Backup: %s", project_id, project["name"], person.editor, zpath)
    return DeleteResult(project_id, project["name"], zpath, leftover)
