"""One-time move to multi-project (section E, approved 10-Nov-2026).

The single-project data folder (%LOCALAPPDATA%\\UATTracker) becomes project "RAMC UAT" (id p001) under the new
root (%LOCALAPPDATA%\\PlanningProgress\\projects\\p001), and 'UAT Planning.xlsx' is renamed to
'Planning Progress RAMC UAT.xlsx' in the same OneDrive folder.

Order (nothing is deleted; the old folder stays as it is):
 1. checks: old DB exists, new root has no projects yet, workbook exists, target name free, workbook not open
    in Excel on this PC;
 2. backups: zip of the whole old folder + a copy of the workbook (new root \\ migration-backup);
 3. counts before: every table + the Traffic Light of every application;
 4. rename the workbook (only with confirm=True; the CLI shows old / new path and asks first);
 5. copy the old folder into projects\\p001 (DB copied with SQLite's backup API), create app.db with the
    project, move holidays + shared settings (PMO email, language, Claude model) into app.db, profile.json;
 6. counts after, through the new code path: must be identical, else the workbook is renamed back and the
    new project is archived (nothing is deleted).
"""

from __future__ import annotations

import logging
import datetime as dt
import json
import shutil
import sqlite3
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from ..excel.reader import copy_shared, sha256_file
from ..excel.sheets import TEST_PROGRESS
from .auth import now_iso
from .projects import ProjectError, file_name_for, open_app_db, open_project
from .traffic_light import compute

log = logging.getLogger(__name__)

SHARED_KEYS = ("settings.pmo_email", "ui.lang", "settings.ai_model")
SKIP_FILES = {"sync.lock", "profile.json", "tracker.db-wal", "tracker.db-shm"}   # the DB is copied with backup()


@dataclass
class MigrationReport:
    old_dir: Path
    new_root: Path
    old_excel: Path
    new_excel: Path
    project_id: str = "p001"
    backups: list[str] = field(default_factory=list)
    before: dict = field(default_factory=dict)
    after: dict = field(default_factory=dict)
    lights_before: dict = field(default_factory=dict)
    lights_after: dict = field(default_factory=dict)
    done: bool = False
    message: str = ""

    @property
    def same(self) -> bool:
        return self.before == self.after and self.lights_before == self.lights_after


def _tables(conn) -> list[str]:
    return [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
                                        " ORDER BY name")]


def _counts(conn, schema: str = "main") -> dict:
    return {t: conn.execute(f"SELECT COUNT(*) FROM {schema}.{t}").fetchone()[0] for t in _tables(conn)}


def _lights(conn, today: dt.date, fail_threshold: float) -> dict:
    conn.row_factory = sqlite3.Row
    out = {}
    for r in conn.execute("SELECT * FROM test_progress WHERE active=1"):
        out[r["id"]] = compute(dict(r), today, fail_threshold).color
    return out


def _fail_threshold(conn, default: float) -> float:
    r = conn.execute("SELECT value FROM settings WHERE key='settings.fail_threshold'").fetchone()
    return float(r[0]) if r and r[0] else default


def plan(base, old_dir: Path, project_name: str = "RAMC UAT") -> MigrationReport:
    old_dir = Path(old_dir)
    old_db = old_dir / "tracker.db"
    if not old_db.exists():
        raise ProjectError(f"Old data not found: {old_db}")
    conn = sqlite3.connect(f"file:{old_db}?mode=ro", uri=True)
    try:
        r = conn.execute("SELECT value FROM settings WHERE key='settings.excel_path'").fetchone()
    finally:
        conn.close()
    old_excel = Path(r[0]) if r and r[0] else Path(base.excel_path)
    new_excel = old_excel.with_name(file_name_for(project_name))
    return MigrationReport(old_dir, base.root_dir, old_excel, new_excel)


def check(base, rep: MigrationReport) -> list[str]:
    from ..excel.writer import excel_windows_for
    problems = []
    appdb = open_app_db(base)
    try:
        if appdb.projects(include_archived=True):
            problems.append(f"{base.root_dir} already has projects: the migration was done before")
    finally:
        appdb.close()
    if not rep.old_excel.exists():
        problems.append(f"Workbook not found: {rep.old_excel}")
    if rep.new_excel.exists() and rep.new_excel != rep.old_excel:
        problems.append(f"Target name already exists: {rep.new_excel}")
    if rep.old_excel.exists() and excel_windows_for(rep.old_excel):
        problems.append("The workbook is open in Excel on this PC: close it first (others may keep it open "
                        "elsewhere; they will see the new name)")
    return problems


def migrate(base, old_dir: Path, *, project_name: str = "RAMC UAT", confirm: bool = False,
            today: dt.date | None = None, editor: str = "migration") -> MigrationReport:
    rep = plan(base, old_dir, project_name)
    problems = check(base, rep)
    if problems:
        raise ProjectError("Migration stopped:\n- " + "\n- ".join(problems))
    if not confirm:
        rep.message = "dry run: nothing changed"
        return rep
    today = today or dt.date.today()
    root = base.root_dir
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    bdir = root / "migration-backup"
    bdir.mkdir(parents=True, exist_ok=True)

    # 2. backups
    zpath = bdir / f"{rep.old_dir.name}_{stamp}.zip"
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for f in rep.old_dir.rglob("*"):
            if f.is_file() and f.name != "sync.lock":
                z.write(f, f.relative_to(rep.old_dir).as_posix())
    xcopy = bdir / f"{rep.old_excel.stem}_{stamp}{rep.old_excel.suffix}"
    copy_shared(rep.old_excel, xcopy)
    if sha256_file(xcopy) != sha256_file(rep.old_excel):
        raise ProjectError("Workbook backup does not match the original; nothing was changed")
    rep.backups = [str(zpath), str(xcopy)]

    # 3. counts before
    old = sqlite3.connect(f"file:{rep.old_dir / 'tracker.db'}?mode=ro", uri=True)
    try:
        rep.before = _counts(old)
        x = _fail_threshold(old, base.alerts.fail_threshold)
        rep.lights_before = _lights(old, today, x)
        shared_vals = {k: v for k, v in old.execute(
            f"SELECT key, value FROM settings WHERE key IN ({','.join('?' * len(SHARED_KEYS))})", SHARED_KEYS)}
        holidays = old.execute("SELECT date, name, source, fetched_at FROM holiday").fetchall()
        fetches = old.execute("SELECT ts, year, result, count, message FROM holiday_fetch ORDER BY id").fetchall()
        admin = old.execute("SELECT email, first_name, last_name FROM app_user WHERE active=1 "
                            "ORDER BY role='Admin' DESC, last_used_at DESC").fetchone()
    finally:
        old.close()

    # 4. rename the workbook
    renamed = False
    if rep.new_excel != rep.old_excel:
        rep.old_excel.rename(rep.new_excel)
        renamed = True
    try:
        # 5. project folder + app.db
        pdir = root / "projects" / rep.project_id
        pdir.mkdir(parents=True, exist_ok=False)
        for f in rep.old_dir.iterdir():
            if f.name in SKIP_FILES or f.name == "tracker.db":
                continue
            (shutil.copytree if f.is_dir() else shutil.copy2)(f, pdir / f.name)
        src = sqlite3.connect(rep.old_dir / "tracker.db")
        dst = sqlite3.connect(pdir / "tracker.db")
        try:
            src.backup(dst)
        finally:
            dst.close()
            src.close()
        appdb = open_app_db(base)
        try:
            ts = now_iso()
            with appdb.transaction():
                appdb.insert_project({"id": rep.project_id, "name": project_name, "excel_path": str(rep.new_excel),
                                      "created_at": ts, "created_by": editor})
                for d in holidays:
                    appdb.conn.execute("INSERT OR REPLACE INTO holiday (date, name, source, fetched_at) VALUES (?,?,?,?)", d)
                for f in fetches:
                    appdb.conn.execute("INSERT INTO holiday_fetch (ts, year, result, count, message) VALUES (?,?,?,?,?)", f)
                for k, v in shared_vals.items():
                    appdb.set_setting(k, v)
                appdb.log_event(ts=ts, project_id=rep.project_id, editor=editor, action="migrated",
                                detail=f"{rep.old_dir} -> {pdir}; {rep.old_excel.name} -> {rep.new_excel.name}")
            project = appdb.get_project(rep.project_id)
        finally:
            appdb.close()
        old_profile = rep.old_dir / "profile.json"
        prof = {}
        if old_profile.exists():
            try:
                prof = json.loads(old_profile.read_text(encoding="utf-8"))
            except ValueError:
                prof = {}
        if admin and not prof.get("email"):
            prof["email"] = admin[0]
        if admin and prof.get("email") == admin[0]:
            prof.update({"first_name": admin[1] or "", "last_name": admin[2] or ""})
        prof["last_project"] = rep.project_id
        new_profile = base.profile_path
        if not new_profile.exists():
            new_profile.parent.mkdir(parents=True, exist_ok=True)
            new_profile.write_text(json.dumps(prof, ensure_ascii=False), encoding="utf-8")

        # 6. counts after (through the new code path)
        cfg, repo = open_project(base, project)
        try:
            after = _counts(repo.conn)
            after["holiday"] = repo.conn.execute("SELECT COUNT(*) FROM shared.holiday").fetchone()[0]
            after["holiday_fetch"] = repo.conn.execute("SELECT COUNT(*) FROM shared.holiday_fetch").fetchone()[0]
            rep.after = after
            items = repo.load_items(TEST_PROGRESS)
            rep.lights_after = {i: compute(r, today, cfg.alerts.fail_threshold).color
                                for i, r in items.items() if r.get("active")}
        finally:
            repo.close()
        if not rep.same:
            raise ProjectError("Counts or Traffic Lights differ after the move")
    except Exception as exc:
        log.error("Migration failed, rolling back", exc_info=True)
        if renamed and rep.new_excel.exists() and not rep.old_excel.exists():
            rep.new_excel.rename(rep.old_excel)
        appdb = open_app_db(base)
        try:
            if appdb.get_project(rep.project_id):
                with appdb.transaction():
                    appdb.update_project(rep.project_id, archived=1, archived_at=now_iso())
                    appdb.log_event(ts=now_iso(), project_id=rep.project_id, editor=editor,
                                    action="migration_failed", detail=str(exc))
        finally:
            appdb.close()
        rep.message = f"FAILED, workbook name restored: {exc}"
        raise ProjectError(rep.message) from exc
    rep.done = True
    rep.message = "ok"
    return rep
