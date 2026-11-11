"""Projects (multi-project add-on, approved 10-Nov-2026).

- 1 project = 1 Excel file = 1 folder <root>/projects/<id>/ with its own tracker.db (items, users + roles,
  per-project settings, logs, attachments, backups). app.db (shared) holds the project list, bank holidays,
  PMO email, language and the Claude model; the Claude API key stays in Windows Credential Manager.
- Create: Project Name only, folder picked every time, file "Planning Progress <name>.xlsx" (characters
  / \\ : * ? " < > | become -), made from the template in code. The display name keeps the original text.
- Import: an existing .xlsx with the 3 sheets and their required headers. The file is never renamed and other
  sheets are never touched. Name = file name without .xlsx and without a leading "Planning Progress ".
- Duplicates (create + import): same final file name already in the chosen folder, or a project (also an
  archived one) already uses that file name -> blocked.
- Who may create / import / relink / archive and change shared settings: Admin or PMO in any project; anyone
  while there is no project yet. The creator becomes the new project's Admin.
- Archive only (hidden, data kept, can be undone). No delete.
- Relink: the new file needs the 3 sheets + required headers, and >= 80% of the project's App IDs.
"""

from __future__ import annotations

import dataclasses
import re
from datetime import date
from pathlib import Path

from ..config import Config, apply_overrides
from ..excel.sheets import SHEETS, TEST_PROGRESS, with_overrides
from ..storage.app_db import AppDb
from .auth import ADMIN, PMO, QA, name_key, now_iso

FILE_PREFIX = "Planning Progress "
INVALID_CHARS = re.compile(r'[/\\:*?"<>|]')
RELINK_MIN_MATCH = 0.8
MANAGER_ROLES = (ADMIN, PMO)


class ProjectError(Exception):
    pass


@dataclasses.dataclass(frozen=True)
class Person:
    """Who uses the app on this PC (profile form)."""
    email: str
    first_name: str = ""
    last_name: str = ""

    @property
    def full_name(self) -> str:
        return f"{self.first_name} {self.last_name}".strip()

    @property
    def editor(self) -> str:
        return f"{self.full_name} <{self.email}>" if self.full_name else self.email


# --------------------------------------------------------------------------- names

def file_name_for(project_name: str) -> str:
    """'A/B' -> 'Planning Progress A-B.xlsx' (approved)."""
    name = (project_name or "").strip()
    if not name:
        raise ProjectError("Project Name is required")
    safe = INVALID_CHARS.sub("-", name)
    safe = "".join(ch for ch in safe if ord(ch) >= 32).rstrip(" .")    # Windows: no control chars, no trailing . or space
    if not safe:
        raise ProjectError("Project Name has no usable characters for a file name")
    return f"{FILE_PREFIX}{safe}.xlsx"


def name_from_file(path: Path | str) -> str:
    stem = Path(path).stem
    if stem.casefold().startswith(FILE_PREFIX.casefold()) and len(stem) > len(FILE_PREFIX):
        stem = stem[len(FILE_PREFIX):]
    return stem.strip()


def duplicate_reason(appdb: AppDb, folder: Path, file_name: str, *, ignore_id: str | None = None) -> str | None:
    """None = OK. Same file name in the folder, or a project already using that file name (approved)."""
    if (Path(folder) / file_name).exists():
        return f"'{file_name}' already exists in {folder}"
    for p in appdb.projects(include_archived=True):
        if p["id"] != ignore_id and Path(p["excel_path"]).name.casefold() == file_name.casefold():
            return f"A project already uses the file name '{file_name}' (project '{p['name']}'" + \
                   (", archived)" if p["archived"] else ")")
    return None


# --------------------------------------------------------------------------- config / repository per project

def project_config(base: Config, project: dict) -> Config:
    root = base.root_dir
    ws = date.fromisoformat(project["window_start"]) if project.get("window_start") else base.window_start
    we = date.fromisoformat(project["window_end"]) if project.get("window_end") else base.window_end
    return dataclasses.replace(base, excel_path=Path(project["excel_path"]), data_dir=root / "projects" / project["id"],
                               app_dir=root, shared_db=root / "app.db", project_id=project["id"],
                               project_name=project["name"], window_start=ws, window_end=we)


def open_project(base: Config, project: dict):
    """-> (cfg with Settings-screen overrides, repository) for one project."""
    from ..alerts.templates import set_project
    from ..storage import open_repository
    cfg = project_config(base, project)
    repo = open_repository(cfg)
    set_project(project["name"])                      # email footer / PMO alert subject name the project
    return apply_overrides(cfg, repo), repo


def open_app_db(base: Config) -> AppDb:
    return AppDb(base.root_dir / "app.db")


def _with_repo(base, project, fn):
    """Look into one project WITHOUT making it the current one: open_project() would also change the project
    name in email footers / PMO alert subjects (bug: looping over all projects left the last one's name)."""
    from ..storage import open_repository
    cfg = project_config(base, project)
    repo = open_repository(cfg)
    cfg = apply_overrides(cfg, repo)
    try:
        return fn(cfg, repo)
    finally:
        repo.close()


# --------------------------------------------------------------------------- who sees / manages what

def role_in(base: Config, project: dict, email: str) -> str | None:
    def go(_cfg, repo):
        u = repo.get_user((email or "").lower())
        return u["role"] if u and u["active"] else None
    return _with_repo(base, project, go)


def projects_for(base: Config, appdb: AppDb, email: str) -> list[dict]:
    """Projects shown in the drop-down: not archived and the user is an active member (approved Q3)."""
    return [p for p in appdb.projects() if role_in(base, p, email)]


def can_manage(base: Config, appdb: AppDb, email: str) -> bool:
    all_p = appdb.projects(include_archived=True)
    if not all_p:
        return True                                   # first use: nobody to ask
    return any(role_in(base, p, email) in MANAGER_ROLES for p in all_p)


def require_manager(base: Config, appdb: AppDb, person: Person) -> None:
    if not can_manage(base, appdb, person.email):
        raise ProjectError("Not allowed: only an Admin or PMO (of any project) can do this")


def find_person(base: Config, appdb: AppDb, email: str) -> Person | None:
    """Stored first / last name of this email in any project (stored names win over typed ones)."""
    email = (email or "").strip().lower()
    for p in appdb.projects(include_archived=True):
        u = _with_repo(base, p, lambda _c, r: r.get_user(email))
        if u and (u.get("first_name") or u.get("last_name")):
            return Person(email, u.get("first_name") or "", u.get("last_name") or "")
    return None


def is_deactivated_everywhere(base: Config, appdb: AppDb, email: str) -> bool:
    found = [u for p in appdb.projects(include_archived=True)
             if (u := _with_repo(base, p, lambda _c, r: r.get_user(email.lower()))) is not None]
    return bool(found) and not any(u["active"] for u in found)


def _emails_in_tp(repo) -> set[str]:
    out = set()
    for r in repo.load_items(TEST_PROGRESS).values():
        if not r.get("active"):
            continue
        for v in (r.get("report_by_email"), r.get("cc_emails")):
            out.update(e.strip().lower() for e in re.split(r"[;,\s]+", v or "") if "@" in e)
    return out


def auto_join(base: Config, appdb: AppDb, person: Person) -> list[str]:
    """Approved Q4: a new email becomes QA in every project whose Test Progress lists it (Report By Email /
    CC Email). Returns the project ids joined."""
    from .auth import audit_auth
    joined = []
    for p in appdb.projects():
        def go(_cfg, repo):
            if repo.get_user(person.email) is not None or person.email not in _emails_in_tp(repo):
                return False
            ts = now_iso()
            with repo.transaction():
                repo.insert_user(email=person.email, first_name=person.first_name or None,
                                 last_name=person.last_name or None, role=QA, created_at=ts, created_by=person.editor)
                if person.full_name:
                    repo.set_user_aliases(person.email, {name_key(person.full_name): person.full_name})
                audit_auth(repo, editor=person.editor, field="profile_created", item_id=person.email, new=QA,
                           reason="Email found in Test Progress (Report By Email / CC Email): role QA", now=ts)
            return True
        if _with_repo(base, p, go):
            joined.append(p["id"])
    return joined


# --------------------------------------------------------------------------- validation of an existing file

def validate_workbook(path: Path, header_overrides: dict | None = None) -> list[str]:
    """[] = OK, else what is missing (sheets / required headers). Read-only, on a copy."""
    import tempfile

    import openpyxl

    from ..excel.reader import copy_shared
    from ..parsing.headers import HeaderError, map_headers
    path = Path(path)
    if not path.exists():
        return [f"File not found: {path}"]
    if path.suffix.lower() != ".xlsx":
        return ["Only .xlsx files can be used"]
    problems = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        copy = Path(tmp) / "check.xlsx"
        copy_shared(path, copy)
        try:
            wb = openpyxl.load_workbook(copy, read_only=True, data_only=True)
        except Exception as exc:
            return [f"Not a readable Excel file: {exc}"]
        try:
            for name, sd in SHEETS.items():
                sd = with_overrides(sd, header_overrides or {})
                if name not in wb.sheetnames:
                    problems.append(f"Sheet '{name}' is missing")
                    continue
                header = next(wb[name].iter_rows(min_row=1, max_row=1, values_only=True), ())
                try:
                    map_headers(list(header), list(sd.fields), name, sd.blank_first_as)
                except HeaderError as exc:
                    problems.append(str(exc))
        finally:
            wb.close()
    return problems


def app_ids_in(path: Path) -> set[str]:
    import tempfile

    import openpyxl

    from ..excel.reader import copy_shared
    from ..excel.sheets import APP_ID_HEADER
    from ..parsing.headers import header_label
    ids = set()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        copy = Path(tmp) / "ids.xlsx"
        copy_shared(path, copy)
        wb = openpyxl.load_workbook(copy, data_only=True)
        try:
            for name in SHEETS:
                if name not in wb.sheetnames:
                    continue
                rows = wb[name].iter_rows(values_only=True)
                header = [header_label(v).casefold() for v in next(rows, ())]
                if APP_ID_HEADER.casefold() not in header:
                    continue
                c = header.index(APP_ID_HEADER.casefold())
                ids.update(str(r[c]).strip() for r in rows if c < len(r) and r[c] not in (None, ""))
        finally:
            wb.close()
    return ids


# --------------------------------------------------------------------------- create / import / relink / archive

def _new_project(base: Config, appdb: AppDb, person: Person, name: str, excel_path: Path, action: str,
                 detail: str) -> dict:
    ts = now_iso()
    with appdb.transaction():
        pid = appdb.next_project_id()
        appdb.insert_project({"id": pid, "name": name, "excel_path": str(excel_path), "created_at": ts,
                              "created_by": person.editor})
        appdb.log_event(ts=ts, project_id=pid, editor=person.editor, action=action, detail=detail)
    project = appdb.get_project(pid)

    def admin(_cfg, repo):                              # creator = Admin of the new project (approved Q5)
        from .auth import audit_auth
        with repo.transaction():
            repo.insert_user(email=person.email, first_name=person.first_name or None,
                             last_name=person.last_name or None, role=ADMIN, created_at=ts, created_by=person.editor)
            if person.full_name:
                repo.set_user_aliases(person.email, {name_key(person.full_name): person.full_name})
            audit_auth(repo, editor=person.editor, field="profile_created", item_id=person.email, new=ADMIN,
                       reason=f"{action}: creator is the project Admin", now=ts)
    _with_repo(base, project, admin)
    return project


def preview_create(appdb: AppDb, name: str, folder: Path) -> tuple[str, str | None]:
    """For the form: (final file name, problem or None)."""
    fname = file_name_for(name)
    return fname, duplicate_reason(appdb, Path(folder), fname)


def create_project(base: Config, appdb: AppDb, person: Person, *, name: str, folder: Path) -> dict:
    from .project_template import create_workbook
    require_manager(base, appdb, person)
    name = name.strip()
    fname = file_name_for(name)
    folder = Path(folder)
    if not folder.is_dir():
        raise ProjectError(f"Folder not found: {folder}")
    why = duplicate_reason(appdb, folder, fname)
    if why:
        raise ProjectError(why)
    path = create_workbook(folder / fname)
    return _new_project(base, appdb, person, name, path, "created", f"new file {path}")


def import_project(base: Config, appdb: AppDb, person: Person, *, path: Path) -> dict:
    """Existing workbook as a new project (never renamed). The first read of the file happens right after."""
    require_manager(base, appdb, person)
    path = Path(path)
    problems = validate_workbook(path, base.header_overrides)
    if problems:
        raise ProjectError("This file can't be imported:\n- " + "\n- ".join(problems))
    for p in appdb.projects(include_archived=True):     # same rule as create: the file name
        if Path(p["excel_path"]).name.casefold() == path.name.casefold():
            raise ProjectError(f"A project already uses the file name '{path.name}' (project '{p['name']}')")
    name = name_from_file(path)
    if not name:
        raise ProjectError("The file name gives no project name")
    return _new_project(base, appdb, person, name, path, "imported", f"existing file {path}")


def relink(base: Config, appdb: AppDb, person: Person, project_id: str, new_path: Path) -> str:
    """Approved Q8. Returns a short summary."""
    require_manager(base, appdb, person)
    project = appdb.get_project(project_id)
    if project is None:
        raise ProjectError(f"Unknown project {project_id}")
    new_path = Path(new_path)
    problems = validate_workbook(new_path, base.header_overrides)
    if problems:
        raise ProjectError("This file can't be linked:\n- " + "\n- ".join(problems))
    for p in appdb.projects(include_archived=True):
        if p["id"] != project_id and Path(p["excel_path"]).name.casefold() == new_path.name.casefold():
            raise ProjectError(f"Project '{p['name']}' already uses the file name '{new_path.name}'")

    def known(_cfg, repo):
        return {i for s in SHEETS for i, r in repo.load_items(s).items() if r.get("active") and r.get("excel_row")}
    expected = _with_repo(base, project, known)
    found = app_ids_in(new_path)
    if expected:
        share = len(expected & found) / len(expected)
        if share < RELINK_MIN_MATCH:
            raise ProjectError(f"Only {share:.0%} of this project's App IDs are in that file (need "
                               f"{RELINK_MIN_MATCH:.0%}). It looks like a different file.")
        summary = f"{len(expected & found)} of {len(expected)} App IDs found"
    else:
        summary = "project has no items yet"
    ts = now_iso()
    with appdb.transaction():
        appdb.update_project(project_id, excel_path=str(new_path))
        appdb.log_event(ts=ts, project_id=project_id, editor=person.editor, action="relinked",
                        detail=f"{project['excel_path']} -> {new_path} ({summary})")
    return summary


def set_archived(base: Config, appdb: AppDb, person: Person, project_id: str, archived: bool) -> None:
    require_manager(base, appdb, person)
    ts = now_iso()
    with appdb.transaction():
        appdb.update_project(project_id, archived=1 if archived else 0, archived_at=ts if archived else None)
        appdb.log_event(ts=ts, project_id=project_id, editor=person.editor,
                        action="archived" if archived else "unarchived")


def set_window(base: Config, appdb: AppDb, person: Person, project_id: str, start: date, end: date) -> None:
    """Per-project date window (approved Q15)."""
    if start > end:
        raise ProjectError("Window start is after window end")
    project = appdb.get_project(project_id)
    if role_in(base, project, person.email) not in MANAGER_ROLES + ("User BU/PM",):
        raise ProjectError("Not allowed: you can't change this project's settings")
    ts = now_iso()
    with appdb.transaction():
        appdb.update_project(project_id, window_start=start.isoformat(), window_end=end.isoformat())
        appdb.log_event(ts=ts, project_id=project_id, editor=person.editor, action="window",
                        detail=f"{start} .. {end}")


def file_missing(project: dict) -> bool:
    return not Path(project["excel_path"]).exists()


# --------------------------------------------------------------------------- profile across projects

def current_person(base: Config) -> Person | None:
    from .profile import read_profile_data
    d = read_profile_data(base.profile_path)
    return Person(d["email"], d.get("first_name") or "", d.get("last_name") or "") if d.get("email") else None


def identify_person(base: Config, appdb: AppDb, *, first_name: str, last_name: str, email: str) -> Person:
    """Profile form across projects (Part 5 v2 rules kept): an email that already exists keeps its stored
    name; a deactivated email can't be used; a new email joins the projects that list it (QA, approved Q4).
    Saves the profile file. Projects are then opened with that email's role in each project."""
    from .profile import ProfileError, norm_email, valid_email, write_profile
    email = norm_email(email)
    if not valid_email(email):
        raise ProfileError("Type one valid email address")
    if is_deactivated_everywhere(base, appdb, email):
        raise ProfileError(f"{email} is deactivated: ask the Admin")
    stored = find_person(base, appdb, email)
    first, last = (stored.first_name, stored.last_name) if stored else (first_name.strip(), last_name.strip())
    if not first or not last:
        raise ProfileError("First name and last name are required")
    person = Person(email, first, last)
    auto_join(base, appdb, person)
    write_profile(base.profile_path, email, first_name=first, last_name=last)
    return person


def remember_project(base: Config, project_id: str) -> None:
    from .profile import write_profile
    write_profile(base.profile_path, None, last_project=project_id)
