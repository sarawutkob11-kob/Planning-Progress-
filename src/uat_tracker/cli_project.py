"""Multi-project CLI: project list|create|import|relink|archive|unarchive|migrate|restore-clear.

Project commands run as the profile saved on this PC by the app (create / import / relink / archive need an
Admin or PMO of any project; anyone while there is no project yet)."""

from __future__ import annotations

from pathlib import Path

from .services import projects
from .services.profile import read_profile_data


def add_parsers(sub) -> None:
    p = sub.add_parser("project", help="projects: list, create, import, relink, archive, migrate, restore-clear")
    ps = p.add_subparsers(dest="pcmd", required=True)
    ps.add_parser("list", help="all projects (also archived)")
    c = ps.add_parser("create", help="new project: new Excel file from the template")
    c.add_argument("name")
    c.add_argument("--folder", required=True, help="folder for 'Planning Progress <name>.xlsx'")
    i = ps.add_parser("import", help="existing Excel file as a new project (never renamed)")
    i.add_argument("path")
    r = ps.add_parser("relink", help="point a project to its moved / renamed Excel file")
    r.add_argument("project")
    r.add_argument("path")
    for name in ("archive", "unarchive"):
        a = ps.add_parser(name)
        a.add_argument("project")
    m = ps.add_parser("migrate", help="one-time: old single-project data -> project 'RAMC UAT'")
    m.add_argument("--old-dir", default=None, help="default: %%LOCALAPPDATA%%\\UATTracker")
    m.add_argument("--name", default="RAMC UAT")
    m.add_argument("--yes", action="store_true", help="really do it (default: dry run, nothing changes)")
    rc = ps.add_parser("restore-clear", help="put back a 'Clear project data' backup zip (app closed)")
    rc.add_argument("project")
    rc.add_argument("zip")


def person(base) -> projects.Person | None:
    d = read_profile_data(base.profile_path)
    if not d.get("email"):
        return None
    return projects.Person(d["email"], d.get("first_name") or "", d.get("last_name") or "")


def handle(base, args) -> int:
    import os
    appdb = projects.open_app_db(base)
    try:
        if args.pcmd == "list":
            rows = appdb.projects(include_archived=True)
            for p in rows:
                miss = "  [FILE NOT FOUND]" if projects.file_missing(p) else ""
                print(f"{p['id']}  {p['name']:<30} {'(archived) ' if p['archived'] else ''}{p['excel_path']}{miss}")
            print(f"{len(rows)} project(s)")
            return 0
        if args.pcmd == "migrate":
            from .services.migrate_multi import migrate, plan
            old = Path(args.old_dir) if args.old_dir else Path(os.path.expandvars("%LOCALAPPDATA%")) / "UATTracker"
            rep = plan(base, old, project_name=args.name)
            print(f"Old data folder : {rep.old_dir}")
            print(f"New data folder : {rep.new_root / 'projects' / rep.project_id}")
            print(f"Excel now       : {rep.old_excel}")
            print(f"Excel after     : {rep.new_excel}")
            migrate(base, old, project_name=args.name, confirm=False)        # checks only
            if not args.yes:
                print("Dry run: nothing changed. Run again with --yes to do it.")
                return 0
            rep = migrate(base, old, project_name=args.name, confirm=True)
            print("Backups:\n  " + "\n  ".join(rep.backups))
            print(f"{'table':<24}{'before':>8}{'after':>8}")
            for t in sorted(set(rep.before) | set(rep.after)):
                b, a = rep.before.get(t), rep.after.get(t)
                print(f"{t:<24}{b!s:>8}{a!s:>8}{'' if a == b else '   <-- DIFFERENT'}")
            print(f"Traffic Light: {len(rep.lights_before)} apps before, {len(rep.lights_after)} after, "
                  f"{'identical' if rep.lights_before == rep.lights_after else 'DIFFERENT'}")
            print(f"Result: {rep.message}")
            return 0 if rep.done else 2
        who = person(base)
        if who is None:
            print("No user profile on this PC. Open the app once and fill in your name and email.")
            return 3
        if args.pcmd == "create":
            p = projects.create_project(base, appdb, who, name=args.name, folder=Path(args.folder))
            print(f"Created {p['id']} '{p['name']}': {p['excel_path']}")
            return 0
        if args.pcmd == "import":
            p = projects.import_project(base, appdb, who, path=Path(args.path))
            from .services.importer import run_import
            cfg, repo = projects.open_project(base, p)
            try:
                res = run_import(cfg, repo, editor=who.editor)
                print(f"Imported {p['id']} '{p['name']}': " +
                      ", ".join(f"{s} {st.rows} rows" for s, st in res.sheets.items()))
            finally:
                repo.close()
            return 0
        proj = appdb.find_project(args.project)
        if proj is None:
            print(f"Unknown project: {args.project}")
            return 2
        if args.pcmd == "relink":
            print(projects.relink(base, appdb, who, proj["id"], Path(args.path)))
            return 0
        if args.pcmd in ("archive", "unarchive"):
            projects.set_archived(base, appdb, who, proj["id"], args.pcmd == "archive")
            print(f"{proj['name']}: {args.pcmd}d")
            return 0
        if args.pcmd == "restore-clear":
            from .services.clear_project import restore_backup
            cfg = projects.project_config(base, proj)
            man = restore_backup(cfg, Path(args.zip), editor=who.editor)
            print(f"Restored backup of {man['created_at']} into {cfg.db_path}")
            return 0
    except Exception as exc:
        from .services.clear_project import ClearError
        if not isinstance(exc, (projects.ProjectError, ClearError)):
            raise
        print(f"NOT DONE: {exc}")
        return 2
    finally:
        appdb.close()
    return 1
