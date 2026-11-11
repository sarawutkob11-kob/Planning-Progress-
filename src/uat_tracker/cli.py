"""Command line: python -m uat_tracker <command>

  import    copy the source workbook, import the 3 sheets, print a summary
  summary   item counts, open review items and active flags by type
  review    list open PMO review items
  flags     list active data flags

Part 2 (alerts): run-daily, lights, emails, jobs, holidays, upload, remind, scheduler.
Part 3 (sync):   sync status|once|watch|init-ids|unblock|log, conflicts, resolve-review, edit, backup.
Part 4 (app):    ui [--no-sync], ui-shortcut.
Part 6 (report): run-weekly, weekly list|show, ai-key set|clear|status.
Part 5: commands that change data run as the user profile saved on this PC by the app.
Multi-project: project list|create|import|relink|archive|unarchive|migrate|restore-clear.
Every other command works on ONE project: --project <id or name> (default: the project last opened in the
app on this PC, or the only project).
Use "<command> -h" for options.
"""

from __future__ import annotations

import argparse
import sys
from collections import Counter

from pathlib import Path

from . import cli_alerts, cli_auth, cli_project, cli_sync, cli_weekly
from .config import DEFAULT_CONFIG_NAME, load_config
from .services.projects import open_app_db, open_project
from .parsing.headers import HeaderError
from .services.auth import PermissionDenied
from .services.importer import run_import


def _print_summary(repo) -> None:
    counts = repo.count_items()
    print("\nItems (active / inactive):")
    for sheet, c in counts.items():
        print(f"  {sheet:<14} {c['active']:>4} / {c['inactive']}")
    reviews = Counter((r["sheet"], r["problem"]) for r in repo.open_reviews())
    print(f"\nPMO review queue (open): {sum(reviews.values())}")
    for (sheet, problem), n in sorted(reviews.items()):
        print(f"  {sheet:<14} {problem:<22} {n:>4}")
    flags = Counter((f["sheet"], f["code"]) for f in repo.active_flags())
    print(f"\nData flags (active): {sum(flags.values())}")
    for (sheet, code), n in sorted(flags.items()):
        print(f"  {sheet:<14} {code:<26} {n:>4}")
    print(f"\nAudit log entries: {repo.count_audit()}")


def cmd_import(cfg, repo, args) -> int:
    try:
        res = run_import(cfg, repo, source=args.source, force=True)   # manual import also ends a pause
    except (HeaderError, KeyError, FileNotFoundError) as exc:
        print(f"IMPORT FAILED: {exc}", file=sys.stderr)
        return 2
    print(f"Import run #{res.run_id}  snapshot: {res.snapshot}")
    print(f"{'Sheet':<14} {'rows':>5} {'new':>5} {'upd':>5} {'same':>5} {'gone':>5} "
          f"{'audit':>6} {'review+':>7} {'flag+':>6} {'flag-':>6}")
    for sheet, s in res.sheets.items():
        print(f"{sheet:<14} {s.rows:>5} {s.created:>5} {s.updated:>5} {s.unchanged:>5} {s.removed:>5} "
              f"{s.audit_added:>6} {s.reviews_added:>7} {s.flags_added:>6} {s.flags_cleared:>6}")
        if s.missing_optional_columns:
            print(f"  note: optional column(s) not in sheet: {', '.join(s.missing_optional_columns)}")
    _print_summary(repo)
    return 0


def cmd_review(cfg, repo, args) -> int:
    rows = repo.open_reviews()
    for r in rows:
        raw = (r["raw_value"] or "").replace("\n", " | ")
        print(f"#{r['id']:<4} {r['sheet']:<14} row {r['excel_row']:<4} {r['item_id']:<9} "
              f"{r['field']:<24} {r['problem']:<20} token={r['token']!r}  cell={raw!r}")
    print(f"{len(rows)} open review item(s)")
    return 0


def cmd_flags(cfg, repo, args) -> int:
    rows = repo.active_flags()
    for f in rows:
        print(f"{f['sheet']:<14} {f['item_id']:<9} {f['code']:<26} {f['detail']}")
    print(f"{len(rows)} active flag(s)")
    return 0


def _pick_project(base, key: str | None):
    from .services.profile import read_profile_data
    appdb = open_app_db(base)
    try:
        if key:
            p = appdb.find_project(key)
            if p is None:
                print(f"Unknown project: {key}. See: project list", file=sys.stderr)
            return p
        active = appdb.projects()
        last = read_profile_data(base.profile_path).get("last_project")
        p = next((x for x in active if x["id"] == last), None)
        if p is None and len(active) == 1:
            p = active[0]
        if p is None:
            print("No project chosen: use --project <id or name> (see: project list)" if active else
                  "No project yet: create or import one in the app (or: project create / import / migrate)",
                  file=sys.stderr)
        return p
    finally:
        appdb.close()


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except AttributeError:
            pass
    p = argparse.ArgumentParser(prog="uat_tracker", description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--config", help="path to config.toml (default: ./config.toml)")
    p.add_argument("--project", help="project id or name (default: last opened in the app, or the only one)")
    sub = p.add_subparsers(dest="cmd", required=True)
    imp = sub.add_parser("import")
    imp.add_argument("--source", help="override the workbook path from config")
    sub.add_parser("summary")
    sub.add_parser("review")
    sub.add_parser("flags")
    cli_alerts.add_parsers(sub)
    cli_sync.add_parsers(sub)
    cli_weekly.add_parsers(sub)
    ui = sub.add_parser("ui", help="open the desktop app")
    ui.add_argument("--no-sync", action="store_true", help="do not watch / write the Excel file")
    sub.add_parser("ui-shortcut", help="create the 'Planning Progress' shortcut on the Desktop")
    cli_project.add_parsers(sub)
    args = p.parse_args(argv)

    base = load_config(args.config)
    config_path = Path(args.config) if args.config else Path.cwd() / DEFAULT_CONFIG_NAME
    if args.cmd == "project":
        return cli_project.handle(base, args)
    if args.cmd == "ui":
        from .ui.app import run
        return run(base, config_path, start_sync=not args.no_sync)
    if args.cmd == "ui-shortcut":
        from .ui.app import create_shortcut
        print(f"Shortcut created: {create_shortcut(config_path)}")
        return 0
    project = _pick_project(base, args.project)
    if project is None:
        return 2
    cfg, repo = open_project(base, project)   # values changed in the app's Settings screen applied
    print(f"Project: {project['name']} ({project['id']})", file=sys.stderr)
    args.actor = None
    try:
        if cli_auth.needs_login(args):
            args.actor = cli_auth.cli_identity(cfg, repo)
            if args.actor is None:
                return 3
        if args.cmd == "import":
            return cmd_import(cfg, repo, args)
        if args.cmd == "summary":
            _print_summary(repo)
            return 0
        if args.cmd == "review":
            return cmd_review(cfg, repo, args)
        if args.cmd == "flags":
            return cmd_flags(cfg, repo, args)
        rc = cli_weekly.handle(cfg, repo, args)
        if rc is None:
            rc = cli_alerts.handle(cfg, repo, args, config_path)
        if rc is None:
            rc = cli_sync.handle(cfg, repo, args)
        if rc is not None:
            return rc
    except PermissionDenied as exc:
        print(f"NOT ALLOWED: {exc}")
        return 2
    finally:
        repo.close()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
