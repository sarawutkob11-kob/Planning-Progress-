"""Part 3 CLI: sync, conflicts, review resolve, edit, backups."""

from __future__ import annotations

import threading

from .excel.sheets import SHEETS
from .mail.sender import OutlookSender
from .parsing.dates import fmt
from .services.auth import PermissionDenied
from .services.edits import ROOT_CAUSES, EditError, edit_item
from .services.resolve import ResolveError, describe_conflict, resolve_conflict, resolve_review
from .services.sync import SyncBusy, SyncService

SHEET_ALIASES = {"uat": "UAT Plan", "plan": "UAT Plan", "tp": "Test Progress", "test": "Test Progress",
                 "defect": "Defect List", "defects": "Defect List"}


def add_parsers(sub) -> None:
    s = sub.add_parser("sync", help="2-way sync with the Excel file")
    ss = s.add_subparsers(dest="scmd", required=True)
    ss.add_parser("status")
    ss.add_parser("once", help="read Excel, then write queued app changes now")
    ss.add_parser("watch", help="keep syncing until Ctrl+C")
    ss.add_parser("init-ids", help="add the hidden App ID column / fill missing IDs")
    ss.add_parser("unblock", help="resume writes after a OneDrive conflict copy was dealt with")
    x = ss.add_parser("log")
    x.add_argument("--limit", type=int, default=20)

    c = sub.add_parser("conflicts", help="Conflict Center")
    cs = c.add_subparsers(dest="ccmd", required=True)
    cs.add_parser("list")
    x = cs.add_parser("resolve")
    x.add_argument("id", type=int)
    x.add_argument("choice", choices=["excel", "app", "new"])
    x.add_argument("--value", help="new value (choice 'new'); dates as YYYY-MM-DD")
    x.add_argument("--reason", required=True)

    r = sub.add_parser("resolve-review", help="resolve a PMO review item (see 'review')")
    r.add_argument("id", type=int)
    r.add_argument("value", nargs="?", default="",
                   help="date YYYY-MM-DD / status / number / time text; empty = drop that part")
    r.add_argument("--reason", default="")

    e = sub.add_parser("edit", help="change one field (same path the UI will use)")
    e.add_argument("sheet", help="uat | tp | defect (or the sheet name)")
    e.add_argument("item")
    e.add_argument("field")
    e.add_argument("value", help="dates as YYYY-MM-DD; '' to clear")
    e.add_argument("--reason")
    e.add_argument("--root-cause", choices=ROOT_CAUSES)

    b = sub.add_parser("backup", help="workbook backups")
    bs = b.add_subparsers(dest="bcmd", required=True)
    x = bs.add_parser("list")
    x.add_argument("--limit", type=int, default=20)
    bs.add_parser("now")
    x = bs.add_parser("restore")
    x.add_argument("id", type=int)
    x.add_argument("--yes", action="store_true", help="confirm replacing the Excel file")


def _svc(cfg, repo):
    return SyncService(cfg, repo, sender=OutlookSender(cfg.alerts.send_account))


def handle(cfg, repo, args) -> int | None:
    actor = args.actor
    if args.cmd == "sync":
        svc = _svc(cfg, repo)
        if args.scmd == "status":
            st = svc.status()
            print(f"state={st['state']}  pending={st['pending']}  failed={st['failed']}  "
                  f"conflicts={st['conflicts']}  last read={st['last_read'] or '-'}")
            if st["blocked"]:
                print(f"BLOCKED: {st['blocked']}")
            if st["last_error"]:
                print(f"last error: {st['last_error']}")
            for p in repo.all_pending():
                print(f"  {p['status']:<8} {p['sheet']:<14} {p['item_id']:<8} {p['field']:<20} "
                      f"{p['value']!r}  tries={p['attempts']}" + (f"  [{p['last_error']}]" if p["last_error"] else ""))
            return 0
        if args.scmd == "once":
            res = svc.pull()
            print("read Excel: " + (svc.repo.sync_events(1)[0]["message"] if res else "no change"))
            fr = svc.flush(force=True)
            print(f"write: {fr.result} {fr.message}" + (f"  backup: {fr.backup}" if fr.backup else ""))
            for f in fr.failed:
                print(f"  FAILED {f}")
            return 0 if fr.result in ("nothing", "written") else 2
        if args.scmd == "watch":
            stop = threading.Event()
            try:
                svc.watch(stop)
            except KeyboardInterrupt:
                stop.set()
            return 0
        if args.scmd == "init-ids":
            from .services.auth import SETTINGS, require
            require(repo, actor, SETTINGS)
            fr = svc.init_ids()
            print(f"{fr.result}: {fr.message}" + (f"  backup: {fr.backup}" if fr.backup else ""))
            return 0 if fr.result == "written" else 2
        if args.scmd == "unblock":
            from .services.auth import SETTINGS, require
            require(repo, actor, SETTINGS)
            svc.unblock()
            print("Writes resumed")
            return 0
        if args.scmd == "log":
            for e in reversed(repo.sync_events(args.limit)):
                print(f"{e['ts'][:19]} {e['kind']:<8} {e['result']:<8} {e['message'] or ''}")
            return 0

    if args.cmd == "conflicts":
        if args.ccmd == "list":
            rows = repo.open_conflicts()
            for c in rows:
                print(describe_conflict(c) + f"  (Excel {c['excel_time'] or '-'} / app {c['app_time'] or '-'})")
            print(f"{len(rows)} open conflict(s)")
            return 0
        try:
            resolve_conflict(cfg, repo, args.id, args.choice, actor=actor, reason=args.reason,
                             new_value=args.value)
        except (ResolveError, PermissionDenied) as exc:
            print(f"NOT RESOLVED: {exc}")
            return 2
        _svc(cfg, repo).pull(force=True)
        print(f"Conflict #{args.id} resolved ({args.choice}). Run 'sync once' (or keep 'sync watch' running) to write Excel.")
        return 0

    if args.cmd == "resolve-review":
        try:
            v = resolve_review(cfg, repo, args.id, args.value or None, actor=actor, reason=args.reason)
        except (ResolveError, PermissionDenied) as exc:
            print(f"NOT RESOLVED: {exc}")
            return 2
        print(f"Review #{args.id} resolved -> {fmt(v) if len(v) == 10 and v[4] == '-' else v or '(dropped)'}; "
              "queued for Excel")
        return 0

    if args.cmd == "edit":
        sheet = SHEET_ALIASES.get(args.sheet.lower(), args.sheet)
        if sheet not in SHEETS:
            print(f"Unknown sheet {args.sheet!r}")
            return 2
        try:
            changed = edit_item(cfg, repo, sheet, args.item, {args.field: args.value}, actor=actor,
                                reason=args.reason, root_cause=args.root_cause)
        except (EditError, PermissionDenied) as exc:
            print(f"NOT SAVED: {exc}")
            return 2
        print(f"Saved {', '.join(changed)}; queued for Excel" if changed else "No change")
        return 0

    if args.cmd == "backup":
        svc = _svc(cfg, repo)
        if args.bcmd == "list":
            for b in reversed(repo.backups(args.limit)):
                print(f"#{b['id']:<4} {b['ts'][:19]} {b['size_bytes'] // 1024:>6} KB  {b['reason'] or ''}  {b['path']}")
            return 0
        if args.bcmd == "now":
            print(svc.backup("manual backup"))
            return 0
        if args.bcmd == "restore":
            if not args.yes:
                print("This replaces the Excel file with the backup (the current file is backed up first). "
                      "Add --yes to confirm.")
                return 2
            try:
                prev = svc.restore(args.id, actor)
            except (ValueError, SyncBusy, PermissionDenied) as exc:
                print(f"NOT RESTORED: {exc}")
                return 2
            print(f"Backup #{args.id} restored. Previous file saved as {prev}")
            return 0
    return None
