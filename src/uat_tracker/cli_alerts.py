"""Part 2 CLI commands: run-daily, lights, emails, jobs, holidays, upload, remind, scheduler."""

from __future__ import annotations

import datetime as dt
from collections import Counter
from pathlib import Path

from .alerts.outbox import mode_of
from .excel.sheets import TEST_PROGRESS
from .mail.sender import OutlookSender, PreviewOnlySender, SendError
from .parsing.dates import fmt
from .services import scheduler
from .services.attachments import UploadError, upload
from .services.auth import SETTINGS, PermissionDenied, require
from .services.daily_job import JOB, run_daily
from .services.holidays import parse_user_date, read_csv, refresh_year, store_holidays
from .services.reminders import ReminderError, send_reminder
from .services.settings import install_schedule
from .services.sync import SyncService
from .services.traffic_light import compute


def add_parsers(sub) -> None:
    p = sub.add_parser("run-daily", help="run the daily alert job now")
    p.add_argument("--date", help="pretend today is YYYY-MM-DD (testing; time is set to 08:30)")
    p.add_argument("--scheduled", action="store_true", help="started by Task Scheduler (applies the deadline)")
    p.add_argument("--force", action="store_true", help="run again although today already ran OK")
    p.add_argument("--preview-only", action="store_true",
                   help="dry-run without Outlook: write .html previews only (and no Excel write)")

    p = sub.add_parser("lights", help="show the Traffic Light for each Test Progress app (nothing saved)")
    p.add_argument("--date", help="as-of date YYYY-MM-DD (default today)")

    p = sub.add_parser("emails", help="email log")
    p.add_argument("--date", help="send date YYYY-MM-DD")
    p.add_argument("--item")
    p.add_argument("--limit", type=int, default=50)

    p = sub.add_parser("jobs", help="daily job runs")
    p.add_argument("--limit", type=int, default=20)

    h = sub.add_parser("holidays", help="bank holidays")
    hs = h.add_subparsers(dest="hcmd", required=True)
    x = hs.add_parser("list")
    x.add_argument("--year", type=int)
    x = hs.add_parser("fetch", help="fetch from the BOT API")
    x.add_argument("--year", type=int, action="append")
    x = hs.add_parser("add")
    x.add_argument("date", help="YYYY-MM-DD or DD-Mon-YYYY")
    x.add_argument("name")
    x = hs.add_parser("import", help="CSV: date,name")
    x.add_argument("csv")
    x = hs.add_parser("remove")
    x.add_argument("date")

    p = sub.add_parser("upload", help="attach a test-result file to a Test Progress item")
    p.add_argument("item")
    p.add_argument("file")

    p = sub.add_parser("remind", help="send a manual reminder for a UAT Plan / Defect List item")
    p.add_argument("item")
    p.add_argument("--to", required=True)
    p.add_argument("--cc", action="append", default=[])
    p.add_argument("--note")
    p.add_argument("--name", help="greeting name (default: Responsible / Assigned To)")
    p.add_argument("--preview-only", action="store_true")

    s = sub.add_parser("scheduler", help="Windows Task Scheduler entry")
    s.add_argument("action", choices=["install", "remove", "show", "print"])
    s.add_argument("--at", default=None, help="HH:MM (default: the email time from Settings, 08:30)")


def _sender(cfg, repo, preview_only: bool):
    if preview_only:
        if not cfg.alerts.dry_run:
            raise SystemExit("--preview-only works only in dry-run mode")
        return PreviewOnlySender()
    return OutlookSender(cfg.alerts.send_account)


def _clock(arg: str | None) -> dt.datetime:
    if not arg:
        return dt.datetime.now()
    return dt.datetime.combine(dt.date.fromisoformat(arg), dt.time(8, 30))


def handle(cfg, repo, args, config_path: Path) -> int | None:
    cmd = args.cmd
    if cmd == "run-daily":
        sender = _sender(cfg, repo, args.preview_only)
        sync = None if args.preview_only else SyncService(cfg, repo, sender=sender)
        res = run_daily(cfg, repo, sender, clock=_clock(args.date),
                        scheduled=args.scheduled, force=args.force, sync=sync)
        print(f"{res.run_date}  {res.result.upper()}  {res.message}")
        if res.missed_recorded:
            print(f"  recorded as missed (no backfill): {', '.join(fmt(d) for d in res.missed_recorded)}")
        if mode_of(cfg) == "dry_run":
            print(f"  DRY RUN: previews in {cfg.preview_dir / res.run_date}")
        return 0 if res.result in ("ok", "skipped") else 3

    if cmd == "lights":
        day = dt.date.fromisoformat(args.date) if args.date else dt.date.today()
        items = repo.load_items(TEST_PROGRESS)
        counts = Counter()
        for item_id, rec in sorted(items.items(), key=lambda kv: kv[1].get("excel_row") or 0):
            if not rec.get("active"):
                continue
            light = compute(rec, day, cfg.alerts.fail_threshold)
            counts[light.color] += 1
            pe = "-" if light.pct_exec is None else f"{light.pct_exec:6.1%}"
            pt = "-" if light.pct_time is None else f"{light.pct_time:6.1%}"
            print(f"{item_id:<8} {str(rec.get('application'))[:24]:<24} {light.color:<9} "
                  f"exec {pe:>6} time {pt:>7}  {light.reason}")
        print(f"As of {fmt(day)}, X = {cfg.alerts.fail_threshold:.0%}: "
              + ", ".join(f"{k} {v}" for k, v in counts.most_common()))
        return 0

    if cmd == "emails":
        rows = repo.email_log(send_date=args.date, item_id=args.item, limit=args.limit)
        for r in reversed(rows):
            print(f"{r['ts'][:16]} {r['rule']:<9} {(r['variant'] or ''):<6} {(r['item_id'] or '-'):<8} "
                  f"{r['mode']:<7} {r['result']:<8} to={r['to_addr'] or '-'}  {r['subject']}"
                  + (f"  [{r['error']}]" if r["error"] else ""))
        print(f"{len(rows)} row(s)")
        return 0

    if cmd == "jobs":
        for r in reversed(repo.job_runs(JOB, args.limit)):
            print(f"{r['run_date']} {r['result']:<8} {(r['mode'] or ''):<7} {r['message'] or ''}")
        return 0

    if cmd == "holidays":
        return _holidays(cfg, repo, args)

    if cmd == "upload":
        try:
            res = upload(cfg, repo, args.item, args.file, actor=args.actor)
        except (UploadError, PermissionDenied) as exc:
            print(f"UPLOAD FAILED: {exc}")
            return 2
        print(f"Saved {res.stored_path} (sha256 {res.sha256[:12]})")
        print("TC/Result Complete set to Complete" if res.status_changed
              else "TC/Result Complete was already Complete")
        return 0

    if cmd == "remind":
        try:
            r = send_reminder(cfg, repo, _sender(cfg, repo, args.preview_only), args.item, to=args.to,
                              cc=args.cc, note=args.note, recipient_name=args.name, actor=args.actor)
        except (ReminderError, SendError, PermissionDenied) as exc:
            print(f"REMINDER NOT SENT: {exc}")
            return 2
        print(f"Reminder {r} ({mode_of(cfg)})")
        return 0 if r in ("sent", "dry_run") else 2

    if cmd == "scheduler":
        from .config import email_time
        args.at = args.at or email_time(repo)
        pid, pname = cfg.project_id, (cfg.project_name or None)      # multi-project: one task pair per project
        if args.action == "print":
            import subprocess
            print(subprocess.list2cmdline(scheduler.install_args(config_path, args.at, pid, pname)))
            print(subprocess.list2cmdline(scheduler.install_weekly_args(config_path, project_id=pid, project_name=pname)))
            print(f"(/TR length {len(scheduler.task_command(config_path, pid))} of max 261)")
            return 0
        if args.action in ("install", "remove"):
            require(repo, args.actor, SETTINGS)
        procs = {"install": lambda: install_schedule(repo, args.actor, config_path, args.at, pid, pname),
                 "remove": lambda: [scheduler.remove(pname), scheduler.remove_weekly(pname)],
                 "show": lambda: [scheduler.show(pname)]}[args.action]()      # Part 6: the weekly task goes with it
        rc = 0
        for proc in procs:
            print(proc.stdout or proc.stderr)
            rc = rc or proc.returncode
        return rc
    return None


def _holidays(cfg, repo, args) -> int:
    now = dt.datetime.now().astimezone().isoformat(timespec="seconds")
    if args.hcmd != "list":
        require(repo, args.actor, SETTINGS)
        editor = args.actor.editor
    if args.hcmd == "list":
        rows = repo.holidays(args.year)
        for d, r in rows.items():
            print(f"{fmt(d)}  {dt.date.fromisoformat(d):%a}  {r['source']:<6} {r['name']}")
        print(f"{len(rows)} holiday(s); years cached: {sorted(repo.holiday_years()) or 'none'}")
        for f in repo.holiday_fetches(5):
            print(f"  fetch {f['ts'][:16]} {f['year']} {f['result']}: {f['message']}")
        return 0
    if args.hcmd == "fetch":
        today = dt.date.today()
        rc = 0
        for y in args.year or [today.year, today.year + 1]:
            fr = refresh_year(repo, cfg, y, now=now, editor=editor)
            print(f"{y}: {'OK' if fr.ok else 'FAILED'} {fr.message}")
            rc = rc or (0 if fr.ok else 2)
        return rc
    if args.hcmd == "add":
        d = parse_user_date(args.date)
        with repo.transaction():
            changed = store_holidays(repo, [(d.isoformat(), args.name)], "manual", now, editor)
        print(f"{fmt(d)} {'added' if changed else 'already there'}")
        return 0
    if args.hcmd == "import":
        rows = read_csv(Path(args.csv))
        with repo.transaction():
            changed = store_holidays(repo, rows, "csv", now, editor)
        print(f"{len(rows)} row(s) read, {changed} new/changed")
        return 0
    if args.hcmd == "remove":
        d = parse_user_date(args.date).isoformat()
        with repo.transaction():
            old = repo.delete_holiday(d)
            if old:
                repo.append_audit({"ts": now, "editor": editor, "channel": "app", "sheet": "Holidays",
                                   "item_id": d, "field": "holiday", "old_value": old["name"],
                                   "new_value": "", "reason": "Holiday removed by PMO"})
        print(f"{fmt(d)} {'removed' if old else 'not found'}")
        return 0
    return 1
