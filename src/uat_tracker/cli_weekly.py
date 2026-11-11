"""Part 6 CLI: run-weekly (the scheduled job), weekly list / show, ai-key set / clear."""

from __future__ import annotations

import datetime as dt
import getpass
from pathlib import Path

from .mail.sender import OutlookSender, PreviewOnlySender
from .services import weekly_ai, weekly_report
from .services.auth import PermissionDenied
from .services.settings import save_ai_settings


def add_parsers(sub) -> None:
    p = sub.add_parser("run-weekly", help="generate and send the weekly report now")
    p.add_argument("--scheduled", action="store_true", help="run by Task Scheduler: only on the report day")
    p.add_argument("--force", action="store_true", help="send even if not the report day / already sent")
    p.add_argument("--date", help="pretend it is this day (YYYY-MM-DD, testing)")
    p.add_argument("--preview-only", action="store_true", help="dry-run without Outlook: preview file only")
    p.add_argument("--no-ai", action="store_true", help="skip the Claude summary")
    w = sub.add_parser("weekly", help="past weekly reports")
    ws = w.add_subparsers(dest="wcmd", required=True)
    x = ws.add_parser("list")
    x.add_argument("--limit", type=int, default=20)
    x = ws.add_parser("show", help="write a past report to an .html file")
    x.add_argument("id", type=int)
    x.add_argument("--out", help="file to write (default: WR-<id>.html here)")
    k = sub.add_parser("ai-key", help="Claude API key for the weekly report (Windows Credential Manager)")
    ks = k.add_subparsers(dest="kcmd", required=True)
    x = ks.add_parser("set", help="type the key (hidden)")
    x.add_argument("--model", help=f"also set the model (default {weekly_ai.DEFAULT_MODEL})")
    ks.add_parser("clear")
    ks.add_parser("status")


def handle(cfg, repo, args) -> int | None:
    if args.cmd == "run-weekly":
        clock = dt.datetime.fromisoformat(args.date + "T09:00") if args.date else dt.datetime.now()
        if args.preview_only and not cfg.alerts.dry_run:
            print("--preview-only works only in dry-run mode")
            return 2
        sender = PreviewOnlySender() if args.preview_only else OutlookSender(cfg.alerts.send_account)
        res = weekly_report.run_weekly(cfg, repo, sender, clock=clock, scheduled=args.scheduled, force=args.force,
                                       use_ai=not args.no_ai)
        print(f"{res.run_date}: {res.result} - {res.message}")
        if res.report_id:
            r = repo.get_weekly(res.report_id)
            print(f"  AI: {'yes (' + str(r['ai_model']) + ')' if r['ai_used'] else 'no - ' + str(r['ai_error'])}"
                  + (f"  tokens {r['input_tokens']}/{r['output_tokens']}  ${r['cost_usd']:.4f}"
                     if r["cost_usd"] is not None else ""))
        return 0 if res.result in ("ok", "skipped") else 2

    if args.cmd == "weekly":
        if args.wcmd == "list":
            for r in reversed(repo.weekly_reports(args.limit)):
                print(f"WR-{r['id']:<4} {r['created_at'][:16]} as of {r['as_of']} {r['trigger']:<9} {r['status']:<9} "
                      f"AI {'yes' if r['ai_used'] else 'no '} "
                      + (f"${r['cost_usd']:.4f}" if r["cost_usd"] is not None else "") + f"  {r['created_by']}")
            return 0
        r = repo.get_weekly(args.id)
        if r is None:
            print(f"No weekly report #{args.id}")
            return 2
        out = Path(args.out or f"WR-{args.id}.html")
        out.write_text(r["html"], encoding="utf-8")
        print(f"Written {out.resolve()}")
        return 0

    if args.cmd == "ai-key":
        if args.kcmd == "status":
            print(f"Model: {weekly_ai.model_of(repo)}  |  API key saved: {'yes' if weekly_ai.KEYS.get() else 'no'}")
            return 0
        try:
            if args.kcmd == "set":
                key = getpass.getpass("Claude API key (hidden): ").strip()
                if not key:
                    print("Nothing typed: key not changed")
                    return 2
                save_ai_settings(repo, args.actor, model=args.model or weekly_ai.model_of(repo), api_key=key)
                print("Saved in Windows Credential Manager")
            else:
                save_ai_settings(repo, args.actor, model=weekly_ai.model_of(repo), clear_key=True)
                print("Key removed")
        except PermissionDenied as exc:
            print(f"NOT ALLOWED: {exc}")
            return 2
        return 0
    return None
