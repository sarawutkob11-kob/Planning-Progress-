"""Daily 08:30 job (PART2 §F): workday check -> re-import -> colors -> R1-R4 -> send -> logs.

Survives restarts: all state is in the DB (job_run, traffic_light_history, email_log).
- A day already run OK is not run again (unless forced; email dedupe still applies).
- Missed workdays since the last run are recorded as 'skipped' (PC off) - never backfilled.
- A scheduled run starting after [alerts] run_deadline (09:30) is skipped and logged.
- Holidays: Monday run refreshes this year + next; the current year is also fetched when it has no
  data. Fetch failure -> keep cache + warn PMO. No data for the current year -> send nothing, alert PMO.
- Outlook not running -> skip the day (log). Re-import fails -> send nothing, alert PMO.
- Switching dry-run -> live: colors already recorded count as notified; only new changes email (6c).
"""

from __future__ import annotations

import logging
import datetime as dt
from dataclasses import dataclass, field

from ..alerts.outbox import DRY, LIVE, Outgoing, alert_pmo, deliver, mode_of
from ..alerts.rules import R3, R3Event, evaluate
from ..alerts.templates import render
from ..domain.models import CHANNEL_SYSTEM
from ..excel.sheets import TEST_PROGRESS
from ..mail.sender import SendError
from .edits import EditError, edit_item
from .auth import system_actor
from .holidays import refresh_year
from .importer import run_import
from .traffic_light import MISSING, effective_previous, r3_qualifies, record_colors
from .workdays import WorkCalendar

log = logging.getLogger(__name__)

JOB = "daily_alerts"
EDITOR = "system (daily job)"
SYSTEM_ACTOR = system_actor(EDITOR)     # Part 5: the scheduled job needs no login
MAX_MISSED_LOOKBACK = 31


@dataclass
class JobResult:
    run_date: str
    result: str                     # ok | skipped | failed
    message: str
    job_id: int | None = None
    emails: dict[str, int] = field(default_factory=dict)
    color_changes: int = 0
    lights_queued: int = 0
    missed_recorded: list[str] = field(default_factory=list)


def _now(clock: dt.datetime) -> str:
    return clock.astimezone().isoformat(timespec="seconds")


def run_daily(cfg, repo, sender, *, clock: dt.datetime | None = None, scheduled: bool = False,
              force: bool = False, fetcher=None, importer=run_import, sync=None) -> JobResult:
    clock = clock or dt.datetime.now()
    today = clock.date()
    day = today.isoformat()
    now = _now(clock)
    mode = mode_of(cfg)

    def finish_simple(result, message) -> JobResult:
        with repo.transaction():
            jid = repo.record_job(job=JOB, run_date=day, ts=now, result=result, mode=mode, message=message)
        return JobResult(day, result, message, jid)

    if repo.job_ok_on(JOB, day) and not force:
        return JobResult(day, "skipped", "Already ran OK today (use --force to run again)")

    if scheduled and clock.time() > cfg.alerts.run_deadline:
        return finish_simple("skipped", f"Started late ({clock:%H:%M}, deadline "
                                        f"{cfg.alerts.run_deadline:%H:%M}); no late sends")

    # ---- holidays
    warnings = []
    years = [today.year, today.year + 1] if today.weekday() == 0 else []
    if today.year not in repo.holiday_years() and today.year not in years:
        years.insert(0, today.year)
    for y in years:
        fr = refresh_year(repo, cfg, y, now=now, editor=EDITOR, fetcher=fetcher)
        if not fr.ok:
            warnings.append(f"BOT holiday fetch for {y} failed: {fr.message}. Cached holidays kept.")
    cal = WorkCalendar.from_repo(repo)
    for w in warnings:
        alert_pmo(cfg, repo, sender, "Holiday fetch failed", w, send_date=day, now=now, editor=EDITOR)
    if not cal.has_year(today.year):
        msg = (f"No bank-holiday data for {today.year}: cannot tell workdays, so no emails were sent. "
               f"Add holidays (holidays fetch / add / import) and run again.")
        res = finish_simple("failed", msg)
        alert_pmo(cfg, repo, sender, "No holiday data - alerts not sent", msg, send_date=day, now=now,
                  editor=EDITOR, job_run_id=res.job_id)
        return res

    missed = _record_missed_days(repo, cal, today, now, mode)

    if not cal.is_workday(today):
        res = finish_simple("skipped", "Not a workday (weekend or bank holiday)")
        res.missed_recorded = missed
        return res

    # ---- Outlook must be running (dry-run with a pmo_email sends through it too)
    if sender.delivers and (mode == LIVE or cfg.alerts.pmo_email):
        try:
            sender.check()
        except SendError as exc:
            res = finish_simple("skipped", f"{exc}; day skipped")
            res.missed_recorded = missed
            return res

    with repo.transaction():
        jid = repo.start_job(job=JOB, run_date=day, started_at=now, mode=mode)
    res = JobResult(day, "running", "", jid, missed_recorded=missed)

    # ---- re-import
    try:
        importer(cfg, repo, editor=EDITOR)
    except Exception as exc:                       # file locked, header changed, ...
        log.error("Daily job: re-import of the Excel file failed", exc_info=True)
        msg = f"Re-import of the Excel file failed, so no emails were sent today: {exc}"
        with repo.transaction():
            repo.finish_job(jid, finished_at=_now(dt.datetime.now()), result="failed", message=msg)
        alert_pmo(cfg, repo, sender, "Excel re-import failed - alerts not sent", msg, send_date=day,
                  now=now, editor=EDITOR, job_run_id=jid)
        res.result, res.message = "failed", msg
        return res

    # ---- colors
    items = repo.load_items(TEST_PROGRESS)
    with repo.transaction():
        lights, changes = record_colors(repo, items, today, cfg.alerts.fail_threshold, now=now,
                                        editor=EDITOR, job_run_id=jid)
    res.color_changes = len(changes)
    res.lights_queued = _queue_traffic_lights(cfg, repo, items, lights, now)
    live_since = _mode_switch(repo, mode, now)

    # ---- rules -> emails
    uploads = {a["item_id"] for a in repo.attachments(TEST_PROGRESS)}
    counts: dict[str, int] = {}
    for item_id, rec in sorted(items.items()):
        if not rec.get("active"):
            continue
        event = _pending_r3(repo, item_id, mode, live_since)
        for alert in evaluate(rec, today, cal, r3_event=event, has_upload=item_id in uploads):
            subject, html = render(alert, rec, lights.get(item_id))
            cc = [c.strip() for c in (rec.get("cc_emails") or "").split(";") if c.strip()]
            out = Outgoing(alert.rule, alert.variant, TEST_PROGRESS, item_id, rec.get("application"),
                           rec.get("report_by_email"), cc, subject, html,
                           event_ref=alert.event.history_id if alert.event else None)
            r = deliver(cfg, repo, sender, out, send_date=day, now=now, editor=EDITOR, job_run_id=jid)
            counts[r] = counts.get(r, 0) + 1
    res.emails = counts
    summary = (f"mode={mode}; colors changed={len(changes)}; emails "
               + (", ".join(f"{k}={v}" for k, v in sorted(counts.items())) or "none"))
    if sync is not None:                          # write queued changes (Traffic Light...) to Excel
        try:
            fr = sync.flush(force=True)
            summary += f"; Excel write: {fr.result}" + (f" ({fr.message})" if fr.message else "")
        except Exception as exc:                  # emails already went out; report, don't fail the day
            log.error("Daily job: Excel write after the emails failed", exc_info=True)
            summary += f"; Excel write error: {exc}"
    with repo.transaction():
        repo.finish_job(jid, finished_at=_now(dt.datetime.now()), result="ok", message=summary)
    res.result, res.message = "ok", summary
    return res


def _queue_traffic_lights(cfg, repo, items, lights, now) -> int:
    """Computed color -> Test Progress 'Traffic Light' column (approved: auto, not typed).
    No color (missing data) leaves the cell as it is."""
    n = 0
    for item_id, light in lights.items():
        if light.color == MISSING:
            continue
        rec = items.get(item_id) or {}
        if (rec.get("traffic_light") or "") == light.color:
            continue
        try:
            if edit_item(cfg, repo, TEST_PROGRESS, item_id, {"traffic_light": light.color}, actor=SYSTEM_ACTOR,
                         reason=f"Calculated: {light.reason}", channel=CHANNEL_SYSTEM, now=now):
                n += 1
        except EditError:                      # e.g. cell in an open conflict: the PMO decides first
            continue
    return n


def _record_missed_days(repo, cal: WorkCalendar, today: dt.date, now: str, mode: str) -> list[str]:
    dates = sorted(repo.job_dates(JOB))
    if not dates:
        return []
    last = dt.date.fromisoformat(dates[-1])
    if last >= today:
        return []
    start = max(last, today - dt.timedelta(days=MAX_MISSED_LOOKBACK))
    missed = [d.isoformat() for d in cal.workdays_between(start, today)]
    if missed:
        with repo.transaction():
            for d in missed:
                repo.record_job(job=JOB, run_date=d, ts=now, result="skipped", mode=mode,
                                message="Not run (PC off or not logged on); no backfill")
    return missed


def _mode_switch(repo, mode: str, now: str) -> str | None:
    """Remember when live mode started. Color changes recorded before then count as notified."""
    with repo.transaction():
        last = repo.get_setting("alerts.last_mode")
        if mode == LIVE and last != LIVE:
            repo.set_setting("alerts.live_since", now)
        if last != mode:
            repo.set_setting("alerts.last_mode", mode)
    return repo.get_setting("alerts.live_since") if mode == LIVE else None


def _pending_r3(repo, item_id: str, mode: str, live_since: str | None) -> R3Event | None:
    """The item's current color entered Amber/Red and that change has not been emailed yet."""
    history = repo.light_history(item_id)
    if not history:
        return None
    cur = history[-1]
    if cur["color"] == MISSING:
        return None
    prev = effective_previous(history)
    if not r3_qualifies(cur["color"], prev):  # also covers Amber -> Missing -> Amber (no re-email)
        return None
    if mode == LIVE and live_since and cur["ts"] < live_since:
        return None                            # changed while in dry-run: treated as notified
    if repo.event_notified(R3, cur["id"], mode):
        return None
    return R3Event(cur["id"], cur["color"], prev, cur["reason"] or "")
