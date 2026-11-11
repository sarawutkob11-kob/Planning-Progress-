"""Weekly report (Part 6, approved 10-Nov-2026): build -> (AI summary) -> HTML -> preview / send -> keep.

Rules:
- Schedule: Windows task 09:00 Mon-Fri; the job sends on Monday, or on the first workday of the week when
  Monday is a bank holiday. Once per week. PC/Outlook off -> skipped + logged, never backfilled.
  No holiday data for the year -> nothing sent, PMO warned (as Part 2).
- Re-import Excel first. Numbers and tables are rendered by code; the AI writes only the summary part.
- AI unavailable / refused / numbers not in the data -> the report goes out WITHOUT the AI part, with a note,
  and the PMO is warned.
- Recipients: every Report By Email + CC Email in Test Progress, deduped; invalid ones skipped + flagged.
  None at all -> the PMO only, with a warning. Dry-run -> PMO only (outbox rules).
- Manual: Dashboard "Generate Weekly Report" -> preview -> Send / Cancel (Admin, PMO, User BU/PM).
- Every report is kept in weekly_report (content, recipients, AI on/off, tokens, cost, status).
"""

from __future__ import annotations

import logging
import datetime as dt
import json
import re
from dataclasses import dataclass

from ..alerts.outbox import DRY, Outgoing, alert_pmo, deliver, mode_of
from ..alerts.templates import _e, footer
from ..domain.models import CHANNEL_APP, CHANNEL_SYSTEM
from ..excel.sheets import TEST_PROGRESS
from ..mail.sender import SendError
from . import weekly_ai
from .auth import WEEKLY, Actor, require, system_actor
from .holidays import refresh_year
from .importer import run_import
from .weekly_data import build_data
from .workdays import WorkCalendar

log = logging.getLogger(__name__)

JOB = "weekly_report"
EDITOR = "system (weekly report)"
SYSTEM_ACTOR = system_actor(EDITOR)
RULE = "WEEKLY"
FLAG_BAD_EMAIL = "invalid_report_email"
_EMAIL = re.compile(r"^[^@\s;<>]+@[^@\s;<>]+\.[^@\s;<>]+$")


def _now(clock: dt.datetime | None = None) -> str:
    return (clock or dt.datetime.now()).astimezone().isoformat(timespec="seconds")


# --------------------------------------------------------------------------- recipients

def recipients(repo, *, flag: bool = False, now: str | None = None) -> tuple[list[str], list[dict]]:
    """Deduped valid addresses (Report By Email + CC Email) and the invalid ones (skipped, flagged)."""
    good: dict[str, None] = {}
    bad: list[dict] = []
    items = [r for r in repo.load_items(TEST_PROGRESS).values() if r.get("active")]
    items.sort(key=lambda r: r.get("excel_row") or 10 ** 6)
    for r in items:
        for fld in ("report_by_email", "cc_emails"):
            for part in re.split(r"[;,\s]+", r.get(fld) or ""):
                a = part.strip().lower()
                if not a:
                    continue
                if _EMAIL.match(a):
                    good.setdefault(a, None)
                else:
                    bad.append({"id": r["id"], "app": r.get("application"), "field": fld, "value": part})
    if flag and bad:
        with repo.transaction():
            for b in bad:
                repo.upsert_flag(TEST_PROGRESS, b["id"], FLAG_BAD_EMAIL,
                                 f"Weekly report: invalid address in {b['field']}: {b['value']!r}", now or _now())
    return list(good), bad


# --------------------------------------------------------------------------- HTML (code renders every number)

_TH_COLOR = {"Red": "แดง", "Amber": "เหลือง", "Green": "เขียว", "Complete": "เสร็จ", "Missing": "ข้อมูลไม่ครบ"}
_KIND = {"overdue": ("เลยกำหนด", "Overdue"), "fail_rate": ("อัตรา fail", "Fail rate"), "time": ("ช้ากว่าเวลา", "Behind time")}
_STYLE_CELL = "padding:4px 8px;border-bottom:1px solid #E3E7EE;vertical-align:top"


def _h(th: str, en: str) -> str:
    return (f'<h3 style="margin:18px 0 6px 0;color:#1F4FBF">{_e(th)} <span style="color:#666;font-weight:normal">'
            f'/ {_e(en)}</span></h3>')


def _table(headers: list[str], rows: list[list]) -> str:
    if not rows:
        return '<p style="color:#666">ไม่มีรายการ / None</p>'
    head = "".join(f'<th style="{_STYLE_CELL};text-align:left;background:#F4F6F9">{_e(h)}</th>' for h in headers)
    body = "".join("<tr>" + "".join(f'<td style="{_STYLE_CELL}">{_e("" if v is None else v)}</td>' for v in r)
                   + "</tr>" for r in rows)
    return f'<table style="border-collapse:collapse;margin:4px 0 8px 0">{head and "<tr>" + head + "</tr>"}{body}</table>'


def _ai_html(ai: dict | None, note: str) -> str:
    if not ai:
        return (f'<div style="background:#FFF1D6;color:#7A4E00;padding:8px 12px;margin:8px 0">'
                f'ไม่มีสรุปจาก AI ในรายงานนี้ / No AI summary in this report: {_e(note)}</div>')
    def items(key, a, b):
        li = "".join(f"<li><b>{_e(i.get(f'{a}_th'))}</b> — {_e(i.get(f'{b}_th'))}<br>"
                     f'<span style="color:#444">{_e(i.get(f"{a}_en"))} — {_e(i.get(f"{b}_en"))}</span></li>'
                     for i in ai.get(key, []))
        return f"<ul>{li}</ul>" if li else '<p style="color:#666">-</p>'
    return (_h("สรุปสำหรับผู้บริหาร (ร่างโดย AI)", "Executive summary (drafted by AI)")
            + f"<p>{_e(ai.get('summary_th'))}</p><p style=\"color:#444\">{_e(ai.get('summary_en'))}</p>"
            + _h("ความเสี่ยงหลัก", "Top risks") + items("risks", "title", "why")
            + _h("สิ่งที่ควรติดตาม", "Suggested follow-ups") + items("follow_ups", "action", "why"))


def render(data: dict, ai: dict | None, ai_note: str = "", extra_notes: list[str] = ()) -> tuple[str, str]:
    o = data["overall"]
    subject = f"[Planning Progress] รายงานประจำสัปดาห์ {data['project']} | Weekly Report - {data['as_of']}"
    c = o["colors"]
    notes = "".join(f'<div style="background:#FDE7E3;color:#A3260F;padding:8px 12px;margin:8px 0">{_e(n)}</div>'
                    for n in extra_notes)
    ra = [[a["app"], f"{_TH_COLOR.get(a['color'], a['color'])} ({a['color']})",
           f"{_KIND[a['reason_kind']][0]} / {_KIND[a['reason_kind']][1]}", a["reason"],
           f"{a['executed']}/{a['total_tc']}"] for a in data["red_amber"]]
    od = ([["UAT Plan", t["id"], t["task"], t["finish_date"], t["days_late"]] for t in data["overdue"]["uat_plan"]]
          + [["Test Progress", t["id"], t["app"], t["planned_finish"], t["days_late"]]
             for t in data["overdue"]["test_progress"]])
    d = data["defects"]
    dov = [[x["id"], x["system"], x["detail"], x["fix_target_date"], x["days_late"]] for x in d["overdue"]]
    dno = [[x["id"], x["system"], x["detail"], x["status"]] for x in d["no_target"]]
    due = [[x["date"], x["page"], x["id"], x["name"]] for x in data["due_next_2_weeks"]]
    rs = [[x["date"], x["page"], x["name"], f"{x['from']} → {x['to']}",
           x["reason"] if x["where"] == "app" else "ไม่ระบุเหตุผล (แก้ใน Excel) / Not given (edited in Excel)"]
          for x in data["reschedules_last_7_days"]]
    w = data["reschedule_window"]
    html = f"""<html><body style="font-family:'IBM Plex Sans Thai',Tahoma,sans-serif;font-size:14px;color:#14213D">
<h2 style="margin:0 0 4px 0">รายงาน UAT ประจำสัปดาห์ {_e(data['project'])} / Weekly UAT Report</h2>
<p style="color:#555;margin:0 0 8px 0">ข้อมูล ณ / Data as of {_e(data['as_of'])}</p>
{notes}{_ai_html(ai, ai_note)}
{_h("ภาพรวม", "Overview")}
{_table(["", ""], [
    ["ทดสอบแล้ว / Executed", f"{o['executed']} / {o['total_tc']} ({o['pct_executed']})"],
    ["Pass", o["pass"]], ["Fail", o["fail"]],
    ["แอป / Apps", f"{o['apps']}: เสร็จ/Complete {c.get('Complete', 0)} · แดง/Red {c.get('Red', 0)} · "
                   f"เหลือง/Amber {c.get('Amber', 0)} · เขียว/Green {c.get('Green', 0)} · "
                   f"ข้อมูลไม่ครบ/Missing {c.get('Missing', 0)}"]])}
{_h("แอปสีแดง / เหลือง", "Red / Amber apps")}
{_table(["แอป / App", "สี / Color", "สาเหตุ / Cause", "รายละเอียด / Detail", "Executed/Total"], ra)}
{_h("งานเลยกำหนด", "Overdue tasks")}
{_table(["หน้า / Page", "ID", "ชื่อ / Name", "กำหนด / Due", "เลยมา (วัน) / Days late"], od)}
{_h("Defect", "Defects")}
<p>เปิดอยู่ / Open: {d['open']} · เลยวันเป้าหมาย / Past fix target: {len(d['overdue'])} · ไม่มีวันเป้าหมาย / No fix target: {len(d['no_target'])}</p>
{_table(["ID", "System", "รายละเอียด / Detail", "วันเป้าหมาย / Fix target", "เลยมา (วัน) / Days late"], dov) if dov else ""}
{_table(["ID", "System", "รายละเอียด / Detail", "สถานะ / Status"], dno) if dno else ""}
{_h("ครบกำหนดใน 2 สัปดาห์", "Due in the next 2 weeks")}
{_table(["วันที่ / Date", "หน้า / Page", "ID", "ชื่อ / Name"], due)}
{_h(f"การเลื่อนวัน {w['from']} – {w['to']}", f"Reschedules {w['from']} – {w['to']}")}
{_table(["วันที่ / Date", "หน้า / Page", "ชื่อ / Name", "จาก → เป็น / From → To", "เหตุผล / Reason"], rs)}
{_h("ข้อขัดแย้งและคิวตรวจ", "Conflicts and review queue")}
<p>ข้อขัดแย้งที่ยังเปิด / Open conflicts: {data['open_conflicts']} · คิวตรวจของ PMO / PMO review queue: {data['review_queue']}</p>
<p style="color:#777;font-size:12px">ตัวเลขและตารางทั้งหมดคำนวณโดยระบบ{' ส่วนสรุปร่างโดย AI (Claude)' if ai else ''} ·
All numbers and tables are computed by the tracker{'; the summary is drafted by AI (Claude)' if ai else ''}.<br>{_e(footer())}</p>
</body></html>"""
    return subject, html


# --------------------------------------------------------------------------- generate / send

@dataclass
class Generated:
    report_id: int
    ai_ok: bool
    ai_error: str
    recipients: list[str]
    invalid: list[dict]


def _summarize(cfg):
    """weekly_ai.summarize with this config's [weekly_ai] fallback switch and prices."""
    import functools
    w = getattr(cfg, "weekly_ai", None)
    if w is None:
        return weekly_ai.summarize
    return functools.partial(weekly_ai.summarize, server_fallback=w.server_fallback, prices=w.prices)


def generate(cfg, repo, *, actor: Actor, as_of: dt.date, trigger: str = "manual", use_ai: bool = True,
             ai_func=None, job_run_id: int | None = None, now: str | None = None) -> Generated:
    """Build the report (status 'preview'). Nothing is sent."""
    require(repo, actor, WEEKLY)
    now = now or _now()
    data = build_data(cfg, repo, as_of)
    to, bad = recipients(repo, flag=True, now=now)
    ai = weekly_ai.AiResult(False, error="AI turned off for this run") if not use_ai else \
        (ai_func or _summarize(cfg))(data, model=weekly_ai.model_of(repo, cfg))
    notes = []
    if not to:
        notes.append("ไม่มีผู้รับใน Report By Email / CC Email: ส่งให้ PMO เท่านั้น · "
                     "No Report By Email / CC Email addresses: sent to the PMO only.")
    if bad:
        notes.append(f"ข้ามอีเมลที่ไม่ถูกต้อง {len(bad)} รายการ · Skipped {len(bad)} invalid address(es): "
                     + ", ".join(f"{b['app']} ({b['value']})" for b in bad[:10]))
    subject, html = render(data, ai.content if ai.ok else None, ai.error, notes)
    with repo.transaction():
        rid = repo.insert_weekly({
            "created_at": now, "as_of": as_of.isoformat(), "trigger": trigger, "created_by": actor.editor,
            "status": "preview", "recipients": json.dumps(to), "invalid_recipients": json.dumps(bad, ensure_ascii=False),
            "subject": subject, "html": html, "data_json": json.dumps(data, ensure_ascii=False),
            "ai_used": 1 if ai.ok else 0, "ai_model": ai.model or None, "ai_error": None if ai.ok else ai.error,
            "input_tokens": ai.input_tokens or None, "output_tokens": ai.output_tokens or None,
            "cost_usd": ai.cost_usd, "job_run_id": job_run_id})
        repo.append_audit({"ts": now, "editor": actor.editor, "channel": CHANNEL_SYSTEM if actor.is_system else CHANNEL_APP,
                           "sheet": "Weekly report", "item_id": f"WR-{rid}", "field": "generated",
                           "old_value": "", "new_value": f"AI {'on' if ai.ok else 'off'}; {len(to)} recipient(s)",
                           "reason": (f"{ai.model}: {ai.input_tokens} in / {ai.output_tokens} out tokens, "
                                      f"${ai.cost_usd:.4f}" if ai.cost_usd is not None else
                                      (ai.error or "no AI call"))})
    return Generated(rid, ai.ok, ai.error, to, bad)


def send(cfg, repo, sender, report_id: int, *, actor: Actor, now: str | None = None) -> str:
    """Send a previewed report. Dry-run -> PMO only. Returns sent | dry_run | failed | skipped."""
    require(repo, actor, WEEKLY)
    now = now or _now()
    r = repo.get_weekly(report_id)
    if r is None or r["status"] != "preview":
        raise ValueError(f"Weekly report #{report_id} is not waiting to be sent")
    to = json.loads(r["recipients"] or "[]") or ([cfg.alerts.pmo_email] if cfg.alerts.pmo_email else [])
    out = Outgoing(RULE, r["trigger"], None, f"WR-{report_id}", None, "; ".join(to) or None, [], r["subject"], r["html"])
    result = deliver(cfg, repo, sender, out, send_date=now[:10], now=now, editor=actor.editor,
                     channel=CHANNEL_SYSTEM if actor.is_system else CHANNEL_APP, job_run_id=r["job_run_id"])
    with repo.transaction():
        repo.update_weekly(report_id, status=result, mode=mode_of(cfg), sent_at=now, sent_by=actor.editor)
    return result


def cancel(repo, report_id: int, *, actor: Actor) -> None:
    require(repo, actor, WEEKLY)
    with repo.transaction():
        repo.update_weekly(report_id, status="cancelled", message=f"Cancelled by {actor.editor}")


# --------------------------------------------------------------------------- scheduled job

@dataclass
class WeeklyResult:
    run_date: str
    result: str                 # ok | skipped | failed
    message: str
    report_id: int | None = None
    send_result: str | None = None


def is_report_day(cal: WorkCalendar, day: dt.date) -> bool:
    """Monday, or the first workday of the week when Monday (and later days) are bank holidays."""
    if day.weekday() > 4 or not cal.is_workday(day):
        return False
    monday = day - dt.timedelta(days=day.weekday())
    return all(not cal.is_workday(monday + dt.timedelta(days=i)) for i in range(day.weekday()))


def _sent_this_week(repo, day: dt.date) -> bool:
    monday = day - dt.timedelta(days=day.weekday())
    return any(r["result"] == "ok" and monday.isoformat() <= r["run_date"] <= day.isoformat()
               for r in repo.job_runs(JOB, 20))


def run_weekly(cfg, repo, sender, *, clock: dt.datetime | None = None, scheduled: bool = True,
               force: bool = False, importer=run_import, ai_func=None, fetcher=None,
               use_ai: bool = True) -> WeeklyResult:
    clock = clock or dt.datetime.now()
    day = clock.date()
    iso, now, mode = day.isoformat(), _now(clock), mode_of(cfg)

    def record(result, message, report_id=None, send_result=None):
        with repo.transaction():
            repo.record_job(job=JOB, run_date=iso, ts=now, result=result, mode=mode, message=message)
        return WeeklyResult(iso, result, message, report_id, send_result)

    if not cal_years_ok(repo, cfg, day, now, fetcher):
        msg = (f"No bank-holiday data for {day.year}: cannot tell workdays, so the weekly report was not sent. "
               f"Add holidays (Settings or 'holidays fetch / add / import').")
        res = record("failed", msg)
        alert_pmo(cfg, repo, sender, "No holiday data - weekly report not sent", msg, send_date=iso, now=now,
                  editor=EDITOR)
        return res
    cal = WorkCalendar.from_repo(repo)
    if scheduled and not force:
        if _sent_this_week(repo, day):
            return WeeklyResult(iso, "skipped", "Weekly report already sent this week")
        if not is_report_day(cal, day):
            return WeeklyResult(iso, "skipped", "Not the weekly report day (Monday, or the first workday after a "
                                                "Monday holiday)")
    if sender.delivers and (mode != DRY or cfg.alerts.pmo_email):
        try:
            sender.check()
        except SendError as exc:
            return record("skipped", f"{exc}; weekly report skipped (not sent later)")
    try:
        importer(cfg, repo, editor=EDITOR)
    except Exception as exc:
        log.error("Weekly job: re-import of the Excel file failed", exc_info=True)
        msg = f"Re-import of the Excel file failed, so the weekly report was not sent: {exc}"
        res = record("failed", msg)
        alert_pmo(cfg, repo, sender, "Excel re-import failed - weekly report not sent", msg, send_date=iso,
                  now=now, editor=EDITOR)
        return res
    g = generate(cfg, repo, actor=SYSTEM_ACTOR, as_of=day, trigger="scheduled", use_ai=use_ai, ai_func=ai_func,
                 now=now)
    result = send(cfg, repo, sender, g.report_id, actor=SYSTEM_ACTOR, now=now)
    if use_ai and not g.ai_ok:
        alert_pmo(cfg, repo, sender, "Weekly report sent without the AI summary",
                  f"Report WR-{g.report_id} was sent without the AI part: {g.ai_error}", send_date=iso, now=now,
                  editor=EDITOR)
    if not g.recipients:
        alert_pmo(cfg, repo, sender, "Weekly report: no recipients",
                  "Test Progress has no Report By Email / CC Email addresses, so the weekly report went to the "
                  "PMO only.", send_date=iso, now=now, editor=EDITOR)
    ok = result in ("sent", "dry_run")
    return record("ok" if ok else "failed", f"WR-{g.report_id}: {result}", g.report_id, result)


def cal_years_ok(repo, cfg, day: dt.date, now: str, fetcher=None) -> bool:
    if day.year in repo.holiday_years():
        return True
    refresh_year(repo, cfg, day.year, now=now, editor=EDITOR, fetcher=fetcher)
    return day.year in repo.holiday_years()
