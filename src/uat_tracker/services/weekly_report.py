"""Weekly report (Part 6, approved 10-Nov-2026): build -> (AI summary) -> HTML -> preview / send -> keep.

Rules:
- Schedule: Windows task 09:00 Mon-Fri; the job sends on Monday, or on the first workday of the week when
  Monday is a bank holiday. Once per week. PC/Outlook off -> skipped + logged, never backfilled.
  No holiday data for the year -> nothing sent, PMO warned (as Part 2).
- Re-import Excel first. Numbers and tables are rendered by code; the AI writes only the summary part.
- AI unavailable / refused / numbers not in the data -> the report goes out WITHOUT the AI part (no note in the
  email, 25-Nov-2026) and the PMO is warned.
- Recipients: every Report By Email + CC Email in Test Progress, deduped; invalid ones skipped + flagged and the
  PMO is warned. None at all -> the PMO only, with a warning. Dry-run -> PMO only (outbox rules).
  These PMO matters are never written in the email itself; the preview dialog shows them.
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
from ..alerts.templates import _e
from ..domain.models import CHANNEL_APP, CHANNEL_SYSTEM
from ..excel.sheets import SHEETS, TEST_PROGRESS
from ..mail.sender import SendError
from . import weekly_ai
from .auth import WEEKLY, Actor, require, system_actor
from .holidays import refresh_year
from .traffic_light import AMBER, COMPLETE, GREEN, LABEL_EN, LABEL_TH, MISSING, RED
from .importer import run_import
from .weekly_data import build_data, pct
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
# Approved layout 25-Nov-2026: Thai text, English only as a small grey line under each heading.
# Email-safe HTML: tables + inline styles only (no flex / grid / <style>), 640 px wide.

WIDTH = 640
_FONT = "font-family:'IBM Plex Sans Thai',Tahoma,sans-serif"
_INK, _MUTED, _LINE, _SOFT = "#14213D", "#6B7280", "#E5E7EB", "#F3F4F6"
_HEX = {COMPLETE: "#2563EB", RED: "#DC2626", AMBER: "#F59E0B", GREEN: "#16A34A", MISSING: "#9CA3AF"}
_TEXT_HEX = {**_HEX, AMBER: "#B45309"}                   # amber text is darker to stay readable
_BAR_ORDER = (COMPLETE, RED, AMBER, GREEN, MISSING)
_CODE_OF = {LABEL_EN[c]: c for c in LABEL_EN}            # the data carries the English label (weekly_data)
_NONE = "ไม่มี"


def _code(label_en: str) -> str:
    return _CODE_OF.get(label_en, label_en)


def _row(inner: str, pad: str = "16px 24px 0 24px") -> str:
    return f'<tr><td style="padding:{pad}">{inner}</td></tr>'


def _heading(th: str, en: str, *, inline: bool = False) -> str:
    """Thai heading in bold; the English name as a small grey line under it (or after it, joined by ·)."""
    en_html = f'<span style="font-size:12px;font-weight:normal;color:{_MUTED}">{_e(en)}</span>' if en else ""
    if inline:
        body = f'{_e(th)} <span style="font-size:12px;font-weight:normal;color:{_MUTED}">· {_e(en)}</span>'
    else:
        body = _e(th) + (f"<br>{en_html}" if en_html else "")
    return (f'<div style="font-size:15px;font-weight:bold;color:{_INK};margin:0 0 8px 0;padding-top:8px;'
            f'border-top:1px solid {_LINE}">{body}</div>')


def _sub(th: str, en: str = "") -> str:
    en_html = f' <span style="font-weight:normal;color:{_MUTED}">· {_e(en)}</span>' if en else ""
    return f'<div style="font-size:13px;font-weight:bold;color:{_INK};margin:10px 0 4px 0">{_e(th)}{en_html}</div>'


def _note(text: str) -> str:
    return f'<div style="font-size:12px;color:{_MUTED};margin:6px 0 0 0">{_e(text)}</div>'


def _empty() -> str:
    return f'<div style="font-size:14px;color:{_MUTED}">{_NONE}</div>'


def _dot(code: str | None) -> str:
    if code is None:
        return ""
    return f'<span style="color:{_HEX.get(code, _MUTED)};font-size:14px">&#9679;</span>'


def _badge(code: str) -> str:
    c = _TEXT_HEX.get(code, _MUTED)
    return (f'<span style="font-size:11px;color:{c};border:1px solid {c};padding:0 6px;'
            f'white-space:nowrap">{_e(LABEL_TH.get(code, code))}</span>')


def _list(rows: list[str]) -> str:
    """Rows of a section: each row is '<tr>..</tr>' with a hairline under it."""
    return f'<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">{"".join(rows)}</table>'


def _item(left: str, main: str, right: str = "", *, right_color: str = _INK) -> str:
    cell = f"padding:6px 0;border-bottom:1px solid {_LINE};vertical-align:top"
    return (f'<tr><td width="18" style="{cell};width:18px">{left}</td><td style="{cell};font-size:14px">{main}</td>'
            f'<td align="right" style="{cell};font-size:13px;white-space:nowrap;color:{right_color}">{right}</td></tr>')


def _small(text: str) -> str:
    return f'<br><span style="font-size:12px;color:{_MUTED}">{_e(text)}</span>'


# ---- sections (each returns '<tr>..</tr>' rows of the 640 px table)

def _header(data: dict) -> str:
    return _row(
        f'<div style="font-size:12px;color:{_MUTED}">Planning Progress</div>'
        f'<div style="font-size:20px;font-weight:bold;color:{_INK};margin:2px 0 0 0">'
        f'รายงาน UAT ประจำสัปดาห์ · {_e(data["project"])}</div>'
        f'<div style="font-size:12px;color:{_MUTED}">Weekly UAT Report</div>'
        f'<div style="font-size:13px;color:{_INK};margin:6px 0 0 0">ข้อมูล ณ {_e(data["as_of"])}</div>',
        "24px 24px 0 24px")


def _summary_line(data: dict) -> str:
    c = data["overall"]["colors"]
    parts = [f'<span style="color:{_TEXT_HEX[k]};font-weight:bold">{_e(LABEL_TH[k])} {c.get(LABEL_EN[k], 0)} แอป</span>'
             for k in (RED, AMBER, MISSING)]
    return _row(f'<div style="background:{_SOFT};padding:10px 14px;font-size:14px">{" · ".join(parts)}</div>')


def _ai_block(ai: dict) -> str:
    """AI text in Thai only (summary_th, title_th / why_th, action_th / why_th). English AI text is not shown."""
    def items(key: str, head: str) -> str:
        rows = [f'<li style="margin:0 0 4px 0"><b>{_e(i.get(f"{head}_th"))}</b>'
                + (f": {_e(i.get('why_th'))}" if i.get("why_th") else "") + "</li>" for i in ai.get(key) or []]
        return f'<ul style="margin:0 0 4px 0;padding-left:20px;font-size:14px">{"".join(rows)}</ul>' if rows else ""
    risks, follow = items("risks", "title"), items("follow_ups", "action")
    return _row(
        _heading("สรุปสำหรับผู้บริหาร (ร่างโดย AI)", "Executive summary")
        + f'<div style="font-size:14px;line-height:1.6">{_e(ai.get("summary_th"))}</div>'
        + (_sub("ความเสี่ยงหลัก", "Top risks") + risks if risks else "")
        + (_sub("สิ่งที่ควรติดตาม", "Suggested follow-ups") + follow if follow else ""))


def _width(pct_text: str) -> str:
    """'70.3%' -> '70.3%' kept between 0 and 100 for the bar."""
    try:
        v = float(pct_text.rstrip("%"))
    except (AttributeError, ValueError):
        v = 0.0
    return f"{min(max(v, 0.0), 100.0):.1f}%"


def _kpis(data: dict) -> str:
    o = data["overall"]
    box = f"padding:10px 12px;background:{_SOFT};vertical-align:top"
    label_ = f"font-size:12px;color:{_MUTED}"
    big = f"font-size:22px;font-weight:bold;color:{_INK}"
    w = _width(o["pct_executed"])
    bar = (f'<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:6px;border-collapse:collapse">'
           f'<tr><td style="background:{_HEX[COMPLETE]};width:{w};height:6px;font-size:0;line-height:0" width="{w}">'
           f'&nbsp;</td><td style="background:{_LINE};height:6px;font-size:0;line-height:0">&nbsp;</td></tr></table>')

    def cell(title: str, value, extra: str = "", width: str = "22%", color: str = _INK) -> str:
        return (f'<td width="{width}" style="{box}"><div style="{label_}">{_e(title)}</div>'
                f'<div style="{big};color:{color}">{_e(value)}</div>{extra}</td>')
    gap = '<td width="8" style="width:8px">&nbsp;</td>'
    first = cell("ทดสอบแล้ว", o["pct_executed"],
                 f'<div style="font-size:12px;color:{_MUTED}">{o["executed"]} จาก {o["total_tc"]} เคส</div>{bar}', "34%")
    return _row(f'<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tr>'
                + first + gap + cell("Pass", o["pass"], color=_HEX[GREEN]) + gap
                + cell("Fail", o["fail"], color=_HEX[RED]) + gap
                + cell("Defect ที่เปิดอยู่", data["defects"]["open"]) + "</tr></table>")


def _status_bar(data: dict) -> str:
    o = data["overall"]
    c = {k: o["colors"].get(LABEL_EN[k], 0) for k in _BAR_ORDER}
    total = sum(c.values())
    segs = "".join(f'<td width="{pct(n, total)}" style="width:{pct(n, total)};background:{_HEX[k]};height:12px;'
                   f'font-size:0;line-height:0">&nbsp;</td>' for k, n in c.items() if n)
    bar = (f'<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:0 0 8px 0">'
           f'<tr>{segs}</tr></table>') if segs else ""
    legend = " &nbsp; ".join(f'<span style="white-space:nowrap">{_dot(k)} {_e(LABEL_TH[k])} <b>{n}</b></span>'
                             for k, n in c.items())
    return _row(_heading(f"สถานะแอป {o['apps']} แอป", "App status", inline=True) + bar
                + f'<div style="font-size:13px;color:{_INK}">{legend}</div>')


def _reason_th(a: dict, threshold: str) -> str:
    """Thai one-line reason for one row of data['delayed_or_at_risk'] (R8)."""
    kind, fail = a.get("reason_kind"), a.get("fail") or 0
    fail_ok = f"มี Fail {fail} เคส ({a.get('pct_fail')}) แต่ยังไม่เกินเกณฑ์ {threshold}"
    if kind == "overdue":
        return f"เลยวันจบตามแผน ({a.get('planned_finish')}) มา {a.get('days_late')} วัน และยังไม่เสร็จ"
    if kind == "fail_rate":
        if _code(a.get("color")) == RED:
            return f"Fail {a.get('pct_fail')} เกินเกณฑ์ {threshold}"
        return fail_ok
    behind = f"ทดสอบได้ {a.get('pct_exec')} แต่เวลาผ่านไปแล้ว {a.get('pct_time')}"
    return f"{behind} · {fail_ok}" if fail >= 1 else behind


def _watch_list(data: dict) -> str:
    th = data.get("fail_threshold") or "-"
    rows = [_item(_dot(_code(a["color"])), f"<b>{_e(a['app'])}</b> {_badge(_code(a['color']))}"
                  + _small(_reason_th(a, th)), f"{a['executed']}/{a['total_tc']}")
            for a in data["delayed_or_at_risk"]]
    rows += [_item(_dot(MISSING), f"<b>{_e(m['app'])}</b> {_badge(MISSING)}"
                   + _small("ยังไม่ได้กรอก " + ", ".join(m["fields"])), "-")
             for m in data.get("missing_data", [])]
    body = (_list(rows) + _note("ตัวเลขด้านขวาคือเคสที่ทดสอบแล้ว / ทั้งหมด")) if rows else _empty()
    return _row(_heading("แอปที่ต้องติดตาม", "Apps to watch") + body)


def _overdue_plan(data: dict) -> str:
    rows = [_item("", f"<b>{_e(t['id'])}</b> {_e(t['task'])}" + _small(f"กำหนด {t['finish_date']}"),
                  f"เลยมา {t['days_late']} วัน", right_color=_HEX[RED])
            for t in data["overdue"]["uat_plan"]]
    return _row(_heading("งาน UAT Plan ที่เลยกำหนด", "Overdue UAT Plan tasks") + (_list(rows) if rows else _empty()))


def _defect_rows(items: list[dict], overdue: bool) -> str:
    def one(x: dict) -> str:
        meta = f"{x.get('status') or '-'} · เป้าหมาย {x['fix_target_date']}" if overdue else (x.get("status") or "-")
        return _item("", f"<b>#{_e(x['id'])}</b> {_e(x.get('system'))} {_e(x.get('detail'))}" + _small(meta),
                     f"เลยมา {x['days_late']} วัน" if overdue else "", right_color=_HEX[RED])
    return _list([one(x) for x in items])


def _defects(data: dict) -> str:
    d = data["defects"]
    line = (f'<div style="font-size:14px">เปิดอยู่ <b>{d["open"]}</b> · เลยวันเป้าหมาย <b>{len(d["overdue"])}</b>'
            f' · ไม่มีวันเป้าหมาย <b>{len(d["no_target"])}</b></div>')
    parts = [_heading("Defect", ""), line]
    if d["overdue"]:
        parts += [_sub("เลยวันเป้าหมาย", "Past fix target"), _defect_rows(d["overdue"], True)]
    if d["no_target"]:
        parts += [_sub("ไม่มีวันเป้าหมาย", "No fix target"), _defect_rows(d["no_target"], False)]
    return _row("".join(parts))


def _due_week(data: dict) -> str:
    due = data["due_next_7_days"]
    watch = {a["id"]: _code(a["color"]) for a in data["delayed_or_at_risk"]}
    watch.update({m["id"]: MISSING for m in data.get("missing_data", [])})
    rows, last, dotted = [], None, False
    for x in due:
        if x["date"] != last:
            last = x["date"]
            rows.append(f'<tr><td colspan="3" style="padding:8px 0 2px 0;font-size:12px;font-weight:bold;'
                        f'color:{_MUTED}">{_e(x["date"])}</td></tr>')
        code = watch.get(x["id"]) if x["page"] == TEST_PROGRESS else None
        dotted = dotted or code is not None
        page = (f'<span style="font-size:11px;color:{_MUTED};border:1px solid {_LINE};padding:0 6px;'
                f'white-space:nowrap">{_e(x["page"])}</span>')
        rows.append(_item(_dot(code), f"{page} <b>{_e(x['id'])}</b> {_e(x['name'])}"))
    body = _empty() if not rows else _list(rows) + (_note("จุดสีหน้าชื่อคือแอปที่อยู่ในรายการต้องติดตาม")
                                                       if dotted else "")
    return _row(_heading(f"ครบกำหนดใน 1 สัปดาห์ ({len(due)} งาน)", "Due within 1 week") + body)


def _field_name(page: str, key: str) -> str:
    sheet = SHEETS.get(page)
    spec = next((f for f in sheet.fields if f.key == key), None) if sheet else None
    return spec.names[0] if spec else key


def _other(data: dict) -> str:
    w, rs = data["reschedule_window"], data["reschedules_last_7_days"]
    title = f"การเลื่อนวัน {w['from']} ถึง {w['to']}"
    if rs:
        rows = [_item("", f"<b>{_e(x['name'])}</b> {_e(_field_name(x['page'], x['field']))} {_e(x['from'])} → "
                          f"{_e(x['to'])}" + _small(f"{x['date']} · {x['page']} · "
                                                    + (f"เหตุผล: {x['reason']}" if x["where"] == "app" and x["reason"]
                                                       else "แก้ใน Excel ไม่ได้ระบุเหตุผล")))
                for x in rs]
        resched = _sub(title, "Reschedules") + _list(rows)
    else:
        resched = _sub(f"{title}: {_NONE}", "Reschedules")
    return _row(_heading("อื่น ๆ", "Other") + resched
                + f'<div style="font-size:14px;margin:10px 0 0 0">ข้อขัดแย้งที่ยังเปิด: <b>{data["open_conflicts"]}</b></div>')


def _footer(data: dict, ai: dict | None) -> str:
    text = f"ตัวเลขทั้งหมดคำนวณโดยระบบ Planning Progress ของ PMO ({data['project']}) และส่งอัตโนมัติ"
    if ai:
        text += " · ส่วนสรุปร่างโดย AI (Claude)"
    return _row(f'<div style="font-size:11px;color:{_MUTED};border-top:1px solid {_LINE};padding-top:10px">'
                f'{_e(text)}</div>', "20px 24px 24px 24px")


def render(data: dict, ai: dict | None) -> tuple[str, str]:
    """(subject, html). Lays out numbers from build_data only; no calculation here.

    `ai` is the AI content when the AI part succeeded, else None (no AI section, no note: R13)."""
    subject = f"[Planning Progress] รายงาน UAT ประจำสัปดาห์ {data['project']} · {data['as_of']}"
    sections = [_header(data), _summary_line(data), _ai_block(ai) if ai else "", _kpis(data), _status_bar(data),
                _watch_list(data), _overdue_plan(data), _defects(data), _due_week(data), _other(data),
                _footer(data, ai)]
    html = (f'<html><body style="margin:0;padding:0;background:{_SOFT}">\n'      # dry-run banner goes after it
            f'<table width="100%" cellpadding="0" cellspacing="0" style="background:{_SOFT}"><tr>'
            f'<td align="center" style="padding:16px 0">'
            f'<table width="{WIDTH}" cellpadding="0" cellspacing="0" style="width:{WIDTH}px;background:#FFFFFF;'
            f'{_FONT};font-size:14px;color:{_INK};text-align:left">'
            + "".join(sections) + "</table></td></tr></table></body></html>")
    return subject, html


def preview_notes(r: dict) -> list[str]:
    """Lines for the preview dialog: invalid recipients (from r['invalid_recipients']), one per address."""
    try:
        bad = json.loads(r.get("invalid_recipients") or "[]")
    except (TypeError, ValueError):
        return []
    return [f"{b.get('app') or b.get('id')} ({b.get('value')})" for b in bad]


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
    subject, html = render(data, ai.content if ai.ok else None)       # PMO matters stay out of the email (R13)
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
    if g.invalid:
        alert_pmo(cfg, repo, sender, "Weekly report: invalid email addresses skipped", _invalid_detail(g),
                  send_date=iso, now=now, editor=EDITOR)
    ok = result in ("sent", "dry_run")
    return record("ok" if ok else "failed", f"WR-{g.report_id}: {result}", g.report_id, result)


def _invalid_detail(g: Generated, limit: int = 10) -> str:
    """Count and list of the skipped addresses (as the note that used to be in the email)."""
    shown = ", ".join(f"{b['app']} ({b['value']})" for b in g.invalid[:limit])
    more = f" and {len(g.invalid) - limit} more" if len(g.invalid) > limit else ""
    return (f"Report WR-{g.report_id}: skipped {len(g.invalid)} invalid address(es) in Report By Email / CC Email: "
            f"{shown}{more}. Fix them in Test Progress.")


def cal_years_ok(repo, cfg, day: dt.date, now: str, fetcher=None) -> bool:
    if day.year in repo.holiday_years():
        return True
    refresh_year(repo, cfg, day.year, now=now, editor=EDITOR, fetcher=fetcher)
    return day.year in repo.holiday_years()
