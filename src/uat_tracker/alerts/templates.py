"""Email templates, one per rule. Thai first, then English (approved). stdlib string.Template.

Layout: greeting to Report By, the message, a small fact table, the one action needed, footer.
"""

from __future__ import annotations

import html
from string import Template

from ..parsing.dates import fmt
from .rules import DAILY, R1, R2, R3, R4A, R4B, TODAY, Alert

COLOR_STYLE = {   # spec §9.1
    "Red": ("#A3260F", "#FDE7E3"), "Amber": ("#7A4E00", "#FFF1D6"),
    "Green": ("#1E6B3E", "#E3F2E8"), "Complete": ("#1E6B3E", "#E3F2E8"),
}
COLOR_TH = {"Red": "แดง (Red)", "Amber": "เหลือง (Amber)", "Green": "เขียว (Green)",
            "Complete": "เสร็จ (Complete)"}

SUBJECTS = {
    (R1, TODAY): Template("[UAT] $app: testing starts today ($date)"),
    (R1, "shifted"): Template("[UAT] $app: testing starts on $date"),
    (R1, DAILY): Template("[UAT] $app: please fill Actual Start (planned $date)"),
    (R2, TODAY): Template("[UAT] $app: planned finish today, please fill Actual Finish"),
    (R2, "shifted"): Template("[UAT] $app: planned finish $date, please fill Actual Finish"),
    (R2, DAILY): Template("[UAT] $app: Actual Finish overdue since $date"),
    (R3, None): Template("[UAT] $app: progress is now $color"),
    (R4A, DAILY): Template("[UAT] $app: testing done, please set Status to Complete"),
    (R4B, DAILY): Template("[UAT] $app: please upload test-result documents"),
}

# (message TH, message EN, action TH, action EN)
BODIES = {
    (R1, TODAY): (
        "การทดสอบ UAT ของ $app เริ่มวันนี้ ($date) ตามแผน",
        "UAT testing for $app starts today ($date) as planned.",
        "เมื่อเริ่มทดสอบแล้ว กรุณากรอก Actual Start ในชีต Test Progress",
        "Once testing has started, please fill in Actual Start in the Test Progress sheet."),
    (R1, "shifted"): (
        "การทดสอบ UAT ของ $app มีแผนเริ่มวันที่ $date ซึ่งเป็นวันหยุด จึงแจ้งล่วงหน้าในวันทำการก่อนหน้า",
        "UAT testing for $app is planned to start on $date, a non-working day, so this notice comes "
        "on the previous working day.",
        "เมื่อเริ่มทดสอบแล้ว กรุณากรอก Actual Start ในชีต Test Progress",
        "Once testing has started, please fill in Actual Start in the Test Progress sheet."),
    (R1, DAILY): (
        "$app มีแผนเริ่มทดสอบวันที่ $date แต่ยังไม่ได้กรอก Actual Start",
        "$app was planned to start testing on $date, but Actual Start is still empty.",
        "กรุณากรอก Actual Start ในชีต Test Progress หากยังไม่เริ่ม กรุณาแจ้ง PMO เพื่อปรับแผน",
        "Please fill in Actual Start in the Test Progress sheet. If testing has not started, "
        "please tell the PMO so the plan can be updated."),
    (R2, TODAY): (
        "วันนี้ ($date) เป็นวันสิ้นสุดการทดสอบตามแผนของ $app",
        "Today ($date) is the planned finish date for $app.",
        "เมื่อทดสอบเสร็จ กรุณากรอก Actual Finish ในชีต Test Progress",
        "When testing is finished, please fill in Actual Finish in the Test Progress sheet."),
    (R2, "shifted"): (
        "$app มีแผนสิ้นสุดการทดสอบวันที่ $date ซึ่งเป็นวันหยุด จึงแจ้งล่วงหน้าในวันทำการก่อนหน้า",
        "$app is planned to finish testing on $date, a non-working day, so this reminder comes "
        "on the previous working day.",
        "เมื่อทดสอบเสร็จ กรุณากรอก Actual Finish ในชีต Test Progress",
        "When testing is finished, please fill in Actual Finish in the Test Progress sheet."),
    (R2, DAILY): (
        "$app เลยวันสิ้นสุดตามแผน ($date) แล้ว แต่ยังไม่ได้กรอก Actual Finish",
        "$app has passed its planned finish date ($date), but Actual Finish is still empty.",
        "กรุณากรอก Actual Finish ในชีต Test Progress หากต้องเลื่อนแผน กรุณาแจ้ง PMO พร้อมเหตุผล",
        "Please fill in Actual Finish in the Test Progress sheet. If the plan must move, please tell "
        "the PMO and give the reason."),
    (R3, None): (
        "สถานะความคืบหน้า (Traffic Light) ของ $app เปลี่ยนเป็น $color_th",
        "The progress status (Traffic Light) of $app is now $color.",
        "กรุณาตรวจสอบความคืบหน้า และแจ้ง PMO ถึงสาเหตุและแผนการแก้ไข",
        "Please check progress and tell the PMO the cause and the recovery plan."),
    (R4A, DAILY): (
        "$app ทดสอบครบทุกเคสแล้ว (Executed = Total TC) และกรอก Actual Finish แล้ว แต่ Status ยังไม่เป็น Complete",
        "$app has executed every test case (Executed = Total TC) and Actual Finish is filled in, "
        "but Status is not Complete yet.",
        "กรุณาเปลี่ยน Status เป็น Complete ในชีต Test Progress",
        "Please set Status to Complete in the Test Progress sheet."),
    (R4B, DAILY): (
        "$app ทดสอบครบทุกเคสแล้ว แต่ยังไม่ได้ส่งเอกสารผลการทดสอบ (TC/Result Complete ยังไม่เป็น Complete)",
        "$app has executed every test case, but the test-result documents have not been submitted "
        "(TC/Result Complete is not Complete).",
        "กรุณาส่งไฟล์ผลการทดสอบให้ PMO เพื่ออัปโหลดเข้าระบบ",
        "Please send the test-result files to the PMO so they can be uploaded."),
}

MANUAL_SUBJECT = Template("[UAT] Reminder: $item")
PMO_ALERT_SUBJECT = Template("[Planning Progress] $project$title")

# Multi-project: the project whose data is being sent (set when a project is opened).
PROJECT = {"name": ""}


def set_project(name: str) -> None:
    PROJECT["name"] = name or ""


def footer() -> str:
    p = f" ({PROJECT['name']})" if PROJECT["name"] else ""
    return (f"อีเมลนี้ส่งอัตโนมัติจากระบบ Planning Progress ของ PMO{p} · "
            f"This email was sent automatically by the PMO's Planning Progress app{p}.")


def _e(s) -> str:
    return html.escape("" if s is None else str(s))


def _num(v) -> str:
    if v is None or v == "":
        return "-"
    return f"{v:g}" if isinstance(v, (int, float)) else str(v)


def _color_badge(color: str | None) -> str:
    if not color:
        return "-"
    fg, bg = COLOR_STYLE.get(color, ("#333", "#eee"))
    return (f'<span style="color:{fg};background:{bg};padding:2px 8px;border-radius:10px;'
            f'font-weight:600">{_e(color)}</span>')


def fact_rows(rec: dict, light=None) -> list[tuple[str, str]]:
    """(label, html value) rows for a Test Progress item."""
    rows = [
        ("Application", _e(rec.get("application"))),
        ("Report By", _e(rec.get("report_by"))),
        ("Executed / Total TC", f"{_num(rec.get('executed'))} / {_num(rec.get('total_tc'))}"),
        ("#Pass / #Fail", f"{_num(rec.get('pass_count'))} / {_num(rec.get('fail_count'))}"),
        ("Planned Start – Finish",
         f"{fmt(rec.get('planned_start')) or '-'} – {fmt(rec.get('planned_finish')) or '-'}"),
        ("Actual Start / Finish",
         f"{fmt(rec.get('actual_start')) or '-'} / {fmt(rec.get('actual_finish')) or '-'}"),
    ]
    if light is not None:
        rows.append(("Traffic Light", _color_badge(light.color if light.color != "Missing" else None)
                     + (f" <span style='color:#555'>{_e(light.reason)}</span>" if light.reason else "")))
    if rec.get("status"):
        rows.append(("Status", _e(rec.get("status"))))
    return rows


def layout(greeting_name: str | None, msg_th: str, msg_en: str, facts: list[tuple[str, str]],
           action_th: str, action_en: str, banner: str = "") -> str:
    name = _e(greeting_name) if greeting_name else ""
    table = "".join(
        f'<tr><td style="padding:4px 12px 4px 0;color:#555;white-space:nowrap">{_e(k)}</td>'
        f'<td style="padding:4px 0">{v}</td></tr>' for k, v in facts)
    return f"""<html><body style="font-family:'IBM Plex Sans Thai',Tahoma,sans-serif;font-size:14px;color:#14213D">
{banner}
<p>เรียน คุณ{name}<br>Dear {name or 'team'},</p>
<p>{_e(msg_th)}<br><span style="color:#444">{_e(msg_en)}</span></p>
<table style="border-collapse:collapse;margin:8px 0 12px 0">{table}</table>
<p style="background:#F4F6F9;border-left:4px solid #1F4FBF;padding:8px 12px">
<b>สิ่งที่ต้องทำ:</b> {_e(action_th)}<br><b>Action:</b> {_e(action_en)}</p>
<p style="color:#777;font-size:12px">{_e(footer())}</p>
</body></html>"""


def render(alert: Alert, rec: dict, light=None) -> tuple[str, str]:
    """Subject + HTML body for an auto alert."""
    variant = "shifted" if alert.variant == TODAY and alert.shifted else alert.variant
    key = (alert.rule, variant)
    color = alert.event.color if alert.event else ""
    vars_ = {
        "app": rec.get("application") or alert.item_id,
        "date": fmt(alert.trigger) if alert.trigger else "",
        "color": color, "color_th": COLOR_TH.get(color, color),
    }
    subject = SUBJECTS[key].substitute(vars_)
    msg_th, msg_en, act_th, act_en = (Template(s).substitute(vars_) for s in BODIES[key])
    facts = fact_rows(rec, light)
    if alert.event is not None:
        facts.append(("Reason / เหตุผล", _e(alert.event.reason)))
    return subject, layout(rec.get("report_by"), msg_th, msg_en, facts, act_th, act_en)


def render_manual(sheet: str, rec: dict, note: str | None, recipient_name: str | None) -> tuple[str, str]:
    title = rec.get("task_name") or rec.get("detail") or rec["id"]
    label = f"{rec['id']} {title}"
    if sheet == "UAT Plan":
        facts = [("ID", _e(rec["id"])), ("Task", _e(rec.get("task_name"))),
                 ("Responsible", _e(rec.get("responsible"))), ("Status", _e(rec.get("status") or "-")),
                 ("Start – Finish", f"{fmt(rec.get('start_date')) or '-'} – {fmt(rec.get('finish_date')) or '-'}")]
    else:
        facts = [("ID", _e(rec["id"])), ("System/App", _e(rec.get("system"))),
                 ("Issue", _e(rec.get("detail"))), ("Status", _e(rec.get("status") or "-")),
                 ("Severity", _e(rec.get("severity") or "-")), ("Assigned To", _e(rec.get("assigned_to"))),
                 ("Fix target date", fmt(rec.get("fix_target_date")) or "-")]
    if note:
        facts.append(("Note from PMO", _e(note)))
    subject = MANUAL_SUBJECT.substitute(item=label)
    body = layout(recipient_name, f"PMO ขอติดตามความคืบหน้าของงาน: {label}",
                  f"The PMO is following up on: {label}", facts,
                  "กรุณาอัปเดตสถานะในไฟล์ Excel ของโปรเจกต์ หรือตอบกลับอีเมลนี้",
                  "Please update the status in the project's Excel file or reply to this email.")
    return subject, body


def render_pmo_alert(title: str, detail: str) -> tuple[str, str]:
    subject = PMO_ALERT_SUBJECT.substitute(title=title, project=f"{PROJECT['name']}: " if PROJECT["name"] else "")
    body = layout("PMO", title, title, [("Detail", _e(detail).replace("\n", "<br>"))],
                  "กรุณาตรวจสอบ", "Please check.")
    return subject, body


def dry_run_banner(to: str, cc: list[str]) -> str:
    return ('<div style="background:#FFF1D6;color:#7A4E00;border:1px solid #7A4E00;padding:8px 12px;'
            'margin-bottom:12px"><b>DRY RUN</b>: this email was NOT sent to the real recipients.<br>'
            f'Real To: {_e(to)}<br>Real CC: {_e("; ".join(cc)) or "(none)"}</div>')


def with_banner(body_html: str, banner: str) -> str:
    return body_html.replace("\n", "\n" + banner, 1) if "<body" in body_html else banner + body_html
