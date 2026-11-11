"""Read models for the Part 4 screens (UAT Plan, Test Progress, Defects, Edit History).

The UI only lays these out. Linking rules (my reading; adjust if you see wrong links):
- a UAT Plan task is linked to a Test Progress app when the app name appears as a whole word in
  the task name (e.g. "CCMS run retrieve data from EDW" -> CCMS);
- defects are linked to that app when their System/App. contains the app name.
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass

from ..alerts.rules import is_complete_text
from ..domain.models import CHANNEL_IMPORT, CHANNEL_RESTORE, DATE_FIELDS
from ..excel.sheets import DEFECT_LIST, NUMBER_FIELDS, TEST_PROGRESS, UAT_PLAN, WRITABLE
from ..parsing.dates import fmt
from .dashboard import CLOSED_DEFECT, first_line
from .edits import EditError, edit_item
from .traffic_light import compute

PLANNED_DATE_FIELDS = ("start_date", "finish_date", "planned_start", "planned_finish", "fix_target_date")
DEFECT_COLUMNS = ("Open", "Assigned", "Fixed", "Retest", "Close/Rejected")


def _active(repo, sheet):
    rows = [r for r in repo.load_items(sheet).values() if r.get("active")]
    rows.sort(key=lambda r: (r.get("excel_row") is None, r.get("excel_row") or 0, r["id"]))
    return rows


def _history_by_item(repo, sheet) -> dict[tuple[str, str], list[dict]]:
    out: dict[tuple[str, str], list[dict]] = {}
    for h in repo.all_date_history(sheet):
        out.setdefault((h["item_id"], h["field"]), []).append(h)
    return out


def _word_in(name: str, text: str) -> bool:
    if not name or not text:
        return False
    return re.search(r"(?<![\w-])" + re.escape(name.casefold()) + r"(?![\w-])", text.casefold()) is not None


# --------------------------------------------------------------------------- UAT Plan

UAT_FILTERS = ("all", "in_progress", "not_start", "no_status", "due7", "resched2")


def uat_plan(repo, as_of: dt.date, flt: str = "all", owner: str | None = None, search: str = "") -> list[dict]:
    """Groups A-E: [{code, title, complete, items: [...]}]. Fully complete groups: complete=True
    (the screen collapses them by default)."""
    hist = _history_by_item(repo, UAT_PLAN)
    A, week = as_of.isoformat(), (as_of + dt.timedelta(days=7)).isoformat()
    groups: dict[str, dict] = {}
    for t in _active(repo, UAT_PLAN):
        g = groups.setdefault(t.get("group_code") or "?", {"code": t.get("group_code") or "?",
                                                           "title": t.get("group_title") or "", "items": [],
                                                           "all_complete": True})
        if t.get("status") != "Complete":
            g["all_complete"] = False
        fh = [h["value"] for h in hist.get((t["id"], "finish_date"), []) if h["source"] == "import"]
        app_h = [h for h in hist.get((t["id"], "finish_date"), []) if h["source"] == "app"]
        resched = len(fh) + len([h for h in app_h if h["value"] not in fh])
        first = (fh[0] if fh else (app_h[0]["value"] if app_h else t.get("finish_date")))
        item = {"id": t["id"], "task": first_line(t.get("task_name")), "owner": first_line(t.get("responsible")),
                "status": t.get("status") or "", "planned_first": first, "current": t.get("finish_date"),
                "reschedules": resched,
                "overdue": bool(t.get("finish_date") and t["finish_date"] < A and t.get("status") != "Complete")}
        keep = {
            "all": True, "in_progress": item["status"] == "In progress", "not_start": item["status"] == "Not Start",
            "no_status": not item["status"],
            "due7": bool(item["current"] and A <= item["current"] <= week and item["status"] != "Complete"),
            "resched2": resched >= 2,
        }[flt]
        if owner and owner.casefold() not in (t.get("responsible") or "").casefold():
            keep = False
        if search and search.casefold() not in f"{t['id']} {t.get('task_name')} {t.get('responsible')}".casefold():
            keep = False
        if keep:
            g["items"].append(item)
    out = []
    for g in sorted(groups.values(), key=lambda g: g["code"]):
        g["complete"] = g.pop("all_complete")
        if g["items"]:
            out.append(g)
    return out


def uat_owners(repo) -> list[str]:
    return sorted({first_line(t.get("responsible")) for t in _active(repo, UAT_PLAN) if t.get("responsible")})


def uat_detail(cfg, repo, item_id: str, as_of: dt.date) -> dict:
    t = repo.get_item(UAT_PLAN, item_id)
    timeline = []
    for f in ("start_date", "finish_date"):
        for h in repo.date_history(UAT_PLAN, item_id, f):
            timeline.append({"field": f, "value": h["value"], "source": h["source"], "reason": h["reason"],
                             "root_cause": h["root_cause"], "at": h["recorded_at"]})
    apps = [r for r in _active(repo, TEST_PROGRESS) if _word_in(first_line(r.get("application")), t.get("task_name"))]
    linked = []
    for r in apps:
        light = compute(r, as_of, cfg.alerts.fail_threshold)
        linked.append({"id": r["id"], "name": first_line(r.get("application")), "executed": int(r.get("executed") or 0),
                       "total": int(r.get("total_tc") or 0), "color": light.color})
    names = [a["name"] for a in linked]
    defects = [{"id": x["id"], "system": first_line(x.get("system")), "detail": first_line(x.get("detail")),
                "status": x.get("status") or ""}
               for x in _active(repo, DEFECT_LIST) if any(_word_in(n, x.get("system")) for n in names)]
    return {"item": t, "timeline": timeline, "linked_apps": linked, "defects": defects}


# --------------------------------------------------------------------------- Test Progress

TP_FILTERS = ("follow", "incomplete", "complete", "all")
_INCOMPLETE_FLAGS = ("missing_report_email", "missing_fail", "missing_data", "pass_fail_mismatch",
                     "complete_no_actual_finish")


def test_progress(cfg, repo, as_of: dt.date, flt: str = "all", search: str = "") -> tuple[dict, list[dict]]:
    flags: dict[str, set[str]] = {}
    for f in repo.active_flags():
        if f["sheet"] == TEST_PROGRESS:
            flags.setdefault(f["item_id"], set()).add(f["code"])
    uploads = {a["item_id"] for a in repo.attachments(TEST_PROGRESS)}
    rows, summary = [], {"executed": 0, "total": 0, "pass": 0, "fail": 0, "amber": 0, "red": 0}
    for r in _active(repo, TEST_PROGRESS):
        light = compute(r, as_of, cfg.alerts.fail_threshold)
        ex, tot = int(r.get("executed") or 0), int(r.get("total_tc") or 0)
        ps, fl = int(r.get("pass_count") or 0), int(r.get("fail_count") or 0)
        summary["executed"] += ex
        summary["total"] += tot
        summary["pass"] += ps
        summary["fail"] += fl
        summary["amber"] += light.color == "Amber"
        summary["red"] += light.color == "Red"
        fl_codes = flags.get(r["id"], set())
        row = {"id": r["id"], "app": first_line(r.get("application")), "reporter": first_line(r.get("report_by")),
               "executed": ex, "total": tot, "pass": ps, "fail": fl,
               "mismatch": "pass_fail_mismatch" in fl_codes,
               "planned_start": r.get("planned_start"), "actual_start": r.get("actual_start"),
               "planned_finish": r.get("planned_finish"), "actual_finish": r.get("actual_finish"),
               "color": light.color, "reason": light.reason, "status": r.get("status") or "",
               "tc_result": "Complete" if is_complete_text(r.get("tc_result_complete")) else (r.get("tc_result_complete") or ""),
               "uploaded": r["id"] in uploads, "flags": sorted(fl_codes), "record": r}
        keep = {"all": True, "complete": light.color == "Complete",
                "follow": light.color in ("Red", "Amber", "Missing"),
                "incomplete": bool(fl_codes & set(_INCOMPLETE_FLAGS))}[flt]
        if search and search.casefold() not in f"{r['id']} {r.get('application')} {r.get('report_by')}".casefold():
            keep = False
        if keep:
            rows.append(row)
    return summary, rows


# --------------------------------------------------------------------------- Defects

def defect_column(status: str | None) -> str:
    if status in ("Close", "Rejected"):
        return "Close/Rejected"
    return status if status in DEFECT_COLUMNS else "Open"


def defects(repo, as_of: dt.date, search: str = "") -> list[dict]:
    A = as_of.isoformat()
    out = []
    for x in _active(repo, DEFECT_LIST):
        if search and search.casefold() not in f"{x['id']} {x.get('system')} {x.get('detail')} {x.get('assigned_to')}".casefold():
            continue
        open_ = x.get("status") not in CLOSED_DEFECT
        target = x.get("fix_target_date")
        sev = str(x.get("severity") or "").strip()
        out.append({
            "id": x["id"], "no": x.get("no") or "", "system": first_line(x.get("system")),
            "detail": first_line(x.get("detail")), "status": x.get("status") or "Open",
            "column": defect_column(x.get("status")), "severity": sev,
            "severe": sev[:1] in ("1", "2"), "owner": first_line(x.get("assigned_to")),
            "days_open": (as_of - dt.date.fromisoformat(x["defect_date"])).days if x.get("defect_date") and open_ else None,
            "overdue": bool(open_ and target and target < A), "no_target": bool(open_ and not target),
            "target": target, "record": x,
        })
    return out


# --------------------------------------------------------------------------- Edit History

HISTORY_TYPES = ("Reschedule", "Status change", "Data edit", "Import", "Access")


def change_type(e: dict) -> str:
    if e["channel"] == CHANNEL_IMPORT:
        return "Import"
    if e["channel"] == "auth":                 # Part 5: login, logout, lockout, users, denied actions
        return "Access"
    if e["field"] == "status":
        return "Status change"
    if e["field"] in PLANNED_DATE_FIELDS and e["old_value"] and e["new_value"]:
        return "Reschedule"
    return "Data edit"


def history(repo, *, text=None, sheet=None, editor=None, date_from=None, date_to=None, kind=None,
            item_id=None) -> list[dict]:
    rows = repo.search_audit(text=text, sheet=sheet, editor=editor, date_from=date_from, date_to=date_to,
                             item_id=item_id)
    out = []
    for e in rows:
        e["type"] = change_type(e)
        e["restorable"] = can_restore(e)
        if kind is None or e["type"] == kind:
            out.append(e)
    return out


def can_restore(e: dict) -> bool:
    return (e.get("sheet") in WRITABLE and e.get("item_id") and e.get("field") in WRITABLE[e["sheet"]]
            and e.get("old_value") is not None and not str(e.get("reason") or "").startswith("CONFLICT"))


def _parse_display(field: str, text: str):
    """Audit values are shown as people see them (DD-Mon-YYYY, chains). Back to storage form."""
    if text is None or text == "":
        return None
    if field in DATE_FIELDS:
        last = text.split("=>")[-1].strip()
        return dt.datetime.strptime(last, "%d-%b-%Y").date().isoformat()
    if field in NUMBER_FIELDS:
        return float(text) if "." in text else int(text)
    return text


def restore(cfg, repo, audit_id: int, *, actor) -> list[str]:
    """Put a field back to the value it had before audit entry #id. A NEW change (and entry)."""
    from .auth import RESTORE, require
    require(repo, actor, RESTORE, item_id=f"history #{audit_id}")
    e = repo.get_audit(audit_id)
    if e is None or not can_restore(e):
        raise EditError("This entry cannot be restored")
    value = _parse_display(e["field"], e["old_value"])
    return edit_item(cfg, repo, e["sheet"], e["item_id"], {e["field"]: value}, actor=actor,
                     reason=f"Restore of change #{audit_id}", channel=CHANNEL_RESTORE)


def export_history(rows: list[dict], path) -> None:
    """Edit History -> new .xlsx (never the source workbook)."""
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Edit History"
    head = ["#", "Date", "Editor", "Type", "Page", "Item", "Field", "Old", "New", "Reason", "Channel"]
    ws.append(head)
    for c in ws[1]:
        c.font = Font(bold=True)
    for e in rows:
        ws.append([e["id"], e["ts"][:19].replace("T", " "), e["editor"], e.get("type") or change_type(e),
                   e["sheet"], e["item_id"], e["field"], e["old_value"], e["new_value"], e["reason"], e["channel"]])
    for col, width in zip("ABCDEFGHIJK", (7, 20, 28, 14, 14, 10, 22, 30, 30, 40, 11)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    wb.save(path)


def show_date(v) -> str:
    return fmt(v) if v else ""
