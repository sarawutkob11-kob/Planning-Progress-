"""New items created in the app (+ Add task, + Add Application, + Add Defect). The sync adds the Excel row:
- UAT Plan task -> a new row inserted right after the last task of its group; a task of a NEW group
  (multi-project, approved Q13) -> group header row (letter + title) + the task, after the last task;
- Test Progress application -> a new row right after the last application, with the %Progress formula;
- Defect -> the first empty row after the last defect (the sheet keeps pre-numbered empty rows).
Until the row exists the item is kept (never marked "removed").
"""

from __future__ import annotations

import json
from datetime import date, datetime

from ..domain.models import CHANNEL_APP, DATE_FIELDS, HISTORY_SUFFIX, NOTE_SUFFIX
from ..excel.sheets import DEFECT_LIST, TEST_PROGRESS, UAT_PLAN
from .auth import ADD_APP, ADD_DEFECT, ADD_TASK, Actor, require
from .edits import EditError, validate_value
from .values import canon

ROW = "(row)"


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _in_window(cfg, f, v):
    if f in DATE_FIELDS and v and not (cfg.window_start.isoformat() <= v <= cfg.window_end.isoformat()):
        raise EditError(f"{f}: {v} is outside the project date window ({cfg.window_start} .. {cfg.window_end})")


def _create(cfg, repo, sheet, item_id, fields: dict, *, editor, now, label):
    for k in list(fields):
        if k in DATE_FIELDS:
            fields.setdefault(k + HISTORY_SUFFIX, ())
            fields.setdefault(k + NOTE_SUFFIX, None)
    with repo.transaction():
        repo.insert_new_item(sheet, item_id, fields, now)
        shown = {k: v for k, v in fields.items() if v not in (None, "", ()) and not k.endswith(HISTORY_SUFFIX)}
        repo.append_audit({"ts": now, "editor": editor, "channel": CHANNEL_APP, "sheet": sheet, "item_id": item_id,
                           "field": "*", "old_value": "", "new_value": json.dumps(shown, ensure_ascii=False),
                           "reason": label})
        repo.upsert_pending(sheet, item_id, ROW, value="", excel_value=json.dumps({"kind": "row"}), now=now,
                            reason=label, editor=editor, channel=CHANNEL_APP)
    return item_id


def groups(repo) -> dict[str, str]:
    """{group letter: title} of the active UAT Plan tasks."""
    return {r.get("group_code"): r.get("group_title") for r in repo.load_items(UAT_PLAN).values()
            if r.get("active") and r.get("group_code")}


def next_group_letter(repo) -> str:
    used = {g.upper() for g in groups(repo)}
    letters = [chr(c) for c in range(ord("A"), ord("Z") + 1)]
    after = [c for c in letters if c > max(used, default="@") and c not in used]
    free = after or [c for c in letters if c not in used]
    if not free:
        raise EditError("All group letters A-Z are used")
    return free[0]


def add_task(cfg, repo, *, group: str, task_name: str, responsible: str, task_no: str = "",
             status: str | None = None, start_date: str | None = None, finish_date: str | None = None,
             note: str = "", remark: str = "", actor: Actor, now: str | None = None,
             new_group_title: str | None = None) -> str:
    """new_group_title: create a new group (next free letter, approved Q13); `group` is then ignored."""
    require(repo, actor, ADD_TASK)
    now = now or _now()
    items = repo.load_items(UAT_PLAN)
    titles = {r.get("group_code"): r.get("group_title") for r in items.values() if r.get("active")}
    if new_group_title is not None:
        if not new_group_title.strip():
            raise EditError("New group title is required")
        group = next_group_letter(repo)
        titles[group] = new_group_title.strip()
    if group not in titles:
        raise EditError(f"Group {group!r} not found (existing: {', '.join(sorted(g for g in titles if g))})")
    if not task_name.strip():
        raise EditError("Task name is required")
    if not responsible.strip():
        raise EditError("Responsible is required")
    fields = {"group_code": group, "group_title": titles[group], "task_no": task_no.strip() or None,
              "task_name": task_name.strip(), "responsible": responsible.strip(),
              "status": validate_value(UAT_PLAN, "status", status),
              "start_date": validate_value(UAT_PLAN, "start_date", start_date),
              "finish_date": validate_value(UAT_PLAN, "finish_date", finish_date),
              "note": note.strip() or None, "remark": remark.strip() or None}
    for f in ("start_date", "finish_date"):
        _in_window(cfg, f, fields[f])
    if fields["start_date"] and fields["finish_date"] and fields["finish_date"] < fields["start_date"]:
        raise EditError("Finish date is before Start date")
    base = f"{group}-{fields['task_no']}" if fields["task_no"] else None
    item_id = _free_id(items, base, f"{group}-x")
    return _create(cfg, repo, UAT_PLAN, item_id, fields, editor=actor.editor, now=now, label="Task added in the app")


def add_application(cfg, repo, *, application: str, report_by: str, planned_start: str, planned_finish: str,
                    report_by_email: str = "", cc_emails: str = "", total_tc=None, remark: str = "",
                    actor: Actor, now: str | None = None) -> str:
    """+ Add Application (multi-project, section C). Required: Application, Report By, Planned Start / Finish
    (the Traffic Light needs them). No. = highest + 1; ID TP-<No.>."""
    require(repo, actor, ADD_APP)
    now = now or _now()
    if not application.strip():
        raise EditError("Application is required")
    if not report_by.strip():
        raise EditError("Report By is required")
    ps = validate_value(TEST_PROGRESS, "planned_start", planned_start)
    pf = validate_value(TEST_PROGRESS, "planned_finish", planned_finish)
    if not ps or not pf:
        raise EditError("Planned Start and Planned Finish are required")
    if pf < ps:
        raise EditError("Planned Finish is before Planned Start")
    for f, v in (("planned_start", ps), ("planned_finish", pf)):
        _in_window(cfg, f, v)
    items = repo.load_items(TEST_PROGRESS)
    if any(r.get("active") and (r.get("application") or "").strip().casefold() == application.strip().casefold()
           for r in items.values()):
        raise EditError(f"Application {application.strip()!r} is already in Test Progress")
    nums = [int(canon("no", r.get("no"))) for r in items.values() if canon("no", r.get("no")).isdigit()]
    no = str(max(nums, default=0) + 1)
    fields = {"no": no, "application": application.strip(), "report_by": report_by.strip(),
              "report_by_email": validate_value(TEST_PROGRESS, "report_by_email", report_by_email),
              "cc_emails": validate_value(TEST_PROGRESS, "cc_emails", cc_emails),
              "total_tc": validate_value(TEST_PROGRESS, "total_tc", total_tc),
              "planned_start": ps, "planned_finish": pf, "remark": remark.strip() or None}
    item_id = _free_id(items, f"TP-{no}", "TP-x")
    return _create(cfg, repo, TEST_PROGRESS, item_id, fields, editor=actor.editor, now=now,
                   label="Application added in the app")


def add_defect(cfg, repo, *, system: str, detail: str, assigned_to: str, fix_target_date: str,
               severity: str = "", raised_by: str = "", data_test: str = "", remark: str = "",
               defect_date: str | None = None, actor: Actor, now: str | None = None) -> str:
    require(repo, actor, ADD_DEFECT)
    now = now or _now()
    if not system.strip() or not detail.strip():
        raise EditError("System/App. and Issue detail are required")
    if not assigned_to.strip():
        raise EditError("Owner (Assigned To) is required - one named person")
    target = validate_value(DEFECT_LIST, "fix_target_date", fix_target_date)
    if not target:
        raise EditError("Fix target date is required")
    ddate = validate_value(DEFECT_LIST, "defect_date", defect_date or date.today().isoformat())
    for f, v in (("fix_target_date", target), ("defect_date", ddate)):
        _in_window(cfg, f, v)
    items = repo.load_items(DEFECT_LIST)
    nums = [int(canon("no", r.get("no"))) for r in items.values() if canon("no", r.get("no")).isdigit()]
    no = str(max(nums, default=0) + 1)
    fields = {"no": no, "defect_date": ddate, "system": system.strip(), "detail": detail.strip(), "status": "Open",
              "severity": severity.strip() or None, "assigned_to": assigned_to.strip(), "fix_target_date": target,
              "data_test": data_test.strip() or None, "raised_by": raised_by.strip() or None,
              "remark": remark.strip() or None}
    item_id = _free_id(items, f"DEF-{no}", "DEF-x")
    return _create(cfg, repo, DEFECT_LIST, item_id, fields, editor=actor.editor, now=now, label="Defect added in the app")


def _free_id(items: dict, base: str | None, blank_prefix: str) -> str:
    if base and base not in items:
        return base
    stem = base or blank_prefix
    n = 1 if not base else 2
    while (f"{stem}{n}" if not base else f"{base}-{n}") in items:
        n += 1
    return f"{stem}{n}" if not base else f"{base}-{n}"
