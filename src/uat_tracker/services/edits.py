"""The one way the app changes item data (Part 3; the Part 4 UI calls this).

Each change: validate -> DB -> audit log -> write queue (Excel write-back by the sync).
Business rules enforced here, not in the UI:
- a date that already had a value and changes = reschedule -> reason REQUIRED (root cause optional
  here; the Part 4 dialog asks for it); the old date is kept in the reschedule history;
- Test Progress: Actual Finish > Planned Finish is refused unless Planned Finish is moved in the
  same edit (with reason);
- Test Progress: Status Complete needs Actual Finish;
- status values must be standard; email fields must be valid addresses; numbers >= 0.
"""

from __future__ import annotations

import json
import re
from datetime import datetime

from ..domain.models import CHANNEL_APP, DATE_FIELDS
from ..excel.sheets import DEFECT_LIST, NUMBER_FIELDS, TEST_PROGRESS, UAT_PLAN, WRITABLE
from ..parsing.dates import DateWindow
from ..parsing.status import normalize_defect_status, normalize_task_status
from .auth import EDIT, Actor, require
from .importer import display
from .values import canon, compose_date_cell, derive_excel_value, typed

ROOT_CAUSES = ("Dev not done", "Waiting other system data", "Defect found", "Environment not ready",
               "Waiting user check")
_EMAIL = re.compile(r"^[^@\s;]+@[^@\s;]+\.[^@\s;]+$")


class EditError(Exception):
    pass


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def validate_value(sheet: str, field: str, value):
    """Return the value in storage form, or raise EditError."""
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return None
    if field in DATE_FIELDS:
        try:
            return datetime.fromisoformat(str(value)[:10]).date().isoformat()
        except ValueError:
            raise EditError(f"{field}: not a date (expected YYYY-MM-DD from the date picker): {value!r}")
    if field in NUMBER_FIELDS:
        try:
            n = typed(field, canon(field, value))
        except (TypeError, ValueError):
            raise EditError(f"{field}: not a number: {value!r}")
        if n < 0:
            raise EditError(f"{field}: must be 0 or more")
        return n
    if field == "status":
        status, problem = (normalize_defect_status if sheet == DEFECT_LIST else normalize_task_status)(value)
        if problem or status is None:
            raise EditError(f"Status {value!r} is not a standard value")
        return status
    if field == "report_by_email":
        v = str(value).strip()
        if not _EMAIL.match(v):
            raise EditError(f"Report By Email must be ONE valid address: {value!r}")
        return v
    if field == "cc_emails":
        parts = [p.strip() for p in str(value).split(";") if p.strip()]
        bad = [p for p in parts if not _EMAIL.match(p)]
        if bad:
            raise EditError(f"CC Email has invalid address(es): {', '.join(bad)}")
        return "; ".join(parts) or None
    return str(value).strip()


def edit_item(cfg, repo, sheet: str, item_id: str, changes: dict, *, actor: Actor, reason: str | None = None,
              root_cause: str | None = None, channel: str = CHANNEL_APP, now: str | None = None,
              in_transaction: bool = True) -> list[str]:
    """Apply changes {field: value}. Returns the fields that actually changed."""
    now = now or _now()
    rec = repo.get_item(sheet, item_id)
    if rec is None or not rec.get("active"):
        raise EditError(f"No active {sheet} item {item_id!r}")
    require(repo, actor, EDIT, sheet=sheet, item_id=item_id, rec=rec)       # Part 5: role / own row
    editor = actor.editor
    bad = [f for f in changes if f not in WRITABLE[sheet]]
    if bad:
        raise EditError(f"Field(s) not editable in {sheet}: {', '.join(bad)}")
    if root_cause is not None and root_cause not in ROOT_CAUSES:
        raise EditError(f"Root cause must be one of: {', '.join(ROOT_CAUSES)}")
    new = {f: validate_value(sheet, f, v) for f, v in changes.items()}
    in_conflict = [c["field"] for c in repo.open_conflicts(sheet) if c["item_id"] == item_id and c["field"] in new]
    if in_conflict:
        raise EditError("Resolve the open conflict first (Conflict Center): " + ", ".join(in_conflict))
    for f, v in new.items():
        if f in DATE_FIELDS and v and not (cfg.window_start.isoformat() <= v <= cfg.window_end.isoformat()):
            raise EditError(f"{f}: {v} is outside the project date window "
                            f"({cfg.window_start} .. {cfg.window_end}); change the window in config.toml first")
    changed = {f: v for f, v in new.items() if canon(f, v) != canon(f, rec.get(f))}
    if not changed:
        return []

    reschedules = [f for f in changed if f in DATE_FIELDS and rec.get(f) and changed[f]
                   and f not in ("actual_start", "actual_finish", "defect_date")]
    if reschedules and not (reason and reason.strip()):
        raise EditError("Reason is required when a planned date moves: " + ", ".join(reschedules))
    if sheet == UAT_PLAN:                      # same rules as + Add task
        after = {**rec, **changed}
        for f, msg in (("task_name", "Task name is required"), ("responsible", "Responsible is required")):
            if f in changed and not after.get(f):
                raise EditError(msg)
        if ({"start_date", "finish_date"} & set(changed) and after.get("start_date") and after.get("finish_date")
                and after["finish_date"] < after["start_date"]):
            raise EditError("Finish date is before Start date")
    if sheet == TEST_PROGRESS:
        after = {**rec, **changed}
        af, pf = after.get("actual_finish"), after.get("planned_finish")
        if af and pf and af > pf and (
                "actual_finish" in changed or "planned_finish" in changed):
            raise EditError("Actual Finish is after Planned Finish: move Planned Finish first (with a reason)")
        if (("status" in changed or "actual_finish" in changed)
                and after.get("status") == "Complete" and not after.get("actual_finish")):
            raise EditError("Status Complete needs Actual Finish")

    window = DateWindow(cfg.window_start, cfg.window_end)
    baseline = repo.load_baseline(sheet)
    pending = repo.pending(sheet)

    def apply():
        for f, v in changed.items():
            old = rec.get(f)
            if f in DATE_FIELDS:
                if v is None:
                    ev = derive_excel_value(f, None)
                else:
                    p = pending.get((item_id, f))
                    raw = baseline.get((item_id, f)) or ""
                    if p and p.get("excel_value"):
                        pv = json.loads(p["excel_value"])
                        raw = str(pv.get("value") or "") if pv["kind"] in ("text", "date") else ""
                    if old and f in reschedules:
                        _, hist, ev = compose_date_cell(raw, window, append_iso=v)
                        repo.set_import_history(sheet, item_id, f, hist, now)
                        repo.insert_app_history(sheet, item_id, f, old, reason=reason, root_cause=root_cause, now=now)
                    else:
                        ev = derive_excel_value(f, v)      # first value or a correction of an actual date
            else:
                ev = derive_excel_value(f, v)
            repo.set_field(sheet, item_id, f, v, now)
            repo.append_audit({"ts": now, "editor": editor, "channel": channel, "sheet": sheet,
                               "item_id": item_id, "field": f, "old_value": display(f, old),
                               "new_value": display(f, v),
                               "reason": (reason or "") + (f" [root cause: {root_cause}]" if root_cause and f in reschedules else "")})
            repo.upsert_pending(sheet, item_id, f, value=canon(f, v), excel_value=ev, now=now,
                                reason=reason or "", editor=editor, channel=channel)

    if in_transaction:
        with repo.transaction():
            apply()
    else:
        apply()
    return list(changed)
