"""PMO decisions: PMO review queue (Part 1) and Conflict Center (Part 3).

Review item resolved -> the clean value goes to the DB, the audit log and the write queue, so the
Excel cell is rewritten in DD-Mon-YYYY (approved). Unresolved bad parts of the same cell stay as typed.

Conflict resolved (reason required) ->
  excel: the Excel value wins (the app's queued value is dropped; the next read applies Excel);
  app:   the app value wins and is written to Excel;
  new:   the PMO's new value goes to both sides.
"""

from __future__ import annotations

from datetime import datetime

from ..domain.models import CHANNEL_APP, DATE_FIELDS, HISTORY_SUFFIX
from ..excel.sheets import DEFECT_LIST, TIME_FIELDS
from ..parsing.dates import DateWindow, fmt
from ..parsing.status import normalize_defect_status, normalize_task_status
from ..parsing.text import parse_number
from .auth import RESOLVE, Actor, require
from .edits import EditError, edit_item, validate_value
from .importer import display
from .values import DROP, canon, compose_date_cell, derive_excel_value, typed


class ResolveError(Exception):
    pass


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _audit(repo, now, editor, sheet, item_id, field, old, new, reason):
    repo.append_audit({"ts": now, "editor": editor, "channel": CHANNEL_APP, "sheet": sheet,
                       "item_id": item_id, "field": field, "old_value": old, "new_value": new,
                       "reason": reason})


# --------------------------------------------------------------------------- review queue

def resolve_review(cfg, repo, review_id: int, value: str | None, *, actor: Actor, reason: str = "",
                   now: str | None = None) -> str:
    """value: ISO date for date problems ('' / None drops that part of the cell), the status, the
    number, or the time-cell text. Returns the value stored."""
    now = now or _now()
    require(repo, actor, RESOLVE, item_id=f"review #{review_id}")
    editor = actor.editor
    r = repo.get_review(review_id)
    if r is None or r["status"] != "open":
        raise ResolveError(f"Review #{review_id} is not open")
    sheet, item_id, field = r["sheet"], r["item_id"], r["field"]
    rec = repo.get_item(sheet, item_id)
    if rec is None or not rec.get("active"):
        raise ResolveError(f"{item_id} is no longer active")
    base = field[: -len(HISTORY_SUFFIX)] if field.endswith(HISTORY_SUFFIX) else field
    window = DateWindow(cfg.window_start, cfg.window_end)
    label = f"Review #{review_id} ({r['problem']} {r['token']!r}) resolved" + (f": {reason}" if reason else "")

    if base in DATE_FIELDS:
        if value:
            try:
                iso = datetime.fromisoformat(value[:10]).date().isoformat()
            except ValueError:
                raise ResolveError(f"Not a date: {value!r} (use YYYY-MM-DD)")
            if not (cfg.window_start.isoformat() <= iso <= cfg.window_end.isoformat()):
                raise ResolveError(f"{iso} is outside the project date window ({cfg.window_start} .. "
                                   f"{cfg.window_end}); change the window in config.toml first")
            replacement = iso
        else:
            replacement = DROP
        raw = repo.load_baseline(sheet).get((item_id, base)) or r["raw_value"] or ""
        try:
            current, hist, ev = compose_date_cell(raw, window, replace_token=r["token"], replacement=replacement)
        except ValueError as exc:
            raise ResolveError(str(exc))
        stored, store_field, store_val = (replacement if replacement is not DROP else ""), base, current
    elif base in TIME_FIELDS:
        stored, store_field, store_val, ev = (value or ""), base, (value or None), derive_excel_value(base, value)
    elif base == "status":
        status, problem = (normalize_defect_status if sheet == DEFECT_LIST else normalize_task_status)(value)
        if problem or status is None:
            raise ResolveError(f"Status {value!r} is not a standard value")
        stored, store_field, store_val, ev = status, base, status, derive_excel_value(base, status)
    else:
        n, problem = parse_number(value)
        if problem or n is None:
            raise ResolveError(f"Not a number: {value!r}")
        stored, store_field, store_val, ev = canon(base, n), base, n, derive_excel_value(base, n)

    with repo.transaction():
        old = rec.get(store_field)
        repo.set_field(sheet, item_id, store_field, store_val, now)
        if base in DATE_FIELDS:
            repo.set_import_history(sheet, item_id, base, hist, now)
        _audit(repo, now, editor, sheet, item_id, store_field, display(store_field, old),
               display(store_field, store_val), label)
        repo.upsert_pending(sheet, item_id, store_field, value=canon(store_field, store_val), excel_value=ev,
                            now=now, reason=label, editor=editor, channel=CHANNEL_APP)
        repo.resolve_review(review_id, value=stored, by=editor, at=now)
    return stored


# --------------------------------------------------------------------------- conflicts

def resolve_conflict(cfg, repo, conflict_id: int, choice: str, *, actor: Actor, reason: str,
                     new_value=None, now: str | None = None) -> None:
    require(repo, actor, RESOLVE, item_id=f"conflict #{conflict_id}")
    editor = actor.editor
    now = now or _now()
    if not (reason and reason.strip()):
        raise ResolveError("A reason is required to resolve a conflict")
    c = repo.get_conflict(conflict_id)
    if c is None or c["status"] != "open":
        raise ResolveError(f"Conflict #{conflict_id} is not open")
    sheet, item_id, field = c["sheet"], c["item_id"], c["field"]
    label = f"Conflict #{conflict_id} resolved ({choice}): {reason}"

    if choice == "excel":
        with repo.transaction():
            repo.clear_pending(sheet, item_id, field)
            repo.resolve_conflict(conflict_id, resolution="excel", resolved_value=c["excel_value"],
                                  reason=reason, resolved_by=editor, resolved_at=now)
            _audit(repo, now, editor, sheet, item_id, field, display(field, c["app_value"]),
                   display(field, c["excel_value"]), label)
        return                                           # the next read applies the Excel value

    if choice == "app":
        value = typed(field, c["app_value"]) if c["app_value"] else None
        with repo.transaction():
            p = repo.pending(sheet).get((item_id, field))
            if p is not None:
                repo.set_pending_state(sheet, item_id, field, status="pending", attempts=0)
            else:
                repo.upsert_pending(sheet, item_id, field, value=c["app_value"] or "",
                                    excel_value=derive_excel_value(field, value), now=now, reason=label,
                                    editor=editor, channel=CHANNEL_APP)
            repo.resolve_conflict(conflict_id, resolution="app", resolved_value=c["app_value"],
                                  reason=reason, resolved_by=editor, resolved_at=now)
            _audit(repo, now, editor, sheet, item_id, field, display(field, c["excel_value"]),
                   display(field, c["app_value"]), label)
        return

    if choice == "new":
        try:
            v = validate_value(sheet, field, new_value)
        except EditError as exc:
            raise ResolveError(str(exc))
        with repo.transaction():
            repo.clear_pending(sheet, item_id, field)
            repo.resolve_conflict(conflict_id, resolution="new", resolved_value=canon(field, v),
                                  reason=reason, resolved_by=editor, resolved_at=now)
            try:
                changed = edit_item(cfg, repo, sheet, item_id, {field: v}, actor=actor, reason=label,
                                    now=now, in_transaction=False)
            except EditError as exc:
                raise ResolveError(str(exc))
            if not changed:                              # same as the DB value: still write it to Excel
                repo.upsert_pending(sheet, item_id, field, value=canon(field, v),
                                    excel_value=derive_excel_value(field, v), now=now, reason=label,
                                    editor=editor, channel=CHANNEL_APP)
                _audit(repo, now, editor, sheet, item_id, field, display(field, c["excel_value"]),
                       display(field, v), label)
        return
    raise ResolveError("Choice must be excel, app or new")


def describe_conflict(c: dict) -> str:
    f = c["field"]
    ev = fmt(c["excel_value"]) if f in DATE_FIELDS and c["excel_value"] else (c["excel_value"] or "(empty)")
    av = fmt(c["app_value"]) if f in DATE_FIELDS and c["app_value"] else (c["app_value"] or "(empty)")
    return f"#{c['id']} {c['sheet']} {c['item_id']} {f}: Excel={ev} vs app={av} [{c['kind']}]"
