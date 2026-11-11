"""Traffic Light for Test Progress (PART2 §A). Calculation only; `record_colors` stores changes.

%Exec = Executed / Total TC
%Time = (today - Planned Start) / (Planned Finish - Planned Start)   (calendar days)
%Fail = #Fail / Executed
First match wins:
  1. Status Complete                                    -> Complete
  2. today > Planned Finish (not Complete) or %Fail > X -> Red
  3. %Exec < %Time, or #Fail >= 1 and %Fail <= X        -> Amber
  4. otherwise                                          -> Green
Approved gap rules: blank #Fail = 0; Executed 0 -> %Fail 0; blank Executed = 0; Total TC blank/0 or a
planned date blank (or waiting in the review queue, which leaves it blank) -> no color ("Missing").
Planned Start = Planned Finish -> %Time is 1 from that day on.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass

from ..domain.models import CHANNEL_SYSTEM
from ..excel.sheets import TEST_PROGRESS

COMPLETE, RED, AMBER, GREEN, MISSING = "Complete", "Red", "Amber", "Green", "Missing"


@dataclass(frozen=True)
class Light:
    color: str
    pct_exec: float | None = None
    pct_time: float | None = None
    pct_fail: float | None = None
    reason: str = ""


def _d(v) -> dt.date | None:
    if v is None or v == "":
        return None
    return v if isinstance(v, dt.date) else dt.date.fromisoformat(str(v))


def _pct(x: float | None) -> str:
    return "-" if x is None else f"{x * 100:.1f}%"


def compute(rec: dict, today: dt.date, fail_threshold: float) -> Light:
    if rec.get("status") == "Complete":
        return Light(COMPLETE, reason="Status is Complete")
    total = rec.get("total_tc")
    ps, pf = _d(rec.get("planned_start")), _d(rec.get("planned_finish"))
    missing = [n for n, v in (("Total TC", total), ("Planned Start", ps), ("Planned Finish", pf)) if not v]
    if missing:
        return Light(MISSING, reason="Missing data: " + ", ".join(missing))
    executed = rec.get("executed") or 0
    failed = rec.get("fail_count") or 0
    pct_exec = executed / total
    if pf == ps:
        pct_time = 1.0 if today >= ps else 0.0
    else:
        pct_time = (today - ps).days / (pf - ps).days
    pct_fail = failed / executed if executed else 0.0
    nums = dict(pct_exec=pct_exec, pct_time=pct_time, pct_fail=pct_fail)
    x = fail_threshold

    if today > pf:
        return Light(RED, **nums, reason=f"Past Planned Finish ({pf:%d-%b-%Y}) and not Complete")
    if pct_fail > x:
        return Light(RED, **nums, reason=f"%Fail {_pct(pct_fail)} > {_pct(x)}")
    if pct_exec < pct_time:
        why = f"%Exec {_pct(pct_exec)} < %Time {_pct(pct_time)}"
        if failed >= 1:
            why += f"; #Fail {failed:g} ({_pct(pct_fail)} <= {_pct(x)})"
        return Light(AMBER, **nums, reason=why)
    if failed >= 1:
        return Light(AMBER, **nums, reason=f"#Fail {failed:g} ({_pct(pct_fail)} <= {_pct(x)})")
    return Light(GREEN, **nums, reason=f"%Exec {_pct(pct_exec)} >= %Time {_pct(pct_time)}, no fails")


# --------------------------------------------------------------------------- change detection

def effective_previous(history: list[dict]) -> str | None:
    """Color before the latest row, ignoring 'Missing' (a data gap is not a color)."""
    for row in reversed(history[:-1]):
        if row["color"] != MISSING:
            return row["color"]
    return None


def r3_qualifies(new_color: str, prev_color: str | None) -> bool:
    """Change INTO Amber/Red: Green/none->Amber, Green/Amber/none->Red, re-entry after Green.
    Red->Amber: no email (approved)."""
    if new_color == RED:
        return prev_color != RED
    if new_color == AMBER:
        return prev_color not in (AMBER, RED)
    return False


@dataclass
class ColorChange:
    item_id: str
    old: str | None
    new: str
    history_id: int


def record_colors(repo, items: dict[str, dict], today: dt.date, fail_threshold: float, *,
                  now: str, editor: str, job_run_id: int | None) -> tuple[dict[str, Light], list[ColorChange]]:
    """Compute colors for active Test Progress items; store a history row + audit entry on change."""
    latest = repo.latest_lights()
    lights, changes = {}, []
    for item_id, rec in sorted(items.items()):
        if not rec.get("active"):
            continue
        light = compute(rec, today, fail_threshold)
        lights[item_id] = light
        prev = latest.get(item_id)
        if prev is not None and prev["color"] == light.color:
            continue
        old = prev["color"] if prev else None
        hid = repo.insert_light({
            "item_id": item_id, "color": light.color, "prev_color": old,
            "pct_exec": light.pct_exec, "pct_time": light.pct_time, "pct_fail": light.pct_fail,
            "reason": light.reason, "calc_date": today.isoformat(), "ts": now, "job_run_id": job_run_id,
        })
        repo.append_audit({"ts": now, "editor": editor, "channel": CHANNEL_SYSTEM, "sheet": TEST_PROGRESS,
                           "item_id": item_id, "field": "traffic_light (calculated)",
                           "old_value": old or "", "new_value": light.color, "reason": light.reason})
        changes.append(ColorChange(item_id, old, light.color, hid))
    return lights, changes
