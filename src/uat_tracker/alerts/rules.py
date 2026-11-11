"""R1-R4 for one Test Progress item on one workday (PART2 §D + approved answers). Calculation only.

- Status Complete stops every rule (6f).
- R1/R2: on the (shifted) trigger day one "today" email, then a "daily" email from the next workday
  while the actual date is empty (6g). Trigger on a non-workday -> previous workday (6h).
- R2 "today" is not sent when Actual Finish is already filled (nothing left to remind).
- R3: passed in by the job (needs the color history + email log).
- R4: Executed = Total TC (> 0) and Actual Finish filled; sent even when Actual Finish > Planned Finish
  (6i, the late finish is flagged separately). R4A = set Status; R4B = upload test results, stops once
  TC/Result Complete is Complete or a file is uploaded (6j).
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from ..services.workdays import WorkCalendar

R1, R2, R3, R4A, R4B = "R1", "R2", "R3", "R4A", "R4B"
TODAY, DAILY = "today", "daily"


@dataclass(frozen=True)
class R3Event:
    history_id: int
    color: str
    prev_color: str | None
    reason: str


@dataclass
class Alert:
    rule: str
    variant: str | None
    item_id: str
    trigger: dt.date | None = None          # the planned date the rule is about
    shifted: bool = False                   # trigger fell on a non-workday
    event: R3Event | None = None
    extra: dict = field(default_factory=dict)


def _d(v) -> dt.date | None:
    return dt.date.fromisoformat(v) if v else None


def _date_rule(rule: str, item_id: str, planned: dt.date | None, actual_filled: bool,
               today: dt.date, cal: WorkCalendar, send_today_if_filled: bool) -> Alert | None:
    if planned is None:
        return None
    eff = cal.on_or_before(planned)
    if today == eff:
        if actual_filled and not send_today_if_filled:
            return None
        return Alert(rule, TODAY, item_id, planned, shifted=eff != planned)
    if today > eff and not actual_filled:
        return Alert(rule, DAILY, item_id, planned)
    return None


def is_complete_text(v) -> bool:
    return str(v or "").strip().lower() in ("complete", "completed")


def evaluate(rec: dict, today: dt.date, cal: WorkCalendar, *, r3_event: R3Event | None = None,
             has_upload: bool = False) -> list[Alert]:
    if not rec.get("active") or rec.get("status") == "Complete":
        return []
    item_id = rec["id"]
    out: list[Alert] = []

    a = _date_rule(R1, item_id, _d(rec.get("planned_start")), bool(rec.get("actual_start")),
                   today, cal, send_today_if_filled=True)
    if a:
        out.append(a)
    a = _date_rule(R2, item_id, _d(rec.get("planned_finish")), bool(rec.get("actual_finish")),
                   today, cal, send_today_if_filled=False)
    if a:
        out.append(a)

    if r3_event is not None:
        out.append(Alert(R3, None, item_id, event=r3_event))

    total, executed = rec.get("total_tc"), rec.get("executed")
    if total and executed == total and rec.get("actual_finish"):
        out.append(Alert(R4A, DAILY, item_id))
        if not is_complete_text(rec.get("tc_result_complete")) and not has_upload:
            out.append(Alert(R4B, DAILY, item_id))
    return out
