"""Suggestions for PMO review items (Conflict Center, approved 09-Nov-2026).

The app still never decides: these are one-click choices the PMO may pick. Only values inside the
project date window are suggested.
- ambiguous date      -> each reading (day/month and month/day) that is inside the window;
- out of window       -> day/month swapped (dates Excel already swapped), year fixed to a window year;
- typo                -> 3-digit year completed (206 -> 2026), missing separator (30/102026), month
                         name typo (06-Now-2026 -> Nov);
- a part of a cell that has other values -> "remove this part";
- date in a time column -> "clear the time cell";
- unknown status      -> the closest standard status.
"""

from __future__ import annotations

import datetime as dt
import difflib
import re

from ..domain.models import DATE_FIELDS, HISTORY_SUFFIX
from ..parsing.dates import MON_ABBR, DateWindow, parse_date_cell
from ..parsing.status import DEFECT_STATUSES, TASK_STATUSES

MONTHS = [m for m in MON_ABBR if m]


def _mk(y, m, d) -> dt.date | None:
    try:
        return dt.date(y - 543 if y > 2500 else y, m, d)
    except ValueError:
        return None


def _near(y: int, window: DateWindow) -> bool:
    y = y - 543 if y > 2500 else y
    return window.start.year - 1 <= y <= window.end.year + 1


def _date_candidates(token: str, window: DateWindow) -> list[tuple[str, str]]:
    """[(why, iso)] for a bad / ambiguous / out-of-window date token.
    Same-year readings (both orders) first; a changed year only when those give nothing and the
    typed year is within one year of the project; obvious typos are always tried."""
    tok = token.strip()
    same: list[tuple[str, dt.date | None]] = []
    years: list[tuple[str, dt.date | None]] = []
    typos: list[tuple[str, dt.date | None]] = []
    wy = range(window.start.year, window.end.year + 1)
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", tok)          # real Excel date (ISO text)
    if m:
        y, mo, d = map(int, m.groups())
        same.append(("swap", _mk(y, d, mo)))
        if _near(y, window):
            years += [("year", _mk(yy, mo, d)) for yy in wy]
    m = re.fullmatch(r"(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{1,4})", tok)
    if m:
        a, b, y = m.groups()
        a, b = int(a), int(b)
        if len(y) == 3:                                          # 206 -> 2026
            for yy in (int(y[:2] + "2" + y[2:]), int("2" + y)):
                typos += [("year", _mk(yy, b, a)), ("year", _mk(yy, a, b))]
        else:
            same += [("both", _mk(int(y), b, a)), ("both", _mk(int(y), a, b))]
            if _near(int(y), window):
                for yy in wy:
                    years += [("year", _mk(yy, b, a)), ("year", _mk(yy, a, b))]
    m = re.fullmatch(r"(\d{1,2})[/.\-](\d{1,2})(\d{4})", tok)    # 30/102026
    if m:
        a, b, y = map(int, m.groups())
        typos += [("year", _mk(y, b, a)), ("year", _mk(y, a, b))]
    m = re.fullmatch(r"(\d{1,2})[\s\-]*([A-Za-z]{3,9})\.?[\s\-]*(\d{4})", tok)   # 06-Now-2026
    if m:
        d, name, y = m.groups()
        close = difflib.get_close_matches(name[:3].title(), MONTHS, n=1, cutoff=0.5)
        if close:
            typos.append(("month", _mk(int(y), MONTHS.index(close[0]) + 1, int(d))))

    def keep(lst):
        seen, res = set(), []
        for why, d in lst:
            if d and d in window and d.isoformat() not in seen:
                seen.add(d.isoformat())
                res.append((why, d.isoformat()))
        return res

    first = keep(same + typos)
    return first or keep(years)


def suggestions(cfg, review: dict) -> list[dict]:
    """[{"why": swap|both|year|month|inwin|drop|clear|status, "value": ...}] for one review item."""
    window = DateWindow(cfg.window_start, cfg.window_end)
    field, problem, token = review["field"], review["problem"], review.get("token") or ""
    base = field[: -len(HISTORY_SUFFIX)] if field.endswith(HISTORY_SUFFIX) else field
    out: list[dict] = []
    if problem == "unknown_status":
        allowed = DEFECT_STATUSES if review["sheet"] == "Defect List" else TASK_STATUSES
        close = difflib.get_close_matches(token.strip().title(), list(allowed), n=2, cutoff=0.4)
        return [{"why": "status", "value": c} for c in close]
    if problem == "date_in_time_column":
        return [{"why": "clear", "value": ""}]
    if base in DATE_FIELDS:
        cands = _date_candidates(token, window)
        if problem == "ambiguous_date" and len(cands) == 1:
            cands = [("inwin", cands[0][1])]
        elif problem == "out_of_window":
            cands = [("swap" if why == "both" else why, v) for why, v in cands]
        elif problem == "ambiguous_date":
            cands = [c for c in cands if c[0] == "both"]
        out = [{"why": why, "value": v} for why, v in cands]
        parsed = parse_date_cell(review.get("raw_value") or None, window)
        if len(parsed.tokens) > 1:                               # other values remain -> can drop this one
            out.append({"why": "drop", "value": ""})
    return out
