"""Cell text / number helpers."""

from __future__ import annotations

import datetime as dt
import re

BAD_NUMBER = "bad_number"

_WS = re.compile(r"\s+")


def cell_text(v) -> str:
    """Cell value as trimmed text. Non-breaking spaces count as blank. Keeps inner newlines."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float):
        return "%.10g" % v
    if isinstance(v, (dt.datetime, dt.date, dt.time)):
        return v.isoformat()
    s = str(v).replace("\xa0", " ").replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.rstrip() for ln in s.split("\n")]
    return "\n".join(lines).strip()


def collapse(s: str) -> str:
    """Single-line, single-spaced, case-folded form for matching."""
    return _WS.sub(" ", s).strip().casefold()


def is_blank(v) -> bool:
    return cell_text(v) == ""


def task_no_text(v) -> str:
    """Task numbers arrive as int, float (14.1) or text."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return "%.10g" % v
    return cell_text(v)


def parse_number(v) -> tuple[int | float | None, str | None]:
    """Return (number, problem). Blank -> (None, None)."""
    if v is None or (isinstance(v, str) and cell_text(v) == ""):
        return None, None
    if isinstance(v, bool):
        return None, BAD_NUMBER
    if isinstance(v, int):
        return v, None
    if isinstance(v, float):
        return (int(v) if v.is_integer() else v), None
    s = cell_text(v).replace(",", "")
    try:
        f = float(s)
    except ValueError:
        return None, BAD_NUMBER
    return (int(f) if f.is_integer() else f), None
