"""Strict date-cell parser. Never guesses: anything uncertain becomes a problem for PMO review.

Rules (PART1 parsing rules 1-5 + approved answers):
- CE dates; year > 2500 -> minus 543 (Thai BE).
- A cell may hold several dates (reschedules). The LAST typed value is current; earlier ones are
  history, in typed order. Separators: '=>', '->', whitespace, newlines.
- Text in parentheses and other non-date words -> note.
- Day/month: if one part > 12 the reading is unambiguous. If both <= 12: zero-padded DD/MM/YYYY
  is read as day/month; anything else (e.g. 10/8/2026) is ambiguous -> review.
- Malformed tokens (24/09/206, 06-Now-2026, 30/102026) -> review.
- Every date must fall inside the project window, otherwise -> review (catches 29/09/2027,
  1/4/1900 and dates Excel already swapped).
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

from .text import cell_text

AMBIGUOUS = "ambiguous_date"
BAD = "bad_date"
OUT_OF_WINDOW = "out_of_window"
DATE_IN_TIME = "date_in_time_column"

MON_ABBR = ("", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
_MONTHS = {m.lower(): i for i, m in enumerate(MON_ABBR) if m}
_MONTHS.update({
    "sept": 9, "january": 1, "february": 2, "march": 3, "april": 4, "june": 6, "july": 7,
    "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
})
_EXCEL_TIME_BASE = dt.date(1899, 12, 30)


def fmt(d: dt.date | str | None) -> str:
    """DD-Mon-YYYY (locale independent). Accepts a date or an ISO string."""
    if d is None or d == "":
        return ""
    if isinstance(d, str):
        d = dt.date.fromisoformat(d)
    return f"{d.day:02d}-{MON_ABBR[d.month]}-{d.year}"


@dataclass(frozen=True)
class DateWindow:
    start: dt.date
    end: dt.date

    def __contains__(self, d: dt.date) -> bool:
        return self.start <= d <= self.end


@dataclass(frozen=True)
class Token:
    raw: str
    value: dt.date | None = None
    problem: str | None = None


@dataclass
class ParsedDate:
    raw: str
    tokens: list[Token] = field(default_factory=list)
    note: str | None = None

    @property
    def current(self) -> dt.date | None:
        if self.tokens and self.tokens[-1].problem is None:
            return self.tokens[-1].value
        return None

    @property
    def current_problem(self) -> Token | None:
        if self.tokens and self.tokens[-1].problem is not None:
            return self.tokens[-1]
        return None

    @property
    def history(self) -> list[dt.date]:
        return [t.value for t in self.tokens[:-1] if t.value is not None]

    @property
    def history_problems(self) -> list[Token]:
        return [t for t in self.tokens[:-1] if t.problem is not None]


_SCAN = re.compile(
    r"""
     (?P<paren>\([^)]*\)?)
    |(?P<arrow>=+>|-+>)
    |(?P<named>\d{1,2}[\s\-]*[A-Za-z]{3,9}\.?[\s\-]*\d{1,4}(?![\d/]))
    |(?P<num>\d+(?:[/.\-]\d+)+)
    |(?P<ws>\s+)
    |(?P<word>[^\s()]+)
    """,
    re.VERBOSE,
)


def _be_to_ce(year: int) -> int:
    return year - 543 if year > 2500 else year


def _finish(raw: str, y: int, m: int, d: int, window: DateWindow) -> Token:
    try:
        value = dt.date(_be_to_ce(y), m, d)
    except ValueError:
        return Token(raw, problem=BAD)
    if value not in window:
        return Token(raw, problem=OUT_OF_WINDOW)
    return Token(raw, value=value)


def _parse_numeric(raw: str, window: DateWindow) -> Token:
    parts = re.split(r"[/.\-]", raw)
    if len(parts) != 3:
        return Token(raw, problem=BAD)
    if len(parts[0]) == 4:  # ISO-like YYYY-MM-DD
        y, m, d = parts
        if len(m) > 2 or len(d) > 2:
            return Token(raw, problem=BAD)
        return _finish(raw, int(y), int(m), int(d), window)
    a, b, y = parts
    if len(y) != 4 or len(a) > 2 or len(b) > 2:
        return Token(raw, problem=BAD)
    ia, ib = int(a), int(b)
    if ia > 12 and ib > 12:
        return Token(raw, problem=BAD)
    if ia > 12:
        day, month = ia, ib
    elif ib > 12:
        day, month = ib, ia
    elif ia == ib:
        day = month = ia
    elif len(a) == 2 and len(b) == 2:
        day, month = ia, ib  # zero-padded DD/MM/YYYY (approved rule)
    else:
        readings = [_finish(raw, int(y), mm, dd, window) for dd, mm in ((ia, ib), (ib, ia))]
        if all(t.problem for t in readings):          # neither reading is a usable date
            return Token(raw, problem=readings[0].problem)
        return Token(raw, problem=AMBIGUOUS)
    return _finish(raw, int(y), month, day, window)


def _parse_named(raw: str, window: DateWindow) -> Token:
    m = re.fullmatch(r"(\d{1,2})[\s\-]*([A-Za-z]{3,9})\.?[\s\-]*(\d{1,4})", raw.strip())
    if not m:
        return Token(raw, problem=BAD)
    day, mon, year = m.groups()
    month = _MONTHS.get(mon.lower())
    if month is None or len(year) != 4:
        return Token(raw, problem=BAD)
    return _finish(raw, int(year), month, int(day), window)


def parse_date_cell(value, window: DateWindow) -> ParsedDate:
    if value is None:
        return ParsedDate(raw="")
    if isinstance(value, dt.datetime) or (isinstance(value, dt.date)):
        d = value.date() if isinstance(value, dt.datetime) else value
        raw = d.isoformat()
        return ParsedDate(raw=raw, tokens=[_finish(raw, d.year, d.month, d.day, window)])
    if isinstance(value, (int, float, dt.time)) and not isinstance(value, bool):
        raw = cell_text(value)
        return ParsedDate(raw=raw, tokens=[Token(raw, problem=BAD)])

    raw = cell_text(value)
    tokens: list[Token] = []
    notes: list[str] = []
    for m in _SCAN.finditer(raw):
        kind, text = m.lastgroup, m.group()
        if kind == "paren":
            notes.append(text.strip())
        elif kind == "named":
            tokens.append(_parse_named(text, window))
        elif kind == "num":
            tokens.append(_parse_numeric(text, window))
        elif kind == "word":
            if any(c.isdigit() for c in text):
                tokens.append(Token(text, problem=BAD))
            else:
                notes.append(text)
    note = " ".join(notes).strip() or None
    return ParsedDate(raw=raw, tokens=tokens, note=note)


def time_cell_problem(value) -> tuple[str, str | None]:
    """Time columns: return (raw text, problem). A date in a time column goes to review, unmoved."""
    if value is None:
        return "", None
    if isinstance(value, dt.time):
        return value.strftime("%H:%M"), None
    if isinstance(value, dt.datetime):
        if value.date() == _EXCEL_TIME_BASE:
            return value.strftime("%H:%M"), None
        return value.date().isoformat(), DATE_IN_TIME
    if isinstance(value, dt.date):
        return value.isoformat(), DATE_IN_TIME
    raw = cell_text(value)
    for m in _SCAN.finditer(raw):
        if m.lastgroup in ("named", "num") and len(re.split(r"[/.\-\s]+", m.group().strip())) >= 3:
            return raw, DATE_IN_TIME
    return raw, None
