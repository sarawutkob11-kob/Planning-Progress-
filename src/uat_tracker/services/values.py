"""Value helpers shared by the importer (pending/conflict checks), edits and the Excel writer.

Excel write-back format (spec §6.5/6.6, approved):
- one date, no history/note -> a real Excel date shown as DD-Mon-YYYY;
- date with reschedule history (or a note) -> text "06-Nov-2026 => 13-Nov-2026" (last = current);
- numbers -> numbers; everything else -> text.
"""

from __future__ import annotations

import datetime as dt
import json

from ..domain.models import DATE_FIELDS
from ..excel.sheets import NUMBER_FIELDS
from ..parsing.dates import DateWindow, fmt, parse_date_cell

DROP = object()          # review resolution: remove a bad token from the cell


def canon(field: str, v) -> str:
    """Comparable text form of a stored value."""
    if v is None or v == "":
        return ""
    if field in DATE_FIELDS:
        return v.isoformat() if isinstance(v, dt.date) else str(v)[:10]
    if field in NUMBER_FIELDS or isinstance(v, (int, float)):
        try:
            f = float(v)
        except (TypeError, ValueError):
            return str(v).strip()
        return str(int(f)) if f.is_integer() else repr(f)
    return str(v).strip()


def typed(field: str, text: str | None):
    """canon() text -> value for the DB column."""
    if text is None or text == "":
        return None
    if field in NUMBER_FIELDS:
        f = float(text)
        return int(f) if f.is_integer() else f
    return text


def excel_value(kind: str, value=None, hist: list[str] | None = None) -> str:
    d = {"kind": kind, "value": value}
    if hist is not None:
        d["hist"] = list(hist)
    return json.dumps(d, ensure_ascii=False)


def derive_excel_value(field: str, value) -> str:
    """Excel value for a plain (non-chain) field edit."""
    if value is None or value == "":
        return excel_value("empty")
    if field in DATE_FIELDS:
        return excel_value("date", canon(field, value), [])
    if field in NUMBER_FIELDS:
        return excel_value("number", typed(field, canon(field, value)))
    return excel_value("text", str(value))


def compose_date_cell(raw: str, window: DateWindow, *, append_iso: str | None = None,
                      replace_token: str | None = None, replacement=None) -> tuple[str | None, list[str], str]:
    """Rebuild a date cell from its raw text.

    append_iso: reschedule -> new current date appended after the existing values.
    replace_token/replacement: review resolution -> that token becomes the ISO date (or DROP).
    Unresolved bad tokens stay as typed; notes in parentheses stay at the end.
    Returns (current ISO or None, history ISO list (valid values only), excel_value JSON).
    """
    parsed = parse_date_cell(raw or None, window)
    parts: list[tuple[str, str | None]] = []          # (display text, iso or None when bad)
    replaced = False
    for t in parsed.tokens:
        if replace_token is not None and not replaced and t.raw == replace_token:
            replaced = True
            if replacement is DROP:
                continue
            parts.append((fmt(replacement), replacement))
        elif t.value is not None:
            parts.append((fmt(t.value), t.value.isoformat()))
        else:
            parts.append((t.raw, None))
    if replace_token is not None and not replaced:
        raise ValueError(f"Token {replace_token!r} not found in cell {raw!r}")
    if append_iso:
        parts.append((fmt(append_iso), append_iso))
    current = parts[-1][1] if parts else None
    hist = [iso for _, iso in parts[:-1] if iso]
    note = parsed.note
    if len(parts) == 1 and current and not note:
        return current, [], excel_value("date", current, [])
    if not parts:
        return None, [], excel_value("text", note) if note else excel_value("empty")
    text = " => ".join(p for p, _ in parts) + (f" {note}" if note else "")
    return current, hist, excel_value("text", text, hist)
