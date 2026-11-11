"""Status normalisation (rule 6). Unknown values -> PMO review; blank -> 'No status' (None)."""

from __future__ import annotations

import re

from .text import cell_text

UNKNOWN_STATUS = "unknown_status"

TASK_STATUSES = ("Not Start", "In progress", "Blocked", "Complete")
DEFECT_STATUSES = ("Open", "Assigned", "Fixed", "Retest", "Close", "Rejected")

_TASK_ALIASES = {
    "notstart": "Not Start",
    "inprogress": "In progress",
    "blocked": "Blocked",
    "complete": "Complete",
    "completed": "Complete",
    "comlete": "Complete",
}
_DEFECT_ALIASES = {s.lower(): s for s in DEFECT_STATUSES}


def _key(raw) -> str:
    return re.sub(r"\s+", "", cell_text(raw).lower())


def normalize_task_status(raw) -> tuple[str | None, str | None]:
    """UAT Plan / Test Progress. Returns (standard value, problem)."""
    if cell_text(raw) == "":
        return None, None
    value = _TASK_ALIASES.get(_key(raw))
    return (value, None) if value else (None, UNKNOWN_STATUS)


def normalize_defect_status(raw) -> tuple[str | None, str | None]:
    if cell_text(raw) == "":
        return None, None
    value = _DEFECT_ALIASES.get(re.sub(r"\s+", " ", cell_text(raw)).strip().lower())
    return (value, None) if value else (None, UNKNOWN_STATUS)
