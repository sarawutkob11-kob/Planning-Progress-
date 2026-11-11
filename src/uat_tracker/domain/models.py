"""Domain objects shared by import now and by sync / UI / email / reports later."""

from __future__ import annotations

from dataclasses import dataclass, field

# Audit channels (later parts add "excel_sync", "restore").
CHANNEL_IMPORT = "import"
CHANNEL_SYSTEM = "system"      # scheduled job: traffic light changes, emails
CHANNEL_APP = "app"            # PMO action in the app (upload, manual reminder, edits)
CHANNEL_EXCEL_SYNC = "excel_sync"   # Part 3: change read from Excel while syncing
CHANNEL_RESTORE = "restore"         # Part 3: backup restored
EDITED_IN_EXCEL = "Edited in Excel"

# Data-issue flags (rule 7 + approved additions).
FLAG_PASS_FAIL_MISMATCH = "pass_fail_mismatch"
FLAG_MISSING_FAIL = "missing_fail"
FLAG_MISSING_OWNER = "missing_owner"
FLAG_MISSING_REPORT_EMAIL = "missing_report_email"
FLAG_COMPLETE_NO_ACTUAL_FINISH = "complete_no_actual_finish"
FLAG_MISSING_DATA = "missing_data"                    # Part 2: no Traffic Light can be computed
FLAG_ACTUAL_FINISH_LATE = "actual_finish_late"        # Part 2: Actual Finish > Planned Finish

# Fields holding dates (ISO text in storage, DD-Mon-YYYY everywhere they are shown).
DATE_FIELDS = frozenset({
    "start_date", "finish_date",
    "planned_start", "planned_finish", "actual_start", "actual_finish",
    "defect_date", "fix_target_date",
})
HISTORY_SUFFIX = "_history"
NOTE_SUFFIX = "_note"


@dataclass(frozen=True)
class Problem:
    """Something the PMO must decide (review queue)."""
    field: str
    raw: str          # whole cell text as read
    code: str         # ambiguous_date | bad_date | out_of_window | date_in_time_column | unknown_status | bad_number
    token: str        # the offending part of the cell


@dataclass(frozen=True)
class Flag:
    """Data-quality issue shown to the PMO (not a decision)."""
    code: str
    detail: str


@dataclass
class Item:
    """One normalised source row, before it has an app ID."""
    sheet: str
    excel_row: int
    key: tuple                      # primary binding key (occurrence index is added by binder)
    alt_key: tuple | None           # fallback binding key (e.g. group + task no when the name changed)
    id_base: str | None             # preferred new ID, e.g. 'E-23'
    id_blank_prefix: str            # used when there is no number, e.g. 'E-x' -> 'E-x1'
    fields: dict = field(default_factory=dict)     # storage form: str/int/float/None; *_history -> tuple[str]
    raw: dict = field(default_factory=dict)        # field -> raw cell text (sync baseline)
    problems: list[Problem] = field(default_factory=list)
    flags: list[Flag] = field(default_factory=list)
    app_id: str | None = None      # Part 3: value of the hidden 'App ID' column, when present
