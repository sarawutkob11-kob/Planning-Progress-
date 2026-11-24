"""The three source sheets and their columns (by header name)."""

from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import dataclass

from ..parsing.headers import FieldSpec
from ..parsing.text import cell_text, collapse, task_no_text

UAT_PLAN = "UAT Plan"
TEST_PROGRESS = "Test Progress"
DEFECT_LIST = "Defect List"


@dataclass(frozen=True)
class SheetDef:
    name: str
    fields: tuple[FieldSpec, ...]
    blank_first_as: str | None = None


F = FieldSpec

SHEETS: dict[str, SheetDef] = {
    UAT_PLAN: SheetDef(UAT_PLAN, (
        F("sub_task", ("Sub Task",), required=False),
        F("group", ("Group",)),
        F("task_no", ("Task No",)),
        F("task_name", ("Task name",)),
        F("note", ("Note",), required=False),
        F("file_path", ("File Path",), required=False),
        F("responsible", ("Responsible name", "Responsible")),
        F("status", ("Status",)),
        F("start_date", ("Start date",)),
        F("start_time", ("Start time",), required=False),
        F("finish_date", ("Finish date",)),
        F("finish_time", ("Finish time",), required=False),
        F("remark", ("Remark",), required=False),
        F("app_id", ("App ID",), required=False),
    )),
    TEST_PROGRESS: SheetDef(TEST_PROGRESS, (
        F("no", ("No.", "No")),
        F("application", ("Application",)),
        F("report_by", ("Report By",)),
        F("report_by_email", ("Report By Email",), required=False),
        F("cc_email", ("CC Email",), required=False),
        F("total_tc", ("Total TC",)),
        F("executed", ("No. of TC Execution",)),
        F("pct_progress", ("%Progress",), required=False),
        F("pass_count", ("#Pass",)),
        F("fail_count", ("#Fail",)),
        F("planned_start", ("Planned Start",)),
        F("planned_finish", ("Planned Finish",)),
        F("actual_start", ("Actual Start",)),
        F("actual_finish", ("Actual Finish",)),
        F("traffic_light", ("Traffic Light",), required=False),
        F("tc_result_complete", ("TC/Result Complete",), required=False),
        F("status", ("Status",)),
        F("remark", ("Remark",), required=False),
        F("app_id", ("App ID",), required=False),
    )),
    DEFECT_LIST: SheetDef(DEFECT_LIST, (
        F("no", ("No.", "No")),
        F("defect_date", ("Date",)),
        F("system", ("System/App.", "System/App")),
        F("detail", ("Issues/detail",)),
        F("status", ("Status",)),
        F("severity", ("Severity",)),
        F("assigned_to", ("Assigned To",)),
        F("fix_target_date", ("Fix target date",)),
        F("data_test", ("Data test",), required=False),
        F("raised_by", ("Raised by",)),
        F("remark", ("Remark",), required=False),
        F("app_id", ("App ID",), required=False),
    ), blank_first_as="No."),
}

APP_ID_KEY = "app_id"
APP_ID_HEADER = "App ID"

# --------------------------------------------------------------------------- which rows are items
# One set of rules for the importer (services.normalize) and the writer's moved-row check.
# `values`: field key -> raw cell value of one row. The hidden App ID column never counts: a row that holds only
# an App ID (or only No. + App ID on Defect List) is not an item, for the importer and the writer alike.

_GROUP_LETTER = re.compile(r"[A-Za-z]")


def _content(values: Mapping[str, object]) -> dict[str, object]:
    """The row's values without the App ID column."""
    return {k: v for k, v in values.items() if k != APP_ID_KEY}


def is_group_header(values: Mapping[str, object]) -> bool:
    """UAT Plan group header row, e.g. 'A' + 'SETUP...' with no Task No."""
    return bool(_GROUP_LETTER.fullmatch(cell_text(values.get("group")))) and not task_no_text(values.get("task_no"))


def is_test_progress_item(values: Mapping[str, object]) -> bool:
    """Not the Grand Total row and not a row without both Application and No."""
    app = cell_text(values.get("application"))
    return collapse(app) != "grand total" and bool(app or task_no_text(values.get("no")))


def is_defect_item(values: Mapping[str, object]) -> bool:
    """A pre-numbered row (only 'No.' filled, App ID not counted) is not a defect yet."""
    return any(cell_text(v) != "" for k, v in _content(values).items() if k != "no")


def is_item_row(sheet: str, values: Mapping[str, object]) -> bool:
    """Whether a row can hold an item of `sheet` (UAT Plan: rows above group A are not seen here).
    App ID is not counted, so a row with only an App ID left in it is not an item row."""
    values = _content(values)
    if all(cell_text(v) == "" for v in values.values()):
        return False
    if sheet == UAT_PLAN:
        return not is_group_header(values)
    if sheet == TEST_PROGRESS:
        return is_test_progress_item(values)
    if sheet == DEFECT_LIST:
        return is_defect_item(values)
    return True


# Part 3: fields the app may write back to Excel (DB field name). Formula cells are never written.
WRITABLE: dict[str, frozenset[str]] = {
    UAT_PLAN: frozenset({"task_no", "sub_task", "task_name", "note", "file_path", "responsible", "status",
                         "start_date", "start_time", "finish_date", "finish_time", "remark"}),
    TEST_PROGRESS: frozenset({"application", "report_by", "report_by_email", "cc_emails", "total_tc",
                              "executed", "pass_count", "fail_count", "planned_start", "planned_finish",
                              "actual_start", "actual_finish", "traffic_light", "tc_result_complete",
                              "status", "remark"}),
    DEFECT_LIST: frozenset({"defect_date", "system", "detail", "status", "severity", "assigned_to",
                            "fix_target_date", "data_test", "raised_by", "remark"}),
}
NUMBER_FIELDS = frozenset({"total_tc", "executed", "pct_progress", "pass_count", "fail_count"})
TIME_FIELDS = frozenset({"start_time", "finish_time"})
# DB field name -> sheet column key, where they differ.
DB_TO_SHEET_KEY = {"cc_emails": "cc_email"}


def with_overrides(sheet: SheetDef, overrides: dict) -> SheetDef:
    """Apply [headers."<sheet>"] key = ["Name", ...] from config."""
    ov = overrides.get(sheet.name, {})
    if not ov:
        return sheet
    fields = tuple(
        FieldSpec(f.key, tuple(ov[f.key]), f.required) if f.key in ov else f for f in sheet.fields
    )
    return SheetDef(sheet.name, fields, sheet.blank_first_as)
