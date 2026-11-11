"""Raw sheet rows -> normalised Items (+ review problems + data flags)."""

from __future__ import annotations

import re

from ..domain.models import (
    FLAG_ACTUAL_FINISH_LATE, FLAG_COMPLETE_NO_ACTUAL_FINISH, FLAG_MISSING_DATA, FLAG_MISSING_FAIL, FLAG_MISSING_OWNER,
    FLAG_MISSING_REPORT_EMAIL, FLAG_PASS_FAIL_MISMATCH, HISTORY_SUFFIX, NOTE_SUFFIX,
    Flag, Item, Problem,
)
from ..excel.reader import RawRow
from ..excel.sheets import DEFECT_LIST, TEST_PROGRESS, UAT_PLAN
from ..parsing.dates import DateWindow, fmt, parse_date_cell, time_cell_problem
from ..parsing.status import normalize_defect_status, normalize_task_status
from ..parsing.text import cell_text, collapse, parse_number, task_no_text


def _text(item: Item, key: str, value) -> str | None:
    t = cell_text(value)
    item.raw[key] = t
    item.fields[key] = t or None
    return t or None


def _date(item: Item, key: str, value, window: DateWindow) -> None:
    p = parse_date_cell(value, window)
    item.raw[key] = p.raw
    item.fields[key] = p.current.isoformat() if p.current else None
    item.fields[key + HISTORY_SUFFIX] = tuple(d.isoformat() for d in p.history)
    item.fields[key + NOTE_SUFFIX] = p.note
    if (cp := p.current_problem) is not None:
        item.problems.append(Problem(key, p.raw, cp.problem, cp.raw))
    for t in p.history_problems:
        item.problems.append(Problem(key + HISTORY_SUFFIX, p.raw, t.problem, t.raw))


def _time(item: Item, key: str, value) -> None:
    raw, problem = time_cell_problem(value)
    item.raw[key] = raw
    item.fields[key] = raw or None
    if problem:
        item.problems.append(Problem(key, raw, problem, raw))


def _number(item: Item, key: str, value):
    n, problem = parse_number(value)
    item.raw[key] = cell_text(value)
    item.fields[key] = n
    if problem:
        item.problems.append(Problem(key, cell_text(value), problem, cell_text(value)))
    return n


def _status(item: Item, value, normalizer) -> str | None:
    raw = cell_text(value)
    item.raw["status"] = raw
    status, problem = normalizer(value)
    item.fields["status"] = status
    if problem:
        item.problems.append(Problem("status", raw, problem, raw))
    return status


# --------------------------------------------------------------------------- UAT Plan

_GROUP_LETTER = re.compile(r"[A-Za-z]")


def normalize_uat_plan(rows: list[RawRow], window: DateWindow) -> list[Item]:
    items: list[Item] = []
    group: str | None = None
    title = ""
    for r in rows:
        v = r.values
        g = cell_text(v.get("group"))
        tno = task_no_text(v.get("task_no"))
        if _GROUP_LETTER.fullmatch(g) and not tno:          # group header row, e.g. 'A' + 'SETUP...'
            group, title = g.upper(), cell_text(v.get("task_name"))
            continue
        if group is None:                                   # rows above group A (Pre-Implement Activity)
            continue
        name = cell_text(v.get("task_name"))
        item = Item(
            sheet=UAT_PLAN, excel_row=r.excel_row,
            key=(group, tno, collapse(name)),
            alt_key=(group, tno) if tno else None,
            id_base=f"{group}-{tno}" if tno else None,
            id_blank_prefix=f"{group}-x",
        )
        f = item.fields
        f["group_code"], f["group_title"] = group, title or None
        item.raw["task_no"] = tno
        f["task_no"] = tno or None
        _text(item, "sub_task", v.get("sub_task"))
        _text(item, "task_name", v.get("task_name"))
        _text(item, "note", v.get("note"))
        _text(item, "file_path", v.get("file_path"))
        owner = _text(item, "responsible", v.get("responsible"))
        _status(item, v.get("status"), normalize_task_status)
        _date(item, "start_date", v.get("start_date"), window)
        _time(item, "start_time", v.get("start_time"))
        _date(item, "finish_date", v.get("finish_date"), window)
        _time(item, "finish_time", v.get("finish_time"))
        _text(item, "remark", v.get("remark"))
        if not owner:
            item.flags.append(Flag(FLAG_MISSING_OWNER, "Responsible name is empty"))
        items.append(item)
    return items


# --------------------------------------------------------------------------- Test Progress

def normalize_test_progress(rows: list[RawRow], window: DateWindow) -> list[Item]:
    items: list[Item] = []
    for r in rows:
        v = r.values
        app = cell_text(v.get("application"))
        no = task_no_text(v.get("no"))
        if collapse(app) == "grand total" or (not app and not no):
            continue
        item = Item(
            sheet=TEST_PROGRESS, excel_row=r.excel_row,
            key=(collapse(app),),
            alt_key=(no,) if no else None,
            id_base=f"TP-{no}" if no else None,
            id_blank_prefix="TP-x",
        )
        item.raw["no"] = no
        item.fields["no"] = no or None
        _text(item, "application", v.get("application"))
        _text(item, "report_by", v.get("report_by"))
        email = _text(item, "report_by_email", v.get("report_by_email"))
        cc_raw = cell_text(v.get("cc_email"))
        item.raw["cc_emails"] = cc_raw
        cc = [p.strip() for p in cc_raw.split(";") if p.strip()]
        item.fields["cc_emails"] = "; ".join(cc) or None
        _number(item, "total_tc", v.get("total_tc"))
        executed = _number(item, "executed", v.get("executed"))
        _number(item, "pct_progress", v.get("pct_progress"))
        passed = _number(item, "pass_count", v.get("pass_count"))
        failed = _number(item, "fail_count", v.get("fail_count"))
        for key in ("planned_start", "planned_finish", "actual_start", "actual_finish"):
            _date(item, key, v.get(key), window)
        _text(item, "traffic_light", v.get("traffic_light"))
        _text(item, "tc_result_complete", v.get("tc_result_complete"))
        status = _status(item, v.get("status"), normalize_task_status)
        _text(item, "remark", v.get("remark"))

        if executed is not None and (passed or 0) + (failed or 0) != executed:
            item.flags.append(Flag(FLAG_PASS_FAIL_MISMATCH,
                                   f"#Pass {passed or 0} + #Fail {failed or 0} != Executed {executed}"))
        if failed is None and (executed or 0) > 0:
            item.flags.append(Flag(FLAG_MISSING_FAIL, "#Fail is empty (counted as 0)"))
        if not email:
            item.flags.append(Flag(FLAG_MISSING_REPORT_EMAIL, "Report By Email is empty"))
        if status == "Complete" and not item.raw["actual_finish"]:
            item.flags.append(Flag(FLAG_COMPLETE_NO_ACTUAL_FINISH, "Status Complete but Actual Finish is empty"))
        f = item.fields
        if status != "Complete":
            missing = [label for label, ok in (("Total TC", f["total_tc"]),
                                               ("Planned Start", f["planned_start"]),
                                               ("Planned Finish", f["planned_finish"])) if not ok]
            if missing:
                item.flags.append(Flag(FLAG_MISSING_DATA,
                                       "No Traffic Light: " + ", ".join(missing) + " empty or under review"))
        if f["actual_finish"] and f["planned_finish"] and f["actual_finish"] > f["planned_finish"]:
            item.flags.append(Flag(FLAG_ACTUAL_FINISH_LATE,
                                   f"Actual Finish {fmt(f['actual_finish'])} > Planned Finish {fmt(f['planned_finish'])}"))
        items.append(item)
    return items


# --------------------------------------------------------------------------- Defect List

def normalize_defects(rows: list[RawRow], window: DateWindow) -> list[Item]:
    items: list[Item] = []
    for r in rows:
        v = r.values
        no = task_no_text(v.get("no"))
        if all(cell_text(val) == "" for k, val in v.items() if k != "no"):
            continue                                        # pre-numbered empty rows
        detail = cell_text(v.get("detail"))
        item = Item(
            sheet=DEFECT_LIST, excel_row=r.excel_row,
            key=("no", no) if no else ("detail", collapse(detail)),
            alt_key=None,
            id_base=f"DEF-{no}" if no else None,
            id_blank_prefix="DEF-x",
        )
        item.raw["no"] = no
        item.fields["no"] = no or None
        _date(item, "defect_date", v.get("defect_date"), window)
        _text(item, "system", v.get("system"))
        _text(item, "detail", v.get("detail"))
        _status(item, v.get("status"), normalize_defect_status)
        _text(item, "severity", v.get("severity"))
        owner = _text(item, "assigned_to", v.get("assigned_to"))
        _date(item, "fix_target_date", v.get("fix_target_date"), window)
        _text(item, "data_test", v.get("data_test"))
        _text(item, "raised_by", v.get("raised_by"))
        _text(item, "remark", v.get("remark"))
        if not owner:
            item.flags.append(Flag(FLAG_MISSING_OWNER, "Assigned To is empty"))
        items.append(item)
    return items


NORMALIZERS = {
    UAT_PLAN: normalize_uat_plan,
    TEST_PROGRESS: normalize_test_progress,
    DEFECT_LIST: normalize_defects,
}
