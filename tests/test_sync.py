"""Part 3: 2-way sync, App ID column, write-back, conflicts, review resolution, backups."""

import dataclasses
import datetime as dt

import openpyxl
import pytest

from uat_tracker.excel.writer import OpenpyxlWorkbook, WorkbookLocked
from uat_tracker.mail.sender import FakeSender
from uat_tracker.services.attachments import upload
from uat_tracker.services.edits import EditError, edit_item
from uat_tracker.services.importer import run_import
from uat_tracker.services.resolve import ResolveError, resolve_conflict, resolve_review
from uat_tracker.services.sync import SyncService

from conftest import build_workbook
from conftest import PMO_ACTOR  # noqa: E402

TP = "Test Progress"
PMO = "pmo@example.test"


@pytest.fixture
def scfg(cfg):
    from uat_tracker.config import AlertsConfig
    return dataclasses.replace(cfg, alerts=AlertsConfig(pmo_email=PMO))


@pytest.fixture
def svc(scfg, repo):
    run_import(scfg, repo, editor="t")
    return SyncService(scfg, repo, opener=OpenpyxlWorkbook, sender=FakeSender())


def sheet_cols(path, sheet):
    ws = openpyxl.load_workbook(path)[sheet]
    return ws, {str(c.value).split("\n")[0].strip(): c.column for c in ws[1] if c.value}


def cell(path, sheet, item_id, header):
    ws, cols = sheet_cols(path, sheet)
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, cols["App ID"]).value == item_id:
            return ws.cell(r, cols[header]).value
    raise KeyError(item_id)


def excel_edit(path, sheet, item_id, header, value):
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    cols = {str(c.value).split("\n")[0].strip(): c.column for c in ws[1] if c.value}
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, cols["App ID"]).value == item_id:
            ws.cell(r, cols[header]).value = value
            wb.save(path)
            return
    raise KeyError(item_id)


def tp(repo, item_id):
    return repo.get_item(TP, item_id)


# ------------------------------------------------------------------ App ID column
def test_app_id_column_added_after_table_hidden_and_used_for_binding(scfg, repo, svc):
    fr = svc.init_ids()
    assert fr.result == "written" and "column added" in fr.message
    ws, cols = sheet_cols(scfg.excel_path, TP)
    assert cols["App ID"] == cols["Remark"] + 1                         # inside the table
    assert ws.cell(1, cols["App ID"] + 2).value == "Status"            # legend shifted right, kept
    assert ws.column_dimensions[ws.cell(1, cols["App ID"]).column_letter].hidden
    assert openpyxl.load_workbook(scfg.excel_path)["NLeadAccount"]["A1"].value == "do not touch"
    assert cell(scfg.excel_path, TP, "TP-38", "Application") == "DMS-RAMC"
    # a row inserted above in Excel: IDs still bind by App ID
    wb = openpyxl.load_workbook(scfg.excel_path)
    wb[TP].insert_rows(2)
    wb.save(scfg.excel_path)
    res = svc.pull()
    assert res.sheets[TP].created == 0 and res.sheets[TP].removed == 0
    assert tp(repo, "TP-38")["excel_row"] == 6


# ------------------------------------------------------------------ app -> Excel
def test_edit_in_app_is_written_to_excel_with_backup_and_audit(scfg, repo, svc):
    edit_item(scfg, repo, TP, "TP-38", {"executed": 130, "pass_count": 128}, actor=PMO_ACTOR, reason="update")
    assert len(repo.all_pending()) == 2
    fr = svc.flush(force=True)
    assert fr.result == "written" and fr.written == 2 and fr.backup
    assert cell(scfg.excel_path, TP, "TP-38", "No. of TC Execution") == 130
    assert repo.all_pending() == []
    assert [b["reason"] for b in repo.backups()] == ["before writing 2 change(s)"]
    fields = [(a["channel"], a["field"]) for a in repo.audit_entries("TP-38")]
    assert ("app", "executed") in fields and ("excel_sync", "executed") in fields


def test_flush_batches_changes_for_quiet_period(scfg, repo, svc):
    edit_item(scfg, repo, TP, "TP-38", {"remark": "x"}, actor=PMO_ACTOR)
    assert svc.flush().result == "waiting"
    assert svc.flush(force=True).result == "written"


def test_reschedule_needs_reason_and_writes_history_chain(scfg, repo, svc):
    with pytest.raises(EditError, match="Reason is required"):
        edit_item(scfg, repo, TP, "TP-38", {"planned_finish": "2026-11-20"}, actor=PMO_ACTOR)
    edit_item(scfg, repo, TP, "TP-38", {"planned_finish": "2026-11-20"}, actor=PMO_ACTOR,
              reason="Waiting data", root_cause="Waiting other system data")
    svc.flush(force=True)
    text = cell(scfg.excel_path, TP, "TP-38", "Planned Finish")
    assert text == "11/06/2026 => 13-Nov-2026 => 20-Nov-2026"   # unresolved bad token kept as typed
    rec = tp(repo, "TP-38")
    assert rec["planned_finish"] == "2026-11-20" and rec["planned_finish_history"][-1] == "2026-11-13"
    assert not [a for a in repo.audit_entries("TP-38")                    # own write is not "Edited in Excel"
                if a["field"] == "planned_finish_history" and a["channel"] == "excel_sync"]
    app_hist = [h for h in repo.date_history(TP, "TP-38", "planned_finish") if h["source"] == "app"]
    assert app_hist[0]["reason"] == "Waiting data" and app_hist[0]["root_cause"] == "Waiting other system data"


def test_single_date_written_as_real_date(scfg, repo, svc):
    edit_item(scfg, repo, TP, "TP-38", {"actual_start": "2026-10-23"}, actor=PMO_ACTOR)
    svc.flush(force=True)
    assert cell(scfg.excel_path, TP, "TP-38", "Actual Start") == dt.datetime(2026, 10, 23)


def test_edit_rules(scfg, repo, svc):
    with pytest.raises(EditError, match="after Planned Finish"):
        edit_item(scfg, repo, TP, "TP-38", {"actual_finish": "2026-11-20"}, actor=PMO_ACTOR)
    edit_item(scfg, repo, TP, "TP-38", {"planned_finish": "2026-11-25", "actual_finish": "2026-11-20"},
              actor=PMO_ACTOR, reason="moved")
    with pytest.raises(EditError, match="Actual Finish"):
        edit_item(scfg, repo, TP, "TP-6", {"status": "Complete"}, actor=PMO_ACTOR)
    with pytest.raises(EditError, match="valid address"):
        edit_item(scfg, repo, TP, "TP-6", {"report_by_email": "a@x.com; b@x.com"}, actor=PMO_ACTOR)
    with pytest.raises(EditError, match="not editable"):
        edit_item(scfg, repo, TP, "TP-6", {"pct_progress": 1}, actor=PMO_ACTOR)
    with pytest.raises(EditError, match="standard"):
        edit_item(scfg, repo, TP, "TP-6", {"status": "Waiting"}, actor=PMO_ACTOR)


def test_formula_cells_are_never_overwritten(scfg, repo, svc):
    svc.init_ids()
    excel_edit(scfg.excel_path, TP, "TP-6", "#Fail", "=1-1")
    svc.pull()
    edit_item(scfg, repo, TP, "TP-6", {"fail_count": 2}, actor=PMO_ACTOR)
    fr = svc.flush(force=True)
    assert any("formula" in f for f in fr.failed)
    assert cell(scfg.excel_path, TP, "TP-6", "#Fail") == "=1-1"
    assert repo.pending(TP)[("TP-6", "fail_count")]["status"] == "failed"
    assert svc.status()["state"] == "error"


# ------------------------------------------------------------------ Excel -> app
def test_excel_edit_reaches_app_with_audit(scfg, repo, svc):
    svc.init_ids()
    excel_edit(scfg.excel_path, TP, "TP-38", "#Pass", 125)
    res = svc.pull()
    assert res.sheets[TP].updated == 1
    assert tp(repo, "TP-38")["pass_count"] == 125
    e = [a for a in repo.audit_entries("TP-38") if a["field"] == "pass_count"][-1]
    assert e["channel"] == "excel_sync" and e["editor"].startswith("Edited in Excel")
    assert svc.pull() is None                                      # unchanged file: nothing to do


# ------------------------------------------------------------------ conflicts
def test_same_cell_changed_both_sides_is_conflict_and_nothing_overwritten(scfg, repo, svc):
    svc.init_ids()
    edit_item(scfg, repo, TP, "TP-38", {"remark": "app note"}, actor=PMO_ACTOR)
    excel_edit(scfg.excel_path, TP, "TP-38", "Remark", "excel note")
    res = svc.pull()
    assert len(res.conflicts_added) == 1
    assert tp(repo, "TP-38")["remark"] == "app note"
    assert svc.flush(force=True).result == "nothing"
    assert cell(scfg.excel_path, TP, "TP-38", "Remark") == "excel note"
    assert svc.sender.sent and "conflict" in svc.sender.sent[-1].subject.lower()
    svc.pull(force=True)
    assert len(repo.open_conflicts()) == 1                          # not duplicated

    cid = repo.open_conflicts()[0]["id"]
    with pytest.raises(ResolveError, match="reason"):
        resolve_conflict(scfg, repo, cid, "app", actor=PMO_ACTOR, reason="")
    resolve_conflict(scfg, repo, cid, "app", actor=PMO_ACTOR, reason="app is right")
    svc.pull(force=True)
    svc.flush(force=True)
    assert cell(scfg.excel_path, TP, "TP-38", "Remark") == "app note"
    assert repo.open_conflicts() == [] and repo.all_pending() == []


def test_conflict_resolved_with_excel_value(scfg, repo, svc):
    svc.init_ids()
    edit_item(scfg, repo, TP, "TP-38", {"remark": "app note"}, actor=PMO_ACTOR)
    excel_edit(scfg.excel_path, TP, "TP-38", "Remark", "excel note")
    svc.pull()
    resolve_conflict(scfg, repo, repo.open_conflicts()[0]["id"], "excel", actor=PMO_ACTOR, reason="Excel ok")
    svc.pull(force=True)
    assert tp(repo, "TP-38")["remark"] == "excel note" and repo.all_pending() == []


def test_conflict_resolved_with_new_value(scfg, repo, svc):
    svc.init_ids()
    edit_item(scfg, repo, TP, "TP-38", {"remark": "app note"}, actor=PMO_ACTOR)
    excel_edit(scfg.excel_path, TP, "TP-38", "Remark", "excel note")
    svc.pull()
    resolve_conflict(scfg, repo, repo.open_conflicts()[0]["id"], "new", actor=PMO_ACTOR, reason="merge",
                     new_value="merged note")
    svc.pull(force=True)
    svc.flush(force=True)
    assert cell(scfg.excel_path, TP, "TP-38", "Remark") == "merged note"
    assert tp(repo, "TP-38")["remark"] == "merged note"


def test_actual_finish_after_planned_finish_in_excel_is_conflict(scfg, repo, svc):
    svc.init_ids()
    excel_edit(scfg.excel_path, TP, "TP-38", "Actual Finish", dt.datetime(2026, 11, 20))
    res = svc.pull()
    c = repo.get_conflict(res.conflicts_added[0])
    assert c["kind"] == "actual_after_planned" and c["excel_value"] == "2026-11-20"
    assert tp(repo, "TP-38")["actual_finish"] is None
    resolve_conflict(scfg, repo, c["id"], "app", actor=PMO_ACTOR, reason="not finished yet")
    svc.pull(force=True)
    svc.flush(force=True)
    assert cell(scfg.excel_path, TP, "TP-38", "Actual Finish") is None


# ------------------------------------------------------------------ review queue
def test_review_resolution_writes_clean_value(scfg, repo, svc):
    reviews = {(r["item_id"], r["token"]): r for r in repo.open_reviews()}
    bad_hist = reviews[("TP-38", "1/4/1900")]                       # history token in Planned Start
    resolve_review(scfg, repo, bad_hist["id"], None, actor=PMO_ACTOR, reason="typo")
    bad_cur = reviews[("TP-12", "24/09/206")]                       # COLLINS planned start
    resolve_review(scfg, repo, bad_cur["id"], "2026-09-24", actor=PMO_ACTOR)
    assert tp(repo, "TP-12")["planned_start"] == "2026-09-24"
    svc.flush(force=True)
    assert cell(scfg.excel_path, TP, "TP-38", "Planned Start") == dt.datetime(2026, 10, 22)
    assert cell(scfg.excel_path, TP, "TP-12", "Planned Start") == dt.datetime(2026, 9, 24)
    still_open = {(r["item_id"], r["token"]) for r in repo.open_reviews()}
    assert ("TP-38", "1/4/1900") not in still_open and ("TP-12", "24/09/206") not in still_open
    assert ("B-7", "24/09/206") in still_open                      # other item: untouched
    assert tp(repo, "TP-12")["planned_start"] == "2026-09-24"


def test_unknown_status_review_resolution(scfg, repo, svc):
    r = next(r for r in repo.open_reviews() if r["problem"] == "unknown_status")
    with pytest.raises(ResolveError):
        resolve_review(scfg, repo, r["id"], "Waiting", actor=PMO_ACTOR)
    resolve_review(scfg, repo, r["id"], "In progress", actor=PMO_ACTOR)
    svc.flush(force=True)
    assert repo.get_item("UAT Plan", r["item_id"])["status"] == "In progress"


# ------------------------------------------------------------------ upload -> Excel (spec acceptance 9)
def test_upload_writes_tc_result_complete_to_excel(scfg, repo, svc, tmp_path):
    f = tmp_path / "r.pdf"
    f.write_bytes(b"%PDF")
    upload(scfg, repo, "TP-38", f, actor=PMO_ACTOR)
    svc.flush(force=True)
    assert cell(scfg.excel_path, TP, "TP-38", "TC/Result Complete") == "Complete"
    assert repo.all_pending() == []


# ------------------------------------------------------------------ failures, block, restore
class LockedBook:
    def __init__(self, path):
        raise WorkbookLocked("locked by another program")


def test_locked_file_retries_then_error_and_one_pmo_email(scfg, repo, svc):
    t = [dt.datetime(2026, 11, 9, 10, 0)]
    locked = SyncService(scfg, repo, opener=LockedBook, sender=FakeSender(), clock=lambda: t[0])
    edit_item(scfg, repo, TP, "TP-38", {"remark": "x"}, actor=PMO_ACTOR, now="2026-11-09T09:00:00+07:00")
    results = []
    for minutes in (0, 1, 2, 5, 10, 15, 30, 31):
        t[0] = dt.datetime(2026, 11, 9, 10, 0) + dt.timedelta(minutes=minutes, seconds=1)
        results.append(locked.flush(force=True).result)
    assert results[:6] == ["retry"] * 6 and results[6] == "error" and results[7] == "nothing"
    assert len([m for m in locked.sender.sent if "Excel sync error" in m.subject]) == 1
    assert locked.status()["state"] == "error"
    locked.retry_now()                                              # file changed -> tried again
    assert locked.flush(force=True).result == "error"


def test_onedrive_conflict_copy_blocks_writes(scfg, repo, svc, monkeypatch):
    monkeypatch.setenv("COMPUTERNAME", "PMO-PC")
    copy = scfg.excel_path.with_name(f"{scfg.excel_path.stem}-PMO-PC.xlsx")
    copy.write_bytes(b"x")
    edit_item(scfg, repo, TP, "TP-38", {"remark": "x"}, actor=PMO_ACTOR)
    assert svc.flush(force=True).result == "blocked"
    assert svc.status()["state"] == "blocked"
    copy.unlink()
    assert svc.flush(force=True).result == "blocked"                # until the PMO unblocks
    svc.unblock()
    assert svc.flush(force=True).result == "written"


def test_backup_restore(scfg, repo, svc):
    edit_item(scfg, repo, TP, "TP-38", {"remark": "after"}, actor=PMO_ACTOR)
    svc.flush(force=True)
    b = repo.backups()[-1]                                          # copy taken before the write
    svc.restore(b["id"], PMO_ACTOR)
    assert tp(repo, "TP-38")["remark"] is None
    assert any(a["channel"] == "restore" for a in repo.audit_entries())
    assert len(repo.backups()) == 2                                 # current file saved first


def test_daily_job_writes_calculated_traffic_light(scfg, repo, svc):
    from uat_tracker.services.daily_job import run_daily
    from test_alerts import fake_fetch
    run_daily(scfg, repo, FakeSender(), clock=dt.datetime(2026, 11, 6, 8, 30), fetcher=fake_fetch, sync=svc)
    assert cell(scfg.excel_path, TP, "TP-38", "Traffic Light") == "Amber"
    assert cell(scfg.excel_path, TP, "TP-1", "Traffic Light") == "Complete"
    assert repo.all_pending() == []


def test_excel_window_title_matching():
    from uat_tracker.excel.writer import title_stem
    assert title_stem("UAT Planning - Excel") == "uat planning"
    assert title_stem("UAT Planning.xlsx [Read-Only] - Excel") == "uat planning"
    assert title_stem("UAT Planning - Saved - Excel") == "uat planning"
    assert title_stem("UAT Planning TEST - Excel") != "uat planning"


# ------------------------------------------------------------------ review fixes
def test_edit_refused_while_conflict_open(scfg, repo, svc):
    svc.init_ids()
    edit_item(scfg, repo, TP, "TP-38", {"remark": "app note"}, actor=PMO_ACTOR)
    excel_edit(scfg.excel_path, TP, "TP-38", "Remark", "excel note")
    svc.pull()
    with pytest.raises(EditError, match="conflict"):
        edit_item(scfg, repo, TP, "TP-38", {"remark": "another"}, actor=PMO_ACTOR)
    assert repo.pending(TP)[("TP-38", "remark")]["status"] == "conflict"


def test_dates_outside_window_refused(scfg, repo, svc):
    with pytest.raises(EditError, match="window"):
        edit_item(scfg, repo, TP, "TP-38", {"actual_start": "2027-09-29"}, actor=PMO_ACTOR)
    r = next(r for r in repo.open_reviews() if r["token"] == "24/09/206" and r["item_id"] == "TP-12")
    with pytest.raises(ResolveError, match="window"):
        resolve_review(scfg, repo, r["id"], "2206-09-24", actor=PMO_ACTOR)


def test_interrupted_write_is_retried(scfg, repo, svc):
    edit_item(scfg, repo, TP, "TP-38", {"remark": "x"}, actor=PMO_ACTOR)
    with repo.transaction():
        repo.set_pending_state(TP, "TP-38", "remark", status="writing")
    assert svc.flush(force=True).result == "written"
    assert cell(scfg.excel_path, TP, "TP-38", "Remark") == "x"


# ------------------------------------------------------------------ review suggestions (Conflict Center)
@pytest.mark.parametrize("field,problem,token,raw,expected", [
    ("finish_date_history", "ambiguous_date", "10/8/2026", "10/8/2026 => 07/10/2026",
     [("both", "2026-08-10"), ("both", "2026-10-08"), ("drop", "")]),
    ("start_date", "ambiguous_date", "4/11/2026", "4/11/2026", [("inwin", "2026-11-04")]),
    ("finish_date", "out_of_window", "10/05/2026", "10/05/2026", [("swap", "2026-10-05")]),
    ("planned_finish", "out_of_window", "2569-05-10", "2569-05-10", [("swap", "2026-10-05")]),
    ("actual_start", "out_of_window", "2568-09-28", "2568-09-28", [("year", "2026-09-28")]),
    ("planned_start", "bad_date", "24/09/206", "24/09/206", [("year", "2026-09-24")]),
    ("actual_finish", "out_of_window", "29/09/2027", "29/09/2027", [("year", "2026-09-29")]),
    ("planned_start_history", "out_of_window", "1/4/1900", "1/4/1900\n22/10/2026", [("drop", "")]),
    ("planned_start", "bad_date", "30/102026", "10/26/2026\n30/102026", [("year", "2026-10-30"), ("drop", "")]),
    ("finish_time", "date_in_time_column", "06-Now-2026", "06-Now-2026", [("clear", "")]),
])
def test_review_suggestions(scfg, field, problem, token, raw, expected):
    from uat_tracker.services.review_suggest import suggestions
    got = suggestions(scfg, {"sheet": TP, "field": field, "problem": problem, "token": token, "raw_value": raw})
    assert [(s["why"], s["value"]) for s in got] == expected


def test_unknown_status_suggestion(scfg):
    from uat_tracker.services.review_suggest import suggestions
    got = suggestions(scfg, {"sheet": "UAT Plan", "field": "status", "problem": "unknown_status",
                             "token": "Inprogres", "raw_value": "Inprogres"})
    assert got[0] == {"why": "status", "value": "In progress"}
