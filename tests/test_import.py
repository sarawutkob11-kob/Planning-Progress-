import json
import sqlite3

import openpyxl
import pytest

from uat_tracker.excel.reader import sha256_file
from uat_tracker.services.importer import run_import

from conftest import DEF_ROWS, TP_ROWS, build_workbook


def items(repo, sheet):
    return {k: v for k, v in repo.load_items(sheet).items() if v["active"]}


def by_name(repo, sheet, field, value):
    matches = [v for v in items(repo, sheet).values() if v[field] == value]
    assert len(matches) == 1, matches
    return matches[0]


def open_reviews(repo, item_id=None):
    return [r for r in repo.open_reviews() if item_id is None or r["item_id"] == item_id]


def flag_codes(repo, item_id):
    return {f["code"] for f in repo.active_flags() if f["item_id"] == item_id}


# ------------------------------------------------------------------ structure / IDs
def test_ids_groups_and_skips(cfg, repo):
    run_import(cfg, repo, editor="tester")
    uat = items(repo, "UAT Plan")
    assert set(uat) == {"A-1", "A-2", "B-6", "B-7", "E-8", "E-8b", "E-11", "E-x1"}
    assert uat["A-1"]["group_title"] == "SETUP for UAT Testing"
    assert uat["E-8"]["task_name"] == "CRD-TAX" and uat["E-8b"]["task_name"] == "EFS - Run Payment File"
    tp = items(repo, "Test Progress")
    assert set(tp) == {"TP-1", "TP-6", "TP-12", "TP-38", "TP-35"}       # Grand Total skipped
    assert set(items(repo, "Defect List")) == {"DEF-1", "DEF-5", "DEF-21"}  # number-only rows skipped


def test_status_normalised_and_unknown_to_review(cfg, repo):
    run_import(cfg, repo, editor="tester")
    uat = items(repo, "UAT Plan")
    assert uat["A-2"]["status"] == "Complete"          # 'comlete'
    assert uat["E-8b"]["status"] == "In progress"      # 'Inprogress'
    assert uat["E-11"]["status"] is None
    assert any(r["problem"] == "unknown_status" and r["raw_value"] == "Waiting"
               for r in open_reviews(repo, "E-11"))


# ------------------------------------------------------------------ acceptance 1-4
def test_acceptance_1_ccms_start_and_history(cfg, repo):
    run_import(cfg, repo, editor="tester")
    t = by_name(repo, "UAT Plan", "task_name", "CCMS run retrieve data from EDW")
    assert t["start_date"] == "2026-09-14"
    hist = repo.date_history("UAT Plan", t["id"], "start_date")
    assert [h["value"] for h in hist] == ["2026-10-20", "2026-09-25", "2026-09-17"]
    assert t["finish_date"] == "2026-10-22"
    assert len(repo.date_history("UAT Plan", t["id"], "finish_date")) == 3


def test_acceptance_2_buddhist_year(cfg, repo):
    run_import(cfg, repo, editor="tester")
    assert items(repo, "UAT Plan")["A-2"]["start_date"] == "2026-09-14"   # text '9/14/2569'
    assert items(repo, "UAT Plan")["A-2"]["finish_date"] == "2026-09-18"  # native 2569 date
    nleads = items(repo, "Test Progress")["TP-1"]
    assert nleads["planned_start"] == "2026-09-14" and nleads["actual_start"] == "2026-09-15"


def test_acceptance_3_typos_go_to_review_not_converted(cfg, repo):
    run_import(cfg, repo, editor="tester")
    reviews = open_reviews(repo)
    tokens = {r["token"]: r for r in reviews}
    for raw, problem in [("24/09/206", "bad_date"), ("29/09/2027", "out_of_window"),
                         ("1/4/1900", "out_of_window"), ("06-Now-2026", "date_in_time_column"),
                         ("10/8/2026", "ambiguous_date")]:
        match = [r for r in reviews if raw in (r["raw_value"] or "")]
        assert match, raw
        assert any(r["problem"] == problem for r in match), (raw, match)
    uat = items(repo, "UAT Plan")
    assert uat["B-7"]["finish_date"] is None                # 24/09/206
    assert uat["E-8"]["finish_date"] is None                # 29/09/2027
    assert uat["E-8b"]["start_date"] is None                # 10/8/2026
    assert items(repo, "Test Progress")["TP-12"]["planned_start"] is None
    dms = items(repo, "Test Progress")["TP-38"]
    assert dms["planned_start"] == "2026-10-22"             # 1/4/1900 only in history -> review
    assert tokens["1/4/1900"]["field"] == "planned_start_history"
    assert uat["E-11"]["finish_time"] == "06-Now-2026\n13-Nov-2026"   # left as-is, not moved
    assert uat["E-11"]["finish_date"] is None
    # native date swapped by Excel (fix target in May for a defect raised in Nov)
    assert items(repo, "Defect List")["DEF-21"]["fix_target_date"] is None
    assert any(r["item_id"] == "DEF-21" and r["problem"] == "out_of_window" for r in reviews)


def test_acceptance_4_pass_fail_mismatch(cfg, repo):
    run_import(cfg, repo, editor="tester")
    atm = by_name(repo, "Test Progress", "application", "ATM")
    assert atm["executed"] == 4 and atm["pass_count"] == 2 and atm["fail_count"] == 0
    assert "pass_fail_mismatch" in flag_codes(repo, atm["id"])
    assert "pass_fail_mismatch" not in flag_codes(repo, "TP-38")       # 124 + 2 = 126


def test_other_flags(cfg, repo):
    run_import(cfg, repo, editor="tester")
    assert "missing_owner" in flag_codes(repo, "B-7")
    assert "missing_owner" in flag_codes(repo, "DEF-5")
    assert "missing_report_email" in flag_codes(repo, "TP-1")          # column not in sheet yet
    assert "complete_no_actual_finish" in flag_codes(repo, "TP-35")
    assert "missing_fail" in flag_codes(repo, "TP-35")
    assert "complete_no_actual_finish" not in flag_codes(repo, "TP-1")


def test_note_from_parentheses(cfg, repo):
    run_import(cfg, repo, editor="tester")
    t = items(repo, "UAT Plan")["E-8"]
    assert t["start_date"] == "2026-11-02" and t["start_date_note"] == "(BD 30/09/26)"


# ------------------------------------------------------------------ acceptance 5: idempotent
def test_acceptance_5_reimport_same_file_changes_nothing(cfg, repo):
    run_import(cfg, repo, editor="tester")
    snap = _table_counts(repo)
    audit_before = repo.count_audit()
    res = run_import(cfg, repo, editor="tester")
    assert repo.count_audit() == audit_before
    assert res.audit_added == 0
    assert _table_counts(repo) == snap
    assert all(s.created == 0 and s.updated == 0 and s.reviews_added == 0 and s.flags_added == 0
               for s in res.sheets.values())
    assert len(list(cfg.snapshot_dir.glob("*.xlsx"))) == 1             # identical snapshot reused


def _table_counts(repo):
    return {t: repo.conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
            for t in ("uat_task", "test_progress", "defect", "date_history", "review_queue",
                      "data_flag", "cell_baseline", "audit_log")}


# ------------------------------------------------------------------ audit log
def test_audit_on_create_and_change(cfg, repo, tmp_path):
    run_import(cfg, repo, editor="tester")
    created = repo.audit_entries("TP-6")
    assert len(created) == 1 and created[0]["field"] == "*" and created[0]["channel"] == "import"
    assert json.loads(created[0]["new_value"])["planned_finish"] == "08-Oct-2026"

    rows = [list(r) for r in TP_ROWS]
    rows[1][6] = 4                                                     # ATM #Pass 2 -> 4
    rows[1][9] = "08/10/2026 => 15/10/2026"                            # reschedule
    build_workbook(cfg.excel_path, tp=rows)
    res = run_import(cfg, repo, editor="tester")
    entries = repo.audit_entries("TP-6")[1:]
    changes = {e["field"]: (e["old_value"], e["new_value"]) for e in entries}
    assert changes["pass_count"] == ("2", "4")
    assert changes["planned_finish"] == ("08-Oct-2026", "15-Oct-2026")
    assert changes["planned_finish_history"] == ("", "08-Oct-2026")
    assert all(e["editor"] == "tester" and e["sheet"] == "Test Progress" and e["ts"] for e in entries)
    assert "pass_fail_mismatch" not in flag_codes(repo, "TP-6")       # flag cleared
    assert res.sheets["Test Progress"].flags_cleared >= 1
    assert repo.verify_audit_chain()


def test_audit_log_is_immutable(cfg, repo):
    run_import(cfg, repo, editor="tester")
    with pytest.raises(sqlite3.DatabaseError):
        repo.conn.execute("UPDATE audit_log SET new_value='x'")
    with pytest.raises(sqlite3.DatabaseError):
        repo.conn.execute("DELETE FROM audit_log")


def test_removed_row_kept_inactive_and_restored(cfg, repo):
    run_import(cfg, repo, editor="tester")
    build_workbook(cfg.excel_path, defects=[r for r in DEF_ROWS if r[0] != 5])
    run_import(cfg, repo, editor="tester")
    rec = repo.load_items("Defect List")["DEF-5"]
    assert rec["active"] == 0
    build_workbook(cfg.excel_path)
    run_import(cfg, repo, editor="tester")
    assert repo.load_items("Defect List")["DEF-5"]["active"] == 1
    fields = [e["new_value"] for e in repo.audit_entries("DEF-5") if e["field"] == "(row)"]
    assert fields == ["removed", "active"]


def test_review_superseded_when_fixed_in_excel(cfg, repo):
    run_import(cfg, repo, editor="tester")
    rows = [list(r) for r in TP_ROWS]
    rows[2][8] = "24/09/2026"                                         # COLLINS typo fixed
    build_workbook(cfg.excel_path, tp=rows)
    run_import(cfg, repo, editor="tester")
    assert not [r for r in open_reviews(repo, "TP-12") if r["field"] == "planned_start"]
    assert items(repo, "Test Progress")["TP-12"]["planned_start"] == "2026-09-24"


# ------------------------------------------------------------------ source safety
def test_source_never_modified_and_other_sheets_untouched(cfg, repo):
    before = sha256_file(cfg.excel_path)
    run_import(cfg, repo, editor="tester")
    run_import(cfg, repo, editor="tester")
    assert sha256_file(cfg.excel_path) == before
    wb = openpyxl.load_workbook(cfg.excel_path)
    assert wb["NLeadAccount"]["A1"].value == "do not touch"


def test_email_columns_when_present(cfg, repo):
    header = ["No.", "Application", "Report By", "Report By Email", "CC Email", "Total TC",
              "No. of TC Execution", "%Progress", "#Pass", "#Fail", "Planned Start", "Planned Finish",
              "Actual Start", "Actual Finish", "Traffic Light", "TC/Result Complete", "Status", "Remark"]
    rows = [[1, "Nleads", "QA G", "g@x.com", "a@x.com; b@x.com;", 1, 1, 1, 1, 0,
             None, None, None, None, None, None, None, None]]
    build_workbook(cfg.excel_path, tp=rows, tp_header=header)
    res = run_import(cfg, repo, editor="tester")
    rec = items(repo, "Test Progress")["TP-1"]
    assert rec["report_by_email"] == "g@x.com" and rec["cc_emails"] == "a@x.com; b@x.com"
    assert "missing_report_email" not in flag_codes(repo, "TP-1")
    assert res.sheets["Test Progress"].missing_optional_columns == ["App ID"]   # added by Part 3 sync
