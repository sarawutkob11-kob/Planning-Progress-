"""Part 4 services behind the screens: dashboard numbers, views, new rows, restore, export, settings."""

import dataclasses
import datetime as dt

import openpyxl
import pytest

from uat_tracker.config import AlertsConfig, apply_overrides
from uat_tracker.excel.writer import OpenpyxlWorkbook
from uat_tracker.mail.sender import FakeSender
from uat_tracker.services import dashboard, views
from uat_tracker.services.edits import EditError, edit_item
from uat_tracker.services.importer import run_import
from uat_tracker.services.items import add_defect, add_task
from uat_tracker.services.sync import SyncService
from conftest import PMO_ACTOR  # noqa: E402

TP = "Test Progress"
DAY = dt.date(2026, 11, 6)


@pytest.fixture
def ucfg(cfg):
    return dataclasses.replace(cfg, alerts=AlertsConfig(pmo_email="pmo@example.test"))


@pytest.fixture
def svc(ucfg, repo):
    run_import(ucfg, repo, editor="t")
    s = SyncService(ucfg, repo, opener=OpenpyxlWorkbook, sender=FakeSender())
    s.init_ids()
    return s


def test_dashboard_numbers_match_db(ucfg, repo, svc):
    d = dashboard.build(ucfg, repo, DAY)
    tp = [r for r in repo.load_items(TP).values() if r["active"]]
    assert d.executed == sum(int(r["executed"] or 0) for r in tp) == 178
    assert d.total_tc == sum(int(r["total_tc"] or 0) for r in tp) == 238
    assert d.open_tasks == 5 and d.all_tasks == 8                 # 3 of 8 fixture tasks Complete
    assert d.no_target_defects == 1 and d.open_defects == 1        # DEF-21 open, target out of window
    tags = {(f.tag, f.item_id) for f in d.follow}
    assert ("Amber", "TP-38") in tags and ("No target", "DEF-21") in tags
    assert all(f.manual == (f.sheet != TP) for f in d.follow)
    assert len(d.missing_email) == 5


def test_uat_plan_view_groups_filters_and_detail(ucfg, repo, svc):
    groups = views.uat_plan(repo, DAY)
    assert [g["code"] for g in groups] == ["A", "B", "E"]
    a = next(g for g in groups if g["code"] == "A")
    assert a["complete"]                                          # collapsed by default in the UI
    b6 = next(i for g in groups for i in g["items"] if i["id"] == "B-6")
    assert b6["reschedules"] == 3 and b6["planned_first"] == "2026-10-10" and b6["current"] == "2026-10-22"
    assert [i["id"] for g in views.uat_plan(repo, DAY, "resched2") for i in g["items"]] == ["B-6"]
    det = views.uat_detail(ucfg, repo, "B-6", DAY)
    assert len(det["timeline"]) == 6


def test_test_progress_view(ucfg, repo, svc):
    summary, rows = views.test_progress(ucfg, repo, DAY)
    assert summary["executed"] == 178 and summary["amber"] >= 1
    atm = next(r for r in rows if r["app"] == "ATM")
    assert atm["mismatch"]
    _, follow = views.test_progress(ucfg, repo, DAY, "follow")
    assert {r["color"] for r in follow} <= {"Red", "Amber", "Missing"}


def test_add_task_and_defect_create_excel_rows(ucfg, repo, svc):
    with pytest.raises(EditError, match="Responsible"):
        add_task(ucfg, repo, group="B", task_name="New", responsible="", actor=PMO_ACTOR)
    tid = add_task(ucfg, repo, group="B", task_no="9", task_name="New check", responsible="QA New",
                   finish_date="2026-11-20", actor=PMO_ACTOR)
    with pytest.raises(EditError, match="target"):
        add_defect(ucfg, repo, system="CCMS", detail="x", assigned_to="Dev A", fix_target_date="", actor=PMO_ACTOR)
    did = add_defect(ucfg, repo, system="CCMS", detail="Wrong total", assigned_to="Dev A",
                     fix_target_date="2026-11-30", actor=PMO_ACTOR)
    assert (tid, did) == ("B-9", "DEF-22")
    svc.pull(force=True)                                          # not written yet: still active
    assert repo.get_item("UAT Plan", tid)["active"]
    fr = svc.flush(force=True)
    assert fr.result == "written" and not fr.failed
    assert repo.all_pending() == []
    ws = openpyxl.load_workbook(ucfg.excel_path)["UAT Plan"]
    names = [ws.cell(r, 5).value for r in range(1, ws.max_row + 1)]
    assert names.index("New check") == names.index("No owner task") + 1   # right after group B
    t = repo.get_item("UAT Plan", tid)
    assert t["excel_row"] and t["task_name"] == "New check" and t["finish_date"] == "2026-11-20"
    d = repo.get_item("Defect List", did)
    assert d["excel_row"] == 5 and d["no"] == "22"                # first empty numbered row
    assert repo.get_item("UAT Plan", "E-8")["task_name"] == "CRD-TAX"   # rows below shifted, still bound


def test_restore_creates_new_entry(ucfg, repo, svc):
    edit_item(ucfg, repo, TP, "TP-38", {"remark": "wrong"}, actor=PMO_ACTOR)
    e = [a for a in repo.audit_entries("TP-38") if a["field"] == "remark"][-1]
    views.restore(ucfg, repo, e["id"], actor=PMO_ACTOR)
    assert repo.get_item(TP, "TP-38")["remark"] is None
    last = repo.audit_entries("TP-38")[-1]
    assert last["channel"] == "restore" and f"#{e['id']}" in last["reason"]
    edit_item(ucfg, repo, TP, "TP-38", {"planned_finish": "2026-11-20"}, actor=PMO_ACTOR, reason="r")
    e2 = [a for a in repo.audit_entries("TP-38") if a["field"] == "planned_finish"][-1]
    views.restore(ucfg, repo, e2["id"], actor=PMO_ACTOR)
    assert repo.get_item(TP, "TP-38")["planned_finish"] == "2026-11-13"


def test_history_types_and_export(ucfg, repo, svc, tmp_path):
    edit_item(ucfg, repo, TP, "TP-38", {"planned_finish": "2026-11-20", "status": "Blocked"}, actor=PMO_ACTOR,
              reason="r")
    rows = views.history(repo, item_id="TP-38")
    types = {(r["field"], r["type"]) for r in rows}
    assert ("planned_finish", "Reschedule") in types and ("status", "Status change") in types
    assert ("*", "Import") in types
    assert views.history(repo, kind="Reschedule", item_id="TP-38")[0]["field"] == "planned_finish"
    out = tmp_path / "history.xlsx"
    views.export_history(rows, out)
    ws = openpyxl.load_workbook(out).active
    assert ws["A1"].value == "#" and ws.max_row == len(rows) + 1


def test_defect_view_and_board_move(ucfg, repo, svc):
    rows = views.defects(repo, DAY)
    d21 = next(r for r in rows if r["id"] == "DEF-21")
    assert d21["column"] == "Open" and d21["no_target"] and d21["days_open"] == 3
    edit_item(ucfg, repo, "Defect List", "DEF-21", {"status": "Fixed"}, actor=PMO_ACTOR)
    assert next(r for r in views.defects(repo, DAY) if r["id"] == "DEF-21")["column"] == "Fixed"


def test_settings_overrides(ucfg, repo):
    with repo.transaction():
        repo.set_setting("settings.fail_threshold", "0.1")
        repo.set_setting("settings.dry_run", "false")
    c = apply_overrides(ucfg, repo)
    assert c.alerts.fail_threshold == 0.1 and c.alerts.dry_run is False
    assert ucfg.alerts.dry_run is True
