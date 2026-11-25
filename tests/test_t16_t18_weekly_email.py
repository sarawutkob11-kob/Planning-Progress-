"""QA T16-T18 (25-Nov-2026): weekly report data, new email layout, PMO matters out of the email.

spec.md section 2026-11-25 (R1-R16, AC-4 to AC-23). No real API / Outlook / Excel: tmp_path + fakes only.
"""

from __future__ import annotations

import copy
import dataclasses
import datetime as dt
import json
import re

import pytest

from conftest import PMO_ACTOR
from uat_tracker.config import AlertsConfig
from uat_tracker.excel.sheets import DEFECT_LIST, TEST_PROGRESS, UAT_PLAN
from uat_tracker.mail.sender import FakeSender
from uat_tracker.services import dashboard, weekly_ai, weekly_report as W
from uat_tracker.services.importer import run_import
from uat_tracker.services.traffic_light import AMBER, COMPLETE, GREEN, MISSING, RED, compute, missing_fields
from uat_tracker.services.weekly_data import build_data

MON = dt.date(2026, 11, 16)
NOW = "2026-11-15T00:00:00+07:00"
HEX = {COMPLETE: "#2563EB", RED: "#DC2626", AMBER: "#F59E0B", GREEN: "#16A34A", MISSING: "#9CA3AF"}


class MemKeys:
    def __init__(self):
        self.key = None

    def get(self):
        return self.key

    def set(self, k):
        self.key = k

    def clear(self):
        self.key = None


@pytest.fixture(autouse=True)
def no_real_keystore(monkeypatch):
    monkeypatch.setattr(weekly_ai, "KEYS", MemKeys())


@pytest.fixture
def wcfg(cfg):
    return dataclasses.replace(cfg, alerts=AlertsConfig(pmo_email="pmo@example.test"))


@pytest.fixture
def repo_data(wcfg, repo):
    run_import(wcfg, repo, editor="t")
    with repo.transaction():
        repo.upsert_holiday("2026-12-07", "King Bhumibol Memorial Day", "manual", "2026-11-01T00:00:00+07:00")
    return repo


def _set(repo, sheet, item_id, **fields):
    with repo.transaction():
        for k, v in fields.items():
            repo.set_field(sheet, item_id, k, v, NOW)


def _tp_ids(repo):
    return sorted(i for i, r in repo.load_items(TEST_PROGRESS).items() if r.get("active"))


# --------------------------------------------------------------------------- synthetic data for render

def base_data(**over) -> dict:
    d = {
        "project": "DEMO UAT", "as_of": "25-Nov-2026",
        "reschedule_window": {"from": "18-Nov-2026", "to": "24-Nov-2026"},
        "overall": {"apps": 7, "executed": 260, "total_tc": 370, "pct_executed": "70.3%", "pass": 255, "fail": 5,
                    "colors": {"Complete": 1, "Delayed": 2, "At risk of delay": 3, "On track": 0,
                               "Missing data": 1}},
        "delayed_or_at_risk": [
            {"id": "TP-1", "app": "AppLate", "color": "Delayed", "reason_kind": "overdue", "reason": "x",
             "planned_finish": "20-Nov-2026", "days_late": 5, "executed": 10, "total_tc": 40, "fail": 0,
             "pct_exec": "25.0%", "pct_time": None, "pct_fail": "0.0%"},
            {"id": "TP-2", "app": "AppFailHigh", "color": "Delayed", "reason_kind": "fail_rate", "reason": "x",
             "planned_finish": "30-Nov-2026", "days_late": None, "executed": 50, "total_tc": 60, "fail": 6,
             "pct_exec": "83.3%", "pct_time": None, "pct_fail": "12.0%"},
            {"id": "TP-3", "app": "AppBehind", "color": "At risk of delay", "reason_kind": "time", "reason": "x",
             "planned_finish": "30-Nov-2026", "days_late": None, "executed": 40, "total_tc": 100, "fail": 0,
             "pct_exec": "40.0%", "pct_time": "60.0%", "pct_fail": "0.0%"},
            {"id": "TP-4", "app": "AppFailLow", "color": "At risk of delay", "reason_kind": "fail_rate",
             "reason": "x", "planned_finish": "30-Nov-2026", "days_late": None, "executed": 66, "total_tc": 70,
             "fail": 2, "pct_exec": "94.3%", "pct_time": None, "pct_fail": "3.0%"},
            {"id": "TP-5", "app": "AppBoth", "color": "At risk of delay", "reason_kind": "time", "reason": "x",
             "planned_finish": "30-Nov-2026", "days_late": None, "executed": 94, "total_tc": 100, "fail": 1,
             "pct_exec": "94.0%", "pct_time": "96.0%", "pct_fail": "1.1%"}],
        "missing_data": [{"id": "TP-6", "app": "AppNoTC", "fields": ["Total TC", "Planned Finish"]}],
        "overdue": {"uat_plan": [], "test_progress": [{"id": "TP-1", "app": "AppLate",
                                                        "planned_finish": "20-Nov-2026", "days_late": 5}]},
        "defects": {"open": 5, "overdue": [], "no_target": []},
        "due_next_7_days": [],
        "reschedules_last_7_days": [],
        "open_conflicts": 0, "review_queue": 3, "fail_threshold": "5.0%",
    }
    d.update(over)
    return d


AI = {"summary_th": "สรุปไทยของสัปดาห์นี้", "summary_en": "ENGLISH SUMMARY TEXT",
      "risks": [{"title_th": "หัวข้อความเสี่ยงไทย", "title_en": "RISK TITLE EN", "why_th": "เหตุผลไทย",
                 "why_en": "WHY EN", "related": []}],
      "follow_ups": [{"action_th": "สิ่งที่ต้องทำไทย", "action_en": "ACTION EN", "why_th": "เพราะไทย",
                      "why_en": "BECAUSE EN"}]}


def html_of(data=None, ai=None) -> str:
    return W.render(data or base_data(), ai)[1]


def section(html: str, start: str, end: str | None) -> str:
    i = html.index(start)
    j = html.index(end, i) if end else len(html)
    return html[i:j]


# =========================================================================== T16: data

def test_t16_due_window_edges(wcfg, repo_data):                                      # AC-14
    repo = repo_data
    a, b, c = _tp_ids(repo)[:3]
    for i in (a, b, c):
        _set(repo, TEST_PROGRESS, i, status="In progress", planned_start="2026-11-01")
    _set(repo, TEST_PROGRESS, a, planned_finish=(MON + dt.timedelta(days=7)).isoformat())
    _set(repo, TEST_PROGRESS, b, planned_finish=(MON + dt.timedelta(days=8)).isoformat())
    _set(repo, TEST_PROGRESS, c, planned_finish=MON.isoformat())                    # today counts
    d = build_data(wcfg, repo, MON)
    ids = {x["id"] for x in d["due_next_7_days"] if x["page"] == "Test Progress"}
    assert a in ids and c in ids and b not in ids
    assert set(d["due_next_7_days"][0]) == {"date", "page", "id", "name"}         # _iso removed


def test_t16_days_late_matches_overdue_list(wcfg, repo_data):                         # AC-11 data
    repo = repo_data
    late = _tp_ids(repo)[0]
    _set(repo, TEST_PROGRESS, late, status="In progress", total_tc=10, executed=9, fail_count=0,
         planned_start="2026-10-01", planned_finish=(MON - dt.timedelta(days=10)).isoformat())
    d = build_data(wcfg, repo, MON)
    row = next(a for a in d["delayed_or_at_risk"] if a["id"] == late)
    assert row["reason_kind"] == "overdue" and row["days_late"] == 10 and row["pct_time"] is None
    assert {t["id"]: t["days_late"] for t in d["overdue"]["test_progress"]}[late] == 10
    for a in d["delayed_or_at_risk"]:
        assert (a["days_late"] is not None) == (a["reason_kind"] == "overdue")
        if a["reason_kind"] == "overdue":
            tp = {t["id"]: t["days_late"] for t in d["overdue"]["test_progress"]}
            assert tp[a["id"]] == a["days_late"]


def test_t16_missing_data_sorted_unique_and_complete_excluded(wcfg, repo_data):       # AC-10 data
    repo = repo_data
    ids = _tp_ids(repo)
    z, y, comp = ids[0], ids[1], ids[2]
    with repo.transaction():
        repo.set_field(TEST_PROGRESS, z, "application", "Zulu app", NOW)
        repo.set_field(TEST_PROGRESS, y, "application", "Alpha app", NOW)
    _set(repo, TEST_PROGRESS, z, status="In progress", total_tc="")
    _set(repo, TEST_PROGRESS, y, status="In progress", total_tc=0, planned_start=None)
    _set(repo, TEST_PROGRESS, comp, status="Complete", total_tc=None, planned_finish=None)
    d = build_data(wcfg, repo, MON)
    md = d["missing_data"]
    assert [m["app"] for m in md] == sorted(m["app"] for m in md)
    by_id = {m["id"]: m for m in md}
    assert by_id[z]["fields"] == ["Total TC"] and by_id[y]["fields"] == ["Total TC", "Planned Start"]
    assert comp not in by_id                                                          # Complete wins
    assert len(md) == d["overall"]["colors"]["Missing data"]
    assert not {a["id"] for a in d["delayed_or_at_risk"]} & set(by_id)
    assert len(by_id) == len(md)


@pytest.mark.parametrize("rec, expected", [
    ({"total_tc": 0, "planned_start": "2026-11-01", "planned_finish": "2026-11-30"}, ["Total TC"]),
    ({"total_tc": "", "planned_start": "2026-11-01", "planned_finish": "2026-11-30"}, ["Total TC"]),
    ({"total_tc": None, "planned_start": "2026-11-01", "planned_finish": "2026-11-30"}, ["Total TC"]),
    ({"total_tc": 3, "planned_start": "", "planned_finish": None}, ["Planned Start", "Planned Finish"]),
    ({}, ["Total TC", "Planned Start", "Planned Finish"]),
    ({"total_tc": 1, "planned_start": "2026-11-01", "planned_finish": "2026-11-01"}, []),
])
def test_t16_missing_fields_edges(rec, expected):
    assert missing_fields(rec) == expected


@pytest.mark.parametrize("rec, color, reason", [
    ({"status": "Complete", "total_tc": None}, COMPLETE, "Status is Complete"),
    ({"total_tc": 0, "planned_start": "2026-11-01"}, MISSING, "Missing data: Total TC, Planned Finish"),
    ({"total_tc": 10, "executed": 5, "planned_start": "2026-11-01", "planned_finish": "2026-11-13"}, RED,
     "Past Planned Finish (13-Nov-2026) and not Complete"),
    ({"total_tc": 10, "executed": 5, "fail_count": 1, "planned_start": "2026-11-01",
      "planned_finish": "2026-11-30"}, RED, "%Fail 20.0% > 5.0%"),
    ({"total_tc": 100, "executed": 10, "fail_count": 0, "planned_start": "2026-11-01",
      "planned_finish": "2026-11-30"}, AMBER, "%Exec 10.0% < %Time 51.7%"),
    ({"total_tc": 100, "executed": 50, "fail_count": 1, "planned_start": "2026-11-01",
      "planned_finish": "2026-11-30"}, AMBER, "%Exec 50.0% < %Time 51.7%; #Fail 1 (2.0% <= 5.0%)"),
    ({"total_tc": 100, "executed": 90, "fail_count": 1, "planned_start": "2026-11-01",
      "planned_finish": "2026-11-30"}, AMBER, "#Fail 1 (1.1% <= 5.0%)"),
    ({"total_tc": 100, "executed": 90, "fail_count": 0, "planned_start": "2026-11-01",
      "planned_finish": "2026-11-30"}, GREEN, "%Exec 90.0% >= %Time 51.7%, no fails"),
])
def test_t16_compute_reason_unchanged(rec, color, reason):
    light = compute(rec, MON, 0.05)
    assert (light.color, light.reason) == (color, reason)


def test_t16_dashboard_still_14_days(wcfg, repo_data):
    import inspect
    assert inspect.signature(dashboard.build).parameters["days_ahead"].default == 14
    assert inspect.signature(build_data).parameters["days_ahead"].default == 7


# =========================================================================== T17: render

def test_ac4_subject():
    subject, _ = W.render(base_data(), None)
    assert subject == "[Planning Progress] รายงาน UAT ประจำสัปดาห์ DEMO UAT · 25-Nov-2026"


def test_ac5_language_headings_and_no_old_bilingual_text():
    d = base_data(
        overdue={"uat_plan": [{"id": "E-1", "task": "Task one", "finish_date": "10-Nov-2026", "days_late": 15}],
                 "test_progress": []},
        defects={"open": 2, "overdue": [{"id": "21", "system": "Sys", "detail": "Det", "status": "Open",
                                         "fix_target_date": "11-May-2026", "days_late": 198}],
                 "no_target": [{"id": "24", "system": "Sys2", "detail": "Det2", "status": "Open"}]})
    html = html_of(d)
    for en in ("Weekly UAT Report", "App status", "Apps to watch", "Overdue UAT Plan tasks", "Past fix target",
               "No fix target", "Due within 1 week", "Reschedules", "Other"):
        assert en in html, en
    for th in ("รายงาน UAT ประจำสัปดาห์ · DEMO UAT", "ข้อมูล ณ 25-Nov-2026", "สถานะแอป 7 แอป", "แอปที่ต้องติดตาม",
               "งาน UAT Plan ที่เลยกำหนด", "เลยวันเป้าหมาย", "ไม่มีวันเป้าหมาย", "ครบกำหนดใน 1 สัปดาห์", "อื่น ๆ"):
        assert th in html, th
    assert "Planning Progress" in html
    for old in ("Executed", "Overview", "Data as of", "ทดสอบแล้ว / Executed", "Delayed / At risk apps",
                "Delayed", "At risk", "On track", "Missing data", "Complete"):
        assert old not in html, old


def test_ac6_summary_line_spans():
    html = html_of()
    assert re.search(r'<span style="color:#DC2626[^"]*">ล่าช้า 2 แอป</span>', html)
    assert re.search(r'<span style="color:#B45309[^"]*">มีความเสี่ยงล่าช้า 3 แอป</span>', html)
    assert re.search(r'<span style="color:#9CA3AF[^"]*">ข้อมูลไม่ครบ 1 แอป</span>', html)
    box = section(html, "ล่าช้า 2 แอป", "ข้อมูลไม่ครบ 1 แอป")
    assert " · " in box


def test_ac7_kpis_use_data_numbers():
    html = html_of(base_data(defects={"open": 9, "overdue": [], "no_target": []}))
    assert "70.3%" in html and "260 จาก 370 เคส" in html
    k = section(html, "ทดสอบแล้ว", "สถานะแอป")
    assert re.search(r">Pass</div><div[^>]*>255</div>", k)
    assert re.search(r">Fail</div><div[^>]*>5</div>", k)
    assert re.search(r">Defect ที่เปิดอยู่</div><div[^>]*>9</div>", k)               # defects.open, not fail
    assert "width:70.3%" in k                                                          # % bar


def test_ac8_status_bar_proportions_and_legend():
    html = html_of()
    bar = section(html, "สถานะแอป 7 แอป", "แอปที่ต้องติดตาม")
    segs = re.findall(r'width:([\d.]+%);background:(#[0-9A-F]{6});height:12px', bar)
    assert segs == [("14.3%", HEX[COMPLETE]), ("28.6%", HEX[RED]), ("42.9%", HEX[AMBER]), ("14.3%", HEX[MISSING])]
    for th, n in (("เสร็จ", 1), ("ล่าช้า", 2), ("มีความเสี่ยงล่าช้า", 3), ("ตามแผน", 0), ("ข้อมูลไม่ครบ", 1)):
        assert f"{th} <b>{n}</b>" in bar, th


def test_ac8_all_zero_has_no_segments():
    d = base_data(overall={**base_data()["overall"], "apps": 0,
                           "colors": {"Complete": 0, "Delayed": 0, "At risk of delay": 0, "On track": 0,
                                      "Missing data": 0}}, delayed_or_at_risk=[], missing_data=[])
    html = html_of(d)
    assert "height:12px" not in html and "ตามแผน <b>0</b>" in html


REASONS = [
    ("AppLate", "เลยวันจบตามแผน (20-Nov-2026) มา 5 วัน และยังไม่เสร็จ"),
    ("AppFailHigh", "Fail 12.0% เกินเกณฑ์ 5.0%"),
    ("AppBehind", "ทดสอบได้ 40.0% แต่เวลาผ่านไปแล้ว 60.0%"),
    ("AppFailLow", "มี Fail 2 เคส (3.0%) แต่ยังไม่เกินเกณฑ์ 5.0%"),
    ("AppBoth", "ทดสอบได้ 94.0% แต่เวลาผ่านไปแล้ว 96.0% · มี Fail 1 เคส (1.1%) แต่ยังไม่เกินเกณฑ์ 5.0%"),
    ("AppNoTC", "ยังไม่ได้กรอก Total TC, Planned Finish"),
]


@pytest.mark.parametrize("app, reason", REASONS, ids=[r[0] for r in REASONS])
def test_ac9_reason_th_each_case(app, reason):
    html = html_of()
    watch = section(html, "แอปที่ต้องติดตาม", "งาน UAT Plan ที่เลยกำหนด")
    i = watch.index(f"<b>{app}</b>")
    nxt = watch.find("<tr>", i)
    assert reason in watch[i:nxt if nxt != -1 else None]


def test_ac9_threshold_comes_from_data():
    html = html_of(base_data(fail_threshold="8.0%"))
    assert "Fail 12.0% เกินเกณฑ์ 8.0%" in html and "เกินเกณฑ์ 5.0%" not in html


def test_ac9_reason_from_build_data_cases(wcfg, repo_data):
    """End to end: compute -> reason_kind -> Thai reason for real records."""
    repo = repo_data
    a, b, c, e = _tp_ids(repo)[:4]
    common = dict(status="In progress", planned_start="2026-11-01", planned_finish="2026-11-30")
    _set(repo, TEST_PROGRESS, a, **common, total_tc=100, executed=10, fail_count=0)     # behind time
    _set(repo, TEST_PROGRESS, b, **common, total_tc=100, executed=50, fail_count=1)     # behind + fail
    _set(repo, TEST_PROGRESS, c, **common, total_tc=100, executed=90, fail_count=1)     # fail ok
    _set(repo, TEST_PROGRESS, e, **common, total_tc=10, executed=5, fail_count=1)       # fail > X
    d = build_data(wcfg, repo, MON)
    rows = {r["id"]: r for r in d["delayed_or_at_risk"]}
    th = d["fail_threshold"]
    assert th == "5.0%"
    assert W._reason_th(rows[a], th) == "ทดสอบได้ 10.0% แต่เวลาผ่านไปแล้ว 51.7%"
    assert W._reason_th(rows[b], th) == ("ทดสอบได้ 50.0% แต่เวลาผ่านไปแล้ว 51.7% · "
                                         "มี Fail 1 เคส (2.0%) แต่ยังไม่เกินเกณฑ์ 5.0%")
    assert W._reason_th(rows[c], th) == "มี Fail 1 เคส (1.1%) แต่ยังไม่เกินเกณฑ์ 5.0%"
    assert W._reason_th(rows[e], th) == "Fail 20.0% เกินเกณฑ์ 5.0%"
    html = W.render(d, None)[1]
    for r in (rows[a], rows[b], rows[c], rows[e]):
        assert W._reason_th(r, th) in html


def test_ac10_missing_data_after_delayed_and_at_risk(wcfg, repo_data):
    html = html_of()
    watch = section(html, "แอปที่ต้องติดตาม", "งาน UAT Plan ที่เลยกำหนด")
    order = [watch.index(f"<b>{n}</b>") for n in
             ("AppLate", "AppFailHigh", "AppBehind", "AppFailLow", "AppBoth", "AppNoTC")]
    assert order == sorted(order)
    row = watch[order[-1]:]
    assert ">ข้อมูลไม่ครบ</span>" in row and ">-</td>" in row
    assert "ตัวเลขด้านขวาคือเคสที่ทดสอบแล้ว / ทั้งหมด" in watch
    assert ">10/40</td>" in watch                                                       # executed/total
    # build_data + render: an app without Total TC
    repo = repo_data
    no_tc = _tp_ids(repo)[0]
    _set(repo, TEST_PROGRESS, no_tc, status="In progress", total_tc=None)
    d = build_data(wcfg, repo, MON)
    html2 = W.render(d, None)[1]
    assert "ยังไม่ได้กรอก Total TC" in section(html2, "แอปที่ต้องติดตาม", "งาน UAT Plan ที่เลยกำหนด")


def test_ac11_tp_overdue_only_in_watch_list():
    html = html_of()
    assert html.count("AppLate") == 1
    assert "มา 5 วัน" in section(html, "แอปที่ต้องติดตาม", "งาน UAT Plan ที่เลยกำหนด")
    plan = section(html, "งาน UAT Plan ที่เลยกำหนด", ">Defect<")
    assert "ไม่มี" in plan and "TP-1" not in plan


def test_ac12_overdue_plan_task_and_empty():
    d = base_data(overdue={"uat_plan": [{"id": "E-11", "task": "DMS task", "finish_date": "10-Nov-2026",
                                         "days_late": 15}], "test_progress": []})
    plan = section(html_of(d), "งาน UAT Plan ที่เลยกำหนด", ">Defect<")
    assert "<b>E-11</b> DMS task" in plan and "กำหนด 10-Nov-2026" in plan
    assert re.search(r'color:#DC2626">เลยมา 15 วัน</td>', plan)
    assert "ไม่มี" not in plan
    empty = section(html_of(), "งาน UAT Plan ที่เลยกำหนด", ">Defect<")
    assert ">ไม่มี</div>" in empty


def test_ac13_defects_summary_only_and_with_lists():
    html = html_of()
    df = section(html, ">Defect<", "ครบกำหนดใน 1 สัปดาห์")
    assert "เปิดอยู่ <b>5</b> · เลยวันเป้าหมาย <b>0</b> · ไม่มีวันเป้าหมาย <b>0</b>" in df
    assert "Past fix target" not in df and "No fix target" not in df and "#" not in re.sub(r"#[0-9A-F]{6}", "", df)
    d = base_data(defects={"open": 5,
                           "overdue": [{"id": "21", "system": "DMS-RAMC", "detail": "Year check", "status": "Open",
                                        "fix_target_date": "11-May-2026", "days_late": 198}],
                           "no_target": [{"id": "24", "system": "Collins", "detail": "Accounts", "status": "New"}]})
    df = section(html_of(d), ">Defect<", "ครบกำหนดใน 1 สัปดาห์")
    assert "เปิดอยู่ <b>5</b> · เลยวันเป้าหมาย <b>1</b> · ไม่มีวันเป้าหมาย <b>1</b>" in df
    over = section(df, "Past fix target", "No fix target")
    assert "<b>#21</b> DMS-RAMC Year check" in over and "Open · เป้าหมาย 11-May-2026" in over
    assert "เลยมา 198 วัน" in over
    none = df[df.index("No fix target"):]
    assert "<b>#24</b> Collins Accounts" in none and "New" in none and "เลยมา" not in none


def test_ac13_only_one_subgroup_shown():
    d = base_data(defects={"open": 1, "overdue": [],
                           "no_target": [{"id": "24", "system": "Collins", "detail": "Accounts", "status": "New"}]})
    df = section(html_of(d), ">Defect<", "ครบกำหนดใน 1 สัปดาห์")
    assert "No fix target" in df and "Past fix target" not in df


def test_ac15_due_grouped_by_date_and_dots():
    due = [{"date": "26-Nov-2026", "page": "UAT Plan", "id": "TP-3", "name": "Plan task same id"},
           {"date": "26-Nov-2026", "page": "Test Progress", "id": "TP-3", "name": "AppBehindDue"},
           {"date": "28-Nov-2026", "page": "Test Progress", "id": "TP-9", "name": "AppGreen"}]
    html = html_of(base_data(due_next_7_days=due))
    sec = section(html, "ครบกำหนดใน 1 สัปดาห์", "อื่น ๆ")
    assert "ครบกำหนดใน 1 สัปดาห์ (3 งาน)" in sec
    assert sec.count("26-Nov-2026") == 1 and sec.count("28-Nov-2026") == 1
    assert sec.count("&#9679;") == 1                                                    # only the watched TP app
    i = sec.index("AppBehindDue")
    row_start = sec.rfind("<tr>", 0, i)
    assert f'color:{HEX[AMBER]}' in sec[row_start:i] and "&#9679;" in sec[row_start:i]
    assert "จุดสีหน้าชื่อคือแอปที่อยู่ในรายการต้องติดตาม" in sec
    assert ">UAT Plan</span>" in sec and ">Test Progress</span>" in sec


def test_ac15_due_empty():
    sec = section(html_of(), "ครบกำหนดใน 1 สัปดาห์", "อื่น ๆ")
    assert "(0 งาน)" in sec and ">ไม่มี</div>" in sec and "&#9679;" not in sec


def test_ac16_reschedules_none():
    sec = section(html_of(base_data(open_conflicts=2)), "อื่น ๆ", None)
    assert "การเลื่อนวัน 18-Nov-2026 ถึง 24-Nov-2026: ไม่มี" in sec
    assert "ข้อขัดแย้งที่ยังเปิด: <b>2</b>" in sec


def test_ac16_reschedules_app_and_excel():
    rs = [{"date": "20-Nov-2026", "page": "Test Progress", "id": "TP-3", "name": "AppBehind",
           "field": "planned_finish", "from": "20-Nov-2026", "to": "27-Nov-2026", "reason": "Waiting data",
           "where": "app"},
          {"date": "21-Nov-2026", "page": "UAT Plan", "id": "E-1", "name": "Plan task", "field": "finish_date",
           "from": "21-Nov-2026", "to": "28-Nov-2026", "reason": None, "where": "excel"}]
    sec = section(html_of(base_data(reschedules_last_7_days=rs)), "อื่น ๆ", None)
    assert "การเลื่อนวัน 18-Nov-2026 ถึง 24-Nov-2026" in sec and ": ไม่มี" not in sec
    assert "<b>AppBehind</b> Planned Finish 20-Nov-2026 → 27-Nov-2026" in sec
    assert "20-Nov-2026 · Test Progress · เหตุผล: Waiting data" in sec
    assert "<b>Plan task</b>" in sec and "21-Nov-2026 → 28-Nov-2026" in sec
    assert "21-Nov-2026 · UAT Plan · แก้ใน Excel ไม่ได้ระบุเหตุผล" in sec
    assert "ข้อขัดแย้งที่ยังเปิด: <b>0</b>" in sec


def test_ac18_email_safe_html_full_and_real(wcfg, repo_data):
    full = base_data(
        overdue={"uat_plan": [{"id": "E-1", "task": "T", "finish_date": "10-Nov-2026", "days_late": 15}],
                 "test_progress": []},
        defects={"open": 2, "overdue": [{"id": "1", "system": "S", "detail": "D", "status": "Open",
                                         "fix_target_date": "11-May-2026", "days_late": 9}],
                 "no_target": [{"id": "2", "system": "S", "detail": "D", "status": "Open"}]},
        due_next_7_days=[{"date": "26-Nov-2026", "page": "Test Progress", "id": "TP-3", "name": "X"}],
        reschedules_last_7_days=[{"date": "20-Nov-2026", "page": "Test Progress", "id": "TP-3", "name": "X",
                                  "field": "planned_finish", "from": "a", "to": "b", "reason": "r",
                                  "where": "app"}])
    htmls = [html_of(full), html_of(full, AI), W.render(build_data(wcfg, repo_data, MON), AI)[1]]
    for html in htmls:
        low = html.lower().replace(" ", "")
        assert "display:flex" not in low and "display:grid" not in low and "<style" not in low
        assert "—" not in html and "–" not in html
        assert 'width="640"' in html and "width:640px" in html
        assert "<table" in html


def test_ac19_numbers_from_build_data(wcfg, repo_data):
    d = build_data(wcfg, repo_data, MON)
    html = W.render(copy.deepcopy(d), None)[1]
    o = d["overall"]
    assert o["pct_executed"] in html and f"{o['executed']} จาก {o['total_tc']} เคส" in html
    assert f"สถานะแอป {o['apps']} แอป" in html
    for a in d["delayed_or_at_risk"]:
        assert f">{a['executed']}/{a['total_tc']}</td>" in html
    assert f"เปิดอยู่ <b>{d['defects']['open']}</b>" in html
    assert f"(<b>{len(d['due_next_7_days'])}</b>" in html or f"({len(d['due_next_7_days'])} งาน)" in html
    before = copy.deepcopy(d)
    W.render(d, None)
    assert d == before                                                                  # render does not mutate


def test_ac22_ai_block_thai_only_and_position():
    html = html_of(None, AI)
    i_sum = html.index("ข้อมูลไม่ครบ 1 แอป")
    i_ai = html.index("สรุปสำหรับผู้บริหาร (ร่างโดย AI)")
    i_kpi = html.index("ทดสอบแล้ว")
    assert i_sum < i_ai < html.index("สรุปไทยของสัปดาห์นี้") < i_kpi
    for s in ("Executive summary", "ความเสี่ยงหลัก", "Top risks", "สิ่งที่ควรติดตาม", "Suggested follow-ups",
              "หัวข้อความเสี่ยงไทย", "เหตุผลไทย", "สิ่งที่ต้องทำไทย", "เพราะไทย"):
        assert s in html, s
    for en in ("ENGLISH SUMMARY TEXT", "RISK TITLE EN", "WHY EN", "ACTION EN", "BECAUSE EN"):
        assert en not in html, en
    assert html.index("หัวข้อความเสี่ยงไทย") < i_kpi and html.index("สิ่งที่ต้องทำไทย") < i_kpi


def test_ac22_ai_empty_lists_and_escaped():
    ai = {"summary_th": "<script>x</script> & co", "summary_en": "", "risks": [], "follow_ups": []}
    html = html_of(None, ai)
    assert "สรุปสำหรับผู้บริหาร" in html and "<script>" not in html and "&lt;script&gt;" in html
    assert "ความเสี่ยงหลัก" not in html and "สิ่งที่ควรติดตาม" not in html


def test_ac22_ac23_no_ai():
    html = html_of(None, None)
    assert "สรุปสำหรับผู้บริหาร" not in html and "Executive summary" not in html
    assert "ส่วนสรุปร่างโดย AI (Claude)" not in html
    assert "ตัวเลขทั้งหมดคำนวณโดยระบบ Planning Progress ของ PMO (DEMO UAT) และส่งอัตโนมัติ" in html


def test_ac23_footer_with_ai_is_last():
    html = html_of(None, AI)
    foot = "ตัวเลขทั้งหมดคำนวณโดยระบบ Planning Progress ของ PMO (DEMO UAT) และส่งอัตโนมัติ"
    assert foot in html and "ส่วนสรุปร่างโดย AI (Claude)" in html
    assert html.index("ส่วนสรุปร่างโดย AI (Claude)") > html.index("ข้อขัดแย้งที่ยังเปิด")


def test_user_text_is_escaped():
    d = base_data(missing_data=[{"id": "TP-6", "app": "<b>Evil</b> & co", "fields": ["Total TC"]}])
    html = html_of(d)
    assert "<b>Evil</b> & co" not in html and "&lt;b&gt;Evil&lt;/b&gt; &amp; co" in html


def test_dry_run_banner_still_shown(wcfg, repo_data):
    g = W.generate(wcfg, repo_data, actor=PMO_ACTOR, as_of=MON, use_ai=False)
    s = FakeSender()
    assert W.send(wcfg, repo_data, s, g.report_id, actor=PMO_ACTOR) == "dry_run"
    html = s.sent[0].html
    assert s.sent[0].to == "pmo@example.test"
    assert html.index("<body") < html.index("DRY RUN") < html.index("รายงาน UAT ประจำสัปดาห์")


# =========================================================================== T18: PMO matters

def _ac17_setup(repo):
    with repo.transaction():
        repo.insert_review({"created_at": NOW, "sheet": TEST_PROGRESS, "item_id": _tp_ids(repo)[0],
                            "excel_row": 2, "field": "planned_start", "raw_value": "??", "token": "??",
                            "problem": "Unreadable date"})
        repo.set_field(TEST_PROGRESS, _tp_ids(repo)[1], "report_by_email", "bad-address", NOW)
    assert repo.open_reviews()


FORBIDDEN = ("No AI summary", "ไม่มีสรุปจาก AI", "คิวตรวจ", "review queue", "Review queue", "ไม่มีผู้รับ",
             "ข้ามอีเมลที่ไม่ถูกต้อง", "bad-address", "No Claude API key", "recipients")


def test_ac17_generate_keeps_internal_matters_out(wcfg, repo_data):
    repo = repo_data
    _ac17_setup(repo)
    g = W.generate(wcfg, repo, actor=PMO_ACTOR, as_of=MON)                # real path, no key -> no AI
    assert not g.ai_ok and "No Claude API key" in g.ai_error
    assert g.recipients == [] and len(g.invalid) == 1
    r = repo.get_weekly(g.report_id)
    for bad in FORBIDDEN:
        assert bad not in r["html"], bad
    assert "สรุปสำหรับผู้บริหาร" not in r["html"]
    assert json.loads(r["invalid_recipients"])[0]["value"] == "bad-address"     # still kept for preview
    assert json.loads(r["data_json"])["review_queue"] >= 1


def test_ac21_run_weekly_alerts_pmo_three_matters_dry_run(wcfg, repo_data):
    repo = repo_data
    _ac17_setup(repo)
    s = FakeSender()
    res = W.run_weekly(wcfg, repo, s, clock=dt.datetime(2026, 11, 16, 9), importer=lambda *a, **k: None)
    assert res.result == "ok" and res.send_result == "dry_run"
    subjects = [m.subject for m in s.sent]
    for t in ("Weekly report sent without the AI summary", "Weekly report: no recipients",
              "Weekly report: invalid email addresses skipped"):
        msgs = [m for m in s.sent if t in m.subject]
        assert len(msgs) == 1, (t, subjects)
        assert msgs[0].to == "pmo@example.test"
    inv = next(m for m in s.sent if "invalid email addresses" in m.subject)
    assert "skipped 1 invalid address(es)" in inv.html and "(bad-address)" in inv.html
    report = next(m for m in s.sent if "[DRY RUN]" in m.subject)
    for bad in FORBIDDEN[:7]:
        assert bad not in report.html.split("DRY RUN", 1)[1], bad


def test_ac21_invalid_list_limited_to_10(wcfg, repo_data):
    repo = repo_data
    bad = ";".join(f"bad{i:02d}" for i in range(1, 13))
    _set(repo, TEST_PROGRESS, _tp_ids(repo)[0], cc_emails=bad, report_by_email="ok@example.test")
    s = FakeSender()
    W.run_weekly(wcfg, repo, s, clock=dt.datetime(2026, 11, 16, 9), importer=lambda *a, **k: None,
                 ai_func=lambda d, model: weekly_ai.AiResult(True, {"summary_th": "", "summary_en": "",
                                                                    "risks": [], "follow_ups": []}, model=model))
    subjects = [m.subject for m in s.sent]
    assert not any("without the AI summary" in x for x in subjects)
    assert not any("no recipients" in x for x in subjects)
    inv = [m for m in s.sent if "invalid email addresses skipped" in m.subject]
    assert len(inv) == 1
    h = inv[0].html
    assert "skipped 12 invalid address(es)" in h and "and 2 more" in h
    assert "(bad10)" in h and "(bad11)" not in h and "(bad12)" not in h


def test_ac21_no_invalid_no_alert(wcfg, repo_data):
    repo = repo_data
    _set(repo, TEST_PROGRESS, _tp_ids(repo)[0], report_by_email="ok@example.test")
    s = FakeSender()
    W.run_weekly(wcfg, repo, s, clock=dt.datetime(2026, 11, 16, 9), importer=lambda *a, **k: None,
                 use_ai=False)
    assert not any("invalid email" in m.subject for m in s.sent)
    assert not any("without the AI summary" in m.subject for m in s.sent)             # AI off: no warning


@pytest.mark.parametrize("raw, expected", [
    (json.dumps([{"id": "TP-1", "app": "App A", "value": "x@"}, {"id": "TP-2", "app": None, "value": "y"}]),
     ["App A (x@)", "TP-2 (y)"]),
    ("[]", []), (None, []), ("", []), ("not json", []),
])
def test_preview_notes(raw, expected):
    assert W.preview_notes({"invalid_recipients": raw}) == expected
