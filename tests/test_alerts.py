"""Part 2: Traffic Light, working days, R1-R4, dry-run, logs, holidays, uploads, manual reminder."""

import dataclasses
import datetime as dt

import pytest

from uat_tracker.alerts.rules import R1, R2, R4A, R4B, evaluate
from uat_tracker.config import AlertsConfig
from uat_tracker.mail.sender import FakeSender
from uat_tracker.services.attachments import UploadError, upload
from uat_tracker.services.daily_job import JOB, run_daily
from uat_tracker.services.holidays import FetchError, parse_bot_response
from uat_tracker.services.importer import run_import
from uat_tracker.services.reminders import ReminderError, send_reminder
from uat_tracker.services.traffic_light import compute, r3_qualifies
from uat_tracker.services.workdays import WorkCalendar

from conftest import build_workbook
from conftest import PMO_ACTOR  # noqa: E402

D = dt.datetime
PMO = "pmo@example.test"

TP_HEADER = ["No.", "Application", "Report By", "Report By Email", "CC Email", "Total TC",
             "No. of TC Execution", "%Progress", "#Pass", "#Fail", "Planned Start", "Planned Finish",
             "Actual Start", "Actual Finish", "Traffic Light", "TC/Result Complete", "Status", "Remark"]


def tp(no, app, email="owner@bank.test", cc=None, total=10, executed=0, passed=None, failed=0,
       ps=None, pf=None, as_=None, af=None, tcr=None, status=None):
    passed = executed - (failed or 0) if passed is None else passed
    return [no, app, f"QA {app}", email, cc, total, executed, None, passed, failed, ps, pf, as_, af,
            None, tcr, status, None]


def dms():
    return tp(38, "DMS-RAMC", total=186, executed=126, passed=124, failed=2,
              ps=D(2026, 10, 22), pf=D(2026, 11, 13), as_=D(2026, 10, 22))


def fake_fetch(year):
    return [(f"{year}-01-01", "New Year's Day"), (f"{year}-12-31", "New Year's Eve")]


def failing_fetch(year):
    raise FetchError("BOT API not reachable: timeout")


@pytest.fixture
def acfg(cfg):
    return dataclasses.replace(cfg, alerts=AlertsConfig(pmo_email=PMO))


def write_tp(cfg, rows):
    build_workbook(cfg.excel_path, tp=rows, tp_header=TP_HEADER)


def run(cfg, repo, sender, day, **kw):
    kw.setdefault("fetcher", fake_fetch)
    return run_daily(cfg, repo, sender, clock=D.combine(day, dt.time(8, 30)), **kw)


def log(repo, rule=None, item=None):
    rows = list(reversed(repo.email_log(item_id=item, limit=1000)))
    return [r for r in rows if rule is None or r["rule"] == rule]


def rec(**kw):
    base = {"id": "TP-1", "active": 1, "status": None, "total_tc": 10, "executed": 0, "fail_count": 0,
            "planned_start": None, "planned_finish": None, "actual_start": None, "actual_finish": None,
            "tc_result_complete": None}
    base.update(kw)
    return base


# ------------------------------------------------------------------ A. Traffic Light
def test_acceptance_1_dms_ramc_is_amber():
    r = rec(total_tc=186, executed=126, fail_count=2, planned_start="2026-10-22", planned_finish="2026-11-13")
    light = compute(r, dt.date(2026, 11, 6), 0.05)
    assert light.color == "Amber"
    assert round(light.pct_exec, 3) == 0.677 and round(light.pct_time, 3) == 0.682
    assert round(light.pct_fail, 3) == 0.016


def test_acceptance_2_fail_10_of_100_is_red():
    r = rec(total_tc=100, executed=100, fail_count=10, planned_start="2026-11-01", planned_finish="2026-11-30")
    assert compute(r, dt.date(2026, 11, 10), 0.05).color == "Red"


@pytest.mark.parametrize("kw,today,color", [
    ({"status": "Complete"}, "2026-12-31", "Complete"),
    ({"planned_start": "2026-11-01", "planned_finish": "2026-11-10", "executed": 10}, "2026-11-11", "Red"),
    ({"planned_start": "2026-11-01", "planned_finish": "2026-11-11", "executed": 6}, "2026-11-06", "Green"),
    ({"planned_start": "2026-11-01", "planned_finish": "2026-11-11", "executed": 4}, "2026-11-06", "Amber"),
    ({"planned_start": "2026-11-01", "planned_finish": "2026-11-11", "executed": 8, "fail_count": None},
     "2026-11-06", "Green"),                                                     # blank #Fail = 0
    ({"planned_start": "2026-11-05", "planned_finish": "2026-11-05", "executed": 0}, "2026-11-05", "Amber"),
    ({"planned_start": "2026-11-05", "planned_finish": "2026-11-05", "executed": 0}, "2026-11-04", "Green"),
    ({"planned_start": None, "planned_finish": "2026-11-05"}, "2026-11-04", "Missing"),
    ({"planned_start": "2026-11-01", "planned_finish": "2026-11-05", "total_tc": 0}, "2026-11-04", "Missing"),
])
def test_traffic_light_rules(kw, today, color):
    assert compute(rec(**kw), dt.date.fromisoformat(today), 0.05).color == color


def test_r3_transitions():
    assert r3_qualifies("Amber", None) and r3_qualifies("Red", None)
    assert r3_qualifies("Amber", "Green") and r3_qualifies("Red", "Green") and r3_qualifies("Red", "Amber")
    assert not r3_qualifies("Amber", "Red")             # approved: no email Red -> Amber
    assert not r3_qualifies("Amber", "Amber") and not r3_qualifies("Green", "Amber")


# ------------------------------------------------------------------ B. working days / D. rules
def cal(*holidays):
    return WorkCalendar({dt.date.fromisoformat(h) for h in holidays}, {2026, 2027})


def test_acceptance_3_finish_on_saturday_reminds_friday_or_thursday():
    r = rec(planned_start="2026-11-02", planned_finish="2026-11-14", actual_start="2026-11-02")   # Sat
    fri, thu = dt.date(2026, 11, 13), dt.date(2026, 11, 12)
    assert [a.rule for a in evaluate(r, fri, cal())] == [R2]
    a = evaluate(r, fri, cal())[0]
    assert a.variant == "today" and a.shifted
    assert evaluate(r, thu, cal()) == []
    assert [a.rule for a in evaluate(r, thu, cal("2026-11-13"))] == [R2]          # Friday holiday
    assert [(a.rule, a.variant) for a in evaluate(r, dt.date(2026, 11, 16), cal())] == [(R2, "daily")]


def test_r2_today_skipped_when_actual_finish_filled():
    r = rec(planned_start="2026-11-02", planned_finish="2026-11-13", actual_start="2026-11-02",
            actual_finish="2026-11-12", executed=5)
    assert evaluate(r, dt.date(2026, 11, 13), cal()) == []


def test_r4_conditions_and_status_complete_stops_everything():
    r = rec(planned_start="2026-11-02", planned_finish="2026-11-13", actual_start="2026-11-02",
            actual_finish="2026-11-12", executed=10)
    day = dt.date(2026, 11, 16)
    assert {a.rule for a in evaluate(r, day, cal())} == {R4A, R4B}
    assert {a.rule for a in evaluate(r, day, cal(), has_upload=True)} == {R4A}
    assert {a.rule for a in evaluate({**r, "tc_result_complete": "Complete"}, day, cal())} == {R4A}
    late = {**r, "actual_finish": "2026-11-20"}                       # approved 6i: R4 anyway
    assert {a.rule for a in evaluate(late, dt.date(2026, 11, 23), cal())} == {R4A, R4B}
    assert evaluate({**r, "status": "Complete"}, day, cal()) == []
    assert evaluate({**r, "executed": 9}, day, cal()) == []


# ------------------------------------------------------------------ job: acceptance 1, 4, 5, 7, 8, 9
def test_acceptance_1_first_run_one_r3_email_second_run_none(acfg, repo):
    write_tp(acfg, [dms()])
    s = FakeSender()
    res = run(acfg, repo, s, dt.date(2026, 11, 6))
    assert res.result == "ok"
    rows = log(repo)
    assert [(r["rule"], r["result"]) for r in rows] == [("R3", "dry_run")]
    assert "progress is now Amber" in rows[0]["subject"]
    run(acfg, repo, s, dt.date(2026, 11, 9))
    assert len(log(repo, "R3")) == 1
    assert len(s.sent) == 1


def test_red_to_amber_no_email_reentry_after_green_emails(acfg, repo):
    item = tp(1, "APP", total=100, executed=100, failed=10, ps=D(2026, 11, 2), pf=D(2026, 11, 30),
              as_=D(2026, 11, 2))
    write_tp(acfg, [item])
    s = FakeSender()
    run(acfg, repo, s, dt.date(2026, 11, 9))                                     # Red
    item[9], item[8] = 2, 98
    write_tp(acfg, [item])
    run(acfg, repo, s, dt.date(2026, 11, 10))                                    # Amber (fail 2%)
    item[9], item[8] = 0, 100
    write_tp(acfg, [item])
    run(acfg, repo, s, dt.date(2026, 11, 11))                                    # Green
    item[9], item[8] = 1, 99
    write_tp(acfg, [item])
    run(acfg, repo, s, dt.date(2026, 11, 12))                                    # Amber again
    colors = [h["color"] for h in repo.light_history("TP-1")]
    assert colors == ["Red", "Amber", "Green", "Amber"]
    r3 = log(repo, "R3")
    assert [r["send_date"] for r in r3] == ["2026-11-09", "2026-11-12"]


def test_acceptance_4_three_r1_emails_none_on_weekend(acfg, repo):
    write_tp(acfg, [tp(1, "APP", ps=D(2026, 11, 5), pf=D(2026, 11, 30))])            # Thu
    s = FakeSender()
    for d in range(5, 10):                                                           # Thu..Mon
        run(acfg, repo, s, dt.date(2026, 11, d))
    r1 = log(repo, R1)
    assert [(r["send_date"], r["variant"]) for r in r1] == [
        ("2026-11-05", "today"), ("2026-11-06", "daily"), ("2026-11-09", "daily")]
    jobs = {j["run_date"]: j["result"] for j in repo.job_runs(JOB, 50)}
    assert jobs["2026-11-07"] == "skipped" and jobs["2026-11-08"] == "skipped"


def test_acceptance_5_r4_daily_until_status_complete(acfg, repo):
    row = tp(1, "APP", executed=10, ps=D(2026, 11, 2), pf=D(2026, 11, 13), as_=D(2026, 11, 2),
             af=D(2026, 11, 12), tcr="Complete")
    write_tp(acfg, [row])
    s = FakeSender()
    run(acfg, repo, s, dt.date(2026, 11, 16))
    run(acfg, repo, s, dt.date(2026, 11, 17))
    assert [r["send_date"] for r in log(repo, R4A)] == ["2026-11-16", "2026-11-17"]
    assert log(repo, R4B) == []                                   # TC/Result already Complete
    row[16] = "Complete"
    write_tp(acfg, [row])
    run(acfg, repo, s, dt.date(2026, 11, 18))
    assert len(log(repo, R4A)) == 2
    assert repo.get_item("Test Progress", "TP-1")["status"] == "Complete"   # set in Excel, never by app


def test_acceptance_6_upload_sets_complete_and_stops_reminders(acfg, repo, tmp_path):
    write_tp(acfg, [tp(1, "APP", executed=10, ps=D(2026, 11, 2), pf=D(2026, 11, 13),
                       as_=D(2026, 11, 2), af=D(2026, 11, 12))])
    s = FakeSender()
    run(acfg, repo, s, dt.date(2026, 11, 16))
    assert len(log(repo, R4B)) == 1
    f = tmp_path / "results.pdf"
    f.write_bytes(b"%PDF-1.4 test")
    res = upload(acfg, repo, "TP-1", f, actor=PMO_ACTOR)
    assert res.status_changed and res.stored_path.exists() and f.exists()
    assert repo.get_item("Test Progress", "TP-1")["tc_result_complete"] == "Complete"
    run(acfg, repo, s, dt.date(2026, 11, 17))                   # re-import keeps the app value
    assert repo.get_item("Test Progress", "TP-1")["tc_result_complete"] == "Complete"
    assert len(log(repo, R4B)) == 1 and len(log(repo, R4A)) == 2
    fields = [a["field"] for a in repo.audit_entries("TP-1")]
    assert "attachment" in fields and "tc_result_complete" in fields
    assert repo.pending("Test Progress")                       # waits for Part 3 write-back


def test_upload_pending_cleared_when_excel_matches(acfg, repo, tmp_path):
    row = tp(1, "APP", executed=10, ps=D(2026, 11, 2), pf=D(2026, 11, 13), as_=D(2026, 11, 2),
             af=D(2026, 11, 12))
    write_tp(acfg, [row])
    run_import(acfg, repo, editor="t")
    f = tmp_path / "r.xlsx"
    f.write_bytes(b"x")
    upload(acfg, repo, "TP-1", f, actor=PMO_ACTOR)
    row[15] = "Complete"
    write_tp(acfg, [row])
    run_import(acfg, repo, editor="t")
    assert repo.pending("Test Progress") == {}


def test_upload_validation(acfg, repo, tmp_path):
    write_tp(acfg, [tp(1, "APP")])
    run_import(acfg, repo, editor="t")
    bad = tmp_path / "x.exe"
    bad.write_bytes(b"x")
    with pytest.raises(UploadError, match="not allowed"):
        upload(acfg, repo, "TP-1", bad, actor=PMO_ACTOR)
    big = tmp_path / "big.pdf"
    big.write_bytes(b"x")
    small = dataclasses.replace(acfg, alerts=dataclasses.replace(acfg.alerts, attachment_max_mb=0))
    with pytest.raises(UploadError, match="limit"):
        upload(small, repo, "TP-1", big, actor=PMO_ACTOR)
    with pytest.raises(UploadError, match="No active"):
        upload(acfg, repo, "TP-99", big, actor=PMO_ACTOR)


def test_acceptance_7_missing_report_email_not_sent_and_flagged(acfg, repo):
    write_tp(acfg, [tp(1, "APP", email=None, ps=D(2026, 11, 9), pf=D(2026, 11, 30))])
    s = FakeSender()
    run(acfg, repo, s, dt.date(2026, 11, 9))
    rows = log(repo, R1)
    assert rows and all(r["result"] == "skipped" and "Report By Email" in r["error"] for r in rows)
    assert s.sent == []
    assert "missing_report_email" in {f["code"] for f in repo.active_flags() if f["item_id"] == "TP-1"}


def test_acceptance_8_dry_run_never_reaches_real_recipients(acfg, repo):
    write_tp(acfg, [tp(1, "APP", email="real@bank.test", cc="cc1@bank.test; cc2@bank.test",
                       ps=D(2026, 11, 9), pf=D(2026, 11, 9))])
    s = FakeSender()
    run(acfg, repo, s, dt.date(2026, 11, 9))
    assert s.sent
    for m in s.sent:
        assert m.to == PMO and m.cc == [] and m.subject.startswith("[DRY RUN]")
        assert "real@bank.test" in m.html and "cc1@bank.test" in m.html            # banner
    previews = list(acfg.preview_dir.rglob("*.html"))
    assert len(previews) == len(s.sent)
    assert all(r["mode"] == "dry_run" for r in log(repo))


def test_dry_run_without_pmo_email_writes_previews_only(cfg, repo):
    write_tp(cfg, [tp(1, "APP", ps=D(2026, 11, 9), pf=D(2026, 11, 30))])
    s = FakeSender(running=False)                      # Outlook not even needed
    res = run(cfg, repo, s, dt.date(2026, 11, 9))
    assert res.result == "ok" and s.sent == []
    assert {r["result"] for r in log(repo, R1)} == {"dry_run"}
    assert list(cfg.preview_dir.rglob("*.html"))


def test_live_mode_sends_to_report_by_and_cc(acfg, repo):
    live = dataclasses.replace(acfg, alerts=dataclasses.replace(acfg.alerts, dry_run=False))
    write_tp(live, [tp(1, "APP", email="real@bank.test", cc="a@bank.test;b@bank.test",
                       ps=D(2026, 11, 9), pf=D(2026, 11, 30))])
    s = FakeSender()
    run(live, repo, s, dt.date(2026, 11, 9))
    m = [m for m in s.sent if "testing starts today" in m.subject][0]
    assert m.to == "real@bank.test" and m.cc == ["a@bank.test", "b@bank.test"]
    assert not m.subject.startswith("[DRY RUN]")
    assert not list(live.preview_dir.rglob("*.html"))


def test_switch_to_live_treats_dry_run_colors_as_notified(acfg, repo):
    write_tp(acfg, [dms()])
    s = FakeSender()
    run(acfg, repo, s, dt.date(2026, 11, 6))                                   # dry-run R3
    live = dataclasses.replace(acfg, alerts=dataclasses.replace(acfg.alerts, dry_run=False))
    run(live, repo, s, dt.date(2026, 11, 9))
    assert [r["mode"] for r in log(repo, "R3")] == ["dry_run"]                # no live R3 for old color
    row = dms()
    row[9], row[8] = 20, 106                                                   # -> Red (fail 15.9%)
    write_tp(live, [row])
    run(live, repo, s, dt.date(2026, 11, 10))
    assert [(r["mode"], r["result"]) for r in log(repo, "R3")] == [("dry_run", "dry_run"), ("live", "sent")]


def test_acceptance_9_every_send_and_skip_in_email_and_audit_log(acfg, repo):
    write_tp(acfg, [dms(), tp(2, "NOMAIL", email=None, ps=D(2026, 11, 6), pf=D(2026, 11, 30))])
    run(acfg, repo, FakeSender(), dt.date(2026, 11, 6))
    rows = log(repo)
    assert {r["result"] for r in rows} == {"dry_run", "skipped"}
    audit = [a for a in repo.audit_entries() if (a["field"] or "").startswith("email")]
    assert len(audit) == len(rows)
    assert all(a["channel"] == "system" for a in audit)
    assert any(a["field"] == "traffic_light (calculated)" for a in repo.audit_entries("TP-38"))
    assert repo.verify_audit_chain()


def test_dedupe_same_rule_item_day(acfg, repo):
    write_tp(acfg, [tp(1, "APP", ps=D(2026, 11, 9), pf=D(2026, 11, 30))])
    s = FakeSender()
    run(acfg, repo, s, dt.date(2026, 11, 9))
    n = len(s.sent)
    assert run(acfg, repo, s, dt.date(2026, 11, 9)).message.startswith("Already ran")
    run(acfg, repo, s, dt.date(2026, 11, 9), force=True)
    assert len(s.sent) == n
    assert any(r["error"] and "dedupe" in r["error"] for r in log(repo, R1))


# ------------------------------------------------------------------ job guards
def test_outlook_not_running_skips_day(acfg, repo):
    write_tp(acfg, [dms()])
    res = run(acfg, repo, FakeSender(running=False), dt.date(2026, 11, 6))
    assert res.result == "skipped" and "Outlook" in res.message
    assert log(repo) == [] and repo.latest_lights() == {}


def test_late_scheduled_start_is_skipped(acfg, repo):
    write_tp(acfg, [dms()])
    res = run_daily(acfg, repo, FakeSender(), clock=D(2026, 11, 6, 9, 45), scheduled=True, fetcher=fake_fetch)
    assert res.result == "skipped" and "late" in res.message
    assert log(repo) == []


def test_missed_workdays_recorded_no_backfill(acfg, repo):
    write_tp(acfg, [tp(1, "APP", ps=D(2026, 11, 3), pf=D(2026, 11, 30))])
    s = FakeSender()
    run(acfg, repo, s, dt.date(2026, 11, 3))                                 # Tue
    res = run(acfg, repo, s, dt.date(2026, 11, 6))                           # Fri (Wed, Thu missed)
    assert res.missed_recorded == ["2026-11-04", "2026-11-05"]
    assert [r["send_date"] for r in log(repo, R1)] == ["2026-11-03", "2026-11-06"]


def test_bank_holiday_is_not_a_workday(acfg, repo):
    write_tp(acfg, [tp(1, "APP", ps=D(2026, 11, 9), pf=D(2026, 11, 30))])
    with repo.transaction():
        repo.upsert_holiday("2026-11-10", "Test holiday", "manual", "now")
    s = FakeSender()
    run(acfg, repo, s, dt.date(2026, 11, 9))
    res = run(acfg, repo, s, dt.date(2026, 11, 10))
    assert res.result == "skipped" and "workday" in res.message


def test_holiday_fetch_failure_keeps_cache_and_warns_pmo(acfg, repo):
    write_tp(acfg, [dms()])
    with repo.transaction():
        repo.upsert_holiday("2026-12-31", "NYE", "bot", "now")
    s = FakeSender()
    res = run(acfg, repo, s, dt.date(2026, 11, 9), fetcher=failing_fetch)     # Monday -> fetch
    assert res.result == "ok"
    assert "2026-12-31" in repo.holidays(2026)
    warn = log(repo, "PMO_ALERT")
    assert warn and all(w["to_addr"] == PMO for w in warn)
    assert {f["result"] for f in repo.holiday_fetches()} == {"failed"}


def test_no_holiday_data_sends_nothing_and_alerts_pmo(acfg, repo):
    write_tp(acfg, [dms()])
    s = FakeSender()
    res = run(acfg, repo, s, dt.date(2026, 11, 6), fetcher=failing_fetch)
    assert res.result == "failed"
    assert [r["rule"] for r in log(repo)] == ["PMO_ALERT", "PMO_ALERT"]
    assert all(m.to == PMO for m in s.sent)


def test_reimport_failure_sends_nothing_and_alerts_pmo(acfg, repo):
    write_tp(acfg, [dms()])
    s = FakeSender()

    def broken(*a, **k):
        raise PermissionError("file is locked")

    res = run(acfg, repo, s, dt.date(2026, 11, 6), importer=broken)
    assert res.result == "failed" and "locked" in res.message
    assert [r["rule"] for r in log(repo)] == ["PMO_ALERT"]


def test_parse_bot_response_shape():
    payload = {"result": {"api": "Financial Institutions' Holidays", "data": [
        {"HolidayWeekDay": "Friday", "Date": "2026-01-02", "HolidayDescription": "Special holiday"},
        {"HolidayWeekDay": "Monday", "Date": "2026-04-13", "HolidayDescription": "Songkran"}]}}
    assert parse_bot_response(payload) == [("2026-01-02", "Special holiday"), ("2026-04-13", "Songkran")]


# ------------------------------------------------------------------ flags + manual reminder
def test_new_flags_missing_data_and_late_finish(acfg, repo):
    write_tp(acfg, [tp(1, "NODATE", pf=D(2026, 11, 30)),
                    tp(2, "LATE", executed=10, ps=D(2026, 11, 2), pf=D(2026, 11, 13),
                       as_=D(2026, 11, 2), af=D(2026, 11, 16))])
    run_import(acfg, repo, editor="t")
    codes = {(f["item_id"], f["code"]) for f in repo.active_flags()}
    assert ("TP-1", "missing_data") in codes and ("TP-2", "actual_finish_late") in codes


def test_manual_reminder_dry_run_and_validation(acfg, repo):
    run_import(acfg, repo, editor="t")                     # default fixture workbook
    s = FakeSender()
    assert send_reminder(acfg, repo, s, "B-6", to="owner@bank.test", note="please update", actor=PMO_ACTOR) == "dry_run"
    assert s.sent[0].to == PMO and "B-6" in s.sent[0].subject
    assert send_reminder(acfg, repo, s, "DEF-21", to="it@bank.test", actor=PMO_ACTOR) == "dry_run"
    with pytest.raises(ReminderError):
        send_reminder(acfg, repo, s, "TP-38", to="x@bank.test", actor=PMO_ACTOR)
    with pytest.raises(ReminderError):
        send_reminder(acfg, repo, s, "B-6", to="not-an-email", actor=PMO_ACTOR)
    assert [r["rule"] for r in log(repo)] == ["MANUAL", "MANUAL"]
    assert any((a["field"] or "").startswith("email MANUAL") for a in repo.audit_entries("B-6"))
