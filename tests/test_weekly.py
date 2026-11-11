"""Part 6: AI weekly report. The Claude API is never called for real here (fake client / fake AI)."""

from __future__ import annotations

import dataclasses
import datetime as dt
import json
from types import SimpleNamespace

import pytest

from conftest import PMO_ACTOR, add_test_user
from uat_tracker.config import AlertsConfig
from uat_tracker.excel.sheets import DEFECT_LIST, TEST_PROGRESS, UAT_PLAN
from uat_tracker.mail.sender import FakeSender
from uat_tracker.services import weekly_ai, weekly_report as W
from uat_tracker.services.auth import PermissionDenied
from uat_tracker.services.dashboard import build as dashboard_build
from uat_tracker.services.edits import edit_item
from uat_tracker.services.importer import run_import
from uat_tracker.services.settings import save_ai_settings
from uat_tracker.services.traffic_light import compute
from uat_tracker.services.weekly_data import build_data

MON = dt.date(2026, 11, 16)
KEY = "sk-ant-api03-TESTKEY-never-store-me-1234567890"
assert MON.weekday() == 0


class MemKeys:
    def __init__(self, key=None):
        self.key = key

    def get(self):
        return self.key

    def set(self, k):
        self.key = k

    def clear(self):
        self.key = None


@pytest.fixture(autouse=True)
def no_real_keystore(monkeypatch):
    monkeypatch.setattr(weekly_ai, "KEYS", MemKeys())          # never touch the real Credential Manager


@pytest.fixture
def wcfg(cfg):
    return dataclasses.replace(cfg, alerts=AlertsConfig(pmo_email="pmo@example.test"))


@pytest.fixture
def data(wcfg, repo):
    run_import(wcfg, repo, editor="t")
    with repo.transaction():                                   # 2026 holiday data present
        repo.upsert_holiday("2026-12-07", "King Bhumibol Memorial Day", "manual", "2026-11-01T00:00:00+07:00")
    return repo


def fake_ai_ok(data, model):
    o = data["overall"]
    return weekly_ai.AiResult(True, {
        "summary_th": f"ทดสอบแล้ว {o['executed']} จาก {o['total_tc']} ({o['pct_executed']})",
        "summary_en": f"{o['executed']} of {o['total_tc']} executed ({o['pct_executed']}).",
        "risks": [{"title_th": "เลยกำหนด", "title_en": "Overdue", "why_th": "x", "why_en": "y", "related": ["DMS-RAMC"]}],
        "follow_ups": []}, model=model, input_tokens=9000, output_tokens=1500,
        cost_usd=weekly_ai.cost_of(model, 9000, 1500))


def fake_ai_down(data, model):
    return weekly_ai.AiResult(False, error="Cannot reach the Claude API: connection refused", model=model)


def no_import(*a, **k):
    return None


def run(wcfg, repo, sender, day, **kw):
    kw.setdefault("ai_func", fake_ai_ok)
    return W.run_weekly(wcfg, repo, sender, clock=dt.datetime.combine(day, dt.time(9, 0)), importer=no_import, **kw)


# --------------------------------------------------------------------------- 4: numbers = DB

def test_numbers_match_database(wcfg, data):
    repo = data
    d = build_data(wcfg, repo, MON)
    tp = [r for r in repo.load_items(TEST_PROGRESS).values() if r["active"]]
    o = d["overall"]
    assert o["executed"] == sum(int(r["executed"] or 0) for r in tp) == 178
    assert o["total_tc"] == sum(int(r["total_tc"] or 0) for r in tp) == 238
    assert o["pct_executed"] == f"{178 / 238 * 100:.1f}%"
    assert o["pass"] == sum(int(r["pass_count"] or 0) for r in tp) and o["fail"] == 2
    colors = [compute(r, MON, wcfg.alerts.fail_threshold).color for r in tp]
    assert len(d["red_amber"]) == sum(c in ("Red", "Amber") for c in colors)
    assert sum(o["colors"].values()) == o["apps"] == len(tp)
    dash = dashboard_build(wcfg, repo, MON)
    assert len(d["overdue"]["uat_plan"]) == dash.overdue_tasks
    assert d["defects"]["open"] == dash.open_defects
    assert len(d["defects"]["no_target"]) == dash.no_target_defects
    assert len(d["defects"]["overdue"]) == dash.overdue_defects
    assert d["review_queue"] == len(repo.open_reviews()) and d["open_conflicts"] == len(repo.open_conflicts())
    blob = json.dumps(d, ensure_ascii=False)
    assert "@" not in blob and "QA Somsak" not in blob                    # no people / emails to the AI


def test_reschedules_last_7_days_with_reason(wcfg, data):
    repo = data
    edit_item(wcfg, repo, TEST_PROGRESS, "TP-38", {"planned_finish": "2026-11-20"}, actor=PMO_ACTOR,
              reason="Waiting data", root_cause="Waiting other system data", now="2026-11-13T10:00:00+07:00")
    edit_item(wcfg, repo, UAT_PLAN, "E-11", {"start_date": "2026-10-21"}, actor=PMO_ACTOR, reason="old",
              now="2026-11-02T10:00:00+07:00")                                 # outside the 7 days
    rs = build_data(wcfg, repo, MON)["reschedules_last_7_days"]
    assert [(x["id"], x["from"], x["to"], x["where"]) for x in rs] == [("TP-38", "13-Nov-2026", "20-Nov-2026", "app")]
    assert "Waiting data" in rs[0]["reason"] and "Waiting other system data" in rs[0]["reason"]


# --------------------------------------------------------------------------- 3: recipients

def test_recipients_deduped_invalid_skipped_and_flagged(wcfg, data):
    repo = data
    edit_item(wcfg, repo, TEST_PROGRESS, "TP-38", {"report_by_email": "QA.A@example.test",
                                                   "cc_emails": "pm@example.test; qa.a@example.test"}, actor=PMO_ACTOR)
    edit_item(wcfg, repo, TEST_PROGRESS, "TP-6", {"report_by_email": "pm@example.test"}, actor=PMO_ACTOR)
    with repo.transaction():                                                   # typed in Excel, not validated
        repo.set_field(TEST_PROGRESS, "TP-12", "report_by_email", "not an address", "2026-11-10T00:00:00+07:00")
    to, bad = W.recipients(repo, flag=True)
    assert to == ["pm@example.test", "qa.a@example.test"] or sorted(to) == ["pm@example.test", "qa.a@example.test"]
    assert len(to) == 2 and {b["id"] for b in bad} == {"TP-12"}
    assert any(f["code"] == W.FLAG_BAD_EMAIL and f["item_id"] == "TP-12" for f in repo.active_flags())


# --------------------------------------------------------------------------- 5: AI never adds numbers

def test_ai_number_check():
    data = {"overall": {"executed": 178, "total_tc": 238, "pct_executed": "74.8%"}, "as_of": "16-Nov-2026"}
    ok = {"summary_th": "ทดสอบแล้ว 178 จาก 238 (74.8%) ณ 16-Nov-2026", "summary_en": "", "risks": [], "follow_ups": []}
    assert weekly_ai.unknown_numbers(ok, data) == []
    bad = {**ok, "summary_en": "About 75% done, 60 left", "risks": [
        {"title_th": "", "title_en": "", "why_th": "", "why_en": "3 apps", "related": []}]}
    assert weekly_ai.unknown_numbers(bad, data) == ["3", "60", "75"]


def _fake_client(text, *, stop="end_turn", model="claude-opus-5-5", raise_exc=None):
    calls = []

    def create(**kw):
        calls.append(kw)
        if raise_exc:
            raise raise_exc
        return SimpleNamespace(content=[SimpleNamespace(type="text", text=text)], stop_reason=stop, model=model,
                               usage=SimpleNamespace(input_tokens=8000, output_tokens=2000,
                                                     cache_read_input_tokens=0, cache_creation_input_tokens=0))
    return SimpleNamespace(beta=SimpleNamespace(messages=SimpleNamespace(create=create))), calls


def test_summarize_request_shape_cost_and_number_rejection():
    data = {"overall": {"executed": 178, "total_tc": 238}, "red_amber": [{"app": "DMS-RAMC"}]}
    good = json.dumps({"summary_th": "178 จาก 238", "summary_en": "178 of 238", "follow_ups": [],
                       "risks": [{"title_th": "a", "title_en": "b", "why_th": "c", "why_en": "d",
                                  "related": ["DMS-RAMC", "Made-up app"]}]})
    client, calls = _fake_client(good)
    r = weekly_ai.summarize(data, model="claude-opus-5-5", key=KEY, client=client)
    assert r.ok and r.content["risks"][0]["related"] == ["DMS-RAMC"]          # unknown names dropped
    kw = calls[0]
    assert kw["model"] == "claude-opus-5-5" and kw["fallbacks"] == "default"
    assert kw["betas"] == ["server-side-fallback-2026-07-01"]
    assert kw["output_config"]["format"]["schema"] == weekly_ai.SCHEMA and kw["output_config"]["effort"] == "medium"
    assert "thinking" not in kw and "temperature" not in kw
    assert r.cost_usd == pytest.approx(8000 * 4 / 1e6 + 2000 * 20 / 1e6)
    client, calls = _fake_client(good)
    weekly_ai.summarize(data, model="claude-haiku-5-5", key=KEY, client=client)
    assert "fallbacks" not in calls[0]                                        # Haiku: no server-side fallback
    client, _ = _fake_client(json.dumps({"summary_th": "ผ่าน 99%", "summary_en": "", "risks": [], "follow_ups": []}))
    r = weekly_ai.summarize(data, key=KEY, client=client)
    assert not r.ok and "99" in r.error and r.cost_usd is not None           # rejected, still costed
    client, _ = _fake_client("", stop="refusal")
    assert "refusal" in weekly_ai.summarize(data, key=KEY, client=client).error


def test_api_errors_never_raise_and_never_show_the_key():
    data = {"overall": {"executed": 1}}
    client, _ = _fake_client("", raise_exc=RuntimeError(f"boom with {KEY}"))
    r = weekly_ai.summarize(data, key=KEY, client=client)
    assert not r.ok and KEY not in r.error and "***" in r.error
    r = weekly_ai.summarize(data, key="")                                     # no key saved
    assert not r.ok and "No Claude API key" in r.error


# --------------------------------------------------------------------------- 6: key never in DB / logs / Excel

def test_api_key_only_in_credential_store(wcfg, data, tmp_path):
    repo = data
    save_ai_settings(repo, PMO_ACTOR, model="claude-opus-5-5", api_key=KEY)
    assert weekly_ai.KEYS.get() == KEY
    client, calls = _fake_client(json.dumps({"summary_th": "", "summary_en": "", "risks": [], "follow_ups": []}))
    ai = lambda d, model: weekly_ai.summarize(d, model=model, client=client)
    run(wcfg, repo, FakeSender(), MON, ai_func=ai)
    dump = "\n".join(str(tuple(r)) for t in ("settings", "audit_log", "email_log", "weekly_report", "job_run")
                     for r in repo.conn.execute(f"SELECT * FROM {t}"))
    assert KEY not in dump and "API key updated" in dump
    for f in list(wcfg.data_dir.rglob("*")) + [wcfg.excel_path]:
        if f.is_file():
            assert KEY.encode() not in f.read_bytes(), f


# --------------------------------------------------------------------------- 1, 2: schedule

def test_monday_workday_sent_dry_run_to_pmo_only(wcfg, data):
    repo = data
    edit_item(wcfg, repo, TEST_PROGRESS, "TP-38", {"report_by_email": "qa@example.test"}, actor=PMO_ACTOR)
    s = FakeSender()
    r = run(wcfg, repo, s, MON)
    assert r.result == "ok" and r.send_result == "dry_run"
    assert [m.to for m in s.sent] == ["pmo@example.test"] and "[DRY RUN]" in s.sent[0].subject
    assert "qa@example.test" in s.sent[0].html                                   # real To shown in the banner only
    w = repo.get_weekly(r.report_id)
    assert w["status"] == "dry_run" and w["ai_used"] == 1 and w["trigger"] == "scheduled"
    assert "178 / 238 (74.8%)" in w["html"]
    assert run(wcfg, repo, s, MON + dt.timedelta(days=1)).message == "Weekly report already sent this week"


def test_monday_holiday_moves_to_next_workday(wcfg, data):
    repo = data
    with repo.transaction():
        repo.upsert_holiday(MON.isoformat(), "Holiday", "manual", "2026-11-01T00:00:00+07:00")
    s = FakeSender()
    assert "Not the weekly report day" in run(wcfg, repo, s, MON).message
    tue = run(wcfg, repo, s, MON + dt.timedelta(days=1))
    assert tue.result == "ok" and tue.send_result == "dry_run"
    assert run(wcfg, repo, s, MON + dt.timedelta(days=2)).result == "skipped"
    assert "Not the weekly report day" in run(wcfg, repo, s, MON + dt.timedelta(days=9)).message  # a normal Wed


def test_no_holiday_data_not_sent_pmo_warned(wcfg, repo):
    def fail(year):
        from uat_tracker.services.holidays import FetchError
        raise FetchError("no key")
    s = FakeSender()
    r = W.run_weekly(wcfg, repo, s, clock=dt.datetime(2026, 11, 16, 9), importer=no_import, ai_func=fake_ai_ok,
                     fetcher=fail)
    assert r.result == "failed" and "No bank-holiday data" in r.message
    assert [m.subject for m in s.sent] == ["[Planning Progress] No holiday data - weekly report not sent"]


def test_outlook_closed_skipped_not_backfilled(wcfg, data):
    r = run(wcfg, data, FakeSender(running=False), MON)
    assert r.result == "skipped" and "Outlook" in r.message and data.weekly_reports() == []


# --------------------------------------------------------------------------- 7: API down -> without AI + PMO warned

def test_api_down_report_sent_without_ai(wcfg, data):
    s = FakeSender()
    r = run(wcfg, data, s, MON, ai_func=fake_ai_down)
    w = data.get_weekly(r.report_id)
    assert r.send_result == "dry_run" and w["ai_used"] == 0 and "connection refused" in w["ai_error"]
    assert "No AI summary" in w["html"] and "178 / 238" in w["html"]
    assert any("without the AI summary" in m.subject for m in s.sent)


def test_live_mode_recipients_and_none_goes_to_pmo(wcfg, data):
    live = dataclasses.replace(wcfg, alerts=dataclasses.replace(wcfg.alerts, dry_run=False))
    s = FakeSender()
    r = run(live, data, s, MON)                                                 # no addresses at all yet
    assert r.send_result == "sent" and s.sent[0].to == "pmo@example.test"
    assert any("no recipients" in m.subject for m in s.sent)
    edit_item(live, data, TEST_PROGRESS, "TP-38", {"report_by_email": "qa@example.test", "cc_emails": "pm@example.test"},
              actor=PMO_ACTOR)
    g = W.generate(live, data, actor=PMO_ACTOR, as_of=MON, ai_func=fake_ai_ok)
    s2 = FakeSender()
    assert W.send(live, data, s2, g.report_id, actor=PMO_ACTOR) == "sent"
    assert s2.sent[0].to == "qa@example.test; pm@example.test"


# --------------------------------------------------------------------------- 8, 9: manual, roles, history

def test_manual_preview_send_cancel_and_qa_cannot(wcfg, data):
    repo = data
    g = W.generate(wcfg, repo, actor=PMO_ACTOR, as_of=MON, ai_func=fake_ai_ok)
    assert repo.get_weekly(g.report_id)["status"] == "preview"                # nothing sent yet
    s = FakeSender()
    assert W.send(wcfg, repo, s, g.report_id, actor=PMO_ACTOR) == "dry_run"
    with pytest.raises(ValueError):
        W.send(wcfg, repo, s, g.report_id, actor=PMO_ACTOR)                   # not twice
    g2 = W.generate(wcfg, repo, actor=PMO_ACTOR, as_of=MON, ai_func=fake_ai_ok)
    W.cancel(repo, g2.report_id, actor=PMO_ACTOR)
    qa = add_test_user(repo, "qa@example.test", "QA")
    bu = add_test_user(repo, "bu@example.test", "User BU/PM")
    with pytest.raises(PermissionDenied):
        W.generate(wcfg, repo, actor=qa, as_of=MON, ai_func=fake_ai_ok)
    g3 = W.generate(wcfg, repo, actor=bu, as_of=MON, ai_func=fake_ai_ok)
    with pytest.raises(PermissionDenied):
        W.send(wcfg, repo, s, g3.report_id, actor=qa)
    hist = repo.weekly_reports()
    assert [h["status"] for h in hist] == ["preview", "cancelled", "dry_run"]
    assert hist[-1]["created_by"] == "Pim Pmo <pmo@example.test>" and hist[-1]["cost_usd"] == pytest.approx(0.066)
    assert any(e["field"] == "denied" and e["new_value"] == "weekly" for e in repo.audit_entries())
