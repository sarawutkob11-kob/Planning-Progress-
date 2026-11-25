"""T13/T14: Traffic Light words users see are Delayed / At risk of delay / On track
(Thai ล่าช้า / มีความเสี่ยงล่าช้า / ตามแผน).

Internal codes Red / Amber / Green stay in the DB, style keys and R3 rules; only the displayed words change.
"""

from __future__ import annotations

import dataclasses
import datetime as dt
import json
import re

import pytest

from uat_tracker.alerts import templates
from uat_tracker.alerts.rules import R3, Alert, R3Event
from uat_tracker.config import AlertsConfig
from uat_tracker.services import daily_job
from uat_tracker.services import weekly_report as W
from uat_tracker.services.importer import run_import
from uat_tracker.services.project_template import template
from uat_tracker.services.traffic_light import AMBER, COMPLETE, GREEN, MISSING, RED, Light, label
from uat_tracker.services.weekly_data import build_data
from uat_tracker.ui import i18n

OLD_EN = re.compile(r"\b(Red|Amber|Green)\b")
OLD_TH = ("แดง", "เหลือง", "เขียว")
MON = dt.date(2026, 11, 16)


def _no_old_words(text: str) -> bool:
    return not OLD_EN.search(text) and not any(w in text for w in OLD_TH)


# --------------------------------------------------------------------------- label table

def test_internal_codes_unchanged():
    assert (RED, AMBER, GREEN, COMPLETE, MISSING) == ("Red", "Amber", "Green", "Complete", "Missing")


@pytest.mark.parametrize("code, en, th", [
    (RED, "Delayed", "ล่าช้า"), (AMBER, "At risk of delay", "มีความเสี่ยงล่าช้า"), (GREEN, "On track", "ตามแผน"),
    (COMPLETE, "Complete", "เสร็จ"), (MISSING, "Missing data", "ข้อมูลไม่ครบ"),
])
def test_label_en_th(code, en, th):
    assert label(code) == en and label(code, "en") == en
    assert label(code, "th") == th


def test_label_unknown_and_empty():
    assert label("Something") == "Something" and label(None) == "" and label("") == ""


# --------------------------------------------------------------------------- UI words

@pytest.fixture
def restore_lang():
    old = i18n.get_lang()
    yield
    i18n.set_lang(old)


def test_ui_tags_and_tp_card(restore_lang):
    i18n.set_lang("en")
    assert [i18n.t(f"tag.{c}") for c in (RED, AMBER, GREEN)] == ["Delayed", "At risk of delay", "On track"]
    assert i18n.t("tp.amber") == "At risk of delay"
    assert i18n.t("tag.Complete") == "Complete" and i18n.t("tag.Missing") == "Missing data"
    i18n.set_lang("th")
    assert [i18n.t(f"tag.{c}") for c in (RED, AMBER, GREEN)] == ["ล่าช้า", "มีความเสี่ยงล่าช้า", "ตามแผน"]
    assert i18n.t("tp.amber") == "มีความเสี่ยงล่าช้า"


# --------------------------------------------------------------------------- email

@pytest.mark.parametrize("code, en, th", [(RED, "Delayed", "ล่าช้า"),
                                           (AMBER, "At risk of delay", "มีความเสี่ยงล่าช้า")])
def test_r3_email_uses_new_words(code, en, th):
    alert = Alert(R3, None, "TP-1", event=R3Event(1, code, None, "Past Planned Finish"))
    rec = {"id": "TP-1", "application": "APP-X", "report_by": "QA", "executed": 5, "total_tc": 10}
    subject, body = templates.render(alert, rec, Light(code, reason="Past Planned Finish"))
    assert subject.endswith(f"progress is now {en}")
    assert f"{th} ({en})" in body and f"is now {en}." in body
    assert f">{en}</span>" in body                                       # badge in the fact table
    assert _no_old_words(subject) and _no_old_words(body)


def test_email_badge_and_color_th():
    assert templates.COLOR_TH == {RED: "ล่าช้า (Delayed)", AMBER: "มีความเสี่ยงล่าช้า (At risk of delay)",
                                  GREEN: "ตามแผน (On track)", COMPLETE: "เสร็จ (Complete)"}
    assert ">On track</span>" in templates._color_badge(GREEN)
    assert templates._color_badge(None) == "-"


# --------------------------------------------------------------------------- weekly report

@pytest.fixture
def wdata(cfg, repo):
    wcfg = dataclasses.replace(cfg, alerts=AlertsConfig(pmo_email="pmo@example.test"))
    run_import(wcfg, repo, editor="t")
    return wcfg, repo


def test_weekly_data_sent_to_ai_has_labels_not_codes(wdata):
    wcfg, repo = wdata
    d = build_data(wcfg, repo, MON)
    assert "red_amber" not in d and d["delayed_or_at_risk"]
    assert set(d["overall"]["colors"]) == {"Complete", "Delayed", "At risk of delay", "On track", "Missing data"}
    assert {a["color"] for a in d["delayed_or_at_risk"]} <= {"Delayed", "At risk of delay"}
    assert _no_old_words(json.dumps(d, ensure_ascii=False))


def test_weekly_report_html_uses_new_words(wdata):
    wcfg, repo = wdata
    d = build_data(wcfg, repo, MON)
    _, html = W.render(d, None)
    assert "แอปที่ต้องติดตาม" in html and "Apps to watch" in html                  # 25-Nov-2026 layout
    for code in (COMPLETE, RED, AMBER, GREEN, MISSING):                             # status-bar legend, Thai only
        assert f"{label(code, 'th')} <b>{d['overall']['colors'][label(code)]}</b>" in html
    th_of = {label(c): label(c, "th") for c in (RED, AMBER)}
    for a in d["delayed_or_at_risk"]:
        assert f">{th_of[a['color']]}</span>" in html                               # badge in the watch list
    assert "At risk of delay" not in html and "Delayed" not in html               # English labels not in the email
    assert _no_old_words(html)


# --------------------------------------------------------------------------- Excel column

def test_queue_traffic_lights_writes_new_words(monkeypatch):
    calls = []

    def fake_edit(cfg, repo, sheet, item_id, values, **kw):
        calls.append((item_id, values["traffic_light"]))
        return True

    monkeypatch.setattr(daily_job, "edit_item", fake_edit)
    items = {"A": {"traffic_light": "Amber"},          # old word in the file -> rewritten
             "B": {"traffic_light": "At risk of delay"},  # already the new word -> untouched
             "G": {"traffic_light": "At risk"},           # T14: previous word -> rewritten
             "C": {"traffic_light": "Red"},            # missing data -> cell left as it is
             "D": {"traffic_light": None},
             "E": {"traffic_light": "Complete"},
             "F": {"traffic_light": "Green"}}
    lights = {"A": Light(AMBER), "B": Light(AMBER), "C": Light(MISSING), "D": Light(RED),
              "E": Light(COMPLETE), "F": Light(GREEN), "G": Light(AMBER)}
    n = daily_job._queue_traffic_lights(None, None, items, lights, "2026-11-25T08:30:00+07:00")
    assert calls == [("A", "At risk of delay"), ("D", "Delayed"), ("F", "On track"),
                     ("G", "At risk of delay")] and n == 4


# --------------------------------------------------------------------------- new-project template

def test_template_dropdown_and_cond_formats_use_new_words():
    tp = next(s for s in template()["sheets"] if s["name"] == "Test Progress")
    assert [v["formula1"] for v in tp["validations"] if v["column"] == "Traffic Light"] == [
        '"On track,At risk of delay,Delayed"']                          # T15
    cf = {f["formula"][0]: f["dxf"]["fill_bg"] for f in tp["cond_formats"] if f["first"] == "Traffic Light"
          and f["type"] == "cellIs"}
    assert cf['"At risk of delay"'] == {"rgb": "FFFFFF00"}               # colors unchanged
    assert '"At risk"' not in cf
    assert cf['"Delayed"'] == {"rgb": "FFFF0000"}
    assert '"On track"' in cf
    assert not any(k in cf for k in ('"Red"', '"Amber"', '"Green"'))


def test_template_legend_uses_new_words():
    rows = template()["sheets"][1]["legend"]["rows"]
    assert [r[0]["value"] for r in rows[1:4]] == ["On track", "At risk of delay", "Delayed"]
    assert [r[1]["value"] for r in rows[1:4]] == ["On Track", "เริ่มช้า, มีโอกาส Delay", "Delay"]  # unchanged
