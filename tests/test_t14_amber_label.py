"""T14 (spec 2026-11-25 AC-1, AC-2, AC-3, AC-3a): Amber shows as "At risk of delay" / "มีความเสี่ยงล่าช้า"."""

from __future__ import annotations

import re

import pytest

from uat_tracker.alerts import templates
from uat_tracker.alerts.rules import R3, Alert, R3Event
from uat_tracker.services import daily_job
from uat_tracker.services.traffic_light import AMBER, LABEL_EN, LABEL_TH, Light, label
from uat_tracker.ui import i18n

# "At risk" that is NOT followed by " of delay" -> the old word leaking somewhere
BARE_AT_RISK = re.compile(r"At risk(?! of delay)")
# "มีความเสี่ยง" that is NOT followed by "ล่าช้า"
BARE_TH = re.compile(r"มีความเสี่ยง(?!ล่าช้า)")


# --------------------------------------------------------------------------- AC-1

def test_ac1_label_tables_exact():
    assert LABEL_TH[AMBER] == "มีความเสี่ยงล่าช้า"
    assert LABEL_EN[AMBER] == "At risk of delay"
    assert label(AMBER) == "At risk of delay"
    assert label(AMBER, "th") == "มีความเสี่ยงล่าช้า"


def test_ac1_other_labels_unchanged():
    assert LABEL_EN == {"Red": "Delayed", "Amber": "At risk of delay", "Green": "On track",
                        "Complete": "Complete", "Missing": "Missing data"}


# --------------------------------------------------------------------------- AC-2

@pytest.fixture
def restore_lang():
    old = i18n.get_lang()
    yield
    i18n.set_lang(old)


@pytest.mark.parametrize("lang, word", [("th", "มีความเสี่ยงล่าช้า"), ("en", "At risk of delay")])
def test_ac2_i18n_exact(restore_lang, lang, word):
    i18n.set_lang(lang)
    assert i18n.t("tag.Amber") == word
    assert i18n.t("tp.amber") == word


def test_ac2_no_bare_old_word_in_i18n_table():
    for key, (en, th) in i18n.TEXT.items():
        assert not BARE_AT_RISK.search(en), key
        assert not BARE_TH.search(th), key


# --------------------------------------------------------------------------- AC-3

def test_ac3_amber_alert_email_has_full_th_en_word():
    alert = Alert(R3, None, "TP-9", event=R3Event(1, AMBER, None, "Behind schedule"))
    rec = {"id": "TP-9", "application": "APP-Y", "report_by": "QA", "executed": 2, "total_tc": 10}
    subject, body = templates.render(alert, rec, Light(AMBER, reason="Behind schedule"))
    assert "มีความเสี่ยงล่าช้า (At risk of delay)" in body
    assert subject.endswith("progress is now At risk of delay")
    assert not BARE_AT_RISK.search(subject) and not BARE_AT_RISK.search(body)
    assert not BARE_TH.search(body)
    assert templates.COLOR_TH[AMBER] == "มีความเสี่ยงล่าช้า (At risk of delay)"


# --------------------------------------------------------------------------- AC-3a

@pytest.fixture
def calls(monkeypatch):
    out: list[tuple[str, object]] = []

    def fake_edit(cfg, repo, sheet, item_id, values, **kw):
        out.append((item_id, values["traffic_light"]))
        return True

    monkeypatch.setattr(daily_job, "edit_item", fake_edit)
    return out


@pytest.mark.parametrize("cell, written", [
    ("At risk", True),               # previous word -> rewritten
    ("Amber", True),                 # oldest word -> rewritten
    ("", True),                      # empty cell -> written
    ("at risk of delay", True),      # different case is not the exact word -> rewritten
    ("At risk of delay", False),     # already the new word -> untouched
])
def test_ac3a_amber_cell_rewrite(calls, cell, written):
    n = daily_job._queue_traffic_lights(None, None, {"X": {"traffic_light": cell}}, {"X": Light(AMBER)},
                                        "2026-11-25T08:30:00+07:00")
    if written:
        assert calls == [("X", "At risk of delay")] and n == 1
    else:
        assert calls == [] and n == 0


def test_ac3a_second_run_does_not_rewrite(calls):
    items = {"X": {"traffic_light": "At risk"}}
    lights = {"X": Light(AMBER)}
    assert daily_job._queue_traffic_lights(None, None, items, lights, "t") == 1
    items["X"]["traffic_light"] = calls[-1][1]          # file now holds what was written
    assert daily_job._queue_traffic_lights(None, None, items, lights, "t") == 0
    assert calls == [("X", "At risk of delay")]


def test_ac3a_edit_error_is_skipped(monkeypatch):
    def boom(*a, **kw):
        raise daily_job.EditError("conflict")

    monkeypatch.setattr(daily_job, "edit_item", boom)
    n = daily_job._queue_traffic_lights(None, None, {"X": {"traffic_light": "At risk"}}, {"X": Light(AMBER)}, "t")
    assert n == 0
