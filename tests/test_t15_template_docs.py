"""T15 (spec 2026-11-25 R1a/R1b, AC-3b): template, README and docstrings use "At risk of delay"."""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl
import pytest

from uat_tracker.config import Config
from uat_tracker.services import projects
from uat_tracker.services import weekly_data
from uat_tracker.services.project_template import template

from conftest import WINDOW

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "src" / "uat_tracker" / "resources" / "project_template.json"
OLD = re.compile(r"At risk(?! of delay)")
OLD_TH = re.compile(r"มีความเสี่ยง(?!ล่าช้า)")
PERSON = projects.Person("pmo@example.test", "Pim", "Pmo")


def test_template_raw_has_no_old_word():
    text = TEMPLATE.read_text(encoding="utf-8")
    assert not OLD.search(text)
    assert "At risk of delay" in text


def test_template_sample_row_uses_new_word():
    text = TEMPLATE.read_text(encoding="utf-8")
    assert '"value": "At risk of delay"' in text


def test_weekly_data_docstring_no_old_word():
    doc = weekly_data.__doc__ or ""
    assert "At risk of delay" in doc
    assert not OLD.search(doc)
    src = Path(weekly_data.__file__).read_text(encoding="utf-8")
    assert not OLD.search(src)


def test_readme_no_old_words():
    text = (ROOT / "README.md").read_text(encoding="utf-8")
    assert not OLD.search(text)
    assert not OLD_TH.search(text)
    assert "At risk of delay" in text and "มีความเสี่ยงล่าช้า" in text


@pytest.fixture
def base(tmp_path) -> Config:
    root = tmp_path / "root"
    return Config(excel_path=tmp_path / "unused.xlsx", data_dir=root, window_start=WINDOW.start,
                  window_end=WINDOW.end, profile_file=root / "profile.json")


def test_new_project_xlsx_dropdown_and_cond_formats(base, tmp_path):
    db = projects.open_app_db(base)
    try:
        folder = tmp_path / "f"
        folder.mkdir()
        p = projects.create_project(base, db, PERSON, name="T15", folder=folder)
    finally:
        db.close()
    wb = openpyxl.load_workbook(p["excel_path"])
    tp = wb["Test Progress"]
    assert [dv.formula1 for dv in tp.data_validations.dataValidation] == ['"On track,At risk of delay,Delayed"']
    rules = [r for cf in tp.conditional_formatting for r in cf.rules]
    assert len(rules) == 5
    by_formula = {r.formula[0]: r for r in rules if r.type == "cellIs"}
    assert '"At risk"' not in by_formula
    amber = by_formula['"At risk of delay"']
    assert amber.dxf.fill.bgColor.rgb == "FFFFFF00"
    assert by_formula['"Delayed"'].dxf.fill.bgColor.rgb == "FFFF0000"
    assert '"On track"' in by_formula
    # legend on the sheet
    legend_col = len(template()["sheets"][1]["columns"]) + 2
    col = [tp.cell(r, legend_col).value for r in range(2, 5)]
    assert col == ["On track", "At risk of delay", "Delayed"]
    # no old word anywhere in the generated sheet
    for row in tp.iter_rows(values_only=True):
        for v in row:
            if isinstance(v, str):
                assert not OLD.search(v)
