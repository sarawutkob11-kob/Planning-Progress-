import datetime as dt

import pytest

from uat_tracker.parsing.dates import (
    AMBIGUOUS, BAD, DATE_IN_TIME, OUT_OF_WINDOW, fmt, parse_date_cell, time_cell_problem,
)

from conftest import WINDOW

d = dt.date


def p(v):
    return parse_date_cell(v, WINDOW)


# Acceptance 1 -------------------------------------------------------------------------------
@pytest.mark.parametrize("cell", [
    "20/10/2026 25/9/2026 => 17/09/2026 14/9/2026",
    "20/10/2026\n25/9/2026 =>\n17/09/2026\n14/9/2026 ",        # exact form in the real file
])
def test_acceptance_1_last_typed_value_is_current_with_3_history(cell):
    r = p(cell)
    assert r.current == d(2026, 9, 14)
    assert fmt(r.current) == "14-Sep-2026"
    assert r.history == [d(2026, 10, 20), d(2026, 9, 25), d(2026, 9, 17)]
    assert r.current_problem is None and r.history_problems == []


def test_rule2_first_example():
    r = p("10/09/2026 => 19/09/2026 => 15/09/2026")
    assert r.current == d(2026, 9, 15)
    assert r.history == [d(2026, 9, 10), d(2026, 9, 19)]


@pytest.mark.parametrize("cell", ["25/09/2026 =>01/10/2026", "25/09/2026 -> 01/10/2026",
                                  "25/09/2026 --> 01/10/2026", "25/09/2026\n01/10/2026"])
def test_separators(cell):
    r = p(cell)
    assert r.current == d(2026, 10, 1) and r.history == [d(2026, 9, 25)]


# Acceptance 2 -------------------------------------------------------------------------------
def test_acceptance_2_buddhist_year_text():
    assert fmt(p("9/14/2569").current) == "14-Sep-2026"


def test_acceptance_2_buddhist_year_native_excel_date():
    assert p(dt.datetime(2569, 9, 14)).current == d(2026, 9, 14)


# Acceptance 3 -------------------------------------------------------------------------------
@pytest.mark.parametrize("cell,problem", [
    ("24/09/206", BAD),
    ("29/09/2027", OUT_OF_WINDOW),
    ("1/4/1900", OUT_OF_WINDOW),
    ("06-Now-2026", BAD),
    ("10/8/2026", AMBIGUOUS),
    ("30/102026", BAD),
    ("31/02/2026", BAD),
])
def test_acceptance_3_typos_are_not_converted(cell, problem):
    r = p(cell)
    assert r.current is None
    assert r.current_problem.problem == problem
    assert r.current_problem.raw == cell


def test_bad_history_value_keeps_valid_current():
    r = p("1/4/1900\n22/10/2026")
    assert r.current == d(2026, 10, 22)
    assert [t.problem for t in r.history_problems] == [OUT_OF_WINDOW]
    assert r.history == []


def test_excel_swapped_native_date_goes_to_review():
    r = p(dt.datetime(2569, 5, 10))         # Excel read 5/10 as May 10 -> outside the window
    assert r.current is None and r.current_problem.problem == OUT_OF_WINDOW


# Day / month rules ---------------------------------------------------------------------------
def test_zero_padded_is_day_month():
    assert p("02/11/2026").current == d(2026, 11, 2)


def test_unambiguous_month_day():
    assert p("9/30/2026").current == d(2026, 9, 30)


def test_padded_but_month_day_in_context_is_caught_by_window():
    r = p("9/30/2026 =>\n10/05/2026")       # really 5-Oct, read as 10-May -> review
    assert r.current is None and r.current_problem.problem == OUT_OF_WINDOW


def test_named_month_formats():
    assert p("10-Oct-2026").current == d(2026, 10, 10)
    r = p("10/6/2026 -> \n20 Oct 2026 -> \n27 Oct 2026  5/11/2026")
    assert r.current is None and r.current_problem.problem == AMBIGUOUS
    assert r.history == [d(2026, 10, 20), d(2026, 10, 27)]


# Rule 3: text -> note -----------------------------------------------------------------------
def test_parenthesised_text_goes_to_note():
    r = p("02/11/2026 (BD 30/09/26)\n")
    assert r.current == d(2026, 11, 2) and r.note == "(BD 30/09/26)" and r.history == []
    r = p("24/09/2026 (actual start 25/09/26)")
    assert r.current == d(2026, 9, 24) and r.note == "(actual start 25/09/26)"


def test_plain_text_only_is_note():
    r = p(".")
    assert r.current is None and r.current_problem is None and r.note == "."


def test_blank():
    r = p(None)
    assert r.current is None and r.tokens == [] and r.note is None


# Rule 5 --------------------------------------------------------------------------------------
def test_format():
    assert fmt(d(2026, 9, 4)) == "04-Sep-2026"
    assert fmt("2026-11-13") == "13-Nov-2026"


# Time columns --------------------------------------------------------------------------------
def test_time_columns():
    assert time_cell_problem(dt.time(15, 0)) == ("15:00", None)
    assert time_cell_problem("16:00") == ("16:00", None)
    assert time_cell_problem("06-Now-2026\n13-Nov-2026")[1] == DATE_IN_TIME
    assert time_cell_problem(dt.datetime(2026, 9, 25))[1] == DATE_IN_TIME
