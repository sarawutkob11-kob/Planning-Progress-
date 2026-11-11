import pytest

from uat_tracker.parsing.headers import FieldSpec, HeaderError, map_headers
from uat_tracker.parsing.status import UNKNOWN_STATUS, normalize_defect_status, normalize_task_status
from uat_tracker.parsing.text import parse_number


@pytest.mark.parametrize("raw,expected", [
    ("completed", "Complete"), ("Completed", "Complete"), ("complete", "Complete"),
    ("comlete", "Complete"), ("Complete", "Complete"), ("Inprogress", "In progress"),
    ("In progress", "In progress"), ("Not Start", "Not Start"), ("Blocked", "Blocked"),
])
def test_task_status(raw, expected):
    assert normalize_task_status(raw) == (expected, None)


def test_status_blank_and_unknown():
    assert normalize_task_status(None) == (None, None)
    assert normalize_task_status("\xa0") == (None, None)
    assert normalize_task_status("Waiting") == (None, UNKNOWN_STATUS)


def test_defect_status():
    assert normalize_defect_status("Close") == ("Close", None)
    assert normalize_defect_status("assigned") == ("Assigned", None)
    assert normalize_defect_status("Done") == (None, UNKNOWN_STATUS)


SPECS = [FieldSpec("status", ("Status",)), FieldSpec("name", ("Task name",)),
         FieldSpec("email", ("Report By Email",), required=False)]


def test_headers_by_name_not_position():
    h = map_headers(["Task name\n(Mandatory)", "X", "Status"], SPECS, "S")
    assert h.columns == {"name": 0, "status": 2}
    h = map_headers(["Status", "Task name"], SPECS, "S")
    assert h.columns == {"status": 0, "name": 1}
    assert h.missing_optional == ["Report By Email"]


def test_headers_leftmost_wins():
    h = map_headers(["Task name", "Status", "Status\n(Optional)", None, "Status"], SPECS, "S")
    assert h.columns["status"] == 1


def test_headers_side_table_after_blank_is_ignored():
    with pytest.raises(HeaderError):
        map_headers(["Task name", None, "Status"], SPECS, "S")


def test_blank_first_header_alias():
    specs = [FieldSpec("no", ("No.",)), FieldSpec("status", ("Status",))]
    assert map_headers([None, "Status"], specs, "D", blank_first_as="No.").columns == {"no": 0, "status": 1}


def test_numbers():
    assert parse_number(4) == (4, None)
    assert parse_number(1.0) == (1, None)
    assert parse_number(None) == (None, None)
    assert parse_number("12") == (12, None)
    assert parse_number("abc")[1] == "bad_number"
