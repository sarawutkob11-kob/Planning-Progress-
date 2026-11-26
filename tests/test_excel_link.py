"""Project Excel file from a SharePoint link (26-Nov-2026).

T32: resolve_excel_link + validate_link_file. T33: import_project_link. T34: relink_link. No registry, no Excel: libraries are injected (SyncedLibrary),
tmp_path stands in for the OneDrive MountPoint, workbooks are made with build_workbook.
"""

from __future__ import annotations

import dataclasses
import hashlib
from pathlib import Path

import openpyxl
import pytest

from uat_tracker.config import Config
from uat_tracker.excel import reader
from uat_tracker.services import projects, sharepoint_link
from uat_tracker.services.projects import AttachmentDirError, ExcelLinkError
from uat_tracker.services.sharepoint_link import SyncedLibrary

from uat_tracker.services.importer import run_import

from conftest import TP_HEADER, WINDOW, build_workbook

HOST = "https://t.sharepoint.com"
LIB_URL = f"{HOST}/sites/Team/Shared Documents"
FILE_NAME = "Planning Progress Demo.xlsx"
LINK = f"{HOST}/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx"


@pytest.fixture(autouse=True)
def _no_registry(monkeypatch):
    def boom():
        raise AssertionError("registry must not be read in tests")
    monkeypatch.setattr(sharepoint_link, "synced_libraries", boom)


@pytest.fixture
def base(tmp_path) -> Config:
    root = tmp_path / "root"
    return Config(excel_path=tmp_path / "unused.xlsx", data_dir=root, window_start=WINDOW.start,
                  window_end=WINDOW.end, profile_file=root / "profile.json")


@pytest.fixture
def mount(tmp_path: Path) -> Path:
    m = tmp_path / "OneDrive - Org" / "Team - Documents"
    m.mkdir(parents=True)
    return m


@pytest.fixture
def libs(mount: Path) -> list[SyncedLibrary]:
    return [SyncedLibrary(LIB_URL, str(mount))]


@pytest.fixture
def target(mount: Path) -> Path:
    return mount / "UAT" / FILE_NAME


def _make(path: Path, **kw) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    return build_workbook(path, **kw)


def _tree(root: Path) -> set[Path]:
    return set(root.rglob("*"))


# ---------------------------------------------------------------- resolve_excel_link

def test_existing_file_returns_clean_url_and_path(libs, target):
    _make(target)
    url, path = projects.resolve_excel_link(f'  "{LINK}?web=1"  ', libs)
    assert url == LINK                                    # plain file URL, no query string
    assert path == target


def test_upper_case_extension_ok(libs, mount):
    f = _make(mount / "UAT" / "Plan.XLSX")
    url, path = projects.resolve_excel_link(f"{HOST}/sites/Team/Shared%20Documents/UAT/Plan.XLSX", libs)
    assert path == f and url.endswith("Plan.XLSX")


@pytest.mark.parametrize("with_parent", [True, False], ids=["UAT-folder-exists", "no-UAT-folder"])
def test_ac10_missing_file_not_found_creates_nothing(libs, mount, target, with_parent):
    if with_parent:
        target.parent.mkdir()
    before = _tree(mount)
    with pytest.raises(ExcelLinkError) as ei:
        projects.resolve_excel_link(LINK, libs)
    exc = ei.value
    assert exc.code == "link_not_found"
    assert exc.path == str(target)
    assert isinstance(exc, projects.ProjectError) and not isinstance(exc, AttachmentDirError)
    assert _tree(mount) == before


def test_folder_named_xlsx_is_not_xlsx(libs, target):
    target.mkdir(parents=True)
    with pytest.raises(ExcelLinkError) as ei:
        projects.resolve_excel_link(LINK, libs)
    assert ei.value.code == "link_not_xlsx"
    assert ei.value.path == str(target)


def test_resolver_errors_pass_through(libs):
    with pytest.raises(ExcelLinkError) as ei:
        projects.resolve_excel_link(f"{HOST}/sites/Team/Shared%20Documents/UAT/Plan.xlsm", libs)
    assert ei.value.code == "link_not_xlsx" and ei.value.path.startswith(HOST)
    with pytest.raises(ExcelLinkError) as ei:
        projects.resolve_excel_link(LINK, [])
    assert ei.value.code == "link_not_synced" and ei.value.path == LINK


def test_resolve_excel_link_reads_no_content(libs, target, monkeypatch):
    _make(target)

    def boom(*_a, **_k):
        raise AssertionError("file content must not be read")
    monkeypatch.setattr(reader, "copy_shared", boom)
    assert projects.resolve_excel_link(LINK, libs) == (LINK, target)


# ---------------------------------------------------------------- validate_link_file

def test_validate_link_file_ok(base, target):
    _make(target)
    assert projects.validate_link_file(base, target) == []


def test_validate_link_file_same_problems_as_validate_workbook(base, target):
    wb = openpyxl.Workbook()
    wb.active.title = "UAT Plan"
    wb.create_sheet("Test Progress")
    target.parent.mkdir(parents=True)
    wb.save(target)
    probs = projects.validate_link_file(base, target)
    assert probs and probs == projects.validate_workbook(target, base.header_overrides)
    assert "Sheet 'Defect List' is missing" in probs


def test_validate_link_file_unreadable(base, target, monkeypatch):
    _make(target)

    def locked(*_a, **_k):
        raise PermissionError("cloud file not available")
    monkeypatch.setattr(reader, "copy_shared", locked)
    with pytest.raises(ExcelLinkError) as ei:
        projects.validate_link_file(base, target)
    assert ei.value.code == "link_unreadable"
    assert ei.value.path == str(target)
    assert isinstance(ei.value.__cause__, OSError)


# ---------------------------------------------------------------- helpers used by T33 / T34

def test_helpers_keep_oserror_without_link(base, target, monkeypatch):
    _make(target)

    def locked(*_a, **_k):
        raise OSError("busy")
    monkeypatch.setattr(reader, "copy_shared", locked)
    with pytest.raises(OSError) as ei:
        projects._validate(base, target, None)
    assert not isinstance(ei.value, ExcelLinkError)
    with pytest.raises(OSError) as ei:
        projects._ids(target, None)
    assert not isinstance(ei.value, ExcelLinkError)
    with pytest.raises(ExcelLinkError) as ei:
        projects._ids(target, LINK)
    assert ei.value.code == "link_unreadable" and ei.value.path == str(target)


def test_ids_reads_app_ids(target):
    _make(target)
    assert projects._ids(target, LINK) == projects.app_ids_in(target)


def test_source():
    assert projects._source(None) == ""
    assert projects._source("") == ""
    assert projects._source(LINK) == f" (SharePoint: {LINK})"


# ---------------------------------------------------------------- QA additions (T32)

@pytest.mark.parametrize("link", [
    LINK,
    f"{LINK}?web=1",
    f"{HOST}/sites/Team/Shared Documents/UAT/Planning Progress Demo.xlsx",
    f'  "{LINK}"  ',
    f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FTeam%2FShared%20Documents"
    f"%2FUAT%2FPlanning%20Progress%20Demo.xlsx&parent=%2Fsites%2FTeam%2FShared%20Documents%2FUAT",
    f"{HOST}/:x:/r/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx?d=w123&csf=1&web=1&e=abc",
], ids=["direct", "web=1", "unencoded", "quoted", "address-bar", "share-r"])
def test_qa_link_forms_resolve_to_file(libs, target, link):
    _make(target)
    url, path = projects.resolve_excel_link(link, libs)
    assert path == target
    # returned url = canonical file URL (spec addition 26-Nov-2026): no query string, ?id= / /:x:/r/ -> file path
    assert url == LINK
    assert url == sharepoint_link.canonical_file_url(link)


@pytest.mark.parametrize("link", [
    f"{HOST}/sites/Team/Shared%20Documents/UAT/Report.xls",
    f"{HOST}/sites/Team/Shared%20Documents/UAT/Report.xlsm",
    f"{HOST}/sites/Team/Shared%20Documents/UAT/Data.csv",
    f"{HOST}/sites/Team/Shared%20Documents/UAT",
    f"{HOST}/:f:/r/sites/Team/Shared%20Documents/UAT",
], ids=["xls", "xlsm", "csv", "folder", "f-r"])
def test_qa_ac6_not_xlsx_before_sync_check(link, tmp_path):
    before = _tree(tmp_path)
    with pytest.raises(ExcelLinkError) as ei:
        projects.resolve_excel_link(link, [])
    assert ei.value.code == "link_not_xlsx"
    assert ei.value.path == link
    assert _tree(tmp_path) == before


def test_qa_ac7_upper_case_via_address_bar(libs, mount):
    f = _make(mount / "UAT" / "PLAN.XLSX")
    link = (f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx"
            f"?id=%2Fsites%2FTeam%2FShared%20Documents%2FUAT%2FPLAN.XLSX")
    assert projects.resolve_excel_link(link, libs)[1] == f


@pytest.mark.parametrize("with_parent", [True, False], ids=["UAT-folder-exists", "no-UAT-folder"])
def test_qa_ac10_nothing_created_anywhere(libs, mount, target, tmp_path, with_parent):
    # other files in the library must not satisfy the link
    _make(mount / "Other" / FILE_NAME)
    if with_parent:
        target.parent.mkdir()
    before = _tree(tmp_path)
    with pytest.raises(ExcelLinkError) as ei:
        projects.resolve_excel_link(f"{LINK}?web=1", libs)
    assert ei.value.code == "link_not_found" and ei.value.path == str(target)
    assert _tree(tmp_path) == before
    assert target.parent.exists() is with_parent


def test_qa_ac10_message_mentions_local_path(libs, target):
    with pytest.raises(ExcelLinkError) as ei:
        projects.resolve_excel_link(LINK, libs)
    assert str(target) in str(ei.value)


def test_qa_dotdot_is_invalid_and_creates_nothing(libs, tmp_path):
    (tmp_path / "x.xlsx").write_bytes(b"")          # a file that exists outside the MountPoint
    before = _tree(tmp_path)
    with pytest.raises(ExcelLinkError) as ei:
        projects.resolve_excel_link(f"{HOST}/sites/Team/Shared%20Documents/UAT/../../x.xlsx", libs)
    assert ei.value.code == "link_invalid"
    assert _tree(tmp_path) == before


@pytest.mark.parametrize("link", ["", "   ", '""', "abc", r"C:\x.xlsx", "ftp://t.sharepoint.com/a.xlsx"])
def test_qa_not_a_url_is_invalid(link, libs):
    with pytest.raises(ExcelLinkError) as ei:
        projects.resolve_excel_link(link, libs)
    assert ei.value.code == "link_invalid"


def test_qa_default_libraries_come_from_synced_libraries(monkeypatch, mount, target):
    _make(target)
    calls = []

    def fake():
        calls.append(1)
        return [SyncedLibrary(LIB_URL, str(mount))]
    monkeypatch.setattr(sharepoint_link, "synced_libraries", fake)
    assert projects.resolve_excel_link(LINK) == (LINK, target)
    assert calls


# ---- AC-18 helper part: OSError family -> link_unreadable only with a link

@pytest.mark.parametrize("err", [OSError("offline"), PermissionError("denied"), FileNotFoundError("gone")],
                         ids=["OSError", "PermissionError", "FileNotFoundError"])
def test_qa_validate_link_file_oserror_family(base, target, monkeypatch, err):
    _make(target)

    def fail(*_a, **_k):
        raise err
    monkeypatch.setattr(reader, "copy_shared", fail)
    with pytest.raises(ExcelLinkError) as ei:
        projects.validate_link_file(base, target)
    assert ei.value.code == "link_unreadable" and ei.value.path == str(target)
    assert ei.value.__cause__ is err
    with pytest.raises(ExcelLinkError) as ei:
        projects._validate(base, target, LINK)
    assert ei.value.code == "link_unreadable" and ei.value.path == str(target)
    with pytest.raises(ExcelLinkError) as ei:
        projects._ids(target, LINK)
    assert ei.value.code == "link_unreadable" and ei.value.path == str(target)


@pytest.mark.parametrize("err_type", [OSError, PermissionError])
def test_qa_no_link_raises_plain_oserror(base, target, monkeypatch, err_type):
    _make(target)
    err = err_type("busy")

    def fail(*_a, **_k):
        raise err
    monkeypatch.setattr(reader, "copy_shared", fail)
    for call in (lambda: projects._validate(base, target, None), lambda: projects._ids(target, None)):
        with pytest.raises(OSError) as ei:
            call()
        assert ei.value is err
        assert not isinstance(ei.value, projects.ProjectError)
    # the public validate_workbook / app_ids_in are unchanged too
    with pytest.raises(err_type):
        projects.validate_workbook(target, base.header_overrides)
    with pytest.raises(err_type):
        projects.app_ids_in(target)


def test_qa_non_oserror_not_translated(base, target, monkeypatch):
    _make(target)

    def fail(*_a, **_k):
        raise ValueError("bug")
    monkeypatch.setattr(reader, "copy_shared", fail)
    with pytest.raises(ValueError):
        projects.validate_link_file(base, target)
    with pytest.raises(ValueError):
        projects._ids(target, LINK)


def test_qa_validate_link_file_accepts_str_and_does_not_modify(base, target):
    _make(target)
    data = target.read_bytes()
    assert projects.validate_link_file(base, str(target)) == []
    assert target.read_bytes() == data


def test_qa_validate_link_file_not_a_workbook(base, target):
    target.parent.mkdir(parents=True)
    target.write_bytes(b"not a zip")
    probs = projects.validate_link_file(base, target)
    assert probs and probs[0].startswith("Not a readable Excel file")


# ---------------------------------------------------------------- canonical_file_url (spec addition 26-Nov-2026)

ADDRESS_BAR = (f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FTeam%2FShared%20Documents"
               f"%2FUAT%2FPlanning%20Progress%20Demo.xlsx&parent=%2Fsites%2FTeam%2FShared%20Documents%2FUAT")


@pytest.mark.parametrize("link", [
    LINK,
    f"{LINK}?web=1",
    f'  "{LINK}?d=w1&csf=1&web=1&e=abc"  ',
    f"{HOST}/sites/Team/Shared Documents/UAT/Planning Progress Demo.xlsx",
    ADDRESS_BAR,
    f"{HOST}/:x:/r/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx?d=w123&csf=1&web=1&e=abc",
    "HTTPS://T.SharePoint.com/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx#frag",
], ids=["direct", "web=1", "quoted-query", "unencoded", "address-bar", "share-r", "upper-host-fragment"])
def test_canonical_file_url_has_no_query(link):
    url = sharepoint_link.canonical_file_url(link)
    assert url == LINK
    assert "?" not in url and "#" not in url


def test_canonical_file_url_encodes_special_characters():
    link = f"{HOST}/sites/Team/Shared%20Documents/A%23B%3F/Plan%20%26%20%25.xlsx"
    assert sharepoint_link.canonical_file_url(link) == \
        f"{HOST}/sites/Team/Shared%20Documents/A%23B%3F/Plan%20%26%20%25.xlsx"


def test_resolve_excel_link_errors_keep_pasted_link(libs):
    link = f"{HOST}/:x:/s/Team/EaBcD123?e=xyz"
    with pytest.raises(ExcelLinkError) as ei:
        projects.resolve_excel_link(f'  "{link}"  ', libs)
    assert ei.value.code == "link_sharing" and ei.value.path == link
    with pytest.raises(ExcelLinkError) as ei:
        projects.resolve_excel_link(f"{LINK}?web=1", [])
    assert ei.value.code == "link_not_synced" and ei.value.path == f"{LINK}?web=1"


# ---------------------------------------------------------------- T33: import_project_link (a)

PERSON = projects.Person("pmo@example.test", "Pim", "Pmo")
OTHER = projects.Person("other@example.test", "Oz", "Other")


def _sha(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


@pytest.fixture
def appdb(base):
    db = projects.open_app_db(base)
    yield db
    db.close()


@pytest.fixture
def synced(monkeypatch, libs):
    """synced_libraries() returns the fake library (import_project_link reads it, no registry)."""
    monkeypatch.setattr(sharepoint_link, "synced_libraries", lambda: list(libs))
    return libs


def _add_member(base: Config, project: dict, email: str, role: str) -> None:
    def go(_cfg, repo):
        with repo.transaction():
            repo.insert_user(email=email, first_name=None, last_name=None, role=role,
                             created_at="2026-11-26T00:00:00", created_by="t")
    projects._with_repo(base, project, go)


def _state(appdb) -> tuple[int, int]:
    return len(appdb.projects(include_archived=True)), len(appdb.events(limit=10_000))


def _imported_events(appdb, pid: str) -> list[dict]:
    return [e for e in appdb.events(pid, limit=10_000) if e["action"] == "imported"]


def test_t33_ac13_first_project_from_link(base, appdb, synced, target):
    _make(target)
    before = _sha(target)
    p = projects.import_project_link(base, appdb, PERSON, link=LINK)
    assert p["name"] == "Demo"
    assert p["excel_path"] == str(target)
    assert projects.role_in(base, p, PERSON.email) == "Admin"
    assert _sha(target) == before


@pytest.mark.parametrize("role", ["Admin", "PMO"])
def test_t33_ac13_manager_of_other_project(base, appdb, synced, target, tmp_path, role):
    first = projects.import_project(base, appdb, OTHER, path=_make(tmp_path / "x" / "Planning Progress Alpha.xlsx"))
    _add_member(base, first, PERSON.email, role)
    _make(target)
    before = _sha(target)
    p = projects.import_project_link(base, appdb, PERSON, link=f"{LINK}?web=1")
    assert p["name"] == "Demo" and p["excel_path"] == str(target)
    assert projects.role_in(base, p, PERSON.email) == "Admin"
    assert _sha(target) == before
    assert len(appdb.projects()) == 2


@pytest.mark.parametrize("link", [f'  "{LINK}?web=1"  ', ADDRESS_BAR,
                                  f"{HOST}/:x:/r/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx"
                                  f"?d=w1&e=abc"], ids=["quoted-web", "address-bar", "share-r"])
def test_t33_ac14_event_detail_has_path_and_plain_link(base, appdb, synced, target, link):
    _make(target)
    p = projects.import_project_link(base, appdb, PERSON, link=link)
    ev = _imported_events(appdb, p["id"])
    assert len(ev) == 1
    assert ev[0]["detail"] == f"existing file {target} (SharePoint: {LINK})"    # no query string
    assert ev[0]["editor"] == PERSON.editor


def test_t33_ac15_qa_everywhere_not_allowed_before_link(base, appdb, monkeypatch, tmp_path):
    first = projects.import_project(base, appdb, OTHER, path=_make(tmp_path / "x" / "Planning Progress Alpha.xlsx"))
    _add_member(base, first, PERSON.email, "QA")
    calls = []

    def no_libs():
        calls.append(1)
        return []
    monkeypatch.setattr(sharepoint_link, "synced_libraries", no_libs)
    before = _state(appdb)
    with pytest.raises(projects.ProjectError) as ei:
        projects.import_project_link(base, appdb, PERSON, link=LINK)
    assert not isinstance(ei.value, ExcelLinkError)
    assert str(ei.value).startswith("Not allowed")
    assert not calls                                         # permission checked before the link
    assert _state(appdb) == before


@pytest.mark.parametrize("archived", [False, True], ids=["normal", "archived"])
def test_t33_ac16_duplicate_file_name(base, appdb, synced, target, tmp_path, archived):
    q = projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "q" / FILE_NAME))
    if archived:
        projects.set_archived(base, appdb, PERSON, q["id"], True)
    _make(target)
    before = _state(appdb)
    with pytest.raises(projects.ProjectError, match="already uses the file name") as ei:
        projects.import_project_link(base, appdb, PERSON, link=LINK)
    assert not isinstance(ei.value, ExcelLinkError)
    assert _state(appdb) == before


def test_t33_ac17_missing_header(base, appdb, synced, target):
    _make(target, tp_header=[h if h != "Total TC" else "Totals" for h in TP_HEADER])
    with pytest.raises(projects.ProjectError, match="can't be imported") as ei:
        projects.import_project_link(base, appdb, PERSON, link=LINK)
    assert "Total TC" in str(ei.value) and not isinstance(ei.value, ExcelLinkError)
    assert _state(appdb) == (0, 0)


def test_t33_ac18_unreadable(base, appdb, synced, target, monkeypatch):
    _make(target)

    def locked(*_a, **_k):
        raise PermissionError("cloud file not available")
    monkeypatch.setattr(reader, "copy_shared", locked)
    with pytest.raises(ExcelLinkError) as ei:
        projects.import_project_link(base, appdb, PERSON, link=LINK)
    assert ei.value.code == "link_unreadable" and ei.value.path == str(target)
    assert _state(appdb) == (0, 0)


AC19_CASES = [
    (f"{HOST}/:x:/s/Team/EaBcD123?e=xyz", "link_sharing"),
    (f"{HOST}/:x:/g/personal/u/EaBcD123", "link_sharing"),
    (f"{HOST}/sites/Team/_layouts/15/Doc.aspx?sourcedoc=%7Babc%7D&file=Plan.xlsx", "link_sharing"),
    (f"{HOST}/sites/Team/Shared%20Documents/UAT/Report.xls", "link_not_xlsx"),
    (f"{HOST}/sites/Team/Shared%20Documents/UAT/Report.xlsm", "link_not_xlsx"),
    (f"{HOST}/:f:/r/sites/Team/Shared%20Documents/UAT", "link_not_xlsx"),
    ("abc", "link_invalid"),
    ("", "link_invalid"),
    (f"{HOST}/sites/Team/Shared%20Documents/UAT/../../x.xlsx", "link_invalid"),
    ("https://other.sharepoint.com/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx",
     "link_not_synced"),
    (f"{HOST}/sites/Other/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx", "link_not_synced"),
    (LINK, "link_not_found"),
]


@pytest.mark.parametrize("link,code", AC19_CASES, ids=[f"{c}-{i}" for i, (_l, c) in enumerate(AC19_CASES)])
def test_t33_ac19_link_errors_save_nothing(base, appdb, synced, tmp_path, link, code):
    projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "x" / "Planning Progress Alpha.xlsx"))
    before = _state(appdb)
    with pytest.raises(ExcelLinkError) as ei:
        projects.import_project_link(base, appdb, PERSON, link=link)
    assert ei.value.code == code
    assert _state(appdb) == before


def test_t33_ac20_reads_same_as_file_import(base, appdb, synced, target, tmp_path):
    _make(target)
    p = projects.import_project_link(base, appdb, PERSON, link=LINK)
    assert projects.project_config(base, p).excel_path == target
    cfg, repo = projects.open_project(base, p)
    try:
        got = {s: r.rows for s, r in run_import(cfg, repo, editor="t").sheets.items()}
    finally:
        repo.close()

    base2 = dataclasses.replace(base, data_dir=tmp_path / "root2", profile_file=tmp_path / "root2" / "p.json")
    db2 = projects.open_app_db(base2)
    try:
        p2 = projects.import_project(base2, db2, PERSON, path=target)
        cfg2, repo2 = projects.open_project(base2, p2)
        try:
            want = {s: r.rows for s, r in run_import(cfg2, repo2, editor="t").sheets.items()}
        finally:
            repo2.close()
    finally:
        db2.close()
    assert got == want and any(want.values())


def test_t33_ac27_import_project_unchanged(base, appdb, tmp_path, monkeypatch):
    f = _make(tmp_path / "g" / "Planning Progress Gamma.xlsx")
    p = projects.import_project(base, appdb, PERSON, path=str(f))           # str still accepted
    ev = _imported_events(appdb, p["id"])
    assert len(ev) == 1 and ev[0]["detail"] == f"existing file {f}"
    assert "SharePoint" not in ev[0]["detail"]

    g = _make(tmp_path / "h" / "Planning Progress Delta.xlsx")

    def busy(*_a, **_k):
        raise OSError("busy")
    monkeypatch.setattr(reader, "copy_shared", busy)
    with pytest.raises(OSError) as ei:                                       # plain OSError as before
        projects.import_project(base, appdb, PERSON, path=g)
    assert not isinstance(ei.value, projects.ProjectError)


def test_t33_import_project_link_is_keyword_only(base, appdb):
    with pytest.raises(TypeError):
        projects.import_project_link(base, appdb, PERSON, LINK)              # type: ignore[misc]


# ---------------------------------------------------------------- QA T33 + spec addition (no query string in detail)

from urllib.parse import quote as _q, urlsplit as _split  # noqa: E402

SPECIAL_DIR = "ทดสอบ #1 & 100% A?B"            # '?' only in URLs (not a valid Windows file name character)
SPECIAL_DISK_DIR = "ทดสอบ #1 & 100% ok"
SPECIAL_FILE = "Planning Progress งาน #2 & 50% x.xlsx"


def _enc(*segs: str) -> str:
    return "/".join(_q(s, safe="") for s in segs)


@pytest.mark.parametrize("seg", ["a#b", "a?b", "100%", "%20", "a&b=c", "sp ace", "ทดสอบ", "a+b", "a;b", "x.xlsx"])
def test_qa_canonical_round_trips_special_segments(libs, mount, seg):
    link = f"{HOST}/sites/Team/Shared%20Documents/{_enc(seg)}/{_enc(SPECIAL_FILE)}"
    url = sharepoint_link.canonical_file_url(link)
    parts = _split(url)
    assert "?" not in url and "#" not in url
    assert parts.query == "" and parts.fragment == ""
    want = mount / seg / SPECIAL_FILE
    assert sharepoint_link.resolve_file_link(link, libs) == want
    assert sharepoint_link.resolve_file_link(url, libs) == want                 # canonical URL -> same file
    assert sharepoint_link.canonical_file_url(url) == url                        # idempotent


@pytest.mark.parametrize("form", ["direct", "address-bar", "share-r", "fragment", "upper-host"])
def test_qa_canonical_same_for_all_forms_of_special_file(libs, mount, form):
    segs = ["sites", "Team", "Shared Documents", SPECIAL_DIR, SPECIAL_FILE]
    direct = f"{HOST}/{_enc(*segs)}"
    if form == "direct":
        link = f"{direct}?web=1"
    elif form == "address-bar":
        link = (f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?id={_q('/' + '/'.join(segs), safe='')}"
                f"&parent={_q('/' + '/'.join(segs[:-1]), safe='')}&viewid=abc")
    elif form == "share-r":
        link = f"{HOST}/:x:/r/{_enc(*segs)}?d=w1&csf=1&web=1&e=Ab%3Fc"
    elif form == "fragment":
        link = f"{direct}#page=2?x=1"
    else:
        link = "HTTPS://T.SHAREPOINT.COM/" + _enc(*segs)
    url = sharepoint_link.canonical_file_url(link)
    assert url == direct
    assert "?" not in url and "#" not in url
    assert sharepoint_link.resolve_file_link(url, libs) == mount / SPECIAL_DIR / SPECIAL_FILE


def test_qa_canonical_unencoded_paste_with_spaces_and_thai(libs, mount):
    # unencoded '#' / '?' would start a fragment / query, so only spaces, '&' and Thai here
    link = f"{HOST}/sites/Team/Shared Documents/ทดสอบ & ok/Planning Progress งาน & x.xlsx"
    url = sharepoint_link.canonical_file_url(link)
    assert " " not in url and "?" not in url and "#" not in url
    assert url.isascii()
    assert sharepoint_link.resolve_file_link(url, libs) == sharepoint_link.resolve_file_link(link, libs)         == mount / "ทดสอบ & ok" / "Planning Progress งาน & x.xlsx"


@pytest.mark.parametrize("form", ["direct-query", "address-bar", "share-r", "fragment"])
def test_qa_t33_event_detail_never_has_query_special_names(base, appdb, synced, mount, form):
    segs = ["sites", "Team", "Shared Documents", SPECIAL_DISK_DIR, SPECIAL_FILE]
    target = _make(mount / SPECIAL_DISK_DIR / SPECIAL_FILE)
    direct = f"{HOST}/{_enc(*segs)}"
    link = {
        "direct-query": f'  "{direct}?web=1&e=x%3Fy"  ',
        "address-bar": (f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?id="
                        f"{_q('/' + '/'.join(segs), safe='')}&parent={_q('/' + '/'.join(segs[:-1]), safe='')}"),
        "share-r": f"{HOST}/:x:/r/{_enc(*segs)}?d=w1&csf=1&web=1&e=abc",
        "fragment": f"{direct}#anchor",
    }[form]
    p = projects.import_project_link(base, appdb, PERSON, link=link)
    assert p["excel_path"] == str(target)
    assert p["name"] == "งาน #2 & 50% x"
    ev = _imported_events(appdb, p["id"])
    assert len(ev) == 1
    detail = ev[0]["detail"]
    assert detail == f"existing file {target} (SharePoint: {direct})"
    url = detail.split("(SharePoint: ", 1)[1][:-1]
    assert "?" not in url and "#" not in url and " " not in url
    assert sharepoint_link.resolve_file_link(url, synced) == target                # URL in the detail finds the file


def test_qa_t33_detail_plain_link_for_unencoded_with_fragment(base, appdb, synced, target):
    _make(target)
    unencoded = f"{HOST}/sites/Team/Shared Documents/UAT/Planning Progress Demo.xlsx#x"
    p = projects.import_project_link(base, appdb, PERSON, link=unencoded)
    assert _imported_events(appdb, p["id"])[0]["detail"] == f"existing file {target} (SharePoint: {LINK})"


@pytest.mark.parametrize("link,code", [
    (f"{HOST}/:x:/s/Team/EaBc?e=x&web=1", "link_sharing"),
    (f"{HOST}/sites/Team/_layouts/15/Doc.aspx?sourcedoc=%7Bg%7D&file=a.xlsx#f", "link_sharing"),
    (f"{HOST}/sites/Team/Shared%20Documents/UAT/Report.xls?web=1", "link_not_xlsx"),
    (f"{HOST}/sites/Team/Shared%20Documents/UAT/../x.xlsx?web=1", "link_invalid"),
    (f"https://other.sharepoint.com/sites/Team/Shared%20Documents/UAT/{_q(FILE_NAME)}?web=1&e=1", "link_not_synced"),
], ids=["sharing", "doc-aspx-fragment", "xls", "dotdot", "other-host"])
def test_qa_t33_error_path_is_cleaned_pasted_link_with_query(base, appdb, synced, link, code):
    before = _state(appdb)
    with pytest.raises(ExcelLinkError) as ei:
        projects.import_project_link(base, appdb, PERSON, link=f'  "{link}"  ')
    assert ei.value.code == code
    assert ei.value.path == link                                                 # query (and fragment) kept
    assert _state(appdb) == before


def test_qa_t33_not_found_path_is_local_path(base, appdb, synced, target):
    with pytest.raises(ExcelLinkError) as ei:
        projects.import_project_link(base, appdb, PERSON, link=f"{LINK}?web=1")
    assert ei.value.code == "link_not_found" and ei.value.path == str(target)


@pytest.mark.parametrize("link", ["", "abc", f"{HOST}/:x:/s/Team/EaBc", LINK])
def test_qa_t33_permission_before_any_link_work(base, appdb, monkeypatch, tmp_path, link):
    first = projects.import_project(base, appdb, OTHER, path=_make(tmp_path / "x" / "Planning Progress Alpha.xlsx"))
    _add_member(base, first, PERSON.email, "User BU/PM")
    calls = []

    def resolve(*a, **_k):
        calls.append(a)
        raise AssertionError("resolved before permission check")
    monkeypatch.setattr(projects, "resolve_excel_link", resolve)
    before = _state(appdb)
    with pytest.raises(projects.ProjectError) as ei:
        projects.import_project_link(base, appdb, PERSON, link=link)
    assert type(ei.value) is projects.ProjectError
    assert str(ei.value) == "Not allowed: only an Admin or PMO (of any project) can do this"
    assert not calls and _state(appdb) == before


def test_qa_t33_import_link_uses_import_file_with_url(base, appdb, synced, target, monkeypatch):
    _make(target)
    seen = []
    real = projects._import_file

    def spy(b, a, person, path, link_url):
        seen.append((path, link_url))
        return real(b, a, person, path, link_url)
    monkeypatch.setattr(projects, "_import_file", spy)
    projects.import_project_link(base, appdb, PERSON, link=f"{LINK}?web=1")
    assert seen == [(target, LINK)]
    other = _make(target.parent / "Planning Progress B.xlsx")
    projects.import_project(base, appdb, PERSON, path=str(other))
    assert seen[-1] == (other, None) and isinstance(seen[-1][0], Path)


# --- import_project (file) behaves exactly as before: messages and check order

def test_qa_t33_import_project_not_allowed_message(base, appdb, tmp_path):
    first = projects.import_project(base, appdb, OTHER, path=_make(tmp_path / "x" / "Planning Progress Alpha.xlsx"))
    _add_member(base, first, PERSON.email, "QA")
    with pytest.raises(projects.ProjectError) as ei:
        projects.import_project(base, appdb, PERSON, path=tmp_path / "nope.xlsx")
    assert type(ei.value) is projects.ProjectError
    assert str(ei.value) == "Not allowed: only an Admin or PMO (of any project) can do this"


def test_qa_t33_import_project_missing_file_message(base, appdb, tmp_path):
    f = tmp_path / "Planning Progress Gone.xlsx"
    with pytest.raises(projects.ProjectError) as ei:
        projects.import_project(base, appdb, PERSON, path=f)
    assert type(ei.value) is projects.ProjectError
    assert str(ei.value) == f"This file can't be imported:\n- File not found: {f}"
    assert _state(appdb) == (0, 0)


def test_qa_t33_import_project_not_xlsx_message(base, appdb, tmp_path):
    f = tmp_path / "Planning Progress X.xlsm"
    f.write_bytes(b"x")
    with pytest.raises(projects.ProjectError) as ei:
        projects.import_project(base, appdb, PERSON, path=f)
    assert str(ei.value) == "This file can't be imported:\n- Only .xlsx files can be used"


def test_qa_t33_import_project_validation_before_duplicate(base, appdb, tmp_path):
    projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "a" / "Planning Progress Dup.xlsx"))
    bad = _make(tmp_path / "b" / "Planning Progress Dup.xlsx",
                tp_header=[h if h != "Total TC" else "T" for h in TP_HEADER])
    with pytest.raises(projects.ProjectError) as ei:
        projects.import_project(base, appdb, PERSON, path=bad)
    assert str(ei.value).startswith("This file can't be imported:\n- ")


def test_qa_t33_import_project_duplicate_message(base, appdb, tmp_path):
    q = projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "a" / "Planning Progress Dup.xlsx"))
    with pytest.raises(projects.ProjectError) as ei:
        projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "b" / "planning progress DUP.xlsx"))
    assert str(ei.value) == \
        f"A project already uses the file name 'planning progress DUP.xlsx' (project '{q['name']}')"


def test_qa_t33_import_project_no_name_message(base, appdb, tmp_path):
    with pytest.raises(projects.ProjectError) as ei:
        projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "a" / " .xlsx"))
    assert str(ei.value) == "The file name gives no project name"
    assert _state(appdb) == (0, 0)


def test_qa_t33_link_validation_before_duplicate_and_no_name(base, appdb, synced, mount, tmp_path):
    projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "q" / FILE_NAME))
    _make(mount / "UAT" / FILE_NAME, tp_header=[h if h != "Total TC" else "T" for h in TP_HEADER])
    with pytest.raises(projects.ProjectError) as ei:
        projects.import_project_link(base, appdb, PERSON, link=LINK)
    assert str(ei.value).startswith("This file can't be imported:\n- ")
    _make(mount / "UAT" / " .xlsx")
    with pytest.raises(projects.ProjectError) as ei:
        projects.import_project_link(base, appdb, PERSON,
                                     link=f"{HOST}/sites/Team/Shared%20Documents/UAT/%20.xlsx")
    assert str(ei.value) == "The file name gives no project name"


def test_qa_t33_link_import_single_event_and_admin_user(base, appdb, synced, target):
    _make(target)
    p = projects.import_project_link(base, appdb, PERSON, link=LINK)
    assert [e["action"] for e in appdb.events(p["id"], limit=100)] == ["imported"]
    assert p["created_by"] == PERSON.editor
    u = projects._with_repo(base, p, lambda _c, repo: repo.get_user(PERSON.email))
    assert u["role"] == "Admin" and u["first_name"] == "Pim"


# ---------------------------------------------------------------- T34: relink_link (b)

import shutil  # noqa: E402

from uat_tracker.excel.writer import OpenpyxlWorkbook  # noqa: E402
from uat_tracker.mail.sender import FakeSender  # noqa: E402
from uat_tracker.services.sync import SyncService  # noqa: E402


@pytest.fixture
def alpha(base, appdb, tmp_path) -> dict:
    """Project Alpha, read once and with App IDs written into its file."""
    p = projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "a" / "Planning Progress Alpha.xlsx"))
    cfg, repo = projects.open_project(base, p)
    try:
        run_import(cfg, repo, editor="t")
        assert SyncService(cfg, repo, opener=OpenpyxlWorkbook, sender=FakeSender()).init_ids().result == "written"
    finally:
        repo.close()
    return appdb.get_project(p["id"])


def _copy_alpha(alpha: dict, dest: Path) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(alpha["excel_path"], dest)
    return dest


def _relinked(appdb, pid: str) -> list[dict]:
    return [e for e in appdb.events(pid, limit=10_000) if e["action"] == "relinked"]


def _known(base, project) -> set[str]:
    def go(_cfg, repo):
        return {i for s in projects.SHEETS for i, r in repo.load_items(s).items()
                if r.get("active") and r.get("excel_row")}
    return projects._with_repo(base, project, go)


@pytest.mark.parametrize("link", [LINK, f'  "{LINK}?web=1"  ', ADDRESS_BAR],
                         ids=["direct", "quoted-web", "address-bar"])
def test_t34_ac21_ac22_relink_from_link(base, appdb, synced, alpha, target, link):
    _copy_alpha(alpha, target)
    before = _sha(target)
    expected = _known(base, alpha)
    assert expected
    summary = projects.relink_link(base, appdb, PERSON, alpha["id"], link)
    n = len(expected & projects.app_ids_in(target))
    assert summary == f"{n} of {len(expected)} App IDs found"
    assert appdb.get_project(alpha["id"])["excel_path"] == str(target)
    assert _sha(target) == before
    ev = _relinked(appdb, alpha["id"])
    assert len(ev) == 1
    assert ev[0]["detail"] == f"{alpha['excel_path']} -> {target} ({summary}) (SharePoint: {LINK})"
    assert ev[0]["editor"] == PERSON.editor
    # same summary as the file-based relink of the same file
    assert projects.relink(base, appdb, PERSON, alpha["id"], target) == summary


def test_t34_ac23_less_than_80_percent(base, appdb, synced, alpha, target):
    _make(target)                                            # fresh workbook: no App IDs of Alpha
    old = alpha["excel_path"]
    before = _state(appdb)
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink_link(base, appdb, PERSON, alpha["id"], LINK)
    assert type(ei.value) is projects.ProjectError
    assert str(ei.value) == ("Only 0% of this project's App IDs are in that file (need 80%). "
                             "It looks like a different file.")
    assert appdb.get_project(alpha["id"])["excel_path"] == old
    assert _state(appdb) == before


def test_t34_ac24_duplicate_file_name(base, appdb, synced, alpha, target, tmp_path):
    q = projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "q" / FILE_NAME))
    _copy_alpha(alpha, target)
    old = alpha["excel_path"]
    before = _state(appdb)
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink_link(base, appdb, PERSON, alpha["id"], LINK)
    assert type(ei.value) is projects.ProjectError
    assert str(ei.value) == f"Project '{q['name']}' already uses the file name '{FILE_NAME}'"
    assert appdb.get_project(alpha["id"])["excel_path"] == old
    assert _state(appdb) == before


def test_t34_same_file_name_of_own_project_is_fine(base, appdb, synced, tmp_path, mount):
    p = projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "p" / FILE_NAME))
    f = _make(mount / "UAT" / FILE_NAME)
    assert projects.relink_link(base, appdb, PERSON, p["id"], LINK) == "project has no items yet"
    assert appdb.get_project(p["id"])["excel_path"] == str(f)


def test_t34_validation_message(base, appdb, synced, alpha, target):
    _make(target, tp_header=[h if h != "Total TC" else "T" for h in TP_HEADER])
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink_link(base, appdb, PERSON, alpha["id"], LINK)
    assert type(ei.value) is projects.ProjectError
    assert str(ei.value).startswith("This file can't be linked:\n- ") and "Total TC" in str(ei.value)


@pytest.mark.parametrize("link", ["", "abc", f"{HOST}/:x:/s/Team/EaBc", LINK])
def test_t34_ac25_not_manager_before_any_link_work(base, appdb, alpha, monkeypatch, link):
    _add_member(base, alpha, OTHER.email, "QA")
    calls = []

    def resolve(*a, **_k):
        calls.append(a)
        raise AssertionError("resolved before permission check")
    monkeypatch.setattr(projects, "resolve_excel_link", resolve)
    before = _state(appdb)
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink_link(base, appdb, OTHER, alpha["id"], link)
    assert type(ei.value) is projects.ProjectError
    assert str(ei.value) == "Not allowed: only an Admin or PMO (of any project) can do this"
    assert not calls and _state(appdb) == before
    assert appdb.get_project(alpha["id"])["excel_path"] == alpha["excel_path"]


def test_t34_ac25_unknown_project_before_link(base, appdb, alpha, monkeypatch):
    calls = []
    monkeypatch.setattr(projects, "resolve_excel_link", lambda *a, **_k: calls.append(a))
    before = _state(appdb)
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink_link(base, appdb, PERSON, "nope", LINK)
    assert type(ei.value) is projects.ProjectError
    assert str(ei.value) == "Unknown project nope"
    assert not calls and _state(appdb) == before


@pytest.mark.parametrize("link,code", AC19_CASES, ids=[f"{c}-{i}" for i, (_l, c) in enumerate(AC19_CASES)])
def test_t34_ac26_link_errors_change_nothing(base, appdb, synced, alpha, link, code):
    before = _state(appdb)
    with pytest.raises(ExcelLinkError) as ei:
        projects.relink_link(base, appdb, PERSON, alpha["id"], link)
    assert ei.value.code == code
    assert appdb.get_project(alpha["id"])["excel_path"] == alpha["excel_path"]
    assert _state(appdb) == before


@pytest.mark.parametrize("fail_on", [1, 2], ids=["validate", "app_ids_in"])
def test_t34_ac26_unreadable(base, appdb, synced, alpha, target, monkeypatch, fail_on):
    _copy_alpha(alpha, target)
    real = reader.copy_shared
    seen = []

    def flaky(src, *a, **k):
        if Path(src) == target:
            seen.append(1)
            if len(seen) == fail_on:
                raise PermissionError("cloud file not available")
        return real(src, *a, **k)
    monkeypatch.setattr(reader, "copy_shared", flaky)
    before = _state(appdb)
    with pytest.raises(ExcelLinkError) as ei:
        projects.relink_link(base, appdb, PERSON, alpha["id"], LINK)
    assert ei.value.code == "link_unreadable" and ei.value.path == str(target)
    assert isinstance(ei.value.__cause__, PermissionError)
    assert len(seen) == fail_on
    assert appdb.get_project(alpha["id"])["excel_path"] == alpha["excel_path"]
    assert _state(appdb) == before


def test_t34_ac27_relink_unchanged(base, appdb, alpha, tmp_path, monkeypatch):
    moved = _copy_alpha(alpha, tmp_path / "moved" / "Alpha renamed.xlsx")
    summary = projects.relink(base, appdb, PERSON, alpha["id"], str(moved))       # str still accepted
    ev = _relinked(appdb, alpha["id"])
    assert len(ev) == 1 and ev[0]["detail"] == f"{alpha['excel_path']} -> {moved} ({summary})"
    assert "SharePoint" not in ev[0]["detail"]
    assert appdb.get_project(alpha["id"])["excel_path"] == str(moved)

    def busy(*_a, **_k):
        raise OSError("busy")
    monkeypatch.setattr(reader, "copy_shared", busy)
    with pytest.raises(OSError) as ei:                                             # plain OSError as before
        projects.relink(base, appdb, PERSON, alpha["id"], moved)
    assert not isinstance(ei.value, projects.ProjectError)


def test_t34_relink_link_uses_relink_file_with_url(base, appdb, synced, alpha, target, monkeypatch, tmp_path):
    _copy_alpha(alpha, target)
    seen = []
    real = projects._relink_file

    def spy(b, a, person, project, path, link_url):
        seen.append((project["id"], path, link_url))
        return real(b, a, person, project, path, link_url)
    monkeypatch.setattr(projects, "_relink_file", spy)
    projects.relink_link(base, appdb, PERSON, alpha["id"], f"{LINK}?web=1")
    assert seen == [(alpha["id"], target, LINK)]
    other = _copy_alpha(alpha, tmp_path / "o" / "Alpha 2.xlsx")
    projects.relink(base, appdb, PERSON, alpha["id"], str(other))
    assert seen[-1] == (alpha["id"], other, None) and isinstance(seen[-1][1], Path)


# ---------------------------------------------------------------- T34 QA

_BAD_TP = [h if h != "Total TC" else "T" for h in TP_HEADER]


@pytest.mark.parametrize("archived", [False, True], ids=["normal", "archived"])
def test_qa_t34_duplicate_with_archived_project(base, appdb, synced, alpha, target, tmp_path, archived):
    q = projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "q" / FILE_NAME.upper()))
    if archived:
        projects.set_archived(base, appdb, PERSON, q["id"], True)
    _copy_alpha(alpha, target)
    before = _state(appdb)
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink_link(base, appdb, PERSON, alpha["id"], LINK)
    assert type(ei.value) is projects.ProjectError
    assert str(ei.value) == f"Project '{q['name']}' already uses the file name '{FILE_NAME}'"
    assert appdb.get_project(alpha["id"])["excel_path"] == alpha["excel_path"]
    assert _state(appdb) == before


@pytest.mark.parametrize("archived", [False, True], ids=["normal", "archived"])
def test_qa_t34_old_relink_duplicate_same_message(base, appdb, alpha, tmp_path, archived):
    q = projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "q" / FILE_NAME))
    if archived:
        projects.set_archived(base, appdb, PERSON, q["id"], True)
    f = _copy_alpha(alpha, tmp_path / "elsewhere" / FILE_NAME)
    before = _state(appdb)
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink(base, appdb, PERSON, alpha["id"], f)
    assert type(ei.value) is projects.ProjectError
    assert str(ei.value) == f"Project '{q['name']}' already uses the file name '{FILE_NAME}'"
    assert _state(appdb) == before


def test_qa_t34_old_relink_messages(base, appdb, alpha, tmp_path):
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink(base, appdb, OTHER, "nope", tmp_path / "x.xlsx")
    assert str(ei.value) == "Not allowed: only an Admin or PMO (of any project) can do this"
    with pytest.raises(projects.ProjectError) as ei:                 # project checked before the file
        projects.relink(base, appdb, PERSON, "nope", tmp_path / "missing.xlsx")
    assert str(ei.value) == "Unknown project nope"
    bad = _make(tmp_path / "b" / "Bad.xlsx", tp_header=_BAD_TP)
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink(base, appdb, PERSON, alpha["id"], bad)
    assert str(ei.value).startswith("This file can't be linked:\n- ") and "Total TC" in str(ei.value)
    fresh = _make(tmp_path / "f" / "Fresh.xlsx")
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink(base, appdb, PERSON, alpha["id"], fresh)
    assert str(ei.value) == ("Only 0% of this project's App IDs are in that file (need 80%). "
                             "It looks like a different file.")
    assert appdb.get_project(alpha["id"])["excel_path"] == alpha["excel_path"]
    assert _relinked(appdb, alpha["id"]) == []


def test_qa_t34_old_relink_validation_before_duplicate(base, appdb, alpha, tmp_path):
    projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "q" / FILE_NAME))
    bad = _make(tmp_path / "b" / FILE_NAME, tp_header=_BAD_TP)
    with pytest.raises(projects.ProjectError, match="can't be linked"):
        projects.relink(base, appdb, PERSON, alpha["id"], bad)


def test_qa_t34_link_validation_before_duplicate(base, appdb, synced, alpha, target, tmp_path):
    projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "q" / FILE_NAME))
    _make(target, tp_header=_BAD_TP)
    with pytest.raises(projects.ProjectError, match="can't be linked") as ei:
        projects.relink_link(base, appdb, PERSON, alpha["id"], LINK)
    assert not isinstance(ei.value, ExcelLinkError)


def test_qa_t34_duplicate_before_app_id_read(base, appdb, synced, alpha, target, tmp_path, monkeypatch):
    projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "q" / FILE_NAME))
    _copy_alpha(alpha, target)

    def no_read(*_a):
        raise AssertionError("App IDs read before the duplicate check")
    monkeypatch.setattr(projects, "app_ids_in", no_read)
    with pytest.raises(projects.ProjectError, match="already uses the file name"):
        projects.relink_link(base, appdb, PERSON, alpha["id"], LINK)


def test_qa_t34_old_relink_unreadable_during_ids_plain_oserror(base, appdb, alpha, tmp_path, monkeypatch):
    moved = _copy_alpha(alpha, tmp_path / "moved" / "Alpha 2.xlsx")
    real = reader.copy_shared
    seen = []

    def flaky(src, *a, **k):
        if Path(src) == moved:
            seen.append(1)
            if len(seen) == 2:
                raise OSError("gone")
        return real(src, *a, **k)
    monkeypatch.setattr(reader, "copy_shared", flaky)
    before = _state(appdb)
    with pytest.raises(OSError) as ei:
        projects.relink(base, appdb, PERSON, alpha["id"], moved)
    assert type(ei.value) is OSError and len(seen) == 2
    assert appdb.get_project(alpha["id"])["excel_path"] == alpha["excel_path"]
    assert _state(appdb) == before


def test_qa_t34_not_manager_and_unknown_project_gives_not_allowed(base, appdb, alpha):
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink_link(base, appdb, OTHER, "nope", "abc")
    assert str(ei.value) == "Not allowed: only an Admin or PMO (of any project) can do this"


@pytest.mark.parametrize("link", ["", "abc", f"{HOST}/:x:/s/Team/EaBc"])
def test_qa_t34_unknown_project_before_bad_link(base, appdb, synced, alpha, link):
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink_link(base, appdb, PERSON, "nope", link)
    assert type(ei.value) is projects.ProjectError
    assert str(ei.value) == "Unknown project nope"


def test_qa_t34_not_found_keeps_state(base, appdb, synced, alpha, mount, target):
    before_tree = _tree(mount)
    before = _state(appdb)
    with pytest.raises(ExcelLinkError) as ei:
        projects.relink_link(base, appdb, PERSON, alpha["id"], LINK)
    assert ei.value.code == "link_not_found" and ei.value.path == str(target)
    assert _tree(mount) == before_tree and _state(appdb) == before


@pytest.mark.parametrize("link", [
    f"{LINK}?web=1",
    f'  "{LINK}?d=w1&csf=1&web=1&e=abc"  ',
    f"{HOST}/sites/Team/Shared Documents/UAT/Planning Progress Demo.xlsx",
    ADDRESS_BAR,
    f"{HOST}/:x:/r/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx?d=w123&csf=1&web=1&e=abc",
    f"{LINK}#frag",
], ids=["web=1", "quoted-query", "unencoded", "address-bar", "share-r", "fragment"])
def test_qa_t34_event_detail_has_no_query(base, appdb, synced, alpha, target, link):
    _copy_alpha(alpha, target)
    summary = projects.relink_link(base, appdb, PERSON, alpha["id"], link)
    ev = _relinked(appdb, alpha["id"])
    assert len(ev) == 1
    d = ev[0]["detail"]
    assert "?" not in d and "#" not in d and "web=1" not in d
    assert d == f"{alpha['excel_path']} -> {target} ({summary}) (SharePoint: {LINK})"


def test_qa_t34_project_without_items(base, appdb, synced, tmp_path, target):
    p = projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "p" / "Planning Progress P.xlsx"))
    _make(target)
    assert projects.relink_link(base, appdb, PERSON, p["id"], LINK) == "project has no items yet"
    ev = _relinked(appdb, p["id"])
    assert len(ev) == 1
    assert ev[0]["detail"] == f"{p['excel_path']} -> {target} (project has no items yet) (SharePoint: {LINK})"


@pytest.mark.parametrize("role", ["Admin", "PMO"])
def test_qa_t34_manager_of_other_project_may_relink(base, appdb, synced, alpha, target, tmp_path, role):
    other = projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "z" / "Planning Progress Z.xlsx"))
    _add_member(base, other, OTHER.email, role)
    _copy_alpha(alpha, target)
    projects.relink_link(base, appdb, OTHER, alpha["id"], LINK)
    assert appdb.get_project(alpha["id"])["excel_path"] == str(target)
    assert _relinked(appdb, alpha["id"])[0]["editor"] == OTHER.editor


def test_qa_t34_unreadable_other_oserror_in_ids(base, appdb, synced, alpha, target, monkeypatch):
    _copy_alpha(alpha, target)

    def ids(_path):
        raise FileNotFoundError("placeholder vanished")
    monkeypatch.setattr(projects, "app_ids_in", ids)
    before = _state(appdb)
    with pytest.raises(ExcelLinkError) as ei:
        projects.relink_link(base, appdb, PERSON, alpha["id"], LINK)
    assert ei.value.code == "link_unreadable" and ei.value.path == str(target)
    assert isinstance(ei.value.__cause__, FileNotFoundError)
    assert _state(appdb) == before


# ---------------------------------------------------------------- T36: UI (a) import from a link

import os  # noqa: E402

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
QtWidgets = pytest.importorskip("PySide6.QtWidgets")

from uat_tracker.ui import i18n  # noqa: E402
from uat_tracker.ui import projects as ui_projects  # noqa: E402
from uat_tracker.ui import worker as ui_worker  # noqa: E402
from uat_tracker.ui.app import AppController  # noqa: E402
from uat_tracker.ui.i18n import t  # noqa: E402


@pytest.fixture(scope="module")
def qapp():
    return QtWidgets.QApplication.instance() or QtWidgets.QApplication([])


@pytest.fixture
def lang_en():
    i18n.set_lang("en")
    yield
    i18n.set_lang("en")


class _Calls:
    """Records the dialogs of the flow; getText / confirm answers are set per test."""

    def __init__(self, monkeypatch, *, text: str = LINK, ok: bool = True, yes: bool = True) -> None:
        self.text, self.ok, self.yes = text, ok, yes
        self.get_text: list[tuple] = []
        self.confirm: list[str] = []
        self.info: list[str] = []
        self.error: list[str] = []
        monkeypatch.setattr(QtWidgets.QInputDialog, "getText", self._get_text)
        monkeypatch.setattr(ui_projects, "confirm", self._confirm)
        monkeypatch.setattr(ui_projects, "info", lambda _p, text: self.info.append(text))
        monkeypatch.setattr(ui_projects, "error", lambda _p, text: self.error.append(text))
        monkeypatch.setattr(ui_worker, "run_in_worker", self.run_now)

    def _get_text(self, *a, **_k):
        self.get_text.append(a)
        return self.text, self.ok

    def _confirm(self, _parent, text: str) -> bool:
        self.confirm.append(text)
        return self.yes

    @staticmethod
    def run_now(_parent, _text, fn, on_done=None, on_error=None, *, name="job"):
        try:
            res = fn()
        except Exception as exc:                     # noqa: BLE001 - mirrors the worker
            if on_error is None:
                raise
            on_error(exc)
            return None
        if on_done:
            on_done(res)
        return None


def _has_thai(text: str) -> bool:
    return any("฀" <= ch <= "๿" for ch in text)


def test_t36_ac28_choice_box_buttons(qapp, lang_en):
    box, by_file, by_link = ui_projects._import_choice_box(None)
    texts = [b.text() for b in box.buttons()]
    assert by_file.text() == t("proj.import_file") and by_link.text() == t("proj.import_link")
    assert t("proj.import_file") in texts and t("proj.import_link") in texts
    assert box.text() == t("proj.import_how")
    assert box.button(QtWidgets.QMessageBox.Cancel) is not None
    box.deleteLater()


def test_t36_ac36_choice_box_thai(qapp, lang_en):
    i18n.set_lang("th")
    box, by_file, by_link = ui_projects._import_choice_box(None)
    assert _has_thai(by_file.text()) and _has_thai(by_link.text()) and _has_thai(box.text())
    box.deleteLater()


@pytest.mark.parametrize("pick,want", [("file", "file"), ("link", "link"), ("cancel", None)])
def test_t36_import_choice_returns_clicked_button(qapp, lang_en, monkeypatch, pick, want):
    real = ui_projects._import_choice_box

    def fake_box(parent):
        box, f, k = real(parent)
        chosen = {"file": f, "link": k, "cancel": box.button(QtWidgets.QMessageBox.Cancel)}[pick]
        monkeypatch.setattr(box, "exec", lambda: chosen.click() or 0)
        return box, f, k
    monkeypatch.setattr(ui_projects, "_import_choice_box", fake_box)
    assert ui_projects.import_choice(None) == want


def _has_import_button(win) -> bool:
    return any(b.text() == t("proj.import") for b in win.findChildren(QtWidgets.QPushButton))


def _close(win) -> None:
    win.on_closed = None
    win.close()


def test_t36_ac28_import_button_shown_by_can_manage(qapp, lang_en, base, appdb, tmp_path):
    ctl = AppController(base, appdb, tmp_path / "c.toml", start_sync=False)
    ctl.person = PERSON
    win = ui_projects.NoProjectWindow(ctl)                           # first use: everyone may manage
    assert _has_import_button(win)
    _close(win)
    projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "x" / "Planning Progress Alpha.xlsx"))
    ctl.person = OTHER                                               # not Admin / PMO anywhere
    win = ui_projects.NoProjectWindow(ctl)
    assert not _has_import_button(win)
    _close(win)


@pytest.mark.parametrize("how,called", [("link", ["link"]), ("file", ["file"]), (None, [])])
def test_t36_controller_routes_by_choice(qapp, base, appdb, tmp_path, monkeypatch, how, called):
    calls: list[str] = []
    monkeypatch.setattr(ui_projects, "import_choice", lambda _parent: how)
    monkeypatch.setattr(ui_projects, "import_link_flow", lambda *a, **k: calls.append("link"))
    monkeypatch.setattr(ui_projects, "import_project_flow", lambda *a, **k: calls.append("file"))
    ctl = AppController(base, appdb, tmp_path / "c.toml", start_sync=False)
    ctl.person = PERSON
    ctl.import_project()
    assert calls == called


@pytest.mark.parametrize("how", ["link", "file"])
def test_t36_controller_flow_opens_new_project(qapp, base, appdb, tmp_path, monkeypatch, how):
    seen: dict = {}

    def flow(parent, b, db, person, on_done=None):
        seen["args"] = (parent, b, db, person)
        on_done({"id": "P-NEW"})
    monkeypatch.setattr(ui_projects, "import_choice", lambda _parent: how)
    monkeypatch.setattr(ui_projects, "import_link_flow" if how == "link" else "import_project_flow", flow)
    ctl = AppController(base, appdb, tmp_path / "c.toml", start_sync=False)
    ctl.person = PERSON
    opened: list[str] = []
    monkeypatch.setattr(ctl, "open_project", lambda pid: opened.append(pid))
    ctl.import_project()
    assert seen["args"] == (None, base, appdb, PERSON)
    assert opened == ["P-NEW"]


def test_t36_ac29_import_from_link(qapp, lang_en, base, appdb, synced, target, monkeypatch):
    _make(target)
    calls = _Calls(monkeypatch, text=f'  "{LINK}?web=1"  ')
    done: list[dict] = []
    p = ui_projects.import_link_flow(None, base, appdb, PERSON, on_done=done.append)
    assert p is not None and p["name"] == "Demo" and p["excel_path"] == str(target)
    assert calls.get_text[0][1:] == (t("proj.xl_link_title"), t("proj.xl_link_prompt"))
    assert calls.confirm == [t("proj.import_link_q", u=LINK, p=target, n="Demo")]
    assert LINK in calls.confirm[0] and str(target) in calls.confirm[0] and "Demo" in calls.confirm[0]
    assert [x["id"] for x in appdb.projects()] == [p["id"]]
    assert calls.error == []
    assert len(calls.info) == 1 and calls.info[0].startswith(t("proj.imported", n="Demo", r=""))
    assert [d["id"] for d in done] == [p["id"]]


@pytest.mark.parametrize("text,ok,yes", [(LINK, False, True), ("", True, True), ("   ", True, True),
                                         (LINK, True, False)], ids=["cancel", "empty", "blank", "confirm-no"])
def test_t36_ac30_cancel_saves_nothing(qapp, lang_en, base, appdb, synced, target, monkeypatch, text, ok, yes):
    _make(target)
    calls = _Calls(monkeypatch, text=text, ok=ok, yes=yes)
    done: list[dict] = []
    assert ui_projects.import_link_flow(None, base, appdb, PERSON, on_done=done.append) is None
    assert _state(appdb) == (0, 0)
    assert calls.error == [] and calls.info == [] and done == []
    assert len(calls.confirm) == (1 if ok and text.strip() else 0)


AC31_CASES = [c for c in AC19_CASES if c[1] in ("link_sharing", "link_not_xlsx", "link_not_synced",
                                                "link_not_found")]


@pytest.mark.parametrize("lang", ["en", "th"])
@pytest.mark.parametrize("link,code", AC31_CASES, ids=[f"{c}-{i}" for i, (_l, c) in enumerate(AC31_CASES)])
def test_t36_ac31_ac36_link_errors(qapp, lang_en, base, appdb, synced, target, monkeypatch, link, code, lang):
    i18n.set_lang(lang)
    calls = _Calls(monkeypatch, text=link)
    assert ui_projects.import_link_flow(None, base, appdb, PERSON) is None
    shown = str(target) if code == "link_not_found" else sharepoint_link.clean_link(link)
    assert calls.error == [t(f"proj.xl_{code}", p=shown)]
    assert shown in calls.error[0]
    assert calls.confirm == [] and calls.info == []
    assert _state(appdb) == (0, 0)
    assert _has_thai(calls.error[0]) == (lang == "th")


def test_t36_invalid_workbook_shows_problems(qapp, lang_en, base, appdb, synced, target, monkeypatch):
    _make(target)
    monkeypatch.setattr(projects, "validate_link_file", lambda _b, _p: ["missing header X", "missing header Y"])
    calls = _Calls(monkeypatch)
    assert ui_projects.import_link_flow(None, base, appdb, PERSON) is None
    assert calls.error == ["missing header X\nmissing header Y"]
    assert calls.confirm == [] and _state(appdb) == (0, 0)


def test_t36_unreadable_shows_link_text(qapp, lang_en, base, appdb, synced, target, monkeypatch):
    _make(target)

    def locked(*_a, **_k):
        raise PermissionError("locked")
    monkeypatch.setattr(reader, "copy_shared", locked)
    calls = _Calls(monkeypatch)
    assert ui_projects.import_link_flow(None, base, appdb, PERSON) is None
    assert calls.error == [t("proj.xl_link_unreadable", p=target)]
    assert calls.confirm == [] and _state(appdb) == (0, 0)


def test_t36_service_error_after_confirm(qapp, lang_en, base, appdb, synced, target, tmp_path, monkeypatch):
    _make(target)
    projects.import_project(base, appdb, PERSON, path=_make(tmp_path / "x" / FILE_NAME))   # same file name
    before = _state(appdb)
    calls = _Calls(monkeypatch)
    done: list[dict] = []
    assert ui_projects.import_link_flow(None, base, appdb, PERSON, on_done=done.append) is None
    assert len(calls.confirm) == 1 and len(calls.error) == 1
    assert calls.error[0] != "" and not calls.error[0].startswith("proj.")
    assert calls.info == [] and done == [] and _state(appdb) == before


def test_t36_service_link_error_after_confirm(qapp, lang_en, base, appdb, synced, target, monkeypatch):
    _make(target)

    def gone(*_a, **_k):
        raise ExcelLinkError("link_not_found", target, "gone")
    monkeypatch.setattr(projects, "import_project_link", gone)
    calls = _Calls(monkeypatch)
    assert ui_projects.import_link_flow(None, base, appdb, PERSON) is None
    assert calls.error == [t("proj.xl_link_not_found", p=target)] and calls.info == []


def test_t36_link_error_text(lang_en):
    exc = ExcelLinkError("link_not_synced", "https://x/y.xlsx", "msg")
    assert ui_projects.link_error_text(exc) == t("proj.xl_link_not_synced", p="https://x/y.xlsx")


def test_t36_file_flow_unchanged(qapp, lang_en, base, appdb, tmp_path, monkeypatch):
    path = _make(tmp_path / "x" / "Planning Progress Alpha.xlsx")
    calls = _Calls(monkeypatch)
    monkeypatch.setattr(QtWidgets.QFileDialog, "getOpenFileName", lambda *a, **k: (str(path), ""))
    names: list[str] = []

    def spy(*a, name="job", **k):
        names.append(name)
        return calls.run_now(*a, name=name, **k)
    monkeypatch.setattr(ui_worker, "run_in_worker", spy)
    done: list[dict] = []
    p = ui_projects.import_project_flow(None, base, appdb, PERSON, on_done=done.append)
    assert p is not None and p["excel_path"] == str(path)
    assert calls.confirm == [t("proj.import_q", f=path.name, n="Alpha")]
    assert calls.get_text == [] and calls.error == []
    assert len(calls.info) == 1 and calls.info[0].startswith(t("proj.imported", n="Alpha", r=""))
    assert [d["id"] for d in done] == [p["id"]] and names == ["import-project"]
