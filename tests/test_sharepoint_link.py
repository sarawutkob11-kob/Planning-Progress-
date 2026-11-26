"""SharePoint link resolvers (T31): characterization of the folder resolver + resolve_file_link (26-Nov-2026).

No registry, no disk: libraries are injected (SyncedLibrary) and tmp_path stands in for the OneDrive MountPoint.
"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from uat_tracker.services import sharepoint_link
from uat_tracker.services.projects import (
    AttachmentDirError, ExcelLinkError, ProjectError, resolve_attachment_link,
)
from uat_tracker.services.sharepoint_link import (
    SyncedLibrary, link_location, resolve_file_link, resolve_link,
)

HOST = "https://t.sharepoint.com"
LIB_URL = f"{HOST}/sites/Team/Shared Documents"
FILE_NAME = "Planning Progress Demo.xlsx"
AC1 = f"{HOST}/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx"


@pytest.fixture(autouse=True)
def _no_registry(monkeypatch):
    def boom():
        raise AssertionError("registry must not be read in tests")
    monkeypatch.setattr(sharepoint_link, "synced_libraries", boom)


@pytest.fixture
def mount(tmp_path: Path) -> Path:
    return tmp_path / "OneDrive - Org" / "Team - Documents"


@pytest.fixture
def libs(mount: Path) -> list[SyncedLibrary]:
    return [SyncedLibrary(LIB_URL, str(mount))]


@pytest.fixture
def target(mount: Path) -> Path:
    return mount / "UAT" / FILE_NAME


# ---------------------------------------------------------------- characterization of the folder resolver (AC-12)

FOLDER_LINKS = [
    f"{HOST}/sites/Team/Shared%20Documents/UAT/Results",
    f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FTeam%2FShared%20Documents%2FUAT"
    f"%2FResults&viewid=abc",
    f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FTeam%2FShared%20Documents"
    f"%2FUAT%2FResults",
    f"{HOST}/:f:/r/sites/Team/Shared%20Documents/UAT/Results?csf=1&web=1&e=x",
]


@pytest.mark.parametrize("link", FOLDER_LINKS)
def test_resolve_link_folder_forms(link: str, libs, mount: Path) -> None:
    assert resolve_link(link, libs) == mount / "UAT" / "Results"


def test_resolve_link_library_view_without_id(libs, mount: Path) -> None:
    link = f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx"
    assert link_location(link) == ("t.sharepoint.com", ["sites", "Team", "Shared Documents"])
    assert resolve_link(link, libs) == mount


def test_resolve_link_longest_prefix_wins(tmp_path: Path) -> None:
    libs = [SyncedLibrary(LIB_URL, str(tmp_path / "a")),
            SyncedLibrary(f"{LIB_URL}/UAT", str(tmp_path / "b"))]
    assert resolve_link(FOLDER_LINKS[0], libs) == tmp_path / "b" / "Results"


def test_resolve_link_dotdot_is_invalid(libs) -> None:
    with pytest.raises(AttachmentDirError) as exc:
        resolve_link(f"{HOST}/sites/Team/Shared%20Documents/UAT/../../x", libs)
    assert exc.value.code == "link_invalid"


def test_resolve_link_not_synced() -> None:
    with pytest.raises(AttachmentDirError) as exc:
        resolve_link(FOLDER_LINKS[0], [])
    assert exc.value.code == "link_not_synced"


def test_resolve_link_opaque_sharing(libs) -> None:
    with pytest.raises(AttachmentDirError) as exc:
        resolve_link(f"{HOST}/:x:/s/Team/EaBcDeF123", libs)
    assert exc.value.code == "link_sharing"
    assert not isinstance(exc.value, ExcelLinkError)


# ---------------------------------------------------------------- resolve_file_link: success (AC-1..AC-4, AC-7)

@pytest.mark.parametrize("link", [
    AC1,
    AC1 + "?web=1",
    f"{HOST}/sites/Team/Shared Documents/UAT/{FILE_NAME}",
    f'  "{AC1}"  ',
])
def test_ac1_direct_path(link: str, libs, target: Path) -> None:
    assert resolve_file_link(link, libs) == target


def test_ac2_address_bar(libs, target: Path) -> None:
    link = (f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FTeam%2FShared%20Documents"
            f"%2FUAT%2FPlanning%20Progress%20Demo.xlsx&parent=%2Fsites%2FTeam%2FShared%20Documents%2FUAT")
    assert resolve_file_link(link, libs) == target


def test_ac3_sharing_path_link(libs, target: Path) -> None:
    link = f"{HOST}/:x:/r/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx?d=w123&csf=1&web=1&e=abc"
    assert resolve_file_link(link, libs) == target


def test_ac4_longest_prefix_wins(tmp_path: Path) -> None:
    libs = [SyncedLibrary(LIB_URL, str(tmp_path / "a")),
            SyncedLibrary(f"{LIB_URL}/UAT", str(tmp_path / "b"))]
    assert resolve_file_link(AC1, libs) == tmp_path / "b" / FILE_NAME


def test_ac7_upper_case_extension(libs, mount: Path) -> None:
    link = f"{HOST}/sites/Team/Shared%20Documents/UAT/DEMO.XLSX"
    assert resolve_file_link(link, libs) == mount / "UAT" / "DEMO.XLSX"


def test_resolver_does_not_touch_disk(libs, target: Path, mount: Path) -> None:
    assert not target.exists()
    resolve_file_link(AC1, libs)
    assert not mount.exists()


# ---------------------------------------------------------------- resolve_file_link: errors

@pytest.mark.parametrize("link", [
    f"{HOST}/:x:/s/Team/EaBcDeF123",
    f"{HOST}/:x:/g/personal/u/EaBc",
    f"{HOST}/sites/Team/_layouts/15/Doc.aspx?sourcedoc=%7BGUID%7D&file=Planning%20Progress%20Demo.xlsx"
    "&action=default",
    f"{HOST}/sites/Team/_layouts/15/WopiFrame.aspx?sourcedoc=%7BGUID%7D&file=Planning%20Progress%20Demo.xlsx"
    "&action=default",
])
def test_ac5_opaque_links(link: str, libs) -> None:
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(f" {link} ", libs)
    assert exc.value.code == "link_sharing"
    assert exc.value.path == link


@pytest.mark.parametrize("link", [
    f"{HOST}/sites/Team/Shared%20Documents/UAT/Report.xls",
    f"{HOST}/sites/Team/Shared%20Documents/UAT/Report.xlsm",
    f"{HOST}/sites/Team/Shared%20Documents/UAT/Data.csv",
    f"{HOST}/sites/Team/Shared%20Documents/UAT",
    f"{HOST}/:f:/r/sites/Team/Shared%20Documents/UAT",
])
def test_ac6_not_xlsx_before_sync_check(link: str) -> None:
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(link, [])
    assert exc.value.code == "link_not_xlsx"
    assert exc.value.path == link


def test_ac6_library_view_without_id_is_not_xlsx(libs) -> None:
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx", libs)
    assert exc.value.code == "link_not_xlsx"


@pytest.mark.parametrize("link", ["abc", "C:\\x.xlsx", "ftp://t.sharepoint.com/a.xlsx"])
def test_ac8_not_a_url(link: str, libs) -> None:
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(link, libs)
    assert exc.value.code == "link_invalid"


@pytest.mark.parametrize("libraries", [
    [],
    [SyncedLibrary("https://other.sharepoint.com/sites/Team/Shared Documents", "X:/other")],
    [SyncedLibrary(f"{HOST}/sites/Other/Shared Documents", "X:/other")],
])
def test_ac9_not_synced(libraries) -> None:
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(AC1, libraries)
    assert exc.value.code == "link_not_synced"
    assert exc.value.path == AC1
    assert AC1 in str(exc.value)


def test_ac11_dotdot_is_invalid_file_error(libs) -> None:
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(f"{HOST}/sites/Team/Shared%20Documents/UAT/../../x.xlsx", libs)
    assert exc.value.code == "link_invalid"
    assert isinstance(exc.value, ProjectError)
    assert not isinstance(exc.value, AttachmentDirError)


def test_default_libraries_come_from_synced_libraries(monkeypatch, libs, target: Path) -> None:
    monkeypatch.setattr(sharepoint_link, "synced_libraries", lambda: libs)
    assert resolve_file_link(AC1) == target


# ---------------------------------------------------------------- QA additions (T31)

def _base(tmp_path: Path) -> SimpleNamespace:
    root = tmp_path / "app"
    (root / "projects").mkdir(parents=True)
    return SimpleNamespace(root_dir=root)


@pytest.mark.parametrize("page", ["Doc.aspx", "WopiFrame.aspx", "doc.ASPX", "WOPIFRAME.aspx"])
@pytest.mark.parametrize("no_library", [False, True])
def test_ac5_office_page_with_file_param_is_sharing(page: str, no_library: bool, libs) -> None:
    """Doc.aspx / WopiFrame.aspx with file=...xlsx: never guessed from file= (E-3), even with no library."""
    link = (f"{HOST}/sites/Team/_layouts/15/{page}?sourcedoc=%7BGUID%7D&file=Planning%20Progress%20Demo.xlsx"
            "&action=default&mobileredirect=true")
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(link, [] if no_library else libs)
    assert exc.value.code == "link_sharing"
    assert exc.value.path == link
    assert not isinstance(exc.value, AttachmentDirError)


def test_ac5_opaque_sharing_without_any_library() -> None:
    link = f"{HOST}/:x:/s/Team/EaBcDeF123"
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(link, [])
    assert exc.value.code == "link_sharing"


def test_ac5_sharing_r_without_path_is_sharing() -> None:
    link = f"{HOST}/:x:/r"
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(link, [])
    assert exc.value.code == "link_sharing"
    assert exc.value.path == link


@pytest.mark.parametrize("link", [
    # ?id= pointing to a folder
    f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FTeam%2FShared%20Documents%2FUAT"
    "&viewid=abc",
    # ?RootFolder= pointing to a folder
    f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FTeam%2FShared%20Documents",
    # library view without ?id=
    f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx",
    f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?viewid=abc",
    # sharing link to a folder (with query)
    f"{HOST}/:f:/r/sites/Team/Shared%20Documents/UAT?csf=1&web=1&e=x",
    # .xlsx only in the middle of the name
    f"{HOST}/sites/Team/Shared%20Documents/UAT/Demo.xlsx.bak",
    # bare host
    f"{HOST}/",
    HOST,
])
def test_ac6_folder_links_not_xlsx_even_without_library(link: str) -> None:
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(link, [])
    assert exc.value.code == "link_not_xlsx"
    assert exc.value.path == link


def test_ac2_onedrive_personal_address_bar(tmp_path: Path) -> None:
    lib = SyncedLibrary("https://t-my.sharepoint.com/personal/u_org_com/Documents", str(tmp_path / "OneDrive"))
    link = ("https://t-my.sharepoint.com/personal/u_org_com/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fu_org_com"
            "%2FDocuments%2FUAT%2FPlanning%20Progress%20Demo.xlsx&parent=%2Fpersonal%2Fu_org_com%2FDocuments%2FUAT")
    assert resolve_file_link(link, [lib]) == tmp_path / "OneDrive" / "UAT" / FILE_NAME


def test_ac2_root_folder_param_points_to_file(libs, target: Path) -> None:
    link = (f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2Fsites%2FTeam%2F"
            "Shared%20Documents%2FUAT%2FPlanning%20Progress%20Demo.xlsx")
    assert resolve_file_link(link, libs) == target


def test_ac1_case_insensitive_host_and_library(libs, target: Path) -> None:
    link = "HTTPS://T.SharePoint.com/sites/team/shared%20documents/UAT/Planning%20Progress%20Demo.xlsx?web=1"
    got = resolve_file_link(link, libs)
    assert got.name == FILE_NAME
    assert str(got).casefold() == str(target).casefold()


def test_ac1_angle_brackets_and_whitespace_are_cleaned(libs, target: Path) -> None:
    assert resolve_file_link(f"\n<{AC1}>\t", libs) == target


def test_ac1_return_type_is_path(libs) -> None:
    assert isinstance(resolve_file_link(AC1, libs), Path)


def test_ac4_longest_prefix_wins_regardless_of_order(tmp_path: Path) -> None:
    libs = [SyncedLibrary(f"{LIB_URL}/UAT", str(tmp_path / "b")),
            SyncedLibrary(LIB_URL, str(tmp_path / "a"))]
    assert resolve_file_link(AC1, libs) == tmp_path / "b" / FILE_NAME


def test_ac4_prefix_must_match_whole_segments(tmp_path: Path) -> None:
    """A library 'Shared Doc' must not match 'Shared Documents'."""
    libs = [SyncedLibrary(f"{HOST}/sites/Team/Shared Doc", str(tmp_path / "x"))]
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(AC1, libs)
    assert exc.value.code == "link_not_synced"


def test_ac7_mixed_case_extension(libs, mount: Path) -> None:
    link = f"{HOST}/sites/Team/Shared%20Documents/UAT/Demo.XlSx"
    assert resolve_file_link(link, libs) == mount / "UAT" / "Demo.XlSx"


@pytest.mark.parametrize("link", [
    "", "   ", '""', "abc", "C:\\x.xlsx", "ftp://t.sharepoint.com/a.xlsx", "file:///C:/x.xlsx",
    "https:///a.xlsx", "//t.sharepoint.com/a.xlsx",
])
def test_ac8_invalid_has_cleaned_path(link: str, libs) -> None:
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(link, libs)
    assert exc.value.code == "link_invalid"
    assert exc.value.path == link.strip().strip('"<>').strip()
    assert not isinstance(exc.value, AttachmentDirError)


def test_ac8_none_is_invalid(libs) -> None:
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(None, libs)  # type: ignore[arg-type]
    assert exc.value.code == "link_invalid"


def test_ac9_library_deeper_than_link_is_not_synced(tmp_path: Path) -> None:
    libs = [SyncedLibrary(f"{LIB_URL}/Other", str(tmp_path / "o"))]
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(f'"{AC1}"', libs)
    assert exc.value.code == "link_not_synced"
    assert exc.value.path == AC1


@pytest.mark.parametrize("link", [
    f"{HOST}/sites/Team/Shared%20Documents/UAT/../../x.xlsx",
    f"{HOST}/sites/Team/Shared%20Documents/UAT/%2E%2E/%2E%2E/x.xlsx",
    f"{HOST}/sites/Team/Shared%20Documents/./UAT/x.xlsx",
    f"{HOST}/sites/Team/Shared%20Documents/UAT/%20..%20/x.xlsx",
    f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FTeam%2FShared%20Documents%2F..%2F"
    "..%2Fx.xlsx",
    f"{HOST}/:x:/r/sites/Team/Shared%20Documents/UAT/../../x.xlsx",
])
@pytest.mark.parametrize("no_library", [False, True])
def test_ac11_dot_segments_invalid(link: str, no_library: bool, libs) -> None:
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(link, [] if no_library else libs)
    assert exc.value.code == "link_invalid"
    assert exc.value.path == link


def test_ac11_dotdot_to_non_xlsx_is_invalid_not_not_xlsx(libs) -> None:
    with pytest.raises(ExcelLinkError) as exc:
        resolve_file_link(f"{HOST}/sites/Team/Shared%20Documents/UAT/../..", libs)
    assert exc.value.code == "link_invalid"


def test_excel_link_error_shape() -> None:
    err = ExcelLinkError("link_not_found", Path("C:/x/y.xlsx"), "msg")
    assert err.code == "link_not_found"
    assert err.path == str(Path("C:/x/y.xlsx"))
    assert isinstance(err.path, str)
    assert str(err) == "msg"
    assert isinstance(err, ProjectError)
    assert not isinstance(err, AttachmentDirError)
    assert not issubclass(AttachmentDirError, ExcelLinkError)


@pytest.mark.parametrize("link, code", [
    (f"{HOST}/:x:/s/Team/EaBcDeF123", "link_sharing"),
    ("abc", "link_invalid"),
    (f"{HOST}/sites/Team/Shared%20Documents/UAT/Report.xls", "link_not_xlsx"),
    (AC1, "link_not_synced"),
])
def test_file_errors_are_excel_link_errors_with_link_in_message(link: str, code: str) -> None:
    with pytest.raises(ProjectError) as exc:
        resolve_file_link(link, [])
    assert type(exc.value) is ExcelLinkError
    assert exc.value.code == code
    assert link in str(exc.value)


# ---- AC-12: folder (attachment) links keep raising AttachmentDirError, never ExcelLinkError

@pytest.mark.parametrize("link, code", [
    ("abc", "link_invalid"),
    (f"{HOST}/:x:/s/Team/EaBcDeF123", "link_sharing"),
    (f"{HOST}/:f:/s/Team/EaBcDeF123", "link_sharing"),
    (f"{HOST}/sites/Team/Shared%20Documents/UAT/../../x", "link_invalid"),
])
def test_ac12_resolve_link_errors_stay_attachment_errors(link: str, code: str, libs) -> None:
    with pytest.raises(AttachmentDirError) as exc:
        resolve_link(link, libs)
    assert exc.value.code == code
    assert not isinstance(exc.value, ExcelLinkError)


def test_ac12_resolve_link_has_no_file_only_rules(libs, mount: Path) -> None:
    """The folder resolver has no .xlsx / Doc.aspx rules (those are file-only)."""
    assert resolve_link(AC1, libs) == mount / "UAT" / FILE_NAME
    assert resolve_link(f"{HOST}/sites/Team/Shared%20Documents/UAT/Report.csv", libs) == mount / "UAT" / "Report.csv"
    doc = f"{HOST}/sites/Team/Shared%20Documents/_layouts/15/Doc.aspx?sourcedoc=x"
    assert resolve_link(doc, libs) == mount / "_layouts" / "15" / "Doc.aspx"


def test_ac12_link_location_unchanged() -> None:
    assert link_location(f' "{HOST}/:f:/r/sites/Team/Shared%20Documents/UAT?e=1" ') == (
        "t.sharepoint.com", ["sites", "Team", "Shared Documents", "UAT"])
    assert link_location(f"{HOST}/sites/Team/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FTeam") == (
        "t.sharepoint.com", ["sites", "Team"])


def test_ac12_resolve_attachment_link_ok(monkeypatch, tmp_path: Path, libs, mount: Path) -> None:
    monkeypatch.setattr(sharepoint_link, "synced_libraries", lambda: libs)
    folder = mount / "UAT" / "Results"
    folder.mkdir(parents=True)
    link = f"{HOST}/:f:/r/sites/Team/Shared%20Documents/UAT/Results"
    assert resolve_attachment_link(_base(tmp_path), link) == folder


@pytest.mark.parametrize("link, code", [
    ("abc", "link_invalid"),
    (f"{HOST}/:x:/s/Team/EaBcDeF123", "link_sharing"),
    (f"{HOST}/sites/Other/Shared%20Documents/UAT", "link_not_synced"),
    (f"{HOST}/sites/Team/Shared%20Documents/UAT/Missing", "link_not_found"),
])
def test_ac12_resolve_attachment_link_errors_are_attachment_errors(monkeypatch, tmp_path: Path, libs,
                                                                   link: str, code: str) -> None:
    monkeypatch.setattr(sharepoint_link, "synced_libraries", lambda: libs)
    with pytest.raises(AttachmentDirError) as exc:
        resolve_attachment_link(_base(tmp_path), link)
    assert exc.value.code == code
    assert not isinstance(exc.value, ExcelLinkError)
