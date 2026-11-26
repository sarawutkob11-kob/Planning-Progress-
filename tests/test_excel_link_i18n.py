"""T35: i18n keys for importing / relinking a project's Excel file from a SharePoint link (AC-35, AC-36, AC-37)."""
from __future__ import annotations

import re
from collections.abc import Iterator

import pytest

from uat_tracker.ui import i18n
from uat_tracker.ui.i18n import TEXT, t

LINK_CODES = ["link_invalid", "link_sharing", "link_not_xlsx", "link_not_synced", "link_not_found",
              "link_unreadable"]
NEW_KEYS = ["proj.import_how", "proj.import_file", "proj.import_link", "proj.relink_link",
            "proj.xl_link_title", "proj.xl_link_prompt", "proj.import_link_q", "proj.relink_link_q",
            *[f"proj.xl_{c}" for c in LINK_CODES]]
THAI = re.compile("[฀-๿]")

# Values of proj.att_link_* before this feature (AC-37).
ATT_LINK_SNAPSHOT: dict[str, tuple[str, str]] = {
    'proj.att_link': (
        'Paste SharePoint link…',
        'วางลิงก์ SharePoint…',
    ),
    'proj.att_link_title': (
        'SharePoint link',
        'ลิงก์ SharePoint',
    ),
    'proj.att_link_prompt': (
        'Open the SharePoint folder in the browser and paste the URL from the address bar:',
        'เปิดโฟลเดอร์ SharePoint ในเบราว์เซอร์ แล้ววาง URL จากแถบที่อยู่ (address bar):',
    ),
    'proj.att_link_q': (
        "Save uploaded result files of '{n}' in SharePoint:\n{u}\n\nThe app copies them to the folder OneDrive syncs on this PC:\n{p}\nOneDrive then uploads them. Continue?",
        "เก็บไฟล์ผลทดสอบที่อัปโหลดของ '{n}' ที่ SharePoint:\n{u}\n\nแอปจะ copy ไฟล์ไปที่โฟลเดอร์ที่ OneDrive sync บนเครื่องนี้:\n{p}\nแล้ว OneDrive จะอัปโหลดขึ้น SharePoint ต่อ ใช่หรือไม่",
    ),
    'proj.att_link_local': (
        'Local folder synced by OneDrive: {p}',
        'โฟลเดอร์บนเครื่องที่ OneDrive sync: {p}',
    ),
    'proj.att_link_invalid': (
        'This is not a SharePoint folder link: {p}\nOpen the folder in the browser and copy the URL from the address bar.',
        'ลิงก์นี้ไม่ใช่ลิงก์โฟลเดอร์ SharePoint: {p}\nกรุณาเปิดโฟลเดอร์ในเบราว์เซอร์ แล้ว copy URL จากแถบที่อยู่ (address bar)',
    ),
    'proj.att_link_sharing': (
        'A sharing link ("Copy link") can\'t be used: {p}\nOpen the folder in the browser and copy the URL from the address bar instead.',
        'ใช้ลิงก์แชร์ ("Copy link") ไม่ได้: {p}\nกรุณาเปิดโฟลเดอร์ในเบราว์เซอร์ แล้ว copy URL จากแถบที่อยู่ (address bar) แทน',
    ),
    'proj.att_link_not_synced': (
        'OneDrive on this PC does not sync the library of this link:\n{p}\n\nTo fix: open the SharePoint library in the browser, click Sync or "Add shortcut to My files", wait until OneDrive has finished, then try again. With a shortcut, use "Choose folder…" and pick the folder inside your OneDrive.',
        'OneDrive บนเครื่องนี้ยังไม่ได้ sync library ของลิงก์นี้:\n{p}\n\nวิธีแก้: เปิด library ใน SharePoint ด้วยเบราว์เซอร์ กด Sync หรือ "Add shortcut to My files" รอ OneDrive sync เสร็จ แล้วลองใหม่ ถ้าใช้ shortcut ให้กด "เลือกโฟลเดอร์…" แล้วเลือกโฟลเดอร์นั้นใน OneDrive ของคุณ',
    ),
    'proj.att_link_not_found': (
        'The link points to this folder, which is not on this PC yet:\n{p}\nWait until OneDrive has finished syncing (or check the link), then try again.',
        'ลิงก์นี้ชี้ไปที่โฟลเดอร์ที่ยังไม่มีบนเครื่องนี้:\n{p}\nกรุณารอ OneDrive sync ให้เสร็จ (หรือตรวจลิงก์อีกครั้ง) แล้วลองใหม่',
    ),
}


@pytest.fixture
def thai() -> Iterator[None]:
    old = i18n.get_lang()
    i18n.set_lang("th")
    yield
    i18n.set_lang(old)


def _fields(s: str) -> set[str]:
    return set(re.findall(r"\{(\w+)\}", s))


def test_new_key_count() -> None:
    assert len(NEW_KEYS) == 14


@pytest.mark.parametrize("key", NEW_KEYS)
def test_ac35_new_keys_en_th_non_empty(key: str) -> None:
    en, th = TEXT[key]
    assert en.strip() and th.strip()
    assert en != th


@pytest.mark.parametrize("key", NEW_KEYS)
def test_ac35_placeholders_match(key: str) -> None:
    en, th = TEXT[key]
    assert _fields(en) == _fields(th)
    assert _fields(en) <= {"p", "u", "n"}


@pytest.mark.parametrize("code", LINK_CODES)
def test_ac35_every_error_code_has_key_with_path(code: str) -> None:
    key = f"proj.xl_{code}"
    assert key in TEXT
    assert _fields(TEXT[key][0]) == {"p"}
    assert "PATH-X" in t(key, p="PATH-X")


@pytest.mark.parametrize("key", ["proj.import_link_q", "proj.relink_link_q"])
def test_confirm_texts_take_url_path_name(key: str) -> None:
    s = t(key, u="URL-X", p="PATH-X", n="NAME-X")
    assert "URL-X" in s and "PATH-X" in s and "NAME-X" in s


def test_ac37_att_link_unchanged() -> None:
    now = {k: v for k, v in TEXT.items() if k.startswith("proj.att_link")}
    assert now == ATT_LINK_SNAPSHOT


def test_file_errors_are_separate_from_folder_errors() -> None:
    for code in LINK_CODES:
        assert TEXT.get(f"proj.xl_{code}") != TEXT.get(f"proj.att_{code}")


def test_ac36_thai_texts(thai: None) -> None:
    assert THAI.search(t("proj.import_link"))
    assert THAI.search(t("proj.relink_link"))
    s = t("proj.xl_link_not_synced", p="https://x.sharepoint.com/sites/A/Doc/UAT.xlsx")
    assert THAI.search(s) and "https://x.sharepoint.com/sites/A/Doc/UAT.xlsx" in s


def test_english_texts_have_no_thai() -> None:
    old = i18n.get_lang()
    i18n.set_lang("en")
    try:
        for key in NEW_KEYS:
            assert not THAI.search(t(key, u="u", p="p", n="n"))
    finally:
        i18n.set_lang(old)


def test_not_synced_names_existing_buttons() -> None:
    en, th = TEXT["proj.xl_link_not_synced"]
    assert TEXT["proj.import"][0] in en and TEXT["proj.relink"][0] in en
    assert TEXT["proj.import"][1] in th and TEXT["proj.relink"][1] in th


# --- QA (T35) -----------------------------------------------------------------------------------------------------
import ast  # noqa: E402
from pathlib import Path  # noqa: E402

SRC = Path(__file__).resolve().parents[1] / "src" / "uat_tracker"

# Exact text of the E-15 table in docs/spec.md (EN, TH) and design.md Q1 keys.
SPEC_E15: dict[str, tuple[str, str]] = {
    "proj.import_how": ("How do you want to import the Excel file?", "ต้องการนำเข้าไฟล์ Excel แบบไหน?"),
    "proj.import_file": ("Choose file…", "เลือกไฟล์…"),
    "proj.import_link": ("Import from SharePoint link…", "นำเข้าจากลิงก์ SharePoint…"),
    "proj.relink_link": ("Relink from SharePoint link…", "เลือกไฟล์ใหม่จากลิงก์ SharePoint…"),
    "proj.xl_link_title": ("SharePoint Excel file link", "ลิงก์ไฟล์ Excel บน SharePoint"),
    "proj.xl_link_prompt": (
        "Paste the link of the .xlsx file on SharePoint (in Excel: File > Info > Copy path):",
        "วางลิงก์ของไฟล์ .xlsx บน SharePoint (ใน Excel: File > Info > Copy path):"),
    "proj.import_link_q": (
        "Import the SharePoint file:\n{u}\n\nOneDrive file on this PC:\n{p}\n\nas project '{n}'? "
        "The file is not renamed and its other sheets are not touched.",
        "นำเข้าไฟล์บน SharePoint:\n{u}\n\nไฟล์ที่ OneDrive sync บนเครื่องนี้:\n{p}\n\nเป็นโปรเจกต์ '{n}' ใช่ไหม? "
        "ไฟล์จะไม่ถูกเปลี่ยนชื่อ และชีตอื่นจะไม่ถูกแตะต้อง"),
    "proj.relink_link_q": (
        "Use the SharePoint file:\n{u}\n\nOneDrive file on this PC:\n{p}\n\nas the Excel file of '{n}'?",
        "ใช้ไฟล์บน SharePoint:\n{u}\n\nไฟล์ที่ OneDrive sync บนเครื่องนี้:\n{p}\n\nเป็นไฟล์ Excel ของ '{n}' ใช่ไหม?"),
    "proj.xl_link_invalid": (
        "This is not a SharePoint file link: {p}\nIn Excel use File > Info > Copy path, then paste it here.",
        "ลิงก์นี้ไม่ใช่ลิงก์ไฟล์บน SharePoint: {p}\nใน Excel ให้ใช้ File > Info > Copy path แล้ววางที่นี่"),
    "proj.xl_link_sharing": (
        "This kind of link can't be used: {p}\nIn Excel use File > Info > Copy path, or in SharePoint use Copy link "
        "with \"People with existing access\", then paste it here.",
        "ใช้ลิงก์แบบนี้ไม่ได้: {p}\nใน Excel ให้ใช้ File > Info > Copy path หรือใน SharePoint ใช้ Copy link แบบ "
        "\"People with existing access\" แล้ววางที่นี่"),
    "proj.xl_link_not_xlsx": (
        "The link must point to an Excel .xlsx file (not a folder or another file type): {p}",
        "ลิงก์ต้องชี้ไปที่ไฟล์ Excel .xlsx (ไม่ใช่โฟลเดอร์หรือไฟล์ชนิดอื่น): {p}"),
    "proj.xl_link_not_synced": (
        "OneDrive on this PC does not sync the library of this link:\n{p}\n\nTo fix: open the SharePoint library in "
        "the browser, click Sync or \"Add shortcut to My files\", wait until OneDrive has finished, then try again. "
        "With a shortcut, use \"Import Excel\" / \"Relink file…\" and pick the file inside your OneDrive.",
        "OneDrive บนเครื่องนี้ยังไม่ได้ sync library ของลิงก์นี้:\n{p}\n\nวิธีแก้: เปิด library ใน SharePoint "
        "ด้วยเบราว์เซอร์ กด Sync หรือ \"Add shortcut to My files\" รอ OneDrive sync เสร็จ แล้วลองใหม่ "
        "ถ้าใช้ shortcut ให้กด \"นำเข้าไฟล์ Excel\" / \"เลือกไฟล์ใหม่ (Relink)…\" แล้วเลือกไฟล์ใน OneDrive ของคุณ"),
    "proj.xl_link_not_found": (
        "The link points to this file, which is not on this PC yet:\n{p}\nWait until OneDrive has finished syncing, "
        "check that its folder is not excluded in OneDrive \"Choose folders\", or check the link, then try again.",
        "ลิงก์นี้ชี้ไปที่ไฟล์ที่ยังไม่มีบนเครื่องนี้:\n{p}\nกรุณารอ OneDrive sync ให้เสร็จ "
        "ตรวจว่าโฟลเดอร์ของไฟล์ไม่ได้ถูกตัดออกใน \"Choose folders\" ของ OneDrive หรือตรวจลิงก์อีกครั้ง แล้วลองใหม่"),
    "proj.xl_link_unreadable": (
        "The file can't be read from OneDrive on this PC:\n{p}\nCheck that OneDrive is running and online, "
        "then try again.",
        "อ่านไฟล์จาก OneDrive บนเครื่องนี้ไม่ได้:\n{p}\nกรุณาตรวจว่า OneDrive ทำงานและเชื่อมต่ออินเทอร์เน็ตอยู่ "
        "แล้วลองใหม่"),
}


def _text_dict_keys() -> list[str]:
    """Literal keys of TEXT in source order (a duplicate literal silently overrides in the dict)."""
    tree = ast.parse((SRC / "ui" / "i18n.py").read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name) and node.target.id == "TEXT":
            assert isinstance(node.value, ast.Dict)
            return [k.value for k in node.value.keys if isinstance(k, ast.Constant)]
    raise AssertionError("TEXT not found")


def _excel_link_codes_in_source() -> set[str]:
    """Codes passed to ExcelLinkError(...) / _file_error(...) / returned in sharepoint_link's file-error map."""
    codes: set[str] = set()
    for rel in ("services/projects.py", "services/sharepoint_link.py"):
        src = (SRC / rel).read_text(encoding="utf-8")
        codes |= set(re.findall(r"(?:ExcelLinkError|_file_error)\(\s*[\"'](link_\w+)[\"']", src))
    return codes


def test_qa_new_keys_match_spec_e15_word_for_word() -> None:
    assert set(SPEC_E15) == set(NEW_KEYS)
    for key, (en, th) in SPEC_E15.items():
        assert TEXT[key] == (en, th), key


def test_qa_no_duplicate_literal_keys_in_text_source() -> None:
    keys = _text_dict_keys()
    dups = sorted({k for k in keys if keys.count(k) > 1})
    assert dups == []
    assert len(keys) == len(TEXT)


def test_qa_new_keys_each_defined_once_in_source() -> None:
    keys = _text_dict_keys()
    for key in NEW_KEYS:
        assert keys.count(key) == 1, key


def test_qa_every_excel_link_code_raised_in_source_has_key() -> None:
    codes = _excel_link_codes_in_source()
    assert codes, "no ExcelLinkError codes found in source"
    assert codes <= set(LINK_CODES)
    for code in codes:
        assert f"proj.xl_{code}" in TEXT


def test_qa_excel_link_error_code_renders_via_t() -> None:
    from uat_tracker.services.projects import ExcelLinkError
    for code in LINK_CODES:
        exc = ExcelLinkError(code, "C:/OneDrive/UAT.xlsx", "msg")
        s = t(f"proj.xl_{exc.code}", p=exc.path)
        assert "C:/OneDrive/UAT.xlsx" in s and "{" not in s


@pytest.mark.parametrize("key", NEW_KEYS)
def test_qa_all_new_keys_render_thai(thai: None, key: str) -> None:
    s = t(key, u="https://x.sharepoint.com/a.xlsx", p="C:/x/a.xlsx", n="Demo")
    assert THAI.search(s)
    assert "{" not in s and "}" not in s


@pytest.mark.parametrize("key", NEW_KEYS)
def test_qa_render_with_only_service_kwargs(key: str) -> None:
    """The UI passes p= for errors, u/p/n for confirms; no other placeholders may be required."""
    en, th = TEXT[key]
    for s in (en, th):
        s.format(u="u", p="p", n="n")  # raises KeyError if an unknown placeholder slipped in


def test_qa_att_link_texts_still_about_folders() -> None:
    for key in ("proj.att_link_invalid", "proj.att_link_not_found"):
        en, th = TEXT[key]
        assert "folder" in en and "โฟลเดอร์" in th
    for code in LINK_CODES:
        en, _ = TEXT[f"proj.xl_{code}"]
        assert "folder link" not in en
