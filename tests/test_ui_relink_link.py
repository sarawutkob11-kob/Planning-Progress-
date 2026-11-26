"""T37: Settings button "Relink from SharePoint link…" + relink_link_flow (AC-32, AC-33, AC-34, AC-36).

Headless (offscreen). No registry: synced_libraries is replaced by a fake library under tmp_path. Dialogs and the
worker are monkeypatched (the worker runs the job at once on the test thread).
"""

from __future__ import annotations

import os
import re
import shutil
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication, QInputDialog  # noqa: E402

from uat_tracker.config import Config  # noqa: E402
from uat_tracker.excel.writer import OpenpyxlWorkbook  # noqa: E402
from uat_tracker.mail.sender import FakeSender  # noqa: E402
from uat_tracker.services import projects, sharepoint_link  # noqa: E402
from uat_tracker.services.importer import run_import  # noqa: E402
from uat_tracker.services.profile import write_profile  # noqa: E402
from uat_tracker.services.sharepoint_link import SyncedLibrary  # noqa: E402
from uat_tracker.services.sync import SyncService  # noqa: E402
from uat_tracker.ui import i18n  # noqa: E402
from uat_tracker.ui import projects as ui_projects  # noqa: E402
from uat_tracker.ui import worker as worker_mod  # noqa: E402
from uat_tracker.ui.i18n import t  # noqa: E402
from uat_tracker.ui.widgets import QSS  # noqa: E402

from conftest import WINDOW, build_workbook  # noqa: E402

HOST = "https://t.sharepoint.com"
LIB_URL = f"{HOST}/sites/Team/Shared Documents"
LINK = f"{HOST}/sites/Team/Shared%20Documents/UAT/Planning%20Progress%20Demo.xlsx"
PERSON = projects.Person("pmo@example.test", "Pim", "Pmo")
THAI = re.compile("[฀-๿]")


# ---------------------------------------------------------------- fixtures

@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance() or QApplication([])
    app.setStyleSheet(QSS)
    return app


@pytest.fixture(autouse=True)
def _english():
    i18n.set_lang("en")
    yield
    i18n.set_lang("en")


@pytest.fixture
def base(tmp_path) -> Config:
    root = tmp_path / "root"
    return Config(excel_path=tmp_path / "unused.xlsx", data_dir=root, window_start=WINDOW.start,
                  window_end=WINDOW.end, profile_file=root / "profile.json")


@pytest.fixture
def appdb(base):
    db = projects.open_app_db(base)
    yield db
    db.close()


@pytest.fixture
def mount(tmp_path: Path) -> Path:
    m = tmp_path / "OneDrive - Org" / "Team - Documents"
    m.mkdir(parents=True)
    return m


@pytest.fixture
def target(mount: Path) -> Path:
    return mount / "UAT" / "Planning Progress Demo.xlsx"


@pytest.fixture(autouse=True)
def synced(monkeypatch, mount):
    monkeypatch.setattr(sharepoint_link, "synced_libraries", lambda: [SyncedLibrary(LIB_URL, str(mount))])


@pytest.fixture
def alpha(base, appdb, tmp_path) -> dict:
    """Project Alpha, read once and with App IDs written into its file."""
    path = tmp_path / "a" / "Planning Progress Alpha.xlsx"
    path.parent.mkdir(parents=True)
    p = projects.import_project(base, appdb, PERSON, path=build_workbook(path))
    cfg, repo = projects.open_project(base, p)
    try:
        run_import(cfg, repo, editor="t")
        assert SyncService(cfg, repo, opener=OpenpyxlWorkbook, sender=FakeSender()).init_ids().result == "written"
    finally:
        repo.close()
    return appdb.get_project(p["id"])


class Calls:
    """Recorded dialog calls + scripted answers."""

    def __init__(self) -> None:
        self.link: tuple[str, bool] = (LINK, True)
        self.answer = True
        self.get_text: list[tuple] = []
        self.confirm: list[str] = []
        self.info: list[str] = []
        self.error: list[str] = []
        self.opened: list[str] = []
        self.service: list[tuple] = []


@pytest.fixture
def calls(monkeypatch) -> Calls:
    c = Calls()

    def get_text(parent, title, prompt, *a, **k):
        c.get_text.append((title, prompt))
        return c.link

    def confirm(parent, text):
        c.confirm.append(text)
        return c.answer

    def run_now(parent, text, fn, on_done=None, on_error=None, *, name="job"):
        try:
            res = fn()
        except Exception as exc:
            if on_error is None:
                raise
            on_error(exc)
            return None
        if on_done:
            on_done(res)
        return None

    real = projects.relink_link

    def spy(*a, **k):
        c.service.append(a)
        return real(*a, **k)

    monkeypatch.setattr(QInputDialog, "getText", staticmethod(get_text))
    monkeypatch.setattr(ui_projects, "confirm", confirm)
    monkeypatch.setattr(ui_projects, "info", lambda parent, text: c.info.append(text))
    monkeypatch.setattr(ui_projects, "error", lambda parent, text: c.error.append(text))
    monkeypatch.setattr(worker_mod, "run_in_worker", run_now)
    monkeypatch.setattr(projects, "relink_link", spy)
    return c


@pytest.fixture
def screen(qapp, base, appdb, alpha, calls, monkeypatch):
    """Settings screen of Alpha opened by its Admin; app.open_project recorded."""
    from uat_tracker.ui import app as app_mod
    write_profile(base.profile_path, PERSON.email, first_name="Pim", last_name="Pmo", last_project=alpha["id"])
    ctl = app_mod.AppController(base, appdb, Path("config.toml"), start_sync=False)
    assert ctl.start()
    w = ctl.win
    w.navigate("settings")
    scr = w.screens["settings"]
    monkeypatch.setattr(ctl, "open_project", lambda pid: calls.opened.append(pid))
    yield scr
    w.on_closed = None
    ctl._close_current()


# ---------------------------------------------------------------- AC-32 / AC-36: button

def test_ac32_button_next_to_relink_for_manager(screen):
    btn = screen.relink_link_btn
    assert btn.text() == t("proj.relink_link")
    assert not btn.isHidden()
    relink = [b for b in screen.findChildren(type(btn)) if b.text() == t("proj.relink")]
    assert len(relink) == 1 and not relink[0].isHidden()
    assert btn.parentWidget() is relink[0].parentWidget()


def test_ac32_hidden_when_not_manager(screen, monkeypatch):
    monkeypatch.setattr(projects, "can_manage", lambda *a: False)
    screen.refresh()
    relink = [b for b in screen.findChildren(type(screen.relink_link_btn)) if b.text() == t("proj.relink")]
    assert screen.relink_link_btn.isHidden()
    assert relink and all(b.isHidden() for b in relink)       # same rule as Relink


def test_ac36_thai_button_text(screen):
    i18n.set_lang("th")
    screen.refresh()
    assert screen.relink_link_btn.text() == t("proj.relink_link")
    assert THAI.search(screen.relink_link_btn.text())


# ---------------------------------------------------------------- AC-33: success

def test_ac33_relink_from_link(screen, appdb, alpha, target, calls):
    shutil.copy2(alpha["excel_path"], target.parent.mkdir(parents=True) or target)
    screen.relink_link_btn.click()
    assert calls.get_text == [(t("proj.xl_link_title"), t("proj.xl_link_prompt"))]
    assert calls.confirm == [t("proj.relink_link_q", u=LINK, p=target, n=alpha["name"])]
    assert LINK in calls.confirm[0] and str(target) in calls.confirm[0]
    assert appdb.get_project(alpha["id"])["excel_path"] == str(target)
    assert len(calls.info) == 1 and calls.error == []
    summary = calls.info[0]
    assert summary.startswith(t("proj.relinked", s="")) and "App IDs found" in summary
    assert calls.opened == [alpha["id"]]
    assert calls.service and calls.service[0][3:] == (alpha["id"], LINK)


def test_ac33_link_is_cleaned_before_service(screen, alpha, target, calls):
    shutil.copy2(alpha["excel_path"], target.parent.mkdir(parents=True) or target)
    calls.link = (f'  "{LINK}?web=1"  ', True)
    screen.relink_link_btn.click()
    assert calls.service[0][4] == LINK and calls.opened == [alpha["id"]]


# ---------------------------------------------------------------- AC-34: cancel / errors

def _unchanged(appdb, alpha, calls) -> None:
    assert appdb.get_project(alpha["id"])["excel_path"] == alpha["excel_path"]
    assert calls.opened == [] and calls.info == []


@pytest.mark.parametrize("answer", [("", False), (LINK, False), ("   ", True)], ids=["cancel", "cancel-text", "blank"])
def test_ac34_get_text_cancel_or_blank(screen, appdb, alpha, target, calls, answer):
    shutil.copy2(alpha["excel_path"], target.parent.mkdir(parents=True) or target)
    calls.link = answer
    screen.relink_link_btn.click()
    assert calls.service == [] and calls.confirm == [] and calls.error == []
    _unchanged(appdb, alpha, calls)


def test_ac34_confirm_no(screen, appdb, alpha, target, calls):
    shutil.copy2(alpha["excel_path"], target.parent.mkdir(parents=True) or target)
    calls.answer = False
    screen.relink_link_btn.click()
    assert len(calls.confirm) == 1 and calls.service == [] and calls.error == []
    _unchanged(appdb, alpha, calls)


def test_ac34_resolve_error_before_confirm(screen, appdb, alpha, target, calls):
    screen.relink_link_btn.click()                          # file not there -> link_not_found, no confirm
    assert calls.confirm == [] and calls.service == []
    assert calls.error == [t("proj.xl_link_not_found", p=str(target))]
    _unchanged(appdb, alpha, calls)


def test_ac34_service_link_not_synced(screen, appdb, alpha, target, calls, monkeypatch):
    shutil.copy2(alpha["excel_path"], target.parent.mkdir(parents=True) or target)

    def boom(*a, **k):
        calls.service.append(a)
        raise projects.ExcelLinkError("link_not_synced", LINK, "not synced")
    monkeypatch.setattr(projects, "relink_link", boom)
    screen.relink_link_btn.click()
    assert len(calls.service) == 1
    assert calls.error == [t("proj.xl_link_not_synced", p=LINK)]
    _unchanged(appdb, alpha, calls)


def test_ac34_app_ids_below_80_percent(screen, appdb, alpha, target, calls):
    target.parent.mkdir(parents=True)
    build_workbook(target)                                  # same layout, no App IDs of Alpha
    screen.relink_link_btn.click()
    assert len(calls.service) == 1 and len(calls.error) == 1
    with pytest.raises(projects.ProjectError) as ei:
        projects.relink(screen.ctx.app.base, appdb, PERSON, alpha["id"], target)
    assert calls.error[0] == str(ei.value)
    _unchanged(appdb, alpha, calls)


def test_ac34_thai_error_text(screen, appdb, alpha, calls, monkeypatch):
    i18n.set_lang("th")
    monkeypatch.setattr(sharepoint_link, "synced_libraries", lambda: [])
    screen.relink_link_btn.click()
    assert calls.error == [t("proj.xl_link_not_synced", p=LINK)] and THAI.search(calls.error[0])
    _unchanged(appdb, alpha, calls)


# ---------------------------------------------------------------- old Relink unchanged

def test_old_relink_button_still_uses_file_picker(screen, appdb, alpha, tmp_path, calls, monkeypatch):
    from PySide6.QtWidgets import QFileDialog
    moved = tmp_path / "moved" / "Alpha.xlsx"
    moved.parent.mkdir()
    shutil.copy2(alpha["excel_path"], moved)
    monkeypatch.setattr(QFileDialog, "getOpenFileName", staticmethod(lambda *a, **k: (str(moved), "")))
    relink = next(b for b in screen.findChildren(type(screen.relink_link_btn)) if b.text() == t("proj.relink"))
    relink.click()
    assert calls.get_text == [] and calls.confirm == [] and calls.service == []
    assert appdb.get_project(alpha["id"])["excel_path"] == str(moved)
    assert len(calls.info) == 1 and calls.opened == [alpha["id"]]
