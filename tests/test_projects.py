"""Multi-project add-on (approved 10-Nov-2026): create / import / duplicates / isolation / forms per project /
RAMC UAT migration / per-project settings and roles / Clear project data / relink / profile + UI switching."""

from __future__ import annotations

import dataclasses
import datetime as dt
import hashlib
import json
import os
import shutil
import sqlite3
from pathlib import Path

import openpyxl
import pytest

from uat_tracker.config import Config, apply_overrides
from uat_tracker.excel.writer import OpenpyxlWorkbook
from uat_tracker.mail.sender import FakeSender
from uat_tracker.services import projects, scheduler
from uat_tracker.services.auth import Actor
from uat_tracker.services.clear_project import ClearError, clear_project, counts, restore_backup
from uat_tracker.services.edits import edit_item
from uat_tracker.services.importer import ImportPaused, S_PAUSED, run_import
from uat_tracker.services.items import add_application, add_task
from uat_tracker.services.profile import ProfileError, read_profile_data, write_profile
from uat_tracker.services.project_template import template
from uat_tracker.services.settings import save_settings
from uat_tracker.services.sync import SyncService
from uat_tracker.storage.sqlite_repo import SqliteRepository

from conftest import DEF_ROWS, TP_ROWS, WINDOW, add_test_user, build_workbook

PERSON = projects.Person("pmo@example.test", "Pim", "Pmo")
ADMIN = Actor(PERSON.email, "Admin", "Pim Pmo")
TP = "Test Progress"


def sha(path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


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


def _wb(path: Path, **kw) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    return build_workbook(path, **kw)


@pytest.fixture
def two(base, appdb, tmp_path):
    """Project Alpha (full fixture workbook) and Beta (smaller), both imported."""
    a = projects.import_project(base, appdb, PERSON, path=_wb(tmp_path / "a" / "Planning Progress Alpha.xlsx"))
    b = projects.import_project(base, appdb, PERSON,
                                path=_wb(tmp_path / "b" / "Beta Plan.xlsx", tp=TP_ROWS[:2], defects=DEF_ROWS[:1]))
    for p in (a, b):
        cfg, repo = projects.open_project(base, p)
        run_import(cfg, repo, editor="t")
        repo.close()
    return a, b


@pytest.fixture
def opened(base):
    repos = []

    def go(project):
        cfg, repo = projects.open_project(base, project)
        repos.append(repo)
        return cfg, repo
    yield go
    for r in repos:
        try:
            r.close()
        except Exception:
            pass


# --------------------------------------------------------------------------- D: create from the template

def test_create_project_file_name_template_and_empty(base, appdb, tmp_path, opened):
    folder = tmp_path / "f"
    folder.mkdir()
    p = projects.create_project(base, appdb, PERSON, name="New/Proj: X", folder=folder)
    assert p["name"] == "New/Proj: X"                                    # display name keeps the original
    path = Path(p["excel_path"])
    assert path == folder / "Planning Progress New-Proj- X.xlsx" and path.exists()
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == ["UAT Plan", "Test Progress", "Defect List"]
    for sd in template()["sheets"]:
        ws = wb[sd["name"]]
        heads = [c["header"] for c in sd["columns"]]
        assert [ws.cell(1, i).value for i in range(1, len(heads) + 1)] == heads
        assert all(ws.cell(r, c).value is None for r in range(2, ws.max_row + 1) for c in range(1, len(heads) + 1))
        for i, c in enumerate(sd["columns"], start=1):
            dim = ws.column_dimensions[openpyxl.utils.get_column_letter(i)]
            assert bool(dim.hidden) == c["hidden"]
            if c["width"] and not c["hidden"]:
                assert abs(dim.width - c["width"]) < 0.01
            assert ws.cell(1, i).font.b == c["head_style"]["font"]["b"]
        assert ws.freeze_panes == "A2"
        assert heads[-1] == "App ID" and ws.column_dimensions[openpyxl.utils.get_column_letter(len(heads))].hidden
    tp = wb["Test Progress"]
    heads = [c.value for c in tp[1]]
    assert heads[2:5] == ["Report By", "Report By Email", "CC Email"]              # approved 12a
    assert [dv.formula1 for dv in tp.data_validations.dataValidation] == ['"Green,Amber,Red"']
    assert len([r for cf in tp.conditional_formatting for r in cf.rules]) == 5
    legend_col = len(template()["sheets"][1]["columns"]) + 2                         # approved 12b
    assert tp.cell(1, legend_col).value == "Status" and tp.cell(1, legend_col + 1).value == "Definition"
    assert projects.validate_workbook(path) == []
    cfg, repo = opened(p)
    res = run_import(cfg, repo, editor="t")
    assert all(s.rows == 0 for s in res.sheets.values())
    assert repo.get_user(PERSON.email)["role"] == "Admin"                             # creator = Admin


def test_invalid_characters_and_duplicates_blocked(base, appdb, tmp_path):
    assert projects.file_name_for("A/B") == "Planning Progress A-B.xlsx"
    assert projects.file_name_for('a\\b:c*d?e"f<g>h|i') == "Planning Progress a-b-c-d-e-f-g-h-i.xlsx"
    with pytest.raises(projects.ProjectError):
        projects.file_name_for("   ")
    f1, f2, f3 = (tmp_path / n for n in ("f1", "f2", "f3"))
    for f in (f1, f2, f3):
        f.mkdir()
    projects.create_project(base, appdb, PERSON, name="A/B", folder=f1)
    with pytest.raises(projects.ProjectError, match="already exists"):            # same file in the folder
        projects.create_project(base, appdb, PERSON, name="A/B", folder=f1)
    with pytest.raises(projects.ProjectError, match="already uses the file name"):  # A:B -> same final name
        projects.create_project(base, appdb, PERSON, name="A:B", folder=f2)
    (f3 / "Planning Progress X.xlsx").write_bytes(b"not mine")
    fname, why = projects.preview_create(appdb, "X", f3)
    assert fname == "Planning Progress X.xlsx" and "already exists" in why
    with pytest.raises(projects.ProjectError):
        projects.create_project(base, appdb, PERSON, name="X", folder=f3)
    assert (f3 / "Planning Progress X.xlsx").read_bytes() == b"not mine"           # never overwritten
    other = _wb(tmp_path / "elsewhere" / "Planning Progress A-B.xlsx")
    with pytest.raises(projects.ProjectError, match="already uses the file name"):  # import: same rule
        projects.import_project(base, appdb, PERSON, path=other)
    assert len(appdb.projects()) == 1


# --------------------------------------------------------------------------- F: import

def test_import_valid_and_invalid_files(base, appdb, tmp_path, opened):
    good = _wb(tmp_path / "g" / "Planning Progress Gamma.xlsx")
    before = sha(good)
    p = projects.import_project(base, appdb, PERSON, path=good)
    assert p["name"] == "Gamma" and Path(p["excel_path"]) == good and good.exists()   # never renamed
    cfg, repo = opened(p)
    run_import(cfg, repo, editor="t")
    assert sha(good) == before                                                        # nothing written
    assert openpyxl.load_workbook(good)["NLeadAccount"]["A1"].value == "do not touch"
    assert projects.name_from_file("C:/x/My Plan.xlsx") == "My Plan"

    wb = openpyxl.Workbook()
    wb.active.title = "UAT Plan"
    wb.create_sheet("Test Progress")
    missing = tmp_path / "bad" / "Missing.xlsx"
    missing.parent.mkdir()
    wb.save(missing)
    probs = projects.validate_workbook(missing)
    assert "Sheet 'Defect List' is missing" in probs and any("UAT Plan" in x for x in probs)
    with pytest.raises(projects.ProjectError, match="Defect List"):
        projects.import_project(base, appdb, PERSON, path=missing)
    from conftest import TP_HEADER
    no_total = _wb(tmp_path / "bad2" / "NoTotal.xlsx", tp_header=[h if h != "Total TC" else "Totals" for h in TP_HEADER])
    with pytest.raises(projects.ProjectError, match="Total TC"):
        projects.import_project(base, appdb, PERSON, path=no_total)
    assert [x["name"] for x in appdb.projects()] == ["Gamma"]


# --------------------------------------------------------------------------- B/G: isolation, per-project settings

def test_projects_isolated_data_settings_and_roles(base, appdb, two, opened):
    a, b = two
    ca, ra = opened(a)
    cb, rb = opened(b)
    assert ra.count_items()[TP]["active"] == 5 and rb.count_items()[TP]["active"] == 2
    assert set(ra.load_items(TP)) != set(rb.load_items(TP)) or len(ra.load_items(TP)) != len(rb.load_items(TP))
    save_settings(ra, ADMIN, fail_pct=12.5, dry_run=False, pmo_email="pmo.shared@example.test", email_time="07:15")
    ca2, cb2 = apply_overrides(projects.project_config(base, a), ra), apply_overrides(projects.project_config(base, b), rb)
    assert ca2.alerts.fail_threshold == 0.125 and not ca2.alerts.dry_run
    assert cb2.alerts.fail_threshold == 0.05 and cb2.alerts.dry_run                   # B untouched
    assert rb.get_setting("settings.email_time") is None
    assert cb2.alerts.pmo_email == "pmo.shared@example.test"                            # shared (app.db)
    with ra.transaction():
        ra.update_user(PERSON.email, role="PMO")
    add_test_user(ra, "qa@example.test", "QA")
    assert ra.get_user(PERSON.email)["role"] == "PMO" and rb.get_user(PERSON.email)["role"] == "Admin"
    assert rb.get_user("qa@example.test") is None
    with ra.transaction():
        ra.upsert_holiday("2026-12-07", "Holiday", "manual", "t")                    # holidays: shared
    assert "2026-12-07" in rb.holidays(2026)
    assert projects.projects_for(base, appdb, "qa@example.test") == [appdb.get_project(a["id"])]
    assert projects.project_config(base, a).data_dir != projects.project_config(base, b).data_dir


def test_scheduler_one_task_pair_per_project():
    args = scheduler.install_args(Path("c.toml"), "07:15", "p002", 'A/B: "x"')
    assert args[args.index("/TN") + 1] == "Planning Progress - Daily Alerts - A-B- -x-"
    assert "--project p002 run-daily --scheduled" in args[-1] and args[args.index("/ST") + 1] == "07:15"
    w = scheduler.install_weekly_args(Path("c.toml"), project_id="p002", project_name="RAMC UAT")
    assert w[w.index("/TN") + 1] == "Planning Progress - Weekly Report - RAMC UAT" and "/ST" in w and "09:00" in w
    assert "--project p002 run-weekly" in w[-1]


# --------------------------------------------------------------------------- C: forms write to that project's file only

def _row_of(ws, cols, item_id):
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, cols["App ID"]).value == item_id:
            return r
    raise KeyError(item_id)


def _cols(ws):
    return {str(c.value).split("\n")[0].strip(): c.column for c in ws[1] if c.value}


def test_form_add_writes_only_to_selected_project_file(base, two, opened):
    a, b = two
    b_hash = sha(b["excel_path"])
    ca, ra = opened(a)
    tp_id = add_application(ca, ra, application="NewApp", report_by="QA X", report_by_email="x@example.test",
                            planned_start="2026-11-10", planned_finish="2026-11-20", total_tc=10, actor=ADMIN)
    t1 = add_task(ca, ra, group=None, new_group_title="New Phase", task_name="First new", responsible="QA X",
                  actor=ADMIN)
    letter = ra.get_item("UAT Plan", t1)["group_code"]
    assert letter == "F"                                         # after A, B, E
    t2 = add_task(ca, ra, group=letter, task_name="Second new", responsible="QA X", actor=ADMIN)
    svc = SyncService(ca, ra, opener=OpenpyxlWorkbook, sender=FakeSender())
    fr = svc.flush(force=True)
    assert fr.result == "written", fr.message
    fr2 = svc.flush(force=True)                                  # the 2nd task of the new group: next round
    assert fr2.result == "written", fr2.message
    assert ra.all_pending() == []
    wb = openpyxl.load_workbook(a["excel_path"])
    ws = wb["Test Progress"]
    c = _cols(ws)
    r = _row_of(ws, c, tp_id)
    assert ws.cell(r, c["Application"]).value == "NewApp"
    assert ws.cell(r, c["%Progress"]).value == f"=E{r}/D{r}"                    # approved Q14
    grand = next(i for i in range(2, ws.max_row + 1) if ws.cell(i, c["Application"]).value == "Grand Total")
    assert r < grand                                                            # above Grand Total
    ws = wb["UAT Plan"]
    c = _cols(ws)
    r1, r2 = _row_of(ws, c, t1), _row_of(ws, c, t2)
    assert ws.cell(r1 - 1, c["Group"]).value == "F" and ws.cell(r1 - 1, c["Task name"]).value == "New Phase"
    assert r2 == r1 + 1
    assert ra.get_item("UAT Plan", t1)["group_title"] == "New Phase"           # read back from Excel
    assert sha(b["excel_path"]) == b_hash                                       # other project's file untouched
    cb, rb = opened(b)
    assert not any(x.get("application") == "NewApp" for x in rb.load_items(TP).values())


def test_add_application_rules(base, two, opened):
    a, _ = two
    ca, ra = opened(a)
    from uat_tracker.services.edits import EditError
    from uat_tracker.services.auth import PermissionDenied
    with pytest.raises(EditError, match="required"):
        add_application(ca, ra, application="X", report_by="", planned_start="2026-11-10",
                        planned_finish="2026-11-20", actor=ADMIN)
    with pytest.raises(EditError, match="already"):
        add_application(ca, ra, application="nleads", report_by="Q", planned_start="2026-11-10",
                        planned_finish="2026-11-20", actor=ADMIN)
    with pytest.raises(EditError, match="before"):
        add_application(ca, ra, application="Y", report_by="Q", planned_start="2026-11-20",
                        planned_finish="2026-11-10", actor=ADMIN)
    qa = add_test_user(ra, "qa@example.test", "QA")
    with pytest.raises(PermissionDenied):
        add_application(ca, ra, application="Z", report_by="Q", planned_start="2026-11-10",
                        planned_finish="2026-11-20", actor=qa)


# --------------------------------------------------------------------------- E: RAMC UAT migration

def test_migration_ramc_uat_counts_lights_and_rename(tmp_path, monkeypatch):
    import uat_tracker.excel.writer as writer_mod
    from uat_tracker.services.migrate_multi import migrate
    monkeypatch.setattr(writer_mod, "excel_windows_for", lambda path: [])
    xl = _wb(tmp_path / "od" / "UAT Planning.xlsx")
    old_dir = tmp_path / "UATTracker"
    old_cfg = Config(excel_path=xl, data_dir=old_dir, window_start=WINDOW.start, window_end=WINDOW.end,
                     profile_file=old_dir / "profile.json")
    old_dir.mkdir()
    repo = SqliteRepository(old_cfg.db_path)
    run_import(old_cfg, repo, editor="t")
    add_test_user(repo, PERSON.email, "Admin", first="Pim", last="Pmo")
    with repo.transaction():
        repo.upsert_holiday("2026-12-07", "H", "manual", "t")
        repo.set_setting("settings.pmo_email", "pmo@example.test")
        repo.set_setting("settings.fail_threshold", "0.1")
        repo.set_setting("ui.lang", "th")
    (old_cfg.attachment_dir / "TP-1").mkdir(parents=True)
    (old_cfg.attachment_dir / "TP-1" / "r.pdf").write_bytes(b"pdf")
    repo.close()
    (old_dir / "profile.json").write_text(json.dumps({"email": PERSON.email}), encoding="utf-8")
    old_hash = sha(xl)
    base = Config(excel_path=xl, data_dir=tmp_path / "PlanningProgress", window_start=WINDOW.start,
                  window_end=WINDOW.end, profile_file=tmp_path / "PlanningProgress" / "profile.json")

    dry = migrate(base, old_dir, confirm=False)
    assert dry.new_excel.name == "Planning Progress RAMC UAT.xlsx" and xl.exists() and not dry.new_excel.exists()
    rep = migrate(base, old_dir, confirm=True, today=dt.date(2026, 11, 6))
    assert rep.done and rep.same and rep.before["uat_task"] > 0 and len(rep.lights_before) == 5
    assert not xl.exists() and rep.new_excel.exists() and sha(rep.new_excel) == old_hash
    assert all(Path(b).exists() for b in rep.backups)
    assert (old_dir / "tracker.db").exists()                                   # old data never deleted
    appdb = projects.open_app_db(base)
    try:
        p = appdb.get_project("p001")
        assert p["name"] == "RAMC UAT" and Path(p["excel_path"]) == rep.new_excel
        cfg, r = projects.open_project(base, p)
        try:
            assert cfg.alerts.fail_threshold == 0.1 and cfg.alerts.pmo_email == "pmo@example.test"
            assert "2026-12-07" in r.holidays(2026) and r.get_setting("ui.lang") == "th"
            assert r.get_user(PERSON.email)["role"] == "Admin"
            assert (cfg.attachment_dir / "TP-1" / "r.pdf").read_bytes() == b"pdf"
            assert r.verify_audit_chain()
        finally:
            r.close()
    finally:
        appdb.close()
    prof = read_profile_data(base.profile_path)
    assert prof["last_project"] == "p001" and prof["first_name"] == "Pim"
    with pytest.raises(projects.ProjectError, match="already has projects"):
        migrate(base, old_dir, confirm=True)


# --------------------------------------------------------------------------- I: Clear project data

def test_clear_project_only_that_project_and_restore(base, appdb, two, opened):
    a, b = two
    ca, ra = opened(a)
    cb, rb = opened(b)
    save_settings(ra, ADMIN, fail_pct=12, dry_run=True, pmo_email=None, email_time="07:00")
    att = ca.attachment_dir / "TP-1"
    att.mkdir(parents=True)
    (att / "result.pdf").write_bytes(b"pdf")
    with ra.transaction():
        ra.insert_attachment({"sheet": TP, "item_id": "TP-1", "original_name": "result.pdf",
                              "stored_path": str(att / "result.pdf"), "sha256": "x", "size_bytes": 3,
                              "uploaded_by": "t", "uploaded_at": "t"})
    before_a, before_b = counts(ra), counts(rb)
    a_hash = sha(a["excel_path"])
    qa = add_test_user(ra, "qa@example.test", "QA")
    with pytest.raises(ClearError, match="Not allowed"):
        clear_project(ca, ra, qa, typed_name="Alpha")
    with pytest.raises(ClearError, match="does not match"):
        clear_project(ca, ra, ADMIN, typed_name="alpha ")
    off = dataclasses.replace(ca, enable_clear_project=False)
    with pytest.raises(ClearError, match="turned off"):
        clear_project(off, ra, ADMIN, typed_name="Alpha")

    res = clear_project(ca, ra, ADMIN, typed_name="Alpha")
    assert res.backup_zip.exists()
    assert all(v == 0 for v in counts(ra).values())
    assert counts(rb) == before_b                                               # other project unchanged
    assert sha(a["excel_path"]) == a_hash                                       # Excel never touched
    assert ra.get_user(PERSON.email) is not None and ra.get_setting("settings.fail_threshold") == "0.12"
    assert not ca.attachment_dir.exists()
    with pytest.raises(sqlite3.IntegrityError):                                 # immutability is back
        with ra.transaction():
            ra.conn.execute("INSERT INTO audit_log (ts, editor, channel, hash) VALUES ('t','t','t','h')")
            ra.conn.execute("DELETE FROM audit_log")
    log = (base.root_dir / "clear-project.log").read_text(encoding="utf-8").splitlines()
    entry = json.loads(log[-1])
    assert entry["project_id"] == a["id"] and entry["user"] == ADMIN.editor and entry["action"] == "clear_project_data"
    # paused until "Import from Excel"
    assert ra.get_setting(S_PAUSED)
    with pytest.raises(ImportPaused):
        run_import(ca, ra, editor="t")
    svc = SyncService(ca, ra, opener=OpenpyxlWorkbook, sender=FakeSender())
    assert svc.pull(force=True) is None and svc.status()["state"] == "paused"
    assert svc.flush(force=True).result == "blocked"
    # restore from the backup zip
    ra.close()
    restore_backup(ca, res.backup_zip, editor="t")
    ra2 = SqliteRepository(ca.db_path, shared=ca.shared_db)
    try:
        assert counts(ra2) == before_a
        assert (ca.attachment_dir / "TP-1" / "result.pdf").read_bytes() == b"pdf"
        assert not ra2.get_setting(S_PAUSED)
    finally:
        ra2.close()


def test_import_from_excel_after_clear_unpauses(base, appdb, two, opened):
    a, _ = two
    ca, ra = opened(a)
    n = ra.count_items()[TP]["active"]
    clear_project(ca, ra, ADMIN, typed_name="Alpha")
    run_import(ca, ra, editor="t", force=True)
    assert not ra.get_setting(S_PAUSED) and ra.count_items()[TP]["active"] == n


# --------------------------------------------------------------------------- relink, archive, managers, profile

def test_relink_checks_app_ids(base, appdb, two, tmp_path, opened):
    a, b = two
    ca, ra = opened(a)
    assert SyncService(ca, ra, opener=OpenpyxlWorkbook, sender=FakeSender()).init_ids().result == "written"
    with pytest.raises(projects.ProjectError, match="App IDs"):
        projects.relink(base, appdb, PERSON, a["id"], _wb(tmp_path / "other" / "Other.xlsx"))   # no matching IDs
    moved = tmp_path / "moved" / "Alpha renamed.xlsx"
    moved.parent.mkdir()
    shutil.copy2(a["excel_path"], moved)
    os.remove(a["excel_path"])
    assert projects.file_missing(appdb.get_project(a["id"]))
    assert "App IDs found" in projects.relink(base, appdb, PERSON, a["id"], moved)
    assert Path(appdb.get_project(a["id"])["excel_path"]) == moved
    assert appdb.events(a["id"])[0]["action"] == "relinked"


def test_archive_hides_and_restores(base, appdb, two):
    a, b = two
    projects.set_archived(base, appdb, PERSON, b["id"], True)
    assert [p["id"] for p in projects.projects_for(base, appdb, PERSON.email)] == [a["id"]]
    assert appdb.get_project(b["id"])["archived"] == 1
    with pytest.raises(sqlite3.IntegrityError):
        with appdb.transaction():
            appdb.conn.execute("DELETE FROM project")                           # never deleted
    projects.set_archived(base, appdb, PERSON, b["id"], False)
    assert len(projects.projects_for(base, appdb, PERSON.email)) == 2


def test_looking_into_other_projects_keeps_email_project_name(base, appdb, two):
    """Bug 10-Nov-2026: email footer showed the LAST project in the list (e.g. RAMC UAT), not the open one."""
    from uat_tracker.alerts import templates
    a, b = two
    _cfg, repo = projects.open_project(base, a)
    repo.close()
    projects.projects_for(base, appdb, PERSON.email)          # loops over every project
    projects.can_manage(base, appdb, PERSON.email)
    projects.find_person(base, appdb, PERSON.email)
    assert templates.PROJECT["name"] == a["name"] != b["name"]
    assert f"({a['name']})" in templates.footer()


def test_new_email_joins_listing_projects_as_qa_and_cannot_manage(base, appdb, two, opened, tmp_path):
    a, b = two
    ca, ra = opened(a)
    edit_item(ca, ra, TP, "TP-38", {"report_by_email": "newqa@example.test"}, actor=ADMIN)
    ra.close()
    p = projects.identify_person(base, appdb, first_name="New", last_name="Qa", email="NewQA@example.test")
    assert p.email == "newqa@example.test"
    assert projects.role_in(base, a, p.email) == "QA" and projects.role_in(base, b, p.email) is None
    assert [x["id"] for x in projects.projects_for(base, appdb, p.email)] == [a["id"]]
    assert not projects.can_manage(base, appdb, p.email)
    (tmp_path / "q").mkdir()
    with pytest.raises(projects.ProjectError, match="Not allowed"):
        projects.create_project(base, appdb, p, name="Mine", folder=tmp_path / "q")
    known = projects.identify_person(base, appdb, first_name="Other", last_name="Name", email=PERSON.email)
    assert (known.first_name, known.last_name) == ("Pim", "Pmo")                # stored name wins
    assert read_profile_data(base.profile_path)["email"] == PERSON.email
    for proj in (a, b):
        cfg, r = projects.open_project(base, proj)
        with r.transaction():
            r.update_user(PERSON.email, active=0)
        r.close()
    with pytest.raises(ProfileError, match="deactivated"):
        projects.identify_person(base, appdb, first_name="", last_name="", email=PERSON.email)


# --------------------------------------------------------------------------- UI (headless)

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture(scope="module")
def qapp():
    pytest.importorskip("PySide6")
    from PySide6.QtWidgets import QApplication

    from uat_tracker.ui.widgets import QSS
    app = QApplication.instance() or QApplication([])
    app.setStyleSheet(QSS)
    return app


def test_profile_form_no_project_page_then_new_project(qapp, base, appdb, tmp_path, monkeypatch):
    from uat_tracker.ui import app as app_mod
    from uat_tracker.ui.profile import ProfileDialog
    from uat_tracker.ui.projects import NewProjectDialog, NoProjectWindow
    dlg = ProfileDialog(base, appdb)
    assert dlg.boot.isVisibleTo(dlg)                                             # no project yet
    dlg.email.setText("not an email")
    dlg.on_save()
    assert "valid email" in dlg.status.text() and dlg.person is None
    dlg.email.setText("Somchai.T@example.test")
    dlg.on_save()
    assert "required" in dlg.status.text()
    dlg.first.setText("Somchai")
    dlg.last.setText("T")
    dlg.on_save()
    assert dlg.person.editor == "Somchai T <somchai.t@example.test>" and dlg.result() == 1
    shown = []
    monkeypatch.setattr(app_mod.AppController, "_form", lambda self, **k: shown.append(1))
    ctl = app_mod.AppController(base, appdb, Path("config.toml"), start_sync=False)
    assert ctl.start() and shown == [] and isinstance(ctl.win, NoProjectWindow)
    # New project form: file name shown, Create enabled only with a folder and no duplicate
    nd = NewProjectDialog(base, appdb, ctl.person)
    nd.name.setText("Demo/1")
    assert "Planning Progress Demo-1.xlsx" in nd.file_lab.text() and not nd.ok.isEnabled()
    nd.folder = tmp_path / "nf"
    nd.folder.mkdir()
    nd._preview()
    assert nd.ok.isEnabled()
    monkeypatch.setattr("uat_tracker.ui.projects.info", lambda *a: None)
    nd._create()
    assert nd.project["name"] == "Demo/1"
    ctl.open_project(nd.project["id"])
    w = ctl.win
    assert "Demo/1" in w.windowTitle() and "Somchai T · Admin" in w.user_lab.text()
    assert w.proj_cb.count() == 1 and w.new_proj_btn.isVisibleTo(w)
    w.navigate("settings")
    assert "Danger zone" in " ".join(l.text() for l in w.findChildren(__import__("PySide6.QtWidgets",
                                                                                 fromlist=["QLabel"]).QLabel))
    w.on_closed = None
    ctl._close_current()


def test_switch_project_shows_only_that_projects_data(qapp, base, appdb, two):
    from uat_tracker.ui import app as app_mod
    a, b = two
    write_profile(base.profile_path, PERSON.email, first_name="Pim", last_name="Pmo", last_project=b["id"])
    ctl = app_mod.AppController(base, appdb, Path("config.toml"), start_sync=False)
    assert ctl.start()
    assert ctl.ctx.cfg.project_id == b["id"]                                      # last project remembered
    ids = {}
    for p in (a, b):
        ctl.open_project(p["id"])
        w = ctl.win
        assert w.proj_cb.count() == 2 and w.proj_cb.currentData() == p["id"] and p["name"] in w.windowTitle()
        for key in ("dashboard", "uat", "tp", "defects", "history", "conflicts"):
            w.navigate(key)
        ids[p["id"]] = {r["id"] for r in w.screens["tp"].rows}
        assert read_profile_data(base.profile_path)["last_project"] == p["id"]
    assert ids[a["id"]] != ids[b["id"]] and len(ids[b["id"]]) == 2
    _, ra = projects.open_project(base, a)
    try:
        assert ids[a["id"]] == {i for i, r in ra.load_items(TP).items() if r.get("active")}
    finally:
        ra.close()
    ctl.win.on_closed = None
    ctl._close_current()


def test_switch_user_logged_and_project_list_follows(qapp, base, appdb, two, monkeypatch):
    from uat_tracker.ui import app as app_mod
    from uat_tracker.ui import profile as prof_mod
    from uat_tracker.ui.projects import NoProjectWindow
    a, _ = two
    write_profile(base.profile_path, PERSON.email, first_name="Pim", last_name="Pmo", last_project=a["id"])
    ctl = app_mod.AppController(base, appdb, Path("config.toml"), start_sync=False)
    assert ctl.start()

    def fake_exec(dlg):
        dlg.email.setText("bee@example.test")
        dlg.first.setText("Bee")
        dlg.last.setText("User")
        dlg.on_save()
        return dlg.result()
    monkeypatch.setattr(prof_mod.ProfileDialog, "exec", fake_exec)
    assert ctl.switch_user()
    assert isinstance(ctl.win, NoProjectWindow)                                   # Bee is in no project
    _, ra = projects.open_project(base, a)
    try:
        e = [x for x in ra.audit_entries() if x["field"] == "switch_user"][-1]
        assert e["new_value"] == "Bee User <bee@example.test>" and e["old_value"] == "Pim Pmo <pmo@example.test>"
    finally:
        ra.close()
    ctl.win.on_closed = None
    ctl.win.close()
