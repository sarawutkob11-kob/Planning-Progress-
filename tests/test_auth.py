"""Part 5 v2: user profile (no login / OTP) and roles. Rules are tested on the services."""

from __future__ import annotations

import pytest

from conftest import PMO_ACTOR, add_test_user
from uat_tracker.mail.sender import FakeSender
from uat_tracker.services import auth, views
from uat_tracker.services.attachments import upload
from uat_tracker.services.auth import PermissionDenied
from uat_tracker.services.edits import edit_item
from uat_tracker.services.importer import run_import
from uat_tracker.services.items import add_defect, add_task
from uat_tracker.services.profile import (FirstAdminNeeded, ProfileError, identify, open_profile, read_profile,
                                          write_profile)
from uat_tracker.services.reminders import send_reminder
from uat_tracker.services.resolve import resolve_review
from uat_tracker.services.settings import save_settings
from uat_tracker.services.users import UserError, add_user, import_candidates, update_user
from uat_tracker.storage.sqlite_repo import SqliteRepository

TP, UAT, DEF = "Test Progress", "UAT Plan", "Defect List"


@pytest.fixture
def data(cfg, repo):
    run_import(cfg, repo, editor="t")
    return repo


@pytest.fixture
def empty_repo(cfg):
    cfg.data_dir.mkdir(parents=True, exist_ok=True)
    r = SqliteRepository(cfg.data_dir / "empty.db")
    yield r
    r.close()


def ident(cfg, repo, first, last, email, **kw):
    return identify(cfg, repo, first_name=first, last_name=last, email=email, **kw)


# --------------------------------------------------------------------------- 1-3: profile, first Admin

def test_first_start_form_then_no_form(cfg, empty_repo):
    repo = empty_repo
    assert open_profile(cfg, repo) is None                            # no profile yet -> the app shows the form
    with pytest.raises(FirstAdminNeeded):
        ident(cfg, repo, "Somchai", "T", "Somchai.T@example.test")       # the UI must confirm first
    a = ident(cfg, repo, "Somchai", "T", "Somchai.T@example.test", confirm_admin=True)
    assert a.role == "Admin" and a.editor == "Somchai T <somchai.t@example.test>"
    assert read_profile(cfg.profile_path) == "somchai.t@example.test"
    b = open_profile(cfg, repo)                                       # next start: straight in
    assert b.email == "somchai.t@example.test" and b.role == "Admin"
    assert repo.get_user("somchai.t@example.test")["last_used_at"]
    assert [a["name"] for a in repo.user_aliases("somchai.t@example.test")] == ["Somchai T"]
    assert [e["field"] for e in repo.audit_entries()] == ["profile_created"]


def test_invalid_email_or_missing_names_cannot_save(cfg, repo):
    for bad in ("", "not-an-email", "a@b", "two@a.com; x@b.com", "Name <a@b.com>"):
        with pytest.raises(ProfileError, match="valid email"):
            ident(cfg, repo, "A", "B", bad)
    with pytest.raises(ProfileError, match="required"):
        ident(cfg, repo, "", "B", "new@example.test")
    assert read_profile(cfg.profile_path) is None


def test_first_admin_once_only_then_new_emails_are_qa(cfg, empty_repo):
    repo = empty_repo
    ident(cfg, repo, "First", "Admin", "first@example.test", confirm_admin=True)
    b = ident(cfg, repo, "Second", "Person", "second@example.test", confirm_admin=True)   # confirm ignored now
    assert b.role == "QA"
    assert [u["role"] for u in repo.list_users() if u["email"] == "first@example.test"] == ["Admin"]
    assert repo.count_active_admins() == 1


def test_existing_email_keeps_stored_name_and_deactivated_refused(cfg, repo):
    a = ident(cfg, repo, "Someone", "Else", "PMO@example.test")
    assert a.role == "PMO" and a.display_name == "Pim Pmo"            # stored name wins
    admin = add_test_user(repo, "admin@example.test", "Admin")
    update_user(repo, admin, "pmo@example.test", active=False)
    with pytest.raises(ProfileError, match="deactivated"):
        ident(cfg, repo, "Pim", "Pmo", "pmo@example.test")
    assert open_profile(cfg, repo) is None                            # saved profile is deactivated -> form
    assert repo.audit_entries()[-1]["field"] == "denied"


def test_switch_user_changes_editor_in_new_entries(cfg, data):
    repo = data
    a = ident(cfg, repo, "Pim", "Pmo", "pmo@example.test")
    edit_item(cfg, repo, TP, "TP-38", {"remark": "by pim"}, actor=a)
    b = ident(cfg, repo, "Bee", "User", "bee@example.test", previous=a)
    assert read_profile(cfg.profile_path) == "bee@example.test"
    sw = [e for e in repo.audit_entries() if e["field"] == "switch_user"][-1]
    assert sw["old_value"] == "Pim Pmo <pmo@example.test>" and sw["new_value"] == "Bee User <bee@example.test>"
    add_defect(cfg, repo, system="CCMS", detail="x", assigned_to="Dev", fix_target_date="2026-11-20", actor=b)
    eds = [e["editor"] for e in repo.audit_entries() if e["channel"] == "app"]
    assert eds[-2:] == ["Pim Pmo <pmo@example.test>", "Bee User <bee@example.test>"]


def test_profile_file_corrupt_or_missing(cfg, repo):
    cfg.profile_path.parent.mkdir(parents=True, exist_ok=True)
    cfg.profile_path.write_text("{not json", encoding="utf-8")
    assert read_profile(cfg.profile_path) is None and open_profile(cfg, repo) is None
    write_profile(cfg.profile_path, " PMO@example.test ")
    assert open_profile(cfg, repo).email == "pmo@example.test"


# --------------------------------------------------------------------------- 5-6: roles in the service layer

def test_qa_edits_only_own_rows(cfg, data):
    repo = data
    qa = add_test_user(repo, "somsak@example.test", "QA", aliases=["QA Somsak"])
    edit_item(cfg, repo, TP, "TP-38", {"report_by_email": "somsak@example.test"}, actor=PMO_ACTOR)
    # Test Progress: own row (Report By Email) yes, other rows no
    edit_item(cfg, repo, TP, "TP-38", {"remark": "mine"}, actor=qa)
    with pytest.raises(PermissionDenied, match="not your row"):
        edit_item(cfg, repo, TP, "TP-6", {"remark": "not mine"}, actor=qa)
    assert repo.get_item(TP, "TP-6")["remark"] is None
    # UAT Plan: own by name in Excel ("QA Somsak" on E-11)
    edit_item(cfg, repo, UAT, "E-11", {"remark": "mine"}, actor=qa)
    with pytest.raises(PermissionDenied):
        edit_item(cfg, repo, UAT, "B-6", {"remark": "no"}, actor=qa)
    # Defect: Assigned To
    with pytest.raises(PermissionDenied):
        edit_item(cfg, repo, DEF, "DEF-21", {"status": "Fixed"}, actor=qa)
    update_user(repo, add_test_user(repo, "admin@example.test", "Admin"), qa.email,
                aliases=["QA Somsak", "IT-DMS-RAMC"])
    edit_item(cfg, repo, DEF, "DEF-21", {"status": "Fixed"}, actor=qa)
    # denial is logged with the real email
    denied = [e for e in repo.audit_entries() if e["field"] == "denied"]
    assert len(denied) == 3 and {e["editor"] for e in denied} == {"Somsak Test <somsak@example.test>"}
    assert {e["item_id"] for e in denied} == {"TP-6", "B-6", "DEF-21"}


def test_multi_name_cell_counts_for_each_person(cfg, data):
    repo = data
    edit_item(cfg, repo, UAT, "B-6", {"responsible": "QA WANNA\nMALEE SRISUKSAWAT / QA Nida"},
              actor=PMO_ACTOR)
    k = add_test_user(repo, "malee@example.test", "QA", aliases=["malee  srisuksawat"])
    edit_item(cfg, repo, UAT, "B-6", {"remark": "ok"}, actor=k)
    r = add_test_user(repo, "nida@example.test", "QA", aliases=["QA Nida"])
    edit_item(cfg, repo, UAT, "B-6", {"remark": "ok too"}, actor=r)


def test_qa_other_rights(cfg, data, tmp_path):
    repo = data
    qa = add_test_user(repo, "qa@example.test", "QA")
    did = add_defect(cfg, repo, system="CCMS", detail="Wrong total", assigned_to="Dev A",
                     fix_target_date="2026-11-20", actor=qa)                    # QA may add defects
    assert repo.get_item(DEF, did)
    with pytest.raises(PermissionDenied):
        edit_item(cfg, repo, DEF, did, {"remark": "x"}, actor=qa)               # created by QA, assigned to Dev A
    with pytest.raises(PermissionDenied):
        add_task(cfg, repo, group="B", task_name="New", responsible="QA x", actor=qa)
    f = tmp_path / "r.pdf"
    f.write_bytes(b"%PDF")
    with pytest.raises(PermissionDenied):
        upload(cfg, repo, "TP-38", f, actor=qa)
    edit_item(cfg, repo, TP, "TP-38", {"report_by_email": "qa@example.test"}, actor=PMO_ACTOR)
    assert upload(cfg, repo, "TP-38", f, actor=qa).status_changed
    for call in (lambda: send_reminder(cfg, repo, FakeSender(), "B-6", to="o@example.test", actor=qa),
                 lambda: resolve_review(cfg, repo, repo.open_reviews()[0]["id"], None, actor=qa),
                 lambda: views.restore(cfg, repo, repo.audit_entries()[-1]["id"], actor=qa),
                 lambda: save_settings(repo, qa, fail_pct=5, dry_run=True, pmo_email="", email_time="08:30"),
                 lambda: add_user(repo, qa, "x@example.test", role="QA")):
        with pytest.raises(PermissionDenied):
            call()


def test_user_bu_pm_has_pmo_rights_but_not_users(cfg, data, tmp_path):
    repo = data
    bu = add_test_user(repo, "bu@example.test", "User BU/PM")
    for action in (auth.EDIT, auth.ADD_TASK, auth.ADD_DEFECT, auth.UPLOAD, auth.RESOLVE, auth.RESTORE,
                   auth.REMIND, auth.SETTINGS):
        assert auth.can(repo, bu, action, TP, repo.get_item(TP, "TP-6"))
        assert auth.can(repo, bu, action, TP, repo.get_item(TP, "TP-6")) == \
            auth.can(repo, PMO_ACTOR, action, TP, repo.get_item(TP, "TP-6"))
    edit_item(cfg, repo, TP, "TP-6", {"remark": "bu"}, actor=bu)
    add_task(cfg, repo, group="B", task_name="BU task", responsible="Someone", actor=bu)
    resolve_review(cfg, repo, repo.open_reviews()[0]["id"], None, actor=bu, reason="drop")
    save_settings(repo, bu, fail_pct=5, dry_run=True, pmo_email="", email_time="08:30")
    for who in (bu, PMO_ACTOR):
        assert not auth.can(repo, who, auth.USERS)
        with pytest.raises(PermissionDenied):
            add_user(repo, who, "x@example.test", role="QA")


def test_role_read_from_db_every_time(cfg, data):
    repo = data
    admin = add_test_user(repo, "admin@example.test", "Admin")
    qa = add_test_user(repo, "q@example.test", "QA")
    with pytest.raises(PermissionDenied):
        edit_item(cfg, repo, TP, "TP-6", {"remark": "a"}, actor=qa)
    update_user(repo, admin, qa.email, role="PMO")
    edit_item(cfg, repo, TP, "TP-6", {"remark": "a"}, actor=qa)                 # same Actor object, new role
    update_user(repo, admin, qa.email, active=False)
    with pytest.raises(PermissionDenied):
        edit_item(cfg, repo, TP, "TP-6", {"remark": "b"}, actor=qa)


# --------------------------------------------------------------------------- 7: user management

def test_user_management_rules_and_audit(cfg, data):
    repo = data
    admin = add_test_user(repo, "admin@example.test", "Admin")
    add_user(repo, admin, " New.User@example.test ", role="QA", first_name="New", last_name="User",
             aliases=["QA New", "qa  new"])
    assert repo.get_user("new.user@example.test")["display_name"] == "New User"
    assert repo.get_user("new.user@example.test")["role"] == "QA"
    assert [a["name"] for a in repo.user_aliases("new.user@example.test")] == ["QA New"]
    with pytest.raises(UserError, match="already exists"):
        add_user(repo, admin, "new.user@example.test", role="QA")
    with pytest.raises(UserError, match="Role"):
        add_user(repo, admin, "z@example.test", role="Boss")
    assert update_user(repo, admin, "new.user@example.test", role="PMO") == ["role"]
    update_user(repo, admin, "new.user@example.test", active=False)
    with pytest.raises(UserError, match="last active Admin"):
        update_user(repo, admin, "admin@example.test", role="PMO")
    with pytest.raises(UserError, match="last active Admin"):
        update_user(repo, admin, "admin@example.test", active=False)
    with pytest.raises(Exception, match="never deleted"):
        repo.conn.execute("DELETE FROM app_user WHERE email='new.user@example.test'")
    fields = [(e["field"], e["editor"]) for e in repo.audit_entries() if e["channel"] == "auth"]
    ed = "Admin Test <admin@example.test>"
    assert fields == [("user_added", ed), ("role", ed), ("active", ed)]
    assert repo.verify_audit_chain()


def test_import_candidates_from_test_progress(cfg, data):
    repo = data
    edit_item(cfg, repo, TP, "TP-38", {"report_by_email": "somsak@example.test",
                                       "cc_emails": "pm@example.test; pmo@example.test"}, actor=PMO_ACTOR)
    c = {x["email"]: x for x in import_candidates(repo)}
    assert set(c) == {"somsak@example.test", "pm@example.test"}               # pmo@ is already a user
    assert c["somsak@example.test"]["role"] == "QA" and c["somsak@example.test"]["names"] == ["QA Somsak"]
    assert c["pm@example.test"]["sources"] == ["CC Email"] and c["pm@example.test"]["apps"] == ["DMS-RAMC"]
    assert repo.get_user("somsak@example.test") is None                     # nothing saved by listing


# --------------------------------------------------------------------------- 7: "First Last <email>" in Edit History

def test_edit_history_shows_first_last_email(cfg, data):
    repo = data
    a = ident(cfg, repo, "x", "y", "pmo@example.test")
    edit_item(cfg, repo, TP, "TP-38", {"remark": "by profile"}, actor=a)
    rows = views.history(repo, item_id="TP-38")
    assert rows[0]["editor"] == "Pim Pmo <pmo@example.test>" and rows[0]["type"] == "Data edit"
    assert repo.pending(TP)[("TP-38", "remark")]["editor"] == "Pim Pmo <pmo@example.test>"
    send_reminder(cfg, repo, FakeSender(), "B-6", to="o@example.test", actor=a)           # email log entry too
    assert repo.audit_entries()[-1]["editor"] == "Pim Pmo <pmo@example.test>"


# --------------------------------------------------------------------------- CLI

def test_cli_uses_the_pc_profile(cfg, repo):
    import argparse
    from uat_tracker import cli_auth
    ns = lambda **k: argparse.Namespace(**k)
    assert cli_auth.needs_login(ns(cmd="edit"))
    assert cli_auth.needs_login(ns(cmd="conflicts", ccmd="resolve"))
    assert not cli_auth.needs_login(ns(cmd="conflicts", ccmd="list"))
    assert not cli_auth.needs_login(ns(cmd="run-daily"))
    assert not cli_auth.needs_login(ns(cmd="sync", scmd="once"))
    out = []
    assert cli_auth.cli_identity(cfg, repo, out=out.append) is None and "Open the app once" in out[-1]
    write_profile(cfg.profile_path, "pmo@example.test")
    a = cli_auth.cli_identity(cfg, repo, out=out.append)
    assert a.editor == "Pim Pmo <pmo@example.test>" and "Pim Pmo" in out[-1]
