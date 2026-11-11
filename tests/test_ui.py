"""Part 4 UI smoke tests (headless). The rules themselves are tested on the services."""

import dataclasses
import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtWidgets import QApplication, QLabel, QLineEdit  # noqa: E402

from uat_tracker.config import AlertsConfig  # noqa: E402
from uat_tracker.excel.writer import OpenpyxlWorkbook  # noqa: E402
from uat_tracker.mail.sender import FakeSender  # noqa: E402
from uat_tracker.services.importer import run_import  # noqa: E402
from uat_tracker.services.sync import SyncService  # noqa: E402
from uat_tracker.ui import i18n  # noqa: E402
from uat_tracker.ui.context import Ctx  # noqa: E402
from uat_tracker.ui.dialogs import UpdateDialog  # noqa: E402
from uat_tracker.ui.main_window import NAV, MainWindow  # noqa: E402
from uat_tracker.ui.widgets import QSS, DatePicker  # noqa: E402
from uat_tracker.ui.worker import wait_for_workers  # noqa: E402

import datetime as dt  # noqa: E402

from test_sync import excel_edit  # noqa: E402
from conftest import PMO_ACTOR  # noqa: E402


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance() or QApplication([])
    app.setStyleSheet(QSS)
    return app


@pytest.fixture
def win(qapp, cfg, repo):
    c = dataclasses.replace(cfg, alerts=AlertsConfig(pmo_email="pmo@example.test"))
    run_import(c, repo, editor="t")
    ctx = Ctx(c, repo)
    ctx.actor = PMO_ACTOR
    ctx.as_of = dt.date(2026, 11, 6)
    w = MainWindow(ctx, start_sync=False)
    yield w
    wait_for_workers()
    w.close()


def texts(widget):
    return " ".join(l.text() for l in widget.findChildren(QLabel))


def test_all_screens_reachable_and_language_switch(win):
    for key, tk in NAV:
        win.nav_buttons[key].click()
        assert win.stack.currentWidget() is win.screens[key]
    win.navigate("dashboard")
    assert "TC executed" in texts(win.screens["dashboard"])
    assert "178 / 238" in texts(win.screens["dashboard"])
    win.th_btn.click()
    assert i18n.get_lang() == "th"
    assert win.nav_buttons["dashboard"].text() == "แดชบอร์ด"
    assert "TC ที่ทดสอบแล้ว" in texts(win.screens["dashboard"])
    assert win.ctx.repo.get_setting("ui.lang") == "th"             # remembered
    for key, _ in NAV:
        win.navigate(key)                                          # every screen renders in Thai
    win.en_btn.click()
    assert win.nav_buttons["uat"].text() == "UAT Plan"


def test_date_picker_rejects_typing(qapp):
    dp = DatePicker("2026-11-06")
    assert dp.edit.isReadOnly()
    dp.edit.setFocus()
    from PySide6.QtTest import QTest
    QTest.keyClicks(dp.edit, "abc 1/1/2026")
    assert dp.value() == "2026-11-06" and dp.edit.text() == "06-Nov-2026"


def test_update_dialog_blocks_reschedule_without_reason(win):
    dlg = UpdateDialog(win.ctx, "Test Progress", "TP-38")
    dlg.inputs["planned_finish"].setValue("2026-11-20", emit=True)
    assert not dlg.save_btn.isEnabled() and "reason" in dlg.warn.text().lower()
    assert "reschedule #" in dlg.moved.text()
    dlg.reason.setText("Waiting data")
    assert dlg.save_btn.isEnabled()
    dlg.inputs["actual_finish"].setValue("2026-11-25", emit=True)  # after Planned Finish
    assert not dlg.save_btn.isEnabled()
    dlg.inputs["actual_finish"].setValue(None, emit=True)
    dlg.inputs["status"].setCurrentIndex(dlg.inputs["status"].findData("Complete"))
    assert not dlg.save_btn.isEnabled()                             # Complete needs Actual Finish
    dlg.inputs["status"].setCurrentIndex(dlg.inputs["status"].findData(None))
    dlg.on_save()
    assert win.ctx.repo.get_item("Test Progress", "TP-38")["planned_finish"] == "2026-11-20"
    assert any(e["field"] == "planned_finish" for e in win.ctx.repo.audit_entries("TP-38"))


def test_excel_edit_refreshes_open_screen(win):
    svc = SyncService(win.ctx.cfg, win.ctx.repo, opener=OpenpyxlWorkbook, sender=FakeSender())
    svc.init_ids()
    win.navigate("tp")
    win._tick()
    excel_edit(win.ctx.cfg.excel_path, "Test Progress", "TP-38", "#Pass", 125)
    svc.pull()                                                     # what the sync thread does
    before = win.screens["tp"].rows
    win._tick()                                                    # timer: audit grew -> refresh
    after = next(r for r in win.screens["tp"].rows if r["id"] == "TP-38")
    assert after["pass"] == 125 and before is not win.screens["tp"].rows


def test_defect_board_move_logs_status(win):
    win.navigate("defects")
    win.screens["defects"].move("DEF-21", "Fixed")
    assert win.ctx.repo.get_item("Defect List", "DEF-21")["status"] == "Fixed"
    e = [a for a in win.ctx.repo.audit_entries("DEF-21") if a["field"] == "status"][-1]
    assert e["channel"] == "app" and e["new_value"] == "Fixed"


def test_search_filters_current_screen(win):
    win.navigate("tp")
    win.search.setText("DMS")
    win._search()
    assert [r["app"] for r in win.screens["tp"].rows] == ["DMS-RAMC"]


def test_conflict_center_resolve_and_review(win, monkeypatch):
    from uat_tracker.services.edits import edit_item
    from uat_tracker.ui.screens import conflicts as conf_mod
    svc = SyncService(win.ctx.cfg, win.ctx.repo, opener=OpenpyxlWorkbook, sender=FakeSender())
    svc.init_ids()
    edit_item(win.ctx.cfg, win.ctx.repo, "Test Progress", "TP-38", {"remark": "app note"}, actor=PMO_ACTOR)
    excel_edit(win.ctx.cfg.excel_path, "Test Progress", "TP-38", "Remark", "excel note")
    svc.pull()
    win.navigate("conflicts")
    scr = win.screens["conflicts"]
    assert "excel note" in texts(scr) and "app note" in texts(scr)
    monkeypatch.setattr(conf_mod, "error", lambda *a: None)
    c = win.ctx.repo.open_conflicts()[0]
    scr._resolve(c, "app", "")                                   # no reason -> nothing happens
    assert win.ctx.repo.open_conflicts()
    # the re-read runs in a worker thread with its own DB connection
    monkeypatch.setattr(conf_mod, "SyncService",
                        lambda cfg, repo: SyncService(cfg, repo, opener=OpenpyxlWorkbook, sender=FakeSender()))
    errors = []
    monkeypatch.setattr(conf_mod, "error", lambda *a: errors.append(a[-1]))
    scr._resolve(c, "app", "app is right")
    wait_for_workers()
    assert errors == [] and win.ctx.repo.open_conflicts() == []
    # review queue: one-click suggestion
    monkeypatch.setattr(conf_mod, "info", lambda *a: None)
    scr._pick("rev")
    r = next(r for r in win.ctx.repo.open_reviews() if r["item_id"] == "TP-12" and r["token"] == "24/09/206")
    scr._apply(r, "2026-09-24")
    assert win.ctx.repo.get_item("Test Progress", "TP-12")["planned_start"] == "2026-09-24"
    win.th_btn.click()
    assert "คิวตรวจของ PMO" in " ".join(b.text() for b in scr.findChildren(type(win.th_btn)))
    win.en_btn.click()


def test_test_progress_upload_tc_result_column(win, tmp_path, monkeypatch):
    from uat_tracker.ui.screens import test_progress as tp_mod
    win.navigate("tp")
    scr = win.screens["tp"]
    tb = scr.table
    assert tb.horizontalHeaderItem(8).text() == "Upload TC/Result"
    rows = {r["id"]: i for i, r in enumerate(scr.rows)}
    assert "Complete" in texts(tb.cellWidget(rows["TP-1"], 8))           # Nleads: Complete in Excel
    assert "Not uploaded" in texts(tb.cellWidget(rows["TP-38"], 8))      # DMS-RAMC: not yet
    f = tmp_path / "result.pdf"
    f.write_bytes(b"%PDF")
    monkeypatch.setattr(tp_mod.QFileDialog, "getOpenFileName", lambda *a, **k: (str(f), ""))
    monkeypatch.setattr(tp_mod, "info", lambda *a: None)
    scr._upload_for("TP-38")                                              # the cell's Upload button
    rows = {r["id"]: i for i, r in enumerate(scr.rows)}
    assert "Complete" in texts(tb.cellWidget(rows["TP-38"], 8))
    win.th_btn.click()
    assert tb.horizontalHeaderItem(8).text() == "Upload TC/Result"
    win.en_btn.click()


def test_uat_update_dialog_edits_every_field(win):
    svc = SyncService(win.ctx.cfg, win.ctx.repo, opener=OpenpyxlWorkbook, sender=FakeSender())
    svc.init_ids()
    dlg = UpdateDialog(win.ctx, "UAT Plan", "B-7")
    for f in ("task_no", "task_name", "responsible", "status", "start_date", "start_time", "finish_date",
              "finish_time", "note", "sub_task", "file_path", "remark"):
        assert f in dlg.inputs, f
    dlg.inputs["task_name"].setPlainText("")
    assert not dlg.save_btn.isEnabled() and "Task name" in dlg.warn.text()
    dlg.inputs["task_name"].setPlainText("Renamed task")
    dlg.inputs["responsible"].setPlainText("QA New Owner")
    dlg.inputs["note"].setPlainText("note line 1\nline 2")
    dlg.inputs["file_path"].setText(r"\share\uat\b7.xlsx")
    dlg.inputs["start_time"].setText("09:00")
    dlg.inputs["sub_task"].setText("7.1")
    assert dlg.save_btn.isEnabled()
    dlg.on_save()
    rec = win.ctx.repo.get_item("UAT Plan", "B-7")
    assert rec["task_name"] == "Renamed task" and rec["responsible"] == "QA New Owner"
    assert rec["note"] == "note line 1\nline 2" and rec["start_time"] == "09:00"
    svc.flush(force=True)
    from test_sync import cell
    assert cell(win.ctx.cfg.excel_path, "UAT Plan", "B-7", "Task name") == "Renamed task"
    assert cell(win.ctx.cfg.excel_path, "UAT Plan", "B-7", "File Path") == r"\share\uat\b7.xlsx"
    assert win.ctx.repo.all_pending() == []


# --------------------------------------------------------------------------- Part 5 v2: profile + roles in the UI

# Profile form + switch user: see test_projects.py (multi-project versions).


def test_qa_sees_read_only_rows_and_no_settings(win):
    from conftest import add_test_user
    from uat_tracker.services.edits import edit_item
    repo = win.ctx.repo
    edit_item(win.ctx.cfg, repo, "Test Progress", "TP-38", {"report_by_email": "qa@example.test"}, actor=PMO_ACTOR)
    win.ctx.actor = add_test_user(repo, "qa@example.test", "QA", aliases=["QA Somsak"])
    win.refresh_current()
    assert "Qa Test · QA" in win.user_lab.text()
    assert not win.nav_buttons["settings"].isVisibleTo(win)
    assert win.nav_buttons["conflicts"].isVisibleTo(win)
    own = UpdateDialog(win.ctx, "Test Progress", "TP-38")
    assert own.save_btn.isEnabled() and not own.read_only
    other = UpdateDialog(win.ctx, "Test Progress", "TP-6")
    assert other.read_only and not other.save_btn.isEnabled() and "Read only" in other.warn.text()
    other.inputs["remark"].setPlainText("try")                  # typing does not re-enable Save
    other.reason.setText("x")
    assert not other.save_btn.isEnabled()
    assert UpdateDialog(win.ctx, "UAT Plan", "E-11").save_btn.isEnabled()      # QA Somsak's task
    assert not UpdateDialog(win.ctx, "UAT Plan", "B-6").save_btn.isEnabled()
    win.navigate("uat")
    assert not win.screens["uat"].add_btn.isVisibleTo(win)
    win.navigate("tp")
    scr = win.screens["tp"]
    rows = {r["id"]: i for i, r in enumerate(scr.rows)}
    from PySide6.QtWidgets import QPushButton
    b6 = scr.table.cellWidget(rows["TP-6"], 8).findChildren(QPushButton)
    assert not b6 or not b6[0].isEnabled()
    b38 = scr.table.cellWidget(rows["TP-38"], 8).findChildren(QPushButton)
    assert b38 and b38[0].isEnabled()
    win.navigate("history")
    assert "Restore" not in " ".join(b.text() for b in win.screens["history"].findChildren(QPushButton))


def test_only_admin_sees_user_management(win, monkeypatch):
    from conftest import add_test_user
    from uat_tracker.ui import users as users_mod
    win.navigate("settings")
    assert not hasattr(win.screens["settings"], "users_table") or not win.screens["settings"].users_table.isVisible()
    assert "Users and roles" not in texts(win.screens["settings"])
    win.ctx.actor = add_test_user(win.ctx.repo, "admin@example.test", "Admin")
    win.navigate("settings")
    scr = win.screens["settings"]
    assert "Users and roles" in texts(scr)
    emails = [scr.users_table.item(i, 2).text() for i in range(scr.users_table.rowCount())]
    assert set(emails) == {"pmo@example.test", "admin@example.test"}
    dlg = users_mod.UserDialog(win.ctx, None)
    dlg.email.setText("new@example.test")
    dlg.first.setText("New")
    dlg.last.setText("Person")
    dlg.role.setCurrentIndex(dlg.role.findData("User BU/PM"))
    for i in range(dlg.names.count()):
        if dlg.names.item(i).text() == "QA Anucha":
            dlg.names.item(i).setCheckState(Qt.Checked)
    dlg.other.setText("K. Somchai")
    dlg.on_save()
    u = win.ctx.repo.get_user("new@example.test")
    assert u["role"] == "User BU/PM" and u["display_name"] == "New Person"
    assert {a["name"] for a in win.ctx.repo.user_aliases("new@example.test")} == {"QA Anucha", "K. Somchai"}


# --------------------------------------------------------------------------- Part 6: weekly report UI

def test_weekly_report_button_preview_send_and_history(win, monkeypatch):
    import datetime as _dt
    from conftest import add_test_user
    from test_weekly import MemKeys, fake_ai_ok
    from uat_tracker.services import weekly_ai, weekly_report as W
    from uat_tracker.ui import weekly as wk
    monkeypatch.setattr(weekly_ai, "KEYS", MemKeys())
    monkeypatch.setattr(wk, "OutlookSender", lambda acct: FakeSender())
    monkeypatch.setattr(wk, "info", lambda *a: None)
    win.navigate("dashboard")
    assert win.screens["dashboard"].weekly_btn.isVisibleTo(win)                 # PMO
    # the background worker (run in this thread here): no API key -> report without AI
    w = wk.GenerateWorker(win.ctx.cfg, win.ctx.actor, _dt.date(2026, 11, 16))
    got = []
    w.done.connect(got.append)
    w.failed.connect(lambda m: got.append(m))
    w.run()
    assert isinstance(got[0], int)
    r = win.ctx.repo.get_weekly(got[0])
    assert r["status"] == "preview" and r["ai_used"] == 0 and "No Claude API key" in r["ai_error"]
    dlg = wk.WeeklyPreviewDialog(win.ctx, got[0])
    assert "DRY RUN" in texts(dlg) and "178 / 238" in dlg.view.toPlainText()
    dlg.on_send()
    assert win.ctx.repo.get_weekly(got[0])["status"] == "dry_run"
    g = W.generate(win.ctx.cfg, win.ctx.repo, actor=win.ctx.actor, as_of=_dt.date(2026, 11, 16), ai_func=fake_ai_ok)
    d2 = wk.WeeklyPreviewDialog(win.ctx, g.report_id)
    d2.reject()                                                                  # Cancel / window X
    assert win.ctx.repo.get_weekly(g.report_id)["status"] == "cancelled"
    win.navigate("weekly")
    scr = win.screens["weekly"]
    assert len(scr.rows) == 2 and "Executive summary" in scr.view.toPlainText()  # newest (AI) shown first
    # QA: no button on the dashboard or the reports screen
    win.ctx.actor = add_test_user(win.ctx.repo, "qa@example.test", "QA")
    win.navigate("dashboard")
    assert not win.screens["dashboard"].weekly_btn.isVisibleTo(win)
    win.navigate("weekly")
    assert not scr.gen_btn.isVisibleTo(win) and len(scr.rows) == 2                # can still read past reports


def test_tp_table_shows_planned_and_actual_start_finish(win):
    win.navigate("tp")
    tb = win.screens["tp"].table
    heads = [tb.horizontalHeaderItem(c).text() for c in range(tb.columnCount())]
    assert heads[4:8] == ["Planned Start", "Planned Finish", "Actual Start", "Actual Finish"]
    rows = {r["id"]: i for i, r in enumerate(win.screens["tp"].rows)}
    assert tb.item(rows["TP-6"], 4).text() == "21-Sep-2026" and tb.item(rows["TP-6"], 5).text() == "08-Oct-2026"


def test_update_dialog_scrolls_with_buttons_pinned(win):
    dlg = UpdateDialog(win.ctx, "UAT Plan", "B-7")
    dlg.show()
    try:
        assert dlg._scroll is not None and dlg._scroll.isAncestorOf(dlg.inputs["task_name"])
        assert not dlg._scroll.isAncestorOf(dlg.save_btn)              # Save/Cancel stay pinned
        cap = dlg.screen().availableGeometry().height() * 0.85
        assert dlg.height() <= cap + 1
    finally:
        dlg.close()


def test_test_progress_columns_not_hidden(win, qapp):
    """Every column fits its header + cells; spare width goes to Application, narrow window -> scrollbar."""
    win.resize(1920, 1080)
    win.show()
    win.stack.setCurrentWidget(win.screens["tp"])
    scr = win.screens["tp"]
    scr.refresh()
    qapp.processEvents()
    tb, hh = scr.table, scr.table.horizontalHeader()
    for col in range(tb.columnCount()):
        need = max(hh.sectionSizeFromContents(col).width(), tb.sizeHintForColumn(col))
        assert tb.columnWidth(col) >= need, (col, tb.columnWidth(col), need)
    assert tb.columnWidth(0) >= 220
    win.resize(900, 700)
    qapp.processEvents()
    qapp.processEvents()
    assert tb.columnWidth(0) == scr._col_min[0]


def test_settings_red_warning_only_when_clear_feature_on_in_live_mode(win):
    scr = win.screens["settings"]

    def banner():
        win.navigate("settings")
        scr.refresh()
        return [l for l in scr.findChildren(QLabel) if l.objectName() == "liveClearWarning"]
    assert banner() == []                                                         # default: feature off
    base = win.ctx.cfg
    win.ctx.cfg = dataclasses.replace(base, enable_clear_project=True)            # on, but dry-run
    assert banner() == []
    win.ctx.cfg = dataclasses.replace(base, enable_clear_project=True,
                                      alerts=dataclasses.replace(base.alerts, dry_run=False))
    shown = banner()
    assert len(shown) == 1 and "enable_clear_project = false" in shown[0].text()
