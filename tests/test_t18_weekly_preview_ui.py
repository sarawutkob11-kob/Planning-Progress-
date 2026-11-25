"""QA T18 (Qt): the preview dialog shows the invalid-address line (PMO matter, not in the email)."""

import dataclasses
import datetime as dt
import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PySide6")

from PySide6.QtWidgets import QApplication, QLabel  # noqa: E402

from conftest import PMO_ACTOR  # noqa: E402
from uat_tracker.config import AlertsConfig  # noqa: E402
from uat_tracker.excel.sheets import TEST_PROGRESS  # noqa: E402
from uat_tracker.services import weekly_ai, weekly_report as W  # noqa: E402
from uat_tracker.services.importer import run_import  # noqa: E402
from uat_tracker.ui import i18n  # noqa: E402
from uat_tracker.ui.context import Ctx  # noqa: E402
from uat_tracker.ui.widgets import QSS  # noqa: E402

MON = dt.date(2026, 11, 16)


class _Keys:
    def get(self):
        return None


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance() or QApplication([])
    app.setStyleSheet(QSS)
    return app


@pytest.fixture
def ctx(qapp, cfg, repo, monkeypatch):
    monkeypatch.setattr(weekly_ai, "KEYS", _Keys())
    c = dataclasses.replace(cfg, alerts=AlertsConfig(pmo_email="pmo@example.test"))
    run_import(c, repo, editor="t")
    x = Ctx(c, repo)
    x.actor = PMO_ACTOR
    x.as_of = MON
    return x


def _texts(w):
    return [l.text() for l in w.findChildren(QLabel)]


def test_preview_dialog_shows_invalid_line(ctx):
    from uat_tracker.ui import weekly as wk
    with ctx.repo.transaction():
        ctx.repo.set_field(TEST_PROGRESS, "TP-12", "report_by_email", "bad-address", "2026-11-15T00:00:00+07:00")
    g = W.generate(ctx.cfg, ctx.repo, actor=PMO_ACTOR, as_of=MON, use_ai=False)
    dlg = wk.WeeklyPreviewDialog(ctx, g.report_id)
    try:
        want = i18n.t("wk.invalid", n=1, w="COLLINS (bad-address)")
        assert want in _texts(dlg)
        assert "bad-address" not in dlg.view.toPlainText()                    # not in the email body
    finally:
        dlg.reject()


def test_preview_dialog_no_invalid_line_when_none(ctx):
    from uat_tracker.ui import weekly as wk
    g = W.generate(ctx.cfg, ctx.repo, actor=PMO_ACTOR, as_of=MON, use_ai=False)
    dlg = wk.WeeklyPreviewDialog(ctx, g.report_id)
    try:
        prefix = i18n.t("wk.invalid", n=0, w="").split("0")[0]
        assert not any(t.startswith(prefix) for t in _texts(dlg))
    finally:
        dlg.reject()
