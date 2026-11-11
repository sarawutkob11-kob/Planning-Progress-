"""Long actions off the UI thread (Hardening phase 1, section D; same pattern as the Weekly report worker).

run_in_worker() shows a modal "please wait" box (which also blocks a second click), runs the job in a QThread
and calls back on the UI thread when it ends. The job gets its OWN database connections: a sqlite3 connection
can only be used in the thread that opened it.

SyncBusy (another sync is writing the workbook right now) -> "Sync running, try again in a moment".
"""

from __future__ import annotations

import logging
from contextlib import contextmanager

from PySide6.QtCore import QThread, Qt, Signal
from PySide6.QtWidgets import QProgressDialog

from .dialogs import error
from .i18n import t

log = logging.getLogger(__name__)
_running: set = set()                                    # every job not finished yet (see wait_for_workers)


def wait_for_workers(timeout_ms: int = 60000) -> None:
    """Wait until all background jobs ended and their callbacks ran (tests; app shutdown)."""
    from PySide6.QtCore import QCoreApplication
    for job in list(_running):
        job.wait(timeout_ms)
    for _ in range(5):                                   # deliver the queued done / failed signals
        QCoreApplication.processEvents()


class Job(QThread):
    done = Signal(object)
    failed = Signal(object)

    def __init__(self, fn, name: str):
        super().__init__()
        self.fn, self.name = fn, name
        self.setObjectName(f"worker:{name}")

    def run(self):
        try:
            res = self.fn()
        except BaseException as exc:                     # noqa: BLE001 - reported on the UI thread
            log.error("Background job '%s' failed", self.name, exc_info=True)
            self.failed.emit(exc)
        else:
            self.done.emit(res)


def friendly(exc: BaseException) -> str:
    from ..services.sync import SyncBusy
    if isinstance(exc, SyncBusy):
        return t("sync.busy_retry")
    return str(exc)


def run_in_worker(parent, text: str, fn, on_done=None, on_error=None, *, name: str = "job") -> Job:
    """Run fn() in a QThread behind a modal progress box. on_done(result) / on_error(exc) run on the UI thread;
    without on_error the error is shown in a message box."""
    prog = QProgressDialog(text, None, 0, 0, parent)
    prog.setWindowTitle(t("app"))
    prog.setWindowModality(Qt.ApplicationModal)          # no second click anywhere while it runs
    prog.setMinimumDuration(0)
    prog.setCancelButton(None)
    job = Job(fn, name)
    jobs = getattr(parent, "_workers", None)
    if jobs is None:
        jobs = parent._workers = set()
    jobs.add(job)                                        # keep a reference while it runs
    _running.add(job)

    def finish():
        prog.close()
        jobs.discard(job)
        _running.discard(job)

    def ok(res):
        finish()
        if on_done:
            on_done(res)

    def bad(exc):
        finish()
        if on_error:
            on_error(exc)
        else:
            error(parent, friendly(exc))
    job.done.connect(ok)
    job.failed.connect(bad)
    prog.show()
    job.start()
    return job


@contextmanager
def project_repo(cfg):
    """A fresh repository for this project, for use inside a worker thread."""
    from ..storage.sqlite_repo import SqliteRepository
    repo = SqliteRepository(cfg.db_path, shared=cfg.shared_db)
    try:
        yield repo
    finally:
        repo.close()


@contextmanager
def sync_locked(cfg):
    """Manual Excel read / import: not while the sync is writing the workbook (-> SyncBusy)."""
    from ..services.sync import file_lock
    with file_lock(cfg.data_dir / "sync.lock"):         # same lock file as SyncService
        yield
