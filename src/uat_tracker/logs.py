"""Log files (Hardening phase 1, 11-Nov-2026).

<root>/logs/app.log   the desktop app and CLI commands
<root>/logs/jobs.log  scheduled runs (--scheduled: daily alerts, weekly report)

<root> is the data folder that holds app.db (default %LOCALAPPDATA%\\PlanningProgress), shared by all projects;
every line names the open project. Each file rotates at 5 MB, 5 old files kept. User-facing CLI output still goes
to the console (print); logs are for diagnosing problems afterwards.
"""

from __future__ import annotations

import logging
import sys
import threading
from logging.handlers import RotatingFileHandler
from pathlib import Path

LOGGER = "uat_tracker"
MAX_BYTES = 5 * 1024 * 1024
BACKUPS = 5
FORMAT = "%(asctime)s %(levelname)-7s [%(project)s] %(threadName)s %(name)s: %(message)s"

_project = "-"
_handler: RotatingFileHandler | None = None


def set_project(project_id: str | None, name: str | None = None) -> None:
    """Called when a project is opened: every following log line names it."""
    global _project
    _project = f"{project_id}:{name}" if project_id and name else (project_id or "-")


class _ProjectFilter(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        record.project = _project
        return True


def log_dir(root: Path) -> Path:
    return Path(root) / "logs"


def setup(root: Path, kind: str = "app", level: int = logging.INFO) -> Path:
    """Send the 'uat_tracker' loggers to <root>/logs/<kind>.log. Safe to call more than once."""
    global _handler
    path = log_dir(root) / f"{kind}.log"
    path.parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger(LOGGER)
    if _handler is not None:
        logger.removeHandler(_handler)
        _handler.close()
    _handler = RotatingFileHandler(path, maxBytes=MAX_BYTES, backupCount=BACKUPS, encoding="utf-8", delay=True)
    _handler.setFormatter(logging.Formatter(FORMAT))
    _handler.addFilter(_ProjectFilter())
    logger.addHandler(_handler)
    logger.setLevel(level)
    _hook_uncaught()
    return path


def _hook_uncaught() -> None:
    """Crashes that nothing caught (also in background threads) end up in the log too."""
    log = logging.getLogger(LOGGER)
    if getattr(sys.excepthook, "_uat_logged", False):
        return
    prev_sys, prev_thread = sys.excepthook, threading.excepthook

    def sys_hook(t, v, tb):
        if not issubclass(t, KeyboardInterrupt):
            log.critical("Uncaught error", exc_info=(t, v, tb))
        prev_sys(t, v, tb)

    def thread_hook(args):
        if args.exc_type is not SystemExit:
            log.critical("Uncaught error in thread %s", getattr(args.thread, "name", "?"),
                         exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
        prev_thread(args)
    sys_hook._uat_logged = True
    sys.excepthook, threading.excepthook = sys_hook, thread_hook
