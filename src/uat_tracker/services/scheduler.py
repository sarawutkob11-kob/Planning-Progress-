"""Windows Task Scheduler entries (approved: Mon-Fri, only when the user is logged on, "run missed task" OFF ->
no backfill; the job itself skips bank holidays and late starts).

Multi-project (approved 10-Nov-2026): one daily-alerts task and one weekly-report task PER PROJECT, named
"Planning Progress - Daily Alerts - <project>" / "Planning Progress - Weekly Report - <project>", each
running the job with --project <id> at that project's own send time (weekly: 09:00 for every project)."""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

PREFIX = "Planning Progress"
DAILY = "Daily Alerts"
WEEKLY = "Weekly Report"
WEEKLY_TIME = "09:00"
# Names used before the rename (removed by the migration).
LEGACY_TASKS = ("UAT Tracker - Daily Alerts", "UAT Tracker - Weekly Report")
# Single-project names (CLI without --project, tests).
TASK_NAME = f"{PREFIX} - {DAILY}"
WEEKLY_TASK_NAME = f"{PREFIX} - {WEEKLY}"


def task_name(kind: str, project_name: str | None = None) -> str:
    base = f"{PREFIX} - {kind}"
    if not project_name:
        return base
    safe = re.sub(r'[/\\:*?"<>|]', "-", project_name).strip(" .")
    return f"{base} - {safe}"


def _pythonw() -> str:
    exe = Path(sys.executable)
    w = exe.with_name("pythonw.exe")
    return str(w if w.exists() else exe)


def _cmd(config_path: Path, job: str, project_id: str | None) -> str:
    proj = f" --project {project_id}" if project_id else ""
    return f'"{_pythonw()}" -m uat_tracker --config "{Path(config_path).resolve()}"{proj} {job} --scheduled'


def task_command(config_path: Path, project_id: str | None = None) -> str:
    return _cmd(config_path, "run-daily", project_id)


def weekly_task_command(config_path: Path, project_id: str | None = None) -> str:
    return _cmd(config_path, "run-weekly", project_id)


def _create_args(name: str, at: str, command: str) -> list[str]:
    # /IT = run only when the user is logged on (Outlook is in that session).
    # schtasks never sets "run task as soon as possible after a scheduled start is missed".
    return ["schtasks", "/Create", "/F", "/TN", name, "/SC", "WEEKLY",
            "/D", "MON,TUE,WED,THU,FRI", "/ST", at, "/IT", "/TR", command]


def install_args(config_path: Path, at: str = "08:30", project_id: str | None = None,
                 project_name: str | None = None) -> list[str]:
    return _create_args(task_name(DAILY, project_name), at, task_command(config_path, project_id))


def install_weekly_args(config_path: Path, at: str = WEEKLY_TIME, project_id: str | None = None,
                        project_name: str | None = None) -> list[str]:
    # Mon-Fri at 09:00: the job itself decides the day (Monday, or the first workday after a Monday holiday).
    return _create_args(task_name(WEEKLY, project_name), at, weekly_task_command(config_path, project_id))


def _run(args: list[str]) -> subprocess.CompletedProcess:
    if len(args[-1]) > 261:                                 # schtasks /TR limit
        raise ValueError("Task command is longer than 261 characters; move config.toml to a shorter path")
    return subprocess.run(args, capture_output=True, text=True)


def install(config_path: Path, at: str = "08:30", project_id: str | None = None,
            project_name: str | None = None) -> subprocess.CompletedProcess:
    return _run(install_args(config_path, at, project_id, project_name))


def install_weekly(config_path: Path, at: str = WEEKLY_TIME, project_id: str | None = None,
                   project_name: str | None = None) -> subprocess.CompletedProcess:
    return _run(install_weekly_args(config_path, at, project_id, project_name))


def remove(project_name: str | None = None) -> subprocess.CompletedProcess:
    return subprocess.run(["schtasks", "/Delete", "/F", "/TN", task_name(DAILY, project_name)],
                          capture_output=True, text=True)


def remove_weekly(project_name: str | None = None) -> subprocess.CompletedProcess:
    return subprocess.run(["schtasks", "/Delete", "/F", "/TN", task_name(WEEKLY, project_name)],
                          capture_output=True, text=True)


def remove_named(name: str) -> subprocess.CompletedProcess:
    return subprocess.run(["schtasks", "/Delete", "/F", "/TN", name], capture_output=True, text=True)


def show(project_name: str | None = None) -> subprocess.CompletedProcess:
    return subprocess.run(["schtasks", "/Query", "/TN", task_name(DAILY, project_name), "/V", "/FO", "LIST"],
                          capture_output=True, text=True)
