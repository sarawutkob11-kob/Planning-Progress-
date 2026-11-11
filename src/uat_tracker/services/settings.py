"""Settings screen actions (Part 4 values, Part 5 role check: Admin / PMO / User BU/PM)."""

from __future__ import annotations

import datetime as dt

from .auth import SETTINGS, Actor, now_iso, require


def save_settings(repo, actor: Actor, *, fail_pct: float, dry_run: bool, pmo_email: str | None, email_time: str,
                  excel_path: str | None = None) -> None:
    """Multi-project: X%, dry-run and send time are per project; pmo_email is shared (app.db) - pass None to
    leave it (the caller lets only an Admin / PMO of any project change it). excel_path: single-file mode only
    (projects use Relink)."""
    require(repo, actor, SETTINGS)
    with repo.transaction():
        repo.set_setting("settings.fail_threshold", repr(round(fail_pct / 100, 4)))
        repo.set_setting("settings.dry_run", "true" if dry_run else "false")
        if pmo_email is not None:
            repo.set_setting("settings.pmo_email", pmo_email.strip())
        repo.set_setting("settings.email_time", email_time)
        if excel_path:
            repo.set_setting("settings.excel_path", excel_path.strip())
        repo.append_audit({"ts": now_iso(), "editor": actor.editor, "channel": "app", "sheet": "Settings",
                           "item_id": None, "field": "settings", "old_value": "",
                           "new_value": f"X={fail_pct:.1f}%, dry_run={dry_run}"
                                        + (f", pmo={pmo_email.strip()}" if pmo_email is not None else "")
                                        + f", time={email_time}" + (f", excel={excel_path}" if excel_path else ""),
                           "reason": "Changed in Settings"})


def fetch_holidays(cfg, repo, actor: Actor, years=None) -> list[str]:
    from .holidays import refresh_year
    require(repo, actor, SETTINGS)
    msgs = []
    today = dt.date.today()
    for y in years or (today.year, today.year + 1):
        r = refresh_year(repo, cfg, y, now=now_iso(), editor=actor.editor)
        msgs.append(f"{y}: {'OK' if r.ok else 'FAILED'} {r.message}")
    return msgs


def install_schedule(repo, actor: Actor, config_path, at: str, project_id: str | None = None,
                     project_name: str | None = None):
    """Daily alerts task + (Part 6) weekly report task, for one project. Returns the two schtasks results."""
    from . import scheduler
    require(repo, actor, SETTINGS)
    daily = scheduler.install(config_path, at, project_id, project_name)
    weekly = scheduler.install_weekly(config_path, project_id=project_id, project_name=project_name)
    if daily.returncode == 0 and weekly.returncode == 0:         # new names in place: drop the old-named tasks
        removed = scheduler.remove_old_tasks(project_name)
        if removed:
            weekly.stdout = (weekly.stdout or "") + "\nRemoved old task(s): " + ", ".join(removed)
    return daily, weekly


def save_ai_settings(repo, actor: Actor, *, model: str, api_key: str | None = None,
                     clear_key: bool = False) -> None:
    """Part 6: Claude model (DB setting; blank = use config [weekly_ai] model) + API key (Windows Credential
    Manager only, never the DB/log)."""
    from . import weekly_ai
    require(repo, actor, SETTINGS)
    model = (model or "").strip()
    if api_key and api_key.strip():
        weekly_ai.KEYS.set(api_key.strip())
    elif clear_key:
        weekly_ai.KEYS.clear()
    with repo.transaction():
        repo.set_setting(weekly_ai.S_MODEL, model)
        repo.append_audit({"ts": now_iso(), "editor": actor.editor, "channel": "app", "sheet": "Settings",
                           "item_id": None, "field": "ai_settings", "old_value": "",
                           "new_value": f"model={model or '(config default)'}" + ("; API key updated" if api_key and api_key.strip()
                                                           else "; API key removed" if clear_key else ""),
                           "reason": "Weekly report AI settings changed"})
