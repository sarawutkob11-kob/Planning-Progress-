"""Roles and permissions (Part 5). Checked here, in the service layer; the UI only hides what a role can't do.

Approved 09-Nov-2026:
- Admin: everything + user management.
- PMO and User BU/PM: the same rights (edit all, Conflict Center + review queue, restore, reminders,
  Settings except users).
- QA: edit OWN rows, upload to OWN Test Progress rows, add defects; read everything else.
  Own = Test Progress Report By Email = login email; UAT Plan Responsible / Defect Assigned To = one of
  the user's "names as written in Excel" (cells often list several people: any listed name counts).
- The role is read from the database on every check, so a role change applies at once.
- A denied action raises PermissionDenied and writes an audit entry.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime

from ..excel.sheets import DEFECT_LIST, TEST_PROGRESS, UAT_PLAN

ADMIN, PMO, QA, BU = "Admin", "PMO", "QA", "User BU/PM"
ROLES = (ADMIN, PMO, QA, BU)
SYSTEM = "system"
CHANNEL_AUTH = "auth"
AUTH_SHEET = "Users"

# actions
EDIT = "edit"
ADD_TASK = "add_task"
ADD_DEFECT = "add_defect"
ADD_APP = "add_app"          # multi-project: + Add Application (Test Progress row); full roles only
UPLOAD = "upload"
RESOLVE = "resolve"          # Conflict Center + PMO review queue
RESTORE = "restore"          # Edit History restore + backup restore
REMIND = "remind"
SETTINGS = "settings"
USERS = "users"
WEEKLY = "weekly"            # Part 6: generate / send the weekly report (Admin, PMO, User BU/PM)

ACTION_TEXT = {EDIT: "edit this row", ADD_TASK: "add tasks", ADD_DEFECT: "add defects", ADD_APP: "add applications",
               UPLOAD: "upload to this row", RESOLVE: "resolve conflicts / review items",
               RESTORE: "restore", REMIND: "send reminders", SETTINGS: "change Settings",
               USERS: "manage users",
               WEEKLY: "generate or send the weekly report"}
FULL_ROLES = (ADMIN, PMO, BU)
OWNER_FIELD = {UAT_PLAN: "responsible", DEFECT_LIST: "assigned_to"}


class PermissionDenied(Exception):
    pass


@dataclass(frozen=True)
class Actor:
    email: str
    role: str
    display_name: str = ""           # "First Last"

    @property
    def is_system(self) -> bool:
        return self.role == SYSTEM

    @property
    def editor(self) -> str:
        """What every audit / email log entry shows (approved v2): "First Last <email>"."""
        if self.is_system or not self.display_name:
            return self.email
        return f"{self.display_name} <{self.email}>"


def system_actor(name: str) -> Actor:
    """Scheduled job / sync: no login, never checked; the name goes into the audit log as before."""
    return Actor(name, SYSTEM)


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


# --------------------------------------------------------------------------- names

def name_key(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip()).casefold()


_SPLIT = re.compile(r"[\n\r/,;]+")


def split_names(cell: str | None) -> list[str]:
    """'QA ANAN SUKSAN\\nMALEE S.' / 'QA Pornpan/QA Nida' -> one entry per person."""
    return [n.strip() for n in _SPLIT.split(cell or "") if n.strip()]


# --------------------------------------------------------------------------- checks

def current_role(repo, actor: Actor) -> str | None:
    """Role from the DB (None = unknown or deactivated user)."""
    if actor.is_system:
        return SYSTEM
    u = repo.get_user(actor.email.lower())
    return u["role"] if u and u["active"] else None


def owns(repo, actor: Actor, sheet: str, rec: dict | None) -> bool:
    if rec is None:
        return False
    if sheet == TEST_PROGRESS:
        return (rec.get("report_by_email") or "").strip().lower() == actor.email.lower()
    field = OWNER_FIELD.get(sheet)
    if field is None:
        return False
    mine = {a["name_key"] for a in repo.user_aliases(actor.email.lower())}
    return any(name_key(n) in mine for n in split_names(rec.get(field)))


def can(repo, actor: Actor | None, action: str, sheet: str | None = None, rec: dict | None = None) -> bool:
    if actor is None:
        return False
    role = current_role(repo, actor)
    if role in (SYSTEM, ADMIN):
        return True
    if role is None or action == USERS:
        return False
    if role in FULL_ROLES:
        return True
    if role == QA:
        if action == ADD_DEFECT:
            return True
        if action == EDIT:
            return owns(repo, actor, sheet, rec)
        if action == UPLOAD:
            return sheet == TEST_PROGRESS and owns(repo, actor, sheet, rec)
    return False


def audit_auth(repo, *, editor: str, field: str, item_id: str | None = None, old="", new="", reason="",
               sheet: str = AUTH_SHEET, now: str | None = None) -> None:
    entry = {"ts": now or now_iso(), "editor": editor, "channel": CHANNEL_AUTH, "sheet": sheet,
             "item_id": item_id, "field": field, "old_value": old, "new_value": new, "reason": reason}
    if getattr(repo, "in_transaction", False):
        repo.append_audit(entry)
    else:
        with repo.transaction():
            repo.append_audit(entry)


def require(repo, actor: Actor | None, action: str, *, sheet: str | None = None, item_id: str | None = None,
            rec: dict | None = None) -> None:
    if actor is not None and can(repo, actor, action, sheet, rec):
        return
    who = actor.editor if actor else "(no user)"
    role = (current_role(repo, actor) if actor else None) or "no active user"
    what = ACTION_TEXT.get(action, action)
    msg = f"Not allowed: {role} cannot {what}" + (f" ({item_id})" if item_id else "")
    if role == QA and action in (EDIT, UPLOAD) and item_id:
        msg += " - it is not your row"
    audit_auth(repo, editor=who, field="denied", sheet=sheet or AUTH_SHEET, item_id=item_id, new=action,
               reason=msg)
    raise PermissionDenied(msg)
