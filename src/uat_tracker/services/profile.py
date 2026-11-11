"""Who is using the app (Part 5 v2, approved 10-Nov-2026). No login, no OTP, no email verification.

- First start on a PC (per Windows user): form First name / Last name / Email. Saved in a local profile file
  (%LOCALAPPDATA%\\PlanningProgress\\profile.json); later starts open straight in.
- Email = user key for roles, audit, "own" rows. Audit editor = "First Last <email>".
- Very first profile ever -> Admin (confirmed), once only. A new email later -> QA (Admin adjusts).
- An email that already exists keeps its stored name; a deactivated email can't be used.
- "First Last" is added to the user's "Names in Excel" so owned UAT Plan tasks / defects match.
Known limit (accepted for Phase 1): anyone can type any email and gets that email's role.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from .auth import ADMIN, QA, Actor, audit_auth, name_key, now_iso

_EMAIL = re.compile(r"^[^@\s;<>]+@[^@\s;<>]+\.[^@\s;<>]+$")
DEFAULT_ROLE = QA


class ProfileError(Exception):
    pass


class FirstAdminNeeded(ProfileError):
    """No users yet: the caller must confirm that this person becomes the Admin."""


def norm_email(email: str) -> str:
    return (email or "").strip().lower()


def valid_email(email: str) -> bool:
    return bool(_EMAIL.match(norm_email(email)))


def needs_first_admin(repo) -> bool:
    return repo.count_users() == 0


# --------------------------------------------------------------------------- local profile file

def read_profile_data(path: Path) -> dict:
    """{"email", "first_name", "last_name", "last_project"} (multi-project add-on: names + last project)."""
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def read_profile(path: Path) -> str | None:
    email = norm_email(read_profile_data(path).get("email", ""))
    return email or None


def write_profile(path: Path, email: str | None = None, **extra) -> None:
    """Merge into the profile file (email, first_name, last_name, last_project)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = read_profile_data(path)
    if email is not None:
        if norm_email(email) != norm_email(data.get("email", "")):
            data = {k: v for k, v in data.items() if k == "last_project"}
        data["email"] = norm_email(email)
    data.update({k: v for k, v in extra.items() if v is not None})
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    tmp.replace(path)


# --------------------------------------------------------------------------- users <-> actors

def _actor(u: dict) -> Actor:
    name = " ".join(x for x in (u.get("first_name"), u.get("last_name")) if x) or (u.get("display_name") or "")
    return Actor(u["email"], u["role"], name)


def lookup(repo, email: str) -> dict | None:
    """For the form: an existing user's stored names (they win over what is typed)."""
    return repo.get_user(norm_email(email))


def open_profile(cfg, repo) -> Actor | None:
    """Start-up: the saved profile's user, or None (show the form). Updates "last used"."""
    email = read_profile(cfg.profile_path)
    u = repo.get_user(email) if email else None
    if u is None or not u["active"]:
        return None
    with repo.transaction():
        repo.update_user(email, last_used_at=now_iso())
    return _actor(u)


def identify(cfg, repo, *, first_name: str, last_name: str, email: str, confirm_admin: bool = False,
             previous: Actor | None = None) -> Actor:
    """Save the profile form (first start or "Switch user"). Returns the Actor now using the app."""
    first, last, email = first_name.strip(), last_name.strip(), norm_email(email)
    if not valid_email(email):
        raise ProfileError("Type one valid email address")
    ts = now_iso()
    existing = repo.get_user(email)
    if existing is not None and not existing["active"]:
        audit_auth(repo, editor=f"{first} {last} <{email}>" if first or last else email, field="denied",
                   item_id=email, new="profile", reason="Deactivated user tried to use the app", now=ts)
        raise ProfileError(f"{email} is deactivated: ask the Admin")
    with repo.transaction():
        u = repo.get_user(email)
        if u is None and (not first or not last):
            raise ProfileError("First name and last name are required")
        if repo.count_users() == 0:
            if not confirm_admin:
                raise FirstAdminNeeded(f"No users yet: {first} {last} <{email}> would become the Admin")
            u = _create(repo, email, first, last, ADMIN, ts, "First profile on this installation: Admin")
        elif u is None:
            u = _create(repo, email, first, last, DEFAULT_ROLE, ts, f"New email: default role {DEFAULT_ROLE}")
        repo.update_user(email, last_used_at=ts)
        actor = _actor(repo.get_user(email))
        if previous is not None and not previous.is_system and previous.email != email:
            audit_auth(repo, editor=actor.editor, field="switch_user", item_id=email, old=previous.editor,
                       new=actor.editor, reason="Switch user", now=ts)
    write_profile(cfg.profile_path, email)
    return actor


def _create(repo, email, first, last, role, ts, why) -> dict:
    repo.insert_user(email=email, first_name=first, last_name=last, role=role, created_at=ts,
                     created_by=f"{first} {last} <{email}>")
    full = f"{first} {last}"
    repo.set_user_aliases(email, {name_key(full): full})
    audit_auth(repo, editor=f"{full} <{email}>", field="profile_created", item_id=email, new=role, reason=why,
               now=ts)
    return repo.get_user(email)
