"""User management (Part 5 v2, Admin only): first / last name, email, role. No delete: deactivate keeps the user and their audit trail.

Every change is in the audit log (channel auth). The last active Admin can't be demoted or deactivated.
"""

from __future__ import annotations

import re

from ..excel.sheets import DEFECT_LIST, TEST_PROGRESS, UAT_PLAN
from .auth import ADMIN, QA, ROLES, USERS, Actor, audit_auth, name_key, now_iso, require, split_names
from .profile import norm_email

_EMAIL = re.compile(r"^[^@\s;]+@[^@\s;]+\.[^@\s;]+$")


class UserError(Exception):
    pass


def _clean_aliases(names) -> dict[str, str]:
    out = {}
    for n in names or ():
        n = re.sub(r"\s+", " ", str(n)).strip()
        if n:
            out.setdefault(name_key(n), n)
    return out


def list_users(repo) -> list[dict]:
    rows = repo.list_users()
    for u in rows:
        u["aliases"] = [a["name"] for a in repo.user_aliases(u["email"])]
    return rows


def add_user(repo, actor: Actor, email: str, *, role: str, first_name: str = "", last_name: str = "",
             aliases=()) -> str:
    require(repo, actor, USERS)
    email = norm_email(email)
    if not _EMAIL.match(email):
        raise UserError(f"Not a valid email address: {email!r}")
    if role not in ROLES:
        raise UserError(f"Role must be one of: {', '.join(ROLES)}")
    if repo.get_user(email) is not None:
        raise UserError(f"{email} already exists (reactivate or edit it instead)")
    ts = now_iso()
    al = _clean_aliases(aliases)
    with repo.transaction():
        repo.insert_user(email=email, first_name=first_name.strip() or None, last_name=last_name.strip() or None,
                         role=role, created_at=ts, created_by=actor.editor)
        repo.set_user_aliases(email, al)
        audit_auth(repo, editor=actor.editor, field="user_added", item_id=email, new=role,
                   reason=f"User added; names in Excel: {', '.join(al.values()) or '-'}", now=ts)
    return email


def update_user(repo, actor: Actor, email: str, *, role: str | None = None, first_name: str | None = None,
                last_name: str | None = None, aliases=None, active: bool | None = None) -> list[str]:
    """Returns the changed fields."""
    require(repo, actor, USERS)
    email = norm_email(email)
    u = repo.get_user(email)
    if u is None:
        raise UserError(f"No user {email}")
    if role is not None and role not in ROLES:
        raise UserError(f"Role must be one of: {', '.join(ROLES)}")
    losing_admin = u["role"] == ADMIN and u["active"] and (
        (role is not None and role != ADMIN) or active is False)
    if losing_admin and repo.count_active_admins() <= 1:
        raise UserError("This is the last active Admin: add another Admin first")
    ts = now_iso()
    changed = []
    with repo.transaction():
        if role is not None and role != u["role"]:
            repo.update_user(email, role=role)
            audit_auth(repo, editor=actor.editor, field="role", item_id=email, old=u["role"], new=role,
                       reason="Role changed", now=ts)
            changed.append("role")
        if active is not None and bool(active) != bool(u["active"]):
            repo.update_user(email, active=1 if active else 0)
            audit_auth(repo, editor=actor.editor, field="active", item_id=email,
                       old="active" if u["active"] else "deactivated",
                       new="active" if active else "deactivated",
                       reason="User reactivated" if active else "User deactivated", now=ts)
            changed.append("active")
        for fld, val in (("first_name", first_name), ("last_name", last_name)):
            if val is not None and (val.strip() or None) != u[fld]:
                repo.update_user(email, **{fld: val.strip() or None})
                audit_auth(repo, editor=actor.editor, field=fld, item_id=email, old=u[fld] or "", new=val.strip(),
                           reason="Name changed", now=ts)
                changed.append(fld)
        if {"first_name", "last_name"} & set(changed):
            nu = repo.get_user(email)
            repo.update_user(email, display_name=" ".join(x for x in (nu["first_name"], nu["last_name"]) if x))
        if aliases is not None:
            old = {a["name_key"]: a["name"] for a in repo.user_aliases(email)}
            new = _clean_aliases(aliases)
            if set(old) != set(new):
                repo.set_user_aliases(email, new)
                audit_auth(repo, editor=actor.editor, field="aliases", item_id=email,
                           old=", ".join(old.values()), new=", ".join(new.values()),
                           reason="Names in Excel changed", now=ts)
                changed.append("aliases")
    return changed


def excel_names(repo) -> list[str]:
    """Every person named in UAT Plan Responsible / Defect Assigned To (for the alias picker)."""
    seen = {}
    for sheet, field in ((UAT_PLAN, "responsible"), (DEFECT_LIST, "assigned_to")):
        for rec in repo.load_items(sheet).values():
            if rec.get("active"):
                for n in split_names(rec.get(field)):
                    seen.setdefault(name_key(n), n)
    return sorted(seen.values(), key=str.casefold)


def import_candidates(repo) -> list[dict]:
    """Addresses in Test Progress Report By Email + CC Email that are not users yet. The Admin reviews
    them (role defaults to QA) before anything is saved."""
    out: dict[str, dict] = {}
    for rec in repo.load_items(TEST_PROGRESS).values():
        if not rec.get("active"):
            continue
        for field, src in (("report_by_email", "Report By Email"), ("cc_emails", "CC Email")):
            for addr in re.split(r"[;,\s]+", rec.get(field) or ""):
                addr = norm_email(addr)
                if not _EMAIL.match(addr) or repo.get_user(addr) is not None:
                    continue
                c = out.setdefault(addr, {"email": addr, "sources": set(), "apps": set(), "role": QA,
                                          "names": set()})
                c["sources"].add(src)
                if rec.get("application"):
                    c["apps"].add(rec["application"])
                if field == "report_by_email" and rec.get("report_by"):
                    c["names"].update(split_names(rec["report_by"]))
    for c in out.values():
        c["sources"], c["apps"], c["names"] = sorted(c["sources"]), sorted(c["apps"]), sorted(c["names"])
    return sorted(out.values(), key=lambda c: c["email"])
