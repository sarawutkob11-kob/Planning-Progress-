"""Manual "Send reminder" for one UAT Plan or Defect List item (PART2 §E). Respects dry-run.

These sheets hold names, not email addresses, so the PMO gives the recipient (--to). The Part 4
button will fill it in.
"""

from __future__ import annotations

import re
from datetime import datetime

from ..alerts.outbox import Outgoing, deliver
from ..alerts.templates import render_manual
from ..domain.models import CHANNEL_APP
from ..excel.sheets import DEFECT_LIST, UAT_PLAN
from .auth import REMIND, Actor, require

_EMAIL = re.compile(r"^[^@\s;]+@[^@\s;]+\.[^@\s;]+$")


class ReminderError(Exception):
    pass


def find_item(repo, item_id: str) -> tuple[str, dict]:
    for sheet in (UAT_PLAN, DEFECT_LIST):
        rec = repo.get_item(sheet, item_id)
        if rec is not None:
            if not rec.get("active"):
                raise ReminderError(f"{item_id} is no longer in the Excel file (inactive)")
            return sheet, rec
    raise ReminderError(f"No UAT Plan or Defect List item {item_id!r} "
                        "(Test Progress items get automatic emails instead)")


def send_reminder(cfg, repo, sender, item_id: str, *, to: str, cc: list[str] | None = None,
                  note: str | None = None, recipient_name: str | None = None,
                  actor: Actor, now: str | None = None) -> str:
    require(repo, actor, REMIND, item_id=item_id)
    editor = actor.editor
    now = now or datetime.now().astimezone().isoformat(timespec="seconds")
    cc = [c.strip() for c in (cc or []) if c.strip()]
    for addr in [to, *cc]:
        if not _EMAIL.match(addr or ""):
            raise ReminderError(f"Not a valid email address: {addr!r}")
    sheet, rec = find_item(repo, item_id)
    name = recipient_name or rec.get("responsible") or rec.get("assigned_to")
    subject, html = render_manual(sheet, rec, note, name)
    out = Outgoing("MANUAL", None, sheet, item_id, rec.get("system") or rec.get("task_name"),
                   to, cc, subject, html)
    if sender.delivers and (not cfg.alerts.dry_run or cfg.alerts.pmo_email):
        sender.check()                    # raises OutlookNotRunning
    return deliver(cfg, repo, sender, out, send_date=now[:10], now=now, editor=editor, channel=CHANNEL_APP)


# --------------------------------------------------------------------------- remembered recipients (Part 4)

def _key(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip()).casefold()


def lookup_contact(repo, name: str | None) -> str | None:
    """Email last used for this owner name (approved: remember name -> email)."""
    if not name:
        return None
    c = repo.get_contact(_key(name))
    return c["email"] if c else None


def remember_contact(repo, name: str | None, email: str) -> None:
    if not name or not _EMAIL.match(email or ""):
        return
    with repo.transaction():
        repo.set_contact(_key(name), name.strip(), email.strip(),
                         datetime.now().astimezone().isoformat(timespec="seconds"))
