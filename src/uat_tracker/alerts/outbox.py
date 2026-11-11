"""One place where every email goes out: dry-run, previews, dedupe, email log, audit log.

Dry-run (default, approved 6k):
- the email goes to [alerts] pmo_email only, subject prefixed "[DRY RUN]", banner with the real To/CC;
- an .html copy is saved under <data_dir>/previews/<date>/;
- with no pmo_email (or --preview-only) only the preview is written. Nothing reaches real recipients.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from ..domain.models import CHANNEL_SYSTEM
from ..mail.sender import Mail, SendError
from .templates import dry_run_banner, render_pmo_alert, with_banner

DRY, LIVE = "dry_run", "live"
SENT, DRY_RESULT, SKIPPED, FAILED = "sent", "dry_run", "skipped", "failed"
NO_DEDUPE_RULES = ("MANUAL", "PMO_ALERT")


@dataclass
class Outgoing:
    rule: str
    variant: str | None
    sheet: str | None
    item_id: str | None
    app: str | None
    to: str | None
    cc: list[str] = field(default_factory=list)
    subject: str = ""
    html: str = ""
    event_ref: int | None = None


def mode_of(cfg) -> str:
    return DRY if cfg.alerts.dry_run else LIVE


def _safe(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", s or "")[:60]


def save_preview(cfg, out: Outgoing, subject: str, html: str, send_date: str, now: str) -> str:
    folder = cfg.preview_dir / send_date
    folder.mkdir(parents=True, exist_ok=True)
    stamp = now[11:19].replace(":", "")
    path = folder / f"{stamp}_{_safe(out.rule)}_{_safe(out.item_id or 'pmo')}.html"
    n = 1
    while path.exists():
        n += 1
        path = folder / f"{stamp}_{_safe(out.rule)}_{_safe(out.item_id or 'pmo')}_{n}.html"
    path.write_text(f"<!-- Subject: {subject} -->\n" + html, encoding="utf-8")
    return str(path)


def deliver(cfg, repo, sender, out: Outgoing, *, send_date: str, now: str, editor: str,
            channel: str = CHANNEL_SYSTEM, job_run_id: int | None = None) -> str:
    """Send (or preview) one email and log it. Returns the result."""
    mode = mode_of(cfg)
    result, error, delivered, preview = None, None, None, None
    subject = out.subject

    if not out.to:
        result, error = SKIPPED, "Report By Email is empty (flagged: missing email)"
    elif out.rule not in NO_DEDUPE_RULES and out.item_id and repo.email_done(out.rule, out.item_id, send_date, mode):
        result, error = SKIPPED, "Already sent today (dedupe: same rule + item + day)"
    elif mode == DRY:
        subject = "[DRY RUN] " + out.subject
        html = with_banner(out.html, dry_run_banner(out.to, out.cc))
        preview = save_preview(cfg, out, subject, html, send_date, now)
        pmo = cfg.alerts.pmo_email
        try:
            if sender.delivers and pmo:
                sender.send(Mail(to=pmo, cc=[], subject=subject, html=html))
                delivered = pmo
            result = DRY_RESULT
        except SendError as exc:
            result, error = FAILED, str(exc)
    else:
        if not sender.delivers:
            raise RuntimeError("Live mode needs a real sender (Outlook); --preview-only is dry-run only")
        try:
            sender.send(Mail(to=out.to, cc=list(out.cc), subject=out.subject, html=out.html))
            result, delivered = SENT, out.to + (("; " + "; ".join(out.cc)) if out.cc else "")
        except SendError as exc:
            result, error = FAILED, str(exc)

    _log(repo, out, mode=mode, result=result, error=error, delivered=delivered, preview=preview,
         subject=subject, send_date=send_date, now=now, editor=editor, channel=channel,
         job_run_id=job_run_id)
    return result


def _log(repo, out: Outgoing, *, mode, result, error, delivered, preview, subject, send_date, now,
         editor, channel, job_run_id) -> None:
    with repo.transaction():
        repo.insert_email({
            "ts": now, "send_date": send_date, "rule": out.rule, "variant": out.variant,
            "sheet": out.sheet, "item_id": out.item_id, "app": out.app, "to_addr": out.to,
            "cc": "; ".join(out.cc), "delivered_to": delivered, "subject": subject, "mode": mode,
            "result": result, "error": error, "event_ref": out.event_ref, "preview_path": preview,
            "job_run_id": job_run_id,
        })
        reason = error or (f"To {out.to}" + (f"; CC {'; '.join(out.cc)}" if out.cc else "")
                           + (f" (dry run, delivered to {delivered})" if mode == DRY and delivered else "")
                           + (" (dry run, preview only)" if mode == DRY and not delivered and result == DRY_RESULT else ""))
        repo.append_audit({"ts": now, "editor": editor, "channel": channel, "sheet": out.sheet,
                           "item_id": out.item_id, "field": f"email {out.rule}" + (f" ({out.variant})" if out.variant else ""),
                           "old_value": "", "new_value": f"{result}: {subject}", "reason": reason})


def alert_pmo(cfg, repo, sender, title: str, detail: str, *, send_date: str, now: str, editor: str,
              job_run_id: int | None = None) -> str:
    """Warning to the PMO only (fetch failure, import failure, no holiday data...).
    Goes to pmo_email for real in both modes: it never reaches anyone else."""
    subject, html = render_pmo_alert(title, detail)
    out = Outgoing("PMO_ALERT", None, None, None, None, cfg.alerts.pmo_email or None, [], subject, html)
    result, error, delivered = None, None, None
    if not out.to:
        result, error = SKIPPED, "No [alerts] pmo_email configured"
    else:
        try:
            sender.check()
            if sender.delivers:
                sender.send(Mail(to=out.to, subject=subject, html=html))
                result, delivered = SENT, out.to
            else:
                result, error = SKIPPED, "Preview-only run: PMO warning not sent"
        except SendError as exc:
            result, error = FAILED, str(exc)
    _log(repo, out, mode=LIVE, result=result, error=error or detail[:500], delivered=delivered, preview=None,
         subject=subject, send_date=send_date, now=now, editor=editor, channel=CHANNEL_SYSTEM,
         job_run_id=job_run_id)
    return result
