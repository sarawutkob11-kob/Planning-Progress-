"""Part 5 v2 CLI identity (approved 10-Nov-2026): commands that change data run as the profile saved on this
PC by the app (no login, no code), with the same role rules. Read-only commands, import / sync and the
scheduled job need no profile."""

from __future__ import annotations

from .services.auth import Actor
from .services.profile import open_profile

# (cmd, sub-command attribute, values) that change data
_WRITE = {
    "edit": None, "resolve-review": None, "upload": None, "remind": None,
    "conflicts": ("ccmd", {"resolve"}),
    "backup": ("bcmd", {"restore"}),
    "holidays": ("hcmd", {"fetch", "add", "import", "remove"}),
    "scheduler": ("action", {"install", "remove"}),
    "sync": ("scmd", {"init-ids", "unblock"}),
    "ai-key": ("kcmd", {"set", "clear"}),
}


def needs_login(args) -> bool:
    if args.cmd not in _WRITE:
        return False
    rule = _WRITE[args.cmd]
    return rule is None or getattr(args, rule[0], None) in rule[1]


def cli_identity(cfg, repo, *, out=print) -> Actor | None:
    actor = open_profile(cfg, repo)
    if actor is None:
        out("No user profile on this PC (or the user is deactivated). Open the app once and fill in "
            "your name and email.")
        return None
    out(f"As {actor.editor} ({actor.role})")
    return actor
