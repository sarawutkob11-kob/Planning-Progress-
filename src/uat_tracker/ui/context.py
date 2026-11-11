"""State shared by the screens: config, repository, logged-in user (Part 5), As-of date, language."""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from ..config import apply_overrides
from ..services import auth
from . import i18n


@dataclass
class Ctx:
    base_cfg: object
    repo: object
    actor: object = None             # services.auth.Actor of the logged-in user
    as_of: dt.date = field(default_factory=dt.date.today)
    search: str = ""
    config_path: object = None       # config.toml (for the Windows scheduler entry)
    navigate: object = None          # callable(screen_key, **kw) set by the main window
    refresh_all: object = None
    switch_user: object = None       # callable() set by the app controller (Part 5 v2)
    app: object = None               # multi-project: the AppController (project list, switch, new, import)

    def __post_init__(self):
        self.cfg = apply_overrides(self.base_cfg, self.repo)
        i18n.set_lang(self.repo.get_setting("ui.lang") or "en")

    @property
    def editor(self) -> str:
        return self.actor.editor if self.actor else ""

    def can(self, action: str, sheet: str | None = None, rec: dict | None = None) -> bool:
        """For hiding / disabling buttons. The services check again (and log a denial)."""
        return auth.can(self.repo, self.actor, action, sheet, rec)

    def role(self) -> str | None:
        return auth.current_role(self.repo, self.actor) if self.actor else None

    def reload_cfg(self):
        self.cfg = apply_overrides(self.base_cfg, self.repo)

    def set_lang(self, lang: str):
        i18n.set_lang(lang)
        with self.repo.transaction():
            self.repo.set_setting("ui.lang", lang)
