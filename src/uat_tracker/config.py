"""Application configuration (config.toml)."""

from __future__ import annotations

import os
import tomllib
from dataclasses import dataclass, field
from datetime import date, time
from pathlib import Path

DEFAULT_CONFIG_NAME = "config.toml"
DEFAULT_DATA_DIR = "%LOCALAPPDATA%/PlanningProgress"
APP_NAME = "Planning Progress"


DEFAULT_ATTACHMENT_TYPES = ("pdf", "xlsx", "xls", "csv", "docx", "doc", "pptx", "png", "jpg",
                            "zip", "msg")
# BOT "Financial Institutions' Holidays". The path below is the legacy one on the new gateway;
# confirm it on portal.api.bot.or.th once the API key is registered.
DEFAULT_BOT_URL = "https://gateway.api.bot.or.th/financial-institutions-holidays/?year={year}"


@dataclass(frozen=True)
class AlertsConfig:
    fail_threshold: float = 0.05          # X in the Traffic Light rules
    dry_run: bool = True                  # REQUIRED default: nothing reaches real recipients
    pmo_email: str = ""                   # dry-run copies + PMO warnings go here
    send_account: str = ""                # Outlook account to send from; blank = Outlook default
    run_deadline: time = time(9, 30)      # scheduled run that starts later than this is skipped
    bot_url: str = DEFAULT_BOT_URL
    bot_api_key: str = ""                 # env UAT_BOT_API_KEY overrides
    bot_key_header: str = "Authorization"
    attachment_max_mb: int = 25
    attachment_types: tuple[str, ...] = DEFAULT_ATTACHMENT_TYPES


@dataclass(frozen=True)
class SyncConfig:
    debounce_s: float = 5.0          # wait for the file to be quiet before reading it
    poll_s: float = 60.0             # fallback check when no file event arrives
    batch_quiet_s: float = 10.0      # group app changes made within this window into one save
    retry_minutes: tuple[int, ...] = (1, 2, 5, 10, 15, 30)   # locked file: retry at these offsets
    backup_dir: Path | None = None   # default <data_dir>/backups


@dataclass(frozen=True)
class Config:
    excel_path: Path
    data_dir: Path
    backend: str = "sqlite"
    window_start: date = date(2026, 8, 1)
    window_end: date = date(2027, 3, 31)
    header_overrides: dict = field(default_factory=dict)
    alerts: AlertsConfig = field(default_factory=AlertsConfig)
    sync: SyncConfig = field(default_factory=SyncConfig)
    profile_file: Path | None = None     # Part 5 v2: who uses the app on this PC (default below)
    # Multi-project add-on: data_dir is the PROJECT folder (<root>/projects/<id>) once a project is open;
    # app_dir is the root that holds app.db (shared) and projects/.
    app_dir: Path | None = None
    shared_db: Path | None = None        # app.db, attached to every project DB (holidays + shared settings)
    project_id: str | None = None
    project_name: str = ""
    enable_clear_project: bool = False   # [features] enable_clear_project - temporary "Clear project data" (off by default)

    @property
    def db_path(self) -> Path:
        return self.data_dir / "tracker.db"

    @property
    def profile_path(self) -> Path:
        """Per Windows user, outside the (later shared) data folder - approved v2."""
        return self.profile_file or (self.root_dir / "profile.json")

    @property
    def root_dir(self) -> Path:
        return self.app_dir or self.data_dir

    @property
    def app_db_path(self) -> Path:
        return self.root_dir / "app.db"

    @property
    def snapshot_dir(self) -> Path:
        return self.data_dir / "snapshots"

    @property
    def attachment_dir(self) -> Path:
        return self.data_dir / "attachments"

    @property
    def preview_dir(self) -> Path:
        return self.data_dir / "previews"

    @property
    def backup_dir(self) -> Path:
        return self.sync.backup_dir or (self.data_dir / "backups")


def _expand(p: str) -> Path:
    return Path(os.path.expandvars(os.path.expanduser(p)))


def _to_date(v) -> date:
    if isinstance(v, date):
        return v
    return date.fromisoformat(str(v))


def _load_alerts(raw: dict, path: Path) -> AlertsConfig:
    a = raw.get("alerts", {})
    d = AlertsConfig()
    x = float(a.get("fail_threshold", d.fail_threshold))
    if not 0 <= x < 1:
        raise ValueError(f"{path}: [alerts] fail_threshold must be between 0 and 1 (5% = 0.05)")
    deadline = a.get("run_deadline", d.run_deadline)
    if not isinstance(deadline, time):
        deadline = time.fromisoformat(str(deadline))
    dry = a.get("dry_run", True)
    if not isinstance(dry, bool):
        raise ValueError(f"{path}: [alerts] dry_run must be true or false")
    return AlertsConfig(
        fail_threshold=x,
        dry_run=dry,
        pmo_email=str(a.get("pmo_email", "")).strip(),
        send_account=str(a.get("send_account", "")).strip(),
        run_deadline=deadline,
        bot_url=str(a.get("bot_url", d.bot_url)),
        bot_api_key=os.environ.get("UAT_BOT_API_KEY") or str(a.get("bot_api_key", "")),
        bot_key_header=str(a.get("bot_key_header", d.bot_key_header)),
        attachment_max_mb=int(a.get("attachment_max_mb", d.attachment_max_mb)),
        attachment_types=tuple(t.lower().lstrip(".") for t in a.get("attachment_types", d.attachment_types)),
    )


def _load_sync(raw: dict) -> SyncConfig:
    s = raw.get("sync", {})
    d = SyncConfig()
    retry = tuple(int(x) for x in s.get("retry_minutes", d.retry_minutes))
    if not retry or list(retry) != sorted(retry):
        raise ValueError("[sync] retry_minutes must be increasing minutes, e.g. [1, 2, 5, 10, 15, 30]")
    return SyncConfig(
        debounce_s=float(s.get("debounce_s", d.debounce_s)),
        poll_s=float(s.get("poll_s", d.poll_s)),
        batch_quiet_s=float(s.get("batch_quiet_s", d.batch_quiet_s)),
        retry_minutes=retry,
        backup_dir=_expand(s["backup_dir"]) if s.get("backup_dir") else None,
    )


def load_config(path: str | Path | None = None) -> Config:
    path = Path(path) if path else Path.cwd() / DEFAULT_CONFIG_NAME
    if not path.exists():
        raise FileNotFoundError(
            f"Config not found: {path}. Copy config.example.toml to config.toml and edit it."
        )
    # utf-8-sig: Notepad / PowerShell 5.1 may save the file with a BOM, which tomllib rejects.
    raw = tomllib.loads(path.read_bytes().decode("utf-8-sig"))
    src = raw.get("source", {})
    storage = raw.get("storage", {})
    rules = raw.get("rules", {})
    if "excel_path" not in src:
        raise ValueError(f"{path}: [source] excel_path is required")
    cfg = Config(
        excel_path=_expand(src["excel_path"]),
        data_dir=_expand(storage.get("data_dir", DEFAULT_DATA_DIR)),
        backend=storage.get("backend", "sqlite"),
        window_start=_to_date(rules.get("date_window_start", "2026-08-01")),
        window_end=_to_date(rules.get("date_window_end", "2027-03-31")),
        header_overrides=raw.get("headers", {}),
        alerts=_load_alerts(raw, path),
        sync=_load_sync(raw),
        enable_clear_project=bool(raw.get("features", {}).get("enable_clear_project", False)),
    )
    if cfg.window_start > cfg.window_end:
        raise ValueError(f"{path}: date_window_start is after date_window_end")
    return cfg


# --------------------------------------------------------------------------- Settings screen (Part 4)
# Values changed in the app's Settings screen are stored in the DB and override config.toml, so the
# app, the CLI and the scheduled job all see the same settings.
OVERRIDE_KEYS = ("fail_threshold", "dry_run", "pmo_email", "email_time")
# Multi-project: these live in app.db (shared by all projects); every other setting is per project.
SHARED_SETTING_KEYS = frozenset({"settings.pmo_email", "ui.lang", "settings.ai_model"})


def apply_overrides(cfg: Config, repo) -> Config:
    import dataclasses
    get = lambda k: repo.get_setting(f"settings.{k}")
    alerts = cfg.alerts
    if (v := get("fail_threshold")) is not None:
        alerts = dataclasses.replace(alerts, fail_threshold=float(v))
    if (v := get("dry_run")) is not None:
        alerts = dataclasses.replace(alerts, dry_run=(v == "true"))
    if (v := get("pmo_email")) is not None:
        alerts = dataclasses.replace(alerts, pmo_email=v)
    out = dataclasses.replace(cfg, alerts=alerts)
    if cfg.project_id is None and (v := get("excel_path")):      # single-file mode (tests / before migration)
        out = dataclasses.replace(out, excel_path=_expand(v))
    return out


def email_time(repo) -> str:
    return repo.get_setting("settings.email_time") or "08:30"
