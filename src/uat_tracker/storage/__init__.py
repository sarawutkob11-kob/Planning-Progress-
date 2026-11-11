"""Storage layer. Business logic depends only on the Repository protocol."""

from __future__ import annotations

from ..config import Config
from .repository import Repository


def open_repository(cfg: Config) -> Repository:
    """The project's own DB; with multi-project, app.db (cfg.shared_db) is attached for shared data."""
    if cfg.backend == "sqlite":
        from .sqlite_repo import SqliteRepository
        cfg.data_dir.mkdir(parents=True, exist_ok=True)
        return SqliteRepository(cfg.db_path, shared=cfg.shared_db)
    raise ValueError(f"Unknown storage backend: {cfg.backend}")
