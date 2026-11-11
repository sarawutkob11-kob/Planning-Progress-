"""app.db: the list of projects + what all projects share (holidays, PMO email, language, Claude model).
Multi-project add-on, approved 10-Nov-2026."""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from importlib import resources
from pathlib import Path

PROJECT_COLS = ("id", "name", "excel_path", "created_at", "created_by", "archived", "archived_at",
                "window_start", "window_end")


class AppDb:
    def __init__(self, path: Path | str):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.path), isolation_level=None, timeout=10)
        self.conn.row_factory = sqlite3.Row
        self._migrate()

    def _migrate(self) -> None:
        version = self.conn.execute("PRAGMA user_version").fetchone()[0]
        files = sorted((f for f in resources.files("uat_tracker.storage").joinpath("migrations_app").iterdir()
                        if f.name.endswith(".sql")), key=lambda f: f.name)
        for f in files:
            n = int(f.name.split("_", 1)[0])
            if n > version:
                self.conn.executescript("BEGIN;\n" + f.read_text(encoding="utf-8")
                                        + f"\nPRAGMA user_version = {n};\nCOMMIT;")

    @contextmanager
    def transaction(self):
        self.conn.execute("BEGIN IMMEDIATE")
        try:
            yield self
        except BaseException:
            self.conn.execute("ROLLBACK")
            raise
        else:
            self.conn.execute("COMMIT")

    def close(self) -> None:
        self.conn.close()

    # ------------------------------------------------------------------ projects
    def projects(self, include_archived: bool = False) -> list[dict]:
        q = "SELECT * FROM project" + ("" if include_archived else " WHERE archived=0") + " ORDER BY name COLLATE NOCASE"
        return [dict(r) for r in self.conn.execute(q)]

    def get_project(self, project_id: str) -> dict | None:
        r = self.conn.execute("SELECT * FROM project WHERE id=?", (project_id,)).fetchone()
        return dict(r) if r else None

    def find_project(self, key: str) -> dict | None:
        """By id, or by name (case-insensitive)."""
        return self.get_project(key) or next(
            (p for p in self.projects(True) if p["name"].casefold() == (key or "").strip().casefold()), None)

    def next_project_id(self) -> str:
        n = [int(r[0][1:]) for r in self.conn.execute("SELECT id FROM project") if r[0][1:].isdigit()]
        return f"p{max(n, default=0) + 1:03d}"

    def insert_project(self, row: dict) -> None:
        cols = [c for c in PROJECT_COLS if c in row]
        self.conn.execute(f"INSERT INTO project ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                          [row[c] for c in cols])

    def update_project(self, project_id: str, **fields) -> None:
        bad = set(fields) - set(PROJECT_COLS[1:])
        if bad:
            raise KeyError(f"project has no field(s) {bad}")
        self.conn.execute(f"UPDATE project SET {', '.join(f'{k}=?' for k in fields)} WHERE id=?",
                          [*fields.values(), project_id])

    def log_event(self, *, ts: str, project_id: str | None, editor: str, action: str, detail: str = "") -> None:
        self.conn.execute("INSERT INTO project_event (ts, project_id, editor, action, detail) VALUES (?,?,?,?,?)",
                          (ts, project_id, editor, action, detail))

    def events(self, project_id: str | None = None, limit: int = 50) -> list[dict]:
        if project_id:
            rows = self.conn.execute("SELECT * FROM project_event WHERE project_id=? ORDER BY id DESC LIMIT ?",
                                     (project_id, limit))
        else:
            rows = self.conn.execute("SELECT * FROM project_event ORDER BY id DESC LIMIT ?", (limit,))
        return [dict(r) for r in rows]

    # ------------------------------------------------------------------ shared settings
    def get_setting(self, key: str) -> str | None:
        r = self.conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return r["value"] if r else None

    def set_setting(self, key: str, value: str | None) -> None:
        self.conn.execute("INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key)"
                          " DO UPDATE SET value=excluded.value", (key, value))
