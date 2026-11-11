"""2-way sync between the app and the OneDrive-synced workbook (Part 3, master spec §7).

Excel -> app (pull): file watcher (+ 60 s fallback check) -> Part 1 import with channel excel_sync.
    Same cell changed on both sides since the last sync -> conflict, neither side overwritten.
App -> Excel (flush): the write queue (pending_writeback) is written through Excel itself:
    - changes made within `batch_quiet_s` (10 s) are saved together;
    - the file must be exactly what was last read (else read it first);
    - backup copy before every write (never deleted);
    - locked / busy -> retry at 1, 2, 5, 10, 15, 30 min, then status "error" + one PMO email;
      the queue stays and the next file change tries again;
    - a OneDrive conflict copy ("UAT Planning-<PC>.xlsx") stops all writes until the PMO clears it.
Only one process writes at a time (lock file in the data folder).
"""

from __future__ import annotations

import datetime as dt
import json
import msvcrt
import os
import shutil
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path

from ..alerts.outbox import alert_pmo
from ..domain.models import CHANNEL_EXCEL_SYNC, CHANNEL_RESTORE
from ..excel.reader import copy_shared, sha256_file
from ..excel.sheets import DB_TO_SHEET_KEY, SHEETS, TEST_PROGRESS, UAT_PLAN, WRITABLE, with_overrides
from ..excel.writer import (ROW_FIELD, CellWrite, ComWorkbook, WorkbookBusy, WorkbookLocked, WriterError,
                            excel_windows_for, write_workbook)
from .importer import S_PAUSED, ImportResult, run_import
from .values import derive_excel_value

SYNC_EDITOR = "system (sync)"
S_BLOCKED = "sync.blocked"            # setting: reason text when writes are stopped
S_ERROR_ALERTED = "sync.error_alerted"


class SyncBusy(Exception):
    pass


@dataclass
class FlushResult:
    result: str                       # nothing | waiting | written | retry | error | blocked | busy
    message: str = ""
    written: int = 0
    failed: list[str] = field(default_factory=list)
    backup: str | None = None


def _iso(t: dt.datetime) -> str:
    return t.astimezone().isoformat(timespec="seconds")


@contextmanager
def file_lock(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    f = open(path, "a+b")
    try:
        try:
            f.seek(0)
            msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            raise SyncBusy("Another sync is running")
        try:
            yield
        finally:
            f.seek(0)
            msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
    finally:
        f.close()


class SyncService:
    def __init__(self, cfg, repo, *, opener=ComWorkbook, sender=None, clock=dt.datetime.now):
        self.cfg, self.repo, self.opener, self.sender, self.clock = cfg, repo, opener, sender, clock
        self.path = Path(cfg.excel_path)
        self.lock_path = cfg.data_dir / "sync.lock"

    # ------------------------------------------------------------------ helpers
    def _now(self) -> str:
        return _iso(self.clock())

    def _log(self, kind, result, message=""):
        with self.repo.transaction():
            self.repo.log_sync(ts=self._now(), kind=kind, result=result, message=message[:2000])

    def _alert(self, title, detail):
        if self.sender is None:
            return
        alert_pmo(self.cfg, self.repo, self.sender, title, detail, send_date=self._now()[:10],
                  now=self._now(), editor=SYNC_EDITOR)

    def conflict_copies(self) -> list[Path]:
        """OneDrive keeps both versions as '<name>-<PC name>.xlsx' when it cannot merge."""
        pc = os.environ.get("COMPUTERNAME", "").lower()
        if not pc or not self.path.parent.exists():
            return []
        stem, ext = self.path.stem.lower(), self.path.suffix.lower()
        return sorted(p for p in self.path.parent.iterdir()
                      if p.suffix.lower() == ext and p.stem.lower().startswith(f"{stem}-{pc}"))

    def status(self) -> dict:
        pend = self.repo.all_pending()
        waiting = [p for p in pend if p["status"] == "pending"]
        failed = [p for p in pend if p["status"] == "failed" or (p["attempts"] and not p["next_attempt_at"])]
        blocked = self.repo.get_setting(S_BLOCKED)
        conflicts = self.repo.open_conflicts()
        paused = self.repo.get_setting(S_PAUSED)
        state = ("paused" if paused else "blocked" if blocked else "error" if failed else "pending" if waiting
                 else "synced")
        last = self.repo.last_ok_import()
        return {"state": state, "pending": len(waiting), "failed": len(failed),
                "conflicts": len(conflicts), "blocked": blocked,
                "last_read": last["finished_at"] if last else None,
                "last_error": next((p["last_error"] for p in failed if p["last_error"]), None)}

    # ------------------------------------------------------------------ Excel -> app
    def file_changed(self) -> bool:
        last = self.repo.last_ok_import()
        if not self.path.exists():
            return False
        return last is None or last["source_sha256"] != sha256_file(self.path)

    def pull(self, force: bool = False) -> ImportResult | None:
        if self.repo.get_setting(S_PAUSED):           # project data cleared: wait for "Import from Excel"
            return None
        if not force and not self.file_changed():
            return None
        try:
            res = run_import(self.cfg, self.repo, channel=CHANNEL_EXCEL_SYNC)
        except Exception as exc:
            self._log("pull", "failed", str(exc))
            raise
        changed = sum(s.updated + s.created + s.removed for s in res.sheets.values())
        self._log("pull", "ok", f"{changed} item(s) changed, {len(res.conflicts_added)} new conflict(s)")
        if res.conflicts_added:
            from .resolve import describe_conflict
            lines = [describe_conflict(self.repo.get_conflict(c)) for c in res.conflicts_added]
            self._alert(f"{len(lines)} sync conflict(s) need your decision",
                        "Neither side was overwritten. Open the Conflict Center.\n" + "\n".join(lines))
        return res

    def retry_now(self) -> None:
        """A file change: give queued writes that were waiting for the file another go."""
        with self.repo.transaction():
            for p in self.repo.all_pending():
                if p["status"] == "pending" and p["attempts"]:
                    self.repo.set_pending_state(p["sheet"], p["item_id"], p["field"], status="pending",
                                                first_failed_at=p["first_failed_at"],
                                                next_attempt_at=self._now(), last_error=p["last_error"])

    # ------------------------------------------------------------------ app -> Excel
    def _due(self) -> list[dict]:
        now = self._now()
        return [p for p in self.repo.all_pending() if p["status"] == "pending"
                and (p["attempts"] == 0 or (p["next_attempt_at"] and p["next_attempt_at"] <= now))]

    def backup(self, reason: str) -> Path:
        self.cfg.backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = self.clock().strftime("%Y%m%d-%H%M%S")
        dest = self.cfg.backup_dir / f"{self.path.stem}_{stamp}{self.path.suffix}"
        n = 1
        while dest.exists():
            n += 1
            dest = self.cfg.backup_dir / f"{self.path.stem}_{stamp}_{n}{self.path.suffix}"
        copy_shared(self.path, dest)
        with self.repo.transaction():
            self.repo.insert_backup({"ts": self._now(), "path": str(dest), "sha256": sha256_file(dest),
                                     "size_bytes": dest.stat().st_size, "reason": reason})
        return dest

    def flush(self, force: bool = False) -> FlushResult:
        try:
            with file_lock(self.lock_path):
                return self._flush(force)
        except SyncBusy as exc:
            return FlushResult("busy", str(exc))

    def _flush(self, force: bool) -> FlushResult:
        if self.repo.get_setting(S_PAUSED):           # rows would be placed without knowing the file
            return FlushResult("blocked", "Excel sync paused: project data was cleared; use Import from Excel")
        blocked = self.repo.get_setting(S_BLOCKED)
        if blocked:
            return FlushResult("blocked", blocked)
        copies = self.conflict_copies()
        if copies:
            msg = ("OneDrive made a conflict copy, so the main file may be out of date: "
                   + ", ".join(p.name for p in copies) + ". Writes stopped. Merge or remove the copy, "
                   "then run 'sync unblock'.")
            with self.repo.transaction():
                self.repo.set_setting(S_BLOCKED, msg)
            self._log("block", "blocked", msg)
            self._alert("Excel sync stopped: OneDrive conflict copy", msg)
            return FlushResult("blocked", msg)

        with self.repo.transaction():                 # 'writing' left by a run that stopped mid-way
            for p in self.repo.all_pending():
                if p["status"] == "writing":
                    self.repo.set_pending_state(p["sheet"], p["item_id"], p["field"], status="pending",
                                                attempts=p["attempts"], last_error="previous write interrupted")
        due = self._due()
        if not due:
            return FlushResult("nothing")
        newest = max(p["created_at"] for p in due)
        quiet = (self.clock().astimezone() - dt.datetime.fromisoformat(newest)).total_seconds()
        if not force and quiet < self.cfg.sync.batch_quiet_s:
            return FlushResult("waiting", f"batching ({quiet:.0f}s since last change)")

        try:
            self.pull()                               # file must match what the rows refer to
        except Exception as exc:
            return self._retry(due, f"Could not read Excel before writing: {exc}", None)
        due = self._due()
        if not due:
            return FlushResult("nothing", "queue cleared or in conflict after reading Excel")

        rows, writes = {}, []
        defs = {name: with_overrides(sd, self.cfg.header_overrides) for name, sd in SHEETS.items()}
        for name in defs:
            items = self.repo.load_items(name)
            rows[name] = {i: r["excel_row"] for i, r in items.items() if r.get("active") and r.get("excel_row")}
        new_rows = {(p["sheet"], p["item_id"]) for p in self.repo.all_pending() if p["field"] == ROW_FIELD}
        new_groups: set[str] = set()                  # a new UAT Plan group gets its header row once per write
        held: set[tuple[str, str]] = set()
        for p in due:
            rec = self.repo.get_item(p["sheet"], p["item_id"])
            if p["field"] == ROW_FIELD:
                if rec is None or not rec.get("active"):
                    self._fail(p, "Item no longer exists in the app")
                    continue
                w = self._row_write(p["sheet"], rec, rows)
                if json.loads(w.excel_value).get("group_header"):
                    if rec.get("group_code") in new_groups:     # 2nd task of the same new group: next round,
                        held.add((p["sheet"], p["item_id"]))   # once the group row exists in Excel
                        continue
                    new_groups.add(rec.get("group_code"))
                writes.append(w)
                continue
            if (p["sheet"], p["item_id"]) in new_rows:
                continue                              # the new row carries this value
            row = rows.get(p["sheet"], {}).get(p["item_id"])
            if rec is None or row is None:
                self._fail(p, "Item is no longer in the Excel file")
                continue
            ev = p["excel_value"] or derive_excel_value(p["field"], rec.get(p["field"]))
            writes.append(CellWrite(p["sheet"], p["item_id"], p["field"], row, ev))

        if not writes and held:
            return FlushResult("nothing", "waiting for a new group row")
        backup = self.backup(f"before writing {len(writes)} change(s)")
        try:
            report = write_workbook(self.opener, self.path, defs, writes, rows)
        except (WorkbookLocked, WorkbookBusy) as exc:
            return self._retry(due, str(exc), backup)
        except WriterError as exc:
            return self._retry(due, f"Excel write failed: {exc}", backup)

        now = self._now()
        with self.repo.transaction():
            for key3, err in report.failed:
                p = next(p for p in due if (p["sheet"], p["item_id"], p["field"]) == key3)
                self._fail(p, err, in_tx=True)
            for sheet, item_id, fld in report.written:
                p = next(p for p in due if (p["sheet"], p["item_id"], p["field"]) == (sheet, item_id, fld))
                shown = json.loads(p["excel_value"]).get("value") if p["excel_value"] else p["value"]
                self.repo.append_audit({"ts": now, "editor": SYNC_EDITOR, "channel": CHANNEL_EXCEL_SYNC,
                                        "sheet": sheet, "item_id": item_id, "field": fld, "old_value": "",
                                        "new_value": "" if shown is None else str(shown),
                                        "reason": "Written to Excel"})
            self.repo.set_setting(S_ERROR_ALERTED, None)
        msg = (f"{len(report.written)} cell(s) written, {report.ids_written} App ID(s), "
               f"{len(report.failed)} failed" + (f"; App ID column added to {', '.join(report.id_column_added)}"
                                                 if report.id_column_added else ""))
        self._log("write", "ok" if not report.failed else "partial", msg)
        if report.failed:
            self._alert("Some app changes could not be written to Excel",
                        "\n".join(f"{s} {i} {f}: {e}" for (s, i, f), e in report.failed))

        written = set(report.written)
        with self.repo.transaction():                 # 'writing' = our own change, not an Excel edit
            for sheet, item_id, fld in written:
                self.repo.set_pending_state(sheet, item_id, fld, status="writing")
        self.pull(force=True)                         # read back: clears the rows Excel now shows
        with self.repo.transaction():
            for p in self.repo.all_pending():
                if p["status"] == "writing":
                    self._fail(p, "Written, but Excel did not read back the same value", in_tx=True)
        return FlushResult("written", msg, len(report.written),
                           [f"{s} {i} {f}: {e}" for (s, i, f), e in report.failed], str(backup))

    def _row_write(self, sheet, rec, rows) -> CellWrite:
        """A new item's row: UAT Plan -> inserted after the last task of its group; Defect -> the next
        empty row after the last defect."""
        items = self.repo.load_items(sheet)
        placed = [r for r in items.values() if r.get("active") and r.get("excel_row")]
        spec_extra = {}
        if sheet == UAT_PLAN:
            same = [r["excel_row"] for r in placed if r.get("group_code") == rec.get("group_code")]
            anchor = max(same or [r["excel_row"] for r in placed] or [1])
            if not same:                              # new group (multi-project, approved Q13): header row first
                spec_extra["group_header"] = {"group": rec.get("group_code") or "",
                                              "title": rec.get("group_title") or ""}
        else:
            anchor = max([r["excel_row"] for r in placed] or [1])
        if sheet == TEST_PROGRESS:
            spec_extra["formulas"] = ["pct_progress"]  # approved Q14: =Executed/Total like the old rows
        fields = set(WRITABLE[sheet]) | ({"task_no"} if sheet == UAT_PLAN else {"no"})
        values = {DB_TO_SHEET_KEY.get(f, f): derive_excel_value(f, rec.get(f))
                  for f in sorted(fields) if rec.get(f) not in (None, "")}
        # UAT Plan + Test Progress rows are inserted (Test Progress: right after the last application, above
        # a Grand Total row if there is one); Defect List keeps pre-numbered empty rows -> next empty row.
        spec = {"kind": "row", "insert": sheet in (UAT_PLAN, TEST_PROGRESS), "values": values, **spec_extra}
        return CellWrite(sheet, rec["id"], ROW_FIELD, anchor, json.dumps(spec, ensure_ascii=False))

    def init_ids(self) -> FlushResult:
        """Add the hidden App ID column / fill missing IDs without any other change."""
        with file_lock(self.lock_path):
            if self.repo.get_setting(S_BLOCKED):
                return FlushResult("blocked", self.repo.get_setting(S_BLOCKED))
            self.pull()
            defs = {name: with_overrides(sd, self.cfg.header_overrides) for name, sd in SHEETS.items()}
            rows = {name: {i: r["excel_row"] for i, r in self.repo.load_items(name).items()
                          if r.get("active") and r.get("excel_row")}
                    for name in defs}
            backup = self.backup("before writing App IDs")
            try:
                report = write_workbook(self.opener, self.path, defs, [], rows)
            except WriterError as exc:
                self._log("write", "failed", str(exc))
                return FlushResult("error", str(exc), backup=str(backup))
            msg = (f"{report.ids_written} App ID(s) written"
                   + (f"; column added to {', '.join(report.id_column_added)}" if report.id_column_added else ""))
            self._log("write", "ok", msg)
            self.pull(force=True)
            return FlushResult("written", msg, backup=str(backup))

    def _fail(self, p, err, in_tx=False):
        def go():
            self.repo.set_pending_state(p["sheet"], p["item_id"], p["field"], status="failed",
                                        attempts=p["attempts"], last_error=err)
        if in_tx:
            go()
        else:
            with self.repo.transaction():
                go()

    def _retry(self, due, err, backup) -> FlushResult:
        offsets = self.cfg.sync.retry_minutes
        now = self.clock().astimezone()
        gave_up = False
        with self.repo.transaction():
            for p in due:
                attempts = p["attempts"] + 1
                first = p["first_failed_at"] or _iso(now)
                if attempts <= len(offsets):
                    nxt = _iso(dt.datetime.fromisoformat(first) + dt.timedelta(minutes=offsets[attempts - 1]))
                else:
                    nxt, gave_up = None, True          # wait for the next file change
                self.repo.set_pending_state(p["sheet"], p["item_id"], p["field"], status="pending",
                                            attempts=attempts, first_failed_at=first, next_attempt_at=nxt,
                                            last_error=err)
        self._log("write", "retry", err)
        backup = backup if backup is None else str(backup)
        if gave_up:
            if not self.repo.get_setting(S_ERROR_ALERTED):       # one email per error episode
                with self.repo.transaction():
                    self.repo.set_setting(S_ERROR_ALERTED, _iso(now))
                self._alert("Excel sync error: changes not written for 30 minutes",
                            f"{err}\n{len(due)} change(s) are still queued. They are retried when the file changes.")
            return FlushResult("error", err, backup=backup)
        return FlushResult("retry", err, backup=backup)

    def unblock(self) -> None:
        with self.repo.transaction():
            self.repo.set_setting(S_BLOCKED, None)
        self._log("block", "cleared", "PMO cleared the block")

    # ------------------------------------------------------------------ restore
    def restore(self, backup_id: int, actor) -> Path:
        from .auth import RESTORE, require
        require(self.repo, actor, RESTORE, item_id=f"backup #{backup_id}")
        editor = actor.editor
        b = self.repo.get_backup(backup_id)
        if b is None:
            raise ValueError(f"No backup #{backup_id}")
        src = Path(b["path"])
        if not src.exists() or sha256_file(src) != b["sha256"]:
            raise ValueError(f"Backup file missing or changed: {src}")
        if excel_windows_for(self.path):
            raise ValueError("The workbook is open in Excel on this PC: close it first, then restore")
        with file_lock(self.lock_path):
            current = self.backup(f"before restoring backup #{backup_id}")
            tmp = self.path.with_name(self.path.name + ".restoring")
            shutil.copy2(src, tmp)
            os.replace(tmp, self.path)
            with self.repo.transaction():
                self.repo.append_audit({"ts": self._now(), "editor": editor, "channel": CHANNEL_RESTORE,
                                        "sheet": None, "item_id": None, "field": "(workbook)",
                                        "old_value": current.name, "new_value": src.name,
                                        "reason": f"Backup #{backup_id} restored"})
            self._log("restore", "ok", f"backup #{backup_id} restored; previous file saved as {current.name}")
        self.pull(force=True)
        return current

    # ------------------------------------------------------------------ watch loop
    def watch(self, stop: threading.Event, on_event=print) -> None:
        from watchdog.events import FileSystemEventHandler
        from watchdog.observers import Observer

        changed_at = {"t": None}
        target = self.path.name.lower()

        class Handler(FileSystemEventHandler):
            def on_any_event(self, event):
                names = {Path(getattr(event, "src_path", "") or "").name.lower(),
                         Path(getattr(event, "dest_path", "") or "").name.lower()}
                if target in names:
                    changed_at["t"] = time.monotonic()

        obs = Observer()
        obs.schedule(Handler(), str(self.path.parent), recursive=False)
        obs.start()
        last_poll = time.monotonic()
        on_event(f"Watching {self.path}")
        try:
            while not stop.is_set():
                t = time.monotonic()
                try:
                    if changed_at["t"] and t - changed_at["t"] >= self.cfg.sync.debounce_s:
                        changed_at["t"] = None
                        res = self.pull()
                        self.retry_now()
                        if res:
                            on_event(f"{self._now()} read Excel: {self.repo.sync_events(1)[0]['message']}")
                    elif t - last_poll >= self.cfg.sync.poll_s:
                        last_poll = t
                        if self.pull():
                            on_event(f"{self._now()} read Excel (periodic check)")
                    fr = self.flush()
                    if fr.result not in ("nothing", "waiting", "busy"):
                        on_event(f"{self._now()} write: {fr.result} {fr.message}")
                except Exception as exc:              # keep watching; the error is logged
                    self._log("error", "failed", str(exc))
                    on_event(f"{self._now()} ERROR {exc}")
                stop.wait(1.0)
        finally:
            obs.stop()
            obs.join(5)
