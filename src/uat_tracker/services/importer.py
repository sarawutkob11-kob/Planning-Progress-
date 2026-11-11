"""Import: Excel snapshot -> normalised items -> DB, with audit log, review queue and flags.

Idempotent: importing an unchanged workbook writes no item rows, audit entries, reviews or flags.
"""

from __future__ import annotations

import logging
import getpass
import html
import json
import re
import zipfile
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

from ..config import Config
from ..domain.models import (CHANNEL_EXCEL_SYNC, CHANNEL_IMPORT, DATE_FIELDS, EDITED_IN_EXCEL,
                             HISTORY_SUFFIX, NOTE_SUFFIX, Item)
from ..excel.reader import read_sheets, take_snapshot
from ..excel.sheets import APP_ID_KEY, SHEETS, TEST_PROGRESS, with_overrides
from ..parsing.dates import DateWindow, fmt
from ..parsing.text import cell_text
from .binding import alt_key_json, bind, fingerprints
from .normalize import NORMALIZERS
from .values import canon

log = logging.getLogger(__name__)

REASON_IMPORT = "Excel import"
REASON_SYNC = "Changed in Excel (sync)"


@dataclass
class SheetStats:
    rows: int = 0
    created: int = 0
    updated: int = 0
    unchanged: int = 0
    removed: int = 0
    restored: int = 0
    audit_added: int = 0
    reviews_added: int = 0
    reviews_superseded: int = 0
    flags_added: int = 0
    flags_cleared: int = 0
    conflicts_added: list[int] = field(default_factory=list)
    missing_optional_columns: list[str] = field(default_factory=list)


@dataclass
class ImportResult:
    run_id: int
    snapshot: str
    sha256: str
    sheets: dict[str, SheetStats]

    @property
    def audit_added(self) -> int:
        return sum(s.audit_added for s in self.sheets.values())

    @property
    def conflicts_added(self) -> list[int]:
        return [c for s in self.sheets.values() for c in s.conflicts_added]


def display(field_name: str, value) -> str:
    """Value as shown to people (audit log, summaries). Dates -> DD-Mon-YYYY."""
    if value is None or value == "" or value == ():
        return ""
    if field_name.endswith(HISTORY_SUFFIX):
        return " => ".join(fmt(v) for v in value)
    if field_name in DATE_FIELDS:
        return fmt(value)
    return str(value)


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def last_saved_by(xlsx: Path) -> str | None:
    """docProps/core.xml lastModifiedBy: who saved the file last (a hint, not per cell)."""
    try:
        with zipfile.ZipFile(xlsx) as z:
            core = z.read("docProps/core.xml").decode("utf-8", "replace")
    except (KeyError, OSError, zipfile.BadZipFile):
        return None
    m = re.search(r"<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>", core, re.S)
    return html.unescape(m.group(1)).strip() or None if m else None


S_PAUSED = "sync.paused"        # multi-project: set by "Clear project data" until "Import from Excel"


class ImportPaused(Exception):
    """The project was cleared: reading Excel waits for the user's "Import from Excel" (approved Q10)."""


def run_import(cfg: Config, repo, editor: str | None = None, source: Path | None = None,
               channel: str = CHANNEL_IMPORT, force: bool = False) -> ImportResult:
    """channel=import: manual import by `editor`. channel=excel_sync: change picked up by the sync;
    editor becomes "Edited in Excel (last saved by <name>)". force=True: also when paused (and un-pause)."""
    if repo.get_setting(S_PAUSED):
        if not force:
            raise ImportPaused("Reading Excel is paused for this project (data was cleared). "
                               "Use 'Import from Excel' to read the file again.")
        with repo.transaction():
            repo.set_setting(S_PAUSED, None)
    editor = editor or getpass.getuser()
    source = Path(source or cfg.excel_path)
    window = DateWindow(cfg.window_start, cfg.window_end)
    started = _now()
    mtime = (datetime.fromtimestamp(source.stat().st_mtime).astimezone().isoformat(timespec="seconds")
             if source.exists() else None)
    run_id = repo.start_run(started_at=started, editor=editor, source_path=str(source), source_mtime=mtime)
    snapshot, sha = None, None
    try:
        snapshot, sha = take_snapshot(source, cfg.snapshot_dir)
        if channel == CHANNEL_EXCEL_SYNC:
            who = last_saved_by(snapshot)
            editor = f"{EDITED_IN_EXCEL} (last saved by {who})" if who else EDITED_IN_EXCEL
        defs = [with_overrides(sd, cfg.header_overrides) for sd in SHEETS.values()]
        data = read_sheets(snapshot, defs)
        normalized = {name: NORMALIZERS[name](sd.rows, window) for name, sd in data.items()}
        for name, items in normalized.items():          # hidden App ID column (Part 3)
            ids = {r.excel_row: cell_text(r.values.get(APP_ID_KEY)) for r in data[name].rows}
            for it in items:
                it.app_id = ids.get(it.excel_row) or None

        stats: dict[str, SheetStats] = {}
        with repo.transaction():
            for name, items in normalized.items():
                st = SheetStats(rows=len(items),
                                missing_optional_columns=data[name].headers.missing_optional)
                _sync_sheet(repo, name, items, st, editor, run_id, _now(), channel, mtime)
                stats[name] = st
    except Exception as exc:
        log.error("Import failed", exc_info=True)
        repo.finish_run(run_id, finished_at=_now(), result="failed", message=str(exc),
                        source_sha256=sha, snapshot_path=str(snapshot) if snapshot else None)
        raise
    repo.finish_run(run_id, finished_at=_now(), result="ok", source_sha256=sha,
                    snapshot_path=str(snapshot),
                    stats_json=json.dumps({k: asdict(v) for k, v in stats.items()}))
    return ImportResult(run_id, str(snapshot), sha, stats)


def _keep_db_value(item: Item, old: dict, k: str) -> None:
    """Freeze field k at its DB value for this import (app value pending or conflict open)."""
    item.fields[k] = old.get(k)
    for suffix in (HISTORY_SUFFIX, NOTE_SUFFIX):
        if k + suffix in item.fields:
            item.fields[k + suffix] = old.get(k + suffix, () if suffix == HISTORY_SUFFIX else None)


def _pending_matches(p: dict, item: Item, k: str) -> bool:
    if canon(k, item.fields.get(k)) != (p["value"] or ""):
        return False
    ev = json.loads(p["excel_value"]) if p.get("excel_value") else None
    if ev and ev.get("hist") is not None and k + HISTORY_SUFFIX in item.fields:
        return list(item.fields[k + HISTORY_SUFFIX]) == list(ev["hist"])
    return True


def _apply_sync_rules(repo, sheet, item_id, item: Item, old: dict, baseline, pending, conflicts,
                      st: "SheetStats", *, now, excel_time, audit) -> None:
    """Part 3 rules for one existing item, before its Excel values are applied to the DB:
    - app value waiting for Excel (write queue): Excel already shows it -> done; Excel unchanged since
      the last sync -> keep the app value; Excel changed too -> CONFLICT, overwrite neither side;
    - open conflict -> field frozen on both sides until the PMO resolves it;
    - Excel edit making Actual Finish > Planned Finish -> CONFLICT (spec §7)."""
    for k in [k for (pid, k) in pending if pid == item_id]:
        p = pending[(item_id, k)]
        if _pending_matches(p, item, k):
            repo.clear_pending(sheet, item_id, k)
            continue
        excel_changed = baseline.get((item_id, k)) != item.raw.get(k)
        if (item_id, k) not in conflicts and excel_changed and p.get("status") != "writing":
            cid = repo.insert_conflict({
                "created_at": now, "sheet": sheet, "item_id": item_id, "field": k, "kind": "both_changed",
                "excel_raw": item.raw.get(k), "excel_value": canon(k, item.fields.get(k)),
                "excel_time": excel_time, "app_value": p["value"], "app_time": p["created_at"],
                "app_editor": p.get("editor")})
            conflicts[(item_id, k)] = {"id": cid}
            repo.set_pending_state(sheet, item_id, k, status="conflict")
            audit(item_id, k, display(k, old.get(k)), display(k, item.fields.get(k)),
                  f"CONFLICT #{cid}: changed in Excel and in the app; neither side overwritten")
            st.conflicts_added.append(cid)
        _keep_db_value(item, old, k)
    for (cid_item, k) in list(conflicts):
        if cid_item == item_id and (item_id, k) not in pending:
            _keep_db_value(item, old, k)
    if sheet == TEST_PROGRESS and (item_id, "actual_finish") not in conflicts:
        af, pf = item.fields.get("actual_finish"), item.fields.get("planned_finish")
        if (af and pf and af > pf and af != old.get("actual_finish")
                and baseline.get((item_id, "actual_finish")) != item.raw.get("actual_finish")):
            cid = repo.insert_conflict({
                "created_at": now, "sheet": sheet, "item_id": item_id, "field": "actual_finish",
                "kind": "actual_after_planned", "excel_raw": item.raw.get("actual_finish"),
                "excel_value": af, "excel_time": excel_time, "app_value": canon("actual_finish", old.get("actual_finish")),
                "app_time": old.get("updated_at"), "app_editor": None})
            conflicts[(item_id, "actual_finish")] = {"id": cid}
            audit(item_id, "actual_finish", display("actual_finish", old.get("actual_finish")),
                  display("actual_finish", af),
                  f"CONFLICT #{cid}: Actual Finish {fmt(af)} is after Planned Finish {fmt(pf)} (edited in Excel)")
            st.conflicts_added.append(cid)
            _keep_db_value(item, old, "actual_finish")


def _sync_sheet(repo, sheet: str, items: list[Item], st: SheetStats, editor: str,
                run_id: int, now: str, channel: str = CHANNEL_IMPORT, excel_time: str | None = None) -> None:
    existing = repo.load_items(sheet)
    baseline = repo.load_baseline(sheet)
    bound = bind(items, existing)
    fps = fingerprints(items)
    pending = repo.pending(sheet)
    conflicts = {(c["item_id"], c["field"]): c for c in repo.open_conflicts(sheet)}

    default_reason = REASON_IMPORT if channel == CHANNEL_IMPORT else REASON_SYNC

    def audit(item_id, field_name, old, new, reason=None):
        repo.append_audit({"ts": now, "editor": editor, "channel": channel, "sheet": sheet,
                           "item_id": item_id, "field": field_name, "old_value": old,
                           "new_value": new, "reason": reason or default_reason, "import_run_id": run_id})
        st.audit_added += 1

    seen: set[str] = set()
    for item, (item_id, is_new), fp in zip(items, bound, fps):
        seen.add(item_id)
        meta = {"excel_row": item.excel_row, "fingerprint": fp, "alt_key": alt_key_json(item), "now": now}
        if is_new:
            repo.insert_item(sheet, item_id, item.fields, meta)
            snapshot = {k: display(k, v) for k, v in item.fields.items() if display(k, v)}
            audit(item_id, "*", "", json.dumps(snapshot, ensure_ascii=False), "Created by Excel import")
            st.created += 1
        else:
            old = existing[item_id]
            _apply_sync_rules(repo, sheet, item_id, item, old, baseline, pending, conflicts, st,
                              now=now, excel_time=excel_time, audit=audit)
            if not old.get("active"):
                repo.set_active(sheet, item_id, True, now)
                audit(item_id, "(row)", "removed", "active", "Row found again in Excel")
                st.restored += 1
            changed = [k for k, v in item.fields.items() if _norm(old.get(k)) != _norm(v)]
            for k in changed:
                audit(item_id, k, display(k, old.get(k)), display(k, item.fields[k]))
            meta_changed = (old.get("excel_row"), old.get("fingerprint"), old.get("alt_key")) != (
                item.excel_row, fp, meta["alt_key"])
            if changed or meta_changed:
                repo.update_item(sheet, item_id, item.fields,
                                 {**meta, "data_changed": bool(changed), "previous": old})
            if changed:
                st.updated += 1
            else:
                st.unchanged += 1
        for k, raw in item.raw.items():
            if baseline.get((item_id, k)) != raw:
                repo.upsert_baseline(sheet, item_id, k, raw, now)

    for item_id, rec in existing.items():
        if (item_id, "(row)") in pending:          # created in the app, Excel row not written yet
            continue
        if item_id not in seen and rec.get("active"):
            repo.set_active(sheet, item_id, False, now)
            audit(item_id, "(row)", "active", "removed", "Row not found in Excel (kept, marked inactive)")
            st.removed += 1

    _sync_reviews(repo, sheet, items, bound, st, run_id, now)
    _sync_flags(repo, sheet, items, bound, st, now)


def _norm(v):
    if isinstance(v, list):
        return tuple(v)
    if v == "":
        return None
    return v


def _sync_reviews(repo, sheet, items, bound, st, run_id, now) -> None:
    current = {}
    for item, (item_id, _) in zip(items, bound):
        for p in item.problems:
            current[(item_id, p.field, p.raw, p.code, p.token)] = item
    existing = repo.load_reviews(sheet)
    known = {(r["item_id"], r["field"], r["raw_value"], r["problem"], r["token"]) for r in existing}
    for key, item in current.items():
        if key in known:
            continue                      # open: still pending; resolved: PMO already decided
        item_id, fld, raw, code, token = key
        repo.insert_review({"created_at": now, "sheet": sheet, "item_id": item_id,
                            "excel_row": item.excel_row, "field": fld, "raw_value": raw,
                            "token": token, "problem": code, "import_run_id": run_id})
        st.reviews_added += 1
    for r in existing:
        key = (r["item_id"], r["field"], r["raw_value"], r["problem"], r["token"])
        if r["status"] == "open" and key not in current:
            repo.set_review_status(r["id"], "superseded")   # value changed or fixed in Excel
            st.reviews_superseded += 1


def _sync_flags(repo, sheet, items, bound, st, now) -> None:
    current = {}
    for item, (item_id, _) in zip(items, bound):
        for f in item.flags:
            current[(item_id, f.code)] = f.detail
    for (item_id, code), detail in current.items():
        if repo.upsert_flag(sheet, item_id, code, detail, now):
            st.flags_added += 1
    for key, row in repo.load_flags(sheet).items():
        if row["active"] and key not in current:
            repo.deactivate_flag(row["id"], now)
            st.flags_cleared += 1
