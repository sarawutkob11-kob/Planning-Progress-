"""Test-result uploads (R4). Files are copied under <data_dir>/attachments/<item-id>/ and never deleted.

On upload: TC/Result Complete = Complete in the DB + audit log. The value goes into the
write queue (pending_writeback) and the sync writes it to Excel.
"""

from __future__ import annotations

import hashlib
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from ..alerts.rules import is_complete_text
from ..domain.models import CHANNEL_APP
from ..excel.sheets import TEST_PROGRESS
from .auth import UPLOAD, Actor, require
from .values import derive_excel_value

FIELD = "tc_result_complete"
COMPLETE = "Complete"


class UploadError(Exception):
    pass


@dataclass
class UploadResult:
    attachment_id: int
    stored_path: Path
    sha256: str
    status_changed: bool


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def upload(cfg, repo, item_id: str, file_path: str | Path, *, actor: Actor,
           now: str | None = None) -> UploadResult:
    editor = actor.editor
    now = now or _now()
    src = Path(file_path)
    rec = repo.get_item(TEST_PROGRESS, item_id)
    if rec is None or not rec.get("active"):
        raise UploadError(f"No active Test Progress item {item_id!r}")
    require(repo, actor, UPLOAD, sheet=TEST_PROGRESS, item_id=item_id, rec=rec)
    if not src.is_file():
        raise UploadError(f"File not found: {src}")
    ext = src.suffix.lower().lstrip(".")
    if ext not in cfg.alerts.attachment_types:
        raise UploadError(f"File type .{ext} not allowed. Allowed: {', '.join(cfg.alerts.attachment_types)}")
    size = src.stat().st_size
    limit = cfg.alerts.attachment_max_mb * 1024 * 1024
    if size > limit:
        raise UploadError(f"File is {size / 1048576:.1f} MB; the limit is {cfg.alerts.attachment_max_mb} MB")
    if size == 0:
        raise UploadError("File is empty")

    folder = cfg.attachment_dir / re.sub(r"[^A-Za-z0-9_.-]+", "_", item_id)
    folder.mkdir(parents=True, exist_ok=True)
    stamp = now[:19].replace(":", "").replace("-", "").replace("T", "-")
    dest = folder / f"{stamp}_{src.name}"
    n = 1
    while dest.exists():
        n += 1
        dest = folder / f"{stamp}_{n}_{src.name}"
    shutil.copy2(src, dest)
    sha = _sha256(dest)
    if sha != _sha256(src):
        dest.unlink()
        raise UploadError("Copy verification failed (hash mismatch); nothing saved")

    with repo.transaction():
        aid = repo.insert_attachment({
            "sheet": TEST_PROGRESS, "item_id": item_id, "original_name": src.name,
            "stored_path": str(dest), "sha256": sha, "size_bytes": size,
            "uploaded_by": editor, "uploaded_at": now})
        repo.append_audit({"ts": now, "editor": editor, "channel": CHANNEL_APP, "sheet": TEST_PROGRESS,
                           "item_id": item_id, "field": "attachment", "old_value": "",
                           "new_value": f"{src.name} ({size} bytes, sha256 {sha[:12]})",
                           "reason": "Test-result document uploaded"})
        changed = not is_complete_text(rec.get(FIELD))
        if changed:
            repo.set_field(TEST_PROGRESS, item_id, FIELD, COMPLETE, now)
            repo.append_audit({"ts": now, "editor": editor, "channel": CHANNEL_APP, "sheet": TEST_PROGRESS,
                               "item_id": item_id, "field": FIELD, "old_value": rec.get(FIELD) or "",
                               "new_value": COMPLETE, "reason": f"Set by upload of {src.name}"})
            repo.upsert_pending(TEST_PROGRESS, item_id, FIELD, value=COMPLETE,
                                excel_value=derive_excel_value(FIELD, COMPLETE), now=now,
                                reason=f"Set by upload of {src.name}", editor=editor, channel=CHANNEL_APP)
    return UploadResult(aid, dest, sha, changed)
