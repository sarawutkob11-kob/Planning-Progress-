"""Test-result uploads (R4). Files are copied to <attachment folder>/<App name>/<original file name> and
never deleted or overwritten: a name already taken gets " (2)", " (3)" ... before the extension.
The App name is made safe for Windows; an empty App name falls back to the item id.
The attachment folder is the project's chosen result files folder (Settings), else <data_dir>/attachments.
A chosen folder that is missing or fails during the copy blocks the upload (no fallback to local).
upload_many() uploads several files one by one; a failed file does not stop the others.

On upload: TC/Result Complete = Complete in the DB + audit log. The value goes into the
write queue (pending_writeback) and the sync writes it to Excel.
"""

from __future__ import annotations

import hashlib
import logging
import re
import shutil
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from ..alerts.rules import is_complete_text
from ..domain.models import CHANNEL_APP
from ..excel.sheets import TEST_PROGRESS
from .auth import UPLOAD, Actor, PermissionDenied, require
from .dashboard import first_line
from .values import derive_excel_value

log = logging.getLogger(__name__)

FIELD = "tc_result_complete"
COMPLETE = "Complete"

_UNSAFE_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED_NAMES = frozenset({"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$",
                             *(f"COM{c}" for c in "123456789¹²³"), *(f"LPT{c}" for c in "123456789¹²³")})
_MAX_FOLDER_NAME = 100          # keeps <folder>/<App name>/<file> well inside Windows path limits


class UploadError(Exception):
    pass


class AttachmentFolderUnavailable(UploadError):
    """The project's chosen result files folder is missing or can't be written. Nothing was saved."""

    def __init__(self, path: Path | str):
        super().__init__(f"Can't save to the result files folder: {path}. Nothing was saved.")
        self.path = str(path)


@dataclass
class UploadResult:
    attachment_id: int
    stored_path: Path
    sha256: str
    status_changed: bool


@dataclass
class BatchUploadResult:
    """upload_many(): saved = (original file name, result); failed = (original file name, error)."""
    saved: list[tuple[str, UploadResult]] = field(default_factory=list)
    failed: list[tuple[str, Exception]] = field(default_factory=list)


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _windows_safe(name: str) -> str:
    """Windows-safe folder name: <>:"/\\|?* and control characters -> "_", no leading spaces, no trailing
    dots/spaces, at most _MAX_FOLDER_NAME characters, reserved device names (CON, NUL, COM1 ...) get a "_".
    Thai characters and inner spaces are kept. May return "" (caller falls back)."""
    safe = _UNSAFE_CHARS.sub("_", name).lstrip(" ")[:_MAX_FOLDER_NAME].rstrip(". ")
    if safe.split(".")[0].rstrip(" ").upper() in _RESERVED_NAMES:
        safe = f"{safe}_"
    return safe


def folder_name(app_name: str | None, item_id: str) -> str:
    """Sub-folder for an item's uploads: the App name (first line), else the item id, made Windows-safe."""
    return _windows_safe(first_line(app_name)) or _windows_safe(item_id) or "_"


def destination_folder(cfg, repo, item_id: str) -> Path:
    """Where uploads of this item go: <attachment folder>/<App name or item id>."""
    rec = repo.get_item(TEST_PROGRESS, item_id) or {}
    return cfg.attachment_dir / folder_name(rec.get("application"), item_id)


def _reserve_dest(folder: Path, name: str) -> Path:
    """Create an empty <folder>/<name>, or "<stem> (2)<ext>", "(3)" ... when the name is taken.
    Creating with "x" mode means an existing file is never overwritten, even by a parallel upload."""
    stem, ext = Path(name).stem, Path(name).suffix
    candidate, n = name, 1
    while True:
        dest = folder / candidate
        try:
            with open(dest, "xb"):
                return dest
        except FileExistsError:
            n += 1
            candidate = f"{stem} ({n}){ext}"


def _dest_for(base: Path, sub: str, src: Path) -> Path:
    """<base>/<sub>/<original name>, reserved so nothing is overwritten.
    Creates the sub folder (never <base> itself when the caller checked it exists)."""
    folder = base / sub
    folder.mkdir(parents=True, exist_ok=True)
    return _reserve_dest(folder, src.name)


def _remove_partial(dest: Path) -> None:
    try:
        dest.unlink(missing_ok=True)
    except OSError:
        log.warning("Partly copied file could not be removed: %s", dest, exc_info=True)


def _copy_verified(base: Path, sub: str, src: Path) -> tuple[Path, str]:
    """Copy to a reserved new name + verify the SHA-256. -> (stored file, sha256).
    On OSError the reserved / partly copied file is removed (when possible) and the error re-raised."""
    dest: Path | None = None
    try:
        dest = _dest_for(base, sub, src)
        shutil.copy2(src, dest)
        sha = _sha256(dest)
    except OSError:
        if dest is not None:
            _remove_partial(dest)
        raise
    if sha != _sha256(src):
        _remove_partial(dest)
        raise UploadError("Copy verification failed (hash mismatch); nothing saved")
    return dest, sha


def _folder_unavailable(cfg, item_id: str, base: Path, *, failed: bool = False) -> AttachmentFolderUnavailable:
    log.warning("Result files folder unavailable (project %s, item %s): %s", cfg.project_id, item_id, base,
                exc_info=failed)
    return AttachmentFolderUnavailable(base)


def _store_in_chosen_folder(cfg, item_id: str, sub: str, src: Path) -> tuple[Path, str]:
    """The project's chosen folder: it is never created here, and any OSError while writing blocks the upload
    (no fallback to <data_dir>/attachments)."""
    base = cfg.attachment_dir
    if not base.is_dir():
        raise _folder_unavailable(cfg, item_id, base)
    try:
        return _copy_verified(base, sub, src)
    except OSError:
        raise _folder_unavailable(cfg, item_id, base, failed=True) from None


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

    sub = folder_name(rec.get("application"), item_id)
    if cfg.uses_attachment_folder:
        dest, sha = _store_in_chosen_folder(cfg, item_id, sub, src)
    else:
        dest, sha = _copy_verified(cfg.attachment_dir, sub, src)

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


def upload_many(cfg, repo, item_id: str, file_paths: Iterable[str | Path], *, actor: Actor,
                now: str | None = None) -> BatchUploadResult:
    """Upload each file separately with upload() (same checks, own attachment row + audit each).
    A file that fails (UploadError, PermissionDenied) is listed in .failed and the others still go on.
    TC/Result Complete is set by the first saved file only (upload() sees it is already Complete after)."""
    now = now or _now()
    result = BatchUploadResult()
    for path in file_paths:
        name = Path(path).name
        try:
            result.saved.append((name, upload(cfg, repo, item_id, path, actor=actor, now=now)))
        except (UploadError, PermissionDenied) as exc:
            result.failed.append((name, exc))
    return result
