"""Read-only Excel access. The source workbook is copied first; only the copy is opened."""

from __future__ import annotations

import logging
import hashlib
import os
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import openpyxl

from ..parsing.headers import HeaderMap, map_headers
from ..parsing.text import is_blank
from .sheets import SheetDef

log = logging.getLogger(__name__)


@dataclass
class RawRow:
    excel_row: int
    values: dict          # field key -> raw cell value


@dataclass
class SheetData:
    sheet: str
    headers: HeaderMap
    rows: list[RawRow]


def open_shared(path: Path):
    """Open a file read-only while letting other programs keep reading, writing AND deleting it.
    Excel with AutoSave (OneDrive) holds the workbook so that a plain open() gets 'Permission denied';
    sharing Delete as well is what it accepts. Read-only: nothing is written to the source."""
    if os.name != "nt":
        return open(path, "rb")
    try:
        import msvcrt
        import win32file
    except ImportError:                                  # pragma: no cover
        return open(path, "rb")
    share = win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE
    try:
        h = win32file.CreateFile(str(path), win32file.GENERIC_READ, share, None, win32file.OPEN_EXISTING, 0, None)
    except Exception as exc:                             # pywintypes.error -> same type callers expect
        log.warning("Cannot open %s for reading", path, exc_info=True)
        raise PermissionError(f"Cannot read {path}: {exc}") from exc
    fd = msvcrt.open_osfhandle(h.Detach(), os.O_RDONLY)
    return os.fdopen(fd, "rb")


def copy_shared(src: Path, dest: Path) -> None:
    """shutil.copy2 for a source another program has open (see open_shared)."""
    with open_shared(src) as f, open(dest, "wb") as g:
        shutil.copyfileobj(f, g, 1 << 20)
    shutil.copystat(src, dest)


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open_shared(path) as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def take_snapshot(source: Path, snapshot_dir: Path) -> tuple[Path, str]:
    """Copy the source workbook into snapshot_dir. Reuses an identical earlier snapshot."""
    if not source.exists():
        raise FileNotFoundError(f"Source workbook not found: {source}")
    snapshot_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    tmp = snapshot_dir / f".incoming-{stamp}.xlsx"
    copy_shared(source, tmp)           # read-only on the source
    sha = sha256_file(tmp)
    existing = sorted(snapshot_dir.glob(f"*_{sha[:12]}.xlsx"))
    if existing:
        tmp.unlink()
        return existing[0], sha
    final = snapshot_dir / f"{source.stem}_{stamp}_{sha[:12]}.xlsx"
    tmp.rename(final)
    return final, sha


def read_sheets(snapshot: Path, sheet_defs: list[SheetDef]) -> dict[str, SheetData]:
    """Open the snapshot read-only and read ONLY the named sheets."""
    wb = openpyxl.load_workbook(snapshot, read_only=True, data_only=True)
    try:
        out: dict[str, SheetData] = {}
        for sd in sheet_defs:
            if sd.name not in wb.sheetnames:
                raise KeyError(f"Sheet '{sd.name}' not found in workbook")
            ws = wb[sd.name]
            rows_iter = ws.iter_rows(values_only=True)
            header = next(rows_iter, ())
            hmap = map_headers(list(header), list(sd.fields), sd.name, sd.blank_first_as)
            rows: list[RawRow] = []
            for r_idx, row in enumerate(rows_iter, start=2):
                values = {k: (row[c] if c < len(row) else None) for k, c in hmap.columns.items()}
                if all(is_blank(v) for v in values.values()):
                    continue
                rows.append(RawRow(r_idx, values))
            out[sd.name] = SheetData(sd.name, hmap, rows)
        return out
    finally:
        wb.close()
