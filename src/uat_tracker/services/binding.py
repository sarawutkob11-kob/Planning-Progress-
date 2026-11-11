"""Bind Excel rows to stable app IDs.

Part 1 (read-only) binds by content, because we cannot add an ID column to Excel yet:
  pass 1: exact key + occurrence number (e.g. 2nd 'E / 8 / efs' row)
  pass 2: fallback key (e.g. group + task no) when exactly one unmatched row and one unmatched
          record share it -> treated as the same item whose name changed.
Unmatched rows get new IDs; unmatched records are marked inactive by the importer.
Part 3 adds a hidden 'App ID' column in Excel; when present it is used first (pass 0).
"""

from __future__ import annotations

import json
import string
from collections import Counter, defaultdict

from ..domain.models import Item


def fingerprints(items: list[Item]) -> list[str]:
    seen: Counter = Counter()
    out = []
    for it in items:
        seen[it.key] += 1
        out.append(json.dumps(list(it.key) + [seen[it.key]], ensure_ascii=False))
    return out


def alt_key_json(item: Item) -> str | None:
    return json.dumps(list(item.alt_key), ensure_ascii=False) if item.alt_key else None


def _new_id(item: Item, used: set[str]) -> str:
    if item.id_base:
        if item.id_base not in used:
            return item.id_base
        for suffix in string.ascii_lowercase[1:]:
            if item.id_base + suffix not in used:
                return item.id_base + suffix
        n = 2
        while f"{item.id_base}-{n}" in used:
            n += 1
        return f"{item.id_base}-{n}"
    n = 1
    while f"{item.id_blank_prefix}{n}" in used:
        n += 1
    return f"{item.id_blank_prefix}{n}"


def bind(items: list[Item], existing: dict[str, dict]) -> list[tuple[str, bool]]:
    """Return [(id, is_new)] aligned with items. `existing` rows need fingerprint/alt_key/active."""
    fps = fingerprints(items)
    ids: list[str | None] = [None] * len(items)
    taken: set[str] = set()

    # pass 0 (Part 3): the hidden 'App ID' column. A copied row repeats an ID: only the first keeps it.
    for i, it in enumerate(items):
        if it.app_id and it.app_id in existing and it.app_id not in taken:
            ids[i] = it.app_id
            taken.add(it.app_id)

    by_fp: dict[str, str] = {}
    for rid, rec in sorted(existing.items(), key=lambda kv: -int(kv[1].get("active") or 0)):
        by_fp.setdefault(rec.get("fingerprint"), rid)
    for i, fp in enumerate(fps):
        rid = by_fp.get(fp)
        if ids[i] is None and rid and rid not in taken:
            ids[i] = rid
            taken.add(rid)

    rows_by_alt: dict[str, list[int]] = defaultdict(list)
    for i, it in enumerate(items):
        if ids[i] is None and (k := alt_key_json(it)):
            rows_by_alt[k].append(i)
    recs_by_alt: dict[str, list[str]] = defaultdict(list)
    for rid, rec in existing.items():
        if rid not in taken and rec.get("alt_key"):
            recs_by_alt[rec["alt_key"]].append(rid)
    for k, rows in rows_by_alt.items():
        recs = recs_by_alt.get(k, [])
        if len(rows) == 1 and len(recs) == 1:
            ids[rows[0]] = recs[0]
            taken.add(recs[0])

    used = set(existing) | taken
    out: list[tuple[str, bool]] = []
    for i, it in enumerate(items):
        if ids[i] is not None:
            out.append((ids[i], False))
        else:
            nid = _new_id(it, used)
            used.add(nid)
            out.append((nid, True))
    return out
