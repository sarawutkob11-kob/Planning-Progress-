"""Bank holidays: BOT "Financial Institutions' Holidays" API + manual entry / CSV.

Approved behaviour:
- Fetch weekly in Monday's run for this year and next, plus whenever the current year has no data.
- Fetch fails -> keep the cache, log it, warn the PMO (the daily job sends the warning).
- No cached data for the current year -> the daily job sends nothing and alerts the PMO.
"""

from __future__ import annotations

import csv
import datetime as dt
import json
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path

from ..domain.models import CHANNEL_APP, CHANNEL_SYSTEM

SHEET = "Holidays"


class FetchError(Exception):
    pass


@dataclass
class FetchResult:
    year: int
    ok: bool
    count: int = 0
    changed: int = 0
    message: str = ""


def parse_bot_response(payload) -> list[tuple[str, str]]:
    """Find holiday rows in the BOT JSON ({"result": {"data": [{"Date": ..., "HolidayDescription": ...}]}}).
    Searches the structure so a small envelope change on the new gateway doesn't break it."""
    found: list[tuple[str, str]] = []

    def walk(node):
        if isinstance(node, list):
            for x in node:
                walk(x)
        elif isinstance(node, dict):
            keys = {k.lower(): k for k in node}
            if "date" in keys and isinstance(node[keys["date"]], str):
                raw = node[keys["date"]][:10]
                try:
                    d = dt.date.fromisoformat(raw)
                except ValueError:
                    raise FetchError(f"Unexpected holiday date format: {node[keys['date']]!r}")
                name_key = next((keys[k] for k in ("holidaydescription", "holidaydescriptionthai", "name")
                                 if k in keys), None)
                found.append((d.isoformat(), str(node[name_key]) if name_key else ""))
                return
            for v in node.values():
                walk(v)

    walk(payload)
    return found


def fetch_bot_year(url_template: str, api_key: str, key_header: str, year: int,
                   timeout: float = 30) -> list[tuple[str, str]]:
    if not api_key:
        raise FetchError("No BOT API key configured ([alerts] bot_api_key or UAT_BOT_API_KEY)")
    req = urllib.request.Request(url_template.format(year=year),
                                 headers={key_header: api_key, "Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise FetchError(f"BOT API HTTP {exc.code}: {exc.reason}") from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise FetchError(f"BOT API not reachable: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise FetchError(f"BOT API returned non-JSON: {exc}") from exc
    rows = [r for r in parse_bot_response(payload) if r[0].startswith(f"{year}-")]
    if not rows:
        raise FetchError(f"BOT API returned no holidays for {year}")
    return rows


def store_holidays(repo, rows, source: str, now: str, editor: str) -> int:
    changed = 0
    for date_iso, name in rows:
        old = repo.holidays(int(date_iso[:4])).get(date_iso)
        if repo.upsert_holiday(date_iso, name, source, now):
            changed += 1
            repo.append_audit({"ts": now, "editor": editor,
                               "channel": CHANNEL_SYSTEM if source == "bot" else CHANNEL_APP,
                               "sheet": SHEET, "item_id": date_iso, "field": "holiday",
                               "old_value": old["name"] if old else "", "new_value": name,
                               "reason": f"Holiday from {source}"})
    return changed


def refresh_year(repo, cfg, year: int, *, now: str, editor: str, fetcher=None) -> FetchResult:
    """Fetch one year from BOT and cache it. Never clears the cache on failure."""
    fetcher = fetcher or (lambda y: fetch_bot_year(cfg.alerts.bot_url, cfg.alerts.bot_api_key,
                                                   cfg.alerts.bot_key_header, y))
    try:
        rows = fetcher(year)
    except FetchError as exc:
        with repo.transaction():
            repo.log_holiday_fetch(ts=now, year=year, result="failed", count=None, message=str(exc))
        return FetchResult(year, False, message=str(exc))
    with repo.transaction():
        changed = store_holidays(repo, rows, "bot", now, editor)
        repo.log_holiday_fetch(ts=now, year=year, result="ok", count=len(rows),
                               message=f"{changed} new/changed")
    return FetchResult(year, True, len(rows), changed, f"{len(rows)} holidays, {changed} new/changed")


def read_csv(path: Path) -> list[tuple[str, str]]:
    """CSV with columns date,name (date as YYYY-MM-DD or DD-Mon-YYYY). Header row optional."""
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        for i, rec in enumerate(csv.reader(f), start=1):
            if not rec or not rec[0].strip():
                continue
            raw = rec[0].strip()
            name = rec[1].strip() if len(rec) > 1 else ""
            try:
                d = parse_user_date(raw)
            except ValueError:
                if i == 1:
                    continue                      # header
                raise ValueError(f"{path}: line {i}: bad date {raw!r}")
            rows.append((d.isoformat(), name))
    return rows


def parse_user_date(raw: str) -> dt.date:
    raw = raw.strip()
    try:
        return dt.date.fromisoformat(raw)
    except ValueError:
        return dt.datetime.strptime(raw, "%d-%b-%Y").date()
