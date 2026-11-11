"""Dashboard numbers (Part 4 screen 1). Calculation only; the UI shows what this returns.

Approved definitions:
- Open tasks = UAT Plan items whose Status is not Complete; overdue = Finish date before the As-of
  date and not Complete.
- Due today = UAT Plan Finish date or Test Progress Planned Finish = As-of date, not Complete.
- Open defects = Status not Close/Rejected; overdue = Fix target date before As-of; no target = empty.
- The As-of date only changes what is shown (colors recomputed); nothing is saved or sent.
- "Send reminder" is offered for UAT Plan and Defect items; Test Progress items get automatic emails.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from ..excel.sheets import DEFECT_LIST, TEST_PROGRESS, UAT_PLAN
from .traffic_light import AMBER, MISSING, RED, compute

CLOSED_DEFECT = ("Close", "Rejected")
TAG_ORDER = ("Overdue", "Today", "No target", "Red", "Amber", "Missing data")


@dataclass
class FollowItem:
    tag: str                  # Overdue | Today | No target | Red | Amber | Missing data
    sheet: str
    item_id: str
    name: str
    detail: dict              # {"kind": ..., ...} - the UI words it in TH/EN
    owner: str
    manual: bool              # Send reminder (UAT Plan / Defect) vs automatic email (Test Progress)


@dataclass
class Dashboard:
    as_of: dt.date
    executed: int = 0
    total_tc: int = 0
    open_tasks: int = 0
    all_tasks: int = 0
    overdue_tasks: int = 0
    due_today: int = 0
    next_due: tuple[str, int] | None = None          # (ISO date, count)
    open_defects: int = 0
    overdue_defects: int = 0
    no_target_defects: int = 0
    follow: list[FollowItem] = field(default_factory=list)
    incomplete_apps: list[dict] = field(default_factory=list)
    due_soon: list[dict] = field(default_factory=list)
    missing_email: list[str] = field(default_factory=list)
    apps: int = 0

    @property
    def pct_executed(self) -> float:
        return self.executed / self.total_tc if self.total_tc else 0.0


def first_line(s) -> str:
    return (str(s or "").split("\n")[0]).strip()


def _num(v) -> int:
    return int(v or 0)


def build(cfg, repo, as_of: dt.date, days_ahead: int = 14) -> Dashboard:
    A = as_of.isoformat()
    soon = (as_of + dt.timedelta(days=days_ahead)).isoformat()
    uat = [r for r in repo.load_items(UAT_PLAN).values() if r.get("active")]
    tp = [r for r in repo.load_items(TEST_PROGRESS).values() if r.get("active")]
    dfs = [r for r in repo.load_items(DEFECT_LIST).values() if r.get("active")]
    uat.sort(key=lambda r: r.get("excel_row") or 10 ** 6)
    tp.sort(key=lambda r: r.get("excel_row") or 10 ** 6)
    dfs.sort(key=lambda r: r.get("excel_row") or 10 ** 6)
    d = Dashboard(as_of)
    d.apps = len(tp)
    d.executed = sum(_num(r.get("executed")) for r in tp)
    d.total_tc = sum(_num(r.get("total_tc")) for r in tp)

    open_t = [t for t in uat if t.get("status") != "Complete"]
    d.all_tasks, d.open_tasks = len(uat), len(open_t)
    upcoming: dict[str, int] = {}
    for t in open_t:
        f = t.get("finish_date")
        if not f:
            continue
        if f < A:
            d.overdue_tasks += 1
            late = (as_of - dt.date.fromisoformat(f)).days
            d.follow.append(FollowItem("Overdue", UAT_PLAN, t["id"], first_line(t.get("task_name")),
                                       {"kind": "finish_late", "date": f, "days": late},
                                       first_line(t.get("responsible")), True))
        elif f == A:
            d.due_today += 1
            d.follow.append(FollowItem("Today", UAT_PLAN, t["id"], first_line(t.get("task_name")),
                                       {"kind": "due_today", "date": f}, first_line(t.get("responsible")), True))
        else:
            upcoming[f] = upcoming.get(f, 0) + 1
        if A <= f <= soon:
            d.due_soon.append({"date": f, "sheet": UAT_PLAN, "item_id": t["id"],
                               "name": first_line(t.get("task_name")), "owner": first_line(t.get("responsible"))})

    open_d = [x for x in dfs if x.get("status") not in CLOSED_DEFECT]
    d.open_defects = len(open_d)
    for x in open_d:
        target = x.get("fix_target_date")
        name = first_line(x.get("system")) or x["id"]
        if not target:
            d.no_target_defects += 1
            d.follow.append(FollowItem("No target", DEFECT_LIST, x["id"], name,
                                       {"kind": "text", "text": first_line(x.get("detail"))[:80]},
                                       first_line(x.get("assigned_to")), True))
        elif target < A:
            d.overdue_defects += 1
            d.follow.append(FollowItem("Overdue", DEFECT_LIST, x["id"], name,
                                       {"kind": "fix_late", "date": target,
                                        "days": (as_of - dt.date.fromisoformat(target)).days},
                                       first_line(x.get("assigned_to")), True))

    for r in tp:
        light = compute(r, as_of, cfg.alerts.fail_threshold)
        name, owner = first_line(r.get("application")), first_line(r.get("report_by"))
        pf = r.get("planned_finish")
        not_done = r.get("status") != "Complete"
        if not r.get("report_by_email"):
            d.missing_email.append(name)
        if not_done and pf == A:
            d.due_today += 1
            d.follow.append(FollowItem("Today", TEST_PROGRESS, r["id"], name, {"kind": "due_today", "date": pf},
                                       owner, False))
        elif not_done and pf and pf > A:
            upcoming[pf] = upcoming.get(pf, 0) + 1
        if not_done and pf and A <= pf <= soon:
            d.due_soon.append({"date": pf, "sheet": TEST_PROGRESS, "item_id": r["id"], "name": name, "owner": owner})
        if light.color == RED:
            kind = ({"kind": "past_finish", "date": pf} if pf and as_of > dt.date.fromisoformat(pf)
                    else {"kind": "fail_high", "pct": light.pct_fail})
            d.follow.append(FollowItem("Red", TEST_PROGRESS, r["id"], name, kind, owner, False))
        elif light.color == AMBER:
            d.follow.append(FollowItem("Amber", TEST_PROGRESS, r["id"], name,
                                       {"kind": "behind", "exec": light.pct_exec, "time": light.pct_time,
                                        "fail": _num(r.get("fail_count"))}, owner, False))
        elif light.color == MISSING:
            d.follow.append(FollowItem("Missing data", TEST_PROGRESS, r["id"], name,
                                       {"kind": "text", "text": light.reason}, owner, False))
        total, done = _num(r.get("total_tc")), _num(r.get("executed"))
        if not_done and done < total:
            d.incomplete_apps.append({"item_id": r["id"], "name": name, "executed": done, "total": total,
                                      "color": light.color})

    if upcoming:
        nxt = min(upcoming)
        d.next_due = (nxt, upcoming[nxt])
    d.follow.sort(key=lambda f: TAG_ORDER.index(f.tag))
    d.due_soon.sort(key=lambda x: (x["date"], x["sheet"] != UAT_PLAN, x["item_id"]))
    return d
