"""Weekly report data (Part 6). Every number and table is computed HERE; the AI only writes prose about it.

Reuses the Part 2 Traffic Light and the Part 4 dashboard definitions so the report matches the app.
No people's names or email addresses are included (approved: not sent to the AI).
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from ..excel.sheets import DEFECT_LIST, TEST_PROGRESS, UAT_PLAN
from ..parsing.dates import fmt
from . import views
from .dashboard import CLOSED_DEFECT, first_line
from .traffic_light import AMBER, COMPLETE, GREEN, MISSING, RED, compute

APP_CHANNELS = ("app", "restore")


def _n(v) -> int:
    try:
        return int(float(v or 0))
    except (TypeError, ValueError):
        return 0


def pct(part: int, whole: int) -> str:
    return f"{part / whole * 100:.1f}%" if whole else "0.0%"


def _p(x) -> str | None:
    return None if x is None else f"{x * 100:.1f}%"


def _reason_kind(light, rec, as_of: dt.date, x: float) -> str:
    pf = rec.get("planned_finish")
    if pf and as_of > dt.date.fromisoformat(pf):
        return "overdue"
    if light.color == RED or (light.pct_exec is not None and light.pct_exec >= (light.pct_time or 0)):
        return "fail_rate"
    return "time"


def build_data(cfg, repo, as_of: dt.date, *, days_ahead: int = 14, days_back: int = 7) -> dict:
    A = as_of.isoformat()
    soon = (as_of + dt.timedelta(days=days_ahead)).isoformat()
    since = (as_of - dt.timedelta(days=days_back)).isoformat()
    uat = sorted((r for r in repo.load_items(UAT_PLAN).values() if r.get("active")),
                 key=lambda r: r.get("excel_row") or 10 ** 6)
    tp = sorted((r for r in repo.load_items(TEST_PROGRESS).values() if r.get("active")),
                key=lambda r: r.get("excel_row") or 10 ** 6)
    dfs = sorted((r for r in repo.load_items(DEFECT_LIST).values() if r.get("active")),
                 key=lambda r: r.get("excel_row") or 10 ** 6)
    x = cfg.alerts.fail_threshold

    executed = sum(_n(r.get("executed")) for r in tp)
    total = sum(_n(r.get("total_tc")) for r in tp)
    colors = {COMPLETE: 0, RED: 0, AMBER: 0, GREEN: 0, MISSING: 0}
    red_amber, tp_overdue, due = [], [], []
    for r in tp:
        light = compute(r, as_of, x)
        colors[light.color] = colors.get(light.color, 0) + 1
        name = first_line(r.get("application")) or r["id"]
        pf = r.get("planned_finish")
        if light.color in (RED, AMBER):
            kind = _reason_kind(light, r, as_of, x)
            red_amber.append({"id": r["id"], "app": name, "color": light.color, "reason_kind": kind,
                              "reason": light.reason, "planned_finish": fmt(pf) if pf else None,
                              "executed": _n(r.get("executed")), "total_tc": _n(r.get("total_tc")),
                              "fail": _n(r.get("fail_count")), "pct_exec": _p(light.pct_exec),
                              "pct_time": _p(light.pct_time) if kind == "time" else None,   # >100% once overdue
                              "pct_fail": _p(light.pct_fail)})
        not_done = r.get("status") != "Complete"
        if not_done and pf and pf < A:
            tp_overdue.append({"id": r["id"], "app": name, "planned_finish": fmt(pf),
                               "days_late": (as_of - dt.date.fromisoformat(pf)).days})
        if not_done and pf and A <= pf <= soon:
            due.append({"date": fmt(pf), "_iso": pf, "page": "Test Progress", "id": r["id"], "name": name})
    red_amber.sort(key=lambda a: (a["color"] != RED, a["app"]))

    uat_overdue = []
    for t in uat:
        f = t.get("finish_date")
        if t.get("status") == "Complete" or not f:
            continue
        if f < A:
            uat_overdue.append({"id": t["id"], "task": first_line(t.get("task_name")), "finish_date": fmt(f),
                                "days_late": (as_of - dt.date.fromisoformat(f)).days})
        elif f <= soon:
            due.append({"date": fmt(f), "_iso": f, "page": "UAT Plan", "id": t["id"],
                        "name": first_line(t.get("task_name"))})
    due.sort(key=lambda d: (d["_iso"], d["page"] != "UAT Plan", d["id"]))
    for d in due:
        d.pop("_iso")

    open_d = [d for d in dfs if d.get("status") not in CLOSED_DEFECT]
    d_overdue, d_none = [], []
    for d in open_d:
        target = d.get("fix_target_date")
        row = {"id": d["id"], "system": first_line(d.get("system")), "detail": first_line(d.get("detail"))[:120],
               "status": d.get("status")}
        if not target:
            d_none.append(row)
        elif target < A:
            d_overdue.append({**row, "fix_target_date": fmt(target),
                              "days_late": (as_of - dt.date.fromisoformat(target)).days})

    names = {(UAT_PLAN, t["id"]): first_line(t.get("task_name")) for t in uat}
    names.update({(TEST_PROGRESS, r["id"]): first_line(r.get("application")) for r in tp})
    resched = []
    for e in reversed(views.history(repo, kind="Reschedule", date_from=since, date_to=(as_of - dt.timedelta(days=1)).isoformat())):
        if e["sheet"] not in (UAT_PLAN, TEST_PROGRESS):
            continue
        in_app = e["channel"] in APP_CHANNELS
        resched.append({"date": fmt(e["ts"][:10]), "page": e["sheet"], "id": e["item_id"],
                        "name": names.get((e["sheet"], e["item_id"]), e["item_id"]), "field": e["field"],
                        "from": (e["old_value"] or "").split("=>")[-1].strip(),
                        "to": (e["new_value"] or "").split("=>")[-1].strip(),
                        "reason": (e["reason"] or "").strip() if in_app else None,
                        "where": "app" if in_app else "excel"})

    return {
        "project": cfg.project_name or Path(cfg.excel_path).stem,
        "as_of": fmt(A),
        "reschedule_window": {"from": fmt(since), "to": fmt((as_of - dt.timedelta(days=1)).isoformat())},
        "overall": {"apps": len(tp), "executed": executed, "total_tc": total, "pct_executed": pct(executed, total),
                    "pass": sum(_n(r.get("pass_count")) for r in tp), "fail": sum(_n(r.get("fail_count")) for r in tp),
                    "colors": {k: v for k, v in colors.items()}},
        "red_amber": red_amber,
        "overdue": {"uat_plan": uat_overdue, "test_progress": tp_overdue},
        "defects": {"open": len(open_d), "overdue": d_overdue, "no_target": d_none},
        "due_next_2_weeks": due,
        "reschedules_last_7_days": resched,
        "open_conflicts": len(repo.open_conflicts()),
        "review_queue": len(repo.open_reviews()),
        "fail_threshold": _p(x),
    }
