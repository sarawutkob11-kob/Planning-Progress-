"""Dev tool: rebuild src/uat_tracker/resources/project_template.json from a COPY of the source workbook.

The app never reads the live workbook to create a project (approved): the template is this JSON, kept in
the code. Run again only if the PMO changes the layout of 'UAT Planning.xlsx':

    python tools/build_template.py "<copy of UAT Planning.xlsx>"

Kept: header cells (text, style), column widths / hidden, header row height, the data-row style (a normal task
row, without highlight fills),
whole-column data validation and conditional formatting, the Test Progress legend, the workbook theme.
Added (approved 10-Nov-2026): Report By Email + CC Email right after "Report By" in Test Progress.
"""

from __future__ import annotations

import base64
import json
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries

SHEETS = ("UAT Plan", "Test Progress", "Defect List")
OUT = Path(__file__).resolve().parents[1] / "src" / "uat_tracker" / "resources" / "project_template.json"


def color(c):
    if c is None:
        return None
    t = getattr(c, "type", None)
    if t == "rgb":
        return {"rgb": c.rgb} if isinstance(c.rgb, str) else None
    if t == "theme":
        return {"theme": c.theme, "tint": c.tint}
    if t == "indexed":
        return {"indexed": c.indexed}
    return None


def style(cell) -> dict:
    f, fl, a, b = cell.font, cell.fill, cell.alignment, cell.border
    return {
        "font": {"name": f.name, "sz": f.sz, "b": bool(f.b), "i": bool(f.i), "color": color(f.color)},
        "fill": ({"type": fl.fill_type, "fg": color(fl.fgColor), "bg": color(fl.bgColor)} if fl.fill_type else None),
        "align": {"h": a.horizontal, "v": a.vertical, "wrap": bool(a.wrap_text)},
        "border": {s: getattr(b, s).style for s in ("left", "right", "top", "bottom")},
        "nf": cell.number_format,
    }


def table_width(ws) -> int:
    """Header columns of the table: up to the first blank header after column A (+ the App ID column)."""
    n = 0
    for c in ws[1]:
        if c.value is None and c.column > 1:
            break
        n = c.column
    return n


def sample_row(ws, width: int) -> int:
    """A normal data row: the first row with values in at least half the table columns (UAT Plan rows 2-3 are
    phase / group headers). Its fills are not copied (PMO highlights), only font / borders / formats."""
    for r in range(2, min(ws.max_row, 200) + 1):
        if sum(ws.cell(r, c).value not in (None, "") for c in range(1, width + 1)) >= width / 2:
            return r
    return 2


def dxf(rule) -> dict:
    d = rule.dxf
    return {"fill_fg": color(d.fill.fgColor) if d and d.fill else None,
            "fill_bg": color(d.fill.bgColor) if d and d.fill else None,
            "font": color(d.font.color) if d and d.font and d.font.color else None}


def main(src: str) -> None:
    wb = openpyxl.load_workbook(src)
    out = {"source": Path(src).name, "theme": base64.b64encode(wb.loaded_theme).decode() if wb.loaded_theme else None,
           "sheets": []}
    for name in SHEETS:
        ws = wb[name]
        width = table_width(ws)
        sample = sample_row(ws, width)
        cols = []
        for i in range(1, width + 1):
            letter = get_column_letter(i)
            h = ws.cell(1, i)
            cd = ws.column_dimensions[letter]
            cols.append({"header": h.value, "width": cd.width, "hidden": bool(cd.hidden),
                         "head_style": style(h), "data_style": {**style(ws.cell(sample, i)), "fill": None}})
        sheet = {"name": name, "header_height": ws.row_dimensions[1].height, "columns": cols,
                 "autofilter": bool(ws.auto_filter.ref), "validations": [], "cond_formats": [], "legend": None}
        headers = [c["header"] for c in cols]
        for dv in ws.data_validations.dataValidation:
            for rng in str(dv.sqref).split():
                c1, r1, c2, r2 = range_boundaries(rng)
                if r2 and r2 >= 1048575 and c1 == c2 and c1 <= width:
                    sheet["validations"].append({"column": headers[c1 - 1], "type": dv.type,
                                                 "formula1": dv.formula1, "allow_blank": bool(dv.allow_blank)})
                    break
        for cf in ws.conditional_formatting:
            rngs = [range_boundaries(r) for r in str(cf.sqref).split()]
            if not any(r2 and r2 >= 1048575 for _, _, _, r2 in rngs):
                continue                                   # single-cell leftovers (M51, N29 ...) are not copied
            c1 = min(r[0] for r in rngs)
            c2 = max(r[2] for r in rngs)
            for rule in cf.rules:
                sheet["cond_formats"].append({"first": headers[c1 - 1], "last": headers[c2 - 1], "type": rule.type,
                                              "operator": rule.operator, "formula": list(rule.formula or []),
                                              "dxf": dxf(rule)})
        if name == "Test Progress":                        # legend table right of the data table (approved 12b)
            start = width + 2
            legend = []
            for r in range(1, ws.max_row + 1):
                vals = [ws.cell(r, start + k) for k in range(2)]
                if all(v.value is None for v in vals):
                    if r > 1:
                        break
                    continue
                legend.append([{"value": v.value, "style": style(v)} for v in vals])
            sheet["legend"] = {"widths": [ws.column_dimensions[get_column_letter(start + k)].width for k in range(2)],
                               "rows": legend}
            # 12a: Report By Email + CC Email after "Report By" (style of Report By, wider column)
            at = next(i for i, c in enumerate(cols) if c["header"] == "Report By") + 1
            for k, head in enumerate(("Report By Email", "CC Email")):
                col = json.loads(json.dumps(cols[at - 1]))
                col["header"], col["width"], col["hidden"] = head, 28.0, False
                cols.insert(at + k, col)
        out["sheets"].append(sheet)
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"written {OUT}")


if __name__ == "__main__":
    main(sys.argv[1])
