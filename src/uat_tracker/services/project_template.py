"""New project workbook from the template kept in code (resources/project_template.json, built once by
tools/build_template.py from a copy of 'UAT Planning.xlsx'). Approved 10-Nov-2026:

- 3 sheets only: UAT Plan, Test Progress, Defect List. Header row + formatting, no data rows.
- Column widths / hidden columns, header row height, header styles, data-row style (pre-applied to the first
  PREFORMAT_ROWS rows so rows added later look like the old file), Traffic Light drop-down, whole-column
  colour rules, Test Progress legend (Status / Definition), the source workbook's theme.
- Hidden App ID column already present; header row frozen.
- Test Progress also has Report By Email + CC Email (approved 12a).
- export_blank_form(): the same blank workbook saved to a user-chosen path (Download blank form, 25-Nov-2026);
  refuses the current project's Excel file.
"""

from __future__ import annotations

import base64
import json
import os
import tempfile
from functools import lru_cache
from importlib import resources
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule, Rule
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

PREFORMAT_ROWS = 200
SHEET_ORDER = ("UAT Plan", "Test Progress", "Defect List")
FORM_FILENAME = "Planning Progress Form.xlsx"         # default name in the blank-form save dialog


class FormIsProjectFile(ValueError):
    """The chosen path is the current project's Excel file. Nothing was written."""


@lru_cache(maxsize=1)
def template() -> dict:
    return json.loads(resources.files("uat_tracker.resources").joinpath("project_template.json")
                      .read_text(encoding="utf-8"))


def headers(sheet: str) -> list:
    return [c["header"] for c in next(s for s in template()["sheets"] if s["name"] == sheet)["columns"]]


def _color(d):
    if not d:
        return None
    if "rgb" in d:
        return Color(rgb=d["rgb"])
    if "theme" in d:
        return Color(theme=d["theme"], tint=d.get("tint") or 0.0)
    if "indexed" in d:
        return Color(indexed=d["indexed"])
    return None


def _apply(cell, st: dict) -> None:
    f = st["font"]
    kw = {"name": f["name"], "sz": f["sz"], "b": f["b"], "i": f["i"]}
    if _color(f.get("color")) is not None:
        kw["color"] = _color(f["color"])
    cell.font = Font(**kw)
    if st.get("fill"):
        fl = st["fill"]
        cell.fill = PatternFill(fill_type=fl["type"], fgColor=_color(fl["fg"]) or Color(),
                                bgColor=_color(fl["bg"]) or Color())
    a = st["align"]
    cell.alignment = Alignment(horizontal=a["h"], vertical=a["v"], wrap_text=a["wrap"])
    cell.border = Border(**{k: Side(style=v) for k, v in st["border"].items() if v})
    cell.number_format = st["nf"]


def _dxf(d: dict) -> DifferentialStyle:
    fill = None
    if d.get("fill_fg") or d.get("fill_bg"):
        fill = PatternFill(fill_type="solid", fgColor=_color(d.get("fill_fg")) or Color(),
                           bgColor=_color(d.get("fill_bg")) or Color())
    font = Font(color=_color(d["font"])) if d.get("font") else None
    return DifferentialStyle(fill=fill, font=font)


def _build_workbook() -> openpyxl.Workbook:
    """Blank workbook from template(): 3 sheets, headers, styles, validations, cond formats, legend. Not saved."""
    tpl = template()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    if tpl.get("theme"):
        wb.loaded_theme = base64.b64decode(tpl["theme"])
    for sd in tpl["sheets"]:
        ws = wb.create_sheet(sd["name"])
        cols = sd["columns"]
        heads = [c["header"] for c in cols]
        for i, c in enumerate(cols, start=1):
            letter = get_column_letter(i)
            cell = ws.cell(1, i, c["header"])
            _apply(cell, c["head_style"])
            for r in range(2, PREFORMAT_ROWS + 2):
                _apply(ws.cell(r, i), c["data_style"])
            dim = ws.column_dimensions[letter]
            if c["width"]:
                dim.width = c["width"]
            dim.hidden = c["hidden"]
        if sd.get("header_height"):
            ws.row_dimensions[1].height = sd["header_height"]
        ws.freeze_panes = "A2"
        last = get_column_letter(len(cols))
        if sd.get("autofilter"):
            ws.auto_filter.ref = f"A1:{last}1"
        for v in sd["validations"]:
            col = get_column_letter(heads.index(v["column"]) + 1)
            dv = DataValidation(type=v["type"], formula1=v["formula1"], allow_blank=v["allow_blank"])
            dv.add(f"{col}2:{col}1048576")
            ws.add_data_validation(dv)
        for cf in sd["cond_formats"]:
            c1 = get_column_letter(heads.index(cf["first"]) + 1)
            c2 = get_column_letter(heads.index(cf["last"]) + 1)
            if cf["type"] == "cellIs":
                rule = CellIsRule(operator=cf["operator"], formula=cf["formula"])
                rule.dxf = _dxf(cf["dxf"])
            else:
                rule = Rule(type=cf["type"], operator=cf["operator"], formula=cf["formula"], dxf=_dxf(cf["dxf"]))
            ws.conditional_formatting.add(f"{c1}2:{c2}1048576", rule)
        if sd.get("legend"):
            start = len(cols) + 2                         # one blank column keeps it out of the data table
            for k, w in enumerate(sd["legend"]["widths"]):
                if w:
                    ws.column_dimensions[get_column_letter(start + k)].width = w
            for r, row in enumerate(sd["legend"]["rows"], start=1):
                for k, cell in enumerate(row):
                    c = ws.cell(r, start + k, cell["value"])
                    _apply(c, cell["style"])
    return wb


def create_workbook(path: Path) -> Path:
    """Write a new, empty project workbook. The caller has checked that `path` does not exist."""
    path = Path(path)
    if path.exists():
        raise FileExistsError(path)
    wb = _build_workbook()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.stem + ".creating.xlsx")
    wb.save(tmp)
    tmp.replace(path)
    return path


def _same_file(a: Path, b: Path) -> bool:
    """True when a and b point to the same path: resolve() + os.path.normcase (case-insensitive on Windows)."""
    return os.path.normcase(str(Path(a).resolve())) == os.path.normcase(str(Path(b).resolve()))


def export_blank_form(dest: Path, project_excel: Path | None) -> Path:
    """Save a blank form at dest, overwriting an existing file (the save dialog already asked). Returns dest.

    Raises FormIsProjectFile when dest is project_excel (checked first; nothing written, no temp file).
    Raises OSError (e.g. PermissionError) when it cannot be saved; the old dest is untouched, the temp is removed.
    No database, Excel COM or auth access.
    """
    dest = Path(dest)
    if project_excel is not None and _same_file(dest, Path(project_excel)):
        raise FormIsProjectFile(str(dest))
    wb = _build_workbook()
    fd, name = tempfile.mkstemp(dir=dest.parent, prefix=".~form-", suffix=".xlsx")
    os.close(fd)
    tmp = Path(name)
    try:
        wb.save(tmp)
        os.replace(tmp, dest)
    except BaseException:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        raise
    return dest
