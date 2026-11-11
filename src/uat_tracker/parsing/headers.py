"""Map columns by header NAME, never by position.

Header label = first non-empty line of the header cell (so 'Task name\\n(Mandatory)' -> 'Task name'),
matched case-insensitively against a field's accepted names.
The table ends at the first blank header cell (this keeps side tables such as the Traffic Light
legend in Test Progress out). When a name appears twice, the leftmost column wins.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .text import cell_text


class HeaderError(Exception):
    pass


@dataclass(frozen=True)
class FieldSpec:
    key: str
    names: tuple[str, ...]
    required: bool = True


@dataclass
class HeaderMap:
    columns: dict[str, int]                 # field key -> 0-based column index
    missing_optional: list[str] = field(default_factory=list)


def header_label(v) -> str:
    for line in cell_text(v).split("\n"):
        line = re.sub(r"\s+", " ", line).strip()
        if line:
            return line
    return ""


def map_headers(header_row, specs: list[FieldSpec], sheet: str,
                blank_first_as: str | None = None) -> HeaderMap:
    labels: list[str] = []
    for i, v in enumerate(header_row):
        label = header_label(v)
        if not label:
            if i == 0 and blank_first_as:
                label = blank_first_as
            elif i == 0:
                labels.append("")
                continue
            else:
                break
        labels.append(label.casefold())

    columns: dict[str, int] = {}
    missing_req, missing_opt = [], []
    for spec in specs:
        wanted = {n.casefold() for n in spec.names}
        idx = next((i for i, lab in enumerate(labels) if lab in wanted), None)
        if idx is None:
            (missing_req if spec.required else missing_opt).append(spec.names[0])
        else:
            columns[spec.key] = idx
    if missing_req:
        raise HeaderError(f"Sheet '{sheet}': required column(s) not found: {', '.join(missing_req)}")
    return HeaderMap(columns=columns, missing_optional=missing_opt)
