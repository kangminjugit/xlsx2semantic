"""Single-pass streaming scanner for sheet XML.

Replaces etree.fromstring() DOM parsing with iterparse streaming.
Collects all cell data, merge info in one pass with immediate memory release.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")
_MERGE_REF_RE = re.compile(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$")


@dataclass
class CellInfo:
    """Lightweight cell data extracted during streaming scan."""

    ref: str
    row: int
    col: int
    cell_type: str
    style_idx: str | None
    raw_value: str | None
    value: str | None
    formula: str | None


@dataclass
class SheetData:
    """All data extracted from a single sheet in one iterparse pass."""

    cells: list[CellInfo] = field(default_factory=list)
    merge_map: dict[str, str] = field(default_factory=dict)
    raw_values: dict[str, str] = field(default_factory=dict)


def scan_sheet(sheet_xml: str, shared_strings: list[str]) -> SheetData:
    """Stream-parse sheet XML and collect all data in a single pass.

    Args:
        sheet_xml: Raw XML string of the worksheet.
        shared_strings: Parsed shared string table.

    Returns:
        SheetData with cells, merge_map, and raw_values populated.
    """
    data = SheetData()

    context = etree.iterparse(
        io.BytesIO(sheet_xml.encode("utf-8")),
        events=("end",),
        tag=(f"{{{NS}}}c", f"{{{NS}}}mergeCell"),
    )

    for _, elem in context:
        if elem.tag == f"{{{NS}}}mergeCell":
            _expand_merge(elem.get("ref", ""), data.merge_map)
            elem.clear()
            continue

        # <c> element
        ref = elem.get("r", "")
        m = _CELL_REF_RE.match(ref)
        if not m:
            elem.clear()
            continue

        col = _col_to_num(m.group(1))
        row = int(m.group(2))
        cell_type = elem.get("t", "")
        style_idx = elem.get("s")

        v_el = elem.find(f"{{{NS}}}v")
        f_el = elem.find(f"{{{NS}}}f")

        raw_value = v_el.text if v_el is not None else None
        formula = f_el.text if f_el is not None else None

        value = _resolve_value(raw_value, cell_type, formula, shared_strings)

        if value and value.strip():
            data.raw_values[f"{col}:{row}"] = value

        data.cells.append(CellInfo(
            ref=ref, row=row, col=col,
            cell_type=cell_type, style_idx=style_idx,
            raw_value=raw_value, value=value, formula=formula,
        ))

        elem.clear()

    return data


def _resolve_value(
    raw_value: str | None,
    cell_type: str,
    formula: str | None,
    shared_strings: list[str],
) -> str | None:
    """Resolve cell value from raw data."""
    if raw_value is not None:
        if cell_type == "s" and shared_strings:
            try:
                idx = int(raw_value)
                if 0 <= idx < len(shared_strings):
                    return shared_strings[idx]
            except ValueError:
                pass
            return raw_value
        return raw_value

    if formula:
        return formula

    return None


def _expand_merge(ref: str, merge_map: dict[str, str]) -> None:
    """Expand a merge range ref into individual coordinate mappings."""
    m = _MERGE_REF_RE.match(ref)
    if not m:
        return

    start_col = _col_to_num(m.group(1))
    start_row = int(m.group(2))
    end_col = _col_to_num(m.group(3))
    end_row = int(m.group(4))

    origin = f"{start_col}:{start_row}"
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            merge_map[f"{c}:{r}"] = origin


def _col_to_num(letters: str) -> int:
    result = 0
    for ch in letters:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result
