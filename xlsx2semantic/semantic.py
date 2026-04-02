"""Semantic table XML transformer.

Converts sheet XML into a semantic structure where:
- Headers become XML tag names
- Data rows become <record> elements
- Title rows become a <title> element
- Row meta column becomes a record attribute

Supports both auto-detection and explicit layout hints.

Result example:
  <semantic-table>
    <title>Public school students by race/ethnicity</title>
    <schema>
      <column index="3" tag="total_number"/>
    </schema>
    <records count="52">
      <record row="8" state="Alabama">
        <total_number>745938</total_number>
      </record>
    </records>
  </semantic-table>
"""

from __future__ import annotations

import logging
import re
from collections import OrderedDict

from lxml import etree

logger = logging.getLogger(__name__)

from xlsx2semantic.layout_hint import TableLayoutHint
from xlsx2semantic.sheet_scanner import SheetData, scan_sheet
from xlsx2semantic.structural_anchor import TableBoundary, detect_tables

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def transform(
    sheet_xml_or_data: str | SheetData,
    shared_strings: list[str],
    hint: TableLayoutHint | None = None,
) -> str:
    """Transform sheet XML into semantic table XML.

    Args:
        sheet_xml_or_data: Raw XML string or pre-scanned SheetData.
        shared_strings: Parsed shared string table.
        hint: Optional layout hint for explicit structure.

    Returns:
        Semantic table XML string.
    """
    try:
        if isinstance(sheet_xml_or_data, SheetData):
            data = sheet_xml_or_data
        else:
            data = scan_sheet(sheet_xml_or_data, shared_strings)

        grid = _build_grid(data)

        all_rows = sorted(grid.keys())
        if not all_rows:
            return "<semantic-table/>"

        if hint is not None:
            max_row = all_rows[-1]
            max_col = max(
                (col for row in grid.values() for col in row.keys()),
                default=1,
            )
            resolved = hint.resolve_defaults(max_row, max_col)
            return _build_with_hint(grid, all_rows, resolved)
        else:
            return _build_auto_detect(grid, all_rows, data.merge_map)
    except Exception as e:
        logger.warning("Failed to transform sheet to semantic XML: %s", e, exc_info=True)
        return f"<semantic-table><error>{e}</error></semantic-table>"


# ── Hint-based transformation ──


def _build_with_hint(
    grid: dict[int, dict[int, str]],
    all_rows: list[int],
    hint: TableLayoutHint,
) -> str:
    root = etree.Element("semantic-table")

    # 1. Title extraction (deduplicated, non-empty text only)
    if hint.title_rows:
        unique_parts: dict[str, None] = OrderedDict()  # ordered set
        for title_row in sorted(hint.title_rows):
            row = grid.get(title_row)
            if not row:
                continue
            for val in row.values():
                if val and val.strip():
                    unique_parts[val.strip()] = None
        if unique_parts:
            title_el = etree.SubElement(root, "title")
            title_el.text = " ".join(unique_parts.keys())

    # 2. Determine header rows
    if hint.has_header_range:
        header_rows = list(range(hint.header_start_row, hint.header_end_row + 1))
        header_start_col = hint.header_start_col
        header_end_col = hint.header_end_col
        first_data_row = hint.header_end_row + 1
    else:
        auto = _hint_fallback_detect_headers(grid, all_rows, hint.title_rows)
        header_rows = auto[0]
        first_data_row = auto[1]
        header_start_col = -1
        header_end_col = -1

    # 3. Build tag names from header rows
    all_data_cols: set[int] = set()
    for row in grid.values():
        all_data_cols.update(row.keys())

    meta_col = hint.row_meta_col_num if hint.has_row_meta_col else -1
    column_tags: dict[int, str] = OrderedDict()

    for col in sorted(all_data_cols):
        if col == meta_col:
            continue
        if header_start_col > 0 and (col < header_start_col or col > header_end_col):
            continue

        parts: list[str] = []
        for hr in header_rows:
            row = grid.get(hr)
            if row:
                val = row.get(col)
                if val and val.strip():
                    cleaned = val.strip()
                    if not parts or parts[-1].lower() != cleaned.lower():
                        parts.append(cleaned)
        if parts:
            column_tags[col] = _to_tag_name("_".join(parts))

    # Row meta attribute name
    meta_attr_name: str | None = None
    if meta_col > 0:
        meta_parts: list[str] = []
        for hr in header_rows:
            row = grid.get(hr)
            if row:
                val = row.get(meta_col)
                if val and val.strip():
                    meta_parts.append(val.strip())
        meta_attr_name = _to_tag_name("_".join(meta_parts)) if meta_parts else "label"

    # 4. Schema
    schema_el = etree.SubElement(root, "schema")
    if meta_col > 0 and meta_attr_name:
        meta_el = etree.SubElement(schema_el, "row-key")
        meta_el.set("index", str(meta_col))
        meta_el.set("attribute", meta_attr_name)
    for col, tag in column_tags.items():
        col_el = etree.SubElement(schema_el, "column")
        col_el.set("index", str(col))
        col_el.set("tag", tag)

    # 5. Records
    records_el = etree.SubElement(root, "records")
    record_count = 0

    skip_rows = set(hint.title_rows) | set(header_rows)

    for row_idx in all_rows:
        if row_idx in skip_rows or row_idx < first_data_row:
            continue
        row_data = grid.get(row_idx)
        if not row_data:
            continue

        has_tagged = any(col in column_tags for col in row_data)
        if not has_tagged:
            continue

        record_el = etree.SubElement(records_el, "record")
        record_el.set("row", str(row_idx))

        if meta_col > 0 and meta_attr_name:
            meta_value = row_data.get(meta_col)
            if meta_value and meta_value.strip():
                record_el.set(meta_attr_name, meta_value.strip())

        has_children = False
        for col, tag in column_tags.items():
            value = row_data.get(col)
            if value and value.strip():
                field_el = etree.SubElement(record_el, tag)
                field_el.text = value.strip()
                has_children = True

        if has_children:
            record_count += 1
        else:
            records_el.remove(record_el)

    records_el.set("count", str(record_count))

    return etree.tostring(root, pretty_print=True, xml_declaration=False, encoding="unicode")


# ── Auto-detect transformation (structural-anchor based) ──


def _build_auto_detect(
    grid: dict[int, dict[int, str]],
    all_rows: list[int],
    merge_map: dict[str, str],
) -> str:
    """Detect table regions via structural anchors, then build semantic XML."""
    tables = detect_tables(grid, merge_map)

    if not tables:
        return "<semantic-table/>"

    if len(tables) == 1:
        return _build_table_xml(grid, tables[0])

    # Multiple tables → wrap in <sheet>
    root = etree.Element("sheet")
    for i, table in enumerate(tables):
        table_str = _build_table_xml(grid, table)
        table_el = etree.fromstring(table_str.encode("utf-8"))
        table_el.set("index", str(i + 1))
        root.append(table_el)

    return etree.tostring(root, pretty_print=True, xml_declaration=False, encoding="unicode")


def _build_table_xml(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
) -> str:
    """Generate ``<semantic-table>`` XML for a single detected table."""
    root = etree.Element("semantic-table")

    # 1. Title
    if table.title_rows:
        unique_parts: dict[str, None] = OrderedDict()
        for tr in sorted(table.title_rows):
            row = grid.get(tr)
            if not row:
                continue
            for val in row.values():
                if val and val.strip():
                    unique_parts[val.strip()] = None
        if unique_parts:
            title_el = etree.SubElement(root, "title")
            title_el.text = " ".join(unique_parts.keys())

    # 2. Column tags from header rows
    all_data_cols: set[int] = set()
    for r in range(table.min_row, table.max_row + 1):
        for c in grid.get(r, {}).keys():
            if table.min_col <= c <= table.max_col:
                all_data_cols.add(c)

    meta_col = table.row_meta_col
    column_tags: dict[int, str] = OrderedDict()

    for col in sorted(all_data_cols):
        if col == meta_col:
            continue
        parts: list[str] = []
        for hr in table.header_rows:
            row = grid.get(hr)
            if row:
                val = row.get(col)
                if val and val.strip():
                    cleaned = val.strip()
                    if not parts or parts[-1].lower() != cleaned.lower():
                        parts.append(cleaned)
        if parts:
            column_tags[col] = _to_tag_name("_".join(parts))

    # If no headers detected, use positional tag names
    if not column_tags and all_data_cols:
        for col in sorted(all_data_cols):
            if col == meta_col:
                continue
            column_tags[col] = f"col_{col}"

    # 3. Row meta attribute name
    meta_attr_name: str | None = None
    if meta_col > 0:
        meta_parts: list[str] = []
        for hr in table.header_rows:
            row = grid.get(hr)
            if row:
                val = row.get(meta_col)
                if val and val.strip():
                    cleaned = val.strip()
                    if not meta_parts or meta_parts[-1].lower() != cleaned.lower():
                        meta_parts.append(cleaned)
        meta_attr_name = _to_tag_name("_".join(meta_parts)) if meta_parts else "label"

    # 4. Schema
    schema_el = etree.SubElement(root, "schema")
    if meta_col > 0 and meta_attr_name:
        meta_el = etree.SubElement(schema_el, "row-key")
        meta_el.set("index", str(meta_col))
        meta_el.set("attribute", meta_attr_name)
    for col, tag in column_tags.items():
        col_el = etree.SubElement(schema_el, "column")
        col_el.set("index", str(col))
        col_el.set("tag", tag)

    # 5. Records
    records_el = etree.SubElement(root, "records")
    record_count = 0

    skip_rows = set(table.title_rows) | set(table.header_rows) | set(table.section_rows)

    for row_idx in range(table.min_row, table.max_row + 1):
        if row_idx in skip_rows or row_idx < table.data_start_row:
            continue
        row_data = grid.get(row_idx)
        if not row_data:
            continue

        has_tagged = any(
            col in column_tags
            for col in row_data
            if table.min_col <= col <= table.max_col
        )
        if not has_tagged:
            continue

        record_el = etree.SubElement(records_el, "record")
        record_el.set("row", str(row_idx))

        if meta_col > 0 and meta_attr_name:
            meta_value = row_data.get(meta_col)
            if meta_value and meta_value.strip():
                record_el.set(meta_attr_name, meta_value.strip())

        has_children = False
        for col, tag in column_tags.items():
            value = row_data.get(col)
            if value and value.strip():
                field_el = etree.SubElement(record_el, tag)
                field_el.text = value.strip()
                has_children = True

        if has_children:
            record_count += 1
        else:
            records_el.remove(record_el)

    records_el.set("count", str(record_count))

    return etree.tostring(root, pretty_print=True, xml_declaration=False, encoding="unicode")


def _hint_fallback_detect_headers(
    grid: dict[int, dict[int, str]],
    all_rows: list[int],
    skip_rows: frozenset[int],
) -> tuple[list[int], int]:
    """Simple text/numeric ratio heuristic used when a hint lacks header_range."""
    header_rows: list[int] = []
    first_data_row = -1

    for row_idx in all_rows:
        if row_idx in skip_rows:
            continue
        row = grid.get(row_idx)
        if not row:
            continue

        text_cells = 0
        numeric_cells = 0
        for val in row.values():
            if _is_numeric(val):
                numeric_cells += 1
            else:
                text_cells += 1

        total = text_cells + numeric_cells
        if total >= 3 and numeric_cells / total >= 0.6:
            first_data_row = row_idx
            break
        if total >= 2:
            header_rows.append(row_idx)

    if first_data_row == -1:
        first_data_row = all_rows[-1] + 1

    header_rows = [r for r in header_rows if r < first_data_row]
    return header_rows, first_data_row


def _build_column_tags(
    grid: dict[int, dict[int, str]],
    header_rows: list[int],
    all_cols: list[int],
) -> dict[int, str]:
    tags: dict[int, str] = OrderedDict()
    for col in all_cols:
        parts: list[str] = []
        for hr in header_rows:
            row = grid.get(hr)
            if row:
                val = row.get(col)
                if val and val.strip():
                    cleaned = val.strip()
                    if not parts or parts[-1].lower() != cleaned.lower():
                        parts.append(cleaned)
        if parts:
            tags[col] = _to_tag_name("_".join(parts))
    return tags


# ── Grid building ──


def _build_grid(data: SheetData) -> dict[int, dict[int, str]]:
    """Build a {row: {col: value}} grid from pre-scanned SheetData."""
    grid: dict[int, dict[int, str]] = {}

    # Apply merge map: expand origin values to all merged coordinates
    all_coords = set(data.raw_values.keys()) | set(data.merge_map.keys())
    for coord in all_coords:
        origin = data.merge_map.get(coord, coord)
        value = data.raw_values.get(origin)
        if not value or not value.strip():
            continue

        parts = coord.split(":")
        col = int(parts[0])
        row = int(parts[1])
        grid.setdefault(row, {})[col] = value

    return grid


# ── Utilities ──


def _to_tag_name(text: str) -> str:
    """Convert text to a valid XML tag name."""
    s = text.strip().lower()
    s = re.sub(r"[\r\n/]+", " ", s)
    s = re.sub(r"[^\w\s]", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"_+", "_", s)
    s = s.strip("_")
    if s and s[0].isdigit():
        s = "col_" + s
    return s or "unknown"


def _is_numeric(s: str | None) -> bool:
    if not s or not s.strip():
        return False
    try:
        float(s.strip())
        return True
    except ValueError:
        return False


