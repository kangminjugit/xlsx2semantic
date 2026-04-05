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
from xlsx2semantic.semantic_scorer import SemanticDecision, decide_semantics
from xlsx2semantic.sheet_scanner import SheetData, scan_sheet
from xlsx2semantic.structural_anchor import TableBoundary, detect_tables

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def transform(
    sheet_xml_or_data: str | SheetData,
    shared_strings: list[str],
    hint: TableLayoutHint | None = None,
    include_trace: bool = False,
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

        if hint is not None and hint.has_structural_hints:
            max_row = all_rows[-1]
            max_col = max(
                (col for row in grid.values() for col in row.keys()),
                default=1,
            )
            resolved = hint.resolve_defaults(max_row, max_col)
            return _build_with_hint(grid, all_rows, resolved, include_trace=include_trace)
        else:
            count = hint.header_row_count if hint is not None else 1
            return _build_auto_detect(grid, all_rows, data.merge_map, count, include_trace=include_trace)
    except Exception as e:
        logger.warning("Failed to transform sheet to semantic XML: %s", e, exc_info=True)
        return f"<semantic-table><error>{e}</error></semantic-table>"


# ── Hint-based transformation ──


def _build_with_hint(
    grid: dict[int, dict[int, str]],
    all_rows: list[int],
    hint: TableLayoutHint,
    include_trace: bool = False,
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
        auto = _hint_fallback_detect_headers(grid, all_rows, hint.title_rows, hint.header_row_count)
        header_rows = auto[0]
        first_data_row = auto[1]
        header_start_col = -1
        header_end_col = -1

    # 3. Build tag names from header rows
    all_data_cols: set[int] = set()
    for row in grid.values():
        all_data_cols.update(row.keys())

    meta_cols: set[int] = set(hint.row_meta_col_nums) if hint.has_row_meta_col else set()

    eligible_cols = [
        col for col in sorted(all_data_cols)
        if col not in meta_cols
        and not (header_start_col > 0 and (col < header_start_col or col > header_end_col))
    ]
    column_tags = _build_min_unique_column_tags(grid, header_rows, eligible_cols)

    # Row meta attribute names (one per meta column)
    meta_attr_names: dict[int, str] = {}
    for mc in sorted(meta_cols):
        meta_parts: list[str] = []
        for hr in header_rows:
            row = grid.get(hr)
            if row:
                val = row.get(mc)
                if val and val.strip():
                    meta_parts.append(val.strip())
        meta_attr_names[mc] = _to_tag_name("_".join(meta_parts)) if meta_parts else f"col_{mc}"

    # 4. Schema
    schema_el = etree.SubElement(root, "schema")
    for mc, attr_name in meta_attr_names.items():
        meta_el = etree.SubElement(schema_el, "row-key")
        meta_el.set("index", str(mc))
        meta_el.set("attribute", attr_name)
    for col, tag in column_tags.items():
        col_el = etree.SubElement(schema_el, "column")
        col_el.set("index", str(col))
        col_el.set("tag", tag)

    # 5. Description rows (full-width merged cells in the data area)
    skip_rows = set(hint.title_rows) | set(header_rows)
    all_cols = sorted(all_data_cols)
    min_col = all_cols[0] if all_cols else 1
    max_col = all_cols[-1] if all_cols else 1
    seen_descriptions: set[str] = set()

    for row_idx in all_rows:
        if row_idx in skip_rows or row_idx < first_data_row:
            continue
        row_data = grid.get(row_idx)
        if not row_data:
            continue
        if _is_merged_description_row(row_data, min_col, max_col):
            text = next(iter(row_data.values())).strip()
            if text not in seen_descriptions:
                desc_el = etree.SubElement(root, "description")
                desc_el.text = text
                seen_descriptions.add(text)
            skip_rows.add(row_idx)

    # 6. Records
    records_el = etree.SubElement(root, "records")
    record_count = 0

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

        for mc, attr_name in meta_attr_names.items():
            meta_value = row_data.get(mc)
            if meta_value and meta_value.strip():
                record_el.set(attr_name, _normalize_cell_value(meta_value))

        has_children = False
        for col, tag in column_tags.items():
            value = row_data.get(col)
            if value and value.strip():
                field_el = etree.SubElement(record_el, tag)
                field_el.text = _normalize_cell_value(value)
                has_children = True

        if has_children:
            record_count += 1
        else:
            records_el.remove(record_el)

    records_el.set("count", str(record_count))

    if include_trace:
        _append_hint_trace(root, hint, header_rows, first_data_row)

    return etree.tostring(root, pretty_print=True, xml_declaration=False, encoding="unicode")


# ── Auto-detect transformation (structural-anchor based) ──


def _build_auto_detect(
    grid: dict[int, dict[int, str]],
    all_rows: list[int],
    merge_map: dict[str, str],
    header_row_count: int = 1,
    include_trace: bool = False,
) -> str:
    """Detect table regions via structural anchors, then build semantic XML."""
    scored_tables = []
    for table in detect_tables(grid, merge_map, header_row_count):
        decision = decide_semantics(grid, table, header_row_count)
        scored_tables.append((_apply_semantic_decision(table, decision), decision))

    if not scored_tables:
        return "<semantic-table/>"

    # Collect title/descriptions from empty tables and forward them to the next non-empty one
    pending_title: str | None = None
    pending_descriptions: list[str] = []
    non_empty_with_context: list[tuple[TableBoundary, str | None, list[str]]] = []

    for table, decision in scored_tables:
        if _has_records(grid, table):
            non_empty_with_context.append((table, decision, pending_title, list(pending_descriptions)))
            pending_title = None
            pending_descriptions = []
        else:
            t, descs = _collect_empty_table_info(grid, table)
            if t:
                pending_title = t
            pending_descriptions.extend(descs)

    if not non_empty_with_context:
        return "<semantic-table/>"

    if len(non_empty_with_context) == 1:
        table, decision, inh_title, inh_descs = non_empty_with_context[0]
        return _build_table_xml(grid, table, inh_title, inh_descs, decision if include_trace else None)

    # Multiple tables → wrap in <sheet>
    root = etree.Element("sheet")
    for index, (table, decision, inh_title, inh_descs) in enumerate(non_empty_with_context, start=1):
        table_str = _build_table_xml(grid, table, inh_title, inh_descs, decision if include_trace else None)
        table_el = etree.fromstring(table_str.encode("utf-8"))
        table_el.set("index", str(index))
        root.append(table_el)

    return etree.tostring(root, pretty_print=True, xml_declaration=False, encoding="unicode")


def _build_table_xml(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    inherited_title: str | None = None,
    inherited_descriptions: list[str] | None = None,
    decision: SemanticDecision | None = None,
) -> str:
    """Generate ``<semantic-table>`` XML for a single detected table."""
    root = etree.Element("semantic-table")

    # 1. Title + descriptions from title_rows
    # Rows with multi-line values are descriptions masquerading as titles.
    own_title: str | None = None
    own_descriptions: list[str] = []
    if table.title_rows:
        title_parts: dict[str, None] = OrderedDict()
        for tr in sorted(table.title_rows):
            row = grid.get(tr)
            if not row:
                continue
            unique_vals = list(dict.fromkeys(v.strip() for v in row.values() if v and v.strip()))
            if not unique_vals:
                continue
            val = unique_vals[0]
            if "\n" in val:
                own_descriptions.append(val)
            else:
                title_parts[val] = None
        if title_parts:
            own_title = " ".join(title_parts.keys())

    title_text = inherited_title or own_title
    if title_text:
        title_el = etree.SubElement(root, "title")
        title_el.text = title_text

    # Descriptions (inherited first, then own)
    seen_descs: set[str] = set()
    for desc in (inherited_descriptions or []) + own_descriptions:
        if desc not in seen_descs:
            desc_el = etree.SubElement(root, "description")
            desc_el.text = desc
            seen_descs.add(desc)

    # 2. Column tags from header rows
    all_data_cols: set[int] = set()
    for r in range(table.min_row, table.max_row + 1):
        for c in grid.get(r, {}).keys():
            if table.min_col <= c <= table.max_col:
                all_data_cols.add(c)

    column_tags = _build_min_unique_column_tags(grid, list(table.header_rows), sorted(all_data_cols))

    if not column_tags and all_data_cols:
        for col in sorted(all_data_cols):
            column_tags[col] = f"col_{col}"

    # 3. Schema
    schema_el = etree.SubElement(root, "schema")
    for col, tag in column_tags.items():
        col_el = etree.SubElement(schema_el, "column")
        col_el.set("index", str(col))
        col_el.set("tag", tag)

    # 4. Description rows (full-width merged cells in the data area)
    skip_rows = set(table.title_rows) | set(table.header_rows) | set(table.section_rows)
    seen_descriptions: set[str] = set()

    # Build the set of values that appear in header rows (for repeat-header detection)
    header_value_set: set[str] = set()
    for hr in table.header_rows:
        hr_data = grid.get(hr, {})
        for v in hr_data.values():
            if v and v.strip():
                header_value_set.add(v.strip())

    for row_idx in range(table.min_row, table.max_row + 1):
        if row_idx in skip_rows or row_idx < table.data_start_row:
            continue
        row_data = grid.get(row_idx)
        if not row_data:
            continue
        if _is_merged_description_row(row_data, table.min_col, table.max_col):
            text = next(iter(row_data.values())).strip()
            if text not in seen_descriptions:
                desc_el = etree.SubElement(root, "description")
                desc_el.text = text
                seen_descriptions.add(text)
            skip_rows.add(row_idx)
            continue
        # Repeated header rows: data row whose non-empty values are all header labels
        if header_value_set:
            row_vals = {v.strip() for v in row_data.values() if v and v.strip()}
            if len(row_vals) >= 2 and row_vals.issubset(header_value_set):
                skip_rows.add(row_idx)

    # 5. Records
    records_el = etree.SubElement(root, "records")
    record_count = 0

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

        has_children = False
        for col, tag in column_tags.items():
            value = row_data.get(col)
            if value and value.strip():
                field_el = etree.SubElement(record_el, tag)
                field_el.text = _normalize_cell_value(value)
                has_children = True

        if has_children:
            record_count += 1
        else:
            records_el.remove(record_el)

    records_el.set("count", str(record_count))

    if decision is not None:
        _append_auto_trace(root, decision)

    return etree.tostring(root, pretty_print=True, xml_declaration=False, encoding="unicode")


def _apply_semantic_decision(
    table: TableBoundary,
    decision: SemanticDecision,
) -> TableBoundary:
    return TableBoundary(
        min_row=table.min_row,
        max_row=table.max_row,
        min_col=table.min_col,
        max_col=table.max_col,
        title_rows=frozenset(decision.title_rows),
        header_rows=decision.header_rows,
        data_start_row=decision.data_start_row,
        section_rows=table.section_rows,
    )


def _append_hint_trace(
    root: etree._Element,
    hint: TableLayoutHint,
    header_rows: list[int],
    first_data_row: int,
) -> None:
    trace_el = etree.SubElement(root, "trace")
    trace_el.set("mode", "hint")
    trace_el.set("data-start-row", str(first_data_row))

    selected_el = etree.SubElement(trace_el, "selected")
    selected_el.set("header-rows", ",".join(str(row) for row in header_rows))
    if hint.title_rows:
        selected_el.set("title-rows", ",".join(str(row) for row in sorted(hint.title_rows)))
    if hint.has_row_meta_col:
        selected_el.set("row-meta-cols", ",".join(str(col) for col in hint.row_meta_col_nums))


def _append_auto_trace(
    root: etree._Element,
    decision: SemanticDecision,
) -> None:
    trace_el = etree.SubElement(root, "trace")
    trace_el.set("mode", "auto")
    trace_el.set("confidence", f"{decision.confidence:.3f}")

    selected_el = etree.SubElement(trace_el, "selected")
    selected_el.set("data-start-row", str(decision.data_start_row))
    selected_el.set("header-rows", ",".join(str(row) for row in decision.header_rows))
    if decision.title_rows:
        selected_el.set("title-rows", ",".join(str(row) for row in decision.title_rows))

    for evidence in decision.evidences:
        evidence_el = etree.SubElement(trace_el, "evidence")
        evidence_el.set("name", evidence.name)
        evidence_el.set("score", f"{evidence.score:.3f}")

    for row_trace in decision.row_traces:
        row_el = etree.SubElement(trace_el, "row")
        row_el.set("index", str(row_trace.row_idx))
        row_el.set("role", row_trace.role)
        row_el.set("title-score", f"{row_trace.title_score:.3f}")
        row_el.set("header-score", f"{row_trace.header_score:.3f}")
        row_el.set("data-score", f"{row_trace.data_score:.3f}")


def _hint_fallback_detect_headers(
    grid: dict[int, dict[int, str]],
    all_rows: list[int],
    skip_rows: frozenset[int],
    header_row_count: int = 1,
) -> tuple[list[int], int]:
    """Take the first ``header_row_count`` non-skip, non-empty rows as headers."""
    header_rows: list[int] = []

    for row_idx in all_rows:
        if row_idx in skip_rows:
            continue
        if not grid.get(row_idx):
            continue
        header_rows.append(row_idx)
        if len(header_rows) >= header_row_count:
            break

    remaining = [r for r in all_rows if r not in skip_rows and (not header_rows or r > header_rows[-1]) and grid.get(r)]
    first_data_row = remaining[0] if remaining else (all_rows[-1] + 1)
    return header_rows, first_data_row


def _build_min_unique_column_tags(
    grid: dict[int, dict[int, str]],
    header_rows: list[int],
    cols: list[int],
) -> dict[int, str]:
    """Build column tag names using the minimum header levels needed for uniqueness.

    Starts from the most specific (bottom) header row and adds parent levels
    only until the tag is unique across all columns.
    """
    # Collect header parts per column (top-to-bottom, consecutive duplicates removed)
    col_parts: dict[int, list[str]] = {}
    for col in cols:
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
            col_parts[col] = parts

    if not col_parts:
        return {}

    sorted_cols = sorted(col_parts.keys())
    depth = {col: 1 for col in sorted_cols}

    while True:
        candidates = {
            col: _to_tag_name("_".join(col_parts[col][-min(depth[col], len(col_parts[col])):]))
            for col in sorted_cols
        }

        tag_to_cols: dict[str, list[int]] = {}
        for col, tag in candidates.items():
            tag_to_cols.setdefault(tag, []).append(col)

        dupe_cols = {
            col
            for cols_with_tag in tag_to_cols.values()
            if len(cols_with_tag) > 1
            for col in cols_with_tag
        }

        if not dupe_cols:
            break

        progress = False
        for col in dupe_cols:
            if depth[col] < len(col_parts[col]):
                depth[col] += 1
                progress = True

        if not progress:
            break

    result: dict[int, str] = OrderedDict()
    used: set[str] = set()
    for col in sorted_cols:
        d = min(depth[col], len(col_parts[col]))
        tag = _to_tag_name("_".join(col_parts[col][-d:]))
        if tag in used:
            tag = f"{tag}_{col}"
        used.add(tag)
        result[col] = tag

    return result


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


def _normalize_cell_value(text: str) -> str:
    """Normalize cell value for record output: remove newlines."""
    return re.sub(r"\s*\n\s*", "", text).strip()


def _to_tag_name(text: str) -> str:
    """Convert text to a valid XML tag name."""
    s = text.strip().lower()
    s = re.sub(r"[\r\n/]+", " ", s)
    s = re.sub(r"\.", "_", s)
    s = re.sub(r"[^\w\s]", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"_+", "_", s)
    s = s.strip("_")
    if s and s[0].isdigit():
        s = "col_" + s
    return s or "unknown"


def _has_records(grid: dict[int, dict[int, str]], table: TableBoundary) -> bool:
    """Return True if the table has at least one non-description data row."""
    skip = set(table.title_rows) | set(table.header_rows) | set(table.section_rows)
    for row_idx in range(table.min_row, table.max_row + 1):
        if row_idx in skip or row_idx < table.data_start_row:
            continue
        row_data = grid.get(row_idx)
        if not row_data:
            continue
        if not _is_merged_description_row(row_data, table.min_col, table.max_col):
            return True
    return False


def _collect_empty_table_info(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
) -> tuple[str | None, list[str]]:
    """Extract title text and description texts from a 0-record table."""
    title_parts: dict[str, None] = OrderedDict()
    desc_from_titles: list[str] = []

    if table.title_rows:
        for tr in sorted(table.title_rows):
            row = grid.get(tr)
            if not row:
                continue
            unique_vals = list(dict.fromkeys(v.strip() for v in row.values() if v and v.strip()))
            if not unique_vals:
                continue
            val = unique_vals[0]
            if "\n" in val:
                desc_from_titles.append(val)
            else:
                title_parts[val] = None

    title_text = " ".join(title_parts.keys()) if title_parts else None

    skip = set(table.title_rows) | set(table.header_rows) | set(table.section_rows)
    descriptions: list[str] = []
    seen: set[str] = set()
    for row_idx in range(table.min_row, table.max_row + 1):
        if row_idx in skip or row_idx < table.data_start_row:
            continue
        row_data = grid.get(row_idx)
        if not row_data:
            continue
        if _is_merged_description_row(row_data, table.min_col, table.max_col):
            text = next(iter(row_data.values())).strip()
            if text not in seen:
                descriptions.append(text)
                seen.add(text)

    return title_text, desc_from_titles + descriptions


def _is_merged_description_row(
    row_data: dict[int, str],
    min_col: int,
    max_col: int,
) -> bool:
    """Return True if all non-empty cells in the column range share the same text value.

    This indicates a horizontally-merged description/subtitle cell that has been
    expanded across all columns by the merge map.
    """
    vals = [
        v.strip()
        for c, v in row_data.items()
        if min_col <= c <= max_col and v and v.strip()
    ]
    return len(vals) >= 2 and len(set(vals)) == 1 and not _is_numeric(vals[0])


def _is_numeric(s: str | None) -> bool:
    if not s or not s.strip():
        return False
    try:
        float(s.strip())
        return True
    except ValueError:
        return False
