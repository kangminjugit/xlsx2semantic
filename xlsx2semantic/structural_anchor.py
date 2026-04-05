"""Structural-anchor-based table detection.

Inspired by SpreadsheetLLM (Chen et al., 2024). Detects table boundaries by:
1. Clustering non-empty cells into connected regions (row/col proximity)
2. Using merge patterns and row heterogeneity to find headers/titles
3. Auto-detecting row-key columns from text/numeric distribution

Replaces manual layout hints for most common spreadsheet layouts.
"""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass

from xlsx2semantic.merge_analyzer import analyze_merges

# ── Tuning constants ──

# Max empty rows allowed within a single table cluster.
_ROW_GAP = 1
# Max empty columns allowed within a single table cluster.
_COL_GAP = 1
# Minimum rows (with data) to be considered a table.
_MIN_ROWS = 2
# Minimum columns to be considered a table.
_MIN_COLS = 2


@dataclass(frozen=True)
class TableBoundary:
    """A detected table region with structural metadata."""

    min_row: int
    max_row: int
    min_col: int
    max_col: int
    title_rows: frozenset[int] = frozenset()
    header_rows: tuple[int, ...] = ()
    data_start_row: int = -1
    section_rows: frozenset[int] = frozenset()


# ── Public API ──


def detect_tables(
    grid: dict[int, dict[int, str]],
    merge_map: dict[str, str],
    header_row_count: int = 1,
) -> list[TableBoundary]:
    """Detect all table regions in a sheet grid.

    Args:
        grid: ``{row: {col: value}}`` mapping of non-empty cells
              (merged cells already expanded).
        merge_map: Raw merge coordinate mappings from sheet scanner.

    Returns:
        List of :class:`TableBoundary`, ordered top-to-bottom then left-to-right.
    """
    if not grid:
        return []

    all_rows = sorted(grid.keys())

    # Step 1 — cluster rows by vertical proximity
    row_clusters = _cluster_rows(all_rows)

    # Step 2 — within each cluster, split by column gaps
    raw_regions: list[tuple[list[int], int, int]] = []
    for rows in row_clusters:
        for min_col, max_col in _split_by_columns(grid, rows):
            raw_regions.append((rows, min_col, max_col))

    # Step 3 — analyse each region
    tables: list[TableBoundary] = []
    for rows, min_col, max_col in raw_regions:
        sub_grid = _extract_sub_grid(grid, rows, min_col, max_col)
        if not sub_grid:
            continue

        actual_rows = sorted(sub_grid.keys())
        actual_cols: set[int] = set()
        for rd in sub_grid.values():
            actual_cols.update(rd.keys())

        if len(actual_rows) < _MIN_ROWS or len(actual_cols) < _MIN_COLS:
            continue

        sub_merge = _extract_sub_merge(merge_map, actual_rows, min_col, max_col)
        boundary = _analyze_region(sub_grid, sub_merge, actual_rows, min_col, max_col, header_row_count)
        if boundary is not None:
            tables.append(boundary)

    return tables


# ── Row / column clustering ──


def _cluster_rows(all_rows: list[int]) -> list[list[int]]:
    """Group rows into clusters separated by gaps > ``_ROW_GAP``."""
    if not all_rows:
        return []

    clusters: list[list[int]] = []
    current = [all_rows[0]]

    for i in range(1, len(all_rows)):
        if all_rows[i] - all_rows[i - 1] - 1 > _ROW_GAP:
            clusters.append(current)
            current = [all_rows[i]]
        else:
            current.append(all_rows[i])

    clusters.append(current)
    return clusters


def _split_by_columns(
    grid: dict[int, dict[int, str]],
    rows: list[int],
) -> list[tuple[int, int]]:
    """Find distinct column groups within *rows*."""
    all_cols: set[int] = set()
    for r in rows:
        all_cols.update(grid.get(r, {}).keys())

    if not all_cols:
        return []

    sorted_cols = sorted(all_cols)
    groups: list[tuple[int, int]] = []
    start = prev = sorted_cols[0]

    for i in range(1, len(sorted_cols)):
        if sorted_cols[i] - prev - 1 > _COL_GAP:
            groups.append((start, prev))
            start = sorted_cols[i]
        prev = sorted_cols[i]

    groups.append((start, prev))
    return groups


# ── Sub-grid extraction ──


def _extract_sub_grid(
    grid: dict[int, dict[int, str]],
    rows: list[int],
    min_col: int,
    max_col: int,
) -> dict[int, dict[int, str]]:
    sub: dict[int, dict[int, str]] = {}
    for r in rows:
        filtered = {
            c: v for c, v in grid.get(r, {}).items() if min_col <= c <= max_col
        }
        if filtered:
            sub[r] = filtered
    return sub


def _extract_sub_merge(
    merge_map: dict[str, str],
    rows: list[int],
    min_col: int,
    max_col: int,
) -> dict[str, str]:
    row_set = set(range(min(rows), max(rows) + 1)) if rows else set()
    sub: dict[str, str] = {}
    for coord, origin in merge_map.items():
        parts = coord.split(":")
        c, r = int(parts[0]), int(parts[1])
        if r in row_set and min_col <= c <= max_col:
            sub[coord] = origin
    return sub


# ── Region analysis ──


def _analyze_region(
    sub_grid: dict[int, dict[int, str]],
    sub_merge: dict[str, str],
    rows: list[int],
    min_col: int,
    max_col: int,
    header_row_count: int = 1,
) -> TableBoundary | None:
    if not sub_grid or not rows:
        return None

    # Use merges only for title detection (full-width merges are reliable).
    # Header/data boundary is always determined by heterogeneity because
    # merge patterns alone cannot distinguish hierarchical headers from
    # vertically-merged group labels (e.g. subtotal rows).
    merge_result = analyze_merges(sub_merge, sub_grid) if sub_merge else None

    merge_title_rows = merge_result.title_rows if merge_result else frozenset()

    # Heterogeneity-based header detection (skip merge-detected title rows)
    skip = merge_title_rows
    het_title, header_rows, data_start = _detect_headers_by_heterogeneity(
        sub_grid, rows, skip, header_row_count,
    )

    # Extend header block downward: rows whose cells are vertically merged
    # from a header-origin cell are still part of the header, not data.
    if sub_merge and header_rows:
        header_rows, data_start = _extend_headers_via_merges(
            header_rows, data_start, rows, sub_merge,
        )

    title_rows = merge_title_rows | het_title
    section_rows: frozenset[int] = (
        merge_result.section_rows if merge_result else frozenset()
    )

    return TableBoundary(
        min_row=rows[0],
        max_row=rows[-1],
        min_col=min_col,
        max_col=max_col,
        title_rows=title_rows,
        header_rows=header_rows,
        data_start_row=data_start,
        section_rows=section_rows,
    )


# ── Header detection ──


def _detect_headers_by_heterogeneity(
    grid: dict[int, dict[int, str]],
    all_rows: list[int],
    skip_rows: frozenset[int] = frozenset(),
    header_row_count: int = 1,
) -> tuple[frozenset[int], tuple[int, ...], int]:
    """Return ``(title_rows, header_rows, data_start_row)``.

    Title rows are single-value rows at the very top (before any header).
    The next ``header_row_count`` non-empty rows become the header rows.
    Everything after that is data.
    """
    if not all_rows:
        return frozenset(), (), -1

    title_rows: set[int] = set()
    header_rows: list[int] = []
    data_start = -1

    for row_idx in all_rows:
        if row_idx in skip_rows:
            continue

        row_data = grid.get(row_idx, {})
        if not row_data:
            continue

        if not header_rows:
            # Before collecting any header, single-value rows are titles
            unique_vals = {v.strip() for v in row_data.values() if v and v.strip()}
            if len(unique_vals) <= 1:
                title_rows.add(row_idx)
                continue

        header_rows.append(row_idx)

        if len(header_rows) >= header_row_count:
            remaining = [
                r for r in all_rows
                if r > row_idx and r not in skip_rows and grid.get(r)
            ]
            data_start = remaining[0] if remaining else row_idx + 1
            break

    if data_start == -1:
        data_start = all_rows[-1] + 1

    return frozenset(title_rows), tuple(header_rows), data_start


def _extend_headers_via_merges(
    header_rows: tuple[int, ...],
    data_start: int,
    all_rows: list[int],
    sub_merge: dict[str, str],
) -> tuple[tuple[int, ...], int]:
    """Extend header_rows to include rows covered by vertical merges from header cells.

    When a header cell (e.g. "대여소번호" at row 1) is vertically merged down
    to rows 2, 3, 4, those rows should still be treated as header rows, not data.
    """
    # Build a set of rows that are origins of merges within the current header block
    header_row_set = set(header_rows)

    # Build a lookup: destination_row → {origin_rows}
    dest_to_origins: dict[int, set[int]] = {}
    for coord, origin in sub_merge.items():
        dest_row = int(coord.split(":")[1])
        origin_row = int(origin.split(":")[1])
        if dest_row != origin_row:  # skip self-references
            dest_to_origins.setdefault(dest_row, set()).add(origin_row)

    # Greedily extend: if the candidate data_start row has cells merged FROM
    # a header row, absorb it into the header block and advance data_start.
    extended = list(header_rows)
    remaining_rows = [r for r in all_rows if r >= data_start]

    for candidate in remaining_rows:
        origins = dest_to_origins.get(candidate, set())
        if origins & header_row_set:
            # At least one cell in this row originates from a header cell
            extended.append(candidate)
            header_row_set.add(candidate)
        else:
            data_start = candidate
            break
    else:
        data_start = all_rows[-1] + 1 if all_rows else data_start

    return tuple(extended), data_start
