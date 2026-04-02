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
    row_meta_col: int = -1
    section_rows: frozenset[int] = frozenset()


# ── Public API ──


def detect_tables(
    grid: dict[int, dict[int, str]],
    merge_map: dict[str, str],
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
        boundary = _analyze_region(sub_grid, sub_merge, actual_rows, min_col, max_col)
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
        sub_grid, rows, skip,
    )

    title_rows = merge_title_rows | het_title
    section_rows: frozenset[int] = (
        merge_result.section_rows if merge_result else frozenset()
    )

    row_meta = _detect_row_meta_col(sub_grid, header_rows, data_start, min_col, max_col)

    return TableBoundary(
        min_row=rows[0],
        max_row=rows[-1],
        min_col=min_col,
        max_col=max_col,
        title_rows=title_rows,
        header_rows=header_rows,
        data_start_row=data_start,
        row_meta_col=row_meta,
        section_rows=section_rows,
    )


# ── Heterogeneity-based header detection ──


def _detect_headers_by_heterogeneity(
    grid: dict[int, dict[int, str]],
    all_rows: list[int],
    skip_rows: frozenset[int] = frozenset(),
) -> tuple[frozenset[int], tuple[int, ...], int]:
    """Return ``(title_rows, header_rows, data_start_row)``."""
    if not all_rows:
        return frozenset(), (), -1

    header_candidates: list[int] = []
    data_start = -1

    for row_idx in all_rows:
        if row_idx in skip_rows:
            continue

        row_data = grid.get(row_idx, {})
        if not row_data:
            continue

        text_count, num_count = _text_num_counts(row_data)
        total = text_count + num_count
        if total == 0:
            continue

        num_ratio = num_count / total

        # Row with enough cells and ≥50 % numeric → first data row
        if total >= 2 and num_ratio >= 0.5:
            data_start = row_idx
            break

        header_candidates.append(row_idx)

    if data_start == -1:
        # No clear numeric data found — treat everything as one block
        if len(header_candidates) <= 1:
            # Probably all-text table; first row is data
            data_start = all_rows[0]
            header_candidates = []
        else:
            data_start = all_rows[-1] + 1

    # Separate titles from headers:
    # Title = single-value row at the very top before multi-cell header rows.
    title_rows: set[int] = set()
    remaining: list[int] = []
    for hr in header_candidates:
        row_data = grid.get(hr, {})
        unique_vals = set(v.strip() for v in row_data.values() if v and v.strip())
        if len(unique_vals) <= 1 and not remaining:
            title_rows.add(hr)
        else:
            remaining.append(hr)

    return frozenset(title_rows), tuple(remaining), data_start


def _text_num_counts(row_data: dict[int, str]) -> tuple[int, int]:
    text = num = 0
    for val in row_data.values():
        if not val or not val.strip():
            continue
        if _is_numeric(val.strip()):
            num += 1
        else:
            text += 1
    return text, num


# ── Row-meta column detection ──


def _detect_row_meta_col(
    grid: dict[int, dict[int, str]],
    header_rows: tuple[int, ...],
    data_start: int,
    min_col: int,
    max_col: int,
) -> int:
    """Find the leftmost mostly-text column in data rows (= row key)."""
    if data_start < 0:
        return -1

    header_set = set(header_rows)
    data_rows = [
        d for r, d in grid.items() if r >= data_start and r not in header_set
    ]
    if not data_rows:
        return -1

    col_text_ratio: dict[int, float] = {}
    for col in range(min_col, max_col + 1):
        text_n = total_n = 0
        for rd in data_rows:
            val = rd.get(col)
            if val and val.strip():
                total_n += 1
                if not _is_numeric(val.strip()):
                    text_n += 1
        if total_n > 0:
            col_text_ratio[col] = text_n / total_n

    if not col_text_ratio:
        return -1

    text_cols = [c for c, r in col_text_ratio.items() if r > 0.8]
    num_cols = [c for c, r in col_text_ratio.items() if r < 0.3]

    if text_cols and num_cols:
        return min(text_cols)

    return -1


# ── Helpers ──


def _is_numeric(s: str) -> bool:
    try:
        float(s)
        return True
    except ValueError:
        return False
