"""Merge-pattern analyzer for automatic table structure detection.

Analyzes merge cell patterns to identify:
- Title rows (full-width merges)
- Multi-level header rows (partial merges with sub-columns below)
- Section separators (mid-table full-width merges)

Core insight: merged cells are the strongest structural signal in Excel.
A horizontal merge with individual sub-cells below = hierarchical header.
A full-width merge = title or section divider.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class MergeRegion:
    """A single merged cell range reconstructed from merge_map."""

    origin_col: int
    origin_row: int
    end_col: int
    end_row: int

    @property
    def col_span(self) -> int:
        return self.end_col - self.origin_col + 1

    @property
    def row_span(self) -> int:
        return self.end_row - self.origin_row + 1


@dataclass
class TableRegion:
    """Detected table structure from merge analysis."""

    title_rows: frozenset[int] = frozenset()
    header_rows: list[int] = field(default_factory=list)
    data_start_row: int = -1
    section_rows: frozenset[int] = frozenset()


def reconstruct_merge_regions(merge_map: dict[str, str]) -> list[MergeRegion]:
    """Convert merge_map (coord->origin) back into MergeRegion objects."""
    origin_extents: dict[str, tuple[int, int, int, int]] = {}

    for coord, origin in merge_map.items():
        parts = coord.split(":")
        col, row = int(parts[0]), int(parts[1])

        if origin not in origin_extents:
            o = origin.split(":")
            oc, or_ = int(o[0]), int(o[1])
            origin_extents[origin] = (oc, or_, col, row)
        else:
            oc, or_, ec, er = origin_extents[origin]
            origin_extents[origin] = (oc, or_, max(ec, col), max(er, row))

    return [
        MergeRegion(origin_col=oc, origin_row=or_, end_col=ec, end_row=er)
        for oc, or_, ec, er in origin_extents.values()
    ]


def analyze_merges(
    merge_map: dict[str, str],
    grid: dict[int, dict[int, str]],
) -> TableRegion | None:
    """Analyze merge patterns to detect table structure.

    Returns None if no useful structure can be inferred from merges.
    """
    if not merge_map:
        return None

    regions = reconstruct_merge_regions(merge_map)
    if not regions:
        return None

    all_rows = sorted(grid.keys())
    if not all_rows:
        return None

    # Calculate data extent
    all_cols: set[int] = set()
    for row_data in grid.values():
        all_cols.update(row_data.keys())
    if not all_cols:
        return None

    data_width = max(all_cols) - min(all_cols) + 1

    # Group merge regions by origin row
    row_merges: dict[int, list[MergeRegion]] = {}
    for region in regions:
        row_merges.setdefault(region.origin_row, []).append(region)

    # Phase 1: Classify rows by merge pattern
    full_width_rows: set[int] = set()
    hierarchical_header_rows: set[int] = set()

    for row_idx, merges in row_merges.items():
        # Full-width merge: single merge covering >= 80% of data width
        has_full_width = any(m.col_span >= data_width * 0.8 for m in merges)

        if has_full_width:
            # Add all rows spanned by the full-width merge
            for m in merges:
                if m.col_span >= data_width * 0.8:
                    for r in range(m.origin_row, m.end_row + 1):
                        full_width_rows.add(r)
        else:
            # Check for hierarchical header pattern:
            # merge spans multiple columns AND the row below has
            # individual cells under those columns
            if _is_hierarchical_header(merges, grid, all_rows):
                for m in merges:
                    if m.col_span >= 2:
                        for r in range(m.origin_row, m.end_row + 1):
                            hierarchical_header_rows.add(r)

    if not full_width_rows and not hierarchical_header_rows:
        return None

    # Phase 2: Separate title rows from section separators
    # Title = full-width merge that appears before the first data/header row
    # Section separator = full-width merge that appears within data rows
    title_rows: set[int] = set()
    section_rows: set[int] = set()

    # Find the first "structural" row (header or non-full-width content)
    first_content_row = -1
    for row_idx in all_rows:
        if row_idx in full_width_rows:
            continue
        if row_idx in hierarchical_header_rows:
            first_content_row = row_idx
            break
        # Non-merged row with data — could be a simple header or data
        row_data = grid.get(row_idx, {})
        if len(row_data) >= 2:
            first_content_row = row_idx
            break

    for row_idx in full_width_rows:
        if first_content_row < 0 or row_idx < first_content_row:
            title_rows.add(row_idx)
        else:
            section_rows.add(row_idx)

    # Phase 3: Build header row list
    # Hierarchical header rows + the sub-header row immediately below them
    header_rows: set[int] = set(hierarchical_header_rows)

    if hierarchical_header_rows:
        last_merge_header = max(hierarchical_header_rows)
        sub_row = _next_row(all_rows, last_merge_header)
        if sub_row is not None and _is_text_dominant(grid.get(sub_row, {})):
            header_rows.add(sub_row)

    # Phase 4: Determine data start row
    all_structural = title_rows | header_rows | section_rows
    data_start = -1
    if all_structural:
        boundary = max(all_structural)
        for row_idx in all_rows:
            if row_idx > boundary and row_idx not in all_structural:
                data_start = row_idx
                break

    if data_start < 0 and header_rows:
        # All rows after last header
        data_start = max(header_rows) + 1

    return TableRegion(
        title_rows=frozenset(title_rows),
        header_rows=sorted(header_rows),
        data_start_row=data_start,
        section_rows=frozenset(section_rows),
    )


def _is_hierarchical_header(
    merges: list[MergeRegion],
    grid: dict[int, dict[int, str]],
    all_rows: list[int],
) -> bool:
    """Check if merged cells form a hierarchical header.

    True when a multi-column merge has individual sub-cells
    in the row immediately below it.
    """
    for m in merges:
        if m.col_span < 2:
            continue

        next_row = _next_row(all_rows, m.end_row)
        if next_row is None:
            continue

        next_row_data = grid.get(next_row, {})
        if not next_row_data:
            continue

        # Count individual cells under the merged region
        sub_cells = sum(
            1 for c in range(m.origin_col, m.end_col + 1)
            if c in next_row_data
        )
        if sub_cells >= 2:
            return True

    return False


def _next_row(all_rows: list[int], after: int) -> int | None:
    """Find the next row index strictly after the given row."""
    for r in all_rows:
        if r > after:
            return r
    return None


def _is_text_dominant(row_data: dict[int, str]) -> bool:
    """Check if a row is predominantly text (non-numeric)."""
    text_count = 0
    num_count = 0
    for val in row_data.values():
        if not val or not val.strip():
            continue
        try:
            float(val.strip())
            num_count += 1
        except ValueError:
            text_count += 1
    total = text_count + num_count
    return total >= 2 and text_count > num_count
