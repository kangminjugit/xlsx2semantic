"""Table layout hints for guiding semantic transformation.

Input formats:
    title_range  = "B2:Z2" | "B2:*2"  (* = end of sheet)
    header_range = "B4:Z6" | "B4:**"  (both col/row to end) | "B4:*6" (col to end)
    row_meta_col = "B" | "A:C"  (single column or column range)

'*' expands to the sheet's last column or last row.
-1 internally means "to end of sheet", resolved by resolve_defaults().
"""

from __future__ import annotations

import re
from dataclasses import dataclass


# B2:Z6 (full range)
_FULL_RANGE = re.compile(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$")
# B2:*6 (end col = *, end row specified)
_STAR_COL = re.compile(r"^([A-Z]+)(\d+):\*(\d+)$")
# B4:Z* (end col specified, end row = *)
_STAR_ROW = re.compile(r"^([A-Z]+)(\d+):([A-Z]+)\*$")
# B4:** (both end col and row = *)
_STAR_BOTH = re.compile(r"^([A-Z]+)(\d+):\*\*$")
# B (column letter only)
_COL_ONLY = re.compile(r"^([A-Z]+)$")
# A:C (column range, no row numbers)
_COL_RANGE = re.compile(r"^([A-Z]+):([A-Z]+)$")


def _col_to_num(letters: str) -> int:
    """Convert column letters to 1-based number. A=1, B=2, ..., Z=26, AA=27."""
    result = 0
    for ch in letters:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


@dataclass(frozen=True)
class TableLayoutHint:
    """Parsed table layout hint for semantic transformation."""

    title_rows: frozenset[int] = frozenset()
    header_start_row: int = -1
    header_end_row: int = -1
    header_start_col: int = -1
    header_end_col: int = -1
    row_meta_col_nums: tuple[int, ...] = ()
    header_row_count: int = 1

    @staticmethod
    def create(
        title_range: str | None = None,
        header_range: str | None = None,
        row_meta_col: str | None = None,
        header_rows: int = 1,
    ) -> TableLayoutHint | None:
        """Create a hint from API string parameters.

        Returns None if all parameters are empty (auto-detect mode).

        Args:
            title_range: e.g. "B2:Z2", "B2:*2"
            header_range: e.g. "B4:Z6", "B4:**", "B4:*6"
            row_meta_col: e.g. "B" or "A:C" (single column or range)
            header_rows: number of header rows (default 1)
        """
        has_title = bool(title_range and title_range.strip())
        has_header = bool(header_range and header_range.strip())
        has_meta = bool(row_meta_col and row_meta_col.strip())
        has_count = header_rows != 1

        if not has_title and not has_header and not has_meta and not has_count:
            return None

        title_rows: frozenset[int] = frozenset()
        h_start_row = h_end_row = h_start_col = h_end_col = -1
        meta_cols: tuple[int, ...] = ()

        if has_title:
            title_rows = _parse_title_range(title_range.strip().upper())

        if has_header:
            h_start_col, h_start_row, h_end_col, h_end_row = _parse_range(header_range.strip().upper())

        if has_meta:
            meta_str = row_meta_col.strip().upper()
            m = _COL_ONLY.match(meta_str)
            if m:
                meta_cols = (_col_to_num(m.group(1)),)
            else:
                m = _COL_RANGE.match(meta_str)
                if m:
                    start = _col_to_num(m.group(1))
                    end = _col_to_num(m.group(2))
                    meta_cols = tuple(range(start, end + 1))

        return TableLayoutHint(
            title_rows=title_rows,
            header_start_row=h_start_row,
            header_end_row=h_end_row,
            header_start_col=h_start_col,
            header_end_col=h_end_col,
            row_meta_col_nums=meta_cols,
            header_row_count=max(1, header_rows),
        )

    def resolve_defaults(self, max_row: int, max_col: int) -> TableLayoutHint:
        """Replace -1 (wildcard) values with actual sheet dimensions."""
        resolved_end_row = max_row if (self.header_end_row == -1 and self.header_start_row > 0) else self.header_end_row
        resolved_end_col = max_col if (self.header_end_col == -1 and self.header_start_col > 0) else self.header_end_col
        return TableLayoutHint(
            title_rows=self.title_rows,
            header_start_row=self.header_start_row,
            header_end_row=resolved_end_row,
            header_start_col=self.header_start_col,
            header_end_col=resolved_end_col,
            row_meta_col_nums=self.row_meta_col_nums,
            header_row_count=self.header_row_count,
        )

    @property
    def has_header_range(self) -> bool:
        return self.header_start_row > 0

    @property
    def has_row_meta_col(self) -> bool:
        return bool(self.row_meta_col_nums)

    @property
    def has_structural_hints(self) -> bool:
        """True if the hint has layout constraints beyond header_row_count."""
        return bool(self.title_rows) or self.header_start_row > 0 or bool(self.row_meta_col_nums)


def _parse_title_range(text: str) -> frozenset[int]:
    """Parse title range. Returns set of row numbers."""
    rows: set[int] = set()

    m = _FULL_RANGE.match(text)
    if m:
        r1, r2 = int(m.group(2)), int(m.group(4))
        rows.update(range(r1, r2 + 1))
        return frozenset(rows)

    m = _STAR_COL.match(text)
    if m:
        r1, r2 = int(m.group(2)), int(m.group(3))
        rows.update(range(r1, r2 + 1))
        return frozenset(rows)

    m = _STAR_ROW.match(text)
    if m:
        rows.add(int(m.group(2)))
        return frozenset(rows)

    m = _STAR_BOTH.match(text)
    if m:
        rows.add(int(m.group(2)))
        return frozenset(rows)

    return frozenset(rows)


def _parse_range(text: str) -> tuple[int, int, int, int]:
    """Parse range string. Returns (start_col, start_row, end_col, end_row).

    '*' positions become -1 (to end of sheet).
    """
    m = _FULL_RANGE.match(text)
    if m:
        return _col_to_num(m.group(1)), int(m.group(2)), _col_to_num(m.group(3)), int(m.group(4))

    m = _STAR_COL.match(text)
    if m:
        return _col_to_num(m.group(1)), int(m.group(2)), -1, int(m.group(3))

    m = _STAR_ROW.match(text)
    if m:
        return _col_to_num(m.group(1)), int(m.group(2)), _col_to_num(m.group(3)), -1

    m = _STAR_BOTH.match(text)
    if m:
        return _col_to_num(m.group(1)), int(m.group(2)), -1, -1

    return -1, -1, -1, -1
