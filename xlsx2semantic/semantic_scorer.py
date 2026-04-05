"""Score-based semantic role detection for table regions.

This module keeps table boundary detection deterministic while making
header/title inference evidence-driven instead of rule-only.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from statistics import pstdev

from xlsx2semantic.structural_anchor import TableBoundary


@dataclass(frozen=True)
class Evidence:
    """A small, explainable piece of evidence contributing to a decision."""

    name: str
    score: float


@dataclass(frozen=True)
class SemanticDecision:
    """Final semantic decision for a detected table region."""

    title_rows: tuple[int, ...]
    header_rows: tuple[int, ...]
    data_start_row: int
    confidence: float
    evidences: tuple[Evidence, ...] = ()
    row_traces: tuple["RowTrace", ...] = ()


@dataclass(frozen=True)
class RowTrace:
    """Per-row scores for the final decision."""

    row_idx: int
    role: str
    title_score: float
    header_score: float
    data_score: float


@dataclass(frozen=True)
class CellFeatures:
    """Feature set for a single cell inside a table boundary."""

    is_text: bool
    is_numeric: bool
    is_empty: bool
    text_length: int
    contains_unit_keyword: bool
    contains_parentheses: bool
    contains_colon: bool
    is_symbolic: bool
    is_date_like: bool
    is_near_top: float
    is_near_left: float
    is_merged: bool
    merge_span_cols: int
    merge_span_rows: int


@dataclass(frozen=True)
class RowFeatures:
    """Feature set summarizing one row in a table boundary."""

    row_coverage_ratio: float
    row_text_ratio: float
    row_numeric_ratio: float
    row_unique_ratio: float
    row_schema_similarity: float
    row_pattern_repeat_score: float
    row_has_single_value: bool
    row_value_variance: float
    row_symbolic_ratio: float


@dataclass(frozen=True)
class ColumnFeatures:
    """Feature set summarizing one column in a table boundary."""

    col_text_ratio: float
    col_numeric_ratio: float
    col_unique_ratio: float
    col_type_consistency: float
    col_category_repetition: float
    col_is_key_candidate: bool
    col_leftness_score: float


def rescore_table_boundary(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    preferred_header_rows: int = 1,
) -> TableBoundary:
    """Return a new boundary whose title/header rows were score-selected."""
    decision = decide_semantics(grid, table, preferred_header_rows)
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


def decide_semantics(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    preferred_header_rows: int = 1,
) -> SemanticDecision:
    """Choose title/header/data rows with a score-based decision."""
    content_rows = [
        row_idx
        for row_idx in range(table.min_row, table.max_row + 1)
        if grid.get(row_idx) and row_idx not in table.section_rows
    ]
    if not content_rows:
        return SemanticDecision((), table.header_rows, table.data_start_row, 0.0)

    explicit_titles = {row for row in table.title_rows if row in content_rows}
    max_header_rows = min(
        3,
        max(preferred_header_rows + 1, len(table.header_rows) + 1, 2),
        max(1, len(content_rows) - 1),
    )

    best: SemanticDecision | None = None
    best_score = float("-inf")

    max_start_idx = max(1, min(4, len(content_rows)))
    for start_idx in range(max_start_idx):
        if content_rows[start_idx] in explicit_titles:
            continue

        prefix_rows = content_rows[:start_idx]
        inferred_titles = tuple(
            row_idx
            for row_idx in prefix_rows
            if row_idx in explicit_titles or _title_score(grid, table, row_idx, content_rows) >= 0.6
        )

        for header_len in range(1, max_header_rows + 1):
            header_rows = tuple(content_rows[start_idx:start_idx + header_len])
            if len(header_rows) != header_len:
                continue

            data_rows = [
                row_idx
                for row_idx in content_rows
                if row_idx > header_rows[-1] and row_idx not in table.section_rows
            ]
            if not data_rows:
                continue

            header_score = _header_block_score(grid, table, header_rows, data_rows, content_rows)
            title_bonus = sum(
                _title_score(grid, table, row_idx, content_rows)
                for row_idx in inferred_titles
            )
            data_bonus = _data_quality_score(grid, table, data_rows)
            explicit_bonus = 0.2 if explicit_titles.issubset(set(inferred_titles)) else 0.0
            total = header_score + title_bonus + data_bonus + explicit_bonus

            if total > best_score:
                best_score = total
                confidence = max(0.0, min(1.0, total / 2.5))
                best = SemanticDecision(
                    title_rows=inferred_titles,
                    header_rows=header_rows,
                    data_start_row=data_rows[0],
                    confidence=confidence,
                    evidences=(
                        Evidence("header_block", round(header_score, 3)),
                        Evidence("title_prefix", round(title_bonus, 3)),
                        Evidence("data_quality", round(data_bonus, 3)),
                    ),
                )

    if best is not None:
        return _attach_row_traces(grid, table, best)

    fallback_header = table.header_rows or (content_rows[0],)
    data_start = next((row for row in content_rows if row > fallback_header[-1]), table.max_row + 1)
    return SemanticDecision(
        title_rows=tuple(sorted(explicit_titles)),
        header_rows=fallback_header,
        data_start_row=data_start,
        confidence=0.2,
    )


def extract_cell_features(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    row_idx: int,
    col_idx: int,
    merge_map: dict[str, str] | None = None,
) -> CellFeatures:
    """Extract cell-level features used by semantic heuristics."""
    raw_value = grid.get(row_idx, {}).get(col_idx, "")
    value = (raw_value or "").strip()
    merge_info = _resolve_merge_info(row_idx, col_idx, merge_map or {})
    return CellFeatures(
        is_text=bool(value) and not _is_numeric(value),
        is_numeric=_is_numeric(value),
        is_empty=not bool(value),
        text_length=len(value),
        contains_unit_keyword=_contains_unit_keyword(value),
        contains_parentheses=("(" in value) or (")" in value),
        contains_colon=":" in value,
        is_symbolic=_is_symbolic(value),
        is_date_like=_is_date_like(value),
        is_near_top=_topness(row_idx, _content_rows(grid, table)),
        is_near_left=_leftness(col_idx, table),
        is_merged=merge_info[0],
        merge_span_cols=merge_info[1],
        merge_span_rows=merge_info[2],
    )


def extract_row_features(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    row_idx: int,
) -> RowFeatures:
    """Extract row-level structural and semantic features."""
    values = _row_values(grid, table, row_idx)
    content_rows = _content_rows(grid, table)
    row_patterns = {r: _row_schema_pattern(grid, table, r) for r in content_rows}
    current_pattern = row_patterns.get(row_idx, ())
    comparable = [p for r, p in row_patterns.items() if r != row_idx]
    schema_similarity = _pattern_similarity(current_pattern, comparable)
    repeat_score = 0.0
    if comparable:
        repeat_score = sum(1 for p in comparable if p == current_pattern) / len(comparable)

    non_empty = [v for v in values if v.strip()]
    normalized = [v.strip().lower() for v in non_empty]
    text_lengths = [len(v) for v in non_empty]
    value_variance = 0.0 if len(text_lengths) <= 1 else min(1.0, pstdev(text_lengths) / 10.0)

    return RowFeatures(
        row_coverage_ratio=_coverage_ratio(grid, table, row_idx),
        row_text_ratio=_text_ratio(values),
        row_numeric_ratio=_numeric_ratio(values),
        row_unique_ratio=len(set(normalized)) / max(1, len(normalized)),
        row_schema_similarity=schema_similarity,
        row_pattern_repeat_score=repeat_score,
        row_has_single_value=len(set(normalized)) == 1 if normalized else False,
        row_value_variance=value_variance,
        row_symbolic_ratio=sum(1 for v in values if _is_symbolic(v)) / max(1, len(values)),
    )


def extract_column_features(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    col_idx: int,
) -> ColumnFeatures:
    """Extract column-level features for key-column and type inference."""
    values = _column_values(grid, table, col_idx)
    non_empty = [v for v in values if v.strip()]
    normalized = [v.strip().lower() for v in non_empty]
    text_ratio = _text_ratio(non_empty)
    numeric_ratio = _numeric_ratio(non_empty)
    symbolic_ratio = sum(1 for v in non_empty if _is_symbolic(v)) / max(1, len(non_empty))
    dominant_ratio = max(text_ratio, numeric_ratio, symbolic_ratio) if non_empty else 0.0
    unique_ratio = len(set(normalized)) / max(1, len(normalized))
    leftness = _leftness(col_idx, table)
    is_key = bool(non_empty) and unique_ratio >= 0.7 and text_ratio >= 0.5 and leftness >= 0.5

    return ColumnFeatures(
        col_text_ratio=text_ratio,
        col_numeric_ratio=numeric_ratio,
        col_unique_ratio=unique_ratio,
        col_type_consistency=dominant_ratio,
        col_category_repetition=1.0 - unique_ratio,
        col_is_key_candidate=is_key,
        col_leftness_score=leftness,
    )


def _title_score(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    row_idx: int,
    content_rows: list[int],
) -> float:
    values = _row_values(grid, table, row_idx)
    if not values:
        return 0.0

    unique_values = {value.strip() for value in values if value.strip()}
    single_value = 1.0 if len(unique_values) == 1 else 0.0
    text_ratio = _text_ratio(values)
    numeric_ratio = _numeric_ratio(values)
    topness = _topness(row_idx, content_rows)
    coverage = len(values) / max(1, table.max_col - table.min_col + 1)

    return (
        0.45 * single_value
        + 0.25 * topness
        + 0.20 * text_ratio
        + 0.10 * min(1.0, coverage)
        - 0.35 * numeric_ratio
    )


def _header_block_score(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    header_rows: tuple[int, ...],
    data_rows: list[int],
    content_rows: list[int],
) -> float:
    row_scores = [
        _header_row_score(grid, table, row_idx, data_rows, content_rows)
        for row_idx in header_rows
    ]
    header_density = sum(_coverage_ratio(grid, table, row_idx) for row_idx in header_rows) / len(header_rows)
    hierarchy_bonus = _hierarchy_bonus(grid, table, header_rows)
    return (sum(row_scores) / len(row_scores)) + 0.15 * header_density + hierarchy_bonus


def _header_row_score(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    row_idx: int,
    data_rows: list[int],
    content_rows: list[int],
) -> float:
    values = _row_values(grid, table, row_idx)
    if not values:
        return -1.0

    text_ratio = _text_ratio(values)
    numeric_ratio = _numeric_ratio(values)
    topness = _topness(row_idx, content_rows)
    unique_ratio = len({value.strip().lower() for value in values if value.strip()}) / max(1, len(values))

    data_sample = data_rows[: min(3, len(data_rows))]
    below_numeric_ratio = 0.0
    if data_sample:
        below_numeric_ratio = sum(
            _numeric_ratio(_row_values(grid, table, sample_idx))
            for sample_idx in data_sample
        ) / len(data_sample)

    type_shift = max(0.0, below_numeric_ratio - numeric_ratio)

    return (
        0.25 * topness
        + 0.30 * text_ratio
        + 0.20 * unique_ratio
        + 0.25 * type_shift
        - 0.35 * numeric_ratio
    )


def _data_quality_score(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    data_rows: list[int],
) -> float:
    sample = data_rows[: min(4, len(data_rows))]
    if not sample:
        return -1.0

    numeric_ratios = [
        _numeric_ratio(_row_values(grid, table, row_idx))
        for row_idx in sample
    ]
    coverage = [
        _coverage_ratio(grid, table, row_idx)
        for row_idx in sample
    ]
    return 0.35 * (sum(numeric_ratios) / len(numeric_ratios)) + 0.25 * (sum(coverage) / len(coverage))


def _data_row_score(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    row_idx: int,
) -> float:
    values = _row_values(grid, table, row_idx)
    if not values:
        return -1.0

    numeric_ratio = _numeric_ratio(values)
    coverage = _coverage_ratio(grid, table, row_idx)
    mixed_bonus = 0.15 if 0.0 < numeric_ratio < 1.0 else 0.0
    return 0.45 * numeric_ratio + 0.30 * coverage + mixed_bonus


def _attach_row_traces(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    decision: SemanticDecision,
) -> SemanticDecision:
    content_rows = [
        row_idx
        for row_idx in range(table.min_row, table.max_row + 1)
        if grid.get(row_idx) and row_idx not in table.section_rows
    ]
    data_rows = [row_idx for row_idx in content_rows if row_idx >= decision.data_start_row]
    title_set = set(decision.title_rows)
    header_set = set(decision.header_rows)

    row_traces = []
    for row_idx in content_rows:
        if row_idx in title_set:
            role = "title"
        elif row_idx in header_set:
            role = "header"
        elif row_idx >= decision.data_start_row:
            role = "data"
        else:
            role = "context"

        row_traces.append(RowTrace(
            row_idx=row_idx,
            role=role,
            title_score=round(_title_score(grid, table, row_idx, content_rows), 3),
            header_score=round(_header_row_score(grid, table, row_idx, data_rows, content_rows), 3),
            data_score=round(_data_row_score(grid, table, row_idx), 3),
        ))

    return SemanticDecision(
        title_rows=decision.title_rows,
        header_rows=decision.header_rows,
        data_start_row=decision.data_start_row,
        confidence=decision.confidence,
        evidences=decision.evidences,
        row_traces=tuple(row_traces),
    )


def _hierarchy_bonus(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    header_rows: tuple[int, ...],
) -> float:
    if len(header_rows) <= 1:
        return 0.0

    bonus = 0.0
    for upper_row, lower_row in zip(header_rows, header_rows[1:]):
        upper_values = _row_values(grid, table, upper_row)
        lower_values = _row_values(grid, table, lower_row)
        if not upper_values or not lower_values:
            continue

        upper_norm = [value.strip().lower() for value in upper_values if value.strip()]
        lower_norm = [value.strip().lower() for value in lower_values if value.strip()]
        upper_unique = len(set(upper_norm))
        lower_unique = len(set(lower_norm))

        if upper_unique < len(upper_norm) and lower_unique >= upper_unique:
            bonus += 0.2
        if _text_ratio(upper_values) > 0.9 and _text_ratio(lower_values) > 0.9:
            bonus += 0.08

    return bonus


def _row_values(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    row_idx: int,
) -> list[str]:
    row = grid.get(row_idx, {})
    return [
        value
        for col, value in sorted(row.items())
        if table.min_col <= col <= table.max_col and value and value.strip()
    ]


def _column_values(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    col_idx: int,
) -> list[str]:
    values: list[str] = []
    for row_idx in range(table.min_row, table.max_row + 1):
        value = grid.get(row_idx, {}).get(col_idx, "")
        if value and value.strip():
            values.append(value)
    return values


def _content_rows(grid: dict[int, dict[int, str]], table: TableBoundary) -> list[int]:
    return [
        row_idx
        for row_idx in range(table.min_row, table.max_row + 1)
        if grid.get(row_idx) and row_idx not in table.section_rows
    ]


def _coverage_ratio(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    row_idx: int,
) -> float:
    return len(_row_values(grid, table, row_idx)) / max(1, table.max_col - table.min_col + 1)


def _topness(row_idx: int, content_rows: list[int]) -> float:
    if len(content_rows) <= 1:
        return 1.0
    position = content_rows.index(row_idx)
    return 1.0 - (position / (len(content_rows) - 1))


def _leftness(col_idx: int, table: TableBoundary) -> float:
    width = max(1, table.max_col - table.min_col)
    return 1.0 - ((col_idx - table.min_col) / width)


def _text_ratio(values: list[str]) -> float:
    return sum(0 if _is_numeric(value) else 1 for value in values) / max(1, len(values))


def _numeric_ratio(values: list[str]) -> float:
    return sum(1 if _is_numeric(value) else 0 for value in values) / max(1, len(values))


def _is_numeric(text: str | None) -> bool:
    if not text or not text.strip():
        return False
    try:
        float(text.strip().replace(",", ""))
        return True
    except ValueError:
        return False


def _is_symbolic(text: str | None) -> bool:
    if not text:
        return False
    return text.strip().lower() in {"-", "n/a", "na", "none", "null", "nan", "tbd"}


def _is_date_like(text: str | None) -> bool:
    if not text:
        return False
    normalized = text.strip()
    patterns = (
        r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$",
        r"^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$",
        r"^\d{4}년\s*\d{1,2}월(\s*\d{1,2}일)?$",
    )
    return any(re.match(pattern, normalized) for pattern in patterns)


def _contains_unit_keyword(text: str) -> bool:
    if not text:
        return False
    keywords = ("단위", "억원", "백만원", "천원", "kg", "km", "m²", "%")
    lowered = text.lower()
    return any(keyword.lower() in lowered for keyword in keywords)


def _row_schema_pattern(
    grid: dict[int, dict[int, str]],
    table: TableBoundary,
    row_idx: int,
) -> tuple[str, ...]:
    row = grid.get(row_idx, {})
    pattern: list[str] = []
    for col_idx in range(table.min_col, table.max_col + 1):
        value = (row.get(col_idx, "") or "").strip()
        if not value:
            pattern.append("E")
        elif _is_symbolic(value):
            pattern.append("S")
        elif _is_numeric(value):
            pattern.append("N")
        else:
            pattern.append("T")
    return tuple(pattern)


def _pattern_similarity(current: tuple[str, ...], candidates: list[tuple[str, ...]]) -> float:
    if not current or not candidates:
        return 0.0
    scores: list[float] = []
    for other in candidates:
        if len(other) != len(current):
            continue
        matched = sum(1 for c1, c2 in zip(current, other) if c1 == c2)
        scores.append(matched / len(current))
    return sum(scores) / len(scores) if scores else 0.0


def _resolve_merge_info(row_idx: int, col_idx: int, merge_map: dict[str, str]) -> tuple[bool, int, int]:
    """Return (is_merged, span_cols, span_rows) for one cell."""
    if not merge_map:
        return False, 1, 1

    coord = f"{col_idx}:{row_idx}"
    origin = merge_map.get(coord)
    if not origin:
        return False, 1, 1

    merged_cells = [key for key, mapped_origin in merge_map.items() if mapped_origin == origin]
    if len(merged_cells) <= 1:
        return False, 1, 1

    merged_cols = [int(cell.split(":")[0]) for cell in merged_cells]
    merged_rows = [int(cell.split(":")[1]) for cell in merged_cells]
    span_cols = (max(merged_cols) - min(merged_cols) + 1) if merged_cols else 1
    span_rows = (max(merged_rows) - min(merged_rows) + 1) if merged_rows else 1
    return True, span_cols, span_rows
