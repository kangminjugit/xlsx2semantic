"""Score-based semantic role detection for table regions.

This module keeps table boundary detection deterministic while making
header/title inference evidence-driven instead of rule-only.
"""

from __future__ import annotations

from dataclasses import dataclass

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
