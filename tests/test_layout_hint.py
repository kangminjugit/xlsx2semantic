"""Tests for TableLayoutHint parsing."""

from xlsx2semantic.layout_hint import TableLayoutHint


def test_create_returns_none_when_all_empty():
    assert TableLayoutHint.create() is None
    assert TableLayoutHint.create("", "", "") is None
    assert TableLayoutHint.create(None, None, None) is None


def test_full_range_title():
    hint = TableLayoutHint.create(title_range="B2:Z3")
    assert hint is not None
    assert hint.title_rows == frozenset({2, 3})


def test_star_col_title():
    hint = TableLayoutHint.create(title_range="B2:*2")
    assert hint is not None
    assert hint.title_rows == frozenset({2})


def test_star_both_title():
    hint = TableLayoutHint.create(title_range="B2:**")
    assert hint is not None
    assert hint.title_rows == frozenset({2})


def test_full_header_range():
    hint = TableLayoutHint.create(header_range="B4:Z6")
    assert hint is not None
    assert hint.header_start_row == 4
    assert hint.header_end_row == 6
    assert hint.header_start_col == 2
    assert hint.header_end_col == 26
    assert hint.has_header_range


def test_star_col_header():
    hint = TableLayoutHint.create(header_range="B4:*6")
    assert hint is not None
    assert hint.header_start_col == 2
    assert hint.header_end_col == -1  # wildcard
    assert hint.header_end_row == 6


def test_star_both_header():
    hint = TableLayoutHint.create(header_range="B4:**")
    assert hint is not None
    assert hint.header_start_col == 2
    assert hint.header_end_col == -1
    assert hint.header_end_row == -1


def test_row_meta_col():
    hint = TableLayoutHint.create(row_meta_col="B")
    assert hint is not None
    assert hint.row_meta_col_num == 2
    assert hint.has_row_meta_col


def test_resolve_defaults():
    hint = TableLayoutHint.create(header_range="B4:**")
    assert hint is not None
    resolved = hint.resolve_defaults(max_row=60, max_col=30)
    assert resolved.header_end_row == 60
    assert resolved.header_end_col == 30


def test_resolve_defaults_no_wildcard():
    hint = TableLayoutHint.create(header_range="B4:Z6")
    assert hint is not None
    resolved = hint.resolve_defaults(max_row=60, max_col=30)
    assert resolved.header_end_row == 6
    assert resolved.header_end_col == 26


def test_combined_params():
    hint = TableLayoutHint.create(
        title_range="B2:*2",
        header_range="B4:*6",
        row_meta_col="B",
    )
    assert hint is not None
    assert hint.title_rows == frozenset({2})
    assert hint.header_start_row == 4
    assert hint.header_end_row == 6
    assert hint.header_end_col == -1
    assert hint.row_meta_col_num == 2
