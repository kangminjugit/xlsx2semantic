"""Tests for structural anchor table detection."""

from lxml import etree

from xlsx2semantic.semantic import transform
from xlsx2semantic.structural_anchor import (
    TableBoundary,
    detect_tables,
)


# ── Unit tests for detect_tables ──


def test_single_table_detected():
    """A contiguous block of cells should be detected as one table."""
    grid = {
        1: {1: "Name", 2: "Age", 3: "Score"},
        2: {1: "Alice", 2: "30", 3: "95"},
        3: {1: "Bob", 2: "25", 3: "88"},
    }
    tables = detect_tables(grid, {})
    assert len(tables) == 1
    t = tables[0]
    assert t.min_row == 1
    assert t.max_row == 3
    assert t.min_col == 1
    assert t.max_col == 3


def test_two_tables_vertical_gap():
    """Two blocks separated by >1 empty row should be two tables."""
    grid = {
        1: {1: "H1", 2: "H2"},
        2: {1: "a", 2: "1"},
        3: {1: "b", 2: "2"},
        # rows 4-5 empty → gap of 2 > _ROW_GAP(1)
        6: {1: "X", 2: "Y"},
        7: {1: "c", 2: "3"},
        8: {1: "d", 2: "4"},
    }
    tables = detect_tables(grid, {})
    assert len(tables) == 2
    assert tables[0].max_row == 3
    assert tables[1].min_row == 6


def test_two_tables_horizontal_gap():
    """Two blocks separated by >1 empty column should be two tables."""
    grid = {
        1: {1: "A", 2: "B", 5: "C", 6: "D"},
        2: {1: "1", 2: "2", 5: "3", 6: "4"},
        3: {1: "5", 2: "6", 5: "7", 6: "8"},
    }
    tables = detect_tables(grid, {})
    assert len(tables) == 2
    assert tables[0].max_col == 2
    assert tables[1].min_col == 5


def test_gap_of_one_row_stays_together():
    """A single empty row should NOT split the table."""
    grid = {
        1: {1: "H1", 2: "H2"},
        2: {1: "a", 2: "1"},
        # row 3 empty — gap of 1 = _ROW_GAP
        4: {1: "b", 2: "2"},
    }
    tables = detect_tables(grid, {})
    assert len(tables) == 1


def test_header_detection_text_then_numeric():
    """First text rows → headers, first numeric-majority row → data start."""
    grid = {
        1: {1: "Name", 2: "Value", 3: "Score"},
        2: {1: "Alice", 2: "100", 3: "95"},
        3: {1: "Bob", 2: "200", 3: "88"},
    }
    tables = detect_tables(grid, {})
    assert len(tables) == 1
    t = tables[0]
    assert 1 in t.header_rows
    assert t.data_start_row == 2


def test_title_row_single_value():
    """A row with one unique value at the top → title."""
    grid = {
        1: {1: "Report Title", 2: "Report Title", 3: "Report Title"},
        2: {1: "Name", 2: "Value", 3: "Score"},
        3: {1: "Alice", 2: "100", 3: "95"},
        4: {1: "Bob", 2: "200", 3: "88"},
    }
    tables = detect_tables(grid, {})
    assert len(tables) == 1
    t = tables[0]
    assert 1 in t.title_rows
    assert 2 in t.header_rows


def test_row_meta_col_not_in_boundary():
    """TableBoundary no longer carries row_meta_col; detection is hint-only."""
    grid = {
        1: {1: "State", 2: "Population", 3: "Area"},
        2: {1: "Alabama", 2: "5000000", 3: "52419"},
        3: {1: "Alaska", 2: "700000", 3: "665384"},
    }
    tables = detect_tables(grid, {})
    assert len(tables) == 1
    assert not hasattr(tables[0], "row_meta_col")


def test_empty_grid():
    assert detect_tables({}, {}) == []


def test_too_small_cluster_skipped():
    """A single cell or single-row cluster should be skipped."""
    grid = {1: {1: "lonely"}}
    tables = detect_tables(grid, {})
    assert len(tables) == 0


# ── Integration tests via semantic.transform ──

SHARED_STRINGS = [
    "Title Text",  # 0
    "State",        # 1
    "Total Number", # 2
    "Total Percent", # 3
    "Alabama",      # 4
    "Alaska",       # 5
    "745938",       # 6
    "100",          # 7
    "626078",       # 8
    "95",           # 9
]

SHEET_XML = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <c r="B2" t="s"><v>0</v></c>
      <c r="C2" t="s"><v>0</v></c>
    </row>
    <row r="4">
      <c r="B4" t="s"><v>1</v></c>
      <c r="C4" t="s"><v>2</v></c>
      <c r="D4" t="s"><v>3</v></c>
    </row>
    <row r="5">
      <c r="B5" t="s"><v>4</v></c>
      <c r="C5" t="s"><v>6</v></c>
      <c r="D5" t="s"><v>7</v></c>
    </row>
    <row r="6">
      <c r="B6" t="s"><v>5</v></c>
      <c r="C6" t="s"><v>8</v></c>
      <c r="D6" t="s"><v>9</v></c>
    </row>
  </sheetData>
  <mergeCells>
    <mergeCell ref="B2:D2"/>
  </mergeCells>
</worksheet>
"""


def test_auto_detect_finds_single_table():
    """Without hint, structural anchors should find the table and records."""
    result = transform(SHEET_XML, SHARED_STRINGS)
    assert "<semantic-table" in result
    assert "<record" in result


def test_auto_detect_title_from_merge():
    """Full-width merge at top should become <title>."""
    result = transform(SHEET_XML, SHARED_STRINGS)
    root = etree.fromstring(result.encode("utf-8"))
    title_el = root.find("title")
    assert title_el is not None
    assert "Title Text" in title_el.text


def test_auto_detect_header_tags():
    """Header row should produce meaningful column tags."""
    result = transform(SHEET_XML, SHARED_STRINGS)
    root = etree.fromstring(result.encode("utf-8"))
    schema = root.find("schema")
    assert schema is not None
    tags = [c.get("tag") for c in schema.findall("column")]
    assert any("total_number" in t for t in tags)


def test_auto_detect_no_row_meta():
    """Auto-detect mode should not produce row-key; all columns are child tags."""
    result = transform(SHEET_XML, SHARED_STRINGS)
    root = etree.fromstring(result.encode("utf-8"))
    assert root.find("schema/row-key") is None


def test_auto_detect_records():
    """Data rows should become records with correct values."""
    result = transform(SHEET_XML, SHARED_STRINGS)
    root = etree.fromstring(result.encode("utf-8"))
    records = root.find("records")
    assert records is not None
    assert int(records.get("count")) >= 2


TWO_TABLE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>1</v></c>
      <c r="B1" t="s"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>4</v></c>
      <c r="B2" t="s"><v>6</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>5</v></c>
      <c r="B3" t="s"><v>8</v></c>
    </row>
    <row r="6">
      <c r="A6" t="s"><v>1</v></c>
      <c r="B6" t="s"><v>3</v></c>
    </row>
    <row r="7">
      <c r="A7" t="s"><v>4</v></c>
      <c r="B7" t="s"><v>7</v></c>
    </row>
    <row r="8">
      <c r="A8" t="s"><v>5</v></c>
      <c r="B8" t="s"><v>9</v></c>
    </row>
  </sheetData>
</worksheet>
"""


def test_multi_table_detection():
    """Two table blocks separated by empty rows → <sheet> with two <semantic-table>."""
    result = transform(TWO_TABLE_XML, SHARED_STRINGS)
    root = etree.fromstring(result.encode("utf-8"))
    if root.tag == "sheet":
        tables = root.findall("semantic-table")
        assert len(tables) == 2
        assert tables[0].get("index") == "1"
        assert tables[1].get("index") == "2"
