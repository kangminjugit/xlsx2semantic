"""Tests for semantic transformer."""

from lxml import etree

from xlsx2semantic.semantic import transform, _to_tag_name


SHARED_STRINGS = [
    "Title Text",
    "State",
    "Total Number",
    "Total Percent",
    "Alabama",
    "Alaska",
    "745938",
    "100",
    "626078",
    "95",
]

# Minimal sheet XML with shared strings, merge cell for title
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


def test_auto_detect():
    """Without hint, should auto-detect headers."""
    result = transform(SHEET_XML, SHARED_STRINGS)
    assert "<semantic-table" in result
    assert "<record" in result


def test_with_hint_title_dedup():
    """Title from merged cells should not duplicate."""
    from xlsx2semantic.layout_hint import TableLayoutHint

    hint = TableLayoutHint.create(title_range="B2:*2", header_range="B4:D4", row_meta_col="B")
    result = transform(SHEET_XML, SHARED_STRINGS, hint)

    root = etree.fromstring(result.encode("utf-8"))
    title_el = root.find("title")
    assert title_el is not None
    # "Title Text" should appear only once despite merge across B2:D2
    assert title_el.text == "Title Text"


def test_with_hint_records():
    """Records should use header as tag names, meta as attribute."""
    from xlsx2semantic.layout_hint import TableLayoutHint

    hint = TableLayoutHint.create(title_range="B2:*2", header_range="B4:D4", row_meta_col="B")
    result = transform(SHEET_XML, SHARED_STRINGS, hint)

    root = etree.fromstring(result.encode("utf-8"))
    records = root.find("records")
    assert records is not None
    assert int(records.get("count")) == 2

    first_record = records.findall("record")[0]
    assert first_record.get("row") == "5"
    # rowMeta (col B = "State") → attribute name "state"
    assert first_record.get("state") == "Alabama"


def test_with_hint_schema():
    from xlsx2semantic.layout_hint import TableLayoutHint

    hint = TableLayoutHint.create(header_range="B4:D4", row_meta_col="B")
    result = transform(SHEET_XML, SHARED_STRINGS, hint)

    root = etree.fromstring(result.encode("utf-8"))
    schema = root.find("schema")
    assert schema is not None

    row_key = schema.find("row-key")
    assert row_key is not None
    assert row_key.get("attribute") == "state"

    columns = schema.findall("column")
    tags = [c.get("tag") for c in columns]
    assert "total_number" in tags
    assert "total_percent" in tags


def test_empty_sheet():
    empty = '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
    result = transform(empty, [])
    assert result == "<semantic-table/>"


def test_to_tag_name():
    assert _to_tag_name("Total Number") == "total_number"
    assert _to_tag_name("  Hello/World  ") == "hello_world"
    assert _to_tag_name("123abc") == "col_123abc"
    assert _to_tag_name("!!!") == "unknown"
    assert _to_tag_name("Race / Ethnicity") == "race_ethnicity"


def test_auto_detect_prefers_two_header_rows_when_data_starts_below():
    shared_strings = ["Region", "Sales", "Costs", "Name", "Total", "East", "West"]
    sheet_xml = """<?xml version="1.0" encoding="UTF-8"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
        <row r="1">
          <c r="A1" t="s"><v>0</v></c>
          <c r="B1" t="s"><v>1</v></c>
          <c r="C1" t="s"><v>2</v></c>
        </row>
        <row r="2">
          <c r="A2" t="s"><v>3</v></c>
          <c r="B2" t="s"><v>4</v></c>
          <c r="C2" t="s"><v>4</v></c>
        </row>
        <row r="3">
          <c r="A3" t="s"><v>5</v></c>
          <c r="B3"><v>100</v></c>
          <c r="C3"><v>40</v></c>
        </row>
        <row r="4">
          <c r="A4" t="s"><v>6</v></c>
          <c r="B4"><v>90</v></c>
          <c r="C4"><v>35</v></c>
        </row>
      </sheetData>
    </worksheet>
    """

    result = transform(sheet_xml, shared_strings)
    root = etree.fromstring(result.encode("utf-8"))
    schema = root.find("schema")
    assert schema is not None

    tags = [c.get("tag") for c in schema.findall("column")]
    assert "name" in tags
    assert "sales_total" in tags
    assert "costs_total" in tags

    records = root.find("records")
    assert records is not None
    first_record = records.findall("record")[0]
    assert first_record.get("row") == "3"


def test_auto_detect_keeps_single_repeated_value_row_as_title():
    shared_strings = ["Quarterly Sales", "Name", "Q1", "Q2", "East"]
    sheet_xml = """<?xml version="1.0" encoding="UTF-8"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
        <row r="1">
          <c r="A1" t="s"><v>0</v></c>
          <c r="B1" t="s"><v>0</v></c>
          <c r="C1" t="s"><v>0</v></c>
        </row>
        <row r="2">
          <c r="A2" t="s"><v>1</v></c>
          <c r="B2" t="s"><v>2</v></c>
          <c r="C2" t="s"><v>3</v></c>
        </row>
        <row r="3">
          <c r="A3" t="s"><v>4</v></c>
          <c r="B3"><v>100</v></c>
          <c r="C3"><v>120</v></c>
        </row>
      </sheetData>
    </worksheet>
    """

    result = transform(sheet_xml, shared_strings)
    root = etree.fromstring(result.encode("utf-8"))

    title_el = root.find("title")
    assert title_el is not None
    assert title_el.text == "Quarterly Sales"


def test_auto_detect_trace_contains_selected_rows_and_scores():
    result = transform(SHEET_XML, SHARED_STRINGS, include_trace=True)
    root = etree.fromstring(result.encode("utf-8"))

    trace = root.find("trace")
    assert trace is not None
    assert trace.get("mode") == "auto"

    selected = trace.find("selected")
    assert selected is not None
    assert selected.get("header-rows") == "4"
    assert selected.get("data-start-row") == "5"

    row_nodes = trace.findall("row")
    assert any(row.get("index") == "4" and row.get("role") == "header" for row in row_nodes)
    assert any(row.get("index") == "5" and row.get("role") == "data" for row in row_nodes)


def test_hint_trace_marks_explicit_layout_source():
    from xlsx2semantic.layout_hint import TableLayoutHint

    hint = TableLayoutHint.create(title_range="B2:*2", header_range="B4:D4", row_meta_col="B")
    result = transform(SHEET_XML, SHARED_STRINGS, hint, include_trace=True)
    root = etree.fromstring(result.encode("utf-8"))

    trace = root.find("trace")
    assert trace is not None
    assert trace.get("mode") == "hint"

    selected = trace.find("selected")
    assert selected is not None
    assert selected.get("header-rows") == "4"
    assert selected.get("title-rows") == "2"
    assert selected.get("row-meta-cols") == "2"
