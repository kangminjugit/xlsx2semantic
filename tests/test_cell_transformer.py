"""Tests for cell transformer."""

from xlsx2semantic.cell_transformer import parse_shared_strings, transform_sheet, _col_to_num


def test_parse_shared_strings():
    xml = """<?xml version="1.0" encoding="UTF-8"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
      <si><t>Hello</t></si>
      <si><t>World</t></si>
      <si><r><t>Rich</t></r><r><t> Text</t></r></si>
    </sst>"""
    strings = parse_shared_strings(xml)
    assert strings == ["Hello", "World", "Rich Text"]


def test_parse_shared_strings_empty():
    assert parse_shared_strings(None) == []
    assert parse_shared_strings("") == []


def test_col_to_num():
    assert _col_to_num("A") == 1
    assert _col_to_num("B") == 2
    assert _col_to_num("Z") == 26
    assert _col_to_num("AA") == 27
    assert _col_to_num("AZ") == 52


def test_transform_sheet():
    sheet_xml = """<?xml version="1.0" encoding="UTF-8"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
        <row r="1">
          <c r="A1" t="s"><v>0</v></c>
          <c r="B1"><v>42</v></c>
        </row>
      </sheetData>
    </worksheet>"""
    shared = ["Hello"]
    result = transform_sheet(sheet_xml, shared)

    assert 'ref="A1"' in result
    assert 'type="sharedString"' in result
    assert 'value="Hello"' in result
    assert 'ref="B1"' in result
    assert 'type="number"' in result
    assert 'value="42"' in result
