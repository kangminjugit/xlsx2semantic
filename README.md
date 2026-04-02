# xlsx2semantic

**Transform XLSX into LLM-friendly semantic XML.**

Raw Excel XML is unreadable — cell references like `<c r="B7" t="s"><v>74</v></c>` mean nothing to an LLM.
**xlsx2semantic** converts that into structured, self-describing XML that any language model can instantly understand.

```
Before (raw OOXML)                          After (semantic XML)
─────────────────────                       ────────────────────
<c r="B7" s="59" t="s">                    <record row="7" state="Alabama">
  <v>74</v>                                   <total_number>745938</total_number>
</c>                                          <total_percent>100</total_percent>
<c r="C7" s="60">                           </record>
  <v>745938</v>
</c>
```

## Why?

| Problem | xlsx2semantic |
|---------|---------------|
| LLMs can't parse raw OOXML cell references | Headers become tag names, data becomes readable records |
| Shared strings are just index numbers | Automatically resolved to actual text |
| Merged cells break structure | Propagated correctly across the grid |
| Multi-level headers are ambiguous | Combined into hierarchical tag names |
| Style indices are opaque | Resolved to human-readable font, color, format info |
| One sheet with multiple tables | Structural anchors auto-detect each table region |

## Quick Start

### Install

```bash
pip install xlsx2semantic
```

### Python

```python
from xlsx2semantic import parse_file

# No hints needed — tables are auto-detected
result = parse_file("enrollment.xlsx")

for sheet, xml in result.semantic_xml.items():
    print(xml)
```

Output:

```xml
<semantic-table>
  <title>Public school enrollment by race/ethnicity</title>
  <schema>
    <row-key index="2" attribute="state"/>
    <column index="3" tag="total_number"/>
    <column index="4" tag="total_percent"/>
  </schema>
  <records count="52">
    <record row="8" state="Alabama">
      <total_number>745938</total_number>
      <total_percent>100</total_percent>
    </record>
    <record row="9" state="Alaska">
      <total_number>132731</total_number>
      <total_percent>100</total_percent>
    </record>
  </records>
</semantic-table>
```

### CLI

```bash
# Semantic XML (default) — auto-detects table structure
xlsx2semantic data.xlsx

# With explicit layout hints (optional override)
xlsx2semantic data.xlsx --title-range "B2:*2" --header-range "B4:*6" --row-meta-col B

# Save to file
xlsx2semantic data.xlsx -o output.xml

# Different output modes
xlsx2semantic data.xlsx --mode cell      # enriched <cell> tags
xlsx2semantic data.xlsx --mode raw-xml   # original OOXML
xlsx2semantic data.xlsx --mode all       # everything
```

Install CLI extras: `pip install xlsx2semantic[cli]`

## Automatic Table Detection (Structural Anchors)

xlsx2semantic uses a [SpreadsheetLLM](https://arxiv.org/abs/2407.09025)-inspired approach to automatically detect table regions without manual hints:

1. **Cell clustering** — Non-empty cells are grouped into connected regions by row/column proximity. Clusters separated by empty rows or columns become separate tables.
2. **Title detection** — Full-width merged cells at the top of a region are recognized as titles.
3. **Header detection** — Row heterogeneity (text/numeric ratio) identifies header rows vs data rows. Text-dominant rows at the top of a cluster become headers; the first numeric-majority row starts the data region.
4. **Multi-level header merging** — Hierarchical headers across multiple rows are combined into structured tag names (e.g. `학생_수용_현황_1_신청_학생_수_1학년`).
5. **Row-key auto-detection** — The leftmost column that is predominantly text (>80%) while other columns are numeric is automatically identified as the row key.

### Single Table

```python
result = parse_file("report.xlsx")  # no hints needed
```

### Multiple Tables in One Sheet

When a sheet contains multiple tables separated by empty rows or columns, each is detected independently:

```xml
<sheet>
  <semantic-table index="1">
    <schema>...</schema>
    <records count="10">...</records>
  </semantic-table>
  <semantic-table index="2">
    <schema>...</schema>
    <records count="5">...</records>
  </semantic-table>
</sheet>
```

### Layout Hints (Optional Override)

For edge cases where auto-detection needs guidance, explicit layout hints still work:

```python
result = parse_file(
    "report.xlsx",
    title_range="B2:*2",       # title rows (B2, cols to end of sheet)
    header_range="B4:*6",      # header area (rows 4-6, cols to end)
    row_meta_col="B",          # row label column → becomes record attribute
)
```

| Pattern | Meaning |
|---------|---------|
| `B4:Z6` | Exact range: col B–Z, row 4–6 |
| `B4:*6` | Col B to **end of sheet**, row 4–6 |
| `B4:Z*` | Col B–Z, row 4 to **end of sheet** |
| `B4:**` | Col B to end, row 4 to end |

## Multi-Level Header Support

Real-world spreadsheets often have hierarchical headers spanning multiple rows. xlsx2semantic merges them into a single, structured tag name — both in auto-detect mode and with explicit hints.

```
Original Excel layout (rows 4–6):

Row 4:  |         | Total   |         | Race/Ethnicity                    | ...
Row 5:  |         |         |         | American Indian | Asian           | ...
Row 6:  | State   | Number  | Percent | Number | Percent | Number | Percent | ...
```

These multi-level headers become:

```xml
<schema>
  <row-key index="2" attribute="state"/>
  <column index="3" tag="total_number"/>
  <column index="4" tag="total_percent"/>
  <column index="5" tag="race_ethnicity_american_indian_number"/>
  <column index="6" tag="race_ethnicity_american_indian_percent"/>
  <column index="7" tag="race_ethnicity_asian_number"/>
  <column index="8" tag="race_ethnicity_asian_percent"/>
</schema>
```

Each header level is joined with `_`, so the full hierarchy is preserved in a flat, LLM-readable tag name. Duplicate values from merged cells are automatically deduplicated.

## Three Output Layers

xlsx2semantic gives you three views of the same spreadsheet:

| Layer | Description | Access |
|-------|-------------|--------|
| **Raw XML** | Original OOXML extracted from ZIP | `result.xml_entries` |
| **Cell XML** | `<c>` → `<cell>` with resolved styles, types, values | `result.cell_xml` |
| **Semantic XML** | Headers as tags, data as records | `result.semantic_xml` |

### Cell XML Example

```xml
<cell ref="B7" row="7" col="2" styleIndex="59"
      font="Arial 11pt Bold" numberFormat="#,##0" fill="solid FF333399"
      type="sharedString" rawValue="74" value="Alabama"/>
```

Every opaque attribute is resolved:
- `t="s"` → `type="sharedString"`, `value="Alabama"`
- `s="59"` → `font="Arial 11pt Bold"`, `numberFormat="#,##0"`

## How It Works

```
┌──────────┐     ┌─────────────┐     ┌──────────────┐     ┌───────────────┐     ┌──────────────┐
│ .xlsx    │────▶│ ZIP Extract │────▶│ Shared Str   │────▶│ Structural    │────▶│ Semantic     │
│ (OOXML)  │     │ Raw XML     │     │ Style Resolve│     │ Anchor Detect │     │ Transform    │
└──────────┘     └─────────────┘     └──────────────┘     └───────────────┘     └──────────────┘
                       │                    │                     │                     │
                  xml_entries           cell_xml            table regions         semantic_xml
```

1. **ZIP Extract** — XLSX is a ZIP archive. Extract all XML entries.
2. **Shared Strings** — Resolve `<v>74</v>` → `"Alabama"` via `sharedStrings.xml`.
3. **Style Resolve** — Map `s="59"` → font, number format, fill, alignment via `styles.xml`.
4. **Structural Anchor Detection** — Cluster non-empty cells into table regions; detect titles, headers, data boundaries, and row-key columns.
5. **Semantic Transform** — Generate semantic XML for each detected table.

### Performance

- **Streaming parser** — `iterparse` instead of DOM parsing for low memory usage on large files.
- **Parallel processing** — Sheets are parsed concurrently via `multiprocessing` when ≥2 sheets are present.

## Use Cases

- **RAG pipelines** — Feed structured spreadsheet data into retrieval systems
- **LLM tool use** — Let agents query spreadsheet data via semantic XML
- **Data extraction** — Convert messy government/financial Excel reports into clean structure
- **Spreadsheet QA** — Ask natural language questions about tabular data

## Development

```bash
git clone https://github.com/kangminjugit/xlsx2semantic.git
cd xlsx2semantic
pip install -e ".[dev,cli]"
pytest -v
```

## License

MIT
