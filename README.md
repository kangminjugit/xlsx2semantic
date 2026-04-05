# xlsx2semantic

**Transform XLSX into LLM-friendly semantic XML.**

Raw Excel XML is unreadable — cell references like `<c r="B7" t="s"><v>74</v></c>` mean nothing to an LLM.
**xlsx2semantic** converts that into structured, self-describing XML that any language model can instantly understand.

```
Before (raw OOXML)                          After (semantic XML)
─────────────────────                       ────────────────────
<c r="B7" s="59" t="s">                    <record row="7">
  <v>74</v>                                   <state>Alabama</state>
</c>                                          <total_number>745938</total_number>
<c r="C7" s="60">                             <total_percent>100</total_percent>
  <v>745938</v>                             </record>
</c>
```

## Why?

| Problem | xlsx2semantic |
|---------|---------------|
| LLMs can't parse raw OOXML cell references | Headers become tag names, data becomes readable records |
| Shared strings are just index numbers | Automatically resolved to actual text |
| Merged cells break structure | Horizontal and vertical merges propagated correctly across the grid |
| Multi-level headers produce redundant tag names | Only the minimum levels needed for uniqueness are used |
| Style indices are opaque | Resolved to human-readable font, color, format info |
| One sheet with multiple tables | Structural anchors auto-detect each table region |
| Formula error values (`#REF!`, `#DIV/0!`, …) | Filtered out — error cells are treated as empty |

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

판단 근거까지 보고 싶으면:

```python
result = parse_file("enrollment.xlsx", include_trace=True)
print(result.semantic_xml["xl/worksheets/sheet1.xml"])
```

이 경우 출력 XML에 `<trace>`가 추가되어 선택된 `title/header/data-start-row`, 최종 confidence, 행별 점수가 함께 기록됩니다.

```xml
<trace mode="auto" confidence="0.922">
  <selected header-rows="4,5" data-start-row="6"/>
  <evidence name="header_block" score="1.411"/>
  <row index="4" role="header" title-score="0.250" header-score="0.913" data-score="0.300"/>
</trace>
```

Output:

```xml
<semantic-table>
  <title>Public school enrollment by race/ethnicity</title>
  <schema>
    <column index="2" tag="state"/>
    <column index="3" tag="total_number"/>
    <column index="4" tag="total_percent"/>
  </schema>
  <records count="52">
    <record row="8">
      <state>Alabama</state>
      <total_number>745938</total_number>
      <total_percent>100</total_percent>
    </record>
    <record row="9">
      <state>Alaska</state>
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
xlsx2semantic data.xlsx --include-trace  # append semantic decision trace
```

Install CLI extras: `pip install xlsx2semantic[cli]`

## Automatic Table Detection (Structural Anchors)

xlsx2semantic uses a [SpreadsheetLLM](https://arxiv.org/abs/2407.09025)-inspired approach to automatically detect table regions without manual hints:

1. **Cell clustering** — Non-empty cells are grouped into connected regions by row/column proximity. Clusters separated by empty rows or columns become separate tables.
2. **Title detection** — Full-width merged cells or sparse single-value rows at the top of a region are recognized as titles.
3. **Header detection** — By default, the first non-title row is treated as the single header row. If header cells are vertically merged downward (spanning multiple rows), those rows are automatically absorbed into the header block.
4. **Minimum-unique tag names** — Multi-level headers are combined into the shortest tag name that remains unique across all columns. Parent levels are added only when the leaf level alone would be ambiguous.
5. **Noise filtering** — Data rows whose values exactly match header labels (repeated template headers) are automatically skipped.

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
    row_meta_col="B",          # single column → becomes record attribute
    # row_meta_col="A:C",      # column range → all three become record attributes
    header_rows=2,             # assume 2 header rows (default: 1)
)
```

`header_rows` controls how many rows are treated as column headers in auto-detect mode. Use it when a sheet has a fixed multi-level header that vertical merge expansion alone cannot determine.

| Pattern | Meaning |
|---------|---------|
| `B4:Z6` | Exact range: col B–Z, row 4–6 |
| `B4:*6` | Col B to **end of sheet**, row 4–6 |
| `B4:Z*` | Col B–Z, row 4 to **end of sheet** |
| `B4:**` | Col B to end, row 4 to end |

When `row_meta_col` is provided, the specified columns are rendered as record **attributes** instead of child tags:

```xml
<!-- row_meta_col="B" -->
<record row="8" state="Alabama">
  <total_number>745938</total_number>
</record>

<!-- row_meta_col="A:C" (multi-column) -->
<record row="8" 대분류="식품" 중분류="음료" 소분류="탄산음료">
  <매출액>1500000</매출액>
</record>
```

## Multi-Level Header Support

Real-world spreadsheets often have hierarchical headers spanning multiple rows. xlsx2semantic uses only the **minimum levels needed** to make each tag name unique — parent levels are added only when the leaf alone would be ambiguous.

```
Original Excel layout (rows 4–6):

Row 4:  |         | Total   |         | Race/Ethnicity                    | ...
Row 5:  |         |         |         | American Indian | Asian           | ...
Row 6:  | State   | Number  | Percent | Number | Percent | Number | Percent | ...
```

These multi-level headers become:

```xml
<schema>
  <column index="2" tag="state"/>
  <column index="3" tag="total_number"/>
  <column index="4" tag="total_percent"/>
  <column index="5" tag="american_indian_number"/>
  <column index="6" tag="american_indian_percent"/>
  <column index="7" tag="asian_number"/>
  <column index="8" tag="asian_percent"/>
</schema>
```

`Number` and `Percent` appear in multiple columns, so one parent level (`American Indian`, `Asian`, etc.) is added to disambiguate. The top-level `Race/Ethnicity` is dropped because it is not needed for uniqueness. Duplicate values from merged cells are automatically deduplicated.

### Vertical Merge Extension

When header cells are vertically merged (e.g. a column label spanning rows 1–3), xlsx2semantic automatically extends the header block to include all covered rows:

```
Row 1: | 대여소번호 (merged ↓) | 소재지(위치) ──────────── | 설치시기 (merged ↓) |
Row 2: | (merged)             | 자치구 | 상세주소 | 위도 | 경도 | (merged)          |
       └── both rows become headers; data starts at row 3
```

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
4. **Structural Anchor Detection** — Cluster non-empty cells into table regions; detect titles, headers (with vertical merge extension), and data boundaries.
5. **Semantic Transform** — Generate minimum-unique tag names and emit one `<record>` per data row.

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
