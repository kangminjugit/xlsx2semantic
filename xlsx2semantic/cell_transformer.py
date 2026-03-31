"""Transform <c> cell tags into enriched <cell> tags.

Before:
  <c r="B7" s="59" t="s"><v>74</v></c>

After:
  <cell ref="B7" row="7" col="2" styleIndex="59"
        font="Arial 11pt Bold" numberFormat="#,##0"
        type="sharedString" rawValue="74" value="Alabama"/>
"""

from __future__ import annotations

import io
import logging
from xml.sax.saxutils import quoteattr

from lxml import etree

logger = logging.getLogger(__name__)

from xlsx2semantic.sheet_scanner import SheetData, scan_sheet
from xlsx2semantic.style_resolver import NS, TYPE_MAP, StyleResolver


def parse_shared_strings(shared_strings_xml: str | None) -> list[str]:
    """Parse sharedStrings.xml and return a list of string values."""
    strings: list[str] = []
    if not shared_strings_xml or not shared_strings_xml.strip():
        return strings

    try:
        context = etree.iterparse(
            io.BytesIO(shared_strings_xml.encode("utf-8")),
            events=("end",),
            tag=f"{{{NS}}}si",
        )
        for _, si in context:
            texts = []
            for t in si.iter(f"{{{NS}}}t"):
                if t.text:
                    texts.append(t.text)
            strings.append("".join(texts))
            si.clear()
    except Exception:
        logger.warning("Failed to parse sharedStrings.xml", exc_info=True)
    return strings


def transform_sheet(
    sheet_xml_or_data: str | SheetData,
    shared_strings: list[str],
    style_resolver: StyleResolver | None = None,
) -> str:
    """Transform sheet XML to enriched cell XML.

    Args:
        sheet_xml_or_data: Raw XML string or pre-scanned SheetData.
        shared_strings: Parsed shared string table.
        style_resolver: Optional style resolver for cell formatting.

    Returns:
        XML string with enriched <cell> elements.
    """
    if isinstance(sheet_xml_or_data, SheetData):
        data = sheet_xml_or_data
    else:
        data = scan_sheet(sheet_xml_or_data, shared_strings)

    return _build_cell_xml(data, style_resolver)


def _build_cell_xml(data: SheetData, style_resolver: StyleResolver | None) -> str:
    """Build cell XML string directly from SheetData (no DOM)."""
    parts: list[str] = [f'<worksheet xmlns="{NS}">', "<sheetData>"]
    current_row = -1

    for cell in data.cells:
        if cell.row != current_row:
            if current_row != -1:
                parts.append("</row>")
            parts.append(f'<row r="{cell.row}">')
            current_row = cell.row

        attrs = [
            f'ref="{cell.ref}"',
            f'row="{cell.row}"',
            f'col="{cell.col}"',
        ]

        if cell.style_idx:
            attrs.append(f'styleIndex="{cell.style_idx}"')
            if style_resolver and not style_resolver.is_empty:
                try:
                    resolved = style_resolver.resolve(int(cell.style_idx))
                    for k, v in resolved.items():
                        attrs.append(f"{k}={quoteattr(v)}")
                except (ValueError, IndexError):
                    pass

        mapped_type = TYPE_MAP.get(cell.cell_type, cell.cell_type) if cell.cell_type else "number"
        attrs.append(f'type="{mapped_type}"')

        if cell.formula:
            attrs.append(f"formula={quoteattr(cell.formula)}")

        if cell.raw_value is not None:
            attrs.append(f"rawValue={quoteattr(cell.raw_value)}")

        if cell.value is not None:
            attrs.append(f"value={quoteattr(cell.value)}")

        parts.append(f'<cell {" ".join(attrs)}/>')

    if current_row != -1:
        parts.append("</row>")

    parts.append("</sheetData>")
    parts.append("</worksheet>")

    return "\n".join(parts)


def _col_to_num(letters: str) -> int:
    result = 0
    for ch in letters:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result
