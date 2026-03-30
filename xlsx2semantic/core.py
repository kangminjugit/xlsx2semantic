"""Core API — main entry points for xlsx2semantic."""

from __future__ import annotations

from pathlib import Path

from xlsx2semantic.cell_transformer import parse_shared_strings, transform_sheet
from xlsx2semantic.layout_hint import TableLayoutHint
from xlsx2semantic.models import ParseResult
from xlsx2semantic.ooxml_support import extract_xml_entries
from xlsx2semantic.semantic import transform as semantic_transform
from xlsx2semantic.style_resolver import StyleResolver


def parse(
    data: bytes,
    *,
    file_name: str = "unknown.xlsx",
    raw: bool = False,
    title_range: str | None = None,
    header_range: str | None = None,
    row_meta_col: str | None = None,
) -> ParseResult:
    """Parse XLSX bytes into structured result.

    Args:
        data: Raw bytes of the XLSX file.
        file_name: Original file name (for metadata).
        raw: If True, return only raw XML entries without transformation.
        title_range: e.g. "B2:Z2", "B2:*2"
        header_range: e.g. "B4:Z6", "B4:**"
        row_meta_col: e.g. "B"

    Returns:
        ParseResult with xml_entries, cell_xml, and semantic_xml.
    """
    xml_entries = extract_xml_entries(data, selective=not raw)

    metadata: dict[str, object] = {
        "xml_entry_count": len(xml_entries),
        "xml_entry_names": list(xml_entries.keys()),
        "raw": raw,
    }

    if raw:
        return ParseResult(
            file_name=file_name,
            xml_entries=xml_entries,
            metadata=metadata,
        )

    # Parse shared strings
    shared_strings_xml = xml_entries.get("xl/sharedStrings.xml")
    shared_strings = parse_shared_strings(shared_strings_xml)

    # Parse styles
    styles_xml = xml_entries.get("xl/styles.xml")
    style_resolver = StyleResolver.parse(styles_xml)

    # Build layout hint
    layout_hint = TableLayoutHint.create(title_range, header_range, row_meta_col)

    # Process each worksheet
    cell_xml: dict[str, str] = {}
    semantic_xml: dict[str, str] = {}

    for entry_name, entry_xml in xml_entries.items():
        if entry_name.startswith("xl/worksheets/") and entry_name.endswith(".xml"):
            # Semantic transformation (from original XML)
            sem = semantic_transform(entry_xml, shared_strings, layout_hint)
            semantic_xml[entry_name] = sem

            # Cell transformation (<c> → <cell>)
            cell = transform_sheet(entry_xml, shared_strings, style_resolver)
            cell_xml[entry_name] = cell

    metadata["shared_string_count"] = len(shared_strings)

    return ParseResult(
        file_name=file_name,
        xml_entries=xml_entries,
        cell_xml=cell_xml,
        semantic_xml=semantic_xml,
        metadata=metadata,
    )


def parse_file(
    path: str | Path,
    *,
    raw: bool = False,
    title_range: str | None = None,
    header_range: str | None = None,
    row_meta_col: str | None = None,
) -> ParseResult:
    """Parse an XLSX file from disk.

    Args:
        path: Path to the XLSX file.
        raw: If True, return only raw XML entries without transformation.
        title_range: e.g. "B2:Z2", "B2:*2"
        header_range: e.g. "B4:Z6", "B4:**"
        row_meta_col: e.g. "B"

    Returns:
        ParseResult with xml_entries, cell_xml, and semantic_xml.
    """
    p = Path(path)
    data = p.read_bytes()
    return parse(
        data,
        file_name=p.name,
        raw=raw,
        title_range=title_range,
        header_range=header_range,
        row_meta_col=row_meta_col,
    )
