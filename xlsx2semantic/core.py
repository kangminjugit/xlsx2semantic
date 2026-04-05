"""Core API — main entry points for xlsx2semantic."""

from __future__ import annotations

import os
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

from xlsx2semantic.cell_transformer import parse_shared_strings, transform_sheet
from xlsx2semantic.layout_hint import TableLayoutHint
from xlsx2semantic.models import ParseResult
from xlsx2semantic.ooxml_support import extract_xml_entries
from xlsx2semantic.semantic import transform as semantic_transform
from xlsx2semantic.sheet_scanner import scan_sheet
from xlsx2semantic.style_resolver import StyleResolver

# Minimum number of sheets to trigger parallel processing.
_PARALLEL_THRESHOLD = 2


def _process_sheet(
    entry_name: str,
    entry_xml: str,
    shared_strings: list[str],
    style_resolver: StyleResolver,
    layout_hint: TableLayoutHint | None,
    include_trace: bool,
    include_cell_roles: bool,
) -> tuple[str, str, str]:
    """Process a single worksheet — designed to run in a worker process.

    Returns:
        (entry_name, semantic_xml, cell_xml)
    """
    sheet_data = scan_sheet(entry_xml, shared_strings)
    sem = semantic_transform(
        sheet_data,
        shared_strings,
        layout_hint,
        include_trace=include_trace,
        include_cell_roles=include_cell_roles,
    )
    cell = transform_sheet(sheet_data, shared_strings, style_resolver)
    return entry_name, sem, cell


def parse(
    data: bytes,
    *,
    file_name: str = "unknown.xlsx",
    raw: bool = False,
    title_range: str | None = None,
    header_range: str | None = None,
    row_meta_col: str | None = None,
    max_workers: int | None = None,
    include_trace: bool = False,
    include_cell_roles: bool = False,
) -> ParseResult:
    """Parse XLSX bytes into structured result.

    Args:
        data: Raw bytes of the XLSX file.
        file_name: Original file name (for metadata).
        raw: If True, return only raw XML entries without transformation.
        title_range: e.g. "B2:Z2", "B2:*2"
        header_range: e.g. "B4:Z6", "B4:**"
        row_meta_col: e.g. "B"
        max_workers: Maximum number of parallel worker processes for sheet
            processing. Defaults to min(sheet_count, cpu_count). Set to 1 to
            disable parallel processing.

    Returns:
        ParseResult with xml_entries, cell_xml, and semantic_xml.
    """
    xml_entries = extract_xml_entries(data, selective=not raw)

    metadata: dict[str, object] = {
        "xml_entry_count": len(xml_entries),
        "xml_entry_names": list(xml_entries.keys()),
        "raw": raw,
        "include_trace": include_trace,
        "include_cell_roles": include_cell_roles,
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

    # Collect worksheet entries
    sheet_entries = {
        name: xml
        for name, xml in xml_entries.items()
        if name.startswith("xl/worksheets/") and name.endswith(".xml")
    }

    cell_xml: dict[str, str] = {}
    semantic_xml: dict[str, str] = {}

    use_parallel = (
        len(sheet_entries) >= _PARALLEL_THRESHOLD
        and max_workers != 1
    )

    if use_parallel:
        workers = max_workers or min(len(sheet_entries), os.cpu_count() or 1)
        with ProcessPoolExecutor(max_workers=workers) as executor:
            futures = {
                executor.submit(
                    _process_sheet,
                    name, xml, shared_strings, style_resolver, layout_hint, include_trace, include_cell_roles,
                ): name
                for name, xml in sheet_entries.items()
            }
            for future in futures:
                entry_name, sem, cell = future.result()
                semantic_xml[entry_name] = sem
                cell_xml[entry_name] = cell
    else:
        for entry_name, entry_xml in sheet_entries.items():
            sheet_data = scan_sheet(entry_xml, shared_strings)
            sem = semantic_transform(
                sheet_data,
                shared_strings,
                layout_hint,
                include_trace=include_trace,
                include_cell_roles=include_cell_roles,
            )
            semantic_xml[entry_name] = sem
            cell = transform_sheet(sheet_data, shared_strings, style_resolver)
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
    max_workers: int | None = None,
    include_trace: bool = False,
    include_cell_roles: bool = False,
) -> ParseResult:
    """Parse an XLSX file from disk.

    Args:
        path: Path to the XLSX file.
        raw: If True, return only raw XML entries without transformation.
        title_range: e.g. "B2:Z2", "B2:*2"
        header_range: e.g. "B4:Z6", "B4:**"
        row_meta_col: e.g. "B"
        max_workers: Maximum number of parallel worker processes for sheet
            processing. Defaults to min(sheet_count, cpu_count). Set to 1 to
            disable parallel processing.

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
        max_workers=max_workers,
        include_trace=include_trace,
        include_cell_roles=include_cell_roles,
    )
