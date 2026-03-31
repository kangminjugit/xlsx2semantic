"""OOXML ZIP extraction utilities."""

from __future__ import annotations

import zipfile
from io import BytesIO


def extract_xml_entries(data: bytes, *, selective: bool = False) -> dict[str, str]:
    """Extract XML/rels entries from an OOXML ZIP archive.

    Args:
        data: Raw bytes of the OOXML file.
        selective: If True, extract only the entries needed for
            transformation (sharedStrings, styles, worksheets).
            If False, extract all XML/rels entries.

    Returns:
        Dict mapping entry paths to XML strings.
    """
    entries: dict[str, str] = {}
    with zipfile.ZipFile(BytesIO(data)) as zf:
        for name in zf.namelist():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            if selective and not _is_needed(name):
                continue
            entries[name] = zf.read(name).decode("utf-8")
    return entries


def _is_needed(name: str) -> bool:
    """Return True if the entry is needed for transformation."""
    return (
        name == "xl/sharedStrings.xml"
        or name == "xl/styles.xml"
        or (name.startswith("xl/worksheets/") and name.endswith(".xml"))
    )
