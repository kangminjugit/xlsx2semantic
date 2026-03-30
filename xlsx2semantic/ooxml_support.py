"""OOXML ZIP extraction utilities."""

from __future__ import annotations

import logging
import zipfile
from io import BytesIO
from xml.dom import minidom

logger = logging.getLogger(__name__)


def extract_xml_entries(data: bytes, *, selective: bool = False) -> dict[str, str]:
    """Extract XML/rels entries from an OOXML ZIP archive.

    Args:
        data: Raw bytes of the OOXML file.
        selective: If True, extract only the entries needed for
            transformation (sharedStrings, styles, worksheets).
            If False, extract all XML/rels entries.

    Returns:
        Dict mapping entry paths to pretty-printed XML strings.
    """
    entries: dict[str, str] = {}
    with zipfile.ZipFile(BytesIO(data)) as zf:
        for name in zf.namelist():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            if selective and not _is_needed(name):
                continue
            raw = zf.read(name).decode("utf-8")
            entries[name] = _pretty_print_xml(raw)
    return entries


def _is_needed(name: str) -> bool:
    """Return True if the entry is needed for transformation."""
    return (
        name == "xl/sharedStrings.xml"
        or name == "xl/styles.xml"
        or (name.startswith("xl/worksheets/") and name.endswith(".xml"))
    )


def _pretty_print_xml(raw_xml: str) -> str:
    """Pretty-print XML string with 2-space indentation."""
    try:
        dom = minidom.parseString(raw_xml)
        pretty = dom.toprettyxml(indent="  ", encoding=None)
        # Remove extra blank lines that minidom adds
        lines = [line for line in pretty.splitlines() if line.strip()]
        return "\n".join(lines)
    except Exception:
        logger.warning("Failed to pretty-print XML entry", exc_info=True)
        return raw_xml
