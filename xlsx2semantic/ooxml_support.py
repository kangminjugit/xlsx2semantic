"""OOXML ZIP extraction utilities."""

from __future__ import annotations

import zipfile
from io import BytesIO
from xml.dom import minidom


def extract_xml_entries(data: bytes) -> dict[str, str]:
    """Extract all XML/rels entries from an OOXML ZIP archive.

    Args:
        data: Raw bytes of the OOXML file.

    Returns:
        Dict mapping entry paths to pretty-printed XML strings.
    """
    entries: dict[str, str] = {}
    with zipfile.ZipFile(BytesIO(data)) as zf:
        for name in zf.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                raw = zf.read(name).decode("utf-8")
                entries[name] = _pretty_print_xml(raw)
    return entries


def _pretty_print_xml(raw_xml: str) -> str:
    """Pretty-print XML string with 2-space indentation."""
    try:
        dom = minidom.parseString(raw_xml)
        pretty = dom.toprettyxml(indent="  ", encoding=None)
        # Remove extra blank lines that minidom adds
        lines = [line for line in pretty.splitlines() if line.strip()]
        return "\n".join(lines)
    except Exception:
        return raw_xml
