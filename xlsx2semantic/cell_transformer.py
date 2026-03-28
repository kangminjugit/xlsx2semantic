"""Transform <c> cell tags into enriched <cell> tags.

Before:
  <c r="B7" s="59" t="s"><v>74</v></c>

After:
  <cell ref="B7" row="7" col="2" styleIndex="59"
        font="Arial 11pt Bold" numberFormat="#,##0"
        type="sharedString" rawValue="74" value="Alabama"/>
"""

from __future__ import annotations

import logging
import re

from lxml import etree

logger = logging.getLogger(__name__)

from xlsx2semantic.style_resolver import NS, TYPE_MAP, StyleResolver

_CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")
_NS_MAP = {"s": NS}


def parse_shared_strings(shared_strings_xml: str | None) -> list[str]:
    """Parse sharedStrings.xml and return a list of string values."""
    strings: list[str] = []
    if not shared_strings_xml or not shared_strings_xml.strip():
        return strings

    try:
        root = etree.fromstring(shared_strings_xml.encode("utf-8"))
        for si in root.findall("s:si", _NS_MAP):
            # Collect all <t> elements (handles rich text <r><t>...</t></r>)
            texts = []
            for t in si.iter(f"{{{NS}}}t"):
                if t.text:
                    texts.append(t.text)
            strings.append("".join(texts))
    except Exception:
        logger.warning("Failed to parse sharedStrings.xml", exc_info=True)
    return strings


def transform_sheet(
    sheet_xml: str,
    shared_strings: list[str],
    style_resolver: StyleResolver | None = None,
) -> str:
    """Transform all <c> elements in sheet XML to enriched <cell> elements."""
    try:
        root = etree.fromstring(sheet_xml.encode("utf-8"))
    except Exception:
        logger.warning("Failed to parse sheet XML for cell transformation", exc_info=True)
        return sheet_xml

    # Collect all <c> elements
    c_elements = root.iter(f"{{{NS}}}c")
    # Need to materialize since we modify the tree
    cells = list(c_elements)

    for c in cells:
        cell = _transform_cell(c, shared_strings, style_resolver)
        parent = c.getparent()
        if parent is not None:
            idx = list(parent).index(c)
            parent.remove(c)
            parent.insert(idx, cell)

    return etree.tostring(root, pretty_print=True, xml_declaration=False, encoding="unicode")


def _transform_cell(
    c: etree._Element,
    shared_strings: list[str],
    style_resolver: StyleResolver | None,
) -> etree._Element:
    cell = etree.Element(f"{{{NS}}}cell")

    ref = c.get("r", "")
    cell.set("ref", ref)

    m = _CELL_REF_RE.match(ref)
    if m:
        cell.set("row", m.group(2))
        cell.set("col", str(_col_to_num(m.group(1))))

    # Style resolution
    s_val = c.get("s")
    if s_val:
        cell.set("styleIndex", s_val)
        if style_resolver and not style_resolver.is_empty:
            try:
                resolved = style_resolver.resolve(int(s_val))
                for k, v in resolved.items():
                    cell.set(k, v)
            except (ValueError, IndexError):
                pass

    # Type mapping
    raw_type = c.get("t", "")
    mapped_type = TYPE_MAP.get(raw_type, raw_type) if raw_type else "number"
    cell.set("type", mapped_type)

    # Value and formula
    v_el = c.find(f"{{{NS}}}v")
    if v_el is None:
        v_el = c.find("v")
    f_el = c.find(f"{{{NS}}}f")
    if f_el is None:
        f_el = c.find("f")

    if f_el is not None and f_el.text:
        cell.set("formula", f_el.text)

    if v_el is not None and v_el.text is not None:
        raw_value = v_el.text
        cell.set("rawValue", raw_value)

        if raw_type == "s" and shared_strings:
            try:
                idx = int(raw_value)
                if 0 <= idx < len(shared_strings):
                    cell.set("value", shared_strings[idx])
                else:
                    cell.set("value", raw_value)
            except ValueError:
                cell.set("value", raw_value)
        else:
            cell.set("value", raw_value)

    return cell


def _col_to_num(letters: str) -> int:
    result = 0
    for ch in letters:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result
