"""Resolve XLSX style indices to meaningful attributes.

Parses styles.xml to map style index (s="59") to human-readable properties:
  - font: name, size, bold, italic, underline, color
  - numberFormat: format code string
  - fill: pattern type and foreground color
  - alignment: horizontal, vertical, wrap
"""

from __future__ import annotations

from dataclasses import dataclass
from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_MAP = {"s": NS}

# Excel built-in number formats (0–49)
BUILTIN_NUM_FMTS: dict[int, str] = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
    18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}

TYPE_MAP: dict[str, str] = {
    "s": "sharedString",
    "str": "formulaString",
    "b": "boolean",
    "e": "error",
    "inlineStr": "inlineString",
}


@dataclass
class FontInfo:
    name: str | None = None
    size: str | None = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: str | None = None

    def describe(self) -> str:
        parts: list[str] = []
        if self.name:
            parts.append(self.name)
        if self.size:
            parts.append(f"{self.size}pt")
        if self.bold:
            parts.append("Bold")
        if self.italic:
            parts.append("Italic")
        if self.underline:
            parts.append("Underline")
        if self.color:
            parts.append(f"color:{self.color}")
        return " ".join(parts)


@dataclass
class FillInfo:
    pattern_type: str | None = None
    fg_color: str | None = None

    def describe(self) -> str | None:
        if not self.pattern_type:
            return None
        parts = [self.pattern_type]
        if self.fg_color:
            parts.append(self.fg_color)
        return " ".join(parts)


@dataclass
class CellXf:
    num_fmt_id: int = 0
    font_id: int = 0
    fill_id: int = 0
    border_id: int = 0
    alignment: str | None = None


class StyleResolver:
    """Resolves XLSX style indices to human-readable attributes."""

    def __init__(
        self,
        num_fmt_map: dict[int, str],
        fonts: list[FontInfo],
        fills: list[FillInfo],
        cell_xfs: list[CellXf],
    ):
        self._num_fmt_map = num_fmt_map
        self._fonts = fonts
        self._fills = fills
        self._cell_xfs = cell_xfs

    @staticmethod
    def parse(styles_xml: str | None) -> StyleResolver:
        """Parse styles.xml string into a resolver."""
        if not styles_xml or not styles_xml.strip():
            return StyleResolver({}, [], [], [])

        try:
            root = etree.fromstring(styles_xml.encode("utf-8"))

            num_fmt_map = dict(BUILTIN_NUM_FMTS)
            _parse_num_fmts(root, num_fmt_map)

            fonts = _parse_fonts(root)
            fills = _parse_fills(root)
            cell_xfs = _parse_cell_xfs(root)

            return StyleResolver(num_fmt_map, fonts, fills, cell_xfs)
        except Exception:
            return StyleResolver({}, [], [], [])

    def resolve(self, style_index: int) -> dict[str, str]:
        """Resolve a style index to a dict of meaningful attributes."""
        result: dict[str, str] = {}
        if style_index < 0 or style_index >= len(self._cell_xfs):
            return result

        xf = self._cell_xfs[style_index]

        num_fmt = self._num_fmt_map.get(xf.num_fmt_id)
        if num_fmt and num_fmt != "General":
            result["numberFormat"] = num_fmt

        if 0 <= xf.font_id < len(self._fonts):
            result["font"] = self._fonts[xf.font_id].describe()

        if 0 <= xf.fill_id < len(self._fills):
            desc = self._fills[xf.fill_id].describe()
            if desc:
                result["fill"] = desc

        if xf.alignment:
            result["alignment"] = xf.alignment

        return result

    @property
    def is_empty(self) -> bool:
        return len(self._cell_xfs) == 0


# ── Internal parsing helpers ──


def _find(el: etree._Element, xpath: str) -> list[etree._Element]:
    return el.findall(xpath, _NS_MAP)


def _find_first(el: etree._Element, xpath: str) -> etree._Element | None:
    return el.find(xpath, _NS_MAP)


def _int_attr(el: etree._Element, attr: str, default: int = 0) -> int:
    val = el.get(attr)
    if val is None:
        return default
    try:
        return int(val)
    except ValueError:
        return default


def _parse_num_fmts(root: etree._Element, out: dict[int, str]) -> None:
    for el in _find(root, ".//s:numFmt"):
        fmt_id = _int_attr(el, "numFmtId", -1)
        code = el.get("formatCode", "")
        if fmt_id >= 0 and code:
            out[fmt_id] = code


def _parse_fonts(root: etree._Element) -> list[FontInfo]:
    fonts_el = _find_first(root, "s:fonts")
    if fonts_el is None:
        return []
    result: list[FontInfo] = []
    for font_el in fonts_el.findall("s:font", _NS_MAP):
        name_el = _find_first(font_el, "s:name")
        sz_el = _find_first(font_el, "s:sz")
        color_el = _find_first(font_el, "s:color")

        color = None
        if color_el is not None:
            color = color_el.get("rgb") or (f"theme:{color_el.get('theme')}" if color_el.get("theme") else None)

        result.append(FontInfo(
            name=name_el.get("val") if name_el is not None else None,
            size=sz_el.get("val") if sz_el is not None else None,
            bold=_find_first(font_el, "s:b") is not None,
            italic=_find_first(font_el, "s:i") is not None,
            underline=_find_first(font_el, "s:u") is not None,
            color=color,
        ))
    return result


def _parse_fills(root: etree._Element) -> list[FillInfo]:
    fills_el = _find_first(root, "s:fills")
    if fills_el is None:
        return []
    result: list[FillInfo] = []
    for fill_el in fills_el.findall("s:fill", _NS_MAP):
        pf = _find_first(fill_el, "s:patternFill")
        if pf is None:
            result.append(FillInfo())
            continue
        pattern_type = pf.get("patternType", "")
        if not pattern_type or pattern_type == "none":
            result.append(FillInfo())
            continue
        fg_el = _find_first(pf, "s:fgColor")
        fg_color = None
        if fg_el is not None:
            fg_color = fg_el.get("rgb") or (f"theme:{fg_el.get('theme')}" if fg_el.get("theme") else None)
        result.append(FillInfo(pattern_type=pattern_type, fg_color=fg_color))
    return result


def _parse_cell_xfs(root: etree._Element) -> list[CellXf]:
    cell_xfs_el = _find_first(root, "s:cellXfs")
    if cell_xfs_el is None:
        return []
    result: list[CellXf] = []
    for xf_el in cell_xfs_el.findall("s:xf", _NS_MAP):
        align_el = _find_first(xf_el, "s:alignment")
        alignment = None
        if align_el is not None:
            alignment = _describe_alignment(align_el)
        result.append(CellXf(
            num_fmt_id=_int_attr(xf_el, "numFmtId"),
            font_id=_int_attr(xf_el, "fontId"),
            fill_id=_int_attr(xf_el, "fillId"),
            border_id=_int_attr(xf_el, "borderId"),
            alignment=alignment,
        ))
    return result


def _describe_alignment(el: etree._Element) -> str | None:
    parts: list[str] = []
    h = el.get("horizontal")
    if h:
        parts.append(h)
    v = el.get("vertical")
    if v:
        parts.append(f"v:{v}")
    if el.get("wrapText") == "1":
        parts.append("wrap")
    return " ".join(parts) if parts else None
