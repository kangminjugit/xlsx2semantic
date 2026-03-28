"""xlsx2semantic: Transform XLSX files into LLM-friendly semantic XML."""

from xlsx2semantic.core import parse, parse_file
from xlsx2semantic.layout_hint import TableLayoutHint
from xlsx2semantic.models import ParseResult

__version__ = "0.1.1"
__all__ = ["parse", "parse_file", "TableLayoutHint", "ParseResult"]
