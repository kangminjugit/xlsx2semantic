"""Data models for parse results."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class ParseResult:
    """Result of parsing an XLSX file."""

    file_name: str
    file_type: str = "xlsx"

    # Raw XML entries extracted from the OOXML ZIP
    xml_entries: dict[str, str] = field(default_factory=dict)

    # Cell-transformed XML (per sheet): <c> → <cell> with resolved attributes
    cell_xml: dict[str, str] = field(default_factory=dict)

    # Semantic table XML (per sheet): headers become tag names
    semantic_xml: dict[str, str] = field(default_factory=dict)

    # Metadata about the parsed file
    metadata: dict[str, object] = field(default_factory=dict)

    parsed_at: datetime = field(default_factory=datetime.now)

    def without_xml_entries(self) -> ParseResult:
        """Return a copy with xml_entries cleared."""
        return ParseResult(
            file_name=self.file_name,
            file_type=self.file_type,
            xml_entries={},
            cell_xml=self.cell_xml,
            semantic_xml=self.semantic_xml,
            metadata=self.metadata,
            parsed_at=self.parsed_at,
        )
