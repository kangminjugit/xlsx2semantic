"""Parse all XLSX files under tests/docs/ and write semantic XML to tests/output/."""

import os
from pathlib import Path

import pytest

from xlsx2semantic import parse_file

DOCS_DIR = Path(__file__).parent / "docs"
OUTPUT_DIR = Path(__file__).parent / "output"
INCLUDE_TRACE = os.getenv("XLSX2SEMANTIC_INCLUDE_TRACE", "").lower() in {"1", "true", "yes", "on"}
MAX_WORKERS = 1


def xlsx_files():
    return sorted(DOCS_DIR.glob("*.xlsx"))


@pytest.mark.parametrize("xlsx_path", xlsx_files(), ids=lambda p: p.name)
def test_parse_and_save(xlsx_path: Path):
    OUTPUT_DIR.mkdir(exist_ok=True)

    result = parse_file(
        xlsx_path,
        include_trace=INCLUDE_TRACE,
        max_workers=MAX_WORKERS,
    )

    output_path = OUTPUT_DIR / (xlsx_path.stem + ".txt")
    with output_path.open("w", encoding="utf-8") as f:
        for sheet, xml in result.semantic_xml.items():
            f.write(f"=== {sheet} ===\n")
            f.write(xml)
            f.write("\n")

    assert output_path.exists()
    assert output_path.stat().st_size > 0
