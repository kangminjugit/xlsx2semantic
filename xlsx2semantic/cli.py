"""CLI interface for xlsx2semantic."""

import sys
from pathlib import Path
from typing import Optional

try:
    import typer
    from rich.console import Console
    from rich.syntax import Syntax
except ImportError:
    print("CLI dependencies not installed. Run: pip install xlsx2semantic[cli]", file=sys.stderr)
    sys.exit(1)

from xlsx2semantic.core import parse_file

app = typer.Typer(
    name="xlsx2semantic",
    help="Transform XLSX files into LLM-friendly semantic XML.",
    add_completion=False,
)
console = Console()


@app.command()
def main(
    file: Path = typer.Argument(..., help="Path to the XLSX file", exists=True),
    output: Optional[Path] = typer.Option(None, "-o", "--output", help="Output file path (default: stdout)"),
    raw: bool = typer.Option(False, "--raw", help="Output raw XML entries only"),
    title_range: Optional[str] = typer.Option(None, "--title-range", help='Title range, e.g. "B2:Z2" or "B2:*2"'),
    header_range: Optional[str] = typer.Option(None, "--header-range", help='Header range, e.g. "B4:Z6" or "B4:**"'),
    row_meta_col: Optional[str] = typer.Option(None, "--row-meta-col", help='Row meta column, e.g. "B"'),
    include_trace: bool = typer.Option(False, "--include-trace", help="Append semantic decision trace to XML output"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="Sheet name filter (default: all sheets)"),
    mode: str = typer.Option(
        "semantic",
        "--mode",
        help="Output mode: semantic (default), cell, raw-xml, all",
    ),
) -> None:
    """Parse an XLSX file and output semantic XML."""
    result = parse_file(
        file,
        raw=raw,
        title_range=title_range,
        header_range=header_range,
        row_meta_col=row_meta_col,
        include_trace=include_trace,
    )

    # Select output based on mode
    if mode == "raw-xml":
        sections = result.xml_entries
    elif mode == "cell":
        sections = result.cell_xml
    elif mode == "all":
        sections = {}
        for name, xml in result.semantic_xml.items():
            sections[f"[semantic] {name}"] = xml
        for name, xml in result.cell_xml.items():
            sections[f"[cell] {name}"] = xml
    else:
        sections = result.semantic_xml

    # Filter by sheet name if specified
    if sheet:
        sections = {k: v for k, v in sections.items() if sheet.lower() in k.lower()}

    if not sections:
        console.print("[yellow]No matching sheets found.[/yellow]")
        raise typer.Exit(1)

    # Output
    combined = ""
    for name, xml in sections.items():
        combined += f"<!-- {name} -->\n{xml}\n\n"
    combined = combined.strip()

    if output:
        output.write_text(combined, encoding="utf-8")
        console.print(f"[green]Written to {output}[/green]")
    else:
        syntax = Syntax(combined, "xml", theme="monokai", line_numbers=True)
        console.print(syntax)


if __name__ == "__main__":
    app()
