#!/usr/bin/env python3
"""Extract text from an HWPX document using lxml (no external hwpx dependency).

Usage:
    python text_extract.py document.hwpx
    python text_extract.py document.hwpx --format markdown
    python text_extract.py document.hwpx --include-tables
"""

import argparse
import sys
import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
}


def _read_section_xml(hwpx_path: str) -> etree._Element:
    with zipfile.ZipFile(hwpx_path, "r") as zf:
        data = zf.read("Contents/section0.xml")
    return etree.parse(BytesIO(data)).getroot()


def _collect_text(elem, *, include_tables: bool = True) -> list[str]:
    """Collect text from <hp:t> nodes, optionally including table cells."""
    lines: list[str] = []
    root = elem

    if include_tables:
        for t in root.findall(".//hp:t", NS):
            text = "".join(t.itertext())
            if text.strip():
                lines.append(text)
    else:
        # Only top-level paragraphs (skip table content)
        top_paras = root.findall("hp:p", NS)
        for p in top_paras:
            for t in p.findall(".//hp:t", NS):
                text = "".join(t.itertext())
                if text.strip():
                    lines.append(text)
    return lines


def extract_plain(hwpx_path: str, *, include_tables: bool = False) -> str:
    root = _read_section_xml(hwpx_path)
    lines = _collect_text(root, include_tables=include_tables)
    return "\n".join(lines)


def extract_markdown(hwpx_path: str) -> str:
    root = _read_section_xml(hwpx_path)
    lines: list[str] = []

    # Top-level paragraphs
    for p in root.findall("hp:p", NS):
        # Check for tables
        tables = p.findall(".//hp:tbl", NS)
        if tables:
            for tbl in tables:
                rows = tbl.findall(".//hp:tr", NS)
                for row in rows:
                    cells = row.findall(".//hp:tc", NS)
                    cell_texts = []
                    for cell in cells:
                        t_texts = []
                        for t in cell.findall(".//hp:t", NS):
                            text = "".join(t.itertext())
                            if text.strip():
                                t_texts.append(text.strip())
                        cell_texts.append(" ".join(t_texts))
                    lines.append("| " + " | ".join(cell_texts) + " |")
            lines.append("")
        else:
            p_texts = []
            for t in p.findall(".//hp:t", NS):
                text = "".join(t.itertext())
                if text.strip():
                    p_texts.append(text)
            if p_texts:
                lines.append("".join(p_texts))

    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract text from an HWPX document"
    )
    parser.add_argument("input", help="Path to .hwpx file")
    parser.add_argument(
        "--format", "-f",
        choices=["plain", "markdown"],
        default="plain",
        help="Output format (default: plain)",
    )
    parser.add_argument(
        "--include-tables",
        action="store_true",
        help="Include text from tables and nested objects (plain mode)",
    )
    parser.add_argument(
        "--output", "-o",
        help="Output file path (default: stdout)",
    )
    args = parser.parse_args()

    if not Path(args.input).is_file():
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if args.format == "markdown":
        result = extract_markdown(args.input)
    else:
        result = extract_plain(args.input, include_tables=args.include_tables)

    if args.output:
        Path(args.output).write_text(result, encoding="utf-8")
        print(f"Extracted to: {args.output}", file=sys.stderr)
    else:
        print(result)


if __name__ == "__main__":
    main()
