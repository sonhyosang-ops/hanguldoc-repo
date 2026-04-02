#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["fastmcp", "lxml"]
# ///
"""Unified HWPX MCP Server — Markdown→HWPX + HWPX Template→HWPX.

Two workflows in one server:
  1. No template: convert Markdown → HWPX (md2hwpx)
  2. With HWPX template: analyze → extract → build → validate → page_guard (hwpx2hwpx)

Run via: uv run python scripts/mcp_server.py
"""

import os
import sys
import subprocess
import tempfile
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPTS_DIR.parent
sys.path.insert(0, str(SCRIPTS_DIR))

from fastmcp import FastMCP

mcp = FastMCP(
    "hwpx",
    instructions=(
        "Unified HWPX (Hangul Word Processor) document toolkit.\n\n"
        "Two workflows:\n"
        "1. Markdown → HWPX (no template needed): use convert_md_to_hwpx\n"
        "   - Converts .md files into .hwpx with headings, tables, code blocks, "
        "lists, blockquotes, bold, cover page, headers/footers, auto-numbering.\n\n"
        "2. HWPX template → HWPX (reference-based): use analyze_hwpx → "
        "extract_hwpx_xml → (edit section0.xml with Python/lxml) → build_hwpx → validate_hwpx\n"
        "   - Replicates an existing HWPX layout, replacing text content.\n"
        "   - Templates available: gonmun(공문), report(보고서), minutes(회의록), proposal(제안서).\n\n"
        "CRITICAL: section0.xml editing rules (workflow 2):\n"
        "  - Use lxml to parse and modify the extracted section0.xml.\n"
        "  - Element order in section0.xml = render order in the document.\n"
        "    Correct: [section_header_table] → [body_paragraphs] → [next_header] → [body]\n"
        "    Wrong: all body paragraphs first, then all headers at the end.\n"
        "  - When removing items (e.g. TOC entries), delete the entire <hp:p> element\n"
        "    from its parent, not just clear <hp:t> text. Empty <hp:p> still renders\n"
        "    (shows dotted tab leaders, page numbers, blank lines).\n"
        "  - When TOC items are reduced, set the cell's vertAlign='TOP' on <hp:subList>\n"
        "    to prevent content from floating to the center of the box.\n"
        "  - After editing, verify the structure by extracting text from the built HWPX\n"
        "    using extract_text_hwpx to confirm content and ordering.\n"
        "  - page_guard_hwpx is for same-structure validation (e.g. filling in values).\n"
        "    When sections are added/removed, it will report expected differences — \n"
        "    use validate_hwpx for structural integrity instead.\n\n"
        "Only .hwpx files are supported — not legacy .hwp binary format."
    ),
)


# ─── Helpers ────────────────────────────────────────────────────────

def _python() -> str:
    """Return the Python executable to use (venv-aware)."""
    if sys.platform == "win32":
        venv_python = SKILL_DIR / ".venv" / "Scripts" / "python.exe"
    else:
        venv_python = SKILL_DIR / ".venv" / "bin" / "python3"
    if venv_python.exists():
        return str(venv_python)
    return sys.executable


def _run(args: list[str]) -> tuple[str, str, int]:
    result = subprocess.run(args, capture_output=True, text=True)
    return result.stdout, result.stderr, result.returncode


# ═══════════════════════════════════════════════════════════════════
# Workflow 1: Markdown → HWPX (no template needed)
# ═══════════════════════════════════════════════════════════════════

@mcp.tool()
def convert_md_to_hwpx(
    input_md: str,
    output_hwpx: str,
    title: str = "",
    author: str = "",
    language: str = "ko",
) -> str:
    """Convert a Markdown file to HWPX format.

    Use this when the user wants to create an HWPX document from a Markdown file
    WITHOUT providing a reference HWPX template. Generates a complete document
    with cover page, page headers/footers, heading auto-numbering, native tables,
    code blocks, lists, and blockquotes.

    Args:
        input_md: Absolute path to the input .md file.
        output_hwpx: Absolute path for the output .hwpx file.
        title: Document title (shown on cover page and page header). Defaults to filename.
        author: Document author (shown on cover page).
        language: Document language code (default: "ko").

    Returns:
        Success message with output file path and size.
    """
    from md2hwpx import convert

    if not os.path.isabs(input_md):
        return f"Error: input_md must be an absolute path. Got: {input_md}"
    if not os.path.isabs(output_hwpx):
        return f"Error: output_hwpx must be an absolute path. Got: {output_hwpx}"
    if not os.path.exists(input_md):
        return f"Error: input file not found: {input_md}"

    try:
        convert(
            input_md=input_md,
            output_hwpx=output_hwpx,
            title=title or None,
            author=author or "Claude",
            language=language,
        )
        size = os.path.getsize(output_hwpx)
        size_kb = size / 1024
        return f"Success! HWPX saved: {output_hwpx} ({size_kb:.1f} KB)"
    except Exception as e:
        return f"Error during conversion: {e}"


# ═══════════════════════════════════════════════════════════════════
# Workflow 2: HWPX Template → HWPX (reference-based)
# ═══════════════════════════════════════════════════════════════════

@mcp.tool()
def analyze_hwpx(hwpx_path: str) -> str:
    """Analyze an existing HWPX file and return its full structure.

    Outputs font definitions, border/fill styles, character styles (charPr),
    paragraph styles (paraPr), page/margin settings, and the complete body
    structure including all paragraphs and tables with cell details.

    Use this as the first step when the user provides a reference HWPX file
    for replication or editing.

    Args:
        hwpx_path: Absolute path to the .hwpx file to analyze.

    Returns:
        Detailed structural analysis as plain text.
    """
    script = str(SCRIPTS_DIR / "analyze_template.py")
    stdout, stderr, rc = _run([_python(), script, hwpx_path])
    if rc != 0:
        return f"ERROR:\n{stderr or stdout}"
    return stdout


@mcp.tool()
def build_hwpx(
    output_hwpx: str,
    template: str = "",
    section_xml: str = "",
    header_xml: str = "",
    title: str = "",
    creator: str = "",
) -> str:
    """Build an HWPX file from a template and optional XML overrides.

    Workflow:
    1. Start from the base template skeleton.
    2. Optionally overlay a document-type template (gonmun/report/minutes).
    3. Optionally override header.xml (styles) and/or section0.xml (body).
    4. Pack as a valid HWPX archive and validate.

    When a user provides a reference HWPX, first call analyze_hwpx() and
    extract_hwpx_xml() to get the header/section XML, then write a new
    section0.xml with only the required text changes, and pass it here.

    Args:
        output_hwpx: Absolute path for the output .hwpx file.
        template: Document type template — one of: gonmun, report, minutes.
                  Leave empty for plain base template.
        section_xml: Absolute path to a custom section0.xml file (body content).
                     Leave empty to use the template's default section.
        header_xml: Absolute path to a custom header.xml file (styles).
                    Leave empty to use the template's default header.
        title: Document title written into content.hpf metadata.
        creator: Document creator written into content.hpf metadata.

    Returns:
        Success message with output path, or error details.
    """
    script = str(SCRIPTS_DIR / "build_hwpx.py")
    args = [_python(), script, "--output", output_hwpx]

    if template:
        args += ["--template", template]
    if section_xml:
        args += ["--section", section_xml]
    if header_xml:
        args += ["--header", header_xml]
    if title:
        args += ["--title", title]
    if creator:
        args += ["--creator", creator]

    stdout, stderr, rc = _run(args)
    if rc != 0:
        return f"ERROR:\n{stderr or stdout}"
    return stdout.strip()


@mcp.tool()
def extract_hwpx_xml(hwpx_path: str, extract_dir: str = "") -> str:
    """Extract header.xml and section0.xml from an HWPX file.

    Use this to get the raw XML from a reference HWPX so you can:
    - Read the exact charPrIDRef/paraPrIDRef/borderFillIDRef values in use
    - Copy the header.xml for reuse in build_hwpx()
    - Edit section0.xml text nodes and pass the modified file to build_hwpx()

    Args:
        hwpx_path: Absolute path to the source .hwpx file.
        extract_dir: Directory to write header.xml and section0.xml into.
                     Defaults to a system temp directory.

    Returns:
        Paths to the extracted header.xml and section0.xml files.
    """
    hwpx_path = str(Path(hwpx_path).resolve())
    if not extract_dir:
        extract_dir = tempfile.mkdtemp(prefix="hwpx_extract_")

    extract_dir = str(Path(extract_dir).resolve())
    os.makedirs(extract_dir, exist_ok=True)

    header_out = str(Path(extract_dir) / "header.xml")
    section_out = str(Path(extract_dir) / "section0.xml")

    script = str(SCRIPTS_DIR / "analyze_template.py")
    args = [
        _python(), script, hwpx_path,
        "--extract-header", header_out,
        "--extract-section", section_out,
    ]
    stdout, stderr, rc = _run(args)
    if rc != 0:
        return f"ERROR:\n{stderr or stdout}"

    return (
        f"Extracted successfully:\n"
        f"  header.xml  → {header_out}\n"
        f"  section0.xml → {section_out}\n\n"
        f"Edit section0.xml with Python/lxml, then pass both paths to build_hwpx().\n\n"
        f"IMPORTANT editing rules:\n"
        f"  - Element order = render order. Keep [header_table] → [body] → [next_header] → [body].\n"
        f"  - To remove items: delete entire <hp:p> from parent, don't just clear <hp:t> text.\n"
        f"  - When reducing TOC items: set vertAlign='TOP' on the TOC cell's <hp:subList>.\n"
        f"  - After build, use extract_text_hwpx to verify content and order."
    )


@mcp.tool()
def validate_hwpx(hwpx_path: str) -> str:
    """Validate the structural integrity of an HWPX file.

    Checks: valid ZIP, required files, mimetype content/position/compression,
    and XML well-formedness.

    Args:
        hwpx_path: Absolute path to the .hwpx file to validate.

    Returns:
        "VALID" with details, or list of errors found.
    """
    script = str(SCRIPTS_DIR / "validate.py")
    stdout, stderr, rc = _run([_python(), script, hwpx_path])
    if rc != 0:
        return f"INVALID:\n{stderr or stdout}"
    return stdout.strip()


@mcp.tool()
def page_guard_hwpx(
    reference_hwpx: str,
    output_hwpx: str,
    max_paragraph_delta: float = 0.25,
    max_text_delta: float = 0.15,
) -> str:
    """Check that output HWPX has not drifted in page count vs a reference.

    Best for SAME-STRUCTURE edits (filling values into a form, replacing text
    in-place without adding/removing paragraphs or tables). When sections are
    added, removed, or restructured, expected differences will be reported as
    failures — in that case use validate_hwpx for integrity checks instead.

    Compares paragraph count, table count, table dimensions, explicit
    page/column breaks, and overall text length between reference and output.

    Args:
        reference_hwpx: Absolute path to the original reference .hwpx file.
        output_hwpx: Absolute path to the newly built .hwpx file.
        max_paragraph_delta: Allowed text-length change per paragraph (0.25 = 25%).
        max_text_delta: Allowed total text-length change (0.15 = 15%).

    Returns:
        "PASS" with summary, or list of drift issues found.
    """
    script = str(SCRIPTS_DIR / "page_guard.py")
    args = [
        _python(), script,
        "--reference", reference_hwpx,
        "--output", output_hwpx,
        "--max-text-delta-ratio", str(max_text_delta),
        "--max-paragraph-delta-ratio", str(max_paragraph_delta),
    ]
    stdout, stderr, rc = _run(args)
    combined = (stdout + stderr).strip()
    return combined


@mcp.tool()
def extract_text_hwpx(
    hwpx_path: str,
    include_tables: bool = True,
    fmt: str = "text",
) -> str:
    """Extract plain text content from an HWPX file.

    Args:
        hwpx_path: Absolute path to the .hwpx file.
        include_tables: Whether to include table cell text (default: True).
        fmt: Output format — "text" (default) or "markdown".

    Returns:
        Extracted text content.
    """
    try:
        from text_extract import extract_plain, extract_markdown

        if fmt == "markdown":
            return extract_markdown(hwpx_path)
        return extract_plain(hwpx_path, include_tables=include_tables)
    except Exception as e:
        return f"ERROR: {e}"


if __name__ == "__main__":
    mcp.run()
