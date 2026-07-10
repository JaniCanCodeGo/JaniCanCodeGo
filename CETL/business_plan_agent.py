#!/usr/bin/env python3
"""CETL Business Plan Agent.

Converts BUSINESS_PLAN.md into a formatted, downloadable Word document.

Usage:
    python3 business_plan_agent.py
    python3 business_plan_agent.py --source BUSINESS_PLAN.md --output-dir outputs

The generated .docx lands in outputs/ by default (gitignored), named
CETL_Business_Plan.docx. Re-run any time the markdown changes — the
markdown file stays the source of truth; the .docx is a build artifact.

Requires: python-docx  (pip install python-docx)
"""

import argparse
import os
import re
import sys

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt, RGBColor
except ImportError:
    print("[Error] python-docx is required: pip install python-docx")
    sys.exit(1)

ACCENT = RGBColor(0x5B, 0x2D, 0x86)  # deep plum — warm, not corporate-gray
BODY_FONT = "Calibri"


def _add_runs_with_inline_formatting(paragraph, text):
    """Render a markdown text fragment into runs: **bold**, *italic*,
    [text](url) links (as 'text (url)'), [^n] footnote refs (superscript),
    and `code` (monospace)."""
    # normalize links first so their brackets don't confuse other patterns
    text = re.sub(r"\[([^\]^][^\]]*)\]\((https?://[^)]+)\)", r"\1 (\2)", text)
    token_re = re.compile(r"(\*\*.+?\*\*|\*[^*]+?\*|`[^`]+`|\[\^\d+\])")
    for part in token_re.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            paragraph.add_run(part[2:-2]).bold = True
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            paragraph.add_run(part[1:-1]).italic = True
        elif part.startswith("`") and part.endswith("`"):
            run = paragraph.add_run(part[1:-1])
            run.font.name = "Consolas"
        elif re.fullmatch(r"\[\^\d+\]", part):
            run = paragraph.add_run(part[2:-1])
            run.font.superscript = True
        else:
            paragraph.add_run(part)


def _add_table(doc, rows):
    """rows: list of lists of cell strings; first row is the header."""
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = "Light Grid Accent 1"
    for r, row in enumerate(rows):
        for c, cell_text in enumerate(row):
            cell = table.cell(r, c)
            cell.paragraphs[0].text = ""
            _add_runs_with_inline_formatting(cell.paragraphs[0], cell_text)
            if r == 0:
                for run in cell.paragraphs[0].runs:
                    run.bold = True
    doc.add_paragraph()


def _style_document(doc):
    style = doc.styles["Normal"]
    style.font.name = BODY_FONT
    style.font.size = Pt(11)
    for level, size in (("Heading 1", 16), ("Heading 2", 13), ("Heading 3", 12)):
        h = doc.styles[level]
        h.font.name = BODY_FONT
        h.font.size = Pt(size)
        h.font.color.rgb = ACCENT
    for section in doc.sections:
        section.left_margin = section.right_margin = Inches(1)


def _split_table_row(line):
    return [c.strip() for c in line.strip().strip("|").split("|")]


def build_docx(source_path, output_path):
    with open(source_path, encoding="utf-8") as f:
        lines = f.read().splitlines()

    doc = Document()
    _style_document(doc)

    footnotes = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        if stripped == "---":
            i += 1
            continue

        # footnote definitions are collected and rendered as an endnotes section
        fn = re.match(r"\[\^(\d+)\]:\s*(.*)", stripped)
        if fn:
            footnotes.append((fn.group(1), fn.group(2)))
            i += 1
            continue

        heading = re.match(r"(#{1,4})\s+(.*)", stripped)
        if heading:
            if heading.group(2).strip() == "Footnotes":
                # rendered as its own endnotes section after the body
                i += 1
                continue
            level = len(heading.group(1))
            text = heading.group(2)
            if level == 1:
                p = doc.add_heading("", level=0)
                _add_runs_with_inline_formatting(p, text)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p = doc.add_heading("", level=min(level - 1, 3))
                _add_runs_with_inline_formatting(p, text)
            i += 1
            continue

        if stripped.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = _split_table_row(lines[i])
                if not all(re.fullmatch(r":?-{3,}:?", c) for c in cells):
                    rows.append(cells)
                i += 1
            if rows:
                width = max(len(r) for r in rows)
                rows = [r + [""] * (width - len(r)) for r in rows]
                _add_table(doc, rows)
            continue

        if stripped.startswith(">"):
            quote = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                quote.append(lines[i].strip().lstrip(">").strip())
                i += 1
            p = doc.add_paragraph(style="Intense Quote")
            _add_runs_with_inline_formatting(p, " ".join(q for q in quote if q))
            continue

        bullet = re.match(r"[-*]\s+(.*)", stripped)
        if bullet:
            p = doc.add_paragraph(style="List Bullet")
            _add_runs_with_inline_formatting(p, bullet.group(1))
            i += 1
            continue

        numbered = re.match(r"\d+\.\s+(.*)", stripped)
        if numbered:
            p = doc.add_paragraph(style="List Number")
            _add_runs_with_inline_formatting(p, numbered.group(1))
            i += 1
            continue

        p = doc.add_paragraph()
        _add_runs_with_inline_formatting(p, stripped)
        i += 1

    if footnotes:
        doc.add_page_break()
        doc.add_heading("Footnotes", level=1)
        for num, text in footnotes:
            p = doc.add_paragraph()
            run = p.add_run(num)
            run.font.superscript = True
            p.add_run("  ")
            _add_runs_with_inline_formatting(p, text)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    return output_path


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    parser = argparse.ArgumentParser(description="Build the CETL business plan .docx from markdown.")
    parser.add_argument("--source", default=os.path.join(here, "BUSINESS_PLAN.md"))
    parser.add_argument("--output-dir", default=os.path.join(here, "outputs"))
    args = parser.parse_args()

    if not os.path.exists(args.source):
        print(f"[Error] Source not found: {args.source}")
        sys.exit(1)

    out = build_docx(args.source, os.path.join(args.output_dir, "CETL_Business_Plan.docx"))
    print(f"[OK] Business plan written to {out}")


if __name__ == "__main__":
    main()
