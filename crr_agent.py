#!/usr/bin/env python3
"""
CRR Civil Rights Review Document Processing Agent

Reads a completed CRR Summary of Findings (LOF), extracts Required Corrective
Actions, fills the Voluntary Compliance Plan (VCP), and generates the appropriate
LOF Cover Letter (Findings or No Findings).

Usage:
  python3 crr_agent.py path/to/Summary_of_Findings.docx [--output-dir outputs/]
"""

import argparse
import copy
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

SCRIPT_DIR = Path(__file__).parent
TEMPLATE_DIR = SCRIPT_DIR / "templates"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def is_none_action(text: str) -> bool:
    """Return True if the text is a placeholder meaning 'no action required'."""
    t = text.strip().lower().rstrip(".")
    return t in ("none", "", "n/a", " " * 5)  # en-space placeholders in blank template


def parse_review_end_date(review_dates_str: str):
    """
    Parse the last date from strings like:
      'October 14–18, 2024'  ->  datetime(2024, 10, 18)
      'October 14, 2024'     ->  datetime(2024, 10, 14)
    Returns None if not parseable.
    """
    months = {
        "january": 1, "february": 2, "march": 3, "april": 4,
        "may": 5, "june": 6, "july": 7, "august": 8,
        "september": 9, "october": 10, "november": 11, "december": 12,
    }
    s = review_dates_str.strip()

    # Range: "Month Day–Day, Year"  (– or -)
    m = re.search(r"(\w+)\s+\d+\s*[–\-]\s*(\d+),\s*(\d{4})", s, re.IGNORECASE)
    if m:
        month = months.get(m.group(1).lower())
        if month:
            return datetime(int(m.group(3)), month, int(m.group(2)))

    # Single day: "Month Day, Year"
    m = re.search(r"(\w+)\s+(\d+),\s*(\d{4})", s, re.IGNORECASE)
    if m:
        month = months.get(m.group(1).lower())
        if month:
            return datetime(int(m.group(3)), month, int(m.group(2)))

    return None


# ---------------------------------------------------------------------------
# Parsing the Summary of Findings
# ---------------------------------------------------------------------------

def parse_summary_of_findings(path: str):
    """
    Parse a completed CRR Summary of Findings document.

    Returns:
        metadata  (dict)  – school_name, cds_code, review_dates, coordinator, reviewer
        crr_sections (list of dicts) – crr_num, crr_title, corrective_actions (list of str)
    """
    from docx.text.paragraph import Paragraph as DocxParagraph
    from docx.table import Table as DocxTable

    doc = Document(path)

    # --- Metadata from the header block ---
    metadata = {
        "school_name": "",
        "cds_code": "",
        "review_dates": "",
        "coordinator": "",
        "reviewer": "",
    }
    for para in doc.paragraphs[:15]:
        text = para.text.strip()
        for key, prefix in [
            ("school_name",  "School Site:"),
            ("cds_code",     "CDS Code:"),
            ("review_dates", "Review Dates:"),
            ("coordinator",  "Site CRR Coordinator:"),
            ("reviewer",     "Program Reviewer:"),
        ]:
            if text.startswith(prefix):
                metadata[key] = text.split(":", 1)[1].strip()

    # --- Walk body elements in document order (paragraphs AND tables) ---
    crr_sections = []
    current_crr = None
    state = None  # 'heading' | 'summary' | 'corrective_actions' | 'observation'

    for element in doc.element.body:
        tag = element.tag.split("}")[-1]  # 'p' or 'tbl'

        if tag == "p":
            para = DocxParagraph(element, doc)
            text = para.text.strip()
            style = para.style.name if para.style else ""

            if style == "Heading 1":
                m = re.match(r"(CRR\s+\d+):\s*(.+)", text)
                if m:
                    if current_crr is not None:
                        crr_sections.append(current_crr)
                    crr_num = re.sub(r"\s+", " ", m.group(1)).strip()
                    crr_title = m.group(2).strip()
                    current_crr = {
                        "crr_num": crr_num,
                        "crr_title": crr_title,
                        "corrective_actions": [],
                        "is_accessible_facilities": "accessible facilit" in text.lower(),
                    }
                    state = "heading"
                # "CRR 13: REMOVED" – flush current and skip
                elif current_crr is not None and "REMOVED" in text:
                    crr_sections.append(current_crr)
                    current_crr = None
                    state = None

            elif style == "Heading 2":
                if "Required Corrective Action" in text:
                    state = "corrective_actions"
                elif "Summary of" in text or "Analysis" in text:
                    state = "summary"
                elif "Observation" in text:
                    state = "observation"
                else:
                    state = "other"

            elif state == "corrective_actions" and current_crr:
                # Strip en-space placeholder chars used in blank template
                clean = text.replace(" ", "").strip()
                if clean:
                    current_crr["corrective_actions"].append(clean)

        elif tag == "tbl":
            # Accessible Facilities section uses a table for corrective actions
            if current_crr and current_crr.get("is_accessible_facilities"):
                table = DocxTable(element, doc)
                items = []
                for row in table.rows[1:]:  # skip header
                    cells = row.cells
                    area       = cells[0].text.strip() if len(cells) > 0 else ""
                    violation  = cells[2].text.strip() if len(cells) > 2 else ""
                    correction = cells[3].text.strip() if len(cells) > 3 else ""
                    if correction and not is_none_action(correction):
                        items.append(
                            f"Area: {area}\n"
                            f"Violation: {violation}\n"
                            f"Required Correction: {correction}"
                        )
                current_crr["corrective_actions"].extend(items)

    if current_crr is not None:
        crr_sections.append(current_crr)

    return metadata, crr_sections


def get_findings(crr_sections: list) -> list:
    """Return only the CRR sections that have real (non-None) corrective actions."""
    findings = []
    for crr in crr_sections:
        real_actions = [a for a in crr["corrective_actions"] if not is_none_action(a)]
        if real_actions:
            findings.append({**crr, "corrective_actions": real_actions})
    return findings


# ---------------------------------------------------------------------------
# Text replacement utilities
# ---------------------------------------------------------------------------

def _replace_in_para(para, old: str, new: str) -> bool:
    """Replace old with new inside a paragraph, handling text split across runs."""
    # Fast path: placeholder lives entirely within one run
    for run in para.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            return True

    # Slow path: placeholder is split across consecutive runs
    combined = "".join(r.text for r in para.runs)
    if old in combined:
        para.runs[0].text = combined.replace(old, new)
        for run in para.runs[1:]:
            run.text = ""
        return True

    return False


def _replace_run_triplet(para, bracket_open: str, inner: str, replacement: str):
    """
    Handle placeholders like '[', 'School Name', ']' stored in three separate runs.
    Collapses them into the middle run.
    """
    runs = para.runs
    for i in range(len(runs) - 2):
        if (
            runs[i].text.strip() == "["
            and runs[i + 1].text.strip() == inner
            and runs[i + 2].text.strip() == "]"
        ):
            runs[i].text = ""
            runs[i + 1].text = replacement
            runs[i + 2].text = ""
            return True
    return False


def replace_in_doc(doc, replacements: dict):
    """Apply replacement dict to all paragraphs (body + headers + footers)."""
    all_para_sets = [doc.paragraphs]
    for section in doc.sections:
        try:
            all_para_sets.append(section.header.paragraphs)
        except Exception:
            pass
        try:
            all_para_sets.append(section.footer.paragraphs)
        except Exception:
            pass

    for paras in all_para_sets:
        for para in paras:
            for old, new in replacements.items():
                _replace_in_para(para, old, new)


def replace_triplets_in_doc(doc, triplets: dict):
    """
    Apply triplet-style replacements {inner_text: replacement} to all paragraphs.
    Handles the '[', 'placeholder', ']' pattern split across three runs.
    """
    all_para_sets = [doc.paragraphs]
    for section in doc.sections:
        try:
            all_para_sets.append(section.header.paragraphs)
        except Exception:
            pass

    for paras in all_para_sets:
        for para in paras:
            for inner, replacement in triplets.items():
                _replace_run_triplet(para, "[", inner, replacement)


# ---------------------------------------------------------------------------
# VCP generation
# ---------------------------------------------------------------------------

def _set_cell_content(cell, title: str, body_lines: list):
    """
    Write bold title + body paragraphs into a table cell.
    Clears existing content first.
    """
    # Clear existing paragraphs (keep at least one)
    for para in cell.paragraphs:
        para.clear()

    # First paragraph: bold CRR title
    first_para = cell.paragraphs[0]
    run = first_para.add_run(title)
    run.bold = True

    # Subsequent paragraphs: corrective action text
    for line in body_lines:
        cell.add_paragraph(line)


def fill_vcp(template_path: str, output_path: str, metadata: dict, findings: list):
    """Fill the VCP template and save to output_path."""
    doc = Document(template_path)

    school_name  = metadata.get("school_name", "") or "[School Name]"
    review_dates = metadata.get("review_dates", "") or "[Date]"

    end_date = parse_review_end_date(review_dates)
    if end_date:
        deadline_str = (end_date + timedelta(days=45)).strftime("%B %d, %Y")
    else:
        deadline_str = "45 days after the date of the review"

    # Simple in-run replacements
    replace_in_doc(doc, {
        "[School Name]": school_name,
        "[Date]":        review_dates,
        "[45 days after Date]": deadline_str,
    })

    # Triplet replacements (text split across [ / text / ] runs in VCP para 1 & 4)
    replace_triplets_in_doc(doc, {
        "School Name":       school_name,
        "Date":              review_dates,
        "45 days after Date": deadline_str,
    })

    # --- Fill the VCP table ---
    table = doc.tables[0]

    # Remove pre-existing blank placeholder rows (rows 1 onward)
    blank_trs = [
        row._tr for row in table.rows[1:]
        if not any(cell.text.strip() for cell in row.cells)
    ]
    for tr in blank_trs:
        tr.getparent().remove(tr)

    if not findings:
        new_row = table.add_row()
        new_row.cells[0].text = "N/A"
        new_row.cells[1].text = "No findings of noncompliance were identified."
        for i in range(2, min(6, len(new_row.cells))):
            new_row.cells[i].text = "N/A"
    else:
        for finding in findings:
            new_row = table.add_row()
            new_row.cells[0].text = finding["crr_num"]
            _set_cell_content(
                new_row.cells[1],
                title=f"{finding['crr_num']}: {finding['crr_title']}",
                body_lines=finding["corrective_actions"],
            )
            # Columns 2–5 left blank for the school to complete

    doc.save(output_path)
    print(f"  [OK] VCP saved: {output_path}")


# ---------------------------------------------------------------------------
# LOF Cover Letter generation
# ---------------------------------------------------------------------------

def fill_cover_letter(template_path: str, output_path: str, metadata: dict, has_findings: bool):
    """Fill the LOF cover letter template and save to output_path."""
    doc = Document(template_path)

    school_name  = metadata.get("school_name", "") or "[School Site Name]"
    review_dates = metadata.get("review_dates", "") or "[dates]"

    end_date = parse_review_end_date(review_dates)
    if end_date:
        deadline_str = (end_date + timedelta(days=45)).strftime("%B %d, %Y")
    else:
        deadline_str = "45 calendar days from the date of this letter"

    today = datetime.now().strftime("%B %d, %Y")

    replace_in_doc(doc, {
        "[Date]":            today,
        "[School Site Name]": school_name,
        "[School Site]":     school_name,
        "[dates]":           review_dates,
        "[45 calendar days]": deadline_str,
    })

    doc.save(output_path)
    letter_type = "Findings" if has_findings else "No Findings"
    print(f"  [OK] LOF Cover Letter ({letter_type}) saved: {output_path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="CRR Civil Rights Review Document Processing Agent",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "input",
        help="Path to the completed CRR Summary of Findings (.docx)",
    )
    parser.add_argument(
        "--output-dir", "-o",
        default="outputs",
        help="Directory where generated files will be saved (default: outputs/)",
    )
    parser.add_argument(
        "--vcp-template",
        default=str(TEMPLATE_DIR / "VCP_template.docx"),
        help="Path to VCP blank template",
    )
    parser.add_argument(
        "--lof-findings-template",
        default=str(TEMPLATE_DIR / "LOF_Findings_template.docx"),
        help="Path to LOF Cover Letter – Findings template",
    )
    parser.add_argument(
        "--lof-no-findings-template",
        default=str(TEMPLATE_DIR / "LOF_NoFindings_template.docx"),
        help="Path to LOF Cover Letter – No Findings template",
    )

    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    print()
    print("=" * 62)
    print("  CRR Document Processing Agent")
    print("=" * 62)
    print(f"  Input : {input_path.name}")
    print()

    # --- Parse ---
    metadata, crr_sections = parse_summary_of_findings(str(input_path))
    findings = get_findings(crr_sections)
    has_findings = bool(findings)

    print(f"  School       : {metadata.get('school_name') or '(not found)'}")
    print(f"  Review Dates : {metadata.get('review_dates') or '(not found)'}")
    print(f"  CRR Sections : {len(crr_sections)} parsed")
    print(f"  Findings     : {len(findings)} section(s) with required corrective action(s)")

    if findings:
        print()
        print("  Sections with findings:")
        for f in findings:
            print(f"    • {f['crr_num']}: {f['crr_title']}")
    else:
        print("  → No findings detected. Generating 'No Findings' cover letter.")

    print()
    print(f"  Cover Letter : {'LOF – Findings' if has_findings else 'LOF – No Findings'}")
    print()

    # --- Output file names ---
    safe = re.sub(r"[^\w\s\-]", "", metadata.get("school_name", "School"))
    safe = safe.strip().replace(" ", "_")

    vcp_out = output_dir / f"{safe}_VCP.docx"
    lof_out = output_dir / f"{safe}_LOF_Cover_Letter.docx"

    # --- Generate VCP ---
    fill_vcp(args.vcp_template, str(vcp_out), metadata, findings)

    # --- Generate Cover Letter ---
    lof_template = (
        args.lof_findings_template if has_findings else args.lof_no_findings_template
    )
    fill_cover_letter(lof_template, str(lof_out), metadata, has_findings)

    print()
    print("  Generated files:")
    print(f"    {vcp_out}")
    print(f"    {lof_out}")
    print("=" * 62)
    print()


if __name__ == "__main__":
    main()
