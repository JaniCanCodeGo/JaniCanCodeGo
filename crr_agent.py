#!/usr/bin/env python3
"""
CRR Civil Rights Review Document Processing Agent

Reads a completed CRR Summary of Findings (LOF), extracts Required Corrective
Actions, fills the Voluntary Compliance Plan (VCP), and generates the appropriate
LOF Cover Letter (Findings or No Findings).

Setup:
  pip install python-docx requests beautifulsoup4

  Place the three blank template .docx files in a templates/ folder:
    templates/VCP_template.docx
    templates/LOF_Findings_template.docx
    templates/LOF_NoFindings_template.docx

Usage:
  python3 crr_agent.py path/to/Summary_of_Findings.docx [--output-dir outputs/]
"""

import argparse
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

from docx import Document

SCRIPT_DIR = Path(__file__).parent
TEMPLATE_DIR = SCRIPT_DIR / "templates"

# California county codes (first 2 digits of CDS code)
CA_COUNTY_CODES = {
    "01": "Alameda",       "02": "Alpine",        "03": "Amador",
    "04": "Butte",         "05": "Calaveras",     "06": "Colusa",
    "07": "Contra Costa",  "08": "Del Norte",     "09": "El Dorado",
    "10": "Fresno",        "11": "Glenn",         "12": "Humboldt",
    "13": "Imperial",      "14": "Inyo",          "15": "Kern",
    "16": "Kings",         "17": "Lake",          "18": "Lassen",
    "19": "Los Angeles",   "20": "Madera",        "21": "Marin",
    "22": "Mariposa",      "23": "Mendocino",     "24": "Merced",
    "25": "Modoc",         "26": "Mono",          "27": "Monterey",
    "28": "Napa",          "29": "Nevada",        "30": "Orange",
    "31": "Placer",        "32": "Plumas",        "33": "Riverside",
    "34": "Sacramento",    "35": "San Benito",    "36": "San Bernardino",
    "37": "San Diego",     "38": "San Francisco", "39": "San Joaquin",
    "40": "San Luis Obispo", "41": "San Mateo",   "42": "Santa Barbara",
    "43": "Santa Clara",   "44": "Santa Cruz",    "45": "Shasta",
    "46": "Sierra",        "47": "Siskiyou",      "48": "Solano",
    "49": "Sonoma",        "50": "Stanislaus",    "51": "Sutter",
    "52": "Tehama",        "53": "Trinity",       "54": "Tulare",
    "55": "Tuolumne",      "56": "Ventura",       "57": "Yolo",
    "58": "Yuba",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def is_none_action(text: str) -> bool:
    """Return True if the text is a placeholder meaning 'no action required'."""
    t = text.strip().lower().rstrip(".")
    return t in ("none", "", "n/a", " " * 5)


def county_from_cds(cds_code: str) -> str:
    """Return the county name from the first 2 digits of a CDS code."""
    cds_clean = re.sub(r"[\s\-]", "", cds_code)
    return CA_COUNTY_CODES.get(cds_clean[:2], "")


def parse_principal(coordinator_str: str):
    """
    Parse 'Eric Preston, Principal' into (full_name, last_name, title).
    Handles 'First Last, Title' or just 'First Last'.
    """
    parts = coordinator_str.split(",", 1)
    full_name = parts[0].strip()
    title = parts[1].strip() if len(parts) > 1 else ""
    name_parts = full_name.split()
    last_name = name_parts[-1] if name_parts else ""
    return full_name, last_name, title


def normalize_cds(cds_code: str) -> str:
    """Normalize CDS code to 14 contiguous digits."""
    digits = re.sub(r"[\s\-]", "", cds_code)
    return digits.ljust(14, "0")[:14]


# ---------------------------------------------------------------------------
# CDE School Directory web lookup
# ---------------------------------------------------------------------------

def lookup_school_cde(cds_code: str) -> dict:
    """
    Look up school address and contact info from the CDE School Directory.
    Returns a dict with whatever fields were found; empty dict on failure.
    """
    import urllib.request

    cds_clean = normalize_cds(cds_code)
    url = f"https://www.cde.ca.gov/schooldirectory/details?cdscode={cds_clean}"
    info = {}

    try:
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                ),
                "Accept": "text/html,application/xhtml+xml",
            },
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
    except Exception as e:
        print(f"  [Info] CDE lookup failed ({type(e).__name__}): address/email will be blank.")
        return info

    if not html or "Host not in allowlist" in html or len(html) < 200:
        print("  [Info] CDE website not reachable from this machine; address/email will be blank.")
        return info

    # --- Parse with BeautifulSoup if available, else use regex ---
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "html.parser")

        def _val(label_lower):
            for row in soup.find_all("tr"):
                cells = row.find_all(["th", "td"])
                if len(cells) >= 2:
                    lbl = cells[0].get_text(" ", strip=True).lower()
                    val = cells[1].get_text(" ", strip=True)
                    if label_lower in lbl and val and val.lower() not in ("n/a", "none", ""):
                        return val
            for dt in soup.find_all("dt"):
                lbl = dt.get_text(strip=True).lower()
                dd = dt.find_next_sibling("dd")
                if dd and label_lower in lbl:
                    val = dd.get_text(strip=True)
                    if val and val.lower() not in ("n/a", "none", ""):
                        return val
            return ""

        info["street"]  = _val("street") or _val("address")
        info["city"]    = _val("city")
        info["state"]   = _val("state") or "CA"
        info["zip"]     = _val("zip")
        info["phone"]   = _val("phone")
        info["email"]   = _val("email")
        info["principal_name"] = _val("administrator") or _val("principal")

    except ImportError:
        # Regex fallback when bs4 not installed
        emails = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", html)
        school_emails = [e for e in emails if "cde.ca.gov" not in e]
        if school_emails:
            info["email"] = school_emails[0]

        phones = re.findall(r"\(?\d{3}\)?[\s\-\.]\d{3}[\s\-\.]\d{4}", html)
        if phones:
            info["phone"] = phones[0]

    # Remove empty strings
    info = {k: v for k, v in info.items() if v and v.strip()}
    if info:
        print(f"  [CDE] Retrieved: {', '.join(info.keys())}")
    return info


# ---------------------------------------------------------------------------
# Parsing the Summary of Findings
# ---------------------------------------------------------------------------

def parse_summary_of_findings(path: str):
    """
    Parse a completed CRR Summary of Findings document.

    Returns:
        metadata (dict) – school_name, cds_code, review_dates, coordinator,
                          reviewer, district, county, principal_name,
                          principal_last_name, principal_title, cde_info
        crr_sections (list of dicts)
    """
    from docx.text.paragraph import Paragraph as DocxParagraph
    from docx.table import Table as DocxTable

    doc = Document(path)

    # --- Body header fields ---
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

    # --- District from page header paragraphs ---
    district = ""
    for section in doc.sections:
        for p in section.header.paragraphs:
            t = p.text.strip()
            if t and t != metadata.get("school_name", "") and any(
                kw in t for kw in ["District", "Unified", "Elementary", "High", "Charter"]
            ):
                district = t
                break
        if district:
            break
    metadata["district"] = district

    # --- County from CDS code ---
    metadata["county"] = county_from_cds(metadata.get("cds_code", ""))

    # --- Principal from coordinator field ---
    coord = metadata.get("coordinator", "")
    if coord:
        principal_name, principal_last, principal_title = parse_principal(coord)
    else:
        principal_name = principal_last = principal_title = ""
    metadata["principal_name"]       = principal_name
    metadata["principal_last_name"]  = principal_last
    metadata["principal_title"]      = principal_title

    # --- CDE web lookup for address / email ---
    print("  Looking up school info from CDE School Directory...")
    metadata["cde_info"] = lookup_school_cde(metadata.get("cds_code", ""))

    # --- Walk body elements for CRR sections ---
    crr_sections = []
    current_crr = None
    state = None

    for element in doc.element.body:
        tag = element.tag.split("}")[-1]

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
                    current_crr = {
                        "crr_num": crr_num,
                        "crr_title": m.group(2).strip(),
                        "corrective_actions": [],
                        "is_accessible_facilities": "accessible facilit" in text.lower(),
                    }
                    state = "heading"
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
                clean = text.replace(" ", "").strip()
                if clean:
                    current_crr["corrective_actions"].append(clean)

        elif tag == "tbl":
            if current_crr and current_crr.get("is_accessible_facilities"):
                table = DocxTable(element, doc)
                for row in table.rows[1:]:
                    cells = row.cells
                    area       = cells[0].text.strip() if len(cells) > 0 else ""
                    violation  = cells[2].text.strip() if len(cells) > 2 else ""
                    correction = cells[3].text.strip() if len(cells) > 3 else ""
                    if correction and not is_none_action(correction):
                        current_crr["corrective_actions"].append(
                            f"Area: {area}\n"
                            f"Violation: {violation}\n"
                            f"Required Correction: {correction}"
                        )

    if current_crr is not None:
        crr_sections.append(current_crr)

    return metadata, crr_sections


def get_findings(crr_sections: list) -> list:
    findings = []
    for crr in crr_sections:
        real = [a for a in crr["corrective_actions"] if not is_none_action(a)]
        if real:
            findings.append({**crr, "corrective_actions": real})
    return findings


# ---------------------------------------------------------------------------
# Text replacement utilities
# ---------------------------------------------------------------------------

def _replace_adjacent_runs(para, old: str, new: str, max_window: int = 6) -> bool:
    """
    Replace `old` text that may be split across up to max_window consecutive runs.
    Only modifies the minimum number of runs needed.
    """
    runs = para.runs
    n = len(runs)
    for size in range(1, min(max_window + 1, n + 1)):
        for i in range(n - size + 1):
            combined = "".join(r.text for r in runs[i:i + size])
            if old in combined:
                runs[i].text = combined.replace(old, new, 1)
                for j in range(i + 1, i + size):
                    runs[j].text = ""
                return True
    return False


def replace_in_doc(doc, replacements: dict):
    """Apply replacement dict across all paragraphs (body + headers + footers)."""
    all_para_sets = [doc.paragraphs]
    for section in doc.sections:
        for attr in ("header", "footer"):
            try:
                all_para_sets.append(getattr(section, attr).paragraphs)
            except Exception:
                pass

    for paras in all_para_sets:
        for para in paras:
            for old, new in replacements.items():
                # Loop to replace every occurrence, not just the first
                while _replace_adjacent_runs(para, old, new):
                    pass


# ---------------------------------------------------------------------------
# VCP generation
# ---------------------------------------------------------------------------

def _set_cell_content(cell, title: str, body_lines: list):
    for para in cell.paragraphs:
        para.clear()
    run = cell.paragraphs[0].add_run(title)
    run.bold = True
    for line in body_lines:
        cell.add_paragraph(line)


def fill_vcp(template_path: str, output_path: str, metadata: dict, findings: list, today: datetime):
    doc = Document(template_path)

    school_name  = metadata.get("school_name") or "[School Name]"
    review_dates = metadata.get("review_dates") or "[Date]"

    # 45-day deadline is from the LOF cover letter date (today), not review end date
    deadline_str = (today + timedelta(days=45)).strftime("%B %d, %Y")

    replace_in_doc(doc, {
        "[School Name]":        school_name,
        "[Date]":               review_dates,
        "[45 days after Date]": deadline_str,
        # Also handle split-run variants
        "School Name":          school_name,
        "45 days after Date":   deadline_str,
    })

    # Fill the table
    table = doc.tables[0]
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

    doc.save(output_path)
    print(f"  [OK] VCP saved: {output_path}")


# ---------------------------------------------------------------------------
# LOF Cover Letter generation
# ---------------------------------------------------------------------------

def _fill_address_block(doc, metadata: dict, today: datetime):
    """Fill in the address block, salutation, and cc list of the cover letter."""
    school_name        = metadata.get("school_name", "")
    principal_name     = metadata.get("principal_name", "")
    principal_last     = metadata.get("principal_last_name", "")
    principal_title    = metadata.get("principal_title", "Principal")
    coordinator_name   = metadata.get("principal_name", "")  # coordinator IS the principal
    district           = metadata.get("district", "")
    county             = metadata.get("county", "")
    cde                = metadata.get("cde_info", {})

    street  = cde.get("street", "")
    city    = cde.get("city", "")
    state   = cde.get("state", "CA")
    zip_    = cde.get("zip", "")
    email   = cde.get("email", "")

    city_state_zip = f"{city}, {state} {zip_}".strip(", ") if city else ""

    paras = doc.paragraphs

    # Para [4]: "Name, Principal"
    if principal_name:
        title_str = principal_title if principal_title else "Principal"
        _replace_adjacent_runs(paras[4], "Name, Principal", f"{principal_name}, {title_str}")

    # Para [5]: "School Site Name"
    if school_name:
        _replace_adjacent_runs(paras[5], "School Site Name", school_name)
        _replace_adjacent_runs(paras[5], "School Site Name ", school_name)

    # Para [6]: "Address"
    if street:
        for run in paras[6].runs:
            if run.text.strip() == "Address":
                run.text = street
                break

    # Para [7]: "City, State Zip Code" / email block / "[Last name]"
    p7 = paras[7]
    if city_state_zip:
        _replace_adjacent_runs(p7, "City, State Zip Code", city_state_zip)

    if email:
        # Replace the whole "Email Address < When finalizing..." instruction with the real email
        email_placeholder = (
            "Email Address < When finalizing letter, this should be a hyperlink (in blue) "
        )
        _replace_adjacent_runs(p7, email_placeholder, email, max_window=8)
    else:
        # Remove the editorial instruction, leave blank line
        _replace_adjacent_runs(
            p7,
            "Email Address < When finalizing letter, this should be a hyperlink (in blue) ",
            "",
            max_window=8,
        )

    if principal_last:
        _replace_adjacent_runs(p7, "[Last name]", principal_last)

    # Para [25]: cc list
    p25 = paras[25]
    if coordinator_name:
        _replace_adjacent_runs(p25, "[Name], Designated CRR Coordinator", f"{coordinator_name}, Designated CRR Coordinator")
    if school_name:
        _replace_adjacent_runs(p25, "[School Site]", school_name)
        _replace_adjacent_runs(p25, "[School Site] ", school_name + " ")
    if district:
        _replace_adjacent_runs(p25, "[School District]", district)
        _replace_adjacent_runs(p25, "[School District] ", district + " ")
    if county:
        _replace_adjacent_runs(p25, "[County]", county)
        _replace_adjacent_runs(p25, "[County] ", county + " ")


def fill_cover_letter(
    template_path: str,
    output_path: str,
    metadata: dict,
    has_findings: bool,
    today: datetime,
):
    doc = Document(template_path)

    school_name  = metadata.get("school_name") or "[School Site Name]"
    review_dates = metadata.get("review_dates") or "[dates]"
    deadline_str = (today + timedelta(days=45)).strftime("%B %d, %Y")
    today_str    = today.strftime("%B %d, %Y")

    # Global text replacements
    replace_in_doc(doc, {
        "[Date]":              today_str,
        "[School Site Name]":  school_name,
        "[School Site]":       school_name,
        "[dates]":             review_dates,
        "[45 calendar days]":  deadline_str,
    })

    # Address block, salutation, cc list
    _fill_address_block(doc, metadata, today)

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
    parser.add_argument("input", help="Path to the completed CRR Summary of Findings (.docx)")
    parser.add_argument("--output-dir", "-o", default="outputs",
                        help="Output directory (default: outputs/)")
    parser.add_argument("--vcp-template",
                        default=str(TEMPLATE_DIR / "VCP_template.docx"))
    parser.add_argument("--lof-findings-template",
                        default=str(TEMPLATE_DIR / "LOF_Findings_template.docx"))
    parser.add_argument("--lof-no-findings-template",
                        default=str(TEMPLATE_DIR / "LOF_NoFindings_template.docx"))

    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    today = datetime.now()

    print()
    print("=" * 62)
    print("  CRR Document Processing Agent")
    print("=" * 62)
    print(f"  Input : {input_path.name}")
    print()

    metadata, crr_sections = parse_summary_of_findings(str(input_path))
    findings    = get_findings(crr_sections)
    has_findings = bool(findings)

    print()
    print(f"  School       : {metadata.get('school_name') or '(not found)'}")
    print(f"  District     : {metadata.get('district') or '(not found)'}")
    print(f"  County       : {metadata.get('county') or '(not found)'}")
    print(f"  Principal    : {metadata.get('principal_name') or '(not found)'}")
    print(f"  Review Dates : {metadata.get('review_dates') or '(not found)'}")
    print(f"  CRR Sections : {len(crr_sections)} parsed")
    print(f"  Findings     : {len(findings)} section(s) with corrective action(s)")

    if findings:
        print()
        print("  Sections with findings:")
        for f in findings:
            print(f"    * {f['crr_num']}: {f['crr_title']}")
    else:
        print("  -> No findings. Generating 'No Findings' cover letter.")

    print()
    print(f"  LOF Date     : {today.strftime('%B %d, %Y')}")
    print(f"  VCP Deadline : {(today + timedelta(days=45)).strftime('%B %d, %Y')} (45 days from LOF date)")
    print(f"  Cover Letter : {'LOF - Findings' if has_findings else 'LOF - No Findings'}")
    print()

    safe = re.sub(r"[^\w\s\-]", "", metadata.get("school_name", "School"))
    safe = safe.strip().replace(" ", "_")

    vcp_out = output_dir / f"{safe}_VCP.docx"
    lof_out = output_dir / f"{safe}_LOF_Cover_Letter.docx"

    fill_vcp(args.vcp_template, str(vcp_out), metadata, findings, today)

    lof_template = (
        args.lof_findings_template if has_findings else args.lof_no_findings_template
    )
    fill_cover_letter(lof_template, str(lof_out), metadata, has_findings, today)

    print()
    print("  Generated files:")
    print(f"    {vcp_out}")
    print(f"    {lof_out}")
    print("=" * 62)
    print()


if __name__ == "__main__":
    main()
