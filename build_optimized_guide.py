"""Build the optimized BLANK Facilities Review Guide.

Reads BLANK_Facilities__Review_Guide.docx, enhances each of the 18 area
sections with:
  - Date picker content control for "Date Constructed"
  - Date picker content control for "ADA Modification Date"
  - Dropdown content control for "Construction Era" (6 options, each
    labeled with the standard that applies)
  - Dropdown content control for "Modification Era" (same options + N/A)
  - A repeating-section content control wrapper so users can click "+"
    to add another instance (Restroom #2, CTE Lab #3, etc.)

No macros. No API keys. Works in Word 2013+, Word for Mac, and Word for
the web.
"""
import copy
import os
import re
from docx import Document
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from docx.shared import Pt, Inches, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from lxml import etree

# Word namespaces
W_NS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"

# Register the w15 namespace for content controls (repeating section)
nsmap['w15'] = W15_NS

SRC = "/root/.claude/uploads/f318c6f4-c9c8-4a12-bcd0-bfd812a66767/60dc1d22-BLANK_Facilities__Review_Guide.docx"
OUT = "outputs/BLANK_Facilities_Review_Guide_optimized.docx"

# The six standard windows — dropdown labels include the standard so the
# choice IS the standard determination (no macros / formulas needed).
ERA_DROPDOWN = [
    ("≤ June 3, 1977 — Existing Facility (Program Access)",       "Program Access"),
    ("June 4, 1977 – January 17, 1991 — ANSI A117.1 (1961 R1971)", "ANSI A117.1"),
    ("January 18, 1991 – January 26, 1992 — UFAS (1984)",          "UFAS"),
    ("January 27, 1992 – September 14, 2010 — 1991 ADA / ADAAG",   "1991 ADA"),
    ("September 15, 2010 – March 14, 2012 — 1991 ADA OR 2010 ADA", "1991 ADA or 2010 ADA"),
    ("March 15, 2012 or later — 2010 ADA Standards",               "2010 ADA"),
]
MOD_DROPDOWN = [("— Not modified —", "N/A")] + ERA_DROPDOWN

# Brand colors matching facilities_lof_tool.html
NAVY = "1A2744"
GOLD = "C8A84B"
CREAM = "F7F4EE"
BORDER = "D4CEBD"

# ───────────────────────────────────────────────────────────────────────
# XML builders
# ───────────────────────────────────────────────────────────────────────
def w(tag, **attrs):
    """Build a w:* element with optional attributes."""
    e = OxmlElement(f"w:{tag}")
    for k, v in attrs.items():
        e.set(qn(f"w:{k}"), str(v))
    return e

def w15(tag):
    """Build a w15:* element."""
    return etree.SubElement(etree.Element("dummy"), f"{{{W15_NS}}}{tag}")

def make_date_picker(label="Click to enter a date", default_text=None):
    """A native Word date-picker content control (sdt)."""
    sdt = OxmlElement("w:sdt")
    sdtPr = OxmlElement("w:sdtPr")
    # ID (required)
    rid = OxmlElement("w:id"); rid.set(qn("w:val"), str(abs(hash(label)) % 99999999))
    sdtPr.append(rid)
    # Placeholder
    ph = OxmlElement("w:placeholder")
    dp = OxmlElement("w:docPart"); dp.set(qn("w:val"), "DefaultPlaceholder_-1854013440")
    ph.append(dp); sdtPr.append(ph)
    # Show placeholder
    sdtPr.append(OxmlElement("w:showingPlcHdr"))
    # Date picker properties
    date = OxmlElement("w:date")
    fmt = OxmlElement("w:dateFormat"); fmt.set(qn("w:val"), "MMMM d, yyyy"); date.append(fmt)
    lid = OxmlElement("w:lid"); lid.set(qn("w:val"), "en-US"); date.append(lid)
    sct = OxmlElement("w:storeMappedDataAs"); sct.set(qn("w:val"), "dateTime"); date.append(sct)
    cal = OxmlElement("w:calendar"); cal.set(qn("w:val"), "gregorian"); date.append(cal)
    sdtPr.append(date)
    sdt.append(sdtPr)
    sdt.append(OxmlElement("w:sdtEndPr"))
    # Content (placeholder text)
    sdtContent = OxmlElement("w:sdtContent")
    p_run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    style = OxmlElement("w:rStyle"); style.set(qn("w:val"), "PlaceholderText"); rpr.append(style)
    p_run.append(rpr)
    t = OxmlElement("w:t"); t.text = default_text or label; p_run.append(t)
    sdtContent.append(p_run)
    sdt.append(sdtContent)
    return sdt

def make_dropdown(options, label="Click to pick"):
    """A native Word dropdown content control (sdt + dropDownList).
       options is a list of (display, value) tuples."""
    sdt = OxmlElement("w:sdt")
    sdtPr = OxmlElement("w:sdtPr")
    rid = OxmlElement("w:id"); rid.set(qn("w:val"), str(abs(hash(label)) % 99999999))
    sdtPr.append(rid)
    ph = OxmlElement("w:placeholder")
    dp = OxmlElement("w:docPart"); dp.set(qn("w:val"), "DefaultPlaceholder_-1854013440")
    ph.append(dp); sdtPr.append(ph)
    sdtPr.append(OxmlElement("w:showingPlcHdr"))
    ddl = OxmlElement("w:dropDownList")
    for disp, val in options:
        li = OxmlElement("w:listItem")
        li.set(qn("w:displayText"), disp)
        li.set(qn("w:value"), val)
        ddl.append(li)
    sdtPr.append(ddl)
    sdt.append(sdtPr)
    sdt.append(OxmlElement("w:sdtEndPr"))
    sdtContent = OxmlElement("w:sdtContent")
    p_run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    style = OxmlElement("w:rStyle"); style.set(qn("w:val"), "PlaceholderText"); rpr.append(style)
    p_run.append(rpr)
    t = OxmlElement("w:t"); t.text = label; p_run.append(t)
    sdtContent.append(p_run)
    sdt.append(sdtContent)
    return sdt

def make_plain_text_cc(label="Click to enter text"):
    sdt = OxmlElement("w:sdt")
    sdtPr = OxmlElement("w:sdtPr")
    rid = OxmlElement("w:id"); rid.set(qn("w:val"), str(abs(hash(label)) % 99999999))
    sdtPr.append(rid)
    ph = OxmlElement("w:placeholder")
    dp = OxmlElement("w:docPart"); dp.set(qn("w:val"), "DefaultPlaceholder_-1854013440")
    ph.append(dp); sdtPr.append(ph)
    sdtPr.append(OxmlElement("w:showingPlcHdr"))
    sdtPr.append(OxmlElement("w:text"))
    sdt.append(sdtPr)
    sdt.append(OxmlElement("w:sdtEndPr"))
    sdtContent = OxmlElement("w:sdtContent")
    p_run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    style = OxmlElement("w:rStyle"); style.set(qn("w:val"), "PlaceholderText"); rpr.append(style)
    p_run.append(rpr)
    t = OxmlElement("w:t"); t.text = label; p_run.append(t)
    sdtContent.append(p_run)
    sdt.append(sdtContent)
    return sdt

def set_cell_shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_cell_borders(cell, color=BORDER, size=4):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), str(size))
        b.set(qn("w:color"), color)
        tcBorders.append(b)
    tcPr.append(tcBorders)

print("Loading source guide…")
doc = Document(SRC)
print(f"  {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")

# ───────────────────────────────────────────────────────────────────────
# Find all Heading 1 paragraphs — these mark the 18 area sections.
# ───────────────────────────────────────────────────────────────────────
body = doc.element.body
heading_para_indexes = []
for i, p in enumerate(doc.paragraphs):
    sty = (p.style.name or "").strip()
    if sty == "Heading 1":
        heading_para_indexes.append(i)
        print(f"  H1 @ p[{i}]: {p.text[:80]}")
print(f"Found {len(heading_para_indexes)} Heading 1 sections")

# ───────────────────────────────────────────────────────────────────────
# For each Heading 1, insert an enhanced "Section Details" table directly
# after it. The new table holds:
#   - Location / Building name (plain text content control)
#   - Date Constructed (date picker content control)
#   - Construction Era / Standard at construction (dropdown)
#   - Date of ADA Modification (date picker content control)
#   - Modification Era / Standard after modification (dropdown)
#   - Applicable Standard guidance
# ───────────────────────────────────────────────────────────────────────

def build_section_details_table(area_name):
    """Build a fresh details table to inject after a Heading 1."""
    tbl = OxmlElement("w:tbl")
    # Table properties
    tblPr = OxmlElement("w:tblPr")
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:w"), "5000"); tblW.set(qn("w:type"), "pct")
    tblPr.append(tblW)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top","left","bottom","right","insideH","insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "4"); b.set(qn("w:color"), BORDER)
        tblBorders.append(b)
    tblPr.append(tblBorders)
    tblLook = OxmlElement("w:tblLook"); tblLook.set(qn("w:val"), "04A0")
    tblPr.append(tblLook)
    tbl.append(tblPr)
    # Grid
    tblGrid = OxmlElement("w:tblGrid")
    for w_ in ("3200", "6500"):
        gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), w_); tblGrid.append(gc)
    tbl.append(tblGrid)

    def row(label, content_element, shade_label=CREAM):
        tr = OxmlElement("w:tr")
        # Label cell
        tc1 = OxmlElement("w:tc")
        tcPr1 = OxmlElement("w:tcPr")
        w1 = OxmlElement("w:tcW"); w1.set(qn("w:w"), "3200"); w1.set(qn("w:type"), "dxa"); tcPr1.append(w1)
        shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), shade_label); tcPr1.append(shd)
        tc1.append(tcPr1)
        p1 = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        spc = OxmlElement("w:spacing"); spc.set(qn("w:after"), "0"); ppr.append(spc)
        p1.append(ppr)
        r1 = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        b = OxmlElement("w:b"); rpr.append(b)
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "20"); rpr.append(sz)
        col = OxmlElement("w:color"); col.set(qn("w:val"), NAVY); rpr.append(col)
        r1.append(rpr)
        t1 = OxmlElement("w:t"); t1.text = label; r1.append(t1)
        p1.append(r1)
        tc1.append(p1)
        tr.append(tc1)
        # Content cell
        tc2 = OxmlElement("w:tc")
        tcPr2 = OxmlElement("w:tcPr")
        w2 = OxmlElement("w:tcW"); w2.set(qn("w:w"), "6500"); w2.set(qn("w:type"), "dxa"); tcPr2.append(w2)
        tc2.append(tcPr2)
        p2 = OxmlElement("w:p")
        ppr2 = OxmlElement("w:pPr")
        spc2 = OxmlElement("w:spacing"); spc2.set(qn("w:after"), "0"); ppr2.append(spc2)
        p2.append(ppr2)
        p2.append(content_element)
        tc2.append(p2)
        tr.append(tc2)
        return tr

    # Title row (full-width)
    tr_hdr = OxmlElement("w:tr")
    tc_hdr = OxmlElement("w:tc")
    tcPr_h = OxmlElement("w:tcPr")
    w_h = OxmlElement("w:tcW"); w_h.set(qn("w:w"), "9700"); w_h.set(qn("w:type"), "dxa"); tcPr_h.append(w_h)
    gs = OxmlElement("w:gridSpan"); gs.set(qn("w:val"), "2"); tcPr_h.append(gs)
    shd_h = OxmlElement("w:shd"); shd_h.set(qn("w:val"), "clear"); shd_h.set(qn("w:color"), "auto"); shd_h.set(qn("w:fill"), NAVY); tcPr_h.append(shd_h)
    tc_hdr.append(tcPr_h)
    p_hdr = OxmlElement("w:p")
    ppr_h = OxmlElement("w:pPr")
    spc_h = OxmlElement("w:spacing"); spc_h.set(qn("w:after"), "60"); ppr_h.append(spc_h)
    p_hdr.append(ppr_h)
    r_hdr = OxmlElement("w:r")
    rpr_h = OxmlElement("w:rPr")
    bh = OxmlElement("w:b"); rpr_h.append(bh)
    szh = OxmlElement("w:sz"); szh.set(qn("w:val"), "22"); rpr_h.append(szh)
    colh = OxmlElement("w:color"); colh.set(qn("w:val"), "C8A84B"); rpr_h.append(colh)
    r_hdr.append(rpr_h)
    th = OxmlElement("w:t"); th.text = f"SECTION DETAILS — {area_name.upper()}"; r_hdr.append(th)
    p_hdr.append(r_hdr)
    tc_hdr.append(p_hdr)
    tr_hdr.append(tc_hdr)
    tbl.append(tr_hdr)

    # Rows
    tbl.append(row("Location / Building name:",   make_plain_text_cc("Click to enter location or building name")))
    tbl.append(row("Date constructed:",            make_date_picker("Click to pick date constructed")))
    tbl.append(row("Era — Standard at construction:", make_dropdown(ERA_DROPDOWN, "Click to choose era → standard")))
    tbl.append(row("Date of ADA modification:",    make_date_picker("Click to pick modification date, or skip")))
    tbl.append(row("Era — Standard after modification:", make_dropdown(MOD_DROPDOWN, "Click to choose era → standard (or N/A)")))
    tbl.append(row("Describe modification:",       make_plain_text_cc("Click to describe what was modified (e.g., restroom gutted and rebuilt)")))

    # Helper / guidance row
    tr_g = OxmlElement("w:tr")
    tc_g = OxmlElement("w:tc")
    tcPr_g = OxmlElement("w:tcPr")
    w_g = OxmlElement("w:tcW"); w_g.set(qn("w:w"), "9700"); w_g.set(qn("w:type"), "dxa"); tcPr_g.append(w_g)
    gs2 = OxmlElement("w:gridSpan"); gs2.set(qn("w:val"), "2"); tcPr_g.append(gs2)
    shd_g = OxmlElement("w:shd"); shd_g.set(qn("w:val"), "clear"); shd_g.set(qn("w:color"), "auto"); shd_g.set(qn("w:fill"), "FEF9C3"); tcPr_g.append(shd_g)
    tc_g.append(tcPr_g)
    p_g = OxmlElement("w:p")
    r_g = OxmlElement("w:r")
    rpr_g = OxmlElement("w:rPr")
    bg = OxmlElement("w:b"); rpr_g.append(bg)
    szg = OxmlElement("w:sz"); szg.set(qn("w:val"), "18"); rpr_g.append(szg)
    cg = OxmlElement("w:color"); cg.set(qn("w:val"), "713F12"); rpr_g.append(cg)
    r_g.append(rpr_g)
    tg = OxmlElement("w:t")
    tg.text = "▸  Applicable Standard = the LATER of the two eras above.  If \"Not modified,\" use the construction era.  Reviewer will verify."
    r_g.append(tg)
    p_g.append(r_g)
    tc_g.append(p_g)
    tr_g.append(tc_g)
    tbl.append(tr_g)
    return tbl

print("\nInjecting Section Details tables after each Heading 1…")

# Walk paragraphs in document order; for each Heading 1 paragraph,
# insert the new table immediately after.
# Use the underlying XML elements.
inserted = 0
for h_idx in heading_para_indexes:
    heading_p = doc.paragraphs[h_idx]
    area_name = re.sub(r"\[.*?\]", "", heading_p.text).strip().rstrip(":")
    if not area_name or area_name.lower().startswith("glossary"):
        continue
    new_tbl = build_section_details_table(area_name)
    heading_p._p.addnext(new_tbl)
    inserted += 1

print(f"Inserted {inserted} Section Details tables.")

# ───────────────────────────────────────────────────────────────────────
# Wrap each new Section Details table in a Repeating Section Content
# Control so users get a native "+ Add Another" button to duplicate the
# block (Restroom #2, CTE Lab #3, etc.). Word 2013+/Word for Mac/Web.
# ───────────────────────────────────────────────────────────────────────

def wrap_in_repeating_section(table_element, area_name):
    """Wrap a w:tbl element inside a Repeating Section Content Control."""
    # Outer SDT (repeatingSection)
    outer = etree.SubElement(table_element.getparent(), qn("w:sdt"))
    outer_sdtPr = etree.SubElement(outer, qn("w:sdtPr"))
    rid = etree.SubElement(outer_sdtPr, qn("w:id"))
    rid.set(qn("w:val"), str(abs(hash("rs"+area_name)) % 99999999))
    alias = etree.SubElement(outer_sdtPr, qn("w:alias"))
    alias.set(qn("w:val"), f"{area_name} — instances")
    tag = etree.SubElement(outer_sdtPr, qn("w:tag"))
    tag.set(qn("w:val"), f"rs_{re.sub(r'[^A-Za-z0-9]+', '_', area_name)}")
    etree.SubElement(outer_sdtPr, f"{{{W15_NS}}}repeatingSection")
    etree.SubElement(outer, qn("w:sdtEndPr"))
    outer_content = etree.SubElement(outer, qn("w:sdtContent"))

    # Inner SDT (repeatingSectionItem)
    inner = etree.SubElement(outer_content, qn("w:sdt"))
    inner_sdtPr = etree.SubElement(inner, qn("w:sdtPr"))
    rid2 = etree.SubElement(inner_sdtPr, qn("w:id"))
    rid2.set(qn("w:val"), str(abs(hash("rsi"+area_name)) % 99999999))
    etree.SubElement(inner_sdtPr, f"{{{W15_NS}}}repeatingSectionItem")
    etree.SubElement(inner, qn("w:sdtEndPr"))
    inner_content = etree.SubElement(inner, qn("w:sdtContent"))

    # Move the original table into the inner content
    parent = table_element.getparent()
    idx = list(parent).index(table_element)
    parent.remove(table_element)
    # Now place outer where the table used to be
    parent.remove(outer)
    parent.insert(idx, outer)
    # And put the table inside the inner content
    inner_content.append(table_element)
    return outer

# Walk the body. For each newly-inserted Section Details table (which we
# can identify by its top row containing "SECTION DETAILS —"), wrap it
# in a repeating-section SDT.
wrapped = 0
for elem in list(body):
    if elem.tag != qn("w:tbl"):
        continue
    # Check first row first cell for our header marker
    text_blob = "".join(elem.itertext())
    if "SECTION DETAILS —" not in text_blob:
        continue
    # Extract the area name from the header
    m = re.search(r"SECTION DETAILS — ([^\n\r]+)", text_blob)
    area = m.group(1).strip() if m else "Area"
    wrap_in_repeating_section(elem, area)
    wrapped += 1

print(f"Wrapped {wrapped} Section Details tables in Repeating Section CCs.")

# ───────────────────────────────────────────────────────────────────────
# Insert a "How to use this guide" instructions block near the top.
# ───────────────────────────────────────────────────────────────────────
# Find the "Accessibility Standards" heading paragraph and insert after it.
insert_before = None
for p in doc.paragraphs:
    if "Accessibility Standards" in p.text and (p.style.name or "").strip() != "List Paragraph":
        insert_before = p
        break

if insert_before is not None:
    # Build an instructions table
    instr = OxmlElement("w:tbl")
    tblPr = OxmlElement("w:tblPr")
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:w"), "5000"); tblW.set(qn("w:type"), "pct"); tblPr.append(tblW)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top","left","bottom","right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "12"); b.set(qn("w:color"), GOLD)
        tblBorders.append(b)
    tblPr.append(tblBorders)
    instr.append(tblPr)
    tblGrid = OxmlElement("w:tblGrid")
    gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), "9700"); tblGrid.append(gc)
    instr.append(tblGrid)
    # Single row
    tr = OxmlElement("w:tr")
    tc = OxmlElement("w:tc")
    tcPr = OxmlElement("w:tcPr")
    w_ = OxmlElement("w:tcW"); w_.set(qn("w:w"), "9700"); w_.set(qn("w:type"), "dxa"); tcPr.append(w_)
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "F7F4EE"); tcPr.append(shd)
    tc.append(tcPr)

    def instr_para(text, bold=False, size=20, color=NAVY, after=80):
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        spc = OxmlElement("w:spacing"); spc.set(qn("w:after"), str(after)); ppr.append(spc)
        p.append(ppr)
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        if bold:
            rpr.append(OxmlElement("w:b"))
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(size)); rpr.append(sz)
        cc = OxmlElement("w:color"); cc.set(qn("w:val"), color); rpr.append(cc)
        r.append(rpr)
        t = OxmlElement("w:t"); t.text = text; r.append(t)
        p.append(r)
        return p

    tc.append(instr_para("How to use this guide — please read", bold=True, size=24))
    tc.append(instr_para(
        "Each of the 18 area sections begins with a SECTION DETAILS box. "
        "Click the highlighted fields to fill them in:", size=20))
    tc.append(instr_para(
        "• Location / Building name — type the location (e.g., \"Main Wing Restroom\").",
        size=20, after=40))
    tc.append(instr_para(
        "• Date constructed / Date of ADA modification — click to open a calendar picker.",
        size=20, after=40))
    tc.append(instr_para(
        "• Era — Standard at construction (and after modification) — pick the date range that applies. The dropdown labels include the standard name (Program Access, ANSI A117.1, UFAS, 1991 ADA, 2010 ADA).",
        size=20, after=120))
    tc.append(instr_para(
        "ADDING ANOTHER INSTANCE (Restroom #2, CTE Lab #3, etc.)", bold=True, size=20))
    tc.append(instr_para(
        "Click any SECTION DETAILS box. A small \"+\" button appears at the right edge. Click \"+\" to add another Section Details box for the same area. Repeat the questions below that section for the second instance, labeling each by location.",
        size=20, after=120))
    tc.append(instr_para(
        "AUTO-DETERMINED STANDARD", bold=True, size=20))
    tc.append(instr_para(
        "The applicable standard is the LATER of the two eras you pick (construction era vs. modification era). The reviewer will verify on site.",
        size=20, after=0))
    tr.append(tc)
    instr.append(tr)
    insert_before._p.addprevious(instr)
    print("Inserted instructions block.")

# ───────────────────────────────────────────────────────────────────────
# Save
# ───────────────────────────────────────────────────────────────────────
os.makedirs("outputs", exist_ok=True)
doc.save(OUT)
print(f"\n✓ Saved → {OUT}")
print(f"  File size: {os.path.getsize(OUT)/1024:.1f} KB")
