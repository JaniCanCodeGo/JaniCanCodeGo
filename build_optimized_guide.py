"""Build the optimized BLANK Facilities Review Guide — v2 (no Repeating Sections).

Fixes from v1:
  * v1 froze Word because each SDT referenced a glossary docPart placeholder
    (DefaultPlaceholder_-1854013440) that doesn't exist in the source doc's
    glossary part — Word retries the placeholder lookup repeatedly.
  * v1 generated SDT IDs by hashing a label that repeated across sections,
    causing ID collisions.
  * v1's Repeating Section Content Controls require a glossary entry for
    the repeat placeholder; without it Word renders blank or hangs.

v2 strategy:
  * Sequential unique IDs (1000 + counter).
  * No docPart placeholder reference — inline placeholder text in sdtContent.
  * No Repeating Section CCs — duplicate via copy-paste of templates at the
    back of the doc instead.
  * Element-by-element ALTERATIONS LOG per section (the correct rule:
    altered elements get the alteration-date standard; unaltered elements
    keep the original construction-date standard).
  * Section Templates page at the back for duplicating multi-instance areas.
"""
import os, re
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Inches, RGBColor

SRC = "/root/.claude/uploads/f318c6f4-c9c8-4a12-bcd0-bfd812a66767/60dc1d22-BLANK_Facilities__Review_Guide.docx"
OUT = "outputs/BLANK_Facilities_Review_Guide_optimized.docx"

ERA_DROPDOWN = [
    ("— Pick the era —", ""),
    ("≤ June 3, 1977 — Existing Facility (Program Access)",       "Program Access"),
    ("June 4, 1977 – January 17, 1991 — ANSI A117.1 (1961 R1971)", "ANSI A117.1"),
    ("January 18, 1991 – January 26, 1992 — UFAS (1984)",          "UFAS"),
    ("January 27, 1992 – September 14, 2010 — 1991 ADA / ADAAG",   "1991 ADA"),
    ("September 15, 2010 – March 14, 2012 — 1991 ADA OR 2010 ADA", "1991 ADA or 2010 ADA"),
    ("March 15, 2012 or later — 2010 ADA Standards",               "2010 ADA"),
]
YES_NO = [("— Pick one —", ""), ("Yes", "Yes"), ("No", "No")]

NAVY   = "1A2744"
GOLD   = "C8A84B"
CREAM  = "F7F4EE"
BORDER = "D4CEBD"
YELLOW = "FEF9C3"

# ── Global unique-ID counter for SDTs ─────────────────────────────────────────
_next_id = 1000
def next_id():
    global _next_id
    _next_id += 1
    return str(_next_id)

# ── SDT builders (no docPart placeholder reference; inline placeholder text) ──
def _placeholder_run(text):
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color"); color.set(qn("w:val"), "808080"); rpr.append(color)
    italic = OxmlElement("w:i"); rpr.append(italic)
    r.append(rpr)
    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text; r.append(t)
    return r

def make_date_picker(placeholder="Click to pick a date"):
    sdt = OxmlElement("w:sdt")
    sdtPr = OxmlElement("w:sdtPr")
    rid = OxmlElement("w:id"); rid.set(qn("w:val"), next_id()); sdtPr.append(rid)
    sdtPr.append(OxmlElement("w:showingPlcHdr"))
    date = OxmlElement("w:date")
    fmt = OxmlElement("w:dateFormat"); fmt.set(qn("w:val"), "MMMM d, yyyy"); date.append(fmt)
    lid = OxmlElement("w:lid"); lid.set(qn("w:val"), "en-US"); date.append(lid)
    stm = OxmlElement("w:storeMappedDataAs"); stm.set(qn("w:val"), "dateTime"); date.append(stm)
    cal = OxmlElement("w:calendar"); cal.set(qn("w:val"), "gregorian"); date.append(cal)
    sdtPr.append(date)
    sdt.append(sdtPr)
    sdt.append(OxmlElement("w:sdtEndPr"))
    content = OxmlElement("w:sdtContent")
    content.append(_placeholder_run(placeholder))
    sdt.append(content)
    return sdt

def make_dropdown(options, placeholder="— Pick one —"):
    sdt = OxmlElement("w:sdt")
    sdtPr = OxmlElement("w:sdtPr")
    rid = OxmlElement("w:id"); rid.set(qn("w:val"), next_id()); sdtPr.append(rid)
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
    content = OxmlElement("w:sdtContent")
    content.append(_placeholder_run(placeholder))
    sdt.append(content)
    return sdt

def make_plain_text(placeholder="Click to enter text"):
    sdt = OxmlElement("w:sdt")
    sdtPr = OxmlElement("w:sdtPr")
    rid = OxmlElement("w:id"); rid.set(qn("w:val"), next_id()); sdtPr.append(rid)
    sdtPr.append(OxmlElement("w:showingPlcHdr"))
    sdtPr.append(OxmlElement("w:text"))
    sdt.append(sdtPr)
    sdt.append(OxmlElement("w:sdtEndPr"))
    content = OxmlElement("w:sdtContent")
    content.append(_placeholder_run(placeholder))
    sdt.append(content)
    return sdt

# ── Paragraph / table builders ────────────────────────────────────────────────
def _para(text=None, *, bold=False, size=22, color="1A1A1A", after=80, runs_extra=None):
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    spc = OxmlElement("w:spacing"); spc.set(qn("w:after"), str(after)); ppr.append(spc)
    p.append(ppr)
    if text is not None:
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        if bold: rpr.append(OxmlElement("w:b"))
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(size)); rpr.append(sz)
        cc = OxmlElement("w:color"); cc.set(qn("w:val"), color); rpr.append(cc)
        r.append(rpr)
        t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text; r.append(t)
        p.append(r)
    if runs_extra:
        for run in runs_extra:
            p.append(run)
    return p

def _cell(width_dxa, paras, shade=None, gridSpan=None):
    tc = OxmlElement("w:tc")
    tcPr = OxmlElement("w:tcPr")
    w_ = OxmlElement("w:tcW"); w_.set(qn("w:w"), str(width_dxa)); w_.set(qn("w:type"), "dxa"); tcPr.append(w_)
    if gridSpan:
        gs = OxmlElement("w:gridSpan"); gs.set(qn("w:val"), str(gridSpan)); tcPr.append(gs)
    if shade:
        shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), shade); tcPr.append(shd)
    tc.append(tcPr)
    for p in paras:
        tc.append(p)
    return tc

def _tbl(grid_cols, rows):
    """grid_cols: list of dxa widths. rows: list of cells lists (or pre-built tr)."""
    tbl = OxmlElement("w:tbl")
    tblPr = OxmlElement("w:tblPr")
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:w"), "5000"); tblW.set(qn("w:type"), "pct"); tblPr.append(tblW)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top","left","bottom","right","insideH","insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "4"); b.set(qn("w:color"), BORDER)
        tblBorders.append(b)
    tblPr.append(tblBorders)
    tblLook = OxmlElement("w:tblLook"); tblLook.set(qn("w:val"), "04A0"); tblPr.append(tblLook)
    tbl.append(tblPr)
    grid = OxmlElement("w:tblGrid")
    for w_ in grid_cols:
        gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), str(w_)); grid.append(gc)
    tbl.append(grid)
    for cells in rows:
        if hasattr(cells, "tag"):  # already a tr element
            tbl.append(cells); continue
        tr = OxmlElement("w:tr")
        for c in cells:
            tr.append(c)
        tbl.append(tr)
    return tbl

# ── Section Details table (per area) ─────────────────────────────────────────
def section_details_table(area_name):
    LBL = 3200
    VAL = 6500
    rows = []

    # Title row — full width
    title_run = OxmlElement("w:r")
    title_rpr = OxmlElement("w:rPr")
    title_rpr.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "24"); title_rpr.append(sz)
    cc = OxmlElement("w:color"); cc.set(qn("w:val"), GOLD); title_rpr.append(cc)
    title_run.append(title_rpr)
    tt = OxmlElement("w:t"); tt.text = f"SECTION DETAILS — {area_name.upper()}"; title_run.append(tt)
    title_p = _para(after=80, runs_extra=[title_run])
    rows.append([_cell(LBL+VAL, [title_p], shade=NAVY, gridSpan=2)])

    def kv_row(label, control):
        lab_p = _para(label, bold=True, size=20, color=NAVY, after=0)
        val_p = _para(after=0)
        val_p.append(control)
        rows.append([_cell(LBL, [lab_p], shade=CREAM), _cell(VAL, [val_p])])

    kv_row("Location / Building name:",
           make_plain_text("Click to type the location (e.g., \"Main Wing Restroom\")"))
    kv_row("Original construction date:",
           make_date_picker("Click to pick the original construction date"))
    kv_row("Original construction era → standard:",
           make_dropdown(ERA_DROPDOWN, "Click to pick original construction era"))
    kv_row("Has this area or its elements been altered?",
           make_dropdown(YES_NO, "Pick Yes or No"))

    # Rule reminder row
    rule_p1 = _para("REVIEWER RULE — Element-by-element alteration analysis",
                    bold=True, size=18, color="713F12", after=40)
    rule_p2 = _para(
        "The original construction era controls UNALTERED elements only. If specific elements were altered (e.g., toilet seat, grab bar, ramp, signage), each altered element is evaluated under the standard in effect at the time of alteration — record those in the Alterations Log below.",
        size=18, color="713F12", after=80)
    rows.append([_cell(LBL+VAL, [rule_p1, rule_p2], shade=YELLOW, gridSpan=2)])

    return _tbl([LBL, VAL], rows)

# ── Alterations Log table (per area) ─────────────────────────────────────────
def alterations_log_table(area_name, n_rows=4):
    EL = 3000
    DT = 2200
    ST = 2500
    DS = 1900
    rows = []

    # Title row
    title_run = OxmlElement("w:r")
    title_rpr = OxmlElement("w:rPr")
    title_rpr.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "22"); title_rpr.append(sz)
    cc = OxmlElement("w:color"); cc.set(qn("w:val"), GOLD); title_rpr.append(cc)
    title_run.append(title_rpr)
    tt = OxmlElement("w:t"); tt.text = f"ALTERATIONS LOG — {area_name.upper()}"; title_run.append(tt)
    rows.append([_cell(EL+DT+ST+DS, [_para(after=80, runs_extra=[title_run])], shade=NAVY, gridSpan=4)])

    instr = _para(
        "List each element altered, the date of alteration, and the era/standard at the time of alteration. Add more rows by tabbing in the last cell, or copy a blank row.",
        size=18, color="6B6659", after=80)
    rows.append([_cell(EL+DT+ST+DS, [instr], shade=CREAM, gridSpan=4)])

    # Header row
    def th(text, w):
        run = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        rpr.append(OxmlElement("w:b"))
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "18"); rpr.append(sz)
        cc = OxmlElement("w:color"); cc.set(qn("w:val"), GOLD); rpr.append(cc)
        run.append(rpr)
        t = OxmlElement("w:t"); t.text = text.upper(); run.append(t)
        return _cell(w, [_para(after=0, runs_extra=[run])], shade=NAVY)
    rows.append([
        th("Element altered", EL),
        th("Date of alteration", DT),
        th("Era → standard at alteration", ST),
        th("Description", DS),
    ])

    # Data rows
    for _ in range(n_rows):
        rows.append([
            _cell(EL, [_para(after=0, runs_extra=[make_plain_text("(e.g., \"Toilet seat\")")])]),
            _cell(DT, [_para(after=0, runs_extra=[make_date_picker("Click to pick date")])]),
            _cell(ST, [_para(after=0, runs_extra=[make_dropdown(ERA_DROPDOWN, "Pick era → standard")])]),
            _cell(DS, [_para(after=0, runs_extra=[make_plain_text("(optional)")])]),
        ])
    return _tbl([EL, DT, ST, DS], rows)

# ── Open source, insert Section Details + Alterations Log after each H1 ──────
print("Loading source guide…")
doc = Document(SRC)
print(f"  {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")

body = doc.element.body

# Collect H1 paragraphs (the 18 area sections + glossary)
heading_paras = []
for p in doc.paragraphs:
    if (p.style.name or "").strip() == "Heading 1":
        heading_paras.append(p)

print(f"Found {len(heading_paras)} Heading 1 paragraphs.")

inserted = 0
multi_instance_areas = set()
for h in heading_paras:
    area = re.sub(r"\[.*?\]", "", h.text).strip().rstrip(":")
    if not area or area.lower().startswith("glossary"):
        continue
    # Insert Alterations Log first (so addnext stacks correctly: H1 → details → log)
    log = alterations_log_table(area, n_rows=4)
    h._p.addnext(log)
    # Then Section Details (will land right after H1, before log)
    det = section_details_table(area)
    h._p.addnext(det)
    inserted += 1
    # Multi-instance areas worth a copy-paste template
    if any(k in area.lower() for k in (
        "restroom", "drinking fountain", "water cooler", "cte classroom",
        "labs/shops", "stairway", "ramp", "curb ramp", "entrance", "room",
        "office", "elevator", "lift",
    )):
        multi_instance_areas.add(area)

print(f"Inserted Section Details + Alterations Log into {inserted} sections.")
print(f"Marked {len(multi_instance_areas)} areas for the copy-paste template page.")

# ── Top-of-doc instructions block ────────────────────────────────────────────
def instructions_block():
    pad = OxmlElement("w:p")
    rows = []
    title = _para("HOW TO USE THIS GUIDE — READ BEFORE COMPLETING", bold=True, size=26, color="C8A84B", after=120)
    p1 = _para("Each of the 18 area sections begins with a SECTION DETAILS table and an ALTERATIONS LOG table.", size=20, after=80)
    p2 = _para("SECTION DETAILS — fill in:", bold=True, size=20, after=40)
    p3 = _para("• Location / Building name (e.g., \"Main Wing Restroom\").", size=20, after=20)
    p4 = _para("• Original construction date — calendar date picker.", size=20, after=20)
    p5 = _para("• Original construction era → standard — dropdown. The era label tells you which standard applies (Program Access / ANSI A117.1 / UFAS / 1991 ADA / 2010 ADA).", size=20, after=80)
    p6 = _para("ALTERATIONS LOG — element-by-element", bold=True, size=20, after=40)
    p7 = _para("Per 28 CFR § 35.151(b) and the 2010 ADA Standards § 202.3, alteration analysis is element-by-element. The ORIGINAL construction era controls UNALTERED elements. Each ALTERED element is evaluated under the standard in effect on the alteration date. Record each altered element in the ALTERATIONS LOG table.", size=20, after=80)
    p8 = _para("ADDING ANOTHER LOCATION (Restroom #2, CTE Lab #3, etc.)", bold=True, size=20, after=40)
    p9 = _para("Multi-instance areas (Restrooms, Drinking Fountains, CTE Classrooms, Labs/Shops, etc.) have copy-paste templates at the end of this guide (page SECTION TEMPLATES). Select the desired template, copy (Ctrl+C / Cmd+C), and paste (Ctrl+V / Cmd+V) at the appropriate point in the guide.", size=20, after=80)
    p10 = _para("CORRECTIVE ACTIONS WILL BE WRITTEN AT 2010 ADA", bold=True, size=20, color="B91C1C", after=40)
    p11 = _para("Per CDE OEO policy and OCR practice, all CORRECTIVE ACTIONS on the LOF are written to the 2010 ADA Standards — irrespective of original construction date or alteration date. The VIOLATION is cited at the applicable standard for that element; the FIX is 2010 ADA.", size=20, color="B91C1C", after=80)
    return [title, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11]

# Find the "Accessibility Standards" paragraph and insert instructions BEFORE it.
insert_before = None
for p in doc.paragraphs:
    if "Accessibility Standards" in p.text and (p.style.name or "").strip() not in ("List Paragraph",):
        insert_before = p
        break

if insert_before is not None:
    for para in instructions_block():
        insert_before._p.addprevious(para)
    print("Instructions block inserted.")

# ── Append "SECTION TEMPLATES TO COPY" page at the back ─────────────────────
print("Appending SECTION TEMPLATES page…")

# Title heading
heading_run = OxmlElement("w:r")
heading_rpr = OxmlElement("w:rPr")
heading_rpr.append(OxmlElement("w:b"))
hsz = OxmlElement("w:sz"); hsz.set(qn("w:val"), "36"); heading_rpr.append(hsz)
hcc = OxmlElement("w:color"); hcc.set(qn("w:val"), NAVY); heading_rpr.append(hcc)
heading_run.append(heading_rpr)
ht = OxmlElement("w:t"); ht.text = "SECTION TEMPLATES — copy & paste to add another instance"; heading_run.append(ht)
heading_p = _para(after=120, runs_extra=[heading_run])

# Page break before the templates page
pb_p = OxmlElement("w:p")
pb_r = OxmlElement("w:r")
pb_br = OxmlElement("w:br"); pb_br.set(qn("w:type"), "page"); pb_r.append(pb_br)
pb_p.append(pb_r)
body.append(pb_p)
body.append(heading_p)

intro = _para(
    "Each template below is a blank SECTION DETAILS + ALTERATIONS LOG pair for one location. "
    "To add another Restroom (Restroom #2), CTE Classroom (CTE Classroom #3), etc.:",
    size=22, after=40)
step1 = _para("1.  Select the appropriate template below (highlight from the SECTION DETAILS title through the last row of the ALTERATIONS LOG).", size=22, after=20)
step2 = _para("2.  Copy:  Ctrl+C (Windows) / Cmd+C (Mac).", size=22, after=20)
step3 = _para("3.  Scroll up to the matching area section earlier in the guide. Click at the END of that area's question list.", size=22, after=20)
step4 = _para("4.  Paste:  Ctrl+V (Windows) / Cmd+V (Mac).", size=22, after=20)
step5 = _para("5.  Fill in the new instance's Location, dates, era, and alteration log.", size=22, after=120)

for p in (intro, step1, step2, step3, step4, step5):
    body.append(p)

# Add one template per multi-instance area
template_areas = sorted(multi_instance_areas)
for area in template_areas:
    spacer = _para(after=120)
    body.append(spacer)
    # Sub-heading
    sh_run = OxmlElement("w:r")
    sh_rpr = OxmlElement("w:rPr")
    sh_rpr.append(OxmlElement("w:b"))
    shsz = OxmlElement("w:sz"); shsz.set(qn("w:val"), "26"); sh_rpr.append(shsz)
    shcc = OxmlElement("w:color"); shcc.set(qn("w:val"), NAVY); sh_rpr.append(shcc)
    sh_run.append(sh_rpr)
    sh_t = OxmlElement("w:t"); sh_t.text = f"Template — {area}"; sh_run.append(sh_t)
    sh_p = _para(after=80, runs_extra=[sh_run])
    body.append(sh_p)

    # Add details + log table for this template
    body.append(section_details_table(area))
    body.append(alterations_log_table(area, n_rows=4))

# ── Save ────────────────────────────────────────────────────────────────────
os.makedirs("outputs", exist_ok=True)
doc.save(OUT)
print(f"\n✓ Saved → {OUT}")
print(f"  File size: {os.path.getsize(OUT)/1024:.1f} KB")
print(f"  Final SDT ID counter: {_next_id}")
