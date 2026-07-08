"""Build the Civil Rights Review site documents (v3, per CDE OEO corrections).

Outputs (both in this directory):
  A) Civil_Rights_Review_Building_Accessibility.docx
     The main guide the LEA / school site fills out. Data collection ONLY:
     no internal reviewer process content of any kind.
  B) Program_Access_Staff_Interview.docx
     The 14 Program Access interview questions as a standalone document
     (conducted by the reviewer with Facilities and Maintenance &
     Operations staff), with response and notes fields.

v3 changes (owner corrections, Murjani McTier, CDE OEO):
  * Title is "Civil Rights Review - Building Accessibility" (already on the
    base document's cover page and footer, preserved verbatim).
  * Standards Applicability Guide page REMOVED (that knowledge lives in the
    LOF agent, not the site document).
  * All reviewer-facing notes and prompts REMOVED: no shaded rule boxes, no
    standard-determination prompts. The site document only collects data.
  * Construction-date dropdowns REMOVED from Section Details. Kept as plain
    data fields: Year Built, an alterations Yes/No gate, and the
    Modifications. The Program Reviewer determines standards later,
    outside this document.
  * Alterations Log tables reduced to TWO columns: Element Altered and
    Date Altered.
  * Corrective Action Summary REMOVED (belongs only in the Letter of
    Findings).
  * Section Templates page REMOVED. Each multi-location section's fillable
    block is wrapped in a native Word Repeating Section Content Control
    (w15:repeatingSection, the proven v5 mechanism from
    build_optimized_guide.py): click in the block and Word shows a "+"
    control to duplicate it in place. Unique sequential SDT ids, no
    docPart placeholder references.
  * No numbered instance labels; each block has a blank Location /
    Building Name field instead.
  * Play Areas section REMOVED (high schools).
  * Standalone Signage section REMOVED (each area carries its own signage
    questions).
  * Standalone Parking Lot Lighting and Surface Condition section REMOVED.
    A surface-condition question was added inside Accessible Parking.
    No lighting question (the base document has none).
  * Medical section renamed "Medical and First-Aid Areas (including
    Nurse's Office)".
  * Program Access staff interview moved out entirely, into document B.

Retained fundamentals:
  * NO em-dash characters anywhere (hard build-blocking house rule); a
    final scrub pass cleans any surviving in preserved front matter.
  * Standard-specific threshold notations beside each measurement field
    (per cde-facilities-toolset/references/thresholds.md).
  * Male / Female / Non-Gender Specific dropdowns (restrooms, locker rooms).
  * Free-text room-name field for Rooms & Offices.
  * Arial font, 12 pt body. Core properties author and last-modified-by
    set to "Murjani McTier".
  * Cover page, Accessibility Standards listing, Tips for Measuring with
    all 8 screenclip images, and Glossary of Terms preserved from the
    base doc (BLANK_Facilities_Review_Guide_optimized.docx).
"""
import os, re, copy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Inches, RGBColor
from lxml import etree

# w15 namespace, Microsoft Word 2012 wordml extensions (repeatingSection)
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"

# Base document: the previously generated guide in this repo. Its cover page,
# Accessibility Standards listing, Tips for Measuring (8 screenclip images),
# and Glossary of Terms are preserved; everything in between is replaced.
_HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.normpath(os.path.join(_HERE, "..", "..", "BLANK_Facilities_Review_Guide_optimized.docx"))
OUT_GUIDE     = os.path.join(_HERE, "Civil_Rights_Review_Building_Accessibility.docx")
OUT_INTERVIEW = os.path.join(_HERE, "Program_Access_Staff_Interview.docx")
OLD_OUT       = os.path.join(_HERE, "Accessibility_Facilities_Checklist_v2.docx")

DOC_AUTHOR = "Murjani McTier"

# ── Dropdowns ────────────────────────────────────────────────────────────────
YES_NO   = [("- Pick one -", ""), ("Yes", "Yes"), ("No", "No")]
YES_NO_NA = [("- Pick one -", ""), ("Yes", "Yes"), ("No", "No"), ("N/A", "N/A")]

# ── Brand palette ────────────────────────────────────────────────────────────
NAVY   = "1A2744"
GOLD   = "C8A84B"
CREAM  = "F7F4EE"
BORDER = "D4CEBD"
YELLOW = "FEF9C3"
LAVENDER = "E9D5FF"
SAGE   = "EDF0EB"

# ── SDT ID counter ───────────────────────────────────────────────────────────
_next_id = 1000
def next_id():
    global _next_id; _next_id += 1; return str(_next_id)

# ── XML builders ─────────────────────────────────────────────────────────────
def _placeholder_run(text):
    r = OxmlElement("w:r"); rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color"); color.set(qn("w:val"), "808080"); rpr.append(color)
    rpr.append(OxmlElement("w:i")); r.append(rpr)
    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text; r.append(t)
    return r

def make_date_picker(placeholder="Click to pick a date"):
    sdt = OxmlElement("w:sdt"); sp = OxmlElement("w:sdtPr")
    rid = OxmlElement("w:id"); rid.set(qn("w:val"), next_id()); sp.append(rid)
    sp.append(OxmlElement("w:showingPlcHdr"))
    date = OxmlElement("w:date")
    for k, v in (("dateFormat","MMMM d, yyyy"),("lid","en-US"),("storeMappedDataAs","dateTime"),("calendar","gregorian")):
        e = OxmlElement(f"w:{k}"); e.set(qn("w:val"), v); date.append(e)
    sp.append(date); sdt.append(sp); sdt.append(OxmlElement("w:sdtEndPr"))
    c = OxmlElement("w:sdtContent"); c.append(_placeholder_run(placeholder)); sdt.append(c)
    return sdt

def make_dropdown(options, placeholder="- Pick one -"):
    sdt = OxmlElement("w:sdt"); sp = OxmlElement("w:sdtPr")
    rid = OxmlElement("w:id"); rid.set(qn("w:val"), next_id()); sp.append(rid)
    sp.append(OxmlElement("w:showingPlcHdr"))
    ddl = OxmlElement("w:dropDownList")
    for disp, val in options:
        li = OxmlElement("w:listItem")
        li.set(qn("w:displayText"), disp); li.set(qn("w:value"), val)
        ddl.append(li)
    sp.append(ddl); sdt.append(sp); sdt.append(OxmlElement("w:sdtEndPr"))
    c = OxmlElement("w:sdtContent"); c.append(_placeholder_run(placeholder)); sdt.append(c)
    return sdt

def make_plain_text(placeholder="Click to enter text"):
    sdt = OxmlElement("w:sdt"); sp = OxmlElement("w:sdtPr")
    rid = OxmlElement("w:id"); rid.set(qn("w:val"), next_id()); sp.append(rid)
    sp.append(OxmlElement("w:showingPlcHdr")); sp.append(OxmlElement("w:text"))
    sdt.append(sp); sdt.append(OxmlElement("w:sdtEndPr"))
    c = OxmlElement("w:sdtContent"); c.append(_placeholder_run(placeholder)); sdt.append(c)
    return sdt

# ── Paragraph / table builders ───────────────────────────────────────────────
def _para(text=None, *, bold=False, size=24, color="1A1A1A", after=80, runs_extra=None, italic=False):
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    spc = OxmlElement("w:spacing"); spc.set(qn("w:after"), str(after)); ppr.append(spc)
    p.append(ppr)
    if text is not None:
        r = OxmlElement("w:r"); rpr = OxmlElement("w:rPr")
        if bold: rpr.append(OxmlElement("w:b"))
        if italic: rpr.append(OxmlElement("w:i"))
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(size)); rpr.append(sz)
        cc = OxmlElement("w:color"); cc.set(qn("w:val"), color); rpr.append(cc)
        r.append(rpr)
        t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text; r.append(t)
        p.append(r)
    if runs_extra:
        for run in runs_extra: p.append(run)
    return p

def _cell(width_dxa, paras, shade=None, gridSpan=None):
    tc = OxmlElement("w:tc"); tcPr = OxmlElement("w:tcPr")
    w_ = OxmlElement("w:tcW"); w_.set(qn("w:w"), str(width_dxa)); w_.set(qn("w:type"), "dxa"); tcPr.append(w_)
    if gridSpan:
        gs = OxmlElement("w:gridSpan"); gs.set(qn("w:val"), str(gridSpan)); tcPr.append(gs)
    if shade:
        shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), shade); tcPr.append(shd)
    tc.append(tcPr)
    for p in paras: tc.append(p)
    return tc

def _tbl(grid_cols, rows):
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
        if hasattr(cells, "tag"): tbl.append(cells); continue
        tr = OxmlElement("w:tr")
        for c in cells: tr.append(c)
        tbl.append(tr)
    return tbl

# ── Section-Details + Alterations-Log tables (v2 retained) ───────────────────
def section_details_table(area_name):
    LBL, VAL = 3200, 6500
    rows = []
    title_run = OxmlElement("w:r")
    rp = OxmlElement("w:rPr"); rp.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "24"); rp.append(sz)
    col = OxmlElement("w:color"); col.set(qn("w:val"), GOLD); rp.append(col)
    title_run.append(rp)
    t = OxmlElement("w:t"); t.text = f"SECTION DETAILS, {area_name.upper()}"; title_run.append(t)
    rows.append([_cell(LBL+VAL, [_para(after=80, runs_extra=[title_run])], shade=NAVY, gridSpan=2)])

    def kv(lbl, ctrl):
        lp = _para(lbl, bold=True, size=24, color=NAVY, after=0)
        vp = _para(after=0); vp.append(ctrl)
        rows.append([_cell(LBL, [lp], shade=CREAM), _cell(VAL, [vp])])

    kv("Location / Building Name:", make_plain_text("Click to type the location"))
    kv("Year Built:", make_plain_text("Click to enter the year built"))
    kv("Have any alterations or ADA modifications been made since original construction? If Yes, complete the Alterations Log below.",
       make_dropdown([("Yes", "Yes"), ("No", "No")]))
    return _tbl([LBL, VAL], rows)

def alterations_log_table(area_name, n_rows=4):
    EL, DT = 5800, 3800
    rows = []
    tr_run = OxmlElement("w:r")
    rp = OxmlElement("w:rPr"); rp.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "24"); rp.append(sz)
    cc = OxmlElement("w:color"); cc.set(qn("w:val"), GOLD); rp.append(cc)
    tr_run.append(rp)
    t = OxmlElement("w:t"); t.text = f"ALTERATIONS LOG, {area_name.upper()}"; tr_run.append(t)
    rows.append([_cell(EL+DT, [_para(after=80, runs_extra=[tr_run])], shade=NAVY, gridSpan=2)])
    instr = _para("List each element altered and the date of the alteration. Add more rows by tabbing in the last cell.", size=18, color="6B6659", after=80)
    rows.append([_cell(EL+DT, [instr], shade=CREAM, gridSpan=2)])
    def th(text, w):
        rn = OxmlElement("w:r"); rp = OxmlElement("w:rPr")
        rp.append(OxmlElement("w:b"))
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "18"); rp.append(sz)
        cc = OxmlElement("w:color"); cc.set(qn("w:val"), GOLD); rp.append(cc)
        rn.append(rp); t = OxmlElement("w:t"); t.text = text.upper(); rn.append(t)
        return _cell(w, [_para(after=0, runs_extra=[rn])], shade=NAVY)
    rows.append([th("Element Altered", EL), th("Date Altered", DT)])
    for _ in range(n_rows):
        rows.append([
            _cell(EL, [_para(after=0, runs_extra=[make_plain_text('(e.g., "Toilet seat")')])]),
            _cell(DT, [_para(after=0, runs_extra=[make_date_picker('Click to pick date')])]),
        ])
    return _tbl([EL, DT], rows)

# ─────────────────────────────────────────────────────────────────────────────
# REPEATING SECTION CONTENT CONTROL helper (proven v5 mechanism, copied from
# build_optimized_guide.py). Wrap the section's fillable block in a
# w15:repeatingSection SDT. When the cursor is inside the block in Word, a
# "+" button appears; clicking it duplicates the inner item's content in
# place. Word 2013+ / Mac 2016+ / Word for the web. Unique sequential SDT
# ids; NO docPart placeholder references (a missing docPart placeholder plus
# duplicate SDT ids froze Word in v1; v5 solved it exactly this way).
# ─────────────────────────────────────────────────────────────────────────────
def wrap_repeating_section(elements, alias_name):
    """Wrap a list of block-level elements in a Repeating Section Content Control."""
    # Outer SDT, declares this is a repeating section
    outer = OxmlElement("w:sdt")
    outer_pr = OxmlElement("w:sdtPr")
    rid = OxmlElement("w:id"); rid.set(qn("w:val"), next_id()); outer_pr.append(rid)
    alias = OxmlElement("w:alias"); alias.set(qn("w:val"), alias_name); outer_pr.append(alias)
    tag = OxmlElement("w:tag")
    tag.set(qn("w:val"), "rs_" + re.sub(r"[^A-Za-z0-9]+", "_", alias_name).strip("_"))
    outer_pr.append(tag)
    etree.SubElement(outer_pr, f"{{{W15_NS}}}repeatingSection")
    outer.append(outer_pr)
    outer.append(OxmlElement("w:sdtEndPr"))
    outer_content = OxmlElement("w:sdtContent")
    outer.append(outer_content)

    # Inner SDT, one repeating-section ITEM (the unit that gets cloned)
    inner = OxmlElement("w:sdt")
    inner_pr = OxmlElement("w:sdtPr")
    rid2 = OxmlElement("w:id"); rid2.set(qn("w:val"), next_id()); inner_pr.append(rid2)
    etree.SubElement(inner_pr, f"{{{W15_NS}}}repeatingSectionItem")
    inner.append(inner_pr)
    inner.append(OxmlElement("w:sdtEndPr"))
    inner_content = OxmlElement("w:sdtContent")
    for el in elements:
        inner_content.append(el)
    inner.append(inner_content)

    outer_content.append(inner)
    return outer

print("Loading source guide...")
doc = Document(SRC)
print(f"  {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
body = doc.element.body

# ─────────────────────────────────────────────────────────────────────────────
# FIELD RENDERER, turns (label, hint, type, threshold) tuples into Word rows
# ─────────────────────────────────────────────────────────────────────────────
# Field types:
#   'yn'    Yes/No checkbox row
#   'yna'   Yes/No/N/A checkbox row
#   'num'   numeric measurement (placeholder: "in" or as specified)
#   'text'  short text
#   'wide'  longer text (single-line wider)
#   'area'  textarea / multi-line text
#   'sel'   dropdown with options
#   'pair'  W: __  D: __ (or any two named numerics)
#   'triple' H: __  W: __  D: __
#   'subs'  subsection heading (renders as gold bar, no input)
#   'note'  inline help/instruction note (blue)
#   'warn'  inline warning note (amber)
#   'meta'  Cover-style label + input grid item
#
# Tuple format:
#   (kind, label, hint, *extra)
#       kind = one of the types above
#       label = field label (or heading text for subs/note/warn)
#       hint = small subtitle line; '' if none
#       extras vary by kind:
#         num: (threshold_text, placeholder='in')
#         sel: (options_list,)
#         pair/triple: nothing extra
#         text/wide/area: (placeholder,)

def _checkbox_glyph():
    return "☐ "

def render_field(field):
    kind = field[0]
    if kind == 'subs':
        # Subsection heading, gold bar
        rn = OxmlElement("w:r")
        rp = OxmlElement("w:rPr"); rp.append(OxmlElement("w:b"))
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "24"); rp.append(sz)
        cc = OxmlElement("w:color"); cc.set(qn("w:val"), NAVY); rp.append(cc)
        rn.append(rp)
        t = OxmlElement("w:t"); t.text = field[1].upper(); rn.append(t)
        return [_para(after=60, runs_extra=[rn])]
    if kind == 'note':
        # Blue inline note
        return [_tbl([9700], [[_cell(9700, [_para(field[1], size=18, color="1A4A6A", after=40)], shade="E8F4FD")]])]
    if kind == 'warn':
        # Amber inline warning
        return [_tbl([9700], [[_cell(9700, [_para(field[1], size=18, color="7A4A00", after=40)], shade=YELLOW)]])]

    label, hint = field[1], field[2]
    # Build label paragraph (with optional hint)
    label_paras = [_para(label, bold=True, size=24, color=NAVY, after=20)]
    if hint:
        label_paras.append(_para(hint, size=18, color="5A6A7A", italic=True, after=0))

    # Build input paragraph
    input_paras = [_para(after=0)]
    p = input_paras[0]
    if kind == 'yn':
        r1 = OxmlElement("w:r"); rp = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "24"); rp.append(sz)
        r1.append(rp)
        t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = "☐ Yes      ☐ No"; r1.append(t)
        p.append(r1)
    elif kind == 'yna':
        r1 = OxmlElement("w:r"); rp = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "24"); rp.append(sz)
        r1.append(rp)
        t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = "☐ Yes      ☐ No      ☐ N/A"; r1.append(t)
        p.append(r1)
    elif kind == 'num':
        ph = field[4] if len(field) > 4 else "in"
        p.append(make_plain_text(f"Click to enter {ph}"))
    elif kind == 'text':
        ph = field[3] if len(field) > 3 else "Click to enter"
        p.append(make_plain_text(ph))
    elif kind == 'wide':
        ph = field[3] if len(field) > 3 else "Click to enter"
        p.append(make_plain_text(ph))
    elif kind == 'area':
        ph = field[3] if len(field) > 3 else "Click to enter narrative"
        p.append(make_plain_text(ph))
    elif kind == 'sel':
        opts = [("- Pick one -", "")] + [(o, o) for o in field[3]]
        p.append(make_dropdown(opts, "Click to pick"))
    elif kind == 'pair':
        names = field[3] if len(field) > 3 else ("W", "D")
        for i, n in enumerate(names):
            if i > 0:
                rspace = OxmlElement("w:r"); ts = OxmlElement("w:t"); ts.set(qn("xml:space"), "preserve"); ts.text = "   "; rspace.append(ts); p.append(rspace)
            rn = OxmlElement("w:r"); rp = OxmlElement("w:rPr"); rp.append(OxmlElement("w:b"))
            rn.append(rp); tt = OxmlElement("w:t"); tt.set(qn("xml:space"), "preserve"); tt.text = f"{n}: "; rn.append(tt); p.append(rn)
            p.append(make_plain_text("in"))
    elif kind == 'triple':
        for i, n in enumerate(("H", "W", "D")):
            if i > 0:
                rspace = OxmlElement("w:r"); ts = OxmlElement("w:t"); ts.set(qn("xml:space"), "preserve"); ts.text = "   "; rspace.append(ts); p.append(rspace)
            rn = OxmlElement("w:r"); rp = OxmlElement("w:rPr"); rp.append(OxmlElement("w:b"))
            rn.append(rp); tt = OxmlElement("w:t"); tt.set(qn("xml:space"), "preserve"); tt.text = f"{n}: "; rn.append(tt); p.append(rn)
            p.append(make_plain_text("in"))
    elif kind == 'meta':
        ph = field[3] if len(field) > 3 else ""
        p.append(make_plain_text(ph or "Click to enter"))

    # Build a 2-column table row: label cell + input cell
    LBL, INP = 4400, 5300
    return [_tbl([LBL, INP], [[
        _cell(LBL, label_paras),
        _cell(INP, input_paras),
    ]])]

def render_section_fields(fields):
    out = []
    for f in fields:
        out.extend(render_field(f))
    return out

def render_section(area_name, fields, multi_instance=True):
    """Render a complete section: Section Details + Alterations Log + the field set.

    Multi-location sections wrap the whole fillable block in a Repeating
    Section Content Control so the site can add another copy inline with
    Word's "+" control. No numbered instance labels; each copy has its own
    blank Location / Building Name field."""
    out = []
    # Section heading
    rn = OxmlElement("w:r")
    rp = OxmlElement("w:rPr"); rp.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "32"); rp.append(sz)
    cc = OxmlElement("w:color"); cc.set(qn("w:val"), NAVY); rp.append(cc)
    rn.append(rp)
    t = OxmlElement("w:t"); t.text = area_name; rn.append(t)
    # Page break first
    pbp = OxmlElement("w:p"); pbr = OxmlElement("w:r")
    pbk = OxmlElement("w:br"); pbk.set(qn("w:type"), "page"); pbr.append(pbk); pbp.append(pbr)
    out.append(pbp)
    out.append(_para(after=120, runs_extra=[rn]))
    # The fillable block: Section Details + Alterations Log + field set
    block = []
    block.append(section_details_table(area_name))
    block.append(_para(after=80))
    block.append(alterations_log_table(area_name, n_rows=4))
    block.append(_para(after=120))
    block.extend(render_section_fields(fields))
    if multi_instance:
        out.append(_para("If this area exists in more than one location, click anywhere inside the block below and use the plus (+) control at its edge to add another blank copy for each additional location.", size=18, color="5A6A7A", italic=True, after=80))
        out.append(wrap_repeating_section(block, f"{area_name} locations"))
        out.append(_para(after=80))
    else:
        out.extend(block)
    return out

# ─────────────────────────────────────────────────────────────────────────────
# SHARED FIELD GROUPS (door, signage, alarms, used by several sections)
# ─────────────────────────────────────────────────────────────────────────────
SIG_FIELDS = [
    ('subs', 'Signage', ''),
    ('yn',   'Identification signage provided?', ''),
    ('num',  'Height of signage, floor to centerline (inches)', '60" centerline, All standards', '57"-63" range', 'in'),
    ('num',  'Height of characters on signage (inches)', '', '', 'in'),
    ('yn',   'Signs mounted adjacent to latch side of door?', ''),
    ('yna',  'At double doors, signs on nearest adjacent wall?', ''),
    ('yn',   'Protruding objects within 3" of sign?', ''),
    ('yn',   'Non-glare finish on signs?', ''),
    ('yn',   'Characters raised and accompanied by Grade II Braille?', ''),
]
DOOR_FIELDS = [
    ('num',  'Width of doorway(s) (inches)', 'Min. 32", All standards', '', 'in'),
    ('num',  'Height of door handle, floor to handle (inches)', 'Max 48", All standards', '', 'in'),
    ('yn',   'Door openable without grasping or twisting wrist?', ''),
    ('text', 'If not, type of hardware on door', '', 'push bar, knob, lever, etc.'),
    ('yna',  'Automatic door opener provided?', ''),
    ('yna',  'Door closer takes at least 3 seconds to close?', ''),
    ('num',  'Door opening force (pounds)', 'Max 5 lbs interior, 1991 ADA/UFAS/2010 ADA | Max 8.5 lbs, ANSI', '', 'lbs'),
    ('num',  'Threshold height, Exterior (inches)', 'Max 1/2", All standards', '', 'in'),
    ('num',  'Threshold height, Interior (inches)', 'Max 1/2", All standards', '', 'in'),
    ('yna',  'Carpeting or mats 1/2 inch high or less?', ''),
]
ALARM_FIELDS = [
    ('subs', 'Alarms', ''),
    ('yn',   'Visual alarm provided?', ''),
    ('yn',   'Light flashes clear or normal white?', ''),
    ('yn',   'Audible alarm provided?', ''),
    ('yn',   'Alarm exceeds prevailing sound level in room?', ''),
]
NOTES_FIELD = [
    ('area', 'Additional Information / Observations', '', 'Narrative observations, photo references, corrective actions noted...'),
]

# ─────────────────────────────────────────────────────────────────────────────
# THE 22 INSTANCE SECTIONS (each with full HTML field set)
# Each entry: ('section_name', 'kind', [fields...])
#   kind = 'multi' , copy-paste template at the back for additional instances
#   kind = 'single', one instance only (Stadium, Cafeteria, Library, Gym, Telephones)
# ─────────────────────────────────────────────────────────────────────────────

PARKING_FIELDS = [
    ('subs', 'Space Counts', ''),
    ('num',  'Total number of parking spaces', '', '', '#'),
    ('num',  'Total number of accessible spaces', '', '', '#'),
    ('num',  'Total number of van accessible spaces', '', '', '#'),
    ('yn',   'At least 1 of every 6 accessible spaces is van accessible?', ''),
    ('yn',   'Accessible spaces located closest to nearest accessible route?', ''),
    ('subs', 'Dimensions', ''),
    ('num',  'Width of accessible parking spaces (inches)', 'Min. 96", All standards', '', 'in'),
    ('num',  'Width of access aisle (inches)', 'Min. 60" standard | Min. 96" van, ANSI/UFAS/1991 ADA | Min. 60" van, 2010 ADA if space is 132"', '', 'in'),
    ('num',  'Width of van accessible space (inches)', 'Min. 132" OR 96"+60" aisle, 2010 ADA | Min. 96"+96" aisle, ANSI/UFAS/1991 ADA', '', 'in'),
    ('num',  'Width of van access aisle (inches)', '', '', 'in'),
    ('num',  'Vertical clearance for van parking/access aisles (inches)', 'Min. 98", All standards', '', 'in'),
    ('yn',   'Ground surface slope less than 1:48?', ''),
    ('yn',   'Access aisle slope less than 1:48?', ''),
    ('yn',   'Parking surface stable, firm, slip-resistant, and free of potholes or broken pavement at accessible spaces and access aisles?', ''),
    ('subs', 'Markings & Signage', ''),
    ('yn',   'Accessible spaces marked with lines?', ''),
    ('yn',   'Access aisle marked to discourage parking in it?', ''),
    ('yn',   'Vertical ISA sign present?', ''),
    ('num',  'Height of ISA signage (inches)', '', '', 'in'),
    ('yn',   'ISA contrasts light-on-dark or dark-on-light?', ''),
    ('yn',   'Van accessible spaces marked with lines?', ''),
    ('yn',   '"Van Accessible" signs posted at van spaces?', ''),
    ('num',  'Height of Van Accessible ISA signage (inches)', '', '', 'in'),
    ('yn',   'Individuals with disabilities required to wheel/walk behind parked cars?', ''),
    ('yna',  'Detectable warnings at curb ramps in transit facilities?', ''),
] + NOTES_FIELD

ROUTES_FIELDS = [
    ('yn',   'At least one accessible route from arrival points to facility entrance?', ''),
    ('yna',  'Inaccessible entrances marked with directional signage?', ''),
    ('num',  'Width of public walkways (inches)', 'Min. 36", All standards', '', 'in'),
    ('num',  'Overhead clearance (inches)', 'Min. 80", All standards', '', 'in'),
    ('yna',  '60" passing space at intervals along route?', 'Required, 2010 ADA where route is less than 60" wide'),
    ('yn',   'Path of travel stable, firm, slip-resistant, and accessible?', ''),
    ('sel',  'Surface material', '', ['Concrete','Asphalt','Pavers','Other']),
    ('yn',   'Walkways continuous, not interrupted by steps or abrupt level changes?', ''),
    ('yn',   'Where paths cross driveways/parking lots, walkways blend to common level?', ''),
    ('yn',   'Curb cuts provided at driveways and parking lots?', ''),
    ('yn',   'Protruding objects present?', 'Objects must not protrude more than 4" into path between 27"-80" above floor'),
    ('wide', 'If yes, describe protruding objects', '', 'Description'),
] + NOTES_FIELD

CURB_RAMPS_FIELDS = [
    ('num',  'Width of ramp run, not including flared sides (inches)', 'Min. 36", All standards', '', 'in'),
    ('num',  'Rise measurement (inches)', '', '', 'in'),
    ('num',  'Length measurement (inches)', '', '', 'in'),
    ('num',  'Running slope (enter denominator, e.g. 12 for 1:12)', 'Max 1:12, All standards', '', '1:?'),
    ('num',  'Cross slope (enter denominator)', 'Max 1:48 (2%), All standards', '', '1:?'),
    ('text', 'Gutter slope measurement', '', 'ratio or %'),
    ('yn',   'Curb ramp has flared sides?', ''),
    ('num',  'Flared sides ratio (enter denominator)', 'Max 1:10, All standards', '', '1:?'),
    ('yn',   'Landing provided at top of curb ramp?', ''),
    ('pair', 'Landing area, Length × Width (inches)', 'Min. 36"×36" at top, All standards', ('L', 'W')),
    ('yn',   'Curb ramp placed diagonally at intersection?', ''),
    ('num',  'If yes, clear space measurement (inches)', '', '', 'in'),
    ('yn',   'Raised islands in crossings?', ''),
    ('pair', 'If yes, curb ramp separation, Length × Width (inches)', '', ('L', 'W')),
    ('yn',   'Surface stable, firm, and slip-resistant?', ''),
    ('yn',   'Detectable warnings provided?', ''),
    ('pair', 'Detectable warning dome size and spacing', '', ('Size', 'Spacing')),
    ('yn',   'Detectable warnings light-on-dark or dark-on-light contrast?', ''),
    ('yn',   'Gratings in walking surface?', ''),
    ('yn',   'Grating space exceeds 1/2" wide in one direction?', ''),
    ('wide', 'Any obstructions? (List)', '', 'Description'),
] + NOTES_FIELD

ENTRANCES_FIELDS = (
    SIG_FIELDS +
    [
        ('yn',   'Entrance marked with ISA (International Symbol of Accessibility)?', ''),
        ('yn',   'ISA contrasts light-on-dark or dark-on-light?', ''),
        ('yna',  'Inaccessible entrances have directional signage to accessible entrance?', ''),
        ('subs', 'Door Hardware & Dimensions', ''),
        ('num',  'Total number of accessible entrances', '', '', '#'),
        ('num',  'Total number of exits', '', '', '#'),
        ('yna',  'Ramp leading to entrance? (If yes, complete Ramps section)', ''),
    ] + DOOR_FIELDS +
    [
        ('yn',   'Floor level within 5 feet of door (direction door swings)?', ''),
        ('num',  'Doorway platform measurement beyond each side of door (inches)', '', '', 'in'),
        ('yn',   'Sharp inclines and abrupt level changes avoided at threshold?', ''),
        ('yn',   'Main entrance a fire door?', ''),
    ] + ALARM_FIELDS + NOTES_FIELD
)

STAIRS_FIELDS = [
    ('sel',  'Stair location', '', ['Interior','Exterior']),
    ('yn',   'Steps avoid abrupt nosing?', ''),
    ('num',  'Stair width (inches)', 'Min. 44" public stair, All standards', '', 'in'),
    ('yn',   'Handrails provided on both sides?', ''),
    ('num',  'Height of handrails (inches)', '30"-34", ANSI/UFAS | 34"-38", 1991 ADA/2010 ADA', '', 'in'),
    ('num',  'Length of handrails (inches)', '', '', 'in'),
    ('yn',   'Do steps have risers?', ''),
    ('num',  'Height of risers (inches)', 'Max 7", All standards', '', 'in'),
    ('num',  'Tread depth (inches)', 'Min. 11", All standards', '', 'in'),
    ('yn',   'Color/tonal contrast on stair nosings?', 'Recommended, 2010 ADA; required by many state standards'),
] + NOTES_FIELD

RAMPS_FIELDS = [
    ('wide', 'Ramp serves which area?', '', 'e.g. Main entrance, Building B north door'),
    ('num',  'Rise measurement, height of run (inches)', '', '', 'in'),
    ('num',  'Length measurement (inches)', '', '', 'in'),
    ('num',  'Maximum rise per run (inches)', 'Max 30" before intermediate landing required, All standards', '', 'in'),
    ('num',  'Running slope (enter denominator)', 'Max 1:12, All standards', '', '1:?'),
    ('num',  'Cross slope (enter denominator)', 'Max 1:48 (2%), All standards', '', '1:?'),
    ('yn',   'Surface stable, firm, and slip-resistant?', ''),
    ('yn',   'Accessible landing at top and bottom?', ''),
    ('pair', 'Top landing, Length × Width (inches)', 'Min. 60"×60", All standards', ('L','W')),
    ('pair', 'Bottom landing, Length × Width (inches)', 'Min. 60"×60", All standards', ('L','W')),
    ('yn',   'Intermediate landing between runs?', ''),
    ('yn',   'Ramp changes direction?', ''),
    ('pair', 'Intermediate / direction-change landing, Length × Width (inches)', 'Min. 60"×60", All standards', ('L','W')),
    ('yn',   'Landing subject to wet conditions?', ''),
    ('text', 'If yes, landing drainage slope (ratio)', '', 'ratio'),
    ('num',  'Clear width of ramp between handrails (inches)', 'Min. 36", All standards', '', 'in'),
    ('yn',   'Handrails on both sides?', ''),
    ('num',  'Height of handrails (inches)', '30"-34", ANSI/UFAS | 34"-38", 1991 ADA/2010 ADA', '', 'in'),
    ('yn',   'Handrails smooth and extend beyond top and bottom of ramp?', ''),
    ('yn',   'Handrails continuous?', ''),
    ('num',  'If not continuous, length of horizontal extensions (inches)', '', '', 'in'),
    ('num',  'Height of edge protection (inches)', '', '', 'in'),
    ('yn',   'Ramp has switchback or dogleg?', ''),
] + NOTES_FIELD

ELEVATORS_FIELDS = [
    ('yn',   'Building has multiple stories?', ''),
    ('yn',   'Elevator available and usable by individuals with physical disabilities?', ''),
    ('yn',   'Unassisted access to elevator?', ''),
    ('subs', 'Call Controls', ''),
    ('num',  'Height of call button, floor to center of call button panel (inches)', '15"-48", All standards', '', 'in'),
    ('num',  'Car call button size, diameter (inches)', 'Min. 3/4", 2010 ADA', '', 'in'),
    ('yn',   'Buttons easy to push or touch sensitive?', ''),
    ('yn',   'Visible and audible signals at each hoist way entrance?', ''),
    ('yn',   'Hall lantern provided?', ''),
    ('yn',   'Call buttons have visible/audible signals when call registered?', ''),
    ('yn',   'Audible signals: once for up, twice for down (or verbal annunciator)?', ''),
    ('subs', 'Hoist Way & Car', ''),
    ('yn',   'Floor designations, raised characters and Braille on both door jambs?', ''),
    ('num',  'Floor designation character height (inches)', 'Min. 2", All standards', '', 'in'),
    ('sel',  'Type of elevator entrance door', '', ['Center Sliding','Side Sliding','Other']),
    ('num',  'Width of elevator entrance door (inches)', 'Min. 32", All standards', '', 'in'),
    ('yn',   'Door automatically stops/reopens if obstructed, remains open min. 20 seconds?', ''),
    ('yn',   'Door reopening device sensitive to light touch or contact?', ''),
    ('pair', 'Car interior, Width × Depth (inches)', 'Min. 51"W × 51"D center door | 68"W × 51"D side door, 2010 ADA', ('W','D')),
    ('num',  'Width of car door (inches)', 'Min. 32", All standards', '', 'in'),
    ('yn',   'Emergency alarm/stop at bottom of panel?', ''),
    ('num',  'Height of emergency communications, floor to bottom of panel (inches)', 'Max 48", All standards', '', 'in'),
    ('yna',  'If not accessible, accessible elevator clearly identified with ISA?', ''),
] + NOTES_FIELD

LIFTS_FIELDS = [
    ('yn',   'Platform lift provides unassisted entry and exit?', ''),
    ('yn',   'Walkways have leveled platform to runway clearance?', ''),
    ('pair', 'Clear space for approach, Length × Width (inches)', '', ('L','W')),
    ('num',  'Height of controls, floor to middle of control panel (inches)', '15"-48", All standards', '', 'in'),
    ('yn',   'Low-energy, power-operated doors or gates?', ''),
    ('yn',   'Doors remain open minimum 20 seconds?', ''),
    ('num',  'Width of end doors and gates (inches)', 'Min. 32", All standards', '', 'in'),
    ('num',  'Width of side doors and gates (inches)', 'Min. 32", All standards', '', 'in'),
    ('pair', 'Lift interior, Width × Depth (inches)', 'Min. 32"W × 48"D, All standards', ('W','D')),
] + NOTES_FIELD

ROOMS_FIELDS = (
    [
        ('wide', 'Room Name / Type (e.g. Wellness Center, Career Center, Counseling Office)', '', 'Enter room name or type'),
        ('text', 'Room Number(s)', '', 'e.g. 101, B-12'),
        ('yn',   'Room/office on accessible route?', ''),
    ] + SIG_FIELDS +
    [
        ('subs', 'Doors', ''),
        ('num',  'Number of entrances and exits', '', '', '#'),
        ('yn',   'Fire door or exterior hinged door?', ''),
    ] + DOOR_FIELDS +
    [
        ('subs', 'Counters & Workstations', ''),
        ('yn',   'Accessible Sales/Services reception countertop?', 'Requires additional measurements'),
        ('num',  'Forward reach approach (inches)', 'Max 48", All standards', '', 'in'),
        ('num',  'Parallel approach (inches)', '', '', 'in'),
        ('yn',   'Accessible reception countertop?', ''),
        ('num',  'Height of reception countertop (inches)', 'Max 34", ANSI/UFAS | Max 36", 1991 ADA/2010 ADA', '', 'in'),
        ('pair', 'Countertop length × width (inches)', '', ('L','W')),
        ('triple', 'Knee clearance, H / W / D (inches)', 'Min. 27"H / 30"W / 19"D, All standards', ),
        ('yn',   'Supplies accessible by persons with disabilities?', ''),
    ] + ALARM_FIELDS + NOTES_FIELD
)

CTE_FIELDS = (
    [
        ('text', 'Room Number(s)', '', 'e.g. 204, Shop B'),
        ('text', 'Program / Course Name', '', 'e.g. Auto Tech, Culinary Arts'),
        ('yn',   'Classroom on accessible route?', ''),
    ] + SIG_FIELDS +
    [
        ('subs', 'Doors', ''),
        ('num',  'Number of entrances and exits', '', '', '#'),
    ] + DOOR_FIELDS +
    [
        ('subs', 'Seating & Workstations', ''),
        ('yn',   'Wheelchair seating at fixed tables/counters?', ''),
        ('num',  'Number of accessible seats', '', '', '#'),
        ('num',  'Height of fixed tables/countertops (inches)', '28"-34", All standards', '', 'in'),
        ('num',  'Aisle space width (inches)', 'Min. 36", All standards', '', 'in'),
        ('triple', 'Knee clearance, H / W / D (inches)', 'Min. 27"H / 30"W / 19"D', ),
        ('yn',   'Accessible desk/workstations have sufficient clearance?', ''),
        ('yna',  'If insufficient, alternative option available?', ''),
        ('yn',   'Supplies accessible by persons with disabilities?', ''),
    ] + ALARM_FIELDS + NOTES_FIELD
)

LABS_FIELDS = (
    [
        ('text', 'Room Number(s)', '', 'e.g. Shop 1, Lab B'),
        ('text', 'Program / Course Name', '', 'e.g. Welding, Biology Lab'),
        ('yn',   'Lab/Shop on accessible route?', ''),
    ] + SIG_FIELDS +
    [
        ('subs', 'Doors', ''),
        ('num',  'Number of entrances and exits', '', '', '#'),
    ] + DOOR_FIELDS +
    [
        ('subs', 'Workstations', ''),
        ('yn',   'Wheelchair seating at fixed tables/counters?', ''),
        ('num',  'Height of fixed tables/countertops (inches)', '28"-34", All standards', '', 'in'),
        ('num',  'Aisle space width (inches)', 'Min. 36", All standards', '', 'in'),
        ('triple', 'Knee clearance, H / W / D (inches)', 'Min. 27"H / 30"W / 19"D', ),
        ('yn',   'Accessible desk/workstations have sufficient clearance?', ''),
        ('yn',   'Supplies accessible by persons with disabilities?', ''),
        ('num',  'Height of soap dispenser (floor to operable part, inches)', 'Max 48", All standards', '', 'in'),
        ('num',  'Height of paper towel dispenser (floor to operable part, inches)', 'Max 48", All standards', '', 'in'),
    ] + ALARM_FIELDS + NOTES_FIELD
)

LOCKER_FIELDS = (
    [
        ('sel',  'Locker Room Type', '', ['Male','Female','Non-Gender Specific']),
        ('text', 'Building No.', '', ''),
        ('yn',   'Locker room on accessible route?', ''),
        ('subs', 'Signage & Entry Door', ''),
    ] + SIG_FIELDS[1:] +
    [
        ('num',  'Width of entrance doorway (inches)', 'Min. 32", All standards', '', 'in'),
        ('num',  'Height of door handle (inches)', 'Max 48", All standards', '', 'in'),
        ('num',  'Door opening force (pounds)', 'Max 5 lbs, 1991 ADA/UFAS/2010 ADA | Max 8.5 lbs, ANSI', '', 'lbs'),
        ('subs', 'Interior Maneuvering', ''),
        ('yn',   'Locker room has foyer?', ''),
        ('num',  'Measurement between doors in foyer (inches)', '', '', 'in'),
        ('num',  'Turning space, diameter (inches)', 'Min. 60", All standards', '', 'in'),
        ('yn',   'Doors swing into turning spaces?', ''),
        ('subs', 'Lockers & Amenities', ''),
        ('num',  'Number of accessible lockers', 'At least 5% of each type, 2010 ADA §225', '', '#'),
        ('num',  'Height of accessible locker operable parts (inches)', 'Max 48" forward reach, All standards', '', 'in'),
        ('yn',   'Mirror provided?', ''),
        ('num',  'Height of mirror, floor to reflecting edge (inches)', 'Max 40", All standards', '', 'in'),
        ('yn',   'Accessible benches provided?', ''),
        ('num',  'Height of benches, floor to top (inches)', '17"-19", All standards', '', 'in'),
        ('num',  'Depth of benches (inches)', 'Min. 20"-24", All standards', '', 'in'),
        ('yn',   'Back support on benches?', ''),
        ('num',  'Clear floor space adjacent to bench (inches)', 'Min. 30"×48", All standards', '', 'in'),
        ('subs', 'Shower Rooms', ''),
        ('yn',   'Showers in use?', ''),
        ('num',  'Number of showers', '', '', '#'),
        ('yn',   'Accessible shower provided?', ''),
        ('sel',  'Shower type', '', ['Transfer (36"×36")','Roll-in (60"×30" min)']),
        ('pair', 'Accessible shower, Width × Depth (inches)', 'Transfer: 36"×36" | Roll-in: 60"×30" min, 2010 ADA', ('W','D')),
        ('yn',   'Shower has curb?', ''),
        ('num',  'Height of curb in shower (inches)', 'Max 1/2", All standards (transfer shower)', '', 'in'),
        ('yn',   'Shower grab bars provided?', ''),
        ('wide', 'Location of shower grab bars', '', 'e.g. side wall, back wall'),
        ('num',  'Height of shower grab bars, floor to centerline (inches)', '33"-36", All standards', '', 'in'),
        ('num',  'Length of shower grab bars (inches)', '', '', 'in'),
        ('yn',   'Accessible shower seat provided?', ''),
        ('num',  'Height of shower seat, floor to top (inches)', '17"-19", All standards', '', 'in'),
        ('yn',   'Handheld shower/spray unit provided?', ''),
        ('num',  'Handheld shower hose length (inches)', 'Min. 59", 2010 ADA', '', 'in'),
        ('subs', 'Restroom Amenities', ''),
        ('yn',   'Separate restroom entrance?', ''),
        ('num',  'Width of restroom entrance (inches)', 'Min. 32", All standards', '', 'in'),
        ('num',  'Turning space, diameter (inches)', 'Min. 60", All standards', '', 'in'),
        ('yn',   'Hot water available?', ''),
        ('yna',  'Hot water pipes under sink covered?', ''),
        ('num',  'Height of sink, floor to rim (inches)', 'Max 34", All standards', '', 'in'),
        ('triple', 'Lavatory knee clearance, H / W / D (inches)', 'Min. 27"H / 30"W / 19"D', ),
        ('num',  'Height of soap dispenser (inches)', 'Max 48", All standards', '', 'in'),
        ('num',  'Height of mirror, floor to reflecting edge (inches)', 'Max 40", All standards', '', 'in'),
        ('num',  'Height of paper towel dispenser (inches)', 'Max 48", All standards', '', 'in'),
        ('subs', 'Accessible Stall', ''),
        ('yn',   'Designated accessible stall?', ''),
        ('num',  'Width of accessible stall (inches)', 'Min. 60", All standards', '', 'in'),
        ('num',  'Depth of accessible stall (inches)', 'Min. 56", All standards', '', 'in'),
        ('yn',   'Stall door swings out?', ''),
        ('num',  'Width of stall door (inches)', 'Min. 32", All standards', '', 'in'),
        ('yn',   'Grab bars in accessible stall?', ''),
        ('sel',  'Location of grab bars', '', ['Side wall','Rear wall','Both']),
        ('num',  'Height of grab bars, floor to centerline (inches)', '33"-36", All standards', '', 'in'),
        ('num',  'Height of toilet seat, floor to seat edge (inches)', '17"-19", All standards', '', 'in'),
        ('num',  'Height of flush controls, floor to controls (inches)', 'Max 44", All standards', '', 'in'),
        ('yn',   'Flush valve on transfer/open side of toilet?', ''),
        ('num',  'Height of toilet paper dispenser (inches)', '', '', 'in'),
        ('num',  'Urinal height, floor to rim (inches)', 'Max 17", All standards', '', 'in'),
        ('num',  'Width of space around urinal (inches)', 'Min. 30", All standards', '', 'in'),
    ] + NOTES_FIELD
)

RESTROOMS_FIELDS = (
    [
        ('sel',  'Restroom Type', '', ['Male','Female','Non-Gender Specific']),
        ('text', 'Building No.', '', ''),
        ('yn',   'Restroom on accessible route?', ''),
        ('num',  'Total number of accessible restrooms in building', '', '', '#'),
        ('yna',  'If no accessible restroom in building, directional signage to nearest?', ''),
        ('subs', 'Entry Door & Signage', ''),
        ('yn',   'Accessible door signage (ISA)?', ''),
        ('num',  'Height of accessible signage, floor to centerline (inches)', '60", All standards (57"-63" range)', '', 'in'),
        ('yn',   'ISA contrasts light-on-dark or dark-on-light?', ''),
        ('yn',   'Characters raised and accompanied by Grade II Braille?', ''),
        ('yn',   'Signage mounted adjacent to latch side of door?', ''),
        ('yn',   'Entrance doors swing out?', ''),
        ('num',  'Width of entrance doorway (inches)', 'Min. 32", All standards', '', 'in'),
        ('num',  'Height of door handle (inches)', 'Max 48", All standards', '', 'in'),
        ('num',  'Door opening force (pounds)', 'Max 5 lbs, 1991 ADA/UFAS/2010 ADA | Max 8.5 lbs, ANSI', '', 'lbs'),
        ('yn',   'Restroom has foyer?', ''),
        ('num',  'Measurement between doors in foyer (inches)', '', '', 'in'),
        ('subs', 'Interior Maneuvering', ''),
        ('num',  'Turning space, diameter (inches)', 'Min. 60" diameter OR T-turn, All standards', '', 'in'),
        ('subs', 'Amenities', ''),
        ('yn',   'Hot water available?', ''),
        ('yna',  'Hot water pipes under sink covered?', ''),
        ('num',  'Height of sink, floor to rim (inches)', 'Max 34", All standards', '', 'in'),
        ('yn',   'Sink usable by wheelchair users?', ''),
        ('triple', 'Lavatory knee clearance, H / W / D (inches)', 'Min. 27"H / 30"W / 19"D, All standards', ),
        ('num',  'Height of soap dispenser (floor to operable part, inches)', 'Max 48", All standards', '', 'in'),
        ('num',  'Height of mirror, floor to reflecting edge (inches)', 'Max 40", All standards', '', 'in'),
        ('num',  'Height of paper towel dispenser (floor to operable part, inches)', 'Max 48", All standards', '', 'in'),
        ('num',  'Height of automatic hand dryer (floor to operable part, inches)', 'Max 48", All standards', '', 'in'),
        ('yn',   'Racks, dispensers, disposal units not protruding into walkways?', ''),
        ('subs', 'Urinals', ''),
        ('num',  'Height of urinal, floor to rim (inches)', 'Max 17", All standards', '', 'in'),
        ('num',  'Width of space around urinal (inches)', 'Min. 30", All standards', '', 'in'),
        ('num',  'Depth of space around urinal (inches)', 'Min. 48", All standards', '', 'in'),
        ('subs', 'Accessible Stall', ''),
        ('yn',   'Designated accessible stall?', ''),
        ('num',  'Width of accessible stall (inches)', 'Min. 60", All standards', '', 'in'),
        ('num',  'Depth of accessible stall (inches)', 'Min. 56", All standards', '', 'in'),
        ('yn',   'Stall door swings out?', ''),
        ('num',  'Width of stall door (inches)', 'Min. 32", All standards', '', 'in'),
        ('yn',   'Grab bars in accessible stall?', ''),
        ('sel',  'Location of grab bars', '', ['Side wall','Rear wall','Both']),
        ('num',  'Height of grab bars, floor to centerline (inches)', '33"-36", All standards', '', 'in'),
        ('num',  'Length of grab bars (inches)', 'Side wall: min. 40" | Rear wall: min. 36"', '', 'in'),
        ('num',  'Depth of grab bars from wall (inches)', '', '', 'in'),
        ('num',  'Height of accessible toilet seat, floor to seat edge (inches)', '17"-19", All standards', '', 'in'),
        ('num',  'Height of flush controls, floor to controls (inches)', 'Max 44", All standards', '', 'in'),
        ('yn',   'Flush valve on transfer/open side of toilet?', ''),
        ('yn',   'Obstructions to reach toilet paper dispenser?', ''),
        ('yn',   'Toilet paper dispenser operable with one hand?', ''),
        ('yn',   'Dispenser allows continuous flow of paper?', ''),
        ('num',  'Height of toilet paper dispenser (inches)', '', '', 'in'),
        ('wide', 'Location of toilet tissue dispenser', '', 'e.g. side wall below grab bar'),
        ('yn',   'Coat hooks/shelves provided?', ''),
        ('num',  'Coat hook height, floor to hook (inches)', 'Max 48" forward reach, All standards', '', 'in'),
        ('sel',  'Type of reach for coat hook', '', ['Forward','Side']),
    ] + NOTES_FIELD
)

FOUNTAINS_FIELDS = [
    ('yn',   'Drinking fountains provided?', ''),
    ('num',  'Number of fountains', '', '', '#'),
    ('num',  'Number of accessible fountains', '', '', '#'),
    ('yn',   'At least 50% of fountains accessible (spout max 36" from floor)?', ''),
    ('yn',   'Fountain on accessible route?', ''),
    ('yn',   '30"×48" clear floor space in front?', ''),
    ('num',  'Height of spout, floor to spout outlet (inches)', 'Max 36", All standards', '', 'in'),
    ('num',  'Height of water flow, spout to top of arc (inches)', 'Min. 4", All standards', '', 'in'),
    ('yn',   'Water flow within 4" of spout towards front of fountain?', ''),
    ('yn',   'Up-front spout and control?', ''),
    ('sel',  'Controls mounted on front or side?', '', ['Front','Side']),
    ('yn',   'Controls operable with closed fist?', ''),
    ('yn',   'Operable parts require tight grasping/pinching/twisting?', ''),
    ('yn',   'Force to activate exceeds 5 lbs?', ''),
    ('yn',   'Fountain protrudes into corridors or traffic ways?', ''),
] + NOTES_FIELD

WATER_COOLERS_FIELDS = [
    ('yn',   'Bottle filler station available?', ''),
    ('yn',   'Station on accessible route?', ''),
    ('yn',   '30"×48" clear floor space in front?', ''),
    ('yn',   'Station hand operated?', ''),
    ('yn',   'Station sensor operated?', ''),
    ('sel',  'Controls mounted on front or side?', '', ['Front','Side']),
    ('num',  'Height of highest operable part (inches)', 'Max 48", All standards', '', 'in'),
    ('yn',   'Controls operable with closed fist?', ''),
    ('yn',   'Operable parts require tight grasping/pinching/twisting?', ''),
    ('yn',   'Force to activate exceeds 5 lbs?', ''),
] + NOTES_FIELD

STADIUM_FIELDS = [
    ('yn',   'Accessible route to area?', ''),
    ('sel',  'Route surface material', '', ['Asphalt','Concrete','Gravel','Dirt']),
    ('num',  'Total number of wheelchair designated seats/spaces', '', '', '#'),
    ('yn',   'Wheelchair designated seating provided?', ''),
    ('yn',   'Accessible route to wheelchair designated seats?', ''),
    ('yn',   'Companion seats provided with wheelchair spaces?', ''),
    ('yn',   'Ramp to wheelchair seating area? (If yes, complete Ramps section)', ''),
    ('num',  'Accessible route clear width (inches)', 'Min. 36", All standards', '', 'in'),
] + NOTES_FIELD

CAFETERIA_FIELDS = (
    SIG_FIELDS +
    [
        ('subs', 'Entrance & Doors', ''),
        ('num',  'Number of entrances and exits', '', '', '#'),
        ('yn',   'Cafeteria on accessible route?', ''),
        ('yna',  'Ramps leading to cafeteria? (If yes, complete Ramps section)', ''),
    ] + DOOR_FIELDS + ALARM_FIELDS +
    [
        ('subs', 'Indoor Fixed Seating', ''),
        ('num',  'Total number of fixed accessible seats indoors', '', '', '#'),
        ('yn',   'Seats and tables on accessible route?', ''),
        ('yn',   'Wheelchair seating space at fixed tables/counters?', ''),
        ('num',  'Height of fixed tables/countertops (inches)', '28"-34", All standards', '', 'in'),
        ('pair', 'Wheelchair space, Width × Depth (inches)', 'Min. 30"×48", All standards', ('W','D')),
        ('triple', 'Knee clearance, H / W / D (inches)', 'Min. 27"H / 30"W / 19"D', ),
        ('yn',   'Accessible route between food line, seating, and exit?', ''),
        ('subs', 'Outdoor Fixed Seating', ''),
        ('num',  'Total number of fixed accessible seats outdoors', '', '', '#'),
        ('yn',   'Seats and tables on accessible route?', ''),
        ('num',  'Height of outdoor fixed tables/countertops (inches)', '28"-34", All standards', '', 'in'),
        ('sel',  'Outdoor surface material', '', ['Concrete','Asphalt','Gravel','Other']),
        ('subs', 'Food Service Line(s)', ''),
        ('note', 'Measure all food service lines including main line, à la carte, and any student store lines.', ''),
        ('num',  'Total number of food lines', '', '', '#'),
        ('num',  'Food service line aisle width (inches)', 'Min. 36" clear, All standards', '', 'in'),
        ('num',  'Height of food line counter / tray slide (inches)', 'Max 34", ANSI/UFAS | Max 36", 1991 ADA/2010 ADA', '', 'in'),
        ('num',  'Height of sneeze guard / food display, lowest accessible item (inches)', 'Max 48" forward reach, All standards', '', 'in'),
        ('yn',   'Self-service items (utensils, condiments, trays) within reach range?', 'Max 48" forward / 54" side, ANSI/UFAS/1991 ADA | Max 48" side, 2010 ADA'),
        ('num',  'Height of highest self-service item, operable part (inches)', 'Max 48", All standards', '', 'in'),
        ('yn',   'Staff-assisted alternative available if self-service not accessible?', ''),
        ('subs', 'Checkout Counter(s)', ''),
        ('num',  'Total number of checkout aisles', '', '', '#'),
        ('num',  'Number of accessible checkout aisles', 'Min. 1 required, All standards', '', '#'),
        ('num',  'Checkout aisle width (inches)', 'Min. 36", All standards', '', 'in'),
        ('num',  'Height of accessible checkout counter (inches)', 'Max 34", ANSI/UFAS | Max 36", 1991 ADA/2010 ADA', '', 'in'),
        ('num',  'Width of lowered counter section (inches)', 'Min. 36", All standards', '', 'in'),
        ('triple', 'Knee clearance, H / W / D (inches)', 'Min. 27"H / 30"W / 19"D', ),
        ('sel',  'Approach type at checkout', '', ['Forward','Parallel']),
        ('num',  'Height of tableware and condiment areas (inches)', 'Max 48" forward reach, All standards', '', 'in'),
        ('subs', 'Tray Return Station', ''),
        ('yn',   'Tray return station provided?', ''),
        ('yna',  'Tray return on accessible route?', ''),
        ('num',  'Height of tray return slot (inches)', 'Max 48", All standards', '', 'in'),
        ('pair', 'Clear floor space in front, Width × Depth (inches)', 'Min. 30"×48", All standards', ('W','D')),
        ('subs', 'Self-Service Beverage Station', ''),
        ('yn',   'Self-service beverage station provided?', ''),
        ('yna',  'Station on accessible route?', ''),
        ('num',  'Height of highest operable part (inches)', 'Max 48", All standards', '', 'in'),
        ('pair', 'Clear floor space, Width × Depth (inches)', 'Min. 30"×48", All standards', ('W','D')),
        ('yn',   'Operable parts require tight grasping/pinching/twisting?', ''),
        ('yn',   'Force to activate exceeds 5 lbs?', ''),
        ('subs', 'Vending Machines', ''),
        ('yn',   'Vending machines provided?', ''),
        ('yna',  'Vending machines on accessible route?', ''),
        ('num',  'Height of highest operable part (inches)', 'Max 48" forward reach, All standards', '', 'in'),
        ('pair', 'Clear floor space in front, Width × Depth (inches)', 'Min. 30"×48", All standards', ('W','D')),
        ('yn',   'Force to activate exceeds 5 lbs?', ''),
    ] + NOTES_FIELD
)

LIBRARY_FIELDS = (
    SIG_FIELDS +
    [
        ('subs', 'Entrance & Doors', ''),
        ('num',  'Number of entrances and exits', '', '', '#'),
        ('yn',   'Library on accessible route?', ''),
        ('yna',  'Ramps leading to library? (If yes, complete Ramps section)', ''),
        ('num',  'Overhead clearance of accessible route (inches)', 'Min. 80", All standards', '', 'in'),
    ] + DOOR_FIELDS + ALARM_FIELDS +
    [
        ('subs', 'Seating & Workstations', ''),
        ('yn',   'Seating space provided for wheelchair users at fixed tables/counters?', ''),
        ('yn',   'Seats and tables on accessible route?', ''),
        ('num',  'Number of accessible seats provided', '', '', '#'),
        ('num',  'Height of fixed tables/countertops (inches)', '28"-34", All standards', '', 'in'),
        ('num',  'Aisle space width (inches)', 'Min. 36", All standards | Min. 60" passing space at intervals, 2010 ADA', '', 'in'),
        ('num',  'Turning space in reading areas, diameter (inches)', 'Min. 60", All standards', '', 'in'),
        ('triple', 'Knee clearance, H / W / D (inches)', 'Min. 27"H / 30"W / 19"D, All standards', ),
        ('yn',   'Accessible desk/workstations have sufficient clear space?', ''),
        ('yna',  'If insufficient, alternative option available for students?', ''),
        ('subs', 'Computer Workstations & Adaptive Technology', ''),
        ('yn',   'Computer workstations available?', ''),
        ('yn',   'Accessible computer workstation(s) provided?', 'Required if workstations provided, All standards'),
        ('num',  'Height of accessible workstation surface (inches)', '28"-34", All standards', '', 'in'),
        ('yn',   'Adaptive technology available? (screen reader, magnification, etc.)', ''),
        ('wide', 'If yes, describe adaptive technology available', '', 'e.g. screen reader, magnification software, large-print keyboard'),
        ('subs', 'Checkout Counter', ''),
        ('num',  'Number of accessible checkout areas', '', '', '#'),
        ('num',  'Height of checkout counter top, floor to surface (inches)', 'Max 34", ANSI/UFAS | Max 36", 1991 ADA/2010 ADA', '', 'in'),
        ('num',  'Width of lowered counter section (inches)', 'Min. 36", All standards', '', 'in'),
        ('num',  'Counter top depth (inches)', '', '', 'in'),
        ('triple', 'Knee clearance, H / W / D (inches)', 'Min. 27"H / 30"W / 19"D', ),
        ('num',  'Forward reach approach (inches)', 'Max 48", All standards', '', 'in'),
        ('yn',   'Parallel approach available?', ''),
        ('subs', 'Book Drop', ''),
        ('note', 'Measuring the Book Drop: measure from finished floor to the bottom edge of the drop slot opening, in inches. Measure clear floor space in front (W×D). Test force required to activate in lbs. Note approach type and interior/exterior.', ''),
        ('yn',   'Book drop provided?', ''),
        ('yna',  'Book drop on accessible route?', ''),
        ('num',  'Height of drop slot, floor to bottom edge of slot opening (inches)', 'Max 48" forward reach, All standards | Max 54" side, ANSI/UFAS/1991 ADA | Max 48" side, 2010 ADA', '', 'in'),
        ('pair', 'Clear floor space in front, Width × Depth (inches)', 'Min. 30"×48", All standards', ('W','D')),
        ('num',  'Force required to activate/open book drop (pounds)', 'Max 5 lbs, All standards', '', 'lbs'),
        ('yn',   'Operable parts require tight grasping, pinching, or twisting?', ''),
        ('sel',  'Approach type', '', ['Forward','Side']),
        ('sel',  'Book drop location', '', ['Exterior','Interior']),
        ('subs', 'Card Catalog & Stack Aisles', ''),
        ('num',  'Card catalog aisle width (inches)', 'Min. 36", All standards', '', 'in'),
        ('num',  'Card catalog operable part height (inches)', 'Max 48" forward reach, All standards', '', 'in'),
        ('num',  'Stack aisle width (inches)', 'Min. 36", All standards', '', 'in'),
        ('num',  'Overhead clearance in stack aisles (inches)', 'Min. 80", All standards', '', 'in'),
    ] + NOTES_FIELD
)

GYM_FIELDS = (
    [
        ('note', 'Complete for Gymnasium, Auditorium, Weight Room, or other large assembly space. If the campus has separate Gym and Auditorium buildings, note both in Additional Information or use separate printed copies.', ''),
        ('sel',  'Space Type', '', ['Gymnasium','Auditorium','Weight Room','Multi-Purpose Room','Black Box / Theater','Other']),
        ('text', 'Room Number(s)', '', 'e.g. Gym A, Aud 1'),
    ] + SIG_FIELDS +
    [
        ('subs', 'Entrance & Doors', ''),
        ('yn',   'Space on accessible route?', ''),
        ('num',  'Number of entrances and exits', '', '', '#'),
    ] + DOOR_FIELDS + ALARM_FIELDS +
    [
        ('subs', 'Wheelchair Seating', ''),
        ('num',  'Total number of all seating', '', '', '#'),
        ('num',  'Total number of wheelchair spaces', 'Required ratio varies by capacity, 2010 ADA §221', '', '#'),
        ('yn',   'Wheelchair seating in different locations throughout venue?', ''),
        ('yn',   'Companion seat provided with each wheelchair space?', ''),
        ('yn',   'Accessible route to wheelchair designated seats?', ''),
        ('yna',  'Ramp to seating area? (If yes, complete Ramps section)', ''),
        ('yn',   'Line of sight for wheelchair spaces comparable to ambulatory seating?', ''),
        ('pair', 'Wheelchair space dimensions, Width × Depth (inches)', 'Min. 36"W × 48"D, All standards', ('W','D')),
        ('yn',   'Can supplies be accessed by persons with disabilities?', ''),
        ('subs', 'Assistive Listening Systems', ''),
        ('yn',   'Audio amplification system provided?', ''),
        ('yn',   'Adequate number of assistive listening systems (ALS) provided?', 'Required where audible communications are integral to use of space, All standards'),
        ('sel',  'Type of assistive listening system', '', ['Induction Loop','FM','Infrared (IR)','Other']),
        ('num',  'Number of receivers available', '', '', '#'),
        ('num',  'If fixed audio amplification, distance from stage (feet)', '', '', 'ft'),
        ('yn',   'Signage indicating availability of assistive listening devices?', ''),
        ('subs', 'Fixed Seating / Tables (Weight Room / Other)', ''),
        ('yna',  'Wheelchair seating space at fixed tables/counters?', ''),
        ('num',  'Number of accessible seats provided', '', '', '#'),
        ('num',  'Height of fixed tables/countertops (inches)', '28"-34", All standards', '', 'in'),
        ('pair', 'Wheelchair space, Width × Depth (inches)', 'Min. 30"×48"', ('W','D')),
        ('triple', 'Knee clearance, H / W / D (inches)', 'Min. 27"H / 30"W / 19"D', ),
        ('yna',  'Accessible desk/workstations have sufficient clearance?', ''),
    ] + NOTES_FIELD
)

TELEPHONES_FIELDS = [
    ('yn',   'Public telephone provided? (If No, mark N/A and skip remaining fields)', ''),
    ('pair', 'Clear floor space, Width × Length (inches)', 'Min. 30"×48", All standards', ('W','L')),
    ('yn',   'Protruding objects?', ''),
    ('yn',   'Push button controls?', ''),
    ('num',  'Cord length, telephone to handset (inches)', 'Min. 29", All standards', '', 'in'),
    ('sel',  'Approach type', '', ['Parallel','Forward']),
    ('num',  'Distance from front edge of counter to face of telephone unit (inches)', '', '', 'in'),
    ('yn',   'Highest operable part within 15"-48"? (54" max if side approach)', ''),
    ('yn',   'Telephone hearing aid compatible?', ''),
    ('yn',   'Volume control provided?', ''),
    ('yna',  'If more than 4 pay telephones, TTY available?', ''),
    ('num',  'Height of TTY keypad touch surface (inches)', 'Max 48", All standards', '', 'in'),
    ('yna',  'Telephones with TTY equipped with shelf and electrical outlet?', ''),
] + NOTES_FIELD

# ── NEW v2 SECTIONS (thresholds per references/thresholds.md) ────────────────

LOADING_ZONES_FIELDS = [
    ('yna',  'Passenger loading zone provided? (If none, mark N/A and skip remaining fields)', ''),
    ('wide', 'Zone serves which area?', '', 'e.g. main entrance drop-off, bus loop'),
    ('yn',   'Access aisle adjacent and parallel to the vehicle pull-up space?', ''),
    ('pair', 'Access aisle, Width x Length', 'Min. 60 in. wide x 20 ft long, all standards', ('W in.','L ft')),
    ('num',  'Vertical clearance at the zone (inches)', 'Min. 114 in., 2010 ADA', '', 'in'),
    ('yn',   'Access aisle marked?', ''),
    ('yn',   'Aisle at same level as pull-up space (no curb between)?', ''),
    ('yna',  'Where a curb separates the aisle from the accessible route, curb ramp provided?', ''),
    ('num',  'Slope of pull-up space and aisle (enter denominator)', 'Max 1:48, all standards', '', '1:?'),
    ('yn',   'Zone connects to an accessible route to the entrance?', ''),
] + NOTES_FIELD

POOLS_FIELDS = [
    ('yna',  'Swimming pool or aquatic facility present? (If none, mark N/A and skip remaining fields)', ''),
    ('text', 'Pool type / use', '', 'e.g. competition pool, instruction pool, spa'),
    ('num',  'Pool wall length (linear feet)', 'Determines required number of accessible means of entry, 2010 ADA 242', '', 'ft'),
    ('yn',   'Accessible means of entry provided?', 'Pools with less than 300 linear ft of wall: min. 1 (pool lift or sloped entry); 300 linear ft or more: min. 2, at least one a pool lift or sloped entry, 2010 ADA 242'),
    ('sel',  'Primary accessible entry type', '', ['Pool lift','Sloped entry','Transfer wall','Transfer system','Accessible pool stairs']),
    ('yn',   'Pool lift available and operable during all open hours?', ''),
    ('num',  'Lift seat height above deck (inches)', '16 in. to 19 in., 2010 ADA', '', 'in'),
    ('yn',   'Accessible route to the pool deck?', ''),
    ('yn',   'Deck surface stable, firm, and slip-resistant?', ''),
    ('yn',   'Accessible route connects deck, entries, and locker/shower facilities?', ''),
] + NOTES_FIELD

EGRESS_FIELDS = [
    ('yn',   'Accessible means of egress provided from all occupied areas?', ''),
    ('num',  'Number of accessible means of egress', 'Same number as required exits, up to code limits', '', '#'),
    ('yna',  'Areas of Rescue Assistance / areas of refuge designated (multi-story buildings)?', ''),
    ('num',  'Number of wheelchair spaces in each area of rescue assistance', 'Min. 30 in. x 48 in. per space', '', '#'),
    ('yn',   'Two-way communication system at each area of rescue assistance?', ''),
    ('yn',   'Communication system tested and operational?', ''),
    ('yn',   'Signage identifying areas of rescue assistance (with ISA)?', ''),
    ('yn',   'Illuminated exit signage on accessible egress routes?', ''),
    ('yn',   'Egress routes free of steps and barriers, or an evacuation alternative provided?', ''),
    ('yn',   'Evacuation procedures address mobility impairments?', ''),
] + NOTES_FIELD

MEDICAL_FIELDS = [
    ('text', 'Room name / number', '', "e.g. Nurse's Office, Health Office, First-Aid Room"),
    ('yn',   'Medical / first-aid area on accessible route?', ''),
    ('num',  'Width of doorway (inches)', 'Min. 32 in., all standards', '', 'in'),
    ('num',  'Height of door handle (inches)', 'Max 48 in., all standards', '', 'in'),
    ('num',  'Door opening force (pounds)', 'Max 5 lbs, 1991 ADA / UFAS / 2010 ADA | Max 8.5 lbs, ANSI', '', 'lbs'),
    ('num',  'Turning space, diameter (inches)', 'Min. 60 in., all standards', '', 'in'),
    ('num',  'Height of examination cot / bed surface (inches)', '17 in. to 19 in. supports transfer', '', 'in'),
    ('yn',   '30 in. x 48 in. clear floor space alongside the cot / bed?', ''),
    ('yn',   'Sink accessible?', 'Rim max 34 in.; knee clearance min. 27 in. H / 30 in. W / 19 in. D'),
    ('num',  'Height of supplies and operable parts (inches)', 'Max 48 in. forward reach, all standards', '', 'in'),
    ('yn',   'Privacy accommodations (curtained or partitioned area) accessible?', ''),
    ('yn',   'Identification signage with raised characters and Braille?', ''),
] + NOTES_FIELD

PORTABLES_FIELDS = [
    ('text', 'Portable / building number', '', 'e.g. P-3'),
    ('num',  'Year installed', '', '', 'yyyy'),
    ('yn',   'Portable on accessible route?', ''),
    ('yn',   'Ramp provided to entrance? (If yes, also complete the Ramps section)', ''),
    ('num',  'Ramp running slope (enter denominator)', 'Max 1:12, all standards', '', '1:?'),
    ('pair', 'Landing at door, Length x Width (inches)', 'Min. 60 in. x 60 in., all standards', ('L','W')),
    ('num',  'Width of doorway (inches)', 'Min. 32 in., all standards', '', 'in'),
    ('num',  'Height of door handle (inches)', 'Max 48 in., all standards', '', 'in'),
    ('yn',   'Door openable without grasping or twisting wrist?', ''),
    ('num',  'Door opening force (pounds)', 'Max 5 lbs, 1991 ADA / UFAS / 2010 ADA | Max 8.5 lbs, ANSI', '', 'lbs'),
    ('num',  'Threshold height (inches)', 'Max 1/2 in., all standards', '', 'in'),
    ('yn',   'Handrails on both sides of ramp where rise exceeds 6 in.?', ''),
    ('num',  'Height of handrails (inches)', '30 in. to 34 in., ANSI / UFAS | 34 in. to 38 in., 1991 ADA / 2010 ADA', '', 'in'),
    ('yn',   'Fire / emergency egress from portable accessible?', ''),
    ('yn',   'Visual and audible alarms provided?', ''),
] + NOTES_FIELD

# Program Access Interview, 14 staff-interview questions (textarea responses).
# These now live in a SEPARATE standalone document (document B), never in the
# site guide.
PA_INTERVIEW_QUESTIONS = [
    ('subs', 'ADA Coordination & Compliance', ''),
    ('area', '1. Who is the designated ADA/Section 504 Coordinator for this facility, and how are accessibility complaints or accommodation requests handled?', '', 'Staff response...'),
    ('area', '2. Has the facility undergone a formal self-evaluation under Section 504? If so, when and what were the findings?', '', 'Staff response...'),
    ('area', '3. Does the facility have a current Transition Plan? When was it last updated and what items remain outstanding?', '', 'Staff response...'),
    ('subs', 'Maintenance of Accessible Features', ''),
    ('area', '4. Are accessible features (ramps, lifts, elevators, accessible restrooms, automatic door openers) on the regular maintenance schedule? How frequently inspected?', '', 'Staff response...'),
    ('area', '5. Are maintenance staff trained to identify and report accessibility barriers? Describe the training.', '', 'Staff response...'),
    ('area', '6. Are accessible parking spaces regularly monitored to prevent unauthorized use? What is the enforcement process?', '', 'Staff response...'),
    ('area', '7. How are temporary accessibility barriers (construction, equipment storage, event setup) managed to maintain accessible routes during disruptions?', '', 'Staff response...'),
    ('subs', 'Program Access & Relocation', ''),
    ('area', '8. Are there areas where a program, class, or activity has been relocated to provide program access? If so, where, why, and how was this communicated?', '', 'Staff response...'),
    ('area', '9. What is the process for notifying students, parents, and staff of accessible routes, features, and accommodations?', '', 'Staff response...'),
    ('area', '10. If a student with a disability needed access to an inaccessible area, what is the process for providing equivalent program access?', '', 'Staff response...'),
    ('subs', 'Emergency Egress & Safety', ''),
    ('area', '11. How are emergency evacuation procedures adapted for mobility impairments? Are Areas of Rescue Assistance designated, equipped, and included in drills?', '', 'Staff response...'),
    ('area', '12. Are two-way communication systems at Areas of Rescue Assistance tested regularly? When last tested?', '', 'Staff response...'),
    ('subs', 'Planned Modifications', ''),
    ('area', '13. Are any accessibility modifications currently planned, funded, approved, or in progress? Describe scope, timeline, and funding source.', '', 'Staff response...'),
    ('area', '14. Have any accessibility modifications been completed in the past five years? Describe scope and dates.', '', 'Staff response...'),
    ('subs', 'Reviewer Observations', ''),
    ('area', 'Program Reviewer observations and notes from interview', '', 'Observations...'),
]

# ─────────────────────────────────────────────────────────────────────────────
# SECTION DEFINITIONS, order matches the HTML sidebar
# (name, kind 'multi'|'single', fields)
# ─────────────────────────────────────────────────────────────────────────────
SECTIONS = [
    ('Accessible Parking',                   'multi',  PARKING_FIELDS),
    ('Passenger Loading Zones',              'multi',  LOADING_ZONES_FIELDS),
    ('Accessible Routes / Walkways',         'multi',  ROUTES_FIELDS),
    ('Curb Ramps',                           'multi',  CURB_RAMPS_FIELDS),
    ('Stadium / Field',                      'single', STADIUM_FIELDS),
    ('Entrances, Doors, and Gates',          'multi',  ENTRANCES_FIELDS),
    ('Stairways and Steps',                  'multi',  STAIRS_FIELDS),
    ('Ramps',                                'multi',  RAMPS_FIELDS),
    ('Elevators',                            'multi',  ELEVATORS_FIELDS),
    ('Lifts',                                'multi',  LIFTS_FIELDS),
    ('Accessible Means of Egress / Areas of Rescue Assistance', 'multi', EGRESS_FIELDS),
    ('Rooms & Offices',                      'multi',  ROOMS_FIELDS),
    ("Medical and First-Aid Areas (including Nurse's Office)", 'multi', MEDICAL_FIELDS),
    ('Cafeteria',                            'single', CAFETERIA_FIELDS),
    ('Library',                              'single', LIBRARY_FIELDS),
    ('CTE Classroom(s)',                     'multi',  CTE_FIELDS),
    ('Labs / Shops',                         'multi',  LABS_FIELDS),
    ('Portable / Temporary Classrooms',      'multi',  PORTABLES_FIELDS),
    ('Gymnasium / Auditorium / Weight Room', 'single', GYM_FIELDS),
    ('Swimming Pools / Aquatic Areas',       'single', POOLS_FIELDS),
    ('Locker Rooms',                         'multi',  LOCKER_FIELDS),
    ('Restrooms (Male / Female / Non-Gender Specific)', 'multi', RESTROOMS_FIELDS),
    ('Accessible Drinking Fountains',        'multi',  FOUNTAINS_FIELDS),
    ('Water Coolers & Bottle Fillers',       'multi',  WATER_COOLERS_FIELDS),
    ('Public Telephones',                    'single', TELEPHONES_FIELDS),
]

# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY: keep the source doc's front-matter (cover, accessibility-standards
# listing, tips-for-measuring with the 8 screenclip images) and back-matter
# (Glossary of Terms), REPLACE every section body in between with new
# schema-driven content. A thresholds-only Standards Guide table sits near
# the top; nothing internal-process-related appears anywhere in document A.
# ─────────────────────────────────────────────────────────────────────────────

# Identify cut points in the base doc (the previously generated guide):
#   - START of replaceable content: the "Accessible Parking" section-title
#     paragraph (the base doc's only Heading 1 is the Glossary, so the first
#     section title is located by its exact text). The page-break-only
#     paragraph immediately before it is removed too.
#   - END of replaceable content (start of preserved tail): the
#     "Glossary of Terms" Heading 1.
para_list = list(doc.paragraphs)
first_cut_idx = None
glossary_idx = None
for i, p in enumerate(para_list):
    if first_cut_idx is None and p.text.strip() == "Accessible Parking":
        first_cut_idx = i
    if (p.style.name or "").strip() == "Heading 1" and "Glossary" in p.text:
        glossary_idx = i
        break

if first_cut_idx is None or glossary_idx is None:
    raise RuntimeError(f"Could not locate cut points: first_cut={first_cut_idx}, glossary={glossary_idx}")
print(f"Cut points: first section title = p[{first_cut_idx}] '{para_list[first_cut_idx].text[:60]}'  |  Glossary = p[{glossary_idx}]")

# We will:
#   1. Walk body children. Identify the XML element of the first section
#      title (plus its preceding page-break paragraph) and the Glossary H1.
#   2. Remove everything between them (inclusive of the section title,
#      exclusive of Glossary).
#   3. Insert NEW content (Standards Guide thresholds table + sections)
#      at the cut position.

first_h1_xml = para_list[first_cut_idx]._p
glossary_xml = para_list[glossary_idx]._p

# Include the page-break-only paragraph right before the first section title.
_prev = first_h1_xml.getprevious()
if _prev is not None and _prev.tag == qn("w:p") and not "".join(t.text or "" for t in _prev.iter(qn("w:t"))).strip():
    first_h1_xml = _prev

# Walk body children to collect those to remove
to_remove = []
collecting = False
for child in list(body):
    if child is first_h1_xml:
        collecting = True
    if collecting:
        if child is glossary_xml:
            break
        to_remove.append(child)

print(f"Removing {len(to_remove)} body elements between first H1 and Glossary...")
for ch in to_remove:
    body.remove(ch)

# Helper to insert a list of elements BEFORE the Glossary
def insert_before_glossary(elements):
    parent = glossary_xml.getparent()
    idx = list(parent).index(glossary_xml)
    for e in elements:
        parent.insert(idx, e)
        idx += 1

def _heading_run(text, half_points="32"):
    rn = OxmlElement("w:r")
    rp = OxmlElement("w:rPr"); rp.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), half_points); rp.append(sz)
    cc = OxmlElement("w:color"); cc.set(qn("w:val"), NAVY); rp.append(cc)
    rn.append(rp)
    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text; rn.append(t)
    return rn

def _page_break_para():
    pbp = OxmlElement("w:p"); pbr = OxmlElement("w:r")
    pbk = OxmlElement("w:br"); pbk.set(qn("w:type"), "page"); pbr.append(pbk); pbp.append(pbr)
    return pbp

# ─────────────────────────────────────────────────────────────────────────────
# BUILD: Standards Guide quick-reference table
# ─────────────────────────────────────────────────────────────────────────────
def standards_guide_block():
    out = []
    # Page break
    pbp = OxmlElement("w:p"); pbr = OxmlElement("w:r")
    pbk = OxmlElement("w:br"); pbk.set(qn("w:type"), "page"); pbr.append(pbk); pbp.append(pbr)
    out.append(pbp)
    # Heading
    rn = OxmlElement("w:r")
    rp = OxmlElement("w:rPr"); rp.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "32"); rp.append(sz)
    cc = OxmlElement("w:color"); cc.set(qn("w:val"), NAVY); rp.append(cc)
    rn.append(rp)
    t = OxmlElement("w:t"); t.text = "Standards Guide, Key Measurement Thresholds"; rn.append(t)
    out.append(_para(after=80, runs_extra=[rn]))
    out.append(_para("Quick reference of key measurement thresholds. These same notations also appear beside each measurement field in the sections that follow.", size=18, color="5A6A7A", italic=True, after=120))

    # Build comparison table
    cols = [3300, 1500, 1500, 1500, 1500]
    rows = []

    def th(text, w):
        rn = OxmlElement("w:r")
        rp = OxmlElement("w:rPr"); rp.append(OxmlElement("w:b"))
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "18"); rp.append(sz)
        cc = OxmlElement("w:color"); cc.set(qn("w:val"), GOLD); rp.append(cc)
        rn.append(rp)
        t = OxmlElement("w:t"); t.text = text.upper(); rn.append(t)
        return _cell(w, [_para(after=0, runs_extra=[rn])], shade=NAVY)

    rows.append([
        th("Element", cols[0]),
        th("ANSI", cols[1]),
        th("UFAS", cols[2]),
        th("1991 ADA", cols[3]),
        th("2010 ADA", cols[4]),
    ])

    def row(label, *vals):
        cells = [_cell(cols[0], [_para(label, size=18, bold=True, color=NAVY, after=0)], shade=CREAM)]
        for i, v in enumerate(vals):
            cells.append(_cell(cols[1+i], [_para(v, size=18, after=0, color="1A2332")]))
        return cells

    THRESHOLDS = [
        ("Door clear width (min)",                     '32"',          '32"',          '32"',          '32"'),
        ("Door opening force, interior (max)",        '8.5 lbs',      '5 lbs',        '5 lbs',        '5 lbs'),
        ("Ramp max running slope",                     '1:12',         '1:12',         '1:12',         '1:12'),
        ("Handrail height",                            '30"-34"',      '30"-34"',      '34"-38"',      '34"-38"'),
        ("Accessible parking width (min)",             '96"',          '96"',          '96"',          '96"'),
        ("Van accessible, space + aisle",             '96"+96" aisle','96"+96" aisle','96"+96" aisle','132" OR 96"+60" aisle'),
        ("Van vertical clearance (min)",               '98"',          '98"',          '98"',          '98"'),
        ("Counter / table height (max)",               '34"',          '34"',          '36"',          '36"'),
        ("Forward reach (max)",                        '48"',          '48"',          '48"',          '48"'),
        ("Side reach (max)",                           '54"',          '54"',          '54"',          '48"'),
        ("Accessible route width (min)",               '36"',          '36"',          '36"',          '36"'),
        ("Overhead clearance (min)",                   '80"',          '80"',          '80"',          '80"'),
        ("Turning space (min)",                        '60" dia.',     '60" dia.',     '60" dia.',     '60" dia. or T-turn'),
        ("Grab bar height, toilet",                   '33"-36"',      '33"-36"',      '33"-36"',      '33"-36"'),
        ("Toilet seat height",                         '17"-19"',      '17"-19"',      '17"-19"',      '17"-19"'),
        ("Drinking fountain spout (max)",              '36"',          '36"',          '36"',          '36"'),
        ("Lavatory rim (max)",                         '34"',          '34"',          '34"',          '34"'),
        ("Mirror bottom edge (max)",                   '40"',          '40"',          '40"',          '40"'),
    ]
    for r in THRESHOLDS:
        rows.append(row(*r))
    out.append(_tbl(cols, rows))
    return out
# ─────────────────────────────────────────────────────────────────────────────
# ASSEMBLE DOCUMENT A and insert before the Glossary
# ─────────────────────────────────────────────────────────────────────────────
all_new = []

# 1) Standards Guide quick-reference (thresholds only, no process content)
all_new.extend(standards_guide_block())

# 2) Per-section bodies (multi-location sections wrapped in a Repeating
#    Section Content Control; no numbered instance labels)
for name, kind, fields in SECTIONS:
    all_new.extend(render_section(name, fields, multi_instance=(kind == "multi")))

# Insert before the Glossary
print(f"Inserting {len(all_new)} top-level elements before Glossary...")
insert_before_glossary(all_new)

# Ensure the document root declares the w15 namespace so Word recognizes the
# repeating-section markers we just embedded (v5 mechanism). The base doc
# already declares it when built by build_optimized_guide.py v5; this guards
# against a base doc that does not.
doc_root = body.getparent()  # <w:document>
if doc_root.nsmap.get("w15") != W15_NS:
    new_nsmap = dict(doc_root.nsmap)
    new_nsmap["w15"] = W15_NS
    new_root = etree.Element(doc_root.tag, attrib=dict(doc_root.attrib), nsmap=new_nsmap)
    for child in list(doc_root):
        new_root.append(child)
    mc_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    ign_attr = f"{{{mc_ns}}}Ignorable"
    cur = new_root.get(ign_attr, "")
    if "w15" not in cur.split():
        new_root.set(ign_attr, (cur + " w15").strip())
    doc.part._element = new_root
    print("Added w15 namespace declaration to document.xml root.")

# ─────────────────────────────────────────────────────────────────────────────
# Finalize helpers: Arial 12 pt base style, author metadata, em-dash scrub
# ─────────────────────────────────────────────────────────────────────────────
EMDASH = "\u2014"  # em dash (U+2014), banned by house style; never write the literal glyph
def _scrub(el):
    fixed = 0
    for node in el.iter():
        if node.text and EMDASH in node.text:
            node.text = node.text.replace(" " + EMDASH + " ", ": ").replace(EMDASH, "-"); fixed += 1
        if node.tail and EMDASH in node.tail:
            node.tail = node.tail.replace(" " + EMDASH + " ", ": ").replace(EMDASH, "-"); fixed += 1
        for k, v in list(node.attrib.items()):
            if EMDASH in v:
                node.set(k, v.replace(" " + EMDASH + " ", ": ").replace(EMDASH, "-")); fixed += 1
    return fixed

def finalize(document):
    """Arial 12 pt Normal style, author metadata, em-dash scrub (all parts)."""
    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
    normal.element.rPr.rFonts.set(qn("w:cs"), "Arial")

    # Authored by Murjani McTier, never Claude.
    cp = document.core_properties
    cp.author = DOC_AUTHOR
    cp.last_modified_by = DOC_AUTHOR

    n_fixed = _scrub(document.element)
    for rel in document.part.rels.values():
        if rel.reltype.endswith("/header") or rel.reltype.endswith("/footer"):
            n_fixed += _scrub(rel.target_part.element)
    print(f"Em-dash scrub: cleaned {n_fixed} node(s).")

# ─────────────────────────────────────────────────────────────────────────────
# Save DOCUMENT A
# ─────────────────────────────────────────────────────────────────────────────
finalize(doc)
doc.save(OUT_GUIDE)
print(f"\nSaved -> {OUT_GUIDE}")
print(f"  File size: {os.path.getsize(OUT_GUIDE)/1024:.1f} KB")

# ─────────────────────────────────────────────────────────────────────────────
# DOCUMENT B: Program Access, Facilities and Maintenance & Operations
# Staff Interview (standalone document)
# ─────────────────────────────────────────────────────────────────────────────
def build_interview_doc():
    idoc = Document()
    ibody = idoc.element.body
    sectPr = ibody.find(qn("w:sectPr"))

    def add(el):
        if sectPr is not None:
            sectPr.addprevious(el)
        else:
            ibody.append(el)

    # Title
    rn = OxmlElement("w:r")
    rp = OxmlElement("w:rPr"); rp.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "32"); rp.append(sz)
    cc = OxmlElement("w:color"); cc.set(qn("w:val"), NAVY); rp.append(cc)
    rn.append(rp)
    t = OxmlElement("w:t")
    t.text = "Program Access, Facilities and Maintenance & Operations Staff Interview"
    rn.append(t)
    add(_para(after=120, runs_extra=[rn]))

    add(_tbl([9700], [[_cell(9700, [
        _para("These questions are asked of the school site's Facilities and Maintenance & Operations staff by the Program Reviewer during the on-site visit. Program Access compliance is determined observationally and through staff responses, not through physical measurement.",
              size=18, color="1A4A6A", after=40),
    ], shade="E8F4FD")]]))
    add(_para(after=80))

    # Header block: school name, date, staff interviewed (name / title), interviewer
    HDR, RSP = 3200, 6500
    hdr_rows = [
        ("School Name:",              make_plain_text("Click to enter the school name")),
        ("Date of Interview:",        make_date_picker("Click to pick the interview date")),
        ("Staff Interviewed, Name:",  make_plain_text("Click to enter name")),
        ("Staff Interviewed, Title:", make_plain_text("Click to enter title")),
        ("Interviewer:",              make_plain_text("Click to enter interviewer name")),
    ]
    for lbl, ctrl in hdr_rows:
        lp = _para(lbl, bold=True, size=24, color=NAVY, after=0)
        vp = _para(after=0); vp.append(ctrl)
        add(_tbl([HDR, RSP], [[_cell(HDR, [lp], shade=CREAM), _cell(RSP, [vp])]]))
    add(_para(after=120))

    # The 14 interview questions, each with a Response field and Notes space
    def _bold_label_para(prefix, placeholder, after):
        p = _para(after=after)
        lb = OxmlElement("w:r")
        lbp = OxmlElement("w:rPr"); lbp.append(OxmlElement("w:b"))
        lbsz = OxmlElement("w:sz"); lbsz.set(qn("w:val"), "22"); lbp.append(lbsz)
        lb.append(lbp)
        lbt = OxmlElement("w:t"); lbt.set(qn("xml:space"), "preserve"); lbt.text = prefix
        lb.append(lbt)
        p.append(lb)
        p.append(make_plain_text(placeholder))
        return p

    for field in PA_INTERVIEW_QUESTIONS:
        kind, label = field[0], field[1]
        if kind == 'subs':
            hr = OxmlElement("w:r")
            hrp = OxmlElement("w:rPr"); hrp.append(OxmlElement("w:b"))
            hsz = OxmlElement("w:sz"); hsz.set(qn("w:val"), "26"); hrp.append(hsz)
            hcc = OxmlElement("w:color"); hcc.set(qn("w:val"), NAVY); hrp.append(hcc)
            hr.append(hrp)
            ht = OxmlElement("w:t"); ht.text = label.upper(); hr.append(ht)
            add(_para(after=80, runs_extra=[hr]))
            continue
        q_para = _para(label, bold=True, size=24, color=NAVY, after=40)
        resp_para = _bold_label_para("Response:  ", "Click to enter the staff response", 40)
        notes_para = _bold_label_para("Notes:  ", "Click to enter notes", 0)
        add(_tbl([9700], [
            [_cell(9700, [q_para], shade=CREAM)],
            [_cell(9700, [resp_para])],
            [_cell(9700, [notes_para])],
        ]))
        add(_para(after=120))

    return idoc

interview_doc = build_interview_doc()
finalize(interview_doc)
interview_doc.save(OUT_INTERVIEW)
print(f"Saved -> {OUT_INTERVIEW}")
print(f"  File size: {os.path.getsize(OUT_INTERVIEW)/1024:.1f} KB")

# Remove the superseded v2 output if it is still present
if os.path.exists(OLD_OUT):
    os.remove(OLD_OUT)
    print(f"Removed superseded output: {OLD_OUT}")

print(f"  Final SDT ID counter: {_next_id}")
