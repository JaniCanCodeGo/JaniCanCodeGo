"""Build the optimized BLANK Facilities Review Guide — v5.

v5 changes from v4:
  * Each multi-instance section's measurement table is wrapped in a
    native Word Repeating Section Content Control (w15:repeatingSection).
    Click anywhere in the table → Word shows a "+" button at the right
    margin → click "+" to insert another blank copy of the entire
    table for an additional set of locations.
  * Works in Word 2013+ (Windows), Word for Mac 2016+, Word for the web.
  * No macros, no security warning, no .docm.

v4 fundamentals retained:
  * NO era dropdown — reviewer determines the standard.
  * Plain text fields throughout — no date pickers.
  * 12pt Arial, black & white.
  * Wide tables: label column + 4 location columns per copy.

PRESERVED VERBATIM from source (non-negotiable):
  * Cover page
  * Accessibility Standards listing
  * Tips for Measuring + all 8 screenclip images
  * Glossary of Terms at the end
"""
import os, re
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Inches
from lxml import etree

# w15 namespace — Microsoft Word 2012 wordml extensions (for repeatingSection)
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"

# Sequential SDT ID counter — guarantees no collisions
_next_id = 1000
def next_id():
    global _next_id
    _next_id += 1
    return str(_next_id)

SRC = "/root/.claude/uploads/f318c6f4-c9c8-4a12-bcd0-bfd812a66767/60dc1d22-BLANK_Facilities__Review_Guide.docx"
OUT = "outputs/BLANK_Facilities_Review_Guide_optimized.docx"

SZ_BODY  = "24"   # 12pt
SZ_FIELD = "24"
SZ_HEAD  = "28"   # 14pt
SZ_TITLE = "32"   # 16pt
SZ_NOTE  = "22"   # 11pt

BLACK = "000000"
GRAY  = "595959"
NUM_LOCATIONS = 4

# ── Low-level XML ──────────────────────────────────────────────────────────
def _run(text, *, bold=False, italic=False, size=SZ_BODY, color=BLACK):
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    if bold: rpr.append(OxmlElement("w:b"))
    if italic: rpr.append(OxmlElement("w:i"))
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), size); rpr.append(sz)
    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"), "Arial"); rf.set(qn("w:hAnsi"), "Arial")
    rpr.append(rf)
    cc = OxmlElement("w:color"); cc.set(qn("w:val"), color); rpr.append(cc)
    r.append(rpr)
    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text
    r.append(t)
    return r

def _para(*runs, after=120, align=None):
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    spc = OxmlElement("w:spacing"); spc.set(qn("w:after"), str(after)); ppr.append(spc)
    if align:
        ja = OxmlElement("w:jc"); ja.set(qn("w:val"), align); ppr.append(ja)
    p.append(ppr)
    for r in runs:
        p.append(r)
    return p

def _t(text, *, bold=False, italic=False, size=SZ_BODY, color=BLACK, after=120, align=None):
    return _para(_run(text, bold=bold, italic=italic, size=size, color=color),
                 after=after, align=align)

def _empty_para():
    return _para(after=0)

def _page_break():
    p = OxmlElement("w:p"); r = OxmlElement("w:r")
    br = OxmlElement("w:br"); br.set(qn("w:type"), "page"); r.append(br); p.append(r)
    return p

def _cell(width_dxa, paras, *, gridSpan=None):
    tc = OxmlElement("w:tc")
    tcPr = OxmlElement("w:tcPr")
    w_ = OxmlElement("w:tcW"); w_.set(qn("w:w"), str(width_dxa)); w_.set(qn("w:type"), "dxa")
    tcPr.append(w_)
    if gridSpan:
        gs = OxmlElement("w:gridSpan"); gs.set(qn("w:val"), str(gridSpan)); tcPr.append(gs)
    va = OxmlElement("w:vAlign"); va.set(qn("w:val"), "top"); tcPr.append(va)
    tc.append(tcPr)
    for p in paras:
        tc.append(p)
    if not paras or paras[-1].tag != qn("w:p"):
        tc.append(_empty_para())
    return tc

def _table(grid_cols, rows):
    tbl = OxmlElement("w:tbl")
    tblPr = OxmlElement("w:tblPr")
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:w"), "5000"); tblW.set(qn("w:type"), "pct")
    tblPr.append(tblW)
    tb = OxmlElement("w:tblBorders")
    for side in ("top","left","bottom","right","insideH","insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), BLACK)
        tb.append(b)
    tblPr.append(tb)
    tblLook = OxmlElement("w:tblLook"); tblLook.set(qn("w:val"), "04A0"); tblPr.append(tblLook)
    tbl.append(tblPr)
    grid = OxmlElement("w:tblGrid")
    for w in grid_cols:
        gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), str(w)); grid.append(gc)
    tbl.append(grid)
    for cells in rows:
        if hasattr(cells, "tag"):
            tbl.append(cells); continue
        tr = OxmlElement("w:tr")
        # keep rows from breaking across pages mid-cell
        trPr = OxmlElement("w:trPr")
        cant = OxmlElement("w:cantSplit"); trPr.append(cant)
        tr.append(trPr)
        for c in cells:
            tr.append(c)
        tbl.append(tr)
    return tbl

# ── Schema field kinds ───────────────────────────────────────────────────────
# Each field is a tuple: (kind, label, *extras)
# kind:
#   'q'    = measurement / open response — blank cells in each location column
#   'yn'   = yes/no row — each location cell renders "☐ Yes   ☐ No"
#   'yna'  = yes/no/N/A row
#   'hdr'  = subsection header — label spans all columns
#   'note' = italic note — text spans all columns

# ── Table row builders ───────────────────────────────────────────────────────
def _yn_runs(with_na=False):
    """Build runs for a Yes/No (or Yes/No/N/A) cell — 12pt Arial inline."""
    items = ["☐ Yes", "☐ No"] + (["☐ N/A"] if with_na else [])
    runs = []
    for i, it in enumerate(items):
        runs.append(_run(it, size=SZ_BODY))
        if i < len(items)-1:
            runs.append(_run("   ", size=SZ_BODY))
    return runs

def _section_row(field, col_widths, num_locs):
    """Render one field as a single table row across (label + N location cells)."""
    kind = field[0]
    label = field[1]
    hint = field[2] if len(field) > 2 else ""

    label_col = col_widths[0]
    loc_widths = col_widths[1:]
    total_locs_width = sum(loc_widths)

    if kind == 'hdr':
        # Subsection heading — single cell spanning all columns
        run_b = _run(label.upper(), bold=True, size=SZ_HEAD, color=BLACK)
        tr = OxmlElement("w:tr")
        trPr = OxmlElement("w:trPr"); trPr.append(OxmlElement("w:cantSplit")); tr.append(trPr)
        tr.append(_cell(label_col + total_locs_width,
                        [_para(run_b, after=40)], gridSpan=1+num_locs))
        return tr

    if kind == 'note':
        run_i = _run(label, italic=True, size=SZ_NOTE, color=GRAY)
        tr = OxmlElement("w:tr")
        trPr = OxmlElement("w:trPr"); trPr.append(OxmlElement("w:cantSplit")); tr.append(trPr)
        tr.append(_cell(label_col + total_locs_width,
                        [_para(run_i, after=40)], gridSpan=1+num_locs))
        return tr

    # Build label cell with optional hint underneath
    label_paras = [_para(_run(label, size=SZ_FIELD), after=0)]
    if hint:
        label_paras.append(_para(_run(hint, italic=True, size=SZ_NOTE, color=GRAY), after=0))

    tr = OxmlElement("w:tr")
    trPr = OxmlElement("w:trPr"); trPr.append(OxmlElement("w:cantSplit")); tr.append(trPr)
    tr.append(_cell(label_col, label_paras))

    for w in loc_widths:
        if kind == 'yn':
            tr.append(_cell(w, [_para(*_yn_runs(False), after=0)]))
        elif kind == 'yna':
            tr.append(_cell(w, [_para(*_yn_runs(True), after=0)]))
        else:
            tr.append(_cell(w, [_empty_para()]))
    return tr

def render_section_table(fields, num_locs=NUM_LOCATIONS):
    """Build a wide multi-location table for one section.
       First row = blank header row (column headers); next rows = identifier rows
       (Location, Year Built, ADA Modification, Identify Modifications); then fields."""
    # Column widths (page is ~9700 dxa wide in normal margins)
    # Label column wider; location columns equal
    LBL = 3800
    LOC = (9700 - LBL) // num_locs
    col_widths = [LBL] + [LOC] * num_locs

    rows = []

    # Identifier rows — match the original guide's per-location header fields
    identifier_rows = [
        ('q', 'Location / Building No.:', ''),
        ('q', 'Year Built:', ''),
        ('q', 'ADA Modification Date(s):', '"N/A" if none'),
        ('q', 'Identify ADA Modifications:', '"N/A" if none'),
    ]
    for f in identifier_rows:
        rows.append(_section_row(f, col_widths, num_locs))

    # Field rows
    for f in fields:
        rows.append(_section_row(f, col_widths, num_locs))

    return _table(col_widths, rows)

def render_single_table(fields):
    """Single-instance section: 2 columns wide (label + value). Larger value column."""
    LBL = 3800
    VAL = 9700 - LBL
    col_widths = [LBL, VAL]
    rows = []

    identifier_rows = [
        ('q', 'Location / Building No.:', ''),
        ('q', 'Year Built:', ''),
        ('q', 'ADA Modification Date(s):', '"N/A" if none'),
        ('q', 'Identify ADA Modifications:', '"N/A" if none'),
    ]
    for f in identifier_rows:
        rows.append(_section_row(f, col_widths, 1))
    for f in fields:
        rows.append(_section_row(f, col_widths, 1))
    return _table(col_widths, rows)

# ── Section heading + reviewer note ─────────────────────────────────────────
def render_section_heading(title, multi=True):
    """Section title page-break + heading + standard-determination note."""
    out = []
    out.append(_page_break())
    out.append(_t(title, bold=True, size=SZ_TITLE, after=120))
    out.append(_t(
        "Reviewer-determined standard: the Program Reviewer determines the applicable accessibility standard (ANSI A117.1, UFAS, 1991 ADA, or 2010 ADA) based on Year Built and ADA Modification Date(s) below. The LEA records measurements only.",
        italic=True, size=SZ_NOTE, color=GRAY, after=80))
    if multi:
        out.append(_t(
            f"This section records up to {NUM_LOCATIONS} locations side-by-side. For more locations, copy the section template from the back of this guide.",
            italic=True, size=SZ_NOTE, color=GRAY, after=120))
    return out

# ─────────────────────────────────────────────────────────────────────────────
# SECTION SCHEMAS — concise, derived from the HTML
# Each field: (kind, label, hint)
#   kind:  q / yn / yna / hdr / note
# ─────────────────────────────────────────────────────────────────────────────

PARKING = [
    ('hdr', 'Space Counts', ''),
    ('q',  'Total number of parking spaces', ''),
    ('q',  'Total number of accessible spaces', ''),
    ('q',  'Total number of van accessible spaces', ''),
    ('yn', 'At least 1 of every 6 accessible spaces is van accessible?', ''),
    ('yn', 'Accessible spaces located closest to nearest accessible route?', ''),
    ('hdr', 'Dimensions (inches)', ''),
    ('q',  'Width of accessible parking space', 'Min. 96"'),
    ('q',  'Width of access aisle', 'Min. 60" std | Min. 96" van — ANSI/UFAS/1991 | Min. 60" — 2010 if space is 132"'),
    ('q',  'Width of van accessible space', 'Min. 132" OR 96"+60" aisle — 2010 | Min. 96"+96" — earlier'),
    ('q',  'Width of van access aisle', ''),
    ('q',  'Vertical clearance — van parking / access aisles', 'Min. 98"'),
    ('yn', 'Ground surface slope less than 1:48?', ''),
    ('yn', 'Access aisle slope less than 1:48?', ''),
    ('hdr', 'Markings & Signage', ''),
    ('yn', 'Accessible spaces marked with lines?', ''),
    ('yn', 'Access aisle marked to discourage parking in it?', ''),
    ('yn', 'Vertical ISA sign present?', ''),
    ('q',  'Height of ISA signage (inches)', ''),
    ('yn', 'ISA contrasts light-on-dark or dark-on-light?', ''),
    ('yn', 'Van accessible spaces marked with lines?', ''),
    ('yn', '"Van Accessible" signs posted at van spaces?', ''),
    ('q',  'Height of Van Accessible ISA signage (inches)', ''),
    ('yn', 'Individuals required to wheel/walk behind parked cars?', ''),
    ('yna','Detectable warnings at curb ramps in transit facilities?', ''),
    ('q',  'Additional Information / Observations', ''),
]

ROUTES = [
    ('yn', 'At least one accessible route from arrival points to entrance?', ''),
    ('yna','Inaccessible entrances marked with directional signage?', ''),
    ('q',  'Width of public walkways (inches)', 'Min. 36"'),
    ('q',  'Overhead clearance (inches)', 'Min. 80"'),
    ('yna','60" passing space at intervals along route?', '2010 ADA if route is <60" wide'),
    ('yn', 'Path stable, firm, slip-resistant, accessible?', ''),
    ('q',  'Surface material (concrete / asphalt / pavers / other)', ''),
    ('yn', 'Walkways continuous, no steps or abrupt level changes?', ''),
    ('yn', 'Walkways blend to common level at driveways / lots?', ''),
    ('yn', 'Curb cuts at driveways and parking lots?', ''),
    ('yn', 'Protruding objects present?', '>4" projection between 27"–80" AFF'),
    ('q',  'If yes, describe protruding objects', ''),
    ('q',  'Additional Information / Observations', ''),
]

CURB_RAMPS = [
    ('q',  'Width of ramp run (not flared sides) — inches', 'Min. 36"'),
    ('q',  'Rise measurement (inches)', ''),
    ('q',  'Length measurement (inches)', ''),
    ('q',  'Running slope — denominator (e.g. 12 = 1:12)', 'Max 1:12'),
    ('q',  'Cross slope — denominator', 'Max 1:48 (2%)'),
    ('q',  'Gutter slope measurement', ''),
    ('yn', 'Curb ramp has flared sides?', ''),
    ('q',  'Flared sides ratio — denominator', 'Max 1:10'),
    ('yn', 'Landing provided at top of curb ramp?', ''),
    ('q',  'Top landing — Length × Width (inches)', 'Min. 36"×36"'),
    ('yn', 'Curb ramp placed diagonally at intersection?', ''),
    ('q',  'If yes, clear space measurement (inches)', ''),
    ('yn', 'Raised islands in crossings?', ''),
    ('q',  'If yes, curb ramp separation — L × W (inches)', ''),
    ('yn', 'Surface stable, firm, slip-resistant?', ''),
    ('yn', 'Detectable warnings provided?', ''),
    ('q',  'Detectable warning dome size + spacing', ''),
    ('yn', 'Detectable warnings light/dark contrast?', ''),
    ('yn', 'Gratings in walking surface?', ''),
    ('yn', 'Grating space >1/2" in one direction?', ''),
    ('q',  'Obstructions present? (describe)', ''),
    ('q',  'Additional Information / Observations', ''),
]

STADIUM = [
    ('yn', 'Accessible route to area?', ''),
    ('q',  'Route surface material (asphalt / concrete / gravel / dirt)', ''),
    ('q',  'Total number of wheelchair designated seats/spaces', ''),
    ('yn', 'Wheelchair designated seating provided?', ''),
    ('yn', 'Accessible route to wheelchair designated seats?', ''),
    ('yn', 'Companion seats provided with wheelchair spaces?', ''),
    ('yn', 'Ramp to wheelchair seating area?', 'If yes, complete Ramps section'),
    ('q',  'Accessible route clear width (inches)', 'Min. 36"'),
    ('q',  'Additional Information / Observations', ''),
]

ENTRANCES = [
    ('hdr', 'Signage', ''),
    ('yn', 'Identification signage provided?', ''),
    ('q',  'Height of signage — floor to centerline (inches)', '60" centerline; 57"–63"'),
    ('q',  'Height of characters on signage (inches)', ''),
    ('yn', 'Signs mounted adjacent to latch side of door?', ''),
    ('yna','At double doors, signs on nearest adjacent wall?', ''),
    ('yn', 'Protruding objects within 3" of sign?', ''),
    ('yn', 'Non-glare finish on signs?', ''),
    ('yn', 'Characters raised + Grade II Braille?', ''),
    ('yn', 'Entrance marked with ISA?', ''),
    ('yn', 'ISA contrasts light-on-dark or dark-on-light?', ''),
    ('yna','Inaccessible entrances have directional signage to accessible entrance?', ''),
    ('hdr', 'Door Hardware & Dimensions', ''),
    ('q',  'Total number of accessible entrances', ''),
    ('q',  'Total number of exits', ''),
    ('yna','Ramp leading to entrance?', 'If yes, complete Ramps section'),
    ('q',  'Width of doorway(s) (inches)', 'Min. 32"'),
    ('q',  'Height of door handle (inches)', 'Max 48"'),
    ('yn', 'Door openable without grasping or twisting wrist?', ''),
    ('q',  'If not, type of hardware on door', ''),
    ('yna','Automatic door opener provided?', ''),
    ('yna','Door closer takes at least 3 seconds to close?', ''),
    ('q',  'Door opening force (pounds)', 'Max 5 lbs — 1991/UFAS/2010 | Max 8.5 — ANSI'),
    ('q',  'Threshold height — Exterior (inches)', 'Max 1/2"'),
    ('q',  'Threshold height — Interior (inches)', 'Max 1/2"'),
    ('yna','Carpeting or mats 1/2" or less?', ''),
    ('yn', 'Floor level within 5 feet of door (direction door swings)?', ''),
    ('q',  'Doorway platform measurement beyond each side (inches)', ''),
    ('yn', 'Sharp inclines and abrupt level changes avoided at threshold?', ''),
    ('yn', 'Main entrance a fire door?', ''),
    ('hdr', 'Alarms', ''),
    ('yn', 'Visual alarm provided?', ''),
    ('yn', 'Light flashes clear or normal white?', ''),
    ('yn', 'Audible alarm provided?', ''),
    ('yn', 'Alarm exceeds prevailing sound level in room?', ''),
    ('q',  'Additional Information / Observations', ''),
]

STAIRS = [
    ('q',  'Stair location (interior / exterior)', ''),
    ('yn', 'Steps avoid abrupt nosing?', ''),
    ('q',  'Stair width (inches)', 'Min. 44" public stair'),
    ('yn', 'Handrails provided on both sides?', ''),
    ('q',  'Height of handrails (inches)', '30"–34" — ANSI/UFAS | 34"–38" — 1991/2010 ADA'),
    ('q',  'Length of handrails (inches)', ''),
    ('yn', 'Do steps have risers?', ''),
    ('q',  'Height of risers (inches)', 'Max 7"'),
    ('q',  'Tread depth (inches)', 'Min. 11"'),
    ('yn', 'Color/tonal contrast on stair nosings?', ''),
    ('q',  'Additional Information / Observations', ''),
]

RAMPS = [
    ('q',  'Ramp serves which area?', ''),
    ('q',  'Rise measurement (inches)', ''),
    ('q',  'Length measurement (inches)', ''),
    ('q',  'Max rise per run (inches)', 'Max 30" before intermediate landing'),
    ('q',  'Running slope — denominator', 'Max 1:12'),
    ('q',  'Cross slope — denominator', 'Max 1:48'),
    ('yn', 'Surface stable, firm, slip-resistant?', ''),
    ('yn', 'Accessible landing at top and bottom?', ''),
    ('q',  'Top landing — L × W (inches)', 'Min. 60"×60"'),
    ('q',  'Bottom landing — L × W (inches)', 'Min. 60"×60"'),
    ('yn', 'Intermediate landing between runs?', ''),
    ('yn', 'Ramp changes direction?', ''),
    ('q',  'Intermediate / direction-change landing — L × W (inches)', 'Min. 60"×60"'),
    ('yn', 'Landing subject to wet conditions?', ''),
    ('q',  'If yes, landing drainage slope (ratio)', ''),
    ('q',  'Clear width of ramp between handrails (inches)', 'Min. 36"'),
    ('yn', 'Handrails on both sides?', ''),
    ('q',  'Height of handrails (inches)', '30"–34" — ANSI/UFAS | 34"–38" — 1991/2010 ADA'),
    ('yn', 'Handrails smooth + extend beyond top and bottom?', ''),
    ('yn', 'Handrails continuous?', ''),
    ('q',  'If not continuous, length of horizontal extensions (inches)', ''),
    ('q',  'Height of edge protection (inches)', ''),
    ('yn', 'Ramp has switchback or dogleg?', ''),
    ('q',  'Additional Information / Observations', ''),
]

ELEVATORS = [
    ('yn', 'Building has multiple stories?', ''),
    ('yn', 'Elevator available and usable by individuals with physical disabilities?', ''),
    ('yn', 'Unassisted access to elevator?', ''),
    ('hdr', 'Call Controls', ''),
    ('q',  'Height of call button — floor to panel center (inches)', '15"–48"'),
    ('q',  'Car call button size — diameter (inches)', 'Min. 3/4" — 2010 ADA'),
    ('yn', 'Buttons easy to push or touch-sensitive?', ''),
    ('yn', 'Visible and audible signals at each hoistway entrance?', ''),
    ('yn', 'Hall lantern provided?', ''),
    ('yn', 'Call buttons have visible/audible signals when registered?', ''),
    ('yn', 'Audible signals: once for up, twice for down (or verbal annunciator)?', ''),
    ('hdr', 'Hoistway & Car', ''),
    ('yn', 'Floor designations — raised characters + Braille on both door jambs?', ''),
    ('q',  'Floor designation character height (inches)', 'Min. 2"'),
    ('q',  'Type of entrance door (center / side / other)', ''),
    ('q',  'Width of elevator entrance door (inches)', 'Min. 32"'),
    ('yn', 'Door reopens if obstructed — remains open min. 20 seconds?', ''),
    ('yn', 'Door reopening device sensitive to light touch or contact?', ''),
    ('q',  'Car interior — W × D (inches)', 'Min. 51"W × 51"D ctr | 68"W × 51"D side — 2010'),
    ('q',  'Width of car door (inches)', 'Min. 32"'),
    ('yn', 'Emergency alarm/stop at bottom of panel?', ''),
    ('q',  'Height of emergency communications — floor to bottom of panel (inches)', 'Max 48"'),
    ('yna','If not accessible, accessible elevator clearly identified with ISA?', ''),
    ('q',  'Additional Information / Observations', ''),
]

LIFTS = [
    ('yn', 'Platform lift provides unassisted entry and exit?', ''),
    ('yn', 'Walkways have leveled platform-to-runway clearance?', ''),
    ('q',  'Clear space for approach — L × W (inches)', ''),
    ('q',  'Height of controls — floor to panel center (inches)', '15"–48"'),
    ('yn', 'Low-energy, power-operated doors or gates?', ''),
    ('yn', 'Doors remain open minimum 20 seconds?', ''),
    ('q',  'Width of end doors and gates (inches)', 'Min. 32"'),
    ('q',  'Width of side doors and gates (inches)', 'Min. 32"'),
    ('q',  'Lift interior — W × D (inches)', 'Min. 32"W × 48"D'),
    ('q',  'Additional Information / Observations', ''),
]

ROOMS = [
    ('q',  'Room name / type', 'e.g. Wellness Center, Counseling Office'),
    ('q',  'Room number(s)', ''),
    ('yn', 'Room/office on accessible route?', ''),
    ('hdr', 'Signage', ''),
    ('yn', 'Identification signage provided?', ''),
    ('q',  'Height of signage — floor to centerline (inches)', '60" centerline; 57"–63"'),
    ('q',  'Height of characters on signage (inches)', ''),
    ('yn', 'Signs mounted adjacent to latch side of door?', ''),
    ('yna','Double doors — signs on nearest adjacent wall?', ''),
    ('yn', 'Protruding objects within 3" of sign?', ''),
    ('yn', 'Non-glare finish on signs?', ''),
    ('yn', 'Characters raised + Grade II Braille?', ''),
    ('hdr', 'Doors', ''),
    ('q',  'Number of entrances and exits', ''),
    ('yn', 'Fire door or exterior hinged door?', ''),
    ('q',  'Width of doorway(s) (inches)', 'Min. 32"'),
    ('q',  'Height of door handle (inches)', 'Max 48"'),
    ('yn', 'Door openable without grasping or twisting wrist?', ''),
    ('q',  'If not, type of hardware on door', ''),
    ('yna','Automatic door opener provided?', ''),
    ('yna','Door closer takes at least 3 seconds to close?', ''),
    ('q',  'Door opening force (pounds)', 'Max 5 lbs — 1991/UFAS/2010 | Max 8.5 — ANSI'),
    ('q',  'Threshold — Exterior / Interior (inches)', 'Max 1/2"'),
    ('yna','Carpeting/mats 1/2" or less?', ''),
    ('hdr', 'Counters & Workstations', ''),
    ('yn', 'Accessible Sales/Services reception countertop?', ''),
    ('q',  'Forward reach approach (inches)', 'Max 48"'),
    ('q',  'Parallel approach (inches)', ''),
    ('yn', 'Accessible reception countertop?', ''),
    ('q',  'Height of reception countertop (inches)', 'Max 34" — ANSI/UFAS | Max 36" — 1991/2010'),
    ('q',  'Countertop length × width (inches)', ''),
    ('q',  'Knee clearance — H / W / D (inches)', 'Min. 27"H / 30"W / 19"D'),
    ('yn', 'Supplies accessible by persons with disabilities?', ''),
    ('hdr', 'Alarms', ''),
    ('yn', 'Visual alarm provided?', ''),
    ('yn', 'Light flashes clear or normal white?', ''),
    ('yn', 'Audible alarm provided?', ''),
    ('yn', 'Alarm exceeds prevailing sound level in room?', ''),
    ('q',  'Additional Information / Observations', ''),
]

CTE = [
    ('q',  'Room number(s)', ''),
    ('q',  'Program / Course name', ''),
    ('yn', 'Classroom on accessible route?', ''),
    ('hdr', 'Signage', ''),
    ('yn', 'Identification signage provided?', ''),
    ('q',  'Height of signage — floor to centerline (inches)', '60" centerline; 57"–63"'),
    ('q',  'Height of characters on signage (inches)', ''),
    ('yn', 'Signs mounted adjacent to latch side of door?', ''),
    ('yna','Double doors — signs on nearest adjacent wall?', ''),
    ('yn', 'Protruding objects within 3" of sign?', ''),
    ('yn', 'Non-glare finish on signs?', ''),
    ('yn', 'Characters raised + Grade II Braille?', ''),
    ('hdr', 'Doors', ''),
    ('q',  'Number of entrances and exits', ''),
    ('q',  'Width of doorway(s) (inches)', 'Min. 32"'),
    ('q',  'Height of door handle (inches)', 'Max 48"'),
    ('yn', 'Door openable without grasping or twisting wrist?', ''),
    ('q',  'If not, type of hardware on door', ''),
    ('yna','Automatic door opener provided?', ''),
    ('yna','Door closer takes at least 3 seconds to close?', ''),
    ('q',  'Door opening force (pounds)', 'Max 5 lbs — 1991/UFAS/2010 | Max 8.5 — ANSI'),
    ('q',  'Threshold — Exterior / Interior (inches)', 'Max 1/2"'),
    ('yna','Carpeting/mats 1/2" or less?', ''),
    ('hdr', 'Seating & Workstations', ''),
    ('yn', 'Wheelchair seating at fixed tables/counters?', ''),
    ('q',  'Number of accessible seats', ''),
    ('q',  'Height of fixed tables/countertops (inches)', '28"–34"'),
    ('q',  'Aisle space width (inches)', 'Min. 36"'),
    ('q',  'Knee clearance — H / W / D (inches)', 'Min. 27"H / 30"W / 19"D'),
    ('yn', 'Accessible desk/workstations have sufficient clearance?', ''),
    ('yna','If insufficient, alternative option available?', ''),
    ('yn', 'Supplies accessible by persons with disabilities?', ''),
    ('hdr', 'Alarms', ''),
    ('yn', 'Visual alarm provided?', ''),
    ('yn', 'Light flashes clear or normal white?', ''),
    ('yn', 'Audible alarm provided?', ''),
    ('yn', 'Alarm exceeds prevailing sound level in room?', ''),
    ('q',  'Additional Information / Observations', ''),
]

LABS = [
    ('q',  'Room number(s)', ''),
    ('q',  'Program / Course name', ''),
    ('yn', 'Lab/Shop on accessible route?', ''),
    ('hdr', 'Signage', ''),
    ('yn', 'Identification signage provided?', ''),
    ('q',  'Height of signage — floor to centerline (inches)', '60" centerline; 57"–63"'),
    ('q',  'Height of characters on signage (inches)', ''),
    ('yn', 'Signs mounted adjacent to latch side of door?', ''),
    ('yna','Double doors — signs on nearest adjacent wall?', ''),
    ('yn', 'Protruding objects within 3" of sign?', ''),
    ('yn', 'Non-glare finish on signs?', ''),
    ('yn', 'Characters raised + Grade II Braille?', ''),
    ('hdr', 'Doors', ''),
    ('q',  'Number of entrances and exits', ''),
    ('q',  'Width of doorway(s) (inches)', 'Min. 32"'),
    ('q',  'Height of door handle (inches)', 'Max 48"'),
    ('yn', 'Door openable without grasping or twisting wrist?', ''),
    ('q',  'If not, type of hardware on door', ''),
    ('yna','Automatic door opener provided?', ''),
    ('yna','Door closer takes at least 3 seconds to close?', ''),
    ('q',  'Door opening force (pounds)', 'Max 5 lbs — 1991/UFAS/2010 | Max 8.5 — ANSI'),
    ('q',  'Threshold — Exterior / Interior (inches)', 'Max 1/2"'),
    ('yna','Carpeting/mats 1/2" or less?', ''),
    ('hdr', 'Workstations', ''),
    ('yn', 'Wheelchair seating at fixed tables/counters?', ''),
    ('q',  'Height of fixed tables/countertops (inches)', '28"–34"'),
    ('q',  'Aisle space width (inches)', 'Min. 36"'),
    ('q',  'Knee clearance — H / W / D (inches)', 'Min. 27"H / 30"W / 19"D'),
    ('yn', 'Accessible desk/workstations have sufficient clearance?', ''),
    ('yn', 'Supplies accessible by persons with disabilities?', ''),
    ('q',  'Height of soap dispenser (inches)', 'Max 48"'),
    ('q',  'Height of paper towel dispenser (inches)', 'Max 48"'),
    ('hdr', 'Alarms', ''),
    ('yn', 'Visual alarm provided?', ''),
    ('yn', 'Light flashes clear or normal white?', ''),
    ('yn', 'Audible alarm provided?', ''),
    ('yn', 'Alarm exceeds prevailing sound level in room?', ''),
    ('q',  'Additional Information / Observations', ''),
]

LOCKER = [
    ('q',  'Locker room type (Male / Female / Non-Gender Specific)', ''),
    ('yn', 'Locker room on accessible route?', ''),
    ('hdr', 'Signage & Entry Door', ''),
    ('yn', 'Identification signage provided?', ''),
    ('q',  'Height of signage — floor to centerline (inches)', '60" centerline; 57"–63"'),
    ('q',  'Height of characters on signage (inches)', ''),
    ('yn', 'Signs mounted adjacent to latch side of door?', ''),
    ('yn', 'Characters raised + Grade II Braille?', ''),
    ('q',  'Width of entrance doorway (inches)', 'Min. 32"'),
    ('q',  'Height of door handle (inches)', 'Max 48"'),
    ('q',  'Door opening force (pounds)', 'Max 5 lbs — 1991/UFAS/2010 | Max 8.5 — ANSI'),
    ('hdr', 'Interior Maneuvering', ''),
    ('yn', 'Locker room has foyer?', ''),
    ('q',  'Measurement between doors in foyer (inches)', ''),
    ('q',  'Turning space — diameter (inches)', 'Min. 60"'),
    ('yn', 'Doors swing into turning spaces?', ''),
    ('hdr', 'Lockers & Amenities', ''),
    ('q',  'Number of accessible lockers', '≥ 5% of each type — 2010 ADA §225'),
    ('q',  'Height of accessible locker operable parts (inches)', 'Max 48" forward reach'),
    ('yn', 'Mirror provided?', ''),
    ('q',  'Height of mirror — floor to reflecting edge (inches)', 'Max 40"'),
    ('yn', 'Accessible benches provided?', ''),
    ('q',  'Height of benches — floor to top (inches)', '17"–19"'),
    ('q',  'Depth of benches (inches)', 'Min. 20"–24"'),
    ('yn', 'Back support on benches?', ''),
    ('q',  'Clear floor space adjacent to bench (inches)', 'Min. 30"×48"'),
    ('hdr', 'Shower Rooms', ''),
    ('yn', 'Showers in use?', ''),
    ('q',  'Number of showers', ''),
    ('yn', 'Accessible shower provided?', ''),
    ('q',  'Shower type (Transfer 36"×36" / Roll-in 60"×30" min)', ''),
    ('q',  'Accessible shower — W × D (inches)', ''),
    ('yn', 'Shower has curb?', ''),
    ('q',  'Height of curb in shower (inches)', 'Max 1/2" — transfer shower'),
    ('yn', 'Shower grab bars provided?', ''),
    ('q',  'Location of shower grab bars', ''),
    ('q',  'Height of shower grab bars — floor to centerline (inches)', '33"–36"'),
    ('q',  'Length of shower grab bars (inches)', ''),
    ('yn', 'Accessible shower seat provided?', ''),
    ('q',  'Height of shower seat — floor to top (inches)', '17"–19"'),
    ('yn', 'Handheld shower/spray unit provided?', ''),
    ('q',  'Handheld shower hose length (inches)', 'Min. 59" — 2010 ADA'),
    ('hdr', 'Restroom Amenities (within Locker Room)', ''),
    ('yn', 'Separate restroom entrance?', ''),
    ('q',  'Width of restroom entrance (inches)', 'Min. 32"'),
    ('q',  'Turning space — diameter (inches)', 'Min. 60"'),
    ('yn', 'Hot water available?', ''),
    ('yna','Hot water pipes under sink covered?', ''),
    ('q',  'Height of sink — floor to rim (inches)', 'Max 34"'),
    ('q',  'Lavatory knee clearance — H/W/D (inches)', 'Min. 27"H / 30"W / 19"D'),
    ('q',  'Height of soap dispenser (inches)', 'Max 48"'),
    ('q',  'Height of mirror — floor to reflecting edge (inches)', 'Max 40"'),
    ('q',  'Height of paper towel dispenser (inches)', 'Max 48"'),
    ('hdr', 'Accessible Stall (within Locker Room)', ''),
    ('yn', 'Designated accessible stall?', ''),
    ('q',  'Width of accessible stall (inches)', 'Min. 60"'),
    ('q',  'Depth of accessible stall (inches)', 'Min. 56"'),
    ('yn', 'Stall door swings out?', ''),
    ('q',  'Width of stall door (inches)', 'Min. 32"'),
    ('yn', 'Grab bars in accessible stall?', ''),
    ('q',  'Location of grab bars (side / rear / both)', ''),
    ('q',  'Height of grab bars — floor to centerline (inches)', '33"–36"'),
    ('q',  'Height of toilet seat — floor to seat edge (inches)', '17"–19"'),
    ('q',  'Height of flush controls (inches)', 'Max 44"'),
    ('yn', 'Flush valve on transfer/open side of toilet?', ''),
    ('q',  'Height of toilet paper dispenser (inches)', ''),
    ('q',  'Urinal height — floor to rim (inches)', 'Max 17"'),
    ('q',  'Width of space around urinal (inches)', 'Min. 30"'),
    ('q',  'Additional Information / Observations', ''),
]

RESTROOMS = [
    ('q',  'Restroom type (Male / Female / Non-Gender Specific)', ''),
    ('yn', 'Restroom on accessible route?', ''),
    ('q',  'Total number of accessible restrooms in building', ''),
    ('yna','If no accessible restroom in building, directional signage to nearest?', ''),
    ('hdr', 'Entry Door & Signage', ''),
    ('yn', 'Accessible door signage (ISA)?', ''),
    ('q',  'Height of accessible signage — floor to centerline (inches)', '60"; 57"–63"'),
    ('yn', 'ISA contrasts light-on-dark or dark-on-light?', ''),
    ('yn', 'Characters raised + Grade II Braille?', ''),
    ('yn', 'Signage mounted adjacent to latch side of door?', ''),
    ('yn', 'Entrance doors swing out?', ''),
    ('q',  'Width of entrance doorway (inches)', 'Min. 32"'),
    ('q',  'Height of door handle (inches)', 'Max 48"'),
    ('q',  'Door opening force (pounds)', 'Max 5 lbs — 1991/UFAS/2010 | Max 8.5 — ANSI'),
    ('yn', 'Restroom has foyer?', ''),
    ('q',  'Measurement between doors in foyer (inches)', ''),
    ('hdr', 'Interior Maneuvering', ''),
    ('q',  'Turning space — diameter (inches)', 'Min. 60" diameter OR T-turn'),
    ('hdr', 'Amenities', ''),
    ('yn', 'Hot water available?', ''),
    ('yna','Hot water pipes under sink covered?', ''),
    ('q',  'Height of sink — floor to rim (inches)', 'Max 34"'),
    ('yn', 'Sink usable by wheelchair users?', ''),
    ('q',  'Lavatory knee clearance — H/W/D (inches)', 'Min. 27"H / 30"W / 19"D'),
    ('q',  'Height of soap dispenser (inches)', 'Max 48"'),
    ('q',  'Height of mirror — floor to reflecting edge (inches)', 'Max 40"'),
    ('q',  'Height of paper towel dispenser (inches)', 'Max 48"'),
    ('q',  'Height of automatic hand dryer (inches)', 'Max 48"'),
    ('yn', 'Racks, dispensers, disposal units not protruding into walkways?', ''),
    ('hdr', 'Urinals', ''),
    ('q',  'Height of urinal — floor to rim (inches)', 'Max 17"'),
    ('q',  'Width of space around urinal (inches)', 'Min. 30"'),
    ('q',  'Depth of space around urinal (inches)', 'Min. 48"'),
    ('hdr', 'Accessible Stall', ''),
    ('yn', 'Designated accessible stall?', ''),
    ('q',  'Width of accessible stall (inches)', 'Min. 60"'),
    ('q',  'Depth of accessible stall (inches)', 'Min. 56"'),
    ('yn', 'Stall door swings out?', ''),
    ('q',  'Width of stall door (inches)', 'Min. 32"'),
    ('yn', 'Grab bars in accessible stall?', ''),
    ('q',  'Location of grab bars (side / rear / both)', ''),
    ('q',  'Height of grab bars — floor to centerline (inches)', '33"–36"'),
    ('q',  'Length of grab bars (inches)', 'Side ≥40" | Rear ≥36"'),
    ('q',  'Depth of grab bars from wall (inches)', ''),
    ('q',  'Height of accessible toilet seat (inches)', '17"–19"'),
    ('q',  'Height of flush controls (inches)', 'Max 44"'),
    ('yn', 'Flush valve on transfer/open side of toilet?', ''),
    ('yn', 'Obstructions to reach toilet paper dispenser?', ''),
    ('yn', 'Toilet paper dispenser operable with one hand?', ''),
    ('yn', 'Dispenser allows continuous flow of paper?', ''),
    ('q',  'Height of toilet paper dispenser (inches)', ''),
    ('q',  'Location of toilet tissue dispenser', ''),
    ('yn', 'Coat hooks/shelves provided?', ''),
    ('q',  'Coat hook height — floor to hook (inches)', 'Max 48" forward reach'),
    ('q',  'Reach type for coat hook (forward / side)', ''),
    ('q',  'Additional Information / Observations', ''),
]

FOUNTAINS = [
    ('yn', 'Drinking fountains provided?', ''),
    ('q',  'Number of fountains', ''),
    ('q',  'Number of accessible fountains', ''),
    ('yn', 'At least 50% accessible (spout ≤36")?', ''),
    ('yn', 'Fountain on accessible route?', ''),
    ('yn', '30"×48" clear floor space in front?', ''),
    ('q',  'Height of spout — floor to spout outlet (inches)', 'Max 36"'),
    ('q',  'Height of water flow — spout to top of arc (inches)', 'Min. 4"'),
    ('yn', 'Water flow within 4" of spout towards front?', ''),
    ('yn', 'Up-front spout and control?', ''),
    ('q',  'Controls mounted on front or side?', ''),
    ('yn', 'Controls operable with closed fist?', ''),
    ('yn', 'Operable parts require tight grasping/pinching/twisting?', ''),
    ('yn', 'Force to activate exceeds 5 lbs?', ''),
    ('yn', 'Fountain protrudes into corridors or traffic ways?', ''),
    ('q',  'Additional Information / Observations', ''),
]

WATER_COOLERS = [
    ('yn', 'Bottle filler station available?', ''),
    ('yn', 'Station on accessible route?', ''),
    ('yn', '30"×48" clear floor space in front?', ''),
    ('yn', 'Station hand operated?', ''),
    ('yn', 'Station sensor operated?', ''),
    ('q',  'Controls mounted on front or side?', ''),
    ('q',  'Height of highest operable part (inches)', 'Max 48"'),
    ('yn', 'Controls operable with closed fist?', ''),
    ('yn', 'Operable parts require tight grasping/pinching/twisting?', ''),
    ('yn', 'Force to activate exceeds 5 lbs?', ''),
    ('q',  'Additional Information / Observations', ''),
]

CAFETERIA = [
    ('hdr', 'Signage', ''),
    ('yn', 'Identification signage provided?', ''),
    ('q',  'Height of signage — floor to centerline (inches)', '60"; 57"–63"'),
    ('q',  'Height of characters on signage (inches)', ''),
    ('yn', 'Signs mounted adjacent to latch side of door?', ''),
    ('yna','Double doors — signs on nearest adjacent wall?', ''),
    ('yn', 'Protruding objects within 3" of sign?', ''),
    ('yn', 'Non-glare finish on signs?', ''),
    ('yn', 'Characters raised + Grade II Braille?', ''),
    ('hdr', 'Entrance & Doors', ''),
    ('q',  'Number of entrances and exits', ''),
    ('yn', 'Cafeteria on accessible route?', ''),
    ('yna','Ramps leading to cafeteria?', 'If yes, complete Ramps section'),
    ('q',  'Width of doorway(s) (inches)', 'Min. 32"'),
    ('q',  'Height of door handle (inches)', 'Max 48"'),
    ('yn', 'Door openable without grasping/twisting wrist?', ''),
    ('q',  'If not, type of hardware on door', ''),
    ('yna','Automatic door opener provided?', ''),
    ('yna','Door closer takes at least 3 seconds to close?', ''),
    ('q',  'Door opening force (pounds)', 'Max 5 lbs — 1991/UFAS/2010 | Max 8.5 — ANSI'),
    ('q',  'Threshold — Exterior / Interior (inches)', 'Max 1/2"'),
    ('yna','Carpeting/mats 1/2" or less?', ''),
    ('hdr', 'Alarms', ''),
    ('yn', 'Visual alarm provided?', ''),
    ('yn', 'Light flashes clear or normal white?', ''),
    ('yn', 'Audible alarm provided?', ''),
    ('yn', 'Alarm exceeds prevailing sound level in room?', ''),
    ('hdr', 'Indoor Fixed Seating', ''),
    ('q',  'Total number of fixed accessible seats indoors', ''),
    ('yn', 'Seats and tables on accessible route?', ''),
    ('yn', 'Wheelchair seating space at fixed tables/counters?', ''),
    ('q',  'Height of fixed tables/countertops (inches)', '28"–34"'),
    ('q',  'Wheelchair space — W × D (inches)', 'Min. 30"×48"'),
    ('q',  'Knee clearance — H/W/D (inches)', 'Min. 27"H / 30"W / 19"D'),
    ('yn', 'Accessible route between food line, seating, and exit?', ''),
    ('hdr', 'Outdoor Fixed Seating', ''),
    ('q',  'Total number of fixed accessible seats outdoors', ''),
    ('yn', 'Seats and tables on accessible route?', ''),
    ('q',  'Height of outdoor fixed tables/countertops (inches)', '28"–34"'),
    ('q',  'Outdoor surface material', ''),
    ('hdr', 'Food Service Line(s)', ''),
    ('note','Measure all food service lines including main line, à la carte, and student store lines.', ''),
    ('q',  'Total number of food lines', ''),
    ('q',  'Food service line aisle width (inches)', 'Min. 36" clear'),
    ('q',  'Height of food line counter / tray slide (inches)', 'Max 34" — ANSI/UFAS | Max 36" — 1991/2010'),
    ('q',  'Height of sneeze guard / food display — lowest accessible item (inches)', 'Max 48" forward reach'),
    ('yn', 'Self-service items within reach range?', ''),
    ('q',  'Height of highest self-service item (inches)', 'Max 48"'),
    ('yn', 'Staff-assisted alternative available if self-service not accessible?', ''),
    ('hdr', 'Checkout Counter(s)', ''),
    ('q',  'Total number of checkout aisles', ''),
    ('q',  'Number of accessible checkout aisles', 'Min. 1 required'),
    ('q',  'Checkout aisle width (inches)', 'Min. 36"'),
    ('q',  'Height of accessible checkout counter (inches)', 'Max 34" — ANSI/UFAS | Max 36" — 1991/2010'),
    ('q',  'Width of lowered counter section (inches)', 'Min. 36"'),
    ('q',  'Knee clearance — H/W/D (inches)', 'Min. 27"H / 30"W / 19"D'),
    ('q',  'Approach type at checkout (Forward / Parallel)', ''),
    ('q',  'Height of tableware and condiment areas (inches)', 'Max 48" forward reach'),
    ('hdr', 'Tray Return / Beverage / Vending', ''),
    ('yn', 'Tray return station provided?', ''),
    ('yna','Tray return on accessible route?', ''),
    ('q',  'Height of tray return slot (inches)', 'Max 48"'),
    ('q',  'Clear floor space in front — W × D (inches)', 'Min. 30"×48"'),
    ('yn', 'Self-service beverage station provided?', ''),
    ('yna','Station on accessible route?', ''),
    ('q',  'Height of highest operable part — beverage (inches)', 'Max 48"'),
    ('q',  'Clear floor space — beverage — W × D (inches)', 'Min. 30"×48"'),
    ('yn', 'Beverage operable parts require tight grasping/pinching/twisting?', ''),
    ('yn', 'Beverage force to activate exceeds 5 lbs?', ''),
    ('yn', 'Vending machines provided?', ''),
    ('yna','Vending machines on accessible route?', ''),
    ('q',  'Height of vending highest operable part (inches)', 'Max 48" forward reach'),
    ('q',  'Vending clear floor space — W × D (inches)', 'Min. 30"×48"'),
    ('yn', 'Vending force to activate exceeds 5 lbs?', ''),
    ('q',  'Additional Information / Observations', ''),
]

LIBRARY = [
    ('hdr', 'Signage', ''),
    ('yn', 'Identification signage provided?', ''),
    ('q',  'Height of signage — floor to centerline (inches)', '60"; 57"–63"'),
    ('q',  'Height of characters on signage (inches)', ''),
    ('yn', 'Signs mounted adjacent to latch side of door?', ''),
    ('yna','Double doors — signs on nearest adjacent wall?', ''),
    ('yn', 'Protruding objects within 3" of sign?', ''),
    ('yn', 'Non-glare finish on signs?', ''),
    ('yn', 'Characters raised + Grade II Braille?', ''),
    ('hdr', 'Entrance & Doors', ''),
    ('q',  'Number of entrances and exits', ''),
    ('yn', 'Library on accessible route?', ''),
    ('yna','Ramps leading to library?', 'If yes, complete Ramps section'),
    ('q',  'Overhead clearance of accessible route (inches)', 'Min. 80"'),
    ('q',  'Width of doorway(s) (inches)', 'Min. 32"'),
    ('q',  'Height of door handle (inches)', 'Max 48"'),
    ('yn', 'Door openable without grasping/twisting wrist?', ''),
    ('q',  'If not, type of hardware on door', ''),
    ('yna','Automatic door opener provided?', ''),
    ('yna','Door closer takes at least 3 seconds to close?', ''),
    ('q',  'Door opening force (pounds)', 'Max 5 lbs — 1991/UFAS/2010 | Max 8.5 — ANSI'),
    ('q',  'Threshold — Exterior / Interior (inches)', 'Max 1/2"'),
    ('yna','Carpeting/mats 1/2" or less?', ''),
    ('hdr', 'Alarms', ''),
    ('yn', 'Visual alarm provided?', ''),
    ('yn', 'Light flashes clear or normal white?', ''),
    ('yn', 'Audible alarm provided?', ''),
    ('yn', 'Alarm exceeds prevailing sound level in room?', ''),
    ('hdr', 'Seating & Workstations', ''),
    ('yn', 'Wheelchair seating at fixed tables/counters?', ''),
    ('yn', 'Seats and tables on accessible route?', ''),
    ('q',  'Number of accessible seats', ''),
    ('q',  'Height of fixed tables/countertops (inches)', '28"–34"'),
    ('q',  'Aisle space width (inches)', 'Min. 36" | Min. 60" passing — 2010 ADA'),
    ('q',  'Turning space in reading areas — diameter (inches)', 'Min. 60"'),
    ('q',  'Knee clearance — H/W/D (inches)', 'Min. 27"H / 30"W / 19"D'),
    ('yn', 'Accessible desk/workstations have sufficient clearance?', ''),
    ('yna','If insufficient, alternative option available?', ''),
    ('hdr', 'Computer Workstations & Adaptive Technology', ''),
    ('yn', 'Computer workstations available?', ''),
    ('yn', 'Accessible computer workstation(s) provided?', ''),
    ('q',  'Height of accessible workstation surface (inches)', '28"–34"'),
    ('yn', 'Adaptive technology available?', ''),
    ('q',  'If yes, describe adaptive technology', ''),
    ('hdr', 'Checkout Counter', ''),
    ('q',  'Number of accessible checkout areas', ''),
    ('q',  'Height of checkout counter top (inches)', 'Max 34" — ANSI/UFAS | Max 36" — 1991/2010'),
    ('q',  'Width of lowered counter section (inches)', 'Min. 36"'),
    ('q',  'Counter top depth (inches)', ''),
    ('q',  'Knee clearance — H/W/D (inches)', 'Min. 27"H / 30"W / 19"D'),
    ('q',  'Forward reach approach (inches)', 'Max 48"'),
    ('yn', 'Parallel approach available?', ''),
    ('hdr', 'Book Drop', ''),
    ('note','Measuring the Book Drop: measure from finished floor to bottom edge of drop slot, in inches. Measure clear floor space in front (W×D). Test activation force. Note approach + location.', ''),
    ('yn', 'Book drop provided?', ''),
    ('yna','Book drop on accessible route?', ''),
    ('q',  'Height of drop slot — floor to bottom edge (inches)', 'Max 48" forward | Max 54" side — pre-2010 | Max 48" side — 2010'),
    ('q',  'Clear floor space in front — W × D (inches)', 'Min. 30"×48"'),
    ('q',  'Force required to activate book drop (pounds)', 'Max 5 lbs'),
    ('yn', 'Operable parts require tight grasping/pinching/twisting?', ''),
    ('q',  'Approach type (Forward / Side)', ''),
    ('q',  'Book drop location (Exterior / Interior)', ''),
    ('hdr', 'Card Catalog & Stack Aisles', ''),
    ('q',  'Card catalog aisle width (inches)', 'Min. 36"'),
    ('q',  'Card catalog operable part height (inches)', 'Max 48" forward reach'),
    ('q',  'Stack aisle width (inches)', 'Min. 36"'),
    ('q',  'Overhead clearance in stack aisles (inches)', 'Min. 80"'),
    ('q',  'Additional Information / Observations', ''),
]

GYM = [
    ('note','Complete for Gymnasium, Auditorium, Weight Room, or other large assembly space. If the campus has separate Gym and Auditorium buildings, note both in Additional Information or use separate printed copies.', ''),
    ('q',  'Space type (Gym / Aud / Weight Room / MP Room / Black Box / Other)', ''),
    ('q',  'Room number(s)', ''),
    ('hdr', 'Signage', ''),
    ('yn', 'Identification signage provided?', ''),
    ('q',  'Height of signage — floor to centerline (inches)', '60"; 57"–63"'),
    ('q',  'Height of characters on signage (inches)', ''),
    ('yn', 'Signs mounted adjacent to latch side of door?', ''),
    ('yna','Double doors — signs on nearest adjacent wall?', ''),
    ('yn', 'Protruding objects within 3" of sign?', ''),
    ('yn', 'Non-glare finish on signs?', ''),
    ('yn', 'Characters raised + Grade II Braille?', ''),
    ('hdr', 'Entrance & Doors', ''),
    ('yn', 'Space on accessible route?', ''),
    ('q',  'Number of entrances and exits', ''),
    ('q',  'Width of doorway(s) (inches)', 'Min. 32"'),
    ('q',  'Height of door handle (inches)', 'Max 48"'),
    ('yn', 'Door openable without grasping/twisting wrist?', ''),
    ('q',  'If not, type of hardware on door', ''),
    ('yna','Automatic door opener provided?', ''),
    ('yna','Door closer takes at least 3 seconds to close?', ''),
    ('q',  'Door opening force (pounds)', 'Max 5 lbs — 1991/UFAS/2010 | Max 8.5 — ANSI'),
    ('q',  'Threshold — Exterior / Interior (inches)', 'Max 1/2"'),
    ('yna','Carpeting/mats 1/2" or less?', ''),
    ('hdr', 'Alarms', ''),
    ('yn', 'Visual alarm provided?', ''),
    ('yn', 'Light flashes clear or normal white?', ''),
    ('yn', 'Audible alarm provided?', ''),
    ('yn', 'Alarm exceeds prevailing sound level in room?', ''),
    ('hdr', 'Wheelchair Seating', ''),
    ('q',  'Total number of all seating', ''),
    ('q',  'Total number of wheelchair spaces', 'Ratio varies by capacity — 2010 ADA §221'),
    ('yn', 'Wheelchair seating in different locations throughout venue?', ''),
    ('yn', 'Companion seat provided with each wheelchair space?', ''),
    ('yn', 'Accessible route to wheelchair designated seats?', ''),
    ('yna','Ramp to seating area?', 'If yes, complete Ramps section'),
    ('yn', 'Line of sight for wheelchair spaces comparable to ambulatory seating?', ''),
    ('q',  'Wheelchair space dimensions — W × D (inches)', 'Min. 36"W × 48"D'),
    ('yn', 'Can supplies be accessed by persons with disabilities?', ''),
    ('hdr', 'Assistive Listening Systems', ''),
    ('yn', 'Audio amplification system provided?', ''),
    ('yn', 'Adequate number of assistive listening systems (ALS) provided?', ''),
    ('q',  'Type of ALS (Induction Loop / FM / IR / Other)', ''),
    ('q',  'Number of receivers available', ''),
    ('q',  'If fixed audio amplification, distance from stage (ft)', ''),
    ('yn', 'Signage indicating availability of assistive listening devices?', ''),
    ('hdr', 'Fixed Seating / Tables (Weight Room / Other)', ''),
    ('yna','Wheelchair seating at fixed tables/counters?', ''),
    ('q',  'Number of accessible seats', ''),
    ('q',  'Height of fixed tables/countertops (inches)', '28"–34"'),
    ('q',  'Wheelchair space — W × D (inches)', 'Min. 30"×48"'),
    ('q',  'Knee clearance — H/W/D (inches)', 'Min. 27"H / 30"W / 19"D'),
    ('yna','Accessible desk/workstations have sufficient clearance?', ''),
    ('q',  'Additional Information / Observations', ''),
]

TELEPHONES = [
    ('yn', 'Public telephone provided?', 'If No, mark N/A and skip remaining fields'),
    ('q',  'Clear floor space — W × L (inches)', 'Min. 30"×48"'),
    ('yn', 'Protruding objects?', ''),
    ('yn', 'Push button controls?', ''),
    ('q',  'Cord length — telephone to handset (inches)', 'Min. 29"'),
    ('q',  'Approach type (Parallel / Forward)', ''),
    ('q',  'Distance from front edge of counter to face of telephone (inches)', ''),
    ('yn', 'Highest operable part within 15"–48"? (54" max if side approach)', ''),
    ('yn', 'Telephone hearing aid compatible?', ''),
    ('yn', 'Volume control provided?', ''),
    ('yna','If >4 pay telephones, TTY available?', ''),
    ('q',  'Height of TTY keypad touch surface (inches)', 'Max 48"'),
    ('yna','Telephones with TTY equipped with shelf and electrical outlet?', ''),
    ('q',  'Additional Information / Observations', ''),
]

# Program Access Interview — single-instance, no location columns
PA_INTERVIEW = [
    ('hdr', 'ADA Coordination & Compliance', ''),
    ('q', '1. Who is the designated ADA/Section 504 Coordinator for this facility, and how are accessibility complaints or accommodation requests handled?', ''),
    ('q', '2. Has the facility undergone a formal self-evaluation under Section 504? If so, when and what were the findings?', ''),
    ('q', '3. Does the facility have a current Transition Plan? When was it last updated and what items remain outstanding?', ''),
    ('hdr', 'Maintenance of Accessible Features', ''),
    ('q', '4. Are accessible features (ramps, lifts, elevators, accessible restrooms, automatic door openers) on the regular maintenance schedule? How frequently inspected?', ''),
    ('q', '5. Are maintenance staff trained to identify and report accessibility barriers? Describe the training.', ''),
    ('q', '6. Are accessible parking spaces regularly monitored to prevent unauthorized use? What is the enforcement process?', ''),
    ('q', '7. How are temporary accessibility barriers (construction, equipment storage, event setup) managed to maintain accessible routes during disruptions?', ''),
    ('hdr', 'Program Access & Relocation', ''),
    ('q', '8. Are there areas where a program, class, or activity has been relocated to provide program access? If so, where, why, and how was this communicated?', ''),
    ('q', '9. What is the process for notifying students, parents, and staff of accessible routes, features, and accommodations?', ''),
    ('q', '10. If a student with a disability needed access to an inaccessible area, what is the process for providing equivalent program access?', ''),
    ('hdr', 'Emergency Egress & Safety', ''),
    ('q', '11. How are emergency evacuation procedures adapted for mobility impairments? Are Areas of Rescue Assistance designated, equipped, and included in drills?', ''),
    ('q', '12. Are two-way communication systems at Areas of Rescue Assistance tested regularly? When last tested?', ''),
    ('hdr', 'Planned Modifications', ''),
    ('q', '13. Are any accessibility modifications currently planned, funded, approved, or in progress? Describe scope, timeline, and funding source.', ''),
    ('q', '14. Have any accessibility modifications been completed in the past five years? Describe scope and dates.', ''),
    ('hdr', 'Reviewer Observations', ''),
    ('q', 'Program Reviewer observations and notes from interview', ''),
]

# ─────────────────────────────────────────────────────────────────────────────
# Section list — order matches the HTML sidebar
# Each entry: (display_name, schema, kind)
#   kind: 'multi' (4 location columns) | 'single' (2 columns)
# ─────────────────────────────────────────────────────────────────────────────
SECTIONS = [
    ('Accessible Parking',                     PARKING,        'multi'),
    ('Accessible Routes / Walkways',           ROUTES,         'multi'),
    ('Curb Ramps',                             CURB_RAMPS,     'multi'),
    ('Stadium / Field',                        STADIUM,        'single'),
    ('Entrances, Doors, and Gates',            ENTRANCES,      'multi'),
    ('Stairways and Steps',                    STAIRS,         'multi'),
    ('Ramps',                                  RAMPS,          'multi'),
    ('Elevators',                              ELEVATORS,      'multi'),
    ('Lifts',                                  LIFTS,          'multi'),
    ('Rooms & Offices',                        ROOMS,          'multi'),
    ('Cafeteria',                              CAFETERIA,      'single'),
    ('Library',                                LIBRARY,        'single'),
    ('CTE Classroom(s)',                       CTE,            'multi'),
    ('Labs / Shops',                           LABS,           'multi'),
    ('Gymnasium / Auditorium / Weight Room',   GYM,            'single'),
    ('Locker Rooms',                           LOCKER,         'multi'),
    ('Restrooms (Male / Female / All-Gender)', RESTROOMS,      'multi'),
    ('Accessible Drinking Fountains',          FOUNTAINS,      'multi'),
    ('Water Coolers & Bottle Fillers',         WATER_COOLERS,  'multi'),
    ('Public Telephones',                      TELEPHONES,     'single'),
]

# ─────────────────────────────────────────────────────────────────────────────
# Open source doc, identify cut points, REPLACE middle, preserve front+back
# ─────────────────────────────────────────────────────────────────────────────
print("Loading source guide…")
doc = Document(SRC)
print(f"  {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
body = doc.element.body

# Find first Heading 1 (start of replaceable content) and Glossary Heading 1
first_h1_xml = None
glossary_xml = None
for p in doc.paragraphs:
    if (p.style.name or "").strip() == "Heading 1":
        if first_h1_xml is None:
            first_h1_xml = p._p
        if "Glossary" in p.text:
            glossary_xml = p._p
            break

if first_h1_xml is None or glossary_xml is None:
    raise RuntimeError("Could not locate cut points")

# Remove everything between first H1 (inclusive) and Glossary (exclusive)
to_remove = []
collecting = False
for child in list(body):
    if child is first_h1_xml: collecting = True
    if collecting:
        if child is glossary_xml: break
        to_remove.append(child)

print(f"Removing {len(to_remove)} body elements between first H1 and Glossary…")
for ch in to_remove:
    body.remove(ch)

# ── Build new content ────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
# REPEATING SECTION CONTENT CONTROL helper (v5)
# Wrap an element (the table) in a w15:repeatingSection SDT.
# When the cursor is inside the table in Word, a "+" button appears on the
# right margin. Click "+" → Word duplicates the inner SDT's content.
# Requires Word 2013+ / Mac 2016+ / Word for the web.
# ─────────────────────────────────────────────────────────────────────────────
def wrap_repeating_section(table_element, alias_name):
    """Wrap a table in a Repeating Section Content Control."""
    # Outer SDT — declares this is a repeating section
    outer = OxmlElement("w:sdt")
    outer_pr = OxmlElement("w:sdtPr")
    rid = OxmlElement("w:id"); rid.set(qn("w:val"), next_id()); outer_pr.append(rid)
    alias = OxmlElement("w:alias"); alias.set(qn("w:val"), alias_name); outer_pr.append(alias)
    tag = OxmlElement("w:tag")
    tag.set(qn("w:val"), "rs_" + re.sub(r"[^A-Za-z0-9]+", "_", alias_name).strip("_"))
    outer_pr.append(tag)
    # The w15 namespace element that marks this as a repeating section
    etree.SubElement(outer_pr, f"{{{W15_NS}}}repeatingSection")
    outer.append(outer_pr)
    outer.append(OxmlElement("w:sdtEndPr"))
    outer_content = OxmlElement("w:sdtContent")
    outer.append(outer_content)

    # Inner SDT — one repeating-section ITEM (the unit that gets cloned)
    inner = OxmlElement("w:sdt")
    inner_pr = OxmlElement("w:sdtPr")
    rid2 = OxmlElement("w:id"); rid2.set(qn("w:val"), next_id()); inner_pr.append(rid2)
    etree.SubElement(inner_pr, f"{{{W15_NS}}}repeatingSectionItem")
    inner.append(inner_pr)
    inner.append(OxmlElement("w:sdtEndPr"))
    inner_content = OxmlElement("w:sdtContent")
    # The actual table to duplicate
    inner_content.append(table_element)
    inner.append(inner_content)

    outer_content.append(inner)
    return outer

# ─────────────────────────────────────────────────────────────────────────────
# Build new content
# ─────────────────────────────────────────────────────────────────────────────
new = []

# Render each section
for name, schema, kind in SECTIONS:
    new.extend(render_section_heading(name, multi=(kind=='multi')))
    if kind == 'multi':
        section_table = render_section_table(schema, num_locs=NUM_LOCATIONS)
        new.append(wrap_repeating_section(section_table, f"{name} — locations"))
    else:
        new.append(render_single_table(schema))

# Program Access Interview (no Location/Year Built rows — single-column responses)
new.append(_page_break())
new.append(_t("Program Access — Facilities & M&O Staff Interview", bold=True, size=SZ_TITLE, after=120))
new.append(_t(
    "These questions are asked of the school site's Facilities and Maintenance & Operations staff by the Program Reviewer during the on-site visit. Program Access compliance is determined OBSERVATIONALLY and through staff responses — not through physical measurement.",
    italic=True, size=SZ_NOTE, color=GRAY, after=120))

# Header (single-cell rows for staff name + title + date + interviewer)
LBL2 = 3000
RSP = 9700 - LBL2
hdr_rows = [
    ('q', 'Staff Member Name:', ''),
    ('q', 'Title:', ''),
    ('q', 'Date of Interview:', ''),
    ('q', 'Interviewer:', ''),
]
new.append(_table([LBL2, RSP], [_section_row(f, [LBL2, RSP], 1) for f in hdr_rows]))
new.append(_empty_para())

# Then the interview questions as a single-column-response table
new.append(_table([LBL2, RSP],
    [_section_row(f, [LBL2, RSP], 1) for f in PA_INTERVIEW]))

# ── Insert before Glossary ───────────────────────────────────────────────────
parent = glossary_xml.getparent()
idx = list(parent).index(glossary_xml)
print(f"Inserting {len(new)} elements before Glossary at position {idx}…")
for e in new:
    parent.insert(idx, e); idx += 1

# ─────────────────────────────────────────────────────────────────────────────
# Ensure the document root declares the w15 namespace so Word recognizes the
# repeating-section markers we just embedded.
# ─────────────────────────────────────────────────────────────────────────────
doc_root = body.getparent()  # <w:document>
if doc_root.nsmap.get("w15") != W15_NS:
    # lxml's nsmap is read-only; rebuild the element with the extra namespace
    new_nsmap = dict(doc_root.nsmap)
    new_nsmap["w15"] = W15_NS
    new_root = etree.Element(doc_root.tag, attrib=dict(doc_root.attrib), nsmap=new_nsmap)
    for child in list(doc_root):
        new_root.append(child)
    # Also copy any Ignorable attribute and add w15 to mc:Ignorable so older
    # Word versions ignore the markers cleanly.
    mc_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    ign_attr = f"{{{mc_ns}}}Ignorable"
    cur = new_root.get(ign_attr, "")
    if "w15" not in cur.split():
        new_root.set(ign_attr, (cur + " w15").strip())
    # Swap the root
    doc_root.getparent().replace(doc_root, new_root) if doc_root.getparent() is not None else None
    # python-docx's document.xml is the root, so we replace directly
    # by reassigning the body's parent reference (handled via the part)
    part = doc.part
    part._element = new_root
    print("✓ Added w15 namespace declaration to document.xml root.")

# ── Save ─────────────────────────────────────────────────────────────────────
os.makedirs("outputs", exist_ok=True)
doc.save(OUT)
print(f"\n✓ Saved → {OUT}")
print(f"  File size: {os.path.getsize(OUT)/1024:.1f} KB")
print(f"  SDT ID counter: {_next_id}")
