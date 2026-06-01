"""Build the CRR 20/21 Accessibility Training PowerPoint deck.

Generates outputs/CRR_Accessibility_Training.pptx — a full-day training for
new CDE OEO reviewers covering ANSI A117.1 (1961), ADA 1991/ADAAG, ADA 2010,
UFAS, and Program Access, with a focus on K-12 / CTE facilities.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ── Palette (mirrors facilities_lof_tool.html) ────────────────────────────────
NAVY      = RGBColor(0x1A, 0x27, 0x44)
NAVY_MID  = RGBColor(0x24, 0x34, 0x60)
GOLD      = RGBColor(0xC8, 0xA8, 0x4B)
GOLD_LT   = RGBColor(0xE8, 0xD0, 0x8A)
CREAM     = RGBColor(0xF7, 0xF4, 0xEE)
SAGE      = RGBColor(0xED, 0xF0, 0xEB)
BORDER    = RGBColor(0xD4, 0xCE, 0xBD)
TEXT      = RGBColor(0x1A, 0x1A, 0x1A)
MUTED     = RGBColor(0x6B, 0x66, 0x59)
RED       = RGBColor(0xB9, 0x1C, 0x1C)
GREEN     = RGBColor(0x16, 0x65, 0x34)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY= RGBColor(0xF3, 0xF4, 0xF6)
INK       = RGBColor(0x37, 0x41, 0x51)

# Standard badge colors (match HTML tool)
B_PA      = RGBColor(0xE5, 0xE7, 0xEB)
B_ANSI    = RGBColor(0xFE, 0xF9, 0xC3)
B_UFAS    = RGBColor(0xDB, 0xEA, 0xFE)
B_1991    = RGBColor(0xD1, 0xFA, 0xE5)
B_2010    = RGBColor(0xFE, 0xE2, 0xE2)

# 16:9 widescreen
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]

# ── Helpers ───────────────────────────────────────────────────────────────────
def add_slide():
    return prs.slides.add_slide(BLANK)

def set_bg(slide, color=WHITE):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.line.fill.background()
    bg.fill.solid(); bg.fill.fore_color.rgb = color
    bg.shadow.inherit = False
    slide.shapes._spTree.remove(bg._element); slide.shapes._spTree.insert(2, bg._element)
    return bg

def add_rect(slide, x, y, w, h, fill, line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(0.75)
    s.shadow.inherit = False
    return s

def add_text(slide, x, y, w, h, text, *, size=18, bold=False, color=TEXT,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Calibri",
             italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        r.font.name = font; r.font.size = Pt(size)
        r.font.bold = bold; r.font.italic = italic
        r.font.color.rgb = color
    return tb

def add_bullets(slide, x, y, w, h, items, *, size=16, color=TEXT, bullet="•",
                spacing=4, bold_first=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    for i, item in enumerate(items):
        if isinstance(item, tuple):
            head, body = item
        else:
            head, body = None, item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(spacing)
        if head:
            r = p.add_run(); r.text = f"{bullet}  {head}"
            r.font.name = "Calibri"; r.font.size = Pt(size)
            r.font.bold = True; r.font.color.rgb = color
            r2 = p.add_run(); r2.text = f"  {body}"
            r2.font.name = "Calibri"; r2.font.size = Pt(size)
            r2.font.color.rgb = color
        else:
            r = p.add_run(); r.text = f"{bullet}  {body}"
            r.font.name = "Calibri"; r.font.size = Pt(size)
            r.font.bold = bold_first if isinstance(bold_first, bool) else False
            r.font.color.rgb = color
    return tb

def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text.strip()

def header_bar(slide, title, sub=None, kicker=None):
    """Standard slide header: thin navy bar, kicker chip, title text."""
    set_bg(slide, WHITE)
    # Top accent bar
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.18), NAVY)
    add_rect(slide, 0, Inches(0.18), SLIDE_W, Inches(0.04), GOLD)
    # Kicker chip
    if kicker:
        chip = add_rect(slide, Inches(0.55), Inches(0.42), Inches(2.6), Inches(0.32), NAVY)
        add_text(slide, Inches(0.55), Inches(0.42), Inches(2.6), Inches(0.32),
                 kicker.upper(), size=10, bold=True, color=GOLD,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    # Title
    ty = Inches(0.85) if kicker else Inches(0.55)
    add_text(slide, Inches(0.5), ty, Inches(12.3), Inches(0.7), title,
             size=30, bold=True, color=NAVY, font="Georgia")
    if sub:
        add_text(slide, Inches(0.5), ty + Inches(0.65), Inches(12.3), Inches(0.4), sub,
                 size=14, color=MUTED, italic=True)
    # Bottom thin rule + footer page number filled per slide
    add_rect(slide, 0, Inches(7.30), SLIDE_W, Inches(0.02), BORDER)

def footer(slide, page_num, total=63):
    add_text(slide, Inches(0.5), Inches(7.18), Inches(8), Inches(0.25),
             "CDE OEO  •  Civil Rights Review  •  CRR 20 / CRR 21 — Accessible Facilities",
             size=9, color=MUTED, font="Consolas")
    add_text(slide, Inches(11.5), Inches(7.18), Inches(1.3), Inches(0.25),
             f"{page_num:02d} / {total:02d}",
             size=9, color=MUTED, font="Consolas", align=PP_ALIGN.RIGHT)

def std_badge(slide, x, y, label, fill, text_color=INK, w=None):
    w = w or Inches(1.2)
    badge = add_rect(slide, x, y, w, Inches(0.28), fill)
    badge.line.color.rgb = BORDER; badge.line.width = Pt(0.5)
    add_text(slide, x, y, w, Inches(0.28), label, size=10, bold=True,
             color=text_color, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             font="Consolas")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 1 — Title
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); set_bg(s, CREAM)
add_rect(s, 0, 0, SLIDE_W, Inches(7.5), NAVY)
add_rect(s, 0, Inches(2.9), SLIDE_W, Inches(0.08), GOLD)
# Seal
seal = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.6), Inches(0.6), Inches(0.9), Inches(0.9))
seal.fill.solid(); seal.fill.fore_color.rgb = GOLD
seal.line.fill.background(); seal.shadow.inherit = False
add_text(s, Inches(0.6), Inches(0.6), Inches(0.9), Inches(0.9), "CDE",
         size=18, bold=True, color=NAVY, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Georgia")
add_text(s, Inches(1.7), Inches(0.7), Inches(8), Inches(0.4),
         "CALIFORNIA DEPARTMENT OF EDUCATION  •  OFFICE OF EQUAL OPPORTUNITY",
         size=11, bold=True, color=GOLD, font="Consolas")
add_text(s, Inches(1.7), Inches(1.05), Inches(10), Inches(0.4),
         "Civil Rights Review  •  Reviewer Onboarding Training",
         size=12, color=GOLD_LT, italic=True)

add_text(s, Inches(0.7), Inches(3.4), Inches(12), Inches(1.2),
         "Accessible Facilities for K–12 and CTE",
         size=42, bold=True, color=WHITE, font="Georgia")
add_text(s, Inches(0.7), Inches(4.3), Inches(12), Inches(0.7),
         "ANSI A117.1 · ADA 1991 / ADAAG · ADA 2010 · UFAS · Program Access",
         size=20, color=GOLD_LT)
add_text(s, Inches(0.7), Inches(4.95), Inches(12), Inches(0.5),
         "Applied to CRR 20 (2025–26) and CRR 21 (2024–25)",
         size=16, color=GOLD_LT, italic=True)

add_text(s, Inches(0.7), Inches(6.4), Inches(12), Inches(0.4),
         "FULL-DAY ONBOARDING  •  ~7 HOURS WITH BREAKS  •  IN-PERSON OR VIRTUAL",
         size=10, bold=True, color=GOLD, font="Consolas")
add_text(s, Inches(0.7), Inches(6.75), Inches(12), Inches(0.4),
         "Office of Equal Opportunity · California Department of Education",
         size=11, color=GOLD_LT, italic=True)

add_notes(s, """
Welcome the class. Introduce yourself and name your role on the CRR team.
This is a full-day onboarding training for new reviewers who will conduct
CRR 20 (Accessible Facilities) or CRR 21 (CTE Facility Site Selection)
reviews. By the end of the day, every reviewer should be able to look at
a Facilities Review Guide, determine which accessibility standard applies
to each area, identify violations, and write a defensible LOF.

Have everyone introduce themselves and share whether they have any prior
ADA / accessibility background. Adjust depth on the technical sections
based on the room.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 2 — Course Objectives
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Course Objectives", kicker="Slide 02 · Orientation")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "By the end of this training, you will be able to:",
         size=16, italic=True, color=MUTED)
add_bullets(s, Inches(0.7), Inches(2.05), Inches(12), Inches(5),
    [
        ("Identify",  "the five accessibility standards used in CRR 20 / 21 and the date window each governs."),
        ("Apply",     "the standard-determination decision tree using construction and ADA modification dates."),
        ("Recognize", "the Program Access rule — and why deficiencies in pre-1977 facilities never become findings."),
        ("Locate",    "the controlling citation for the most common dimensional requirements (handrails, ramps, doors, signage, restrooms)."),
        ("Evaluate",  "CTE-specific spaces — shops, labs, kitchens, ag, health science — against the applicable standard."),
        ("Use",       "the Facilities LOF Generator tool to draft an editable findings table from a self-assessment packet."),
        ("Write",     "an Accessibility Violation and Corrective Action that will stand up to OCR review."),
    ], size=15, spacing=6)
footer(s, 2)
add_notes(s, """
Walk the class through each objective. Emphasize that this is not a code
class — they are not architects. Their job is to (a) determine the
correct standard, (b) compare measurements against it, and (c) document
what they found in defensible language.

Stress objective #3 (Program Access) early — it is the single most common
mistake new reviewers make. They see a barrier in a 1968 building and
write it up as a finding. That is wrong, and we will spend significant
time on why.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 3 — Agenda
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Full-Day Agenda", kicker="Slide 03 · Orientation")

agenda = [
    ("9:00",  "Orientation & legal foundations",         "Slides 04–10"),
    ("9:45",  "The five standards & decision tree",       "Slides 11–18"),
    ("10:30", "Break"),
    ("10:45", "Program Access doctrine",                 "Slides 19–22"),
    ("11:30", "ANSI A117.1 (1961) deep-dive",            "Slides 23–25"),
    ("12:00", "Lunch"),
    ("1:00",  "UFAS · 1991 ADA · 2010 ADA",              "Slides 26–38"),
    ("2:30",  "Break"),
    ("2:45",  "Measurements reference + CTE deep-dives", "Slides 39–50"),
    ("3:45",  "Two reviewer rules: element-by-element & corrective = 2010 ADA", "Slides 51–54"),
    ("4:15",  "Writing findings + case studies",         "Slides 55–61"),
    ("5:00",  "Resources & wrap-up",                     "Slides 62–63"),
]
ay = Inches(1.55)
for i, row in enumerate(agenda):
    y = ay + Inches(0.42 * i)
    if len(row) == 2:
        add_rect(s, Inches(0.7), y, Inches(12), Inches(0.36), SAGE)
        add_text(s, Inches(0.9), y, Inches(1.5), Inches(0.36), row[0],
                 size=12, bold=True, color=MUTED, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
        add_text(s, Inches(2.4), y, Inches(8), Inches(0.36), row[1],
                 size=13, color=MUTED, italic=True, anchor=MSO_ANCHOR.MIDDLE)
    else:
        add_text(s, Inches(0.9), y, Inches(1.5), Inches(0.36), row[0],
                 size=12, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
        add_text(s, Inches(2.4), y, Inches(8.5), Inches(0.36), row[1],
                 size=14, bold=True, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(11), y, Inches(2), Inches(0.36), row[2],
                 size=11, color=MUTED, anchor=MSO_ANCHOR.MIDDLE, font="Consolas",
                 align=PP_ALIGN.RIGHT)
footer(s, 3)
add_notes(s, """
Times are notional — adjust to your start time. Two ten-minute breaks and
a one-hour lunch are built in. Tell the class when the breaks land so they
can plan. If you are running virtual, swap the two breaks for 5-minute
stretches and a 45-minute lunch.

If you are short on time, the sections you can compress (in order):
(1) survey of remaining CTE sectors at slide 49, (2) UFAS deep-dive at 26-27
(it rarely controls a finding in K-12), (3) one of the three case studies.
Never compress the decision tree or Program Access sections.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 4 — Why this matters
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Why this matters", kicker="Slide 04 · Orientation")

add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.5),
         "A facility barrier is not just a code issue — it is an exclusion from public education.",
         size=18, italic=True, color=NAVY)

# Three stat cards
cards = [
    ("1 in 4", "California adults reports a disability (CDC, 2024).",
     "Every campus we review has students, parents, and staff who depend on access."),
    ("60+ yrs", "of federal accessibility law — yet most CTE shops we visit predate it.",
     "California schools built between 1955 and 1990 dominate the CRR queue."),
    ("$0",     "is the cost of getting the standard right the first time.",
     "Misapplying ANSI to a 2015 building, or 2010 ADA to a 1968 building, can invalidate the entire LOF."),
]
for i, (big, head, body) in enumerate(cards):
    x = Inches(0.55 + 4.2*i)
    add_rect(s, x, Inches(2.3), Inches(4), Inches(4.4), CREAM, line=BORDER)
    add_rect(s, x, Inches(2.3), Inches(4), Inches(0.08), GOLD)
    add_text(s, x + Inches(0.2), Inches(2.5), Inches(3.6), Inches(1.0), big,
             size=44, bold=True, color=NAVY, font="Georgia")
    add_text(s, x + Inches(0.2), Inches(3.7), Inches(3.6), Inches(1.0), head,
             size=14, bold=True, color=TEXT)
    add_text(s, x + Inches(0.2), Inches(4.8), Inches(3.6), Inches(1.8), body,
             size=12, color=MUTED, italic=True)
footer(s, 4)
add_notes(s, """
Frame the day. The reviewer's report goes to the LEA, to OEO leadership,
and — if there is a complaint — potentially to the U.S. Department of
Education's Office for Civil Rights (OCR). Treat every LOF as something
that might be quoted in a federal complaint or in a settlement agreement.

Personalize: share a brief example of a real (anonymized) finding that
made a difference for a student. CTE access matters because it is the
on-ramp to wages — if a student with a mobility disability can't get into
the welding bay, they are effectively locked out of the sector.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 5 — Your role as a CDE OEO reviewer
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Your role as a CDE OEO reviewer", kicker="Slide 05 · Orientation")

# Two-column: You DO / You DO NOT
add_rect(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(0.5), GREEN)
add_text(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(0.5),
         "YOU DO",
         size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_rect(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(0.5), RED)
add_text(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(0.5),
         "YOU DO NOT",
         size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")

add_rect(s, Inches(0.55), Inches(2.05), Inches(6.0), Inches(5.0), CREAM, line=BORDER)
add_bullets(s, Inches(0.75), Inches(2.15), Inches(5.7), Inches(4.9), [
    "Verify the LEA's self-assessment against documentation (blueprints, work orders, DSA approvals).",
    "Determine the applicable standard for each area by date.",
    "Compare measured field conditions to the applicable standard.",
    "Document violations with the controlling citation in the LOF.",
    "Issue corrective actions tied to that citation, with a 45-day deadline.",
    "Coordinate with the LEA's ADA/Title II and Section 504 coordinator.",
], size=13, spacing=6)

add_rect(s, Inches(6.85), Inches(2.05), Inches(6.0), Inches(5.0), CREAM, line=BORDER)
add_bullets(s, Inches(7.05), Inches(2.15), Inches(5.7), Inches(4.9), [
    "Design fixes or recommend specific construction details (LEA hires the architect).",
    "Apply a newer standard to an older building (you are not retroactive).",
    "Issue findings against Program Access — there is no measurable standard to cite.",
    "Replace OCR enforcement — your LOF triggers corrective action, not litigation.",
    "Interpret the standard yourself when text is ambiguous — escalate to your lead.",
    "Sign off without DSA-approved date evidence in the file.",
], size=13, spacing=6, bullet="✕")
footer(s, 5)
add_notes(s, """
Reviewers are fact-finders, not designers. Their job is to determine
what the standard requires, what the field shows, and whether the two
match. The LEA hires architects and contractors to design fixes — the
reviewer does not.

The "you do not" list captures the top mistakes new reviewers make. Walk
through each. The most common one is the second: applying the 2010
standard to a 1985 building because the reviewer recognizes a barrier
the 2010 standard would prohibit. That is a misapplication of law and
will be reversed on appeal.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 6 — Section 504
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Section 504 of the Rehabilitation Act (1973)",
                            kicker="Slide 06 · Legal foundations")
# Left: history
add_text(s, Inches(0.55), Inches(1.55), Inches(6.5), Inches(0.4),
         "The original federal civil-rights statute on disability.",
         size=15, italic=True, color=MUTED)
add_bullets(s, Inches(0.55), Inches(2.05), Inches(6.5), Inches(5), [
    ("29 U.S.C. § 794.", "Prohibits discrimination on the basis of disability by any program receiving federal financial assistance."),
    ("34 CFR Part 104.", "U.S. Dept. of Education's implementing regulation — the rule book for K-12."),
    ("§ 104.22.", "Program accessibility for existing facilities (built or altered on/before June 3, 1977)."),
    ("§ 104.23.", "New construction and alteration accessibility (built or altered after June 3, 1977)."),
    ("Why it matters.", "Federal funds flow to every CA LEA — Section 504 is the umbrella under which ANSI, UFAS, 1991 ADA, and 2010 ADA are sequentially incorporated."),
], size=13, spacing=8)
# Right: timeline card
add_rect(s, Inches(7.4), Inches(1.55), Inches(5.4), Inches(5.3), CREAM, line=BORDER)
add_rect(s, Inches(7.4), Inches(1.55), Inches(5.4), Inches(0.4), NAVY)
add_text(s, Inches(7.4), Inches(1.55), Inches(5.4), Inches(0.4), "504 IMPLEMENTATION TIMELINE",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
events = [
    ("Sept 26, 1973", "Rehabilitation Act signed"),
    ("May 4, 1977",   "DHEW issues § 504 regulations"),
    ("June 3, 1977",  "Existing facilities cutoff date"),
    ("June 4, 1977",  "ANSI A117.1-1961 (R1971) becomes the standard for new construction"),
    ("Jan 18, 1991",  "UFAS replaces ANSI for new construction"),
    ("Jan 27, 1992",  "1991 ADA enters the chain"),
    ("Mar 15, 2012",  "2010 ADA becomes the default"),
]
for i, (date, ev) in enumerate(events):
    y = Inches(2.15 + 0.6*i)
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.6), y + Inches(0.08), Inches(0.18), Inches(0.18))
    dot.fill.solid(); dot.fill.fore_color.rgb = GOLD; dot.line.fill.background()
    add_text(s, Inches(7.9), y, Inches(2.0), Inches(0.35), date,
             size=11, bold=True, color=NAVY, font="Consolas")
    add_text(s, Inches(9.8), y, Inches(3.0), Inches(0.35), ev,
             size=12, color=TEXT)
footer(s, 6)
add_notes(s, """
Section 504 is the legal anchor for everything we do. The CRR program
operates under 504 — not under the ADA — even though the ADA standards
are incorporated by reference. That distinction matters when you write
findings: the citation chain is "34 CFR § 104.23 (Section 504) → 1991 ADA
Standards → § 4.13.5 (door width)."

The June 3, 1977 cutoff date is the single most important date in this
training. Pre-June-4, 1977 buildings are governed by Program Access only —
no measurable standard applies. Drill this date into the class.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 7 — ADA Title II
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "ADA Title II (1990) — Public Entities",
                            kicker="Slide 07 · Legal foundations")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Title II reaches every state and local government program — including public schools — regardless of federal funding.",
         size=14, italic=True, color=MUTED)
add_bullets(s, Inches(0.55), Inches(2.1), Inches(12), Inches(4.5), [
    ("42 U.S.C. § 12131 et seq.",   "Prohibits discrimination by public entities. CA public schools and county offices of education are 'public entities.'"),
    ("28 CFR Part 35.",             "DOJ implementing regulation."),
    ("§ 35.150.",                   "Existing facilities — program accessibility (mirrors § 104.22)."),
    ("§ 35.151.",                   "New construction and alterations — adopts UFAS, 1991 ADA, and 2010 ADA as the design standards."),
    ("Standards incorporated.",     "Title II does NOT have its own dimensional code — it borrows from the date-controlled chain we'll cover next."),
    ("Dual coverage.",              "K-12 facilities are covered by both § 504 and Title II. Both citations appear on every CRR 20 finding."),
], size=14, spacing=6)

# Bottom callout
add_rect(s, Inches(0.55), Inches(6.3), Inches(12.3), Inches(0.7), B_2010, line=RED)
add_text(s, Inches(0.75), Inches(6.3), Inches(12), Inches(0.7),
         "Reviewer rule: cite both authorities. \"28 CFR § 35.151(a); 34 CFR § 104.23(a)\" — every time.",
         size=14, bold=True, color=RED, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 7)
add_notes(s, """
The reason we cite both 504 and Title II together is that they are the
two federal authorities OCR can enforce. A district could conceivably
argue "we don't take federal funds for this CTE program" — which would
take § 504 off the table — but Title II still applies because the
district is a public entity. Belt and suspenders.

Walk through the CFR section numbers slowly. New reviewers find the cite
chain confusing. Show them the boilerplate from the CRR 20 Word
template — that same chain appears verbatim at the top of every
Facilities LOF.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 8 — OCR & DOJ Enforcement
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "OCR and DOJ enforcement — the bigger context",
                            kicker="Slide 08 · Legal foundations")
# Two-column
add_rect(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(5.3), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(0.45), NAVY)
add_text(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(0.45),
         "U.S. DEPT. OF EDUCATION — OCR",
         size=12, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(0.75), Inches(2.1), Inches(5.7), Inches(4.8), [
    "Enforces § 504 in K-12 and higher education.",
    "Investigates individual and systemic complaints.",
    "Resolution typically a Resolution Agreement with corrective action plan.",
    "Can refer to DOJ for litigation.",
    "Maintains list of open and closed K-12 access cases — publicly searchable.",
    "CDE LOFs sit in the OCR file when complaints overlap.",
], size=13, spacing=8)

add_rect(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(5.3), CREAM, line=BORDER)
add_rect(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(0.45), NAVY)
add_text(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(0.45),
         "U.S. DEPT. OF JUSTICE — DOJ",
         size=12, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(7.05), Inches(2.1), Inches(5.7), Inches(4.8), [
    "Enforces ADA Title II — including the 1991 and 2010 Standards.",
    "Publishes the 2010 Standards (28 CFR Part 35 Appendix B).",
    "Issues binding guidance ('Technical Assistance').",
    "Can sue public entities, intervene in private suits.",
    "Settlement agreements often impose 2010 ADA standards across a portfolio.",
    "Project Civic Access has produced 250+ K-12 settlement agreements.",
], size=13, spacing=8)
footer(s, 8)
add_notes(s, """
Reviewers should know that their LOF is one of several documents that may
end up in front of OCR. If a district disputes a finding, OCR may be the
final arbiter. That is why citations must be exact — vague language is
the easiest thing to push back on.

Point reviewers to www2.ed.gov/ocr (OCR case search) and ada.gov
(DOJ technical assistance) as live resources. They will use these during
their first reviews.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 9 — CDE OEO authority
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "CDE OEO — your authority to review",
                            kicker="Slide 09 · Legal foundations")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.45),
         "California Department of Education · Office of Equal Opportunity",
         size=15, color=MUTED, italic=True)
add_bullets(s, Inches(0.55), Inches(2.1), Inches(12), Inches(5), [
    ("Federal subrecipient monitoring.", "CDE distributes federal Perkins V CTE funds and is required to monitor subrecipients (LEAs) for civil-rights compliance under 34 CFR § 100.7 (Title VI), § 104.61 (§ 504), and § 106.71 (Title IX)."),
    ("California Education Code § 220 et seq.", "State-law parallel that gives CDE authority to investigate equity issues."),
    ("CRR program.", "Three-year on-site review cycle covering all civil-rights authorities. CRR 20 = Accessible Facilities. CRR 21 = CTE Facility Site Selection. Both touch building accessibility."),
    ("Outputs.", "Letter of Findings (LOF), Voluntary Compliance Plan (VCP) with 45-day corrective-action deadlines, and a follow-up evidence review."),
    ("Enforcement path.", "Persistent non-compliance can result in CDE withholding Perkins funds; serious cases referred to OCR or DOJ."),
], size=14, spacing=8)
footer(s, 9)
add_notes(s, """
Reviewers sometimes feel they lack 'real' authority compared to OCR
investigators. They have plenty. CDE OEO findings carry the weight of
federal subrecipient monitoring — if a district ignores a VCP, CDE can
withhold Perkins funds, which is often a more immediate consequence
than an OCR case that takes years.

Spend a minute on the three-year cycle — every CA LEA running CTE will
see a CRR every three years, so each reviewer's docket is steady.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 10 — CRR 20 / CRR 21 at a glance
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "CRR 20 and CRR 21 at a glance",
                            kicker="Slide 10 · Legal foundations")
# Two cards side-by-side
add_rect(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(5.5), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(0.55), NAVY)
add_text(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(0.55),
         "CRR 20 — Accessible Facilities", size=15, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Georgia")
add_bullets(s, Inches(0.75), Inches(2.25), Inches(5.7), Inches(4.7), [
    ("Cycle.",     "2025–26 (current)."),
    ("Authorities.", "34 CFR § 104.22 & § 104.23; 28 CFR § 35.151; 28 CFR Part 36 App. D."),
    ("Scope.",     "Buildings and facilities used to deliver CTE — and the common-use areas that serve them (restrooms, drinking fountains, parking, paths of travel)."),
    ("Evidence.",  "Site plans, DSA-approved blueprints, alteration/modification records, work orders, completed Facilities Review Packet."),
    ("Output.",    "18-row findings table → CRR boilerplate + LOF cover letter (Word)."),
], size=12, spacing=6)

add_rect(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(5.5), CREAM, line=BORDER)
add_rect(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(0.55), NAVY)
add_text(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(0.55),
         "CRR 21 — CTE Facility Site Selection", size=15, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Georgia")
add_bullets(s, Inches(7.05), Inches(2.25), Inches(5.7), Inches(4.7), [
    ("Cycle.",     "2024–25 (most recent prior cycle)."),
    ("Authorities.", "Same Section 504 + Title II chain, plus 34 CFR § 100.3 (Title VI siting analysis)."),
    ("Scope.",     "Where CTE programs are located — does siting create disparate impact by disability, race, language, or sex?"),
    ("Evidence.",  "Maps showing CTE locations, student demographics before/after modifications, route accessibility."),
    ("Overlap.",   "Same five accessibility standards, same self-assessment, same 18-area LOF table — only the narrative framing changes."),
], size=12, spacing=6)
footer(s, 10)
add_notes(s, """
Reviewers need to know which CRR they are assigned to before they walk
on site. The technical content is nearly identical — same five
standards, same dimensions, same 18-area structure. The narrative
framing differs: CRR 20 asks "is this building accessible?" while
CRR 21 asks "is the siting of CTE programs equitable across protected
classes?"

For today, treat CRR 20 and CRR 21 as one body of work. Differences
are in the cover memo and which authority appears first in the cite
chain.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 11 — Five standards comparison
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "The five standards at a glance",
                            kicker="Slide 11 · The standards")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Each standard governs a specific date window. Pick the wrong window — pick the wrong standard.",
         size=14, italic=True, color=MUTED)

# Table-style layout
headers = ["Standard", "Date window", "Authority", "Findings?", "Tag"]
col_x = [Inches(0.55), Inches(3.4), Inches(5.7), Inches(9.4), Inches(11.0)]
col_w = [Inches(2.85), Inches(2.3), Inches(3.7), Inches(1.6), Inches(2.2)]

# Header row
for i, h in enumerate(headers):
    add_rect(s, col_x[i], Inches(2.05), col_w[i], Inches(0.4), NAVY)
    add_text(s, col_x[i], Inches(2.05), col_w[i], Inches(0.4), h.upper(),
             size=10, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")

rows = [
    ("Program Access", "≤ June 3, 1977",            "34 CFR § 104.22  ·  28 CFR § 35.150",  "Never",  "PA",   B_PA,   INK),
    ("ANSI A117.1 (1961)", "June 4, 1977 – Jan 17, 1991", "34 CFR § 104.23(a)",         "Yes",    "ANSI", B_ANSI, RGBColor(0x71,0x3F,0x12)),
    ("UFAS",           "Jan 18, 1991 – Jan 26, 1992", "28 CFR § 35.151(a); 41 CFR 101-19.6 App.A", "Yes", "UFAS", B_UFAS, RGBColor(0x1E,0x3A,0x8A)),
    ("1991 ADA / ADAAG", "Jan 27, 1992 – Sep 14, 2010 (and overlap to Mar 14, 2012)", "28 CFR Pt 36 App. D", "Yes", "ADA 1991", B_1991, RGBColor(0x06,0x4E,0x3B)),
    ("2010 ADA",       "≥ Mar 15, 2012",            "28 CFR Pt 35 App. B",                 "Yes",    "ADA 2010", B_2010, RGBColor(0x7F,0x1D,0x1D)),
]
for i, (name, dates, auth, findings, tag, fill, tcol) in enumerate(rows):
    y = Inches(2.45 + 0.78*i)
    row_bg = WHITE if i % 2 == 0 else CREAM
    for j in range(5):
        add_rect(s, col_x[j], y, col_w[j], Inches(0.74), row_bg, line=BORDER)
    add_text(s, col_x[0]+Inches(0.1), y, col_w[0], Inches(0.74), name,
             size=12, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, col_x[1]+Inches(0.1), y, col_w[1], Inches(0.74), dates,
             size=10, color=TEXT, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    add_text(s, col_x[2]+Inches(0.1), y, col_w[2], Inches(0.74), auth,
             size=10, color=TEXT, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    fcolor = RED if findings == "Never" else GREEN
    add_text(s, col_x[3], y, col_w[3], Inches(0.74), findings,
             size=12, bold=True, color=fcolor, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
    std_badge(s, col_x[4]+Inches(0.4), y+Inches(0.23), tag, fill, text_color=tcol, w=Inches(1.4))

footer(s, 11)
add_notes(s, """
This is the most-referenced slide in the deck. Tell the class to bookmark
it. They will use this table on every review.

Walk down the rows. Highlight that Program Access is the only standard
where the "Findings?" answer is NEVER — not "rarely" or "depends." A
common new-reviewer instinct is to write up a barrier in a 1970 building
because the barrier feels obvious. That is incorrect: the standard
specifies that pre-June-4-1977 facilities are evaluated under program
access, which is not a measurable code.

The 2010-and-2012 overlap window (Sept 2010 – Mar 14, 2012) is the only
place where two standards may apply. The LEA chooses; the reviewer
honors the LEA's choice if it is documented in the construction record.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 12 — Why date matters
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Why construction date controls the standard",
                            kicker="Slide 12 · The standards")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Accessibility law is not retroactive. The standard in force when a building was built or last altered governs forever — unless and until the building is altered again.",
         size=14, italic=True, color=MUTED)

# Concept boxes
boxes = [
    ("Built in 1965 · never altered",
     "Program Access only.  Barriers can exist legally — what matters is whether the program as a whole is accessible.",
     B_PA, INK),
    ("Built in 1985 · never altered",
     "ANSI A117.1 (1961 R1971).  Cite ANSI sections only.  2010 ADA is not retroactive.",
     B_ANSI, RGBColor(0x71,0x3F,0x12)),
    ("Built in 1995 · re-roofed in 2018",
     "1991 ADA / ADAAG governs the building.  Re-roofing is not an 'accessibility-affecting' alteration, so it does not bump up the standard.",
     B_1991, RGBColor(0x06,0x4E,0x3B)),
    ("Built in 1968 · restrooms gutted and rebuilt in 2019",
     "Original building stays Program Access.  Altered restrooms are evaluated under 2010 ADA.",
     B_2010, RGBColor(0x7F,0x1D,0x1D)),
]
for i, (label, body, fill, tcol) in enumerate(boxes):
    col = i % 2
    row = i // 2
    x = Inches(0.55 + 6.2*col)
    y = Inches(2.15 + 2.45*row)
    add_rect(s, x, y, Inches(6.0), Inches(2.3), WHITE, line=BORDER)
    add_rect(s, x, y, Inches(0.18), Inches(2.3), fill)
    add_text(s, x+Inches(0.35), y+Inches(0.15), Inches(5.5), Inches(0.5),
             label, size=14, bold=True, color=NAVY)
    add_text(s, x+Inches(0.35), y+Inches(0.7), Inches(5.5), Inches(1.5),
             body, size=12, color=TEXT)
footer(s, 12)
add_notes(s, """
The fourth example is the most important — and the one new reviewers miss
most often. Partial alterations 'isolate' the altered element to the
newer standard. The original (pre-1977) building remains under Program
Access; the gutted restrooms become 2010 ADA work.

Write the four scenarios on the board (if in person) and have the class
call out which standard applies before you reveal the answer. This is the
single best diagnostic of whether they understand the rule.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 13 — Decision tree (flowchart)
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Standard determination — the decision tree",
                            kicker="Slide 13 · Decision tree")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Use this every time. If you cannot answer the date question, you cannot issue a finding.",
         size=14, italic=True, color=MUTED)

# Decision tree drawn as diamond + boxes
# Center diamond top
def diamond(x, y, w, h, text):
    d = s.shapes.add_shape(MSO_SHAPE.DIAMOND, x, y, w, h)
    d.fill.solid(); d.fill.fore_color.rgb = NAVY
    d.line.color.rgb = GOLD; d.line.width = Pt(1.5)
    d.shadow.inherit = False
    tf = d.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = GOLD; r.font.name = "Calibri"
    return d

def outcome(x, y, w, h, label, fill, tcol=INK):
    b = add_rect(s, x, y, w, h, fill, line=BORDER)
    add_text(s, x, y, w, h, label, size=11, bold=True, color=tcol,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")

def arrow(x1, y1, x2, y2, label=None):
    ln = s.shapes.add_connector(1, x1, y1, x2, y2)
    ln.line.color.rgb = MUTED; ln.line.width = Pt(1.25)
    if label:
        add_text(s, (x1+x2)//2 - Inches(0.3), (y1+y2)//2 - Inches(0.15),
                 Inches(0.6), Inches(0.3), label,
                 size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE, font="Consolas")

# Top diamond
diamond(Inches(5.3), Inches(2.1), Inches(2.7), Inches(1.0),
        "Was this element altered after construction?")
# Left path: no altered
diamond(Inches(1.8), Inches(3.7), Inches(2.7), Inches(1.0),
        "Use ORIGINAL construction date")
diamond(Inches(8.8), Inches(3.7), Inches(2.7), Inches(1.0),
        "Use ALTERATION / ADA-mod date")
arrow(Inches(6.65), Inches(3.1), Inches(3.15), Inches(3.7), "NO")
arrow(Inches(6.65), Inches(3.1), Inches(10.15), Inches(3.7), "YES")

# Bottom outcomes
outcomes = [
    (Inches(0.4),  Inches(5.05), "≤ Jun 3, 1977",        "PROGRAM ACCESS", B_PA),
    (Inches(2.7),  Inches(5.05), "Jun 4, 1977 – Jan 17, 1991", "ANSI A117.1",   B_ANSI),
    (Inches(5.0),  Inches(5.05), "Jan 18 – Jan 26, 1992", "UFAS",            B_UFAS),
    (Inches(7.3),  Inches(5.05), "Jan 27, 1992 – Sep 14, 2010", "1991 ADA",     B_1991),
    (Inches(9.6),  Inches(5.05), "Sep 15, 2010 – Mar 14, 2012", "1991 or 2010 ADA", B_2010),
    (Inches(11.9), Inches(5.05), "≥ Mar 15, 2012",      "2010 ADA",        B_2010),
]
# Draw arrows from both middle diamonds down to a horizontal rail
rail_y = Inches(4.85)
for x, y, label, std, fill in outcomes:
    add_rect(s, x, y, Inches(2.2), Inches(0.4), CREAM, line=BORDER)
    add_text(s, x, y, Inches(2.2), Inches(0.4), label,
             size=10, color=TEXT, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    add_rect(s, x, y+Inches(0.45), Inches(2.2), Inches(0.45), fill, line=BORDER)
    add_text(s, x, y+Inches(0.45), Inches(2.2), Inches(0.45), std,
             size=11, bold=True, color=INK, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")

# Reviewer rule callout
add_rect(s, Inches(0.55), Inches(6.5), Inches(12.3), Inches(0.55), GOLD_LT, line=GOLD)
add_text(s, Inches(0.55), Inches(6.5), Inches(12.3), Inches(0.55),
         "Reviewer rule: If construction date is missing or disputed, request DSA-approved blueprints before issuing any finding.",
         size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
footer(s, 13)
add_notes(s, """
This is the single most-used diagram. Have a printed handout version
that reviewers carry on site visits. The rule at the bottom is critical:
if the LEA cannot produce DSA-approved blueprints or alteration records,
the reviewer cannot determine the applicable standard, and therefore
cannot issue a finding. Document the missing evidence in the LOF as a
records deficiency.

The 'September 2010 – March 2012' window is the only one with two
possible standards. In that window the LEA can elect either 1991 or
2010 ADA — the reviewer should honor the LEA's documented election. If
the election is not documented, default to 2010 ADA (more stringent).
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 14 — The six date windows visualized
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "The six date windows — visualized",
                            kicker="Slide 14 · Decision tree")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Read left to right. Every facility falls into exactly one band based on its construction or last-alteration date.",
         size=13, italic=True, color=MUTED)

# Horizontal timeline
bar_y = Inches(3.2)
bar_h = Inches(0.9)
total_w = Inches(12.3)
x0 = Inches(0.55)

# Bands proportional to date span (approximate)
bands = [
    ("Program\nAccess",    "≤ Jun 3, 1977",        B_PA,   INK, 0.27),
    ("ANSI A117.1",        "Jun 4, '77 – Jan 17, '91", B_ANSI, RGBColor(0x71,0x3F,0x12), 0.18),
    ("UFAS",               "Jan 18 – Jan 26, '92", B_UFAS, RGBColor(0x1E,0x3A,0x8A), 0.02),
    ("1991 ADA / ADAAG",   "Jan 27, '92 – Sep 14, 2010", B_1991, RGBColor(0x06,0x4E,0x3B), 0.20),
    ("Overlap\n('91 or '10)", "Sep 15, '10 – Mar 14, '12", B_2010, RGBColor(0x7F,0x1D,0x1D), 0.05),
    ("2010 ADA",           "≥ Mar 15, 2012",       B_2010, RGBColor(0x7F,0x1D,0x1D), 0.28),
]
# Build a timeline rail
add_rect(s, x0, bar_y - Inches(0.05), total_w, Inches(0.04), MUTED)
# Year tick marks
years = [(1955, 0.0), (1977, 0.27), (1991, 0.45), (1992, 0.47), (2010, 0.67), (2012, 0.72), (2026, 1.0)]
for year, frac in years:
    tx = x0 + Emu(int(total_w * frac))
    add_rect(s, tx, bar_y - Inches(0.18), Inches(0.015), Inches(0.13), MUTED)
    add_text(s, tx - Inches(0.3), bar_y - Inches(0.38), Inches(0.6), Inches(0.25),
             str(year), size=9, color=MUTED, align=PP_ALIGN.CENTER, font="Consolas")

# Band rectangles
cur = 0.0
for label, sub, fill, tcol, frac in bands:
    bx = x0 + Emu(int(total_w * cur))
    bw = Emu(int(total_w * frac))
    add_rect(s, bx, bar_y, bw, bar_h, fill, line=BORDER)
    tb = add_text(s, bx, bar_y, bw, bar_h, label, size=11, bold=True, color=tcol,
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Calibri")
    # Sub-label below
    add_text(s, bx, bar_y + bar_h + Inches(0.05), bw, Inches(0.3), sub,
             size=9, color=MUTED, align=PP_ALIGN.CENTER, font="Consolas")
    cur += frac

# Examples below
add_text(s, Inches(0.55), Inches(5.4), Inches(12.3), Inches(0.4),
         "Where common CA campuses fall:",
         size=14, bold=True, color=NAVY)
exs = [
    ("• 1957 ag building, never altered  →  Program Access (PA)"),
    ("• 1989 wood shop, never altered  →  ANSI A117.1"),
    ("• 2003 culinary classroom, never altered  →  1991 ADA"),
    ("• 2011 STEM addition, no documented election  →  default to 2010 ADA"),
    ("• 2019 welding pad, brand new construction  →  2010 ADA"),
]
add_bullets(s, Inches(0.55), Inches(5.8), Inches(12.3), Inches(1.4), exs, size=12, spacing=2, bullet="")
footer(s, 14)
add_notes(s, """
The timeline is not to actual scale — Program Access and 2010 ADA bands
are stretched to show context. The microscopic UFAS band (Jan 18 1991 –
Jan 26 1992 — barely over one year) deserves a callout: it almost never
controls a finding in K-12 because so few schools were built in that
narrow window. Reviewers should still know it exists.

Walk through each example. Have the class predict the standard before
you reveal it.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 15 — Construction vs ADA modification date
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Element-by-element — the alteration rule",
                            kicker="Slide 15 · Decision tree")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.45),
         "Only the altered elements get the new standard. Unaltered elements keep the original construction-date standard. The building does not \"upgrade\" as a whole.",
         size=14, italic=True, color=MUTED)

# Access Board quote
add_rect(s, Inches(0.55), Inches(2.1), Inches(12.3), Inches(1.05), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(2.1), Inches(0.18), Inches(1.05), GOLD)
add_text(s, Inches(0.85), Inches(2.2), Inches(11.8), Inches(0.3),
         "U.S. Access Board — 2010 ADA Scoping § 202.3 (Alterations):",
         size=11, bold=True, color=NAVY, font="Consolas")
add_text(s, Inches(0.85), Inches(2.5), Inches(11.7), Inches(0.65),
         "\"Only those elements or spaces altered are required to comply… If a room or space is completely altered (or built new as part of an alteration), the entire room or space is fully subject to the standards.\"",
         size=12, italic=True, color=TEXT)

# Two columns
add_rect(s, Inches(0.55), Inches(3.3), Inches(6.0), Inches(3.7), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(3.3), Inches(6.0), Inches(0.4), NAVY)
add_text(s, Inches(0.55), Inches(3.3), Inches(6.0), Inches(0.4),
         "UNALTERED ELEMENTS", size=11, bold=True, color=GOLD,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(0.75), Inches(3.8), Inches(5.7), Inches(3.2), [
    "Standard = original construction-date era.",
    "Pre-1977 unaltered → Program Access (no findings).",
    "1985 unaltered → ANSI A117.1 controls.",
    "Re-roofing, paint, HVAC, electrical upgrades do NOT alter accessibility — element stays at original standard.",
], size=12, spacing=6)

add_rect(s, Inches(6.85), Inches(3.3), Inches(6.0), Inches(3.7), CREAM, line=BORDER)
add_rect(s, Inches(6.85), Inches(3.3), Inches(6.0), Inches(0.4), NAVY)
add_text(s, Inches(6.85), Inches(3.3), Inches(6.0), Inches(0.4),
         "ALTERED ELEMENTS ONLY", size=11, bold=True, color=GOLD,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(7.05), Inches(3.8), Inches(5.7), Inches(3.2), [
    "Standard = standard in effect on alteration date.",
    "Each altered ELEMENT gets evaluated separately.",
    "A 1968 restroom with 2018 grab-bar replacement → grab bars under 2010 ADA; toilet, lavatory, mirror still under Program Access (unless they were also altered).",
    "Whole-room rule: if an entire room is gutted and rebuilt, the entire room is subject to the alteration-date standard.",
], size=12, spacing=6)

add_rect(s, Inches(0.55), Inches(7.05), Inches(12.3), Inches(0.25), GOLD)
add_text(s, Inches(0.55), Inches(7.01), Inches(12.3), Inches(0.3),
         "Reviewer rule — request work orders / DSA records for each altered element. No proof of alteration = default to original construction date.",
         size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 15)
add_notes(s, """
This is the single most-misunderstood rule. New reviewers want to assign
"the building" one standard. The Access Board's 2010 ADA Scoping
guidance is explicit: alteration analysis is element-by-element.

Concrete example to walk through aloud:
A 1968 restroom has these elements: WC, lavatory, mirror, grab bars,
signage, dispensers. In 2018 the LEA replaced the grab bars only —
work orders confirm. Result:
- WC, lavatory, mirror, signage, dispensers = Program Access
  (1968 construction date, no alteration to those elements).
- Grab bars = 2010 ADA (altered 2018).

This means the reviewer can issue a finding ONLY on the grab bars
(if they don't meet 2010 ADA 609.4 / 604.5). The reviewer CANNOT
issue findings on the WC seat height or lavatory rim under any
standard — those elements remain Program Access.

The "whole room" exception: if the LEA gutted the entire restroom in
2018 (toilet, lav, walls, plumbing), the WHOLE room becomes 2010 ADA.
That's the difference between an "element alteration" and a "room
alteration." Look at the scope of work.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 16 — Worked example #1
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Worked example — Mixed-era campus",
                            kicker="Slide 16 · Decision tree")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.5),
         "Pacific Career & Technology HS — 4 buildings, 1 site",
         size=16, bold=True, color=NAVY)

# Table of elements
hd = ["Element", "Original built", "Last accessibility alteration", "Applicable standard"]
xs = [Inches(0.55), Inches(4.6), Inches(6.4), Inches(9.5)]
ws = [Inches(4.05), Inches(1.8), Inches(3.1), Inches(3.3)]
for i, h in enumerate(hd):
    add_rect(s, xs[i], Inches(2.2), ws[i], Inches(0.4), NAVY)
    add_text(s, xs[i], Inches(2.2), ws[i], Inches(0.4), h.upper(),
             size=10, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")

data = [
    ("Main classroom wing (unaltered)",             "1962", "None",                "PROGRAM ACCESS",  B_PA),
    ("Ag bldg & greenhouse (unaltered)",            "1971", "None",                "PROGRAM ACCESS",  B_PA),
    ("Wood / metal shop bays (unaltered)",          "1986", "None",                "ANSI A117.1",     B_ANSI),
    ("Culinary bldg — entire building (new build)", "2014", "None",                "2010 ADA",        B_2010),
    ("Parking-lot striping ONLY (route element)",   "1962", "Restriped 2022",      "2010 ADA",        B_2010),
    ("Front entry restrooms — GUTTED rebuild",      "1962", "Whole room rebuilt 2019", "2010 ADA",   B_2010),
    ("Library entry ramp (new install, no prior)",  "2003", "None",                "1991 ADA",        B_1991),
    ("Locker-room grab bars ONLY (element swap)",   "1962", "Bars replaced 2018",  "2010 ADA (bars only)", B_2010),
]
for i, (el, built, alt, std, fill) in enumerate(data):
    y = Inches(2.6 + 0.48*i)
    row_bg = WHITE if i % 2 == 0 else CREAM
    for j in range(4):
        add_rect(s, xs[j], y, ws[j], Inches(0.48), row_bg, line=BORDER)
    add_text(s, xs[0]+Inches(0.1), y, ws[0], Inches(0.48), el, size=10, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, xs[1], y, ws[1], Inches(0.48), built, size=10, color=TEXT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    add_text(s, xs[2]+Inches(0.1), y, ws[2], Inches(0.48), alt, size=9, color=TEXT,
             anchor=MSO_ANCHOR.MIDDLE)
    std_badge(s, xs[3]+Inches(0.8), y+Inches(0.10), std, fill, w=Inches(1.7))

footer(s, 16)
add_notes(s, """
This is the kind of campus you will actually walk. Most CA schools are
mixed-era. The reviewer's job is to keep the standards straight
ELEMENT BY ELEMENT — not building by building.

Key teaching moments per row:

Row 4 (Culinary 2014): whole BUILDING is new construction → entire
building under 2010 ADA. Every element. This is the only kind of
"whole building" answer that's clean.

Row 5 (parking-lot striping): striping is itself an accessibility
ELEMENT. Restriping in 2022 makes the striping subject to 2010 ADA
502.3.3 (van aisle), 502.6 (ISA), etc. The pavement, vertical clearance,
and route surface continue to be 1962 elements unless those were also
altered.

Row 6 (restrooms — GUTTED): the whole-room rule applies — entire
restroom rebuilt in 2019 = entire restroom under 2010 ADA.

Row 8 (locker-room grab bars): contrast with row 6. Only the bars were
swapped. The bars become 2010 ADA (must meet 609.4). The toilet,
lavatory, mirror, and stall geometry remain 1962 elements → Program
Access. Reviewer can issue a finding ONLY on the bars; cannot issue
findings on the toilet seat height under any standard.

This is the rule the user MUST internalize before walking a site.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 17 — Worked example #2 (Mid-century, no alterations)
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Worked example — A pre-1977 ag campus",
                            kicker="Slide 17 · Decision tree")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.5),
         "Valley Union High School — Ag program, original 1958 buildings",
         size=16, bold=True, color=NAVY)

add_bullets(s, Inches(0.55), Inches(2.1), Inches(6.0), Inches(4.5), [
    ("Built 1958.", "Three single-story ag classrooms, one shop, one greenhouse."),
    ("No documented alterations.", "Re-roofed twice; HVAC replaced — none of which affect accessibility."),
    ("Field measurements show:",
     "Door clear width 28 inches (would violate 2010 ADA 404.2.3); ramp slope 1:10 (would violate 2010 ADA 405.2); no ISA signage at restrooms."),
], size=14, spacing=10)

# Right side: the rule applied
add_rect(s, Inches(7.0), Inches(2.1), Inches(5.9), Inches(4.5), CREAM, line=BORDER)
add_rect(s, Inches(7.0), Inches(2.1), Inches(5.9), Inches(0.5), NAVY)
add_text(s, Inches(7.0), Inches(2.1), Inches(5.9), Inches(0.5),
         "REVIEWER'S FINDING", size=12, bold=True, color=GOLD,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(7.2), Inches(2.7), Inches(5.6), Inches(3.8), [
    "Applicable standard: PROGRAM ACCESS.",
    "Accessibility Violation column: \"None.\"",
    "Corrective Actions column: \"None.\"",
    "No violation citation possible — ANSI, UFAS, and ADA do not apply to pre-Jun-4-1977 facilities.",
    "Program-as-a-whole analysis: Does the LEA offer the ag program in an accessible alternate location? If yes — compliant. If no — a separate Title II program-access analysis is warranted.",
], size=12, spacing=6)

# Big alert
add_rect(s, Inches(0.55), Inches(6.7), Inches(12.3), Inches(0.45), B_2010, line=RED)
add_text(s, Inches(0.55), Inches(6.7), Inches(12.3), Inches(0.45),
         "DO NOT cite 2010 ADA. DO NOT issue a finding on this building. Document the program-access analysis instead.",
         size=14, bold=True, color=RED, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 17)
add_notes(s, """
This example is the inverse of the previous one. A 1958 building with
clear physical barriers, but no findings can be issued.

The pushback from new reviewers is always: 'But students can't use that
ramp.' That is correct, and that is exactly what program-access analysis
is designed to address — by relocating the program, not by citing a
standard that doesn't apply.

Tell the class: if the LEA cannot demonstrate program access by other
means, the reviewer recommends relocation in the LOF narrative. The
reviewer does NOT cite a dimensional standard against a 1958 building.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 18 — Knowledge check #1
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); set_bg(s, NAVY)
# Side band
add_rect(s, 0, 0, Inches(0.6), SLIDE_H, GOLD)
add_text(s, Inches(0.85), Inches(0.6), Inches(12), Inches(0.5),
         "KNOWLEDGE CHECK · 01",
         size=12, bold=True, color=GOLD, font="Consolas")
add_text(s, Inches(0.85), Inches(1.1), Inches(12), Inches(0.9),
         "Which standard applies?",
         size=36, bold=True, color=WHITE, font="Georgia")

scenarios = [
    ("A.", "A 1972 high school gym — locker rooms gutted and rebuilt in 2017."),
    ("B.", "A 1995 culinary classroom — no documented alterations."),
    ("C.", "A 1965 wood shop — restriped parking lot in 2024; building never altered."),
    ("D.", "A 2011 STEM building — district records do not document an election between 1991 and 2010 ADA."),
]
for i, (lbl, sc) in enumerate(scenarios):
    y = Inches(2.3 + 0.85*i)
    add_rect(s, Inches(0.85), y, Inches(0.6), Inches(0.65), GOLD)
    add_text(s, Inches(0.85), y, Inches(0.6), Inches(0.65), lbl, size=20, bold=True,
             color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Georgia")
    add_rect(s, Inches(1.55), y, Inches(11.3), Inches(0.65), NAVY_MID)
    add_text(s, Inches(1.75), y, Inches(11.0), Inches(0.65), sc,
             size=14, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Inches(0.85), Inches(6.6), Inches(12), Inches(0.4),
         "Take 3 minutes. Write the standard for each. Answers in the next slide.",
         size=12, color=GOLD, italic=True, font="Consolas")
footer(s, 18)
add_notes(s, """
Answers (do not reveal until the class commits to an answer):

A. Two standards in one building. Original 1972 gym = ANSI A117.1 (still
   reads "PROGRAM ACCESS" if you go by the simpler dichotomy, but
   technically the 1972 build date puts it in the ANSI window — wait,
   1972 is BEFORE June 4, 1977 so it is PROGRAM ACCESS). The locker
   rooms gutted in 2017 = 2010 ADA.

   CORRECT: Gym building = Program Access (built before June 4, 1977).
   Locker rooms (rebuilt 2017) = 2010 ADA.

B. 1995 build, no alterations = 1991 ADA / ADAAG.

C. 1965 building = Program Access. Parking lot restriped in 2024 = 2010
   ADA for the parking lot. The wood shop building itself remains
   Program Access.

D. September 15, 2010 to March 14, 2012 is the overlap window. If the
   LEA cannot document its election, the reviewer defaults to 2010 ADA
   (the more stringent of the two).

Common error on A: students will say "the gym is 1972 so use ANSI." Walk
them through the date table — 1972 is BEFORE June 4, 1977, so it is
Program Access. ANSI starts at June 4, 1977.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 19 — Program Access doctrine
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Program Access — the doctrine",
                            kicker="Slide 19 · Program Access")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "34 CFR § 104.22  ·  28 CFR § 35.150  —  the rule for existing facilities.",
         size=14, italic=True, color=MUTED)

# Quote block
add_rect(s, Inches(0.55), Inches(2.1), Inches(12.3), Inches(2.0), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(2.1), Inches(0.18), Inches(2.0), GOLD)
add_text(s, Inches(0.85), Inches(2.25), Inches(12), Inches(0.3),
         "34 CFR § 104.22(a) — verbatim:",
         size=11, bold=True, color=NAVY, font="Consolas")
add_text(s, Inches(0.85), Inches(2.6), Inches(11.7), Inches(1.5),
         "\"A recipient shall operate its program or activity so that when each part is viewed in its entirety, it is readily accessible to handicapped persons. This paragraph does not require a recipient to make each of its existing facilities or every part of a facility accessible to and usable by handicapped persons.\"",
         size=13, color=TEXT, italic=True)

# Key concepts
add_text(s, Inches(0.55), Inches(4.4), Inches(12.3), Inches(0.4),
         "What this means in practice:",
         size=15, bold=True, color=NAVY)
add_bullets(s, Inches(0.55), Inches(4.85), Inches(12.3), Inches(2.5), [
    ("Program-as-a-whole.", "Accessibility is evaluated at the program level, not the building level. The question is: can a student with a disability access the program?"),
    ("No retroactive code.", "Pre-Jun-4-1977 facilities do not need to comply with ANSI, UFAS, or ADA dimensional standards."),
    ("Methods are flexible.", "Compliance can be achieved by relocating the program, reassigning sections, providing auxiliary aids, or — only when nothing else works — making structural changes."),
    ("Structural change is last resort.", "Title II § 35.150(b)(1): structural changes are required only when there is no other feasible method."),
], size=13, spacing=6)
footer(s, 19)
add_notes(s, """
Read the regulation aloud — slowly. "Viewed in its entirety" and
"readily accessible" are the operative phrases. They mean: the program
must be accessible, the building need not be.

This is the doctrinal heart of why Program Access does not yield
findings. There is no measurable code to cite. A reviewer cannot write
"violation: handrail at 30 inches" against a 1965 building, because the
controlling regulation does not specify a handrail height. The
regulation specifies a program-level outcome.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 20 — "Viewed in its entirety"
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "What \"viewed in its entirety\" means",
                            kicker="Slide 20 · Program Access")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "The four methods of providing program access — in priority order.",
         size=14, italic=True, color=MUTED)

methods = [
    ("1", "Relocate the program",
     "Move the inaccessible class section to a building or room that IS accessible. The most common and cheapest method.",
     B_PA),
    ("2", "Reassign students or staff",
     "If multiple sections exist, offer the accessible section to the student who needs it. Works for shared programs (math, English).",
     B_PA),
    ("3", "Provide auxiliary aids and services",
     "Sign-language interpreters, large-print materials, assistive technology. Effective for sensory and learning barriers — less so for mobility.",
     B_PA),
    ("4", "Structural alteration",
     "Last resort. Triggers full 2010 ADA design compliance on the altered element. Required only when no other method works.",
     B_2010),
]
for i, (num, head, body, fill) in enumerate(methods):
    y = Inches(2.1 + 1.2*i)
    add_rect(s, Inches(0.55), y, Inches(0.9), Inches(1.05), NAVY)
    add_text(s, Inches(0.55), y, Inches(0.9), Inches(1.05), num,
             size=36, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Georgia")
    add_rect(s, Inches(1.55), y, Inches(11.3), Inches(1.05), CREAM, line=BORDER)
    add_text(s, Inches(1.75), y+Inches(0.1), Inches(11), Inches(0.35), head,
             size=15, bold=True, color=NAVY)
    add_text(s, Inches(1.75), y+Inches(0.45), Inches(11), Inches(0.6), body,
             size=12, color=TEXT)
    std_badge(s, Inches(11.55), y+Inches(0.15), "PA-COMPLIANT" if i < 3 else "TRIGGERS 2010 ADA",
              fill, w=Inches(1.3))
footer(s, 20)
add_notes(s, """
The priority order matters. OCR (and CDE OEO) expect LEAs to consider
the cheaper, non-structural methods first. A district that jumps
straight to "we'll build a new ramp" without considering whether the
section could be moved to an already-accessible building has not done
program-access analysis.

For the LOF narrative, if you find that an LEA has not done a
program-access analysis for an old building, the correct framing is:
"The LEA has not demonstrated program access for [building] under
34 CFR § 104.22. Within 45 days, conduct and document a program-access
analysis identifying which method (relocation, reassignment, auxiliary
aids, or structural alteration) will provide access to the program."
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 21 — The "None / None" rule
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "The \"None / None\" rule",
                            kicker="Slide 21 · Program Access")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.45),
         "If the applicable standard is Program Access, the Accessibility Violation and Corrective Actions columns both read \"None.\" — every time.",
         size=15, italic=True, color=NAVY)

# Show what a row looks like
add_text(s, Inches(0.55), Inches(2.2), Inches(12.3), Inches(0.4),
         "What the LOF table row looks like:",
         size=13, bold=True, color=MUTED)

hdr = ["Area Reviewed", "Applicable Standard", "Accessibility Violation", "Corrective Actions"]
xs = [Inches(0.55), Inches(4.2), Inches(6.5), Inches(9.6)]
ws = [Inches(3.65), Inches(2.3), Inches(3.1), Inches(3.3)]
for i, h in enumerate(hdr):
    add_rect(s, xs[i], Inches(2.65), ws[i], Inches(0.42), NAVY)
    add_text(s, xs[i], Inches(2.65), ws[i], Inches(0.42), h.upper(),
             size=10, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
y = Inches(3.07)
for j in range(4):
    add_rect(s, xs[j], y, ws[j], Inches(1.2), CREAM, line=BORDER)
add_text(s, xs[0]+Inches(0.15), y+Inches(0.1), ws[0], Inches(1.1),
         "CTE Classrooms:\nMain shop building\nConstructed: 1965\nModified: N/A",
         size=11, color=TEXT)
std_badge(s, xs[1]+Inches(0.55), y+Inches(0.4), "PROGRAM ACCESS", B_PA, w=Inches(1.6))
add_text(s, xs[2]+Inches(0.15), y+Inches(0.4), ws[2], Inches(0.6), "None.",
         size=20, bold=True, color=MUTED, italic=True, align=PP_ALIGN.CENTER)
add_text(s, xs[3]+Inches(0.15), y+Inches(0.4), ws[3], Inches(0.6), "None.",
         size=20, bold=True, color=MUTED, italic=True, align=PP_ALIGN.CENTER)

# Why
add_rect(s, Inches(0.55), Inches(4.6), Inches(12.3), Inches(2.5), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(4.6), Inches(0.18), Inches(2.5), RED)
add_text(s, Inches(0.85), Inches(4.75), Inches(12), Inches(0.4),
         "Why this rule exists — and why you must follow it:",
         size=14, bold=True, color=NAVY)
add_bullets(s, Inches(0.85), Inches(5.15), Inches(11.7), Inches(2), [
    "Program Access is a program-level standard, not a dimensional code. There is no \"violation\" of a number that does not exist.",
    "Citing 2010 ADA against a pre-1977 building is legally incorrect — and reversible on appeal.",
    "Program-access concerns are addressed in narrative, NOT in the table. Use the LOF cover letter or recommendation section.",
    "If the LEA has a documented program-access plan, the row is fully compliant — no further action required.",
], size=12, spacing=4)
footer(s, 21)
add_notes(s, """
This rule is etched into the CRR procedures manual and the CRR 20
boilerplate — it is non-negotiable:
"Program Access → ALWAYS None. / None." That language came from us.
Every reviewer must internalize it.

If a reviewer writes a finding against a Program Access row, the
reviewer's lead will reject the LOF and send it back for revision. Save
everyone time by getting this right the first time.

If the reviewer observes a barrier in a pre-1977 building and is
concerned that the program is not accessible, the correct mechanism is
the LOF narrative section — recommend a documented program-access plan,
not a dimensional finding.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 22 — Knowledge check #2 (PA pitfall)
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); set_bg(s, NAVY)
add_rect(s, 0, 0, Inches(0.6), SLIDE_H, GOLD)
add_text(s, Inches(0.85), Inches(0.6), Inches(12), Inches(0.5),
         "KNOWLEDGE CHECK · 02",
         size=12, bold=True, color=GOLD, font="Consolas")
add_text(s, Inches(0.85), Inches(1.1), Inches(12), Inches(0.9),
         "Spot the Program Access pitfall",
         size=32, bold=True, color=WHITE, font="Georgia")

# Scenario card
add_rect(s, Inches(0.85), Inches(2.2), Inches(11.95), Inches(1.5), NAVY_MID)
add_text(s, Inches(1.1), Inches(2.3), Inches(11.5), Inches(1.4),
         "You walk a 1969 ag building. The greenhouse has a 1:8 ramp slope (steeper than the 1:12 limit in 2010 ADA 405.2), the entrance door is 28\" clear width (below the 32\" minimum in 2010 ADA 404.2.3), and no ISA signage is mounted at the door (below the 48\"–60\" requirement in 2010 ADA 703.4.1).\n\nThe LEA has not altered this building since 1969. There are no work orders, no DSA alteration approvals, and no ADA modification record.",
         size=14, color=WHITE, anchor=MSO_ANCHOR.TOP)

add_text(s, Inches(0.85), Inches(3.95), Inches(12), Inches(0.45),
         "QUESTION:",
         size=13, bold=True, color=GOLD, font="Consolas")
add_text(s, Inches(0.85), Inches(4.4), Inches(12), Inches(0.6),
         "What do you write in the Accessibility Violation column for this row?",
         size=18, bold=True, color=WHITE)

opts = [
    ("A", "Cite ADA 2010 405.2 for the ramp, 404.2.3 for the door, 703.4.1 for the signage."),
    ("B", "Cite ANSI A117.1 4.8.2 for the ramp because the building is from 1969."),
    ("C", "\"None.\"  Document the program-access concern in the narrative section instead."),
    ("D", "Cite 1991 ADA because it is the lowest common denominator."),
]
for i, (lbl, txt) in enumerate(opts):
    y = Inches(5.1 + 0.45*i)
    add_rect(s, Inches(0.85), y, Inches(0.4), Inches(0.4), GOLD)
    add_text(s, Inches(0.85), y, Inches(0.4), Inches(0.4), lbl, size=14, bold=True,
             color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    add_text(s, Inches(1.35), y, Inches(11.5), Inches(0.4), txt,
             size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

footer(s, 22)
add_notes(s, """
Correct answer: C.

A 1969 building falls in the Program Access window (≤ June 3, 1977). No
dimensional standard applies. ANSI starts June 4, 1977, so option B is
wrong — 1969 is before that date.

Common wrong answers:
- A is the most tempting wrong answer because the field conditions LOOK
  like 2010 ADA violations. They are not. The 2010 ADA does not apply
  retroactively to a 1969 building.
- B is the second most common wrong answer — reviewers see "old
  building" and pick "old standard." ANSI applies only from June 4,
  1977 onward.

The correct reviewer move: write "None." in the violation column,
"None." in the corrective action column, and address the program-
access concern in the narrative. Recommend the LEA conduct and
document a program-access analysis for the ag building.

This is the most-tested concept in the entire training. Spend 5 minutes
on it.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 23 — ANSI A117.1 (1961) overview
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "ANSI A117.1 (1961, R1971) — overview",
                            kicker="Slide 23 · ANSI")
std_badge(s, Inches(11), Inches(1.0), "ANSI A117.1", B_ANSI,
          text_color=RGBColor(0x71,0x3F,0x12), w=Inches(2.0))

# Lead
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "The first national accessibility standard — and the only one applicable to CA schools built 1977–1991.",
         size=14, italic=True, color=MUTED)

# Left: scope facts
add_text(s, Inches(0.55), Inches(2.05), Inches(6.0), Inches(0.4),
         "Scope facts",
         size=14, bold=True, color=NAVY)
add_bullets(s, Inches(0.55), Inches(2.5), Inches(6.0), Inches(4.5), [
    ("Authority.",   "34 CFR § 104.23(a) — adopts ANSI as the § 504 design standard for new construction."),
    ("Date window.", "June 4, 1977 – January 17, 1991, inclusive."),
    ("Published.",   "American National Standards Institute, 1961, reaffirmed 1971."),
    ("Title.",       "\"American National Standard Specifications for Making Buildings and Facilities Accessible to, and Usable by, the Physically Disabled.\""),
    ("Critical caveat.", "Later versions of ANSI A117.1 (1980, 1986, 1992, 2003, 2009, 2017) DO NOT APPLY to § 504. Cite only the 1961 R1971 version."),
], size=12, spacing=6)

# Right: representative sections
add_rect(s, Inches(7.0), Inches(2.05), Inches(5.85), Inches(5.0), CREAM, line=BORDER)
add_rect(s, Inches(7.0), Inches(2.05), Inches(5.85), Inches(0.5), NAVY)
add_text(s, Inches(7.0), Inches(2.05), Inches(5.85), Inches(0.5),
         "ANSI 1961 — KEY SECTIONS",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
ansi_rows = [
    ("4.2",   "Site development"),
    ("4.3",   "Walks"),
    ("4.4",   "Parking"),
    ("4.5",   "Ramps"),
    ("4.6",   "Entrances"),
    ("4.7",   "Doors"),
    ("4.8",   "Stairs"),
    ("4.8.2", "Ramps — 1:12 maximum running slope"),
    ("4.8.5", "Handrails — 34\"–38\" above ramp / stair surface"),
    ("4.13",  "Toilet rooms"),
    ("4.15",  "Drinking fountains"),
    ("5",     "Special requirements (signage, controls)"),
]
for i, (sec, body) in enumerate(ansi_rows):
    y = Inches(2.65 + 0.36*i)
    add_text(s, Inches(7.2), y, Inches(1.0), Inches(0.35), sec,
             size=11, bold=True, color=NAVY, font="Consolas")
    add_text(s, Inches(8.2), y, Inches(4.5), Inches(0.35), body,
             size=11, color=TEXT)
footer(s, 23)
add_notes(s, """
The "later versions don't apply" rule is critical. ANSI A117.1 has been
revised many times since 1961. The §504 regulation locks in the
1961-R1971 version specifically. If you cite ANSI A117.1 4.8.5 you are
citing the 1961 R1971 text, not the 2017 version.

This trips up architects more than reviewers. Architects working on
school projects today are familiar with the 2017 ANSI standard (the
current technical reference for the IBC). They are NOT designing to the
1961 standard — but for §504 purposes on pre-1991 buildings, that is
the only ANSI version that applies.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 24 — ANSI dimensions
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "ANSI A117.1 — key dimensions reviewers cite",
                            kicker="Slide 24 · ANSI")

# Dimension cards (3x2 grid)
dims = [
    ("Ramp slope",      "Max 1:12  (8.33%)",      "ANSI 4.8.2"),
    ("Handrail height", "34\" – 38\"",            "ANSI 4.8.5"),
    ("Handrail extensions", "12\" min — both ends", "ANSI 4.8.5"),
    ("Door clear width", "32\" minimum",          "ANSI 4.7"),
    ("Toilet stall width", "60\" minimum",        "ANSI 4.13"),
    ("Drinking fountain", "Max 36\" spout height", "ANSI 4.15"),
]
for i, (dim, val, cite) in enumerate(dims):
    col = i % 3
    row = i // 3
    x = Inches(0.55 + 4.2*col)
    y = Inches(1.95 + 2.55*row)
    add_rect(s, x, y, Inches(4.0), Inches(2.35), CREAM, line=BORDER)
    add_rect(s, x, y, Inches(4.0), Inches(0.4), B_ANSI)
    add_text(s, x, y, Inches(4.0), Inches(0.4), cite,
             size=11, bold=True, color=RGBColor(0x71,0x3F,0x12),
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    add_text(s, x+Inches(0.2), y+Inches(0.5), Inches(3.6), Inches(0.5), dim,
             size=14, bold=True, color=NAVY)
    add_text(s, x+Inches(0.2), y+Inches(1.05), Inches(3.6), Inches(1.0), val,
             size=22, bold=True, color=TEXT, font="Georgia")

# Bottom note
add_rect(s, Inches(0.55), Inches(7.0), Inches(12.3), Inches(0.25), GOLD)
add_text(s, Inches(0.55), Inches(6.97), Inches(12.3), Inches(0.3),
         "Many ANSI 1961 dimensions match 1991/2010 ADA — cite the ANSI section for ANSI-era buildings.",
         size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 24)
add_notes(s, """
Reviewers will memorize these six numbers within the first few site
visits. Note that several values are identical across ANSI, 1991, and
2010 — that is intentional, the standards build on each other.

The trap: even though the NUMBER is the same, the CITATION is not. For
a 1985 building, cite ANSI 4.8.5 — not ADA 2010 505.4. The "30-inch
handrail" violation is the same physical problem, but the legal hook
must match the applicable standard.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 25 — ANSI gaps
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "What ANSI 1961 DOESN'T cover",
                            kicker="Slide 25 · ANSI")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.45),
         "ANSI 1961 predates a lot of modern access concepts. If the issue is in this list, ANSI is silent — and there is no finding to issue against an ANSI-era building for these.",
         size=14, italic=True, color=MUTED)

gaps = [
    ("Detectable warnings",     "(truncated domes)",       "Not in ANSI 1961. First appears in ADAAG (1991)."),
    ("Van-accessible parking",  "(96\" aisle width)",       "Not in ANSI 1961. Specified in 1991 ADA 4.6 and 2010 ADA 502.3.3."),
    ("Accessible routes — slope ratio for cross slope", "(2% / 1:50 limit)", "Not codified in ANSI 1961. Added in 1991/2010 ADA."),
    ("ISA tactile signage",     "(48\"–60\" mounting)",    "Not specified in ANSI 1961. Specified in 1991 ADA 4.30 and 2010 ADA 703."),
    ("Communication features",  "(VRS, captioning equipment)", "Outside ANSI 1961 scope — covered later under ADA Title II auxiliary aids analysis."),
    ("Play structures, athletic field accessibility", "", "Not in ANSI 1961. Specified in 2010 ADA Chapter 10."),
]
for i, (topic, paren, body) in enumerate(gaps):
    y = Inches(2.15 + 0.75*i)
    add_rect(s, Inches(0.55), y, Inches(0.18), Inches(0.65), B_PA)
    add_text(s, Inches(0.85), y+Inches(0.05), Inches(5.5), Inches(0.3), topic,
             size=14, bold=True, color=NAVY)
    if paren:
        add_text(s, Inches(0.85), y+Inches(0.35), Inches(5.5), Inches(0.3), paren,
                 size=11, color=MUTED, italic=True)
    add_text(s, Inches(6.5), y, Inches(6.4), Inches(0.65), body,
             size=12, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)

footer(s, 25)
add_notes(s, """
If you observe one of these issues in an ANSI-era (1977–1991) building,
the answer is NOT to cite the 2010 ADA section. ANSI is silent on these
items; therefore there is no violation under §504 for those buildings.

These are the gaps that get reviewers in trouble. New reviewers see a
1980s building with no truncated domes and want to write it up. They
cannot. ANSI 1961 does not require truncated domes.

If the building has been ALTERED to add (for example) new curb ramps in
2018, the altered curb ramps are evaluated under 2010 ADA and detectable
warnings ARE required — but only on the altered element, not on the
original building.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 26 — UFAS overview
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "UFAS — Uniform Federal Accessibility Standards",
                            kicker="Slide 26 · UFAS")
std_badge(s, Inches(11), Inches(1.0), "UFAS",
          B_UFAS, text_color=RGBColor(0x1E,0x3A,0x8A), w=Inches(2.0))

add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.45),
         "The shortest window in §504 — but you still need to know it.",
         size=15, italic=True, color=MUTED)

add_bullets(s, Inches(0.55), Inches(2.15), Inches(6.5), Inches(5), [
    ("Authority.",      "28 CFR § 35.151(a); 34 CFR § 104.23(c); Appendix A to 41 CFR § 101-19.6."),
    ("Date window.",    "January 18, 1991 – January 26, 1992, inclusive.  About 13 months."),
    ("Published.",      "U.S. General Services Administration, 1984."),
    ("Origin.",         "Standardized accessibility for federally-funded construction across all federal agencies (ABA / Architectural Barriers Act, 1968)."),
    ("Departures rule.", "Departures from technical / scoping requirements are permitted when substantially equivalent or greater access is provided. (28 CFR § 35.151(c))."),
    ("Why short.",      "DOJ replaced UFAS as the Title II standard when the 1991 ADA Standards were published. Education adopted the change effective Jan 27, 1992."),
], size=13, spacing=6)

# Right side: at-a-glance
add_rect(s, Inches(7.3), Inches(2.15), Inches(5.55), Inches(4.8), CREAM, line=BORDER)
add_text(s, Inches(7.5), Inches(2.3), Inches(5.3), Inches(0.4),
         "When you'll see it",
         size=14, bold=True, color=NAVY)
add_bullets(s, Inches(7.5), Inches(2.75), Inches(5.3), Inches(4.1), [
    "Very rarely in CA K-12 — most schools were not built or altered in this 13-month window.",
    "Possible candidates: facilities completed in 1991 (calendar year) — check the DSA stamp date.",
    "If you see UFAS in the standard column, double-check the alteration / construction date against the table.",
    "When UFAS does apply, dimensions are similar to 1991 ADA — cite UFAS sections directly (UFAS 4.5 Ramps, 4.8 Doors, etc.)."
], size=12, spacing=8)
footer(s, 26)
add_notes(s, """
UFAS is the standard most reviewers never use. It exists in the table
because §504 specifically incorporates it, but in practice you'll see
fewer than 1 in 50 reviews with a UFAS-controlled element.

If a self-assessment shows a building constructed in 1991, check the
month carefully. January 18 is the cutoff — anything built or
altered on or before January 17, 1991 is ANSI. From January 18, 1991
through January 26, 1992 is UFAS. From January 27, 1992 onward is
1991 ADA.

The departure rule (departures permitted where equivalent access is
provided) is one reason UFAS is less prescriptive than the ADA
standards. Cite this rule when an LEA argues their UFAS-era building
provides equivalent access by alternate means.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 27 — UFAS in K-12 context
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "UFAS in K-12 context — what to look for",
                            kicker="Slide 27 · UFAS")

# Side-by-side: similarities and a few quirks
add_rect(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(5.4), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(0.5), NAVY)
add_text(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(0.5),
         "WHERE UFAS MATCHES 1991 ADA",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(0.75), Inches(2.15), Inches(5.7), Inches(4.7), [
    "Ramp running slope (max 1:12).",
    "Handrail height (34\"–38\").",
    "Door clear width (32\" min).",
    "Toilet seat height (17\"–19\").",
    "Grab bar height (33\"–36\").",
    "Drinking fountain spout height (max 36\")."
], size=13, spacing=8)

add_rect(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(5.4), CREAM, line=BORDER)
add_rect(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(0.5), NAVY)
add_text(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(0.5),
         "UFAS QUIRKS",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(7.05), Inches(2.15), Inches(5.7), Inches(4.7), [
    ("Departure provision.", "UFAS 2.2 — substantial equivalent access is acceptable."),
    ("Different section numbering.", "UFAS organizes by section number, not by use-case. 4.5 = ramps, 4.6 = parking."),
    ("No detectable warnings.", "Predates 1991 ADAAG, no truncated domes required."),
    ("Drinking fountain dual-height.", "Spec includes high and low fountain — pre-1991 ADA refinements."),
], size=12, spacing=6)
footer(s, 27)
add_notes(s, """
Reviewers don't need to memorize UFAS dimensions — they're nearly
identical to 1991 ADA for the most common areas. If you encounter a
UFAS-controlled element, look it up.

The departure provision is the practical difference. An LEA can argue
that a UFAS-era element provides equivalent access through alternate
design. That argument does not extend to 1991 or 2010 ADA elements,
which require strict compliance.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 28 — 1991 ADA / ADAAG overview
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "1991 ADA Standards / ADAAG — overview",
                            kicker="Slide 28 · 1991 ADA")
std_badge(s, Inches(10.5), Inches(1.0), "1991 ADA",
          B_1991, text_color=RGBColor(0x06,0x4E,0x3B), w=Inches(2.5))

add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.45),
         "The first generation of ADA dimensional standards — controls 18+ years of CA school construction.",
         size=14, italic=True, color=MUTED)

add_bullets(s, Inches(0.55), Inches(2.1), Inches(6.5), Inches(5), [
    ("Authority.",   "28 CFR Part 36, Appendix D; adopted into §504 via 34 CFR § 104.23."),
    ("Date window.", "January 27, 1992 – September 14, 2010. Overlap with 2010 ADA from Sep 15, 2010 to Mar 14, 2012."),
    ("Published.",   "ADA Accessibility Guidelines for Buildings and Facilities (ADAAG), Department of Justice / Architectural and Transportation Barriers Compliance Board, 1991."),
    ("Structure.",   "Sections 4.1–4.34: building elements. Section 5 onwards: special occupancy types (restaurants, libraries, transient lodging, etc.)."),
    ("Key advance.", "First standard to include detectable warnings, van-accessible parking, ISA signage requirements, and comprehensive toilet-room geometry."),
], size=13, spacing=6)

# Right column: subsections
add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(4.85), CREAM, line=BORDER)
add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45), NAVY)
add_text(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45),
         "1991 ADAAG — STRUCTURE",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
rows1991 = [
    ("4.1",  "Minimum requirements / scoping"),
    ("4.3",  "Accessible routes"),
    ("4.5",  "Ground & floor surfaces"),
    ("4.6",  "Parking & passenger loading"),
    ("4.8",  "Ramps"),
    ("4.13", "Doors"),
    ("4.16", "Water closets"),
    ("4.17", "Toilet stalls"),
    ("4.25", "Storage / coat hooks"),
    ("4.30", "Signage"),
    ("4.32", "Fixed seating and tables"),
]
for i, (sec, body) in enumerate(rows1991):
    y = Inches(2.6 + 0.38*i)
    add_text(s, Inches(7.5), y, Inches(0.8), Inches(0.36), sec,
             size=10, bold=True, color=NAVY, font="Consolas")
    add_text(s, Inches(8.4), y, Inches(4.4), Inches(0.36), body,
             size=11, color=TEXT)
footer(s, 28)
add_notes(s, """
1991 ADA is what controls the bulk of California school construction
from the 1990s and 2000s — a huge chunk of the CRR docket. Reviewers
will cite this standard most days.

ADAAG (the Accessibility Guidelines) is the original 1991 document.
Some practitioners say "1991 ADA Standards" — same document. The DOJ
adopted ADAAG as the 1991 Standards.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 29 — 1991 ADA key citations
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "1991 ADA — key citations reviewers use",
                            kicker="Slide 29 · 1991 ADA")

hd = ["Element", "Requirement", "Cite"]
xs = [Inches(0.55), Inches(5.0), Inches(10.3)]
ws = [Inches(4.4), Inches(5.25), Inches(2.55)]
for i, h in enumerate(hd):
    add_rect(s, xs[i], Inches(1.55), ws[i], Inches(0.4), NAVY)
    add_text(s, xs[i], Inches(1.55), ws[i], Inches(0.4), h.upper(),
             size=10, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")

rows1991k = [
    ("Door clear width",       "32 inches minimum",                            "4.13.5"),
    ("Ramp running slope",     "1:12 maximum (8.33%)",                         "4.8.2"),
    ("Ramp cross slope",       "2% maximum (1:50)",                            "4.8.6"),
    ("Handrail height",        "34\" – 38\" above ramp / stair nosing",        "4.8.5"),
    ("Handrail extensions",    "12\" minimum, top and bottom",                 "4.8.5"),
    ("ISA signage mounting",   "48\" – 60\" floor to centerline of tactile",    "4.30.6"),
    ("Grab bar height",        "33\" – 36\" above floor",                       "4.26.2"),
    ("Toilet seat height",     "17\" – 19\" above floor",                       "4.16.3"),
    ("Accessible stall width", "60\" minimum",                                  "4.17.3"),
    ("Coat hooks",             "48\" max forward reach",                        "4.25.3"),
    ("Knee clearance",         "27\"H × 30\"W × 19\"D minimum",                "4.32.3"),
    ("Counter height",         "34\" maximum",                                  "7.2"),
    ("Drinking fountain",      "36\" max spout height",                         "4.15.2"),
]
for i, (el, req, cite) in enumerate(rows1991k):
    y = Inches(1.95 + 0.38*i)
    bg = WHITE if i % 2 == 0 else CREAM
    for j in range(3):
        add_rect(s, xs[j], y, ws[j], Inches(0.38), bg, line=BORDER)
    add_text(s, xs[0]+Inches(0.1), y, ws[0], Inches(0.38), el,
             size=11, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, xs[1]+Inches(0.1), y, ws[1], Inches(0.38), req,
             size=11, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, xs[2]+Inches(0.1), y, ws[2], Inches(0.38), f"1991 ADA {cite}",
             size=10, color=NAVY, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
footer(s, 29)
add_notes(s, """
This is a reference slide. Reviewers will refer back to it constantly
during a 1991 ADA-era review. Many of these numbers carry forward
unchanged to 2010 ADA — but the citation numbers do not (see next
slide).

Tell the class: when you cite, always include the standard year AND the
section. "ADA 1991 4.13.5" — not just "4.13.5."
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 30 — Knowledge check #3 (1991 ADA finding)
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); set_bg(s, NAVY)
add_rect(s, 0, 0, Inches(0.6), SLIDE_H, GOLD)
add_text(s, Inches(0.85), Inches(0.6), Inches(12), Inches(0.5),
         "KNOWLEDGE CHECK · 03",
         size=12, bold=True, color=GOLD, font="Consolas")
add_text(s, Inches(0.85), Inches(1.1), Inches(12), Inches(0.9),
         "Cite the right standard — 1991 ADA",
         size=30, bold=True, color=WHITE, font="Georgia")

add_rect(s, Inches(0.85), Inches(2.2), Inches(11.95), Inches(1.7), NAVY_MID)
add_text(s, Inches(1.1), Inches(2.35), Inches(11.5), Inches(1.5),
         "Madera HS culinary classroom built 1998. Last alteration: 2005 paint and HVAC (not accessibility-affecting).\n\nField measurements: counter prep surface at 38\" above floor; ISA signage at 44\" to centerline of tactile characters; door clear width 31\".\n\nWrite the violation row.",
         size=14, color=WHITE)

add_text(s, Inches(0.85), Inches(4.0), Inches(12), Inches(0.4),
         "WRITE OUT YOUR ANSWER — ALL THREE VIOLATIONS WITH CITES:",
         size=12, bold=True, color=GOLD, font="Consolas")

# Answer template box
add_rect(s, Inches(0.85), Inches(4.5), Inches(11.95), Inches(2.5), NAVY_MID, line=GOLD)
add_text(s, Inches(1.1), Inches(4.65), Inches(11.5), Inches(2.4),
         "Standard applies: ____________________________________\n\nViolation 1: ____________________________________________\n\nViolation 2: ____________________________________________\n\nViolation 3: ____________________________________________",
         size=14, color=GOLD_LT, font="Consolas")
footer(s, 30)
add_notes(s, """
Correct answers:

Standard applies: 1991 ADA (built 1998, no accessibility-affecting
alterations).

Violation 1 (counter): Counter prep surface measured at 38" above the
floor exceeds the 34" maximum specified in 1991 ADA 7.2.
Cite: 1991 ADA 7.2 (Counters).

Violation 2 (signage): ISA signage centerline measured at 44" above
the floor is below the required 48"–60" mounting height range
specified in 1991 ADA 4.30.6.
Cite: 1991 ADA 4.30.6 (Mounting Location and Height).

Violation 3 (door): Door clear width measured at 31" is below the
32" minimum specified in 1991 ADA 4.13.5.
Cite: 1991 ADA 4.13.5 (Clear Width).

Common errors:
- Citing 2010 ADA. The building is 1998, no alterations — 2010 ADA
  does not retroactively apply.
- Citing only one cite when multiple violations exist. Each
  measurement is its own line in the violation column.
- Forgetting the standard year. "4.13.5" is ambiguous — is it 1991
  ADA or 2010 ADA? Always include the year.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 31 — 2010 ADA overview
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "2010 ADA Standards — the current default",
                            kicker="Slide 31 · 2010 ADA")
std_badge(s, Inches(10.5), Inches(1.0), "2010 ADA",
          B_2010, text_color=RGBColor(0x7F,0x1D,0x1D), w=Inches(2.5))

add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.45),
         "This is the standard that applies to anything built or altered today. Most modern CRR findings cite 2010 ADA sections.",
         size=14, italic=True, color=MUTED)

add_bullets(s, Inches(0.55), Inches(2.1), Inches(12.3), Inches(5.0), [
    ("Authority.",   "28 CFR Part 35, Appendix B (Title II); 28 CFR Part 36, Appendix A (Title III). Adopted into §504 via 34 CFR § 104.23."),
    ("Date window.", "Required for facilities designed for construction or alteration on or after March 15, 2012.  Available as an election from September 15, 2010."),
    ("Published.",   "U.S. Access Board / DOJ — final rule September 15, 2010."),
    ("Structure.",   "Reorganized from 1991 ADAAG. Renumbered: Chapters 1 (application), 2 (scoping), 3 (building blocks), 4 (accessible routes), 5 (general site & building elements), 6 (plumbing), 7 (communication), 8 (special rooms), 9 (built-in elements), 10 (recreation)."),
    ("Safe harbor.", "Existing elements that complied with 1991 ADA and remain unaltered are deemed compliant under 2010 ADA — no retrofit required unless the element is altered."),
    ("Key advances.", "Comprehensive play-area scoping (Chapter 10), updated reach ranges, refined clear floor space and turning space, expanded signage and assistive listening, recreational facility requirements."),
], size=13, spacing=6)
footer(s, 31)
add_notes(s, """
The 2010 standard is the most-cited standard on the CRR docket — nearly
every alteration or new build since 2012 falls here, and the safe-harbor
rule sometimes works in reverse (an LEA argues "we already complied with
1991 ADA and have not altered the element, so 2010 ADA does not require
retrofit"). The reviewer's response: confirm the 1991 compliance with
field measurements; if it complies, no finding; if not, the element was
never compliant and the finding stands under 1991.

Make sure reviewers understand "safe harbor" doesn't mean "exempt from
ADA." It means "deemed-compliant-as-of-1991 until next alteration."
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 32 — 2010 ADA reach ranges
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "2010 ADA — reach ranges",
                            kicker="Slide 32 · 2010 ADA")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Reach ranges control where operable parts (controls, switches, dispensers, hooks) can be mounted.",
         size=14, italic=True, color=MUTED)

# Two big visual cards — forward and side reach
def reach_card(x, y, title, cite, items):
    add_rect(s, x, y, Inches(5.85), Inches(4.5), CREAM, line=BORDER)
    add_rect(s, x, y, Inches(5.85), Inches(0.5), NAVY)
    add_text(s, x, y, Inches(5.85), Inches(0.5), title.upper(),
             size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    add_text(s, x+Inches(0.2), y+Inches(0.6), Inches(5.5), Inches(0.4), cite,
             size=11, bold=True, color=NAVY, font="Consolas")
    add_bullets(s, x+Inches(0.2), y+Inches(1.05), Inches(5.5), Inches(3.4),
                items, size=12, spacing=8)

reach_card(Inches(0.55), Inches(2.1), "Forward Reach", "2010 ADA 308.2",
    [
        ("Unobstructed forward reach:", "15\" min — 48\" max above the floor."),
        ("Obstructed high forward reach (≤ 20\" deep obstruction):", "48\" max."),
        ("Obstructed high forward reach (20\"–25\" deep):", "44\" max."),
        ("Obstructed reach depths > 25\":", "not permitted — find another location."),
    ])

reach_card(Inches(6.95), Inches(2.1), "Side Reach", "2010 ADA 308.3",
    [
        ("Unobstructed side reach:", "15\" min — 48\" max above the floor."),
        ("Side reach over an obstruction (≤ 10\" deep, ≤ 34\" high):", "48\" max."),
        ("Obstructed side reach (10\"–24\" deep):", "46\" max."),
        ("Obstructed reach depths > 24\":", "not permitted."),
    ])

# Common findings strip
add_rect(s, Inches(0.55), Inches(6.7), Inches(12.3), Inches(0.45), B_2010, line=RED)
add_text(s, Inches(0.85), Inches(6.7), Inches(12), Inches(0.45),
         "Most common reach violations:  coat hooks > 48\"  ·  paper towel dispensers > 48\"  ·  switches behind deep counters",
         size=12, bold=True, color=RED, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 32)
add_notes(s, """
Reach ranges sound simple but are easy to mis-measure on site. The
measurement is to the OPERABLE PART of the control — the lever on a
faucet, the slot in a paper-towel dispenser, the centerline of a switch
plate. Reviewers should bring a tape measure long enough to measure
from floor to control (typically 5–6 feet).

In CTE classrooms, watch for items mounted on the rear of deep counters
(welding setups, lab benches). These often exceed the 24-25" obstructed
reach depths and can never be reached.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 33 — 2010 ADA clear floor space & turning space
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "2010 ADA — clear floor space & turning space",
                            kicker="Slide 33 · 2010 ADA")

# Four spec cards
specs = [
    ("Clear floor space",   "30\" × 48\" minimum",     "2010 ADA 305",
     "The footprint a wheelchair user occupies. Required at every accessible element."),
    ("Turning space — circle", "60\" diameter",        "2010 ADA 304.3.1",
     "Permits a 180° turn. Used in toilet rooms, kitchens, and dead-end spaces."),
    ("Turning space — T-shape", "60\" × 60\" with 36\"-wide arms", "2010 ADA 304.3.2",
     "Alternate to the circle. Required minimum arm length 36\", base 60\"."),
    ("Knee & toe clearance", "27\"H × 30\"W × 19\"D", "2010 ADA 306",
     "Required under counters, lavatories, drinking fountains, and other elements designed for forward approach."),
]
for i, (head, val, cite, body) in enumerate(specs):
    col = i % 2
    row = i // 2
    x = Inches(0.55 + 6.2*col)
    y = Inches(1.85 + 2.6*row)
    add_rect(s, x, y, Inches(6.0), Inches(2.45), CREAM, line=BORDER)
    add_rect(s, x, y, Inches(0.18), Inches(2.45), B_2010)
    add_text(s, x+Inches(0.35), y+Inches(0.15), Inches(5.6), Inches(0.4), head,
             size=14, bold=True, color=NAVY)
    add_text(s, x+Inches(0.35), y+Inches(0.6), Inches(5.6), Inches(0.55), val,
             size=22, bold=True, color=TEXT, font="Georgia")
    add_text(s, x+Inches(0.35), y+Inches(1.2), Inches(5.6), Inches(0.4), cite,
             size=11, bold=True, color=NAVY, font="Consolas")
    add_text(s, x+Inches(0.35), y+Inches(1.55), Inches(5.6), Inches(0.85), body,
             size=11, color=TEXT, italic=True)

footer(s, 33)
add_notes(s, """
Clear floor space is the single most-tested spec because almost every
accessible element requires one. In CTE classrooms, clear floor space
in front of welding booths, lab benches, and cooktops is often blocked
by equipment, storage carts, or task lighting. Measure with the room
in its "real-world" configuration, not staged for the review.

Turning space requirements scale up in CTE labs — accessible workstations
often need both forward approach (with knee/toe clearance under) and a
nearby turning space.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 34 — 2010 ADA ramps & slopes
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "2010 ADA — ramps, slopes, and curb ramps",
                            kicker="Slide 34 · 2010 ADA")

# Diagram-ish: ramp running slope, cross slope, landing
add_rect(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(2.5), CREAM, line=BORDER)
# Ramp silhouette
ramp = s.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(2.5), Inches(2.8), Inches(4.5), Inches(1.0))
ramp.fill.solid(); ramp.fill.fore_color.rgb = NAVY_MID
ramp.line.color.rgb = NAVY; ramp.shadow.inherit = False
# Landing
add_rect(s, Inches(7.0), Inches(2.8), Inches(1.5), Inches(1.0), GOLD, line=BORDER)
add_text(s, Inches(7.0), Inches(2.8), Inches(1.5), Inches(1.0), "LANDING\n60\"×60\"",
         size=10, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
# Slope label
add_text(s, Inches(2.5), Inches(2.3), Inches(4.5), Inches(0.4), "Running slope 1:12 max  ·  Cross slope 1:48 (2.1%) max",
         size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
add_text(s, Inches(2.5), Inches(3.85), Inches(4.5), Inches(0.3), "Rise max 30\" between landings",
         size=11, color=MUTED, align=PP_ALIGN.CENTER, italic=True, font="Consolas")
# Handrail callout
add_text(s, Inches(8.7), Inches(2.0), Inches(4), Inches(0.4),
         "Handrails 34\"–38\"  ·  12\" extensions both ends",
         size=11, bold=True, color=NAVY, font="Consolas")
add_text(s, Inches(8.7), Inches(2.4), Inches(4), Inches(0.3),
         "Required when rise > 6\" or run > 72\"",
         size=10, color=MUTED, italic=True)

# Citation cards
cites = [
    ("Ramp running slope", "Max 1:12 (8.33%)",     "2010 ADA 405.2"),
    ("Ramp cross slope",   "Max 2% (1:48)",        "2010 ADA 405.3"),
    ("Landings",           "60\" min length each end",  "2010 ADA 405.7"),
    ("Curb ramp slope",    "Max 1:12 running, 2% cross",  "2010 ADA 406.1"),
    ("Handrails",          "34\"–38\", 12\" extensions",  "2010 ADA 505.4 / .10"),
    ("Edge protection",    "4\" min curb / wall",  "2010 ADA 405.9"),
]
for i, (head, val, cite) in enumerate(cites):
    col = i % 3
    row = i // 3
    x = Inches(0.55 + 4.2*col)
    y = Inches(4.25 + 1.4*row)
    add_rect(s, x, y, Inches(4.0), Inches(1.25), WHITE, line=BORDER)
    add_text(s, x+Inches(0.2), y+Inches(0.1), Inches(3.6), Inches(0.3),
             cite, size=11, bold=True, color=NAVY, font="Consolas")
    add_text(s, x+Inches(0.2), y+Inches(0.4), Inches(3.6), Inches(0.4),
             head, size=13, bold=True, color=TEXT)
    add_text(s, x+Inches(0.2), y+Inches(0.8), Inches(3.6), Inches(0.4),
             val, size=12, color=TEXT)
footer(s, 34)
add_notes(s, """
Ramp findings are among the most common in CRR 20 / 21. Common errors
the reviewer will see:

1. Running slope just over 1:12. Often 1:10 or 1:11. The LEA says
   "close enough" — it isn't. Cite the section.
2. Cross slope creep. The 2% (1:48) tolerance is small and often
   exceeded by drainage requirements. Measure across the ramp width.
3. Missing landings — especially at door swings or directional
   changes. Need 60" minimum, level (max 2% in any direction).
4. Handrail extensions missing or wrong direction. Need 12" beyond
   top and bottom of the ramp run, parallel to the floor at the
   bottom.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 35 — 2010 ADA signage (ISA)
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "2010 ADA — signage and the ISA",
                            kicker="Slide 35 · 2010 ADA")

# Left: requirements
add_text(s, Inches(0.55), Inches(1.55), Inches(6.5), Inches(0.4),
         "International Symbol of Accessibility (ISA) — required at the most-used elements.",
         size=14, italic=True, color=MUTED)
add_bullets(s, Inches(0.55), Inches(2.05), Inches(6.5), Inches(5.0), [
    ("Mounting height.", "48\"–60\" from finished floor to BASELINE of tactile characters. (2010 ADA 703.4.1)"),
    ("Location.",        "Latch side of door, 18\" centered (max) from door frame. Single-leaf doors only. (2010 ADA 703.4.2)"),
    ("Tactile characters.", "Raised 1/32\" min; sans-serif uppercase. Grade 2 Braille required adjacent. (2010 ADA 703.2 / 703.3)"),
    ("ISA appearance.",  "White on blue, or contrasting colors. 6\" minimum size at room entries. (2010 ADA 703.7.2.1)"),
    ("Required at:",     "Accessible parking, accessible entrances, accessible restrooms, accessible routes when not obvious."),
    ("Directional signs.", "Where accessible route differs from main route — direct users to accessible alternative. (2010 ADA 216.3 / 216.6)"),
], size=12, spacing=6)

# Right: visual mock-up of an ISA sign with measurements
add_rect(s, Inches(7.5), Inches(2.05), Inches(5.4), Inches(4.95), CREAM, line=BORDER)
add_text(s, Inches(7.5), Inches(2.15), Inches(5.4), Inches(0.4), "Sample ISA placement",
         size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
# Door
add_rect(s, Inches(8.4), Inches(2.75), Inches(2.4), Inches(3.8), WHITE, line=NAVY)
add_text(s, Inches(8.4), Inches(2.75), Inches(2.4), Inches(0.3), "DOOR",
         size=10, color=NAVY, align=PP_ALIGN.CENTER, font="Consolas")
# Sign on latch side
add_rect(s, Inches(11.0), Inches(4.05), Inches(1.0), Inches(1.0), GOLD, line=NAVY)
add_text(s, Inches(11.0), Inches(4.05), Inches(1.0), Inches(1.0), "ISA",
         size=14, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
# Vertical dimension line
add_rect(s, Inches(12.3), Inches(4.55), Inches(0.02), Inches(2.0), MUTED)
add_text(s, Inches(11.95), Inches(4.5), Inches(0.6), Inches(0.3), "60\"",
         size=10, color=NAVY, font="Consolas")
add_text(s, Inches(11.95), Inches(6.3), Inches(0.6), Inches(0.3), "48\"",
         size=10, color=NAVY, font="Consolas")
add_text(s, Inches(11.95), Inches(6.6), Inches(1.0), Inches(0.3), "FLOOR",
         size=8, color=MUTED, font="Consolas")
# Latch-side note
add_text(s, Inches(7.5), Inches(6.6), Inches(5.4), Inches(0.4),
         "Latch side · 18\" max to frame · centered between 48\"–60\"",
         size=10, color=MUTED, align=PP_ALIGN.CENTER, italic=True, font="Consolas")

footer(s, 35)
add_notes(s, """
Signage is the most-cited element in CRR 20 because it's checked at
every door. Common failures:
- Signs mounted at 66" (too high — over 60" to baseline of tactile).
- Signs on the hinge side instead of the latch side.
- Signs on a double-leaf door (not permitted — only single-leaf).
- Missing Braille adjacent to tactile characters.
- ISA size below 6" at room entries.

Reviewers should measure the BASELINE of the tactile characters, not
the top of the sign. A common mistake is measuring to the top of the
sign panel; that produces a wrong reading and a wrong finding.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 36 — 2010 ADA restroom essentials
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "2010 ADA — restroom essentials",
                            kicker="Slide 36 · 2010 ADA")

# Two column: water closet & stall / accessories
add_rect(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(5.4), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(0.5), NAVY)
add_text(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(0.5),
         "WATER CLOSET & STALL",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(0.75), Inches(2.15), Inches(5.7), Inches(4.8), [
    ("Stall width (accessible)", "60\" minimum  ·  604.8.1.1"),
    ("Stall depth",              "56\" wall-mount / 59\" floor-mount min  ·  604.8.1.1"),
    ("Toilet seat height",       "17\"–19\" above floor  ·  604.4"),
    ("Centerline of WC",         "16\"–18\" from sidewall  ·  604.2"),
    ("Side grab bar",            "42\" min, 12\" from rear wall  ·  604.5.1"),
    ("Rear grab bar",            "36\" min  ·  604.5.2"),
    ("Grab bar height",          "33\"–36\" above floor  ·  609.4"),
    ("Flush controls",           "Open side of WC  ·  604.6"),
    ("TP dispenser",             "15\"–48\" above floor; 7\"–9\" in front of WC  ·  604.7"),
], size=11, spacing=4)

add_rect(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(5.4), CREAM, line=BORDER)
add_rect(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(0.5), NAVY)
add_text(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(0.5),
         "LAV, MIRROR, ACCESSORIES",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(7.05), Inches(2.15), Inches(5.7), Inches(4.8), [
    ("Lavatory rim height",      "Max 34\" above floor  ·  606.3"),
    ("Knee clearance under lav", "27\"H × 30\"W × 11\"D min  ·  606.2 / 306"),
    ("Lav faucet controls",      "Operable with closed fist; 5 lbf max  ·  309.4"),
    ("Hot water / drain pipes",  "Insulated / configured to prevent contact  ·  606.5"),
    ("Mirror bottom edge",       "Max 40\" above floor  ·  603.3"),
    ("Coat hook",                "Max 48\" forward reach  ·  308.2.1 / 604.8.3"),
    ("Soap dispenser",           "Reach within 48\" max  ·  308"),
    ("Paper towel dispenser",    "Operable parts within 48\" max  ·  308"),
    ("Door swing into stall",    "Permitted only if 60\"×60\" maneuvering space remains  ·  604.8.1.2"),
], size=11, spacing=4)
footer(s, 36)
add_notes(s, """
Restrooms account for the largest share of CRR 20 findings — they have
the most elements and the tightest dimensional tolerances. New
reviewers should plan to spend 20-30 minutes per restroom.

Order of operations for measuring a restroom: walk in, find the
accessible stall, measure stall dimensions, measure grab bar location
and height, measure TP dispenser, measure lavatory height + knee
clearance, measure mirror bottom edge, measure soap/towel reach,
measure door clear width on the way out.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 37 — 2010 ADA parking & accessible routes
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "2010 ADA — parking & accessible routes",
                            kicker="Slide 37 · 2010 ADA")

# Parking spec
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Parking lot scoping is the first thing a reviewer evaluates — it determines whether students arrive at the building at all.",
         size=14, italic=True, color=MUTED)

# Parking scoping table
hd = ["Total spaces in lot", "Required accessible", "Required van-accessible"]
xs = [Inches(0.55), Inches(4.6), Inches(8.65)]
ws = [Inches(4.05), Inches(4.05), Inches(4.2)]
for i, h in enumerate(hd):
    add_rect(s, xs[i], Inches(2.1), ws[i], Inches(0.4), NAVY)
    add_text(s, xs[i], Inches(2.1), ws[i], Inches(0.4), h.upper(),
             size=10, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
pk = [
    ("1 – 25",     "1",  "1"),
    ("26 – 50",    "2",  "1"),
    ("51 – 75",    "3",  "1"),
    ("76 – 100",   "4",  "1"),
    ("101 – 150",  "5",  "1"),
    ("151 – 200",  "6",  "1"),
    ("201 – 300",  "7",  "1"),
    ("301 – 400",  "8",  "1"),
    ("401 – 500",  "9",  "2"),
]
for i, (a, b, c) in enumerate(pk):
    y = Inches(2.5 + 0.36*i)
    bg = WHITE if i % 2 == 0 else CREAM
    for j, val in enumerate([a, b, c]):
        add_rect(s, xs[j], y, ws[j], Inches(0.36), bg, line=BORDER)
        add_text(s, xs[j], y, ws[j], Inches(0.36), val,
                 size=11, color=TEXT, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE, font="Consolas")

# Right-side notes
add_text(s, Inches(0.55), Inches(5.8), Inches(12.3), Inches(0.4),
         "Critical dimensions:",
         size=14, bold=True, color=NAVY)
add_bullets(s, Inches(0.55), Inches(6.2), Inches(12.3), Inches(1.0), [
    ("Standard accessible aisle:", "60\" minimum width  ·  2010 ADA 502.3.1"),
    ("Van-accessible aisle:",      "96\" minimum width  ·  2010 ADA 502.3.3"),
    ("Surface slope (parking + aisle):", "Max 2% in any direction  ·  2010 ADA 502.4"),
    ("ISA signage:",               "Mounted at each accessible space; min 60\" to bottom of sign  ·  2010 ADA 502.6"),
], size=11, spacing=4, bullet="·")
footer(s, 37)
add_notes(s, """
Parking is foundational. If parking is wrong, the entire accessible-route
analysis fails — the student can't get from the lot to the building.

The 502.4 surface slope rule (2% maximum in ANY direction) is one of the
most-violated specs. Drainage requirements often push parking slopes
above 2%. Measure with a digital level in both directions; record the
worst-case slope.

The 1-in-N pattern continues above 500 spaces — 1 van-accessible per 6
accessible spaces. Most CA high schools have 100-300 spaces in the main
lot.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 38 — Knowledge check #4
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); set_bg(s, NAVY)
add_rect(s, 0, 0, Inches(0.6), SLIDE_H, GOLD)
add_text(s, Inches(0.85), Inches(0.6), Inches(12), Inches(0.5),
         "KNOWLEDGE CHECK · 04",
         size=12, bold=True, color=GOLD, font="Consolas")
add_text(s, Inches(0.85), Inches(1.1), Inches(12), Inches(0.9),
         "2010 ADA — six measurements",
         size=30, bold=True, color=WHITE, font="Georgia")

add_text(s, Inches(0.85), Inches(2.15), Inches(12), Inches(0.5),
         "Building constructed 2018. For each measurement, mark COMPLIANT or VIOLATION.",
         size=14, color=GOLD_LT, italic=True)

# Six measurement rows
meas = [
    ("Parking surface slope: 2.4% in one direction",                       "Violation"),
    ("Coat hook in accessible restroom: 47\" forward reach",                "Compliant"),
    ("Toilet seat height: 16\" above floor",                                "Violation"),
    ("Door clear width: 34\"",                                              "Compliant"),
    ("Ramp running slope: 1:14",                                            "Compliant"),
    ("ISA signage baseline of tactile characters: 62\" above floor",        "Violation"),
]
for i, (sc, ans) in enumerate(meas):
    y = Inches(2.85 + 0.6*i)
    add_rect(s, Inches(0.85), y, Inches(8.5), Inches(0.5), NAVY_MID)
    add_text(s, Inches(1.05), y, Inches(8.3), Inches(0.5), sc,
             size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    # Empty boxes for compliant / violation
    add_rect(s, Inches(9.5), y, Inches(1.6), Inches(0.5), CREAM, line=BORDER)
    add_text(s, Inches(9.5), y, Inches(1.6), Inches(0.5), "COMPLIANT",
             size=10, color=MUTED, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    add_rect(s, Inches(11.2), y, Inches(1.6), Inches(0.5), CREAM, line=BORDER)
    add_text(s, Inches(11.2), y, Inches(1.6), Inches(0.5), "VIOLATION",
             size=10, color=MUTED, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")

footer(s, 38)
add_notes(s, """
Answers:
1. Parking surface slope 2.4% — VIOLATION. 2010 ADA 502.4 caps surface
   slope at 2% in any direction. Cite 502.4.
2. Coat hook at 47" — COMPLIANT. 2010 ADA 308.2.1 allows up to 48"
   forward reach.
3. Toilet seat 16" — VIOLATION. 2010 ADA 604.4 requires 17"–19".
   Cite 604.4.
4. Door clear width 34" — COMPLIANT. 2010 ADA 404.2.3 requires
   minimum 32". 34" exceeds the minimum.
5. Ramp slope 1:14 — COMPLIANT. 2010 ADA 405.2 requires max 1:12,
   meaning slope cannot be STEEPER than 1:12. 1:14 is shallower
   (gentler) and is fine.
6. ISA tactile baseline 62" — VIOLATION. 2010 ADA 703.4.1 requires
   48"–60" baseline. 62" exceeds the maximum.

Most common errors:
- #5 confuses students. "1:14 is bigger than 1:12, so it's worse, right?"
  No — ramp slope is rise over run. 1:14 means 1" rise per 14" run, which
  is gentler than 1" rise per 12" run. The MAXIMUM steepness is 1:12.
- #4 is also commonly missed. Minimums work in the opposite direction
  from maximums. 34" exceeds the 32" minimum, so it complies.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 39 — Measurements ref: parking & routes
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Quick reference — parking & accessible routes",
                            kicker="Slide 39 · Reference")

def quickref(slide, x, y, w, h, title, rows):
    add_rect(slide, x, y, w, h, CREAM, line=BORDER)
    add_rect(slide, x, y, w, Inches(0.4), NAVY)
    add_text(slide, x, y, w, Inches(0.4), title.upper(),
             size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    for i, (item, spec, cite) in enumerate(rows):
        ry = y + Inches(0.45) + Inches(0.4*i)
        add_text(slide, x+Inches(0.15), ry, w-Inches(2.0), Inches(0.35),
                 item, size=11, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, x+w-Inches(3.5), ry, Inches(1.8), Inches(0.35),
                 spec, size=11, color=TEXT, anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
        add_text(slide, x+w-Inches(1.7), ry, Inches(1.6), Inches(0.35),
                 cite, size=10, color=MUTED, anchor=MSO_ANCHOR.MIDDLE, font="Consolas",
                 align=PP_ALIGN.RIGHT)

quickref(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(5.4), "Parking",
         [("Std accessible aisle",     "60\" min",          "502.3.1"),
          ("Van accessible aisle",     "96\" min",          "502.3.3"),
          ("Surface slope (any dir.)", "2% max",            "502.4"),
          ("Vertical clearance",       "98\" min van",      "502.5"),
          ("ISA sign — bottom",        "60\" min from grade", "502.6"),
          ("Marking",                  "Painted; ISA visible", "502.6"),
          ("Path from aisle to route", "Connects to acc. rt.", "502.3"),
          ("Curb ramps if separate",   "1:12 max",          "406.1"),
          ("Detectable warnings",      "At curb ramps + bus stops", "705"),
          ("Re-striping = alteration", "Triggers 2010 ADA", "Title II"),
          ("Quantity (1–25 lot)",      "1 acc + 1 van",     "208.2"),
          ("Quantity (26–50 lot)",     "2 acc + 1 van",     "208.2"),
         ])
quickref(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(5.4), "Accessible Routes",
         [("Minimum width",            "36\" (32\" at point)", "403.5"),
          ("Width at 90° turn",        "48\" if route < 48\"", "403.5.2"),
          ("Running slope (route)",    "1:20 max",          "403.3"),
          ("Cross slope",              "1:48 (2.1%) max",   "403.3"),
          ("Passing spaces",           "60\"×60\" every 200'", "403.5.3"),
          ("Surface — firm, stable",   "No loose gravel",   "302.1"),
          ("Surface openings",         "< ½\" any direction", "302.3"),
          ("Vertical change ≤ ¼\"",    "OK unbeveled",      "303.2"),
          ("Vertical change ¼\"–½\"",  "Beveled 1:2",       "303.3"),
          ("Vertical change > ½\"",    "Ramp required",     "303.4"),
          ("Protruding objects",       "≤ 4\" if 27\"–80\" AFF", "307.2"),
          ("Headroom",                 "80\" min clear",    "307.4"),
         ])
add_text(s, Inches(0.55), Inches(7.05), Inches(12.3), Inches(0.25),
         "All cites are 2010 ADA. Match standard to construction date before using these numbers.",
         size=10, color=MUTED, italic=True, align=PP_ALIGN.CENTER, font="Consolas")
footer(s, 39)
add_notes(s, """
Hand out a one-page print of this slide. Reviewers will tab back to it
constantly. Note the cites are 2010 ADA only — for 1991 ADA buildings,
look up the corresponding section (most carry over with similar numbers
but different organization).
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 40 — Measurements ref: doors & ramps
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Quick reference — doors, ramps & stairs",
                            kicker="Slide 40 · Reference")
quickref(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(5.4), "Doors",
         [("Clear width",          "32\" min",          "404.2.3"),
          ("Maneuvering — pull",   "18\" beyond latch", "404.2.4.1"),
          ("Maneuvering — push",   "12\" beyond latch w/ closer", "404.2.4.1"),
          ("Threshold",            "½\" max, beveled",  "404.2.5"),
          ("Hardware operation",   "Closed-fist usable; 5 lbf", "404.2.7"),
          ("Hardware height",      "34\"–48\" above floor", "404.2.7"),
          ("Opening force int.",   "5 lbf max",         "404.2.9"),
          ("Closing speed",        "≥ 5 sec to 12° from latch", "404.2.8"),
          ("Vision lite (if any)", "≤ 43\" to bottom",  "404.2.11"),
          ("Manual revolving",     "Not permitted as sole", "404.1"),
          ("Auto operators",       "5 lb max actuation", "404.3"),
          ("Surface ≥ 10\" base",  "Smooth, kick plate", "404.2.10"),
         ])
quickref(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(5.4), "Ramps & Stairs",
         [("Ramp running slope",   "1:12 max",          "405.2"),
          ("Ramp cross slope",     "1:48 (2.1%) max",   "405.3"),
          ("Rise between landings", "30\" max",          "405.6"),
          ("Landings",             "60\" min length",   "405.7"),
          ("Landing at door",      "60\" × maneuv. clr.", "405.7.5"),
          ("Edge protection",      "4\" min curb or wall", "405.9"),
          ("Handrails ramp",       "Both sides if rise > 6\"", "405.8"),
          ("Handrail height",      "34\"–38\"",         "505.4"),
          ("Handrail extensions",  "12\" top, 12\"+tread bottom", "505.10"),
          ("Stair treads",         "11\" min depth",    "504.2"),
          ("Stair risers",         "4\"–7\"; uniform",  "504.2"),
          ("Stair nosings",        "Contrast stripe; no projection", "504.5"),
         ])
footer(s, 40)
add_notes(s, """
Door findings are the second-most common after restrooms. The "32"
clear width" rule sounds simple but is measured at the door at 90°,
between the face of the door and the stop on the latch side. With a
standard 36" door, you usually have 33-34" clear — but with a thick
weatherstrip or a swung door that doesn't open fully, you can lose
those inches.

Maneuvering clearance is where push vs pull doors trip people up.
Pull-side needs 18" past the latch; push-side needs 12" when there is
a closer. New reviewers often forget the closer condition.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 41 — Measurements ref: restrooms
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Quick reference — restrooms",
                            kicker="Slide 41 · Reference")
quickref(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(5.4), "Water closet + grab bars",
         [("Stall width",         "60\" min",         "604.8.1.1"),
          ("Stall depth wall-mt", "56\" min",         "604.8.1.1"),
          ("Stall depth floor-mt", "59\" min",         "604.8.1.1"),
          ("WC centerline",       "16\"–18\" from sidewall", "604.2"),
          ("Seat height",         "17\"–19\"",        "604.4"),
          ("Side grab bar length", "42\" min",         "604.5.1"),
          ("Side grab bar offset", "12\" max from rear", "604.5.1"),
          ("Rear grab bar length", "36\" min",         "604.5.2"),
          ("Grab bar height",     "33\"–36\"",        "609.4"),
          ("Flush controls side", "Open side (transfer side)", "604.6"),
          ("TP dispenser height", "15\"–48\"",        "604.7"),
          ("TP dispenser front",  "7\"–9\" in front of WC", "604.7"),
         ])
quickref(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(5.4), "Lavatory + accessories",
         [("Lav rim height",      "34\" max",          "606.3"),
          ("Knee clearance",      "27\"H × 30\"W × 11\"D", "306"),
          ("Toe clearance ext",   "9\" min height under", "306.2.3"),
          ("Faucet control",      "Closed-fist usable; 5 lbf", "309.4"),
          ("Hot / drain pipes",   "Insulated",         "606.5"),
          ("Mirror bottom",       "40\" max above floor", "603.3"),
          ("Mirror — full-len.",  "Bottom 35\" max",    "603.3"),
          ("Coat hook",           "48\" max forward",   "604.8.3 / 308.2.1"),
          ("Soap dispenser",      "Within 48\" reach",  "308"),
          ("Towel / dryer",       "Within 48\" reach",  "308"),
          ("Sanitary disposal",   "Within reach",       "Title II program"),
          ("Door swing into stall", "OK if 60×60 remains",  "604.8.1.2"),
         ])
footer(s, 41)
add_notes(s, """
Bring out the tape measure for this slide. Walk through how to measure
each one. Critical:

- WC centerline is measured to the centerline of the WATER CLOSET, not
  the centerline of the stall.
- Side grab bar "offset" 12" max from rear means the leading edge of
  the bar is no more than 12" from the rear wall — and the bar
  extends 42" toward the door from there.
- TP dispenser "7-9" in front of WC" is measured from the front edge
  of the seat to the centerline of the dispenser roll.
- Coat hook reach is total HEIGHT from floor, not depth.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 42 — Measurements ref: fountains, counters, signage
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Quick reference — fountains, counters & signage",
                            kicker="Slide 42 · Reference")
quickref(s, Inches(0.55), Inches(1.55), Inches(6.0), Inches(5.4), "Fountains & built-in counters",
         [("Drinking fountain spout", "36\" max above floor", "602.4"),
          ("Spout location",          "≤ 15\" from front",   "602.5"),
          ("Spout angle",             "≤ 30° from vertical", "602.6"),
          ("Knee clearance — DF",     "27\"H × 30\"W × 17\"D min", "306"),
          ("Hi-Lo paired DF",         "One ≤36\" + one for standing", "602.7"),
          ("Service counter (acc.)",  "34\" max height",    "904.4.1"),
          ("Service counter length",  "36\" min long",      "904.4.1"),
          ("Sales / service — POS",   "≤ 38\" if integral", "904.4"),
          ("Library check-out",       "34\" max writing srf", "904.4"),
          ("Cafeteria self-serve",    "≤ 34\" tray slide",  "226.2 / 904"),
          ("Counter slope",           "Level",              "904.4"),
          ("Forward reach over",      "Apply 308.2",        "308"),
         ])
quickref(s, Inches(6.85), Inches(1.55), Inches(6.0), Inches(5.4), "Signage & wayfinding",
         [("ISA size — room entry",   "6\" min ISA",       "703.7.2.1"),
          ("Tactile mounting",        "48\"–60\" baseline", "703.4.1"),
          ("Tactile location",        "Latch side, 18\" max from frame", "703.4.2"),
          ("Tactile character size",  "5/8\"–2\" uppercase", "703.2.5"),
          ("Tactile raised",          "1/32\" min",         "703.2.1"),
          ("Braille",                 "Grade 2 adjacent",   "703.3"),
          ("Visual character contrast", "70% min",         "703.5.1"),
          ("Visual char. height (rm)", "5/8\" min",         "703.5.5"),
          ("ISA color",               "White on blue (or contrast)", "703.7.2.1"),
          ("Directional signs",       "Required when route ≠ main", "216.3 / 216.6"),
          ("Restroom door designation", "ISA + 'WOMEN' / 'MEN' tactile", "216.8"),
          ("Stairway tactile",        "Exit sign tactile", "216.4.3 / 703.4"),
         ])
footer(s, 42)
add_notes(s, """
Three reference slides done. These are intentionally information-dense
— hand them out as a print packet for site visits. Reviewers should
plan to refer to them many times before they internalize the numbers.

Counters: the cafeteria tray-slide is a frequent issue. Required max
34" — many CTE-affiliated school kitchens have 36"–38" tray slides
from older renovations.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 43 — CTE sectors at a glance (all 15)
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "California's 15 CTE sectors — at a glance",
                            kicker="Slide 43 · CTE classrooms")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.45),
         "Every CRR 20 / 21 review touches at least one CTE sector. Know which sector you're in — it shapes what to measure.",
         size=14, italic=True, color=MUTED)

sectors = [
    ("Agriculture & Natural Resources",       "Greenhouses, ag mech shops, animal labs, vet science"),
    ("Arts, Media & Entertainment",           "Studios, photo darkrooms, performance spaces, edit bays"),
    ("Building & Construction Trades",        "Carpentry, masonry, electrical, HVAC, plumbing bays"),
    ("Business & Finance",                    "Computer labs, mock-bank classrooms"),
    ("Education, Child Development & Family", "Lab schools, preschool classrooms, kitchens"),
    ("Energy, Environment & Utilities",       "Solar / wind / EV labs, HVAC labs"),
    ("Engineering & Architecture",            "CAD labs, drafting tables, fabrication labs"),
    ("Fashion & Interior Design",             "Sewing labs, fitting rooms, interior design studios"),
    ("Health Science & Medical Technology",   "Nursing skills labs, dental, biotech, sports med"),
    ("Hospitality, Tourism & Recreation",     "Culinary kitchens, hotel mock-rooms, baking labs"),
    ("Information & Communication Tech",      "Networking labs, server rooms, cyber labs"),
    ("Manufacturing & Product Development",   "Welding bays, CNC labs, machine shops, robotics"),
    ("Marketing, Sales & Service",            "Retail mock-stores, service counters"),
    ("Public Services",                       "Fire science, criminal justice, EMR labs, simulators"),
    ("Transportation",                        "Auto shop, diesel, aviation, marine bays"),
]
# 3 columns x 5 rows
for i, (name, examples) in enumerate(sectors):
    col = i % 3
    row = i // 3
    x = Inches(0.55 + 4.25*col)
    y = Inches(2.1 + 0.95*row)
    add_rect(s, x, y, Inches(4.05), Inches(0.85), CREAM, line=BORDER)
    add_rect(s, x, y, Inches(0.12), Inches(0.85), GOLD)
    add_text(s, x+Inches(0.2), y+Inches(0.07), Inches(3.8), Inches(0.32),
             name, size=10, bold=True, color=NAVY)
    add_text(s, x+Inches(0.2), y+Inches(0.4), Inches(3.8), Inches(0.45),
             examples, size=9, color=MUTED, italic=True)
footer(s, 43)
add_notes(s, """
Every CA high school is required to offer at least one CTE pathway, and
most offer four or more. The CRR queue is dominated by Manufacturing,
Building & Construction Trades, Health Science, and Hospitality —
because those sectors require purpose-built shops, labs, and kitchens
that have the most accessibility risk.

Tell the class: when you arrive on site, find out which pathways the
school offers BEFORE you walk. The Facilities Review Packet should list
them. If it doesn't, ask the CTE director on arrival.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 44 — Building / Construction Trades shops
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "CTE deep-dive — Building & Construction Trades",
                            kicker="Slide 44 · CTE classrooms")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Carpentry, masonry, electrical, HVAC, plumbing — high-equipment-density labs with mobility, reach, and clearance challenges.",
         size=13, italic=True, color=MUTED)

# What to look for
add_text(s, Inches(0.55), Inches(2.1), Inches(6.5), Inches(0.4),
         "What to look for:",
         size=14, bold=True, color=NAVY)
add_bullets(s, Inches(0.55), Inches(2.55), Inches(6.5), Inches(4.5), [
    ("Workbench accessibility.", "At least one workbench with 27\"H × 30\"W × 19\"D knee/toe clearance for forward approach. Adjustable-height benches strongly preferred."),
    ("Power tool reach.",        "Outlet strips and tool controls within 15\"–48\" reach range. Plug strips mounted on wall above 48\" = violation."),
    ("Aisle width between stations.", "36\" minimum accessible route; 48\" preferred for wheelchair turnaround between active stations."),
    ("Tool storage / cabinets.",  "Shelves and drawers accessible — top shelf ≤ 48\" reach. Heavy tools at lower shelves."),
    ("Emergency stop / e-stop.",  "Within 15\"–48\" reach. Mushroom heads visible — color contrast required for safety (CA OSHA also)."),
    ("Eyewash / safety shower.",  "Hardware operable with closed fist; spray heads within reach range; clear floor space at base."),
], size=12, spacing=4)

# Common findings box
add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(4.95), CREAM, line=RED)
add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45), RED)
add_text(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45),
         "COMMON FINDINGS",
         size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(7.5), Inches(2.7), Inches(5.2), Inches(4.2), [
    "No workbench with knee/toe clearance — 100% of stations are fixed at 36\" with full skirts.",
    "Tool cabinets with controls > 48\" forward reach.",
    "Compressed-air quick connects at 64\" wall mount (out of reach).",
    "Welded-pipe storage racks reducing aisle to 28\".",
    "Safety shower behind a sliding cabinet door — no clear floor space.",
    "Sawdust collection ducts blocking accessible route at chest height.",
], size=11, spacing=4, color=TEXT)
footer(s, 44)
add_notes(s, """
Construction shops are operationally messy by nature. The accessibility
question is not "can you keep this immaculate?" — it is "is there a
configuration in which a student using a wheelchair can fully
participate in the curriculum?"

Look for at least one fully-accessible workstation. Programs often
designate one "ADA bench" and call it good — verify it actually meets
reach and knee clearance specs, and verify it is reachable via an
accessible route from the classroom door.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 45 — Manufacturing / welding bays
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "CTE deep-dive — Manufacturing & welding",
                            kicker="Slide 45 · CTE classrooms")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Welding bays, CNC, machining, robotics, automotive — the highest-risk CTE category for accessibility issues.",
         size=13, italic=True, color=MUTED)

add_text(s, Inches(0.55), Inches(2.1), Inches(6.5), Inches(0.4),
         "Critical issues:",
         size=14, bold=True, color=NAVY)
add_bullets(s, Inches(0.55), Inches(2.55), Inches(6.5), Inches(4.5), [
    ("Welding booth dimensions.", "Booth opening 32\" minimum clear. Interior 60\"×60\" turning space if booth is enclosed."),
    ("Welding table height.",     "At least one table at 32\"–34\" with knee/toe clearance for seated operation."),
    ("Booth ventilation reach.",  "Hood positioning controls and gas-valve controls within 48\" reach."),
    ("CNC interface.",            "Touch-screen or operator panel within 15\"–48\" reach; clear floor space in front."),
    ("Auto-shop lifts.",          "Lift controls within reach; accessible route to lift bay; clear floor space adjacent."),
    ("Toolbox / tool crib.",      "Service counter at 34\" max if students approach; or alternate accessible counter."),
], size=12, spacing=4)

add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(4.95), CREAM, line=RED)
add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45), RED)
add_text(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45),
         "COMMON FINDINGS",
         size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(7.5), Inches(2.7), Inches(5.2), Inches(4.2), [
    "Every welding booth has a 30\" frame opening (must be 32\" min).",
    "Welding-table-height fixed at 38\"; no accessible alternate.",
    "Gas valves mounted at 70\" wall height (out of reach).",
    "CNC pendant control on a fixed swing-arm at 52\" (1991 ADA range, but 48\" max under 2010 ADA for forward unobstructed reach).",
    "Auto lift bay reached only via 1:8 internal driveway slope.",
    "Tool-crib service window at 42\" with no lower counter — violates 904.4.1 (34\" max).",
], size=11, spacing=4, color=TEXT)
footer(s, 45)
add_notes(s, """
Welding bays are tightly engineered for safety and ventilation, which
sometimes conflicts with access. The accessible workstation does NOT
have to be welding-equipped — it can be a clean-room or layout bench
that participates in the welding curriculum even if it doesn't have a
torch. The question is curriculum participation, not 100% station
parity.

If you find a welding shop with no accessible booth and no accessible
alternative workstation, that is a hard finding — students with
mobility disabilities cannot participate. Cite the booth opening
(404.2.3) plus the workstation knee/toe (306) and write up corrective
actions for both.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 46 — Culinary kitchens
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "CTE deep-dive — Culinary & hospitality",
                            kicker="Slide 46 · CTE classrooms")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Commercial kitchens, baking labs, hotel mock-rooms — purpose-built with heavy fixed equipment.",
         size=13, italic=True, color=MUTED)

add_text(s, Inches(0.55), Inches(2.1), Inches(6.5), Inches(0.4),
         "Critical issues:",
         size=14, bold=True, color=NAVY)
add_bullets(s, Inches(0.55), Inches(2.55), Inches(6.5), Inches(4.5), [
    ("Accessible prep station.", "Min one prep counter at 34\" max height with 27\"H × 30\"W × 19\"D knee/toe clearance."),
    ("Cooktop / range.",         "Min one cooktop at 34\" max with knee clearance, OR an accessible alternate cooktop on a movable cart."),
    ("Sink height.",             "Lab sink: 34\" max with knee/toe under. Hot/drain pipes insulated."),
    ("Cabinet & pantry reach.",  "Drawers and lower shelves accessible — top shelf ≤ 48\". Walk-in pantry door 32\" min clear width."),
    ("Refrigerator / freezer.",  "Door handle within reach (48\" max); pull force ≤ 5 lbf; clear floor space at door."),
    ("Dishwashing station.",     "Spray-arm controls within 48\" reach; rinse-station entry 32\" min."),
    ("Service / dining counter.", "Per 904.4.1 — 34\" max counter, 36\" min length."),
], size=12, spacing=3)

add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(4.95), CREAM, line=RED)
add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45), RED)
add_text(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45),
         "COMMON FINDINGS",
         size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(7.5), Inches(2.7), Inches(5.2), Inches(4.2), [
    "Every prep station at 36\" with full-skirt cabinets — zero knee clearance.",
    "Cooktops at 38\" with no knee clearance and no accessible alternate.",
    "Dish-pit spray arm controls behind a fixed stainless splash — 52\" reach over 12\" deep counter (max allowed = 48\" under 308.2).",
    "Walk-in pantry door 30\" clear width (must be 32\" min).",
    "Tray-slide service line at 36\" (must be 34\" max).",
    "Walk-in freezer threshold > ½\" — trips wheelchair casters.",
    "Hand-wash sink at 40\" rim height (must be 34\" max).",
], size=11, spacing=3, color=TEXT)
footer(s, 46)
add_notes(s, """
Culinary classrooms are heavily fixed-built and expensive to alter.
That doesn't change the standard — a student in a wheelchair must be
able to fully participate in the curriculum. Look for at least one
fully-compliant prep station with knee/toe clearance, an accessible
cooktop or hot plate alternate, and an accessible hand-wash sink.

Common LEA pushback: "Health Code requires the equipment to be at this
height." The response: cite the accessibility standard, and the LEA
must reconcile with Cal/OSHA and California Retail Food Code. Both
codes have accommodation mechanisms.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 47 — Agriculture & natural resources
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "CTE deep-dive — Agriculture & natural resources",
                            kicker="Slide 47 · CTE classrooms")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Ag mechanics, greenhouses, animal labs, vet science — often the oldest buildings on campus.",
         size=13, italic=True, color=MUTED)

add_text(s, Inches(0.55), Inches(2.1), Inches(6.5), Inches(0.4),
         "Critical issues:",
         size=14, bold=True, color=NAVY)
add_bullets(s, Inches(0.55), Inches(2.55), Inches(6.5), Inches(4.5), [
    ("Greenhouse access.",         "Door clear width 32\" min; bench/table height with knee/toe per 306; aisle 36\" min between benches."),
    ("Greenhouse surface.",        "Gravel floors fail 302.1 firm/stable. Concrete or composite paver paths required for accessible route."),
    ("Ag mech shop.",              "Same workbench, reach, and aisle requirements as Building & Construction Trades shops."),
    ("Animal-science labs.",       "Hose-bibs and wash stations within reach; clear floor space at restraint chutes."),
    ("Outdoor classrooms / pens.", "Accessible route from building to pen; gates 32\" min clear; ground surfaces firm and stable."),
    ("Farm-to-school kitchen.",    "Same culinary specs (see Slide 46)."),
], size=12, spacing=4)

add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(4.95), CREAM, line=RED)
add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45), RED)
add_text(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45),
         "COMMON FINDINGS",
         size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(7.5), Inches(2.7), Inches(5.2), Inches(4.2), [
    "Greenhouse door 30\" clear (must be 32\" min).",
    "Decomposed-granite pathway to pens — not firm or stable, fails 302.1.",
    "Pen gates with drop-bolt closures requiring tight grasping (309.4 violation).",
    "Greenhouse benches all at 40\" with no knee clearance.",
    "Ag mech shop in a pre-1977 building — Program Access; NO findings (this is the trap!).",
    "Welding pad next to ag building — added 2018, becomes 2010 ADA.",
], size=11, spacing=4, color=TEXT)
footer(s, 47)
add_notes(s, """
Ag buildings on California campuses are commonly pre-1977 — the
program-access trap is most common here. New reviewers will see a
gravel path and clearly inaccessible greenhouse and want to write up
2010 ADA violations. Catch yourself: check the date. If pre-1977 and
unaltered, NO findings. Document the program-access concern in the
narrative.

A reasonable corrective recommendation for a pre-1977 ag building: have
the LEA designate an alternate accessible location for the ag
curriculum, or document a program-access analysis that explains how
students with mobility disabilities will participate in the program.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 48 — Health Science skills labs
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "CTE deep-dive — Health Science & Medical Tech",
                            kicker="Slide 48 · CTE classrooms")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Nursing skills labs, dental labs, biotech, sports medicine — usually newer (post-2010), so 2010 ADA applies.",
         size=13, italic=True, color=MUTED)

add_text(s, Inches(0.55), Inches(2.1), Inches(6.5), Inches(0.4),
         "Critical issues:",
         size=14, bold=True, color=NAVY)
add_bullets(s, Inches(0.55), Inches(2.55), Inches(6.5), Inches(4.5), [
    ("Patient simulation bays.",   "Aisle 36\" min; clear floor space at hospital bed for student approach. Curtain hardware operable with closed fist."),
    ("Medication cart / Pyxis.",   "Drawer pulls within reach; operable parts at 15\"–48\"; door swing maneuvering clearance."),
    ("Dental chair stations.",     "Min one station with student approach from accessible side. Lap-tray accessible counter alternate."),
    ("Lab benches (biotech).",     "Same workbench rules — at least one bench with knee/toe; reagent shelves within reach."),
    ("Microscope / fume hood.",    "Knee clearance under hood sash; control sash and valves within reach. Pull force ≤ 5 lbf."),
    ("Hand-wash sink locations.",  "34\" rim, knee clearance, insulated pipes — same as restroom lavs."),
], size=12, spacing=4)

add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(4.95), CREAM, line=RED)
add_rect(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45), RED)
add_text(s, Inches(7.3), Inches(2.1), Inches(5.55), Inches(0.45),
         "COMMON FINDINGS",
         size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(7.5), Inches(2.7), Inches(5.2), Inches(4.2), [
    "Sim-bay curtain track requires two-hand operation (309.4 violation).",
    "Med-cart drawers with knob-only pulls (need lever / paddle).",
    "Biotech bench at 36\" with full skirt — no accessible alternate.",
    "Fume hood sash control at 56\" forward reach (must be ≤ 48\").",
    "Eye-wash mounted in a corner with no clear floor space.",
    "Specimen-prep counter at 38\" — must be 34\" max under 904.4.",
], size=11, spacing=4, color=TEXT)
footer(s, 48)
add_notes(s, """
Health Science labs are usually built recently and to a high standard.
Most findings here are dimensional misses, not gross access failures.
That said, the curriculum is exam-driven — students must demonstrate
skills, often physically — so workstation parity matters.

The dental chair example is illustrative: students learning to be
dental assistants work from a specific approach side. If only one chair
allows an alternate approach, accessibility is constrained to that one
station. Treat the rotation schedule as part of the program-access
analysis: does every student get equitable time at the accessible
station?
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 49 — Other CTE sectors survey
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Survey of remaining CTE sectors",
                            kicker="Slide 49 · CTE classrooms")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Quick orientation to the remaining sectors — at least one issue to look for in each.",
         size=13, italic=True, color=MUTED)

quick = [
    ("Arts, Media & Entertainment", "Darkrooms (32\" door); edit-bay desk knee/toe; performance stage routing to/from."),
    ("Business & Finance",          "Computer labs: workstation height + accessible chair pull-up; service window 34\" max."),
    ("Education, Child Dev.",       "Lab schools — child-height fixtures, BUT accessible-route + restroom under 2010 ADA."),
    ("Energy, Environment & Util.", "Solar/wind lab outdoor pads — surface stable, route 36\" min."),
    ("Engineering & Architecture",  "CAD stations: keyboard at 28\"–34\" preferred; drafting tables w/ knee clearance."),
    ("Fashion & Interior Design",   "Sewing stations: machine controls within reach; fitting rooms 60\"×60\" turning."),
    ("Information & Comm. Tech",    "Server racks have no accessibility scoping; student stations DO — apply standard workstation rules."),
    ("Marketing, Sales & Service",  "Mock-retail service counters — apply 904.4.1 (34\" max, 36\" min length)."),
    ("Public Services",             "Simulators and fire/police training props — at least one accessible station; restroom in vehicle bay."),
    ("Transportation",              "Auto / diesel / aviation bays — same as Manufacturing. Lifts, hoists, work pits = clear floor space."),
]
for i, (sector, body) in enumerate(quick):
    col = i % 2
    row = i // 2
    x = Inches(0.55 + 6.2*col)
    y = Inches(2.1 + 1.0*row)
    add_rect(s, x, y, Inches(6.0), Inches(0.9), CREAM, line=BORDER)
    add_rect(s, x, y, Inches(0.12), Inches(0.9), GOLD)
    add_text(s, x+Inches(0.2), y+Inches(0.07), Inches(5.7), Inches(0.3),
             sector, size=11, bold=True, color=NAVY)
    add_text(s, x+Inches(0.2), y+Inches(0.4), Inches(5.7), Inches(0.5),
             body, size=10, color=TEXT)
footer(s, 49)
add_notes(s, """
This slide is intentionally light. Reviewers are unlikely to see all
fifteen sectors in their first year — but they will see most. The point
is to give them a starting frame so they aren't surprised the first
time they walk into a fashion-design fitting room or a fire-science
training pad.

If you have extra time in this segment, ask the class which sectors
they've already visited or had questions about — and dive deeper there.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 50 — CTE accommodation strategies
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "How LEAs actually accommodate — patterns that work",
                            kicker="Slide 50 · CTE classrooms")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Reviewers should recognize compliant configurations — not just non-compliant ones.",
         size=14, italic=True, color=MUTED)

patterns = [
    ("Adjustable-height workbenches",
     "Crank or motorized tables that rise from 28\" to 42\". One per shop satisfies the accessible-workstation requirement under 2010 ADA 902.4."),
    ("Mobile prep carts (culinary)",
     "Wheeled cart with 34\" max counter and knee clearance can supplement fixed prep stations — must be present and used, not stored."),
    ("Designated accessible welding booth",
     "32\" min frame opening + 60\"×60\" interior turning space + accessible-height table. Programs that designate ONE such booth and rotate students through it are compliant."),
    ("Roll-under microscope / fume-hood station",
     "Sash control at 48\" max, open knee space, insulated piping. Common in newer biotech labs."),
    ("Service counter combo",
     "Standard 38\" counter PLUS a 34\" lower section ≥ 36\" long. Satisfies 904.4.1."),
    ("Greenhouse paver path",
     "Concrete or composite paver path replacing gravel — meets 302.1 firm/stable for outdoor accessible route."),
]
for i, (head, body) in enumerate(patterns):
    col = i % 2
    row = i // 2
    x = Inches(0.55 + 6.2*col)
    y = Inches(2.1 + 1.7*row)
    add_rect(s, x, y, Inches(6.0), Inches(1.6), CREAM, line=GREEN)
    add_rect(s, x, y, Inches(0.18), Inches(1.6), GREEN)
    add_text(s, x+Inches(0.35), y+Inches(0.1), Inches(5.5), Inches(0.4), head,
             size=13, bold=True, color=GREEN)
    add_text(s, x+Inches(0.35), y+Inches(0.55), Inches(5.5), Inches(1.0), body,
             size=11, color=TEXT)
footer(s, 50)
add_notes(s, """
This slide flips the lens. If you only train reviewers on what FAILS,
they walk into compliant programs ready to find fault. Train them to
recognize what SUCCEEDS so they can call out and celebrate good
practice in the LOF narrative.

Mobile carts deserve a callout: if the cart is in the storage closet
and only rolled out for the review, that is not compliance — it must be
in active use as part of the program's regular configuration.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 51 — The 18 area review categories
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "The 18 area review categories",
                            kicker="Slide 51 · Reviewer rules")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Every CRR 20 / 21 self-assessment and LOF table uses this exact sequence — in this exact order.",
         size=14, italic=True, color=MUTED)

areas = [
    "Accessible Parking",
    "Accessible Routes / Walkways from accessible parking",
    "Stairways and Steps",
    "Ramps",
    "Curb Ramps",
    "Entrances, Doors, and Gates",
    "Rooms, Offices, and Administration",
    "Elevators / Lifts",
    "Accessible Drinking Fountains",
    "Cafeteria",
    "Library",
    "CTE Classrooms",
    "Labs / Shops",
    "Gymnasium / Auditorium / Weight Room / Other",
    "Stadium / Field",
    "Dressing, Fitting, and Locker Rooms",
    "Restrooms",
    "Public Telephones",
]
# Two-column numbered list
for i, area in enumerate(areas):
    col = i // 9
    row = i % 9
    x = Inches(0.55 + 6.4*col)
    y = Inches(2.15 + 0.55*row)
    add_rect(s, x, y, Inches(0.6), Inches(0.45), NAVY)
    add_text(s, x, y, Inches(0.6), Inches(0.45), f"{i+1:02d}",
             size=12, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    add_rect(s, x+Inches(0.6), y, Inches(5.5), Inches(0.45), CREAM, line=BORDER)
    add_text(s, x+Inches(0.75), y, Inches(5.3), Inches(0.45), area,
             size=12, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)

add_rect(s, Inches(0.55), Inches(7.0), Inches(12.3), Inches(0.25), GOLD)
add_text(s, Inches(0.55), Inches(6.97), Inches(12.3), Inches(0.3),
         "Don't skip rows. If an area is not present, the row reads N/A — but the row still appears.",
         size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 51)
add_notes(s, """
The 18-row format is sacred. Every CRR 20 / 21 self-assessment, every
LOF table, every boilerplate — they all assume these 18 areas in this
exact order. Don't let an LEA submit a self-assessment that combines
rows or skips categories.

For areas that don't exist on a campus (e.g., a K-8 school with no
cafeteria), the row still appears in the LOF — marked N/A in the
applicable standard column, with "None." in both finding columns.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 52 — Element-by-element alteration rule (DEEP DIVE)
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Reviewer rule #1 — element-by-element alterations",
                            kicker="Slide 52 · Reviewer rules")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.45),
         "The single most common reviewer mistake — applying a newer standard to an older facility because part of it was altered. Don't.",
         size=14, italic=True, color=MUTED)

# Three example scenarios — same building, different element-level analyses
scen_h = ["WHAT THE LEA REPORTS", "WHAT THE REVIEWER ASSIGNS"]
for i, h in enumerate(scen_h):
    add_rect(s, Inches(0.55+6.2*i), Inches(2.05), Inches(6.0), Inches(0.4), NAVY)
    add_text(s, Inches(0.55+6.2*i), Inches(2.05), Inches(6.0), Inches(0.4), h,
             size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")

scenarios52 = [
    ("1968 restroom — grab bars replaced 2018.\nNo other alterations.",
     "Grab bars → 2010 ADA (609.4).\nWC seat, lav, mirror, stall, signage → Program Access. NO findings possible on those elements."),
    ("1985 ramp — entire ramp + handrails rebuilt 2020.\nWhole-element rebuild.",
     "Entire ramp + handrails → 2010 ADA (405, 505).\nThe ramp is the element; rebuilding it is a whole-element alteration."),
    ("1962 cafeteria — new tray-slide installed 2016.\nKitchen otherwise unchanged.",
     "Tray-slide → 2010 ADA (904.4.1, max 34\").\nFixed tables, counter heights elsewhere → Program Access. NO findings."),
    ("1970 wood shop — entire room gutted & rebuilt 2019.",
     "Whole-room rule (§ 202.3): entire rebuilt room → 2010 ADA. Every element (workbenches, signage, doors, outlets, eye-wash) evaluated to 2010 ADA."),
]
for i, (lea, rev) in enumerate(scenarios52):
    y = Inches(2.5 + 1.1*i)
    add_rect(s, Inches(0.55), y, Inches(6.0), Inches(1.0), CREAM, line=BORDER)
    add_text(s, Inches(0.7), y+Inches(0.1), Inches(5.75), Inches(0.85), lea,
             size=11, color=TEXT, italic=True)
    add_rect(s, Inches(6.75), y, Inches(6.0), Inches(1.0), CREAM, line=GOLD)
    add_text(s, Inches(6.9), y+Inches(0.1), Inches(5.75), Inches(0.85), rev,
             size=11, color=NAVY)

footer(s, 52)
add_notes(s, """
Walk row by row. Anchor every example back to the Access Board language:
"Only those elements or spaces altered are required to comply… If a
room or space is completely altered, the entire room or space is fully
subject to the standards."

Row 1 (grab bars only): this is the most common scenario. The bars were
the only altered element. The toilet, lav, mirror were not altered and
do not become 2010 ADA — they remain Program Access (1968).

Row 2 (ramp): the ramp IS the element. If you rebuilt the whole ramp,
you altered the whole element. 2010 ADA applies to the new ramp.

Row 3 (tray-slide): only the tray-slide was installed/altered. Other
fixed-counter elements in the kitchen remain at their original
construction standard.

Row 4 (whole-room rebuild): the "completely altered room" rule from
§ 202.3 — the entire rebuilt room is fully subject to the standards.
Every fixture, every outlet, every signage element.

Reviewer takeaway: walk the LEA's work orders element-by-element and
mark each individually. Do not paint a whole building with one standard.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 53 — Corrective actions = always 2010 ADA (DEEP DIVE)
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Reviewer rule #2 — corrective actions = always 2010 ADA",
                            kicker="Slide 53 · Reviewer rules")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.45),
         "Cite the violation at the standard that applies to the element. Write the corrective action to bring the element into compliance with 2010 ADA — always.",
         size=14, italic=True, color=MUTED)

# CRR boilerplate quote
add_rect(s, Inches(0.55), Inches(2.1), Inches(12.3), Inches(1.0), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(2.1), Inches(0.18), Inches(1.0), GOLD)
add_text(s, Inches(0.85), Inches(2.2), Inches(11.8), Inches(0.3),
         "CRR 20 boilerplate (Summary of Analysis):",
         size=11, bold=True, color=NAVY, font="Consolas")
add_text(s, Inches(0.85), Inches(2.5), Inches(11.7), Inches(0.6),
         "\"The Office of Civil Rights requires all corrective actions to be made in accordance with 2010 ADA standards; therefore, certain areas will not require corrective action as the noted deficiency is within the 2010 ADA standards.\"",
         size=12, italic=True, color=TEXT)

# Side-by-side: VIOLATION cite vs CORRECTIVE cite
add_rect(s, Inches(0.55), Inches(3.3), Inches(6.0), Inches(3.5), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(3.3), Inches(6.0), Inches(0.45), NAVY)
add_text(s, Inches(0.55), Inches(3.3), Inches(6.0), Inches(0.45),
         "VIOLATION COLUMN — cite at the APPLICABLE standard",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(0.75), Inches(3.85), Inches(5.7), Inches(2.9), [
    "1985 handrail at 30\" → Cite: ANSI A117.1 § 4.8.5 (34–38\" required).",
    "2003 ramp at 1:9 → Cite: 1991 ADA § 4.8.2 (1:12 max).",
    "2016 toilet seat at 16\" → Cite: 2010 ADA § 604.4 (17–19\")."
], size=12, spacing=6)

add_rect(s, Inches(6.85), Inches(3.3), Inches(6.0), Inches(3.5), CREAM, line=GOLD)
add_rect(s, Inches(6.85), Inches(3.3), Inches(6.0), Inches(0.45), GOLD)
add_text(s, Inches(6.85), Inches(3.3), Inches(6.0), Inches(0.45),
         "CORRECTIVE COLUMN — always cite 2010 ADA",
         size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(7.05), Inches(3.85), Inches(5.7), Inches(2.9), [
    "Reset handrail to 34–38\" per 2010 ADA § 505.4.",
    "Re-grade ramp to max 1:12 running, 1:48 cross per 2010 ADA § 405.2 / 405.3.",
    "Reset water closet seat to 17–19\" per 2010 ADA § 604.4.",
], size=12, spacing=6)

# Bottom call-out
add_rect(s, Inches(0.55), Inches(6.95), Inches(12.3), Inches(0.4), GOLD)
add_text(s, Inches(0.55), Inches(6.93), Inches(12.3), Inches(0.4),
         "Why: 2010 ADA is the federal target state. Bringing an old element \"up to its old standard\" doesn't fix the access barrier — only 2010 ADA does.",
         size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 53)
add_notes(s, """
This is reviewer rule #2 — and it's a rule I deliberately did NOT cover
on the citation-placement slide earlier because it gets confused with
the "cite the standard that controls the element" rule. Both rules are
true at the same time, in different columns:

VIOLATION column: cite the standard that controls the ELEMENT
(ANSI for 1977-1991 elements, 1991 ADA for 1992-2010 elements, 2010
ADA for 2012+ elements).

CORRECTIVE column: always cite 2010 ADA, because the corrective work
itself is new construction / alteration happening today — and today's
work is 2010 ADA.

Example to drive this home aloud:

Violation: "Handrail on the wood shop ramp measured 30\" — does not
meet ANSI A117.1 § 4.8.5 (34–38\" required at time of 1985
construction). Cite: ANSI A117.1 § 4.8.5."

Corrective action: "Replace handrail at a height of 34–38\" above the
ramp surface in conformance with 2010 ADA § 505.4, with 12\" horizontal
extensions at the top and bottom per § 505.10. Provide photographic
documentation of completed installation. Complete within 45 days.
Cite: 2010 ADA § 505.4 and § 505.10."

Note the corrective cites 2010 ADA — NOT ANSI. The fix is being built
new, so 2010 ADA controls.

This is the rule the CRR 20 boilerplate codifies. Read the boilerplate
quote at top of the slide aloud. That is the rule.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 54 — Knowledge check on rules #1 and #2
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); set_bg(s, NAVY)
add_rect(s, 0, 0, Inches(0.6), SLIDE_H, GOLD)
add_text(s, Inches(0.85), Inches(0.6), Inches(12), Inches(0.5),
         "KNOWLEDGE CHECK · 04½",
         size=12, bold=True, color=GOLD, font="Consolas")
add_text(s, Inches(0.85), Inches(1.1), Inches(12), Inches(0.9),
         "Both rules at once",
         size=32, bold=True, color=WHITE, font="Georgia")

add_rect(s, Inches(0.85), Inches(2.15), Inches(11.95), Inches(1.7), NAVY_MID)
add_text(s, Inches(1.1), Inches(2.3), Inches(11.5), Inches(1.5),
         "1972 high-school cafeteria. Original 1972 fixed counters, 1972 tray-slide. In 2018 the LEA installed a NEW tray-slide (only) — work orders confirm.\n\nField measurement: new (2018) tray-slide at 36\" above floor.\nField measurement: original 1972 counter at 37\" above floor.",
         size=14, color=WHITE)

# Two-part question
add_text(s, Inches(0.85), Inches(4.0), Inches(12), Inches(0.4),
         "ANSWER BOTH PARTS:",
         size=12, bold=True, color=GOLD, font="Consolas")

add_rect(s, Inches(0.85), Inches(4.5), Inches(5.8), Inches(2.45), NAVY_MID, line=GOLD)
add_text(s, Inches(1.05), Inches(4.6), Inches(5.5), Inches(0.4),
         "PART A — 2018 tray-slide @ 36\":",
         size=13, bold=True, color=GOLD)
add_text(s, Inches(1.05), Inches(5.0), Inches(5.5), Inches(1.85),
         "Standard at violation: ____________\n\nCite at violation: ____________\n\nCite at corrective: ____________",
         size=12, color=GOLD_LT, font="Consolas")

add_rect(s, Inches(7.05), Inches(4.5), Inches(5.8), Inches(2.45), NAVY_MID, line=GOLD)
add_text(s, Inches(7.25), Inches(4.6), Inches(5.5), Inches(0.4),
         "PART B — 1972 counter @ 37\":",
         size=13, bold=True, color=GOLD)
add_text(s, Inches(7.25), Inches(5.0), Inches(5.5), Inches(1.85),
         "Standard at violation: ____________\n\nCite at violation: ____________\n\nCite at corrective: ____________",
         size=12, color=GOLD_LT, font="Consolas")
footer(s, 54)
add_notes(s, """
Answers:

PART A — 2018 tray-slide:
The tray-slide is a newly-installed ELEMENT in 2018. It is evaluated
under the standard in effect on the alteration date.
- Standard at violation: 2010 ADA
- Cite at violation: 2010 ADA § 904.4.1 (max 34" tray slide height).
- Cite at corrective: 2010 ADA § 904.4.1.
(In this case both cites are 2010 ADA because the element was already
constructed under 2010 ADA — but the corrective rule still applies.)

PART B — 1972 counter:
The counter is an UNALTERED element on a 1972 building.
- Standard at violation: PROGRAM ACCESS. No findings can be issued —
  the original 1972 counter is governed by program access, not by a
  dimensional standard.
- Cite at violation: NONE — leave as "None."
- Cite at corrective: NONE — leave as "None."
- (Address the access concern in the LOF narrative instead, e.g.,
  "LEA should designate an accessible alternate counter location.")

The trap: students will want to cite the 1972 counter as a 2010 ADA
violation because it's at 37" (over the 34" max). That is wrong on two
counts: (1) the counter was not altered, so 2010 ADA does not apply;
(2) Program Access has no dimensional standard, so no findings issue.

This double-check is exactly what new reviewers will be tested on by
their team leads. Master it now.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 55 — Anatomy of a Violation entry
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Anatomy of an Accessibility Violation entry",
                            kicker="Slide 55 · Writing findings")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "A defensible violation entry has four parts. Every one must be present.",
         size=14, italic=True, color=MUTED)

# Full-width quote-style finding box
add_rect(s, Inches(0.55), Inches(2.05), Inches(12.3), Inches(1.6), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(2.05), Inches(0.18), Inches(1.6), GOLD)
add_text(s, Inches(0.85), Inches(2.2), Inches(12), Inches(0.4),
         "EXAMPLE — fully-formed violation entry:",
         size=11, bold=True, color=NAVY, font="Consolas")
add_text(s, Inches(0.85), Inches(2.6), Inches(11.7), Inches(1.0),
         "Van-accessible parking aisle adjacent to space #1 measured at 84 inches wide. The aisle does not meet the 96-inch minimum width specified in 2010 ADA Standards § 502.3.3 for van-accessible aisles. Cite: 2010 ADA 502.3.3.",
         size=13, italic=True, color=TEXT)

# Breakdown of parts
parts = [
    ("WHAT",     "The element and its measured condition.", "\"Van-accessible parking aisle adjacent to space #1 measured at 84 inches wide.\""),
    ("WHY",      "The standard's requirement — quoted or paraphrased.", "\"The aisle does not meet the 96-inch minimum width specified in 2010 ADA § 502.3.3.\""),
    ("WHERE",    "Specific location identifier.", "\"adjacent to space #1\""),
    ("CITE",     "The controlling section, with standard year.", "\"Cite: 2010 ADA 502.3.3.\""),
]
for i, (head, body, ex) in enumerate(parts):
    y = Inches(3.85 + 0.75*i)
    add_rect(s, Inches(0.55), y, Inches(1.0), Inches(0.65), NAVY)
    add_text(s, Inches(0.55), y, Inches(1.0), Inches(0.65), head,
             size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    add_rect(s, Inches(1.6), y, Inches(11.3), Inches(0.65), CREAM, line=BORDER)
    add_text(s, Inches(1.8), y+Inches(0.05), Inches(11), Inches(0.3),
             body, size=11, bold=True, color=NAVY)
    add_text(s, Inches(1.8), y+Inches(0.35), Inches(11), Inches(0.3),
             ex, size=10, color=MUTED, italic=True)
footer(s, 55)
add_notes(s, """
The four-part structure is non-negotiable. If any part is missing, the
LEA can push back and the LOF may not survive review.

Common errors:
- "Aisle too narrow" — missing the measurement.
- "Aisle measures 84"" — missing the standard reference.
- "Violates ADA" — missing the specific section.
- "Cite: 502.3.3" — missing the standard year.

The order in the sentence doesn't matter — what matters is that all
four parts are present.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 56 — Anatomy of a Corrective Action
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Anatomy of a Corrective Action entry",
                            kicker="Slide 56 · Writing findings")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "A defensible corrective action tells the LEA exactly what to do, by when, against what standard.",
         size=14, italic=True, color=MUTED)

add_rect(s, Inches(0.55), Inches(2.05), Inches(12.3), Inches(1.6), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(2.05), Inches(0.18), Inches(1.6), GOLD)
add_text(s, Inches(0.85), Inches(2.2), Inches(12), Inches(0.4),
         "EXAMPLE — fully-formed corrective action:",
         size=11, bold=True, color=NAVY, font="Consolas")
add_text(s, Inches(0.85), Inches(2.6), Inches(11.7), Inches(1.0),
         "Restripe the front parking lot to provide a van-accessible parking space with an access aisle of not less than 96 inches in width, in conformance with 2010 ADA Standards § 502.3.3. Provide photographic documentation of completed restriping. Complete within 45 days of receipt of this Letter of Findings. Cite: 2010 ADA 502.3.3.",
         size=12, italic=True, color=TEXT)

parts = [
    ("ACTION",     "What the LEA must do — concrete and verifiable.",       "\"Restripe the front parking lot to provide a van-accessible parking space...\""),
    ("STANDARD",   "The 2010 ADA dimensional target — ALWAYS.",             "\"...with an access aisle of not less than 96 inches in width per 2010 ADA § 502.3.3.\""),
    ("EVIDENCE",   "How the LEA demonstrates completion.",                  "\"Provide photographic documentation of completed restriping.\""),
    ("DEADLINE",   "Always 45 days from LOF receipt.",                      "\"Complete within 45 days of receipt of this Letter of Findings.\""),
    ("CITE",       "Cite 2010 ADA — even when the violation cites ANSI / UFAS / 1991 ADA.",
     "\"Cite: 2010 ADA 502.3.3.\""),
]
for i, (head, body, ex) in enumerate(parts):
    y = Inches(3.85 + 0.62*i)
    add_rect(s, Inches(0.55), y, Inches(1.0), Inches(0.55), NAVY)
    add_text(s, Inches(0.55), y, Inches(1.0), Inches(0.55), head,
             size=10, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    add_rect(s, Inches(1.6), y, Inches(11.3), Inches(0.55), CREAM, line=BORDER)
    add_text(s, Inches(1.8), y+Inches(0.03), Inches(11), Inches(0.26),
             body, size=10, bold=True, color=NAVY)
    add_text(s, Inches(1.8), y+Inches(0.28), Inches(11), Inches(0.26),
             ex, size=10, color=MUTED, italic=True)
footer(s, 56)
add_notes(s, """
Corrective actions must be specific enough that the LEA cannot
satisfy them with a vague "we'll look into it" response. Photographic
documentation is the gold standard for evidence — it's harder to fake
than a self-attested statement.

THE 2010 ADA CITATION RULE in this column is per CRR 20 boilerplate
and OCR practice. Even if the violation is cited at ANSI A117.1 for an
old element, the corrective action is written to bring the element into
compliance with 2010 ADA — because the corrective work itself is new
work, and 2010 ADA is the federal target state.

The 45-day deadline is set by CDE OEO's Voluntary Compliance Plan
process. Some corrective actions cannot reasonably be completed in 45
days (a full restroom remodel, for example) — in those cases, the
corrective action requires the LEA to submit a project schedule and
funding commitment within 45 days, with the work to follow.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 57 — Citation placement rules
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Citation placement — two rules, one row",
                            kicker="Slide 57 · Writing findings")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.45),
         "VIOLATION cite = standard in effect when the element was built/altered.  CORRECTIVE cite = always 2010 ADA.",
         size=14, italic=True, color=MUTED)

add_text(s, Inches(0.55), Inches(2.1), Inches(12.3), Inches(0.4),
         "Same physical problem — handrail at 30\" — three different VIOLATION cites; SAME corrective cite:",
         size=13, bold=True, color=NAVY)

scenarios57 = [
    ("Ramp installed 1988", "ANSI A117.1 controls",
     "Violation:  ANSI A117.1 § 4.8.5",
     "Corrective:  2010 ADA § 505.4",
     B_ANSI, RGBColor(0x71,0x3F,0x12)),
    ("Ramp installed 2003", "1991 ADA controls",
     "Violation:  1991 ADA § 4.8.5",
     "Corrective:  2010 ADA § 505.4",
     B_1991, RGBColor(0x06,0x4E,0x3B)),
    ("Ramp installed 2016", "2010 ADA controls",
     "Violation:  2010 ADA § 505.4",
     "Corrective:  2010 ADA § 505.4",
     B_2010, RGBColor(0x7F,0x1D,0x1D)),
]
for i, (when, std, viol_cite, corr_cite, fill, tcol) in enumerate(scenarios57):
    y = Inches(2.6 + 1.4*i)
    add_rect(s, Inches(0.55), y, Inches(3.2), Inches(1.2), CREAM, line=BORDER)
    add_rect(s, Inches(0.55), y, Inches(0.18), Inches(1.2), fill)
    add_text(s, Inches(0.75), y+Inches(0.15), Inches(3.0), Inches(0.4), when,
             size=14, bold=True, color=NAVY)
    add_text(s, Inches(0.75), y+Inches(0.55), Inches(3.0), Inches(0.5), std,
             size=11, color=TEXT, italic=True)
    # Violation cite box
    add_rect(s, Inches(3.95), y, Inches(4.3), Inches(1.2), fill, line=BORDER)
    add_text(s, Inches(4.1), y+Inches(0.25), Inches(4.0), Inches(0.3), "VIOLATION column:",
             size=10, bold=True, color=tcol, font="Consolas")
    add_text(s, Inches(4.1), y+Inches(0.55), Inches(4.0), Inches(0.5), viol_cite,
             size=13, bold=True, color=tcol, font="Consolas")
    # Corrective cite box — always 2010 ADA gold
    add_rect(s, Inches(8.45), y, Inches(4.4), Inches(1.2), B_2010, line=GOLD)
    add_text(s, Inches(8.6), y+Inches(0.25), Inches(4.1), Inches(0.3), "CORRECTIVE column:",
             size=10, bold=True, color=RGBColor(0x7F,0x1D,0x1D), font="Consolas")
    add_text(s, Inches(8.6), y+Inches(0.55), Inches(4.1), Inches(0.5), corr_cite,
             size=13, bold=True, color=RGBColor(0x7F,0x1D,0x1D), font="Consolas")

footer(s, 57)
add_notes(s, """
This is the dual-citation rule. Reviewers must internalize both halves:

VIOLATION column — cite the standard in effect when the ELEMENT was
constructed or altered. ANSI for 1977-1991 elements. 1991 ADA for
1992-2010 elements. 2010 ADA for 2012+ elements.

CORRECTIVE column — always cite 2010 ADA. Why: the corrective work is
new construction happening today, and today's federal standard is
2010 ADA. The CRR 20 boilerplate is explicit on this.

Note the row 3 in the table: when both the violation and the corrective
are 2010 ADA (because the element was built under 2010 ADA), the two
cites look identical. They're still two cites — one in each column.

Common mistakes by new reviewers:
- Citing 2010 ADA in the violation column for a 1985 element (wrong —
  that's a misapplication of law).
- Citing ANSI in the corrective column for a 1985 element (wrong — the
  fix must be 2010 ADA).
- Omitting the cite from the corrective column entirely.

Reviewer tip: if you're unsure which standard applies to the violation,
look at the element's date in the SECTION DETAILS box of the
self-assessment. That date drives the violation cite.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 58 — Knowledge check #5
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); set_bg(s, NAVY)
add_rect(s, 0, 0, Inches(0.6), SLIDE_H, GOLD)
add_text(s, Inches(0.85), Inches(0.6), Inches(12), Inches(0.5),
         "KNOWLEDGE CHECK · 05",
         size=12, bold=True, color=GOLD, font="Consolas")
add_text(s, Inches(0.85), Inches(1.1), Inches(12), Inches(0.9),
         "Rewrite this violation entry",
         size=30, bold=True, color=WHITE, font="Georgia")

add_text(s, Inches(0.85), Inches(2.15), Inches(12), Inches(0.4),
         "ROOKIE REVIEWER WROTE:",
         size=12, bold=True, color=GOLD, font="Consolas")
add_rect(s, Inches(0.85), Inches(2.6), Inches(11.95), Inches(1.3), NAVY_MID, line=RED)
add_text(s, Inches(1.05), Inches(2.7), Inches(11.5), Inches(1.1),
         "\"The handrail on the ramp at the wood shop is too low. The shop was built in 1985. It violates the ADA. Fix it.\"",
         size=15, color=WHITE, italic=True)

add_text(s, Inches(0.85), Inches(4.05), Inches(12), Inches(0.4),
         "WHAT'S MISSING / WRONG? (think before flipping to the next slide)",
         size=12, bold=True, color=GOLD, font="Consolas")

issues = [
    "Specific measurement?",
    "Correct standard for a 1985 building?",
    "Right section citation?",
    "Specific location identifier?",
    "Concrete corrective action with deadline?",
]
for i, q in enumerate(issues):
    y = Inches(4.55 + 0.42*i)
    add_text(s, Inches(1.0), y, Inches(0.3), Inches(0.35), "□",
             size=18, color=GOLD, font="Calibri")
    add_text(s, Inches(1.4), y, Inches(11.5), Inches(0.35), q,
             size=13, color=WHITE)

footer(s, 58)
add_notes(s, """
Rewrite to model:

Violation: "Handrail on the south-side ramp serving the wood shop
measured at 30 inches above the ramp surface. The handrail does not
meet the 34-inch minimum height specified in ANSI A117.1-1961 (R1971)
§ 4.8.5. The building was constructed in 1985 and the handrail has not
been altered; ANSI A117.1 (1961 R1971) is the applicable standard
under 34 CFR § 104.23. Cite: ANSI A117.1 § 4.8.5."

Corrective: "Replace the south-side wood shop ramp handrail at a
height not less than 34 inches and not more than 38 inches above the
ramp surface, with 12-inch horizontal extensions at the top and bottom
parallel to the floor at the bottom, in conformance with 2010 ADA
§ 505.4 and § 505.10. Provide photographic documentation of the
completed installation including dimensional measurements. Complete
within 45 days of receipt of this Letter of Findings.
Cite: 2010 ADA § 505.4 and § 505.10."

NOTE the corrective cites 2010 ADA — NOT ANSI — because corrective
work is new work and 2010 ADA controls all new work.

Errors in rookie version:
1. "Too low" — no measurement
2. "Violates the ADA" — wrong: 1985 unaltered = ANSI
3. No section citation
4. No specific location (which ramp? which side?)
5. "Fix it" — not a corrective action; no deadline; no 2010 ADA target

Have students rewrite individually, then compare to the model. Time: 8
minutes for rewrite + share.

This is also a good moment to revisit Knowledge Check #4½ (slide 54) —
the dual-cite rule lives here.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 59 — CASE STUDY #1 — Mixed-era HS welding shop
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Case study #1 — Mixed-era HS, welding shop",
                            kicker="Slide 59 · Case study")
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Work in pairs. 25 minutes to read, decide, draft. Then compare with the room.",
         size=14, italic=True, color=MUTED)

# Scenario card
add_rect(s, Inches(0.55), Inches(2.05), Inches(8.0), Inches(5.0), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(2.05), Inches(8.0), Inches(0.45), NAVY)
add_text(s, Inches(0.55), Inches(2.05), Inches(8.0), Inches(0.45),
         "SCENARIO",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_text(s, Inches(0.75), Inches(2.55), Inches(7.6), Inches(4.4),
         "Sierra Vista High School — Manufacturing pathway.\n\n"
         "Welding shop is in a free-standing building constructed 1979. "
         "Original 1979 construction never altered structurally. In 2018 the LEA installed:\n"
         "  • a new concrete pad outside the south door\n"
         "  • a new south-side ramp from the pad to the door (rise 18\", run 16'-0\")\n"
         "  • new handrails on the south ramp\n"
         "  • ISA signage at the south door\n\n"
         "Field measurements you took:\n"
         "  • South ramp running slope: 1:11 (steeper than 1:12)\n"
         "  • South ramp handrail height: 32\"\n"
         "  • ISA sign baseline tactile characters: 62\" above floor\n"
         "  • Inside the shop: 8 welding booths, all with 30\" frame openings\n"
         "  • Inside the shop: workbenches all at 36\" with full skirts (no knee/toe)",
         size=11, color=TEXT)

# Tasks
add_rect(s, Inches(8.85), Inches(2.05), Inches(4.0), Inches(5.0), CREAM, line=BORDER)
add_rect(s, Inches(8.85), Inches(2.05), Inches(4.0), Inches(0.45), NAVY)
add_text(s, Inches(8.85), Inches(2.05), Inches(4.0), Inches(0.45),
         "YOUR TASKS",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(9.05), Inches(2.6), Inches(3.7), Inches(4.4), [
    "Which standard applies to the original shop building?",
    "Which standard applies to the 2018 ramp / signage?",
    "Which standard applies to the welding booths?",
    "List each violation with its cite.",
    "Draft the corrective action for the ramp slope.",
    "Identify the Program Access issue (if any).",
], size=11, spacing=8, bullet="?")
footer(s, 59)
add_notes(s, """
Walk the room while pairs work. Common stumbles:
1. Some pairs apply 1991 ADA to the 1979 building (wrong — 1979 is in
   the ANSI window).
2. Some pairs cite ANSI for the 2018 ramp (wrong — 2018 = 2010 ADA).
3. Some pairs miss that the welding booths are 1979 ANSI elements
   (frame opening 30" is below ANSI's 32" door clear width — but
   ANSI doesn't have a specific welding booth spec; treat as door).

Answer key (next slide).
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 60 — Case study #1 answers
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Case study #1 — answer key",
                            kicker="Slide 60 · Case study")
# Standards by element
add_text(s, Inches(0.55), Inches(1.55), Inches(12.3), Inches(0.4),
         "Element-by-element standards application:",
         size=14, bold=True, color=NAVY)

ans = [
    ("Original shop building (1979, no alterations)",
     "ANSI A117.1 (1961 R1971)", B_ANSI, RGBColor(0x71,0x3F,0x12)),
    ("South ramp + handrails (built 2018)",
     "2010 ADA", B_2010, RGBColor(0x7F,0x1D,0x1D)),
    ("ISA signage at south door (installed 2018)",
     "2010 ADA", B_2010, RGBColor(0x7F,0x1D,0x1D)),
    ("Welding booths (original 1979 construction)",
     "ANSI A117.1 (1961 R1971)", B_ANSI, RGBColor(0x71,0x3F,0x12)),
    ("Workbenches (original 1979)",
     "ANSI A117.1 — but ANSI is silent on workstation specs", B_ANSI, RGBColor(0x71,0x3F,0x12)),
]
for i, (el, std, fill, tcol) in enumerate(ans):
    y = Inches(2.05 + 0.6*i)
    add_rect(s, Inches(0.55), y, Inches(7.5), Inches(0.5), CREAM, line=BORDER)
    add_text(s, Inches(0.75), y, Inches(7.3), Inches(0.5), el,
             size=11, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    std_badge(s, Inches(8.3), y+Inches(0.11), std, fill, text_color=tcol, w=Inches(4.5))

# Violations
add_text(s, Inches(0.55), Inches(5.15), Inches(12.3), Inches(0.4),
         "Violations + correctives (note: violation cite varies; corrective always 2010 ADA):",
         size=14, bold=True, color=NAVY)
viols = [
    "Ramp slope 1:11  →  VIOLATION: 2010 ADA § 405.2 (built 2018; 1:12 max).  CORRECTIVE: 2010 ADA § 405.2.",
    "Ramp handrail 32\"  →  VIOLATION: 2010 ADA § 505.4 (built 2018).  CORRECTIVE: 2010 ADA § 505.4.",
    "ISA sign 62\"  →  VIOLATION: 2010 ADA § 703.4.1 (installed 2018).  CORRECTIVE: 2010 ADA § 703.4.1.",
    "Welding booth frame 30\" (1979 element)  →  No coded violation; ANSI 1961 silent on welding booths. Program-access narrative recommendation only.",
]
add_bullets(s, Inches(0.55), Inches(5.55), Inches(12.3), Inches(1.8), viols, size=11, spacing=4)
footer(s, 60)
add_notes(s, """
Walk through each row. Key teaching moments:

- Welding booth opening 30" in a 1979 building: ANSI 4.7 (Doors)
  requires 32" clear — but a welding booth opening is debatable as
  a "door." Be careful here. If you treat it as a door, you can cite
  ANSI 4.7. If you treat it as a fixed shop element, ANSI is silent
  — then the LEA might claim no violation. Best practice: cite ANSI
  4.7 if you believe a wheelchair user must transit the opening,
  and explain in narrative.

- The 1979 workbenches with no knee/toe: ANSI 1961 has no
  workstation requirements. So this is NOT a coded violation. But
  it IS a Program Access concern — if no student in a wheelchair
  can use any workbench, the program is not accessible.

  Narrative recommendation: "The LEA should evaluate whether at
  least one workbench provides knee/toe clearance for a wheelchair
  user, consistent with the program-access analysis under
  34 CFR § 104.22. If no accessible workstation exists, the LEA
  should make a reasonable modification or designate an alternate
  station."
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 61 — CASE STUDY #2 — Modern culinary
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Case study #2 — 2014 culinary classroom",
                            kicker="Slide 61 · Case study")

# Two-column: scenario + abbreviated answer in speaker notes
add_rect(s, Inches(0.55), Inches(1.55), Inches(8.0), Inches(5.5), CREAM, line=BORDER)
add_rect(s, Inches(0.55), Inches(1.55), Inches(8.0), Inches(0.45), NAVY)
add_text(s, Inches(0.55), Inches(1.55), Inches(8.0), Inches(0.45),
         "SCENARIO",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_text(s, Inches(0.75), Inches(2.05), Inches(7.6), Inches(4.95),
         "Bay Shore High School — Hospitality, Tourism & Recreation pathway.\n\n"
         "Culinary classroom built 2014 (single construction event, no alterations).\n\n"
         "Field measurements:\n"
         "  • 12 student prep stations — all at 36\" with full-skirt cabinets, no knee clearance\n"
         "  • 4 cooktops — all at 36\" with no knee clearance\n"
         "  • 1 designated 'ADA prep cart' — 34\" stainless cart with knee clearance, currently stored in a side closet\n"
         "  • Dish-pit spray-arm controls at 50\" reach over a 14\" deep counter\n"
         "  • Walk-in pantry door clear width 31\"\n"
         "  • Hand-wash sink lavatory rim height 37\"\n"
         "  • Tray-slide service line at 36\"\n\n"
         "LEA's position: 'We have an accessible prep cart that we wheel out as needed.'",
         size=11, color=TEXT)

add_rect(s, Inches(8.85), Inches(1.55), Inches(4.0), Inches(5.5), CREAM, line=BORDER)
add_rect(s, Inches(8.85), Inches(1.55), Inches(4.0), Inches(0.45), NAVY)
add_text(s, Inches(8.85), Inches(1.55), Inches(4.0), Inches(0.45),
         "YOUR TASKS",
         size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
add_bullets(s, Inches(9.05), Inches(2.1), Inches(3.7), Inches(4.8), [
    "Which standard applies?",
    "How do you respond to the LEA's 'mobile cart' position?",
    "Identify each dimensional violation with its cite.",
    "Draft corrective actions.",
    "Does the dish-pit reach problem need an obstructed-reach analysis?",
], size=11, spacing=10, bullet="?")
footer(s, 61)
add_notes(s, """
Answer summary (for the trainer):

Standard: 2010 ADA (constructed 2014, no alterations).

LEA position pushback: The mobile prep cart, currently stored in a side
closet, is NOT an active part of the classroom configuration. The
"alternate accessible workstation" approach is acceptable under 2010
ADA only if the alternate is in regular active use and is reachable via
an accessible route. Stored-in-closet doesn't qualify.

Violations:
1. No fixed accessible prep station with knee/toe clearance — cite
   2010 ADA 306 (knee/toe) and 902.4 (work surfaces) for the lack of
   an accessible station integrated into normal classroom config.
2. Cooktops all at 36" with no knee clearance — cite 2010 ADA 306
   (no accessible alternate in regular use).
3. Dish-pit spray-arm at 50" over 14" deep counter — obstructed
   forward reach analysis:
   - 14" deep is within 20"–25" range
   - 2010 ADA 308.2.2 allows max 44" reach when obstruction is
     20"–25" deep
   - At 14" deep, the obstruction is < 20" so 48" max applies
   - 50" exceeds 48" max → VIOLATION
   - Cite 2010 ADA 308.2.2.
4. Walk-in pantry door clear width 31" — cite 2010 ADA 404.2.3.
5. Hand-wash lavatory rim 37" — cite 2010 ADA 606.3 (34" max).
6. Tray-slide at 36" — cite 2010 ADA 904.4.1 (34" max).

Corrective actions — all cited at 2010 ADA (even though here the
violation cites are also 2010 ADA because the building is 2014):

1. Provide minimum one fixed accessible prep station with knee/toe
   clearance per 2010 ADA § 306 and accessible work surface per § 902.4.
   Cite: 2010 ADA § 306 and § 902.4.
2. Provide minimum one cooktop at 34" max height with knee/toe
   clearance under, per 2010 ADA § 306. Cite: 2010 ADA § 306.
3. Relocate dish-pit spray-arm controls so the operable parts are
   within the 48" max forward reach per 2010 ADA § 308.2.
   Cite: 2010 ADA § 308.2.
4. Widen walk-in pantry door to provide 32" minimum clear width per
   2010 ADA § 404.2.3. Cite: 2010 ADA § 404.2.3.
5. Re-mount hand-wash sink so lavatory rim is at 34" max above floor
   per 2010 ADA § 606.3. Cite: 2010 ADA § 606.3.
6. Re-mount tray-slide service line at 34" max above floor per 2010
   ADA § 904.4.1. Cite: 2010 ADA § 904.4.1.

All within 45 days of LOF receipt. All cites in the corrective column
are 2010 ADA — that is the rule from Slide 53.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 62 — Resources & references
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); header_bar(s, "Resources & references",
                            kicker="Slide 62 · Wrap-up")
# Three columns: standards docs, agency guidance, internal tools
def res_col(x, y, w, h, title, items):
    add_rect(s, x, y, w, h, CREAM, line=BORDER)
    add_rect(s, x, y, w, Inches(0.45), NAVY)
    add_text(s, x, y, w, Inches(0.45), title.upper(),
             size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font="Consolas")
    for i, (head, body) in enumerate(items):
        ry = y + Inches(0.55) + Inches(0.65*i)
        add_text(slide=s, x=x+Inches(0.15), y=ry, w=w-Inches(0.25), h=Inches(0.25),
                 text=head, size=11, bold=True, color=NAVY)
        add_text(slide=s, x=x+Inches(0.15), y=ry+Inches(0.25), w=w-Inches(0.25), h=Inches(0.4),
                 text=body, size=10, color=TEXT, italic=True)

res_col(Inches(0.55), Inches(1.55), Inches(4.0), Inches(5.4), "Standards Documents",
    [("28 CFR Part 35",       "Title II ADA regulation; App. B = 2010 ADA"),
     ("28 CFR Part 36 App. D", "1991 ADA Standards (ADAAG)"),
     ("34 CFR Part 104",      "§ 504 regulation. §§ 104.22 / 104.23 are core."),
     ("ANSI A117.1-1961 (R1971)", "Original ANSI accessibility standard"),
     ("UFAS (1984)",          "41 CFR § 101-19.6 App. A"),
     ("Access Board guides",   "access-board.gov — Title II tech assistance"),
    ])
res_col(Inches(4.7), Inches(1.55), Inches(4.0), Inches(5.4), "Federal Guidance",
    [("ADA.gov",              "DOJ Technical Assistance, settlement archive"),
     ("Project Civic Access",  "DOJ K-12 settlement agreements"),
     ("OCR Case Search",       "www2.ed.gov/ocr — open case decisions"),
     ("OCR Dear Colleague letters", "Disability-rights guidance for ed agencies"),
     ("U.S. Access Board",     "Technical bulletins, animations, training"),
     ("DSA (CA)",              "dgs.ca.gov/dsa — CA construction approvals"),
    ])
res_col(Inches(8.85), Inches(1.55), Inches(4.0), Inches(5.4), "CDE / Internal",
    [("CRR Procedures Manual", "Reviewer SOPs, evidence checklist"),
     ("Facilities Review Guide", "LEA self-assessment template (Word, with date pickers + era dropdowns)"),
     ("CRR 20 / CRR 21 boilerplate", "Standard LOF cover + 18-row finding table"),
     ("CDE OEO team lead",     "Your team lead — escalate ambiguity"),
     ("CDE accessibility liaison", "Coordinates with DSA on construction-date evidence"),
     ("Knowledge-check slides",   "Slides 18, 22, 30, 38, 54, 58 — rehearse before site visits"),
    ])
footer(s, 62)
add_notes(s, """
Hand out printed copies of: (1) the decision tree (Slide 13), (2) the
six date windows (Slide 14), and (3) the three quick-reference slides
(39-42). Reviewers will use these on every site visit.

Direct reviewers to bookmark ADA.gov and the U.S. Access Board site —
both are excellent for quick lookups in the field.
""")

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 63 — Course summary & next steps
# ─────────────────────────────────────────────────────────────────────────────
s = add_slide(); set_bg(s, CREAM)
add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
add_rect(s, 0, Inches(2.4), SLIDE_W, Inches(0.08), GOLD)

add_text(s, Inches(0.7), Inches(0.7), Inches(12), Inches(0.4),
         "COURSE SUMMARY",
         size=11, bold=True, color=GOLD, font="Consolas")
add_text(s, Inches(0.7), Inches(1.1), Inches(12), Inches(1.1),
         "You can now run a CRR 20 / 21 review.",
         size=36, bold=True, color=WHITE, font="Georgia")

add_text(s, Inches(0.7), Inches(2.7), Inches(12), Inches(0.4),
         "What you take away today:",
         size=14, color=GOLD_LT, italic=True)
takeaways = [
    "The five standards, the six date windows, and the decision tree.",
    "The Program-Access doctrine and the absolute None / None rule.",
    "Key dimensions for parking, routes, doors, ramps, restrooms, and signage.",
    "How to walk a CTE shop, kitchen, lab, or ag building.",
    "How to drive the Facilities LOF Generator from packet to Word output.",
    "How to write a defensible violation and corrective action — every time.",
]
for i, t in enumerate(takeaways):
    y = Inches(3.2 + 0.42*i)
    add_text(s, Inches(0.9), y, Inches(0.5), Inches(0.35), "✓",
             size=18, bold=True, color=GOLD, font="Calibri")
    add_text(s, Inches(1.4), y, Inches(11), Inches(0.35), t,
             size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

# Next steps
add_text(s, Inches(0.7), Inches(6.0), Inches(12), Inches(0.4),
         "NEXT STEPS",
         size=11, bold=True, color=GOLD, font="Consolas")
add_text(s, Inches(0.7), Inches(6.4), Inches(12), Inches(0.4),
         "Shadow two reviews · run one mock LOF · sign up for the Access Board's webinar series",
         size=14, color=GOLD_LT)

add_text(s, Inches(0.7), Inches(6.9), Inches(12), Inches(0.4),
         "Questions? Reach out to your CDE OEO team lead.",
         size=12, color=GOLD_LT, italic=True)
footer(s, 63)
add_notes(s, """
Wrap-up. Three concrete next steps for every new reviewer:

1. SHADOW two complete reviews with a senior reviewer before flying
   solo. Watch how they walk a site, how they measure, how they handle
   pushback from the LEA.

2. RUN one mock LOF using the Facilities LOF Generator with a prior
   completed packet (from the archives). Have a senior reviewer check
   your output before exporting to Word.

3. SIGN UP for the U.S. Access Board's monthly webinar series — free,
   excellent technical content, counts toward your CEU requirements.

Open the floor for final questions. Distribute the printed handouts
(decision tree, quick-reference cards). Thank everyone for their
attention. End of training.
""")

# ─────────────────────────────────────────────────────────────────────────────
# Save
# ─────────────────────────────────────────────────────────────────────────────
import os
os.makedirs('outputs', exist_ok=True)
out_path = 'outputs/CRR_Accessibility_Training.pptx'
prs.save(out_path)
print(f"Saved {len(prs.slides)} slides → {out_path}")
