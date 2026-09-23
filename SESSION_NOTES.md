# Session resume notes — CDE OEO Accessibility Training & Tooling

Paste this whole file (or its contents) into a new Claude conversation
to pick up where we left off.

---

## Project context

I am Marisol McTier (`mmctier@cde.ca.gov`), Program Manager at the
California Department of Education (CDE), **Office of Equal Opportunity
(OEO)**. I oversee the **Civil Rights Review (CRR)** and Education Equity
Review programs and I have stepped into Program Reviewer due to staff
vacancies.

This repo (`JaniCanCodeGo/JaniCanCodeGo`) contains my CRR document
automation:

- **`crr_agent.py`** — CLI that takes a completed `Summary of Findings`
  and produces a Voluntary Compliance Plan (VCP) + LOF cover letter.
- **`crr_mcp_server.py`** — MCP wrapper exposing `process_crr_document`
  and `list_crr_templates` for Claude Desktop integration.
- **`templates/`** — VCP, LOF (Findings + No-Findings), Summary of
  Findings templates.
- **`facilities_lof_tool.html`** — standalone in-browser tool for
  manually building the **Facilities LOF** (CRR 20 / CRR 21). API key
  and Claude analyze step REMOVED — purely manual data entry now.

## Brand / palette

Mirrors the existing LOF tool exactly:
- **Navy** `#1A2744` · **Navy mid** `#243460`
- **Gold** `#C8A84B` · **Gold light** `#E8D08A`
- **Cream** `#F7F4EE` · **Sage** `#EDF0EB`
- **Border** `#D4CEBD` · **Muted** `#6B6659`
- Fonts: **DM Serif Display** (headings), **DM Sans** (body),
  **DM Mono** (kickers / labels)

---

## The two reviewer rules — corrected

These rules are baked into the PPT, the Word guide, and the LOF
boilerplate. Do not re-litigate them.

### Rule #1 — Element-by-element alteration analysis

Per 2010 ADA § 202.3 (U.S. Access Board scoping guidance) and
28 CFR § 35.151(b):

> "Only those elements or spaces altered are required to comply… If
> a room or space is completely altered (or built new as part of an
> alteration), the entire room or space is fully subject to the
> standards."

What this means in practice:

- **A 1968 restroom with grab bars replaced in 2018** → grab bars
  evaluated under 2010 ADA. The WC, lavatory, mirror, stall geometry,
  signage all remain Program Access (1968 construction). Reviewer can
  cite only the grab bars.
- **A 1985 building with re-roofing in 2020** → re-roofing is not an
  accessibility-affecting alteration. The whole building stays ANSI.
- **A 1962 cafeteria with a new tray-slide installed 2016** → only the
  tray-slide is 2010 ADA. Fixed counters, doorways, signage that were
  not altered remain Program Access.
- **Whole-room exception**: if an entire room is gutted and rebuilt,
  the entire rebuilt room is subject to the alteration-date standard.

### Rule #2 — Corrective actions = always 2010 ADA

Per CRR 20 boilerplate:

> "The Office of Civil Rights requires all corrective actions to be
> made in accordance with 2010 ADA standards; therefore, certain areas
> will not require corrective action as the noted deficiency is within
> the 2010 ADA standards."

**Two columns, two rules, same row:**

| Column | Cite |
|---|---|
| Violation | The standard in effect when the element was built/altered (ANSI / UFAS / 1991 ADA / 2010 ADA) |
| Corrective | **Always 2010 ADA** — irrespective of construction or alteration date |

A 1985 handrail at 30" produces:
- VIOLATION: "...does not meet ANSI A117.1 § 4.8.5..." → **Cite: ANSI A117.1 § 4.8.5**
- CORRECTIVE: "Reset handrail to 34–38" per 2010 ADA § 505.4..." → **Cite: 2010 ADA § 505.4**

### Standard determination by date (unchanged)

| Date built / last altered | Standard |
|---|---|
| ≤ Jun 3, 1977 | Program Access |
| Jun 4, 1977 – Jan 17, 1991 | ANSI A117.1 (1961 R1971) |
| Jan 18, 1991 – Jan 26, 1992 | UFAS (1984) |
| Jan 27, 1992 – Sep 14, 2010 | 1991 ADA / ADAAG |
| Sep 15, 2010 – Mar 14, 2012 | 1991 ADA OR 2010 ADA |
| ≥ Mar 15, 2012 | 2010 ADA |

### Program Access = None / None

Program Access has no measurable dimensional standard. Both columns of
the LOF table read `None.` for any element under Program Access. No
exceptions.

---

## What's been built

### 1. Full-day training deck — DONE ✓ (74 slides)

- **File**: `outputs/CRR_Accessibility_Training.pptx`
- **Script**: `build_training_deck.py`
- **74 slides**, 16:9 widescreen, ~80,000 chars of speaker notes
- Covers ANSI A117.1 (1961), ADA 1991/ADAAG, ADA 2010, UFAS, Program
  Access, applied to K-12 and CTE facilities
- **All five U.S. Access Board guides** incorporated: Alterations,
  Accessible Routes, Entrances/Doors/Gates, Drinking Fountains,
  Lavatories & Sinks. Dimensions and rules added to slides 36, 39,
  40, 42.
- **All tool references removed** (Slides 52–54 are reviewer-rule
  deep dives: element-by-element + corrective = 2010 ADA + KC #4½)
- **Light lavender overlap color** (`#E9D5FF`) on Slides 13 + 14 for
  the September 15, 2010 – March 14, 2012 overlap window
- **6 knowledge checks** (Slides 18, 22, 30, 38, 54, 58)
- **12 case studies** (Slides 59–71): 2 original + 10 new from the
  Access Board guides
- **Slide 60** is a dedicated answer-key slide for Case Study #1.
  All other answers live in Speaker Notes and the answer-key doc.
- **Slide 72** explicitly tells trainees where the answers live

### 2. Standalone Answer Key — DONE ✓

- **File**: `outputs/CRR_Training_Answer_Key.docx` (~49 KB, ~4,560
  words)
- Cover page + Part 1 (all 6 knowledge checks) + Part 2 (all 12 case
  studies)
- Each entry: prompt + focus area + full answer
- Designed for printing; distribute AFTER trainees attempt each
  problem

### 4. Optimized BLANK Facilities Review Guide — DONE ✓ (v2)

- **File**: `outputs/BLANK_Facilities_Review_Guide_optimized.docx`
- **Script**: `build_optimized_guide.py`
- **v1 froze Word** (Repeating Section CCs + missing docPart placeholder
  + duplicate SDT IDs). v2 fixes all three.
- All 18 area sections now have:
  - **SECTION DETAILS** table (location, original construction date,
    original construction era → standard dropdown, "altered?" Yes/No)
  - **ALTERATIONS LOG** table — element-by-element capture (element
    altered, date altered, era → standard at alteration, description)
  - 4 blank alteration rows per area
- **Section Templates page at the back** with copy-paste templates for
  the 14 multi-instance area types — no Repeating Section CCs needed
- Instructions block near the top explains the element-by-element rule
  and the corrective = 2010 ADA rule explicitly
- 661 SDTs, all with unique IDs, no docPart placeholder dependencies

### 5. Facilities LOF tool (`facilities_lof_tool.html`) — UPDATED ✓

- **API key step REMOVED**
- **Claude analyze step REMOVED**
- Three steps: School Info → Upload → Enter & Edit Findings
- 18 area rows auto-populate blank when a packet is uploaded
- All four LOF columns editable (location/dates inline, standard
  dropdown, violation / corrective free-text)
- Export to Word still works
- No network call to Anthropic anywhere

### 6. Interactive training HTML — PENDING

To be built. Will mirror the PPT as a standalone self-paced course
with inline multiple-choice knowledge checks, scenario walkthroughs
for the case studies, progress bar in `localStorage`, trainer/trainee
mode toggle, same navy/gold/DM-fonts palette.

This is Claude's equivalent of Gemini Canvas — same approach as
`facilities_lof_tool.html` (single self-contained HTML).

### 7. Correct self-evaluation HTML — STILL AWAITING UPLOAD

The user mentioned there is a SEPARATE HTML for LEAs to self-evaluate
their facilities. We have NOT yet seen it. Once uploaded, the
optimized BLANK Word guide will be cross-checked against its question
schema. If the schema differs, the guide will be regenerated to match.

---

## Open design questions for #6 (interactive training HTML)

- One big HTML file covering all 74 slides, or modular (one file per
  major section)?
- Include a "test out" certification exam at the end?
- Save progress in `localStorage` only, or also export a completion
  PDF?
- Pull case-study answers from `CRR_Training_Answer_Key.docx`, or
  re-embed the content directly in the HTML?

---

## Reference materials uploaded (USE THESE)

The user uploaded U.S. Access Board technical guides as the knowledge
base for the 2010 ADA. Path: `/root/.claude/uploads/e22f8a78-…`:

- `71e301ba-alterations.pdf` — **THE** alteration / addition guide
  (Chapter 2 of Access Board's "ADA Scoping" technical guide)
- `a609c205-accessibleroutes.pdf` — Accessible routes guide (18 pp)
- `c45859ba-Entrances_Doors_and_Gates.pdf` — Entrances/doors (23 pp)
- `a190ea6d-drinkingfountains.pdf` — Drinking fountains (62 pp)
- `daa66d70-lavssinks.pdf` — Lavatories & sinks (170 pp)

These are 2010 ADA; the core principles (especially the
element-by-element alteration rule) are authoritative.

---

## Branch / git state

- Branch: `claude/accessibility-training-powerpoint-JJTQJ`
- Repo: `JaniCanCodeGo/JaniCanCodeGo`
- `.gitignore` excludes `outputs/` — generated docs force-added with
  `git add -f`.

---

## How to pick this up in a new chat

Paste the entire contents of this file into a new conversation, then
say one of:

> "Resume from SESSION NOTES. The user uploaded the correct self-eval
> HTML — cross-check the optimized guide schema."

> "Resume from SESSION NOTES. I want to build the interactive training
> HTML (item #4)."

> "Resume from SESSION NOTES. The Word guide still has [problem] —
> please fix."

…and Claude will continue from this state.
