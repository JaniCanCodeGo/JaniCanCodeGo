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
  generating the **Facilities LOF** (CRR 20 / CRR 21) from an LEA-
  completed Facilities Review Guide. Calls Anthropic API directly from
  the browser using the LEA reviewer's own API key.
- **`Facilities_LOF_Session_Prompt.md`** — the master skill prompt that
  encodes the standard-determination rules.

## Brand / palette (use across all artifacts)

Mirrors the HTML tool exactly:
- **Navy** `#1A2744` · **Navy mid** `#243460`
- **Gold** `#C8A84B` · **Gold light** `#E8D08A`
- **Cream** `#F7F4EE` · **Sage** `#EDF0EB`
- **Border** `#D4CEBD` · **Muted** `#6B6659`
- Fonts: **DM Serif Display** (headings), **DM Sans** (body),
  **DM Mono** (kickers / labels)

## Reviewer rules already encoded everywhere

**Standard determination by date** (memorize):
| Date built / last altered | Standard |
|---|---|
| ≤ Jun 3, 1977 | Program Access |
| Jun 4, 1977 – Jan 17, 1991 | ANSI A117.1 (1961 R1971) |
| Jan 18, 1991 – Jan 26, 1992 | UFAS (1984) |
| Jan 27, 1992 – Sep 14, 2010 | 1991 ADA / ADAAG |
| Sep 15, 2010 – Mar 14, 2012 | 1991 ADA OR 2010 ADA |
| ≥ Mar 15, 2012 | 2010 ADA |

**Hard rule:** Program Access = `None.` / `None.` in both finding
columns of the LOF table. **Always.**

Other reviewer rules (see `Facilities_LOF_Session_Prompt.md` for full
list):
- 18 area rows in fixed order, every LOF.
- All sub-locations of an area collapse to ONE row.
- Cite the standard that matches the date — NOT always 2010 ADA.
- Per CRR 2 guidance from Randi Solís Thompson (May 12, 2026): the APN
  is a standalone annual requirement; do NOT issue a CRR 2 finding for
  individual CTE flyers missing the APN.

---

## What's been built in this conversation

### 1. Full-day training deck — DONE ✓
- **File**: `outputs/CRR_Accessibility_Training.pptx`
- **Script**: `build_training_deck.py` (regenerates the deck)
- **63 slides**, 16:9 widescreen, ~37,000 chars of speaker notes
- Covers ANSI A117.1 (1961), ADA 1991/ADAAG, ADA 2010, UFAS, Program
  Access, applied to K-12 and CTE facilities
- Sections: orientation → legal foundations → 5 standards + decision
  tree → Program Access doctrine → ANSI deep-dive → UFAS → 1991 ADA →
  2010 ADA → measurements reference → CTE deep-dives (15 sectors,
  detailed: Building Trades, Manufacturing/welding, Culinary, Ag,
  Health Science) → LOF tool walkthrough → writing findings → 2 case
  studies with answer keys → resources & wrap-up
- **5 knowledge-check slides** with answers in speaker notes
- **2 case studies** with answer-key slides / notes

### 2. Optimized BLANK Facilities Review Guide — DONE ✓
- **File**: `BLANK_Facilities_Review_Guide_optimized.docx`
  (also copy in `outputs/`)
- **Script**: `build_optimized_guide.py`
- All 18 area sections now have a **"SECTION DETAILS"** block with:
  - **Location / Building name** — plain text content control
  - **Date constructed** — native Word calendar date picker
  - **Era — Standard at construction** — dropdown with 6 options, each
    labeled with the standard name (Program Access / ANSI A117.1 /
    UFAS / 1991 ADA / 1991 ADA OR 2010 ADA / 2010 ADA)
  - **Date of ADA modification** — calendar date picker
  - **Era — Standard after modification** — same dropdown + "Not
    modified" option
  - **Describe modification** — plain text content control
- Each SECTION DETAILS block is wrapped in a **Repeating Section
  Content Control** — Word shows a small `+` button to add another
  instance (Restroom #2, CTE Lab #3, etc.)
- **No macros**, **no API key**, **no security warning**. Works in
  Word 2013+, Word for Mac, Word for the Web.
- Instructions block injected near the top explaining how to use
  the new features.

### 3. HTML tool updated — DONE ✓
- **File**: `facilities_lof_tool.html`
- Added **"📥 Download Blank Guide (.docx)"** button in Step 02.
  When clicked, fetches `./BLANK_Facilities_Review_Guide_optimized.docx`
  from the same directory and triggers download, optionally pre-named
  with the school name from Step 02.
- Updated `systemPrompt` in `runAnalysis()` to recognize both the
  legacy format AND the new optimized format. Looks for
  `SECTION DETAILS —` headers and era-dropdown values; collapses
  multiple instances (Restroom #1, Restroom #2) into one LOF row.

### 4. Interactive training HTML — PENDING
- Discussed in detail. The plan is to build
  `CRR_Accessibility_Training.html` as a standalone self-paced course
  mirroring the PPT, with inline multiple-choice knowledge checks,
  scenario walkthroughs for the 2 case studies, progress bar in
  `localStorage`, trainer/trainee mode toggle, and the same
  navy/gold/DM-fonts palette.
- This is Claude's equivalent of Gemini Canvas — same approach as
  the existing `facilities_lof_tool.html` (single self-contained HTML).

---

## Decisions already made (do not re-litigate)

| Question | Decision |
|---|---|
| Date input mechanism | **Calendar date picker + era dropdown** (label includes standard) |
| Duplicate-section mechanism | **Native Word Repeating Section Content Control** (+ button) |
| Which sections duplicatable | **All 18 areas** |
| Auto-applicable-standard mech | **Dropdown labels include the standard** (no macros, no formulas) |
| HTML tool integration | **Parser update + Download Blank Guide button** |
| HTML download mechanism | **Fetch from same directory** (no base64 bloat) |
| PPT audience | **New CDE OEO reviewers**, full-day (~50–70 slides) |
| PPT visual style | **Clean modern + navy/gold accents** matching tool |
| PPT interactivity | **Knowledge-check slides + 2–3 case study scenarios** |
| Training tool platform | **Standalone HTML** (Claude Artifacts-equivalent), not LMS |

## Open question

**Resume the build of `CRR_Accessibility_Training.html`** (item #4).
Should it:
- be one big HTML covering all 63 slides as interactive pages, or
- be modular (one HTML per major section: legal foundations, the 5
  standards, CTE deep-dives, LOF tool, finding-writing, case studies)?
- include a "test out" exam at the end (20-30 questions) for
  certification?
- save progress in `localStorage` only, or also export a completion
  certificate as a PDF?

---

## Branch / git state

- Branch: `claude/accessibility-training-powerpoint-JJTQJ`
- Repo: `JaniCanCodeGo/JaniCanCodeGo`
- Files committed: `build_training_deck.py`,
  `outputs/CRR_Accessibility_Training.pptx`,
  `build_optimized_guide.py`,
  `BLANK_Facilities_Review_Guide_optimized.docx`,
  `outputs/BLANK_Facilities_Review_Guide_optimized.docx`,
  `facilities_lof_tool.html` (updated)
- `.gitignore` excludes `outputs/` by default — generated docs are
  force-added with `git add -f`.

---

## How to pick this up in a new chat

Paste the entire contents of this file into a new conversation, then
say:
> "Resume from the SESSION NOTES. I want to build the interactive
> training HTML (item #4 / pending)."

…and Claude will continue from this state.
