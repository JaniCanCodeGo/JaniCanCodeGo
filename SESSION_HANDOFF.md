# CRR Project — Session Handoff

**Purpose of this file:** Hand this project off to a new Claude session with full context.
Read this file first. It explains what exists, why, the key decisions, and what's left to do.

**Owner:** Murjani McTier (mmctier@cde.ca.gov), California Department of Education,
Office of Equal Opportunity.

**Repository:** `janicancodego/janicancodego`
**Working branch:** `claude/crr-compliance-agent-pjosJ`
**Last updated:** 2026 review cycle work (2025-26 CRR instrument)

---

## What this project is

Tools to help CDE reviewers run the **Civil Rights Review (CRR)** — the annual
compliance review of California schools/LEAs for their CTE programs. The CRR
instrument has **20 sections (CRR 01–20)** for the **2025-26 cycle**, each with
multiple "evidence requests" a school submits to prove compliance.

The instrument **changes every year.** The current tools are built on the **2025-26**
instrument. Future years will need the instrument text updated (see "Yearly updates" below).

---

## The files in this repo (what each one does)

### 1. `crr_project_instructions.txt`  ← the main deliverable, most recent work
Plain-text **system prompt / project instructions** for a **Claude.ai Project**.
The owner has **Claude Pro** (no API key, no billing). Pasting this text into a
Claude.ai Project's custom instructions turns any chat in that project into a
CRR evidence reviewer.

How it behaves once installed:
- Asks which of the 20 CRR sections (shows the numbered list).
- Shows that section's evidence requests as a **numbered list** with abbreviations.
- Asks for school/LEA name and first-submission-vs-resubmission.
- Asks the user to upload the evidence document(s).
- Runs a **4-step review protocol** (identify → inventory → requirements check → verdict).
- Produces **two outputs**: a **CMT Comment** (professional prose for the CDE
  Monitoring Tool) and a **Reviewer Note** (internal candid notes).

Every one of the 20 sections has **per-evidence-request checkbox requirements**
taken word-for-word from the 2025-26 instrument. CRR 01 is fully broken down
(coordinator contact elements per document type), matching how CRR 02 is broken
down (elements 2.0–2.6).

### 2. `crr_evidence_reviewer.html`  ← standalone browser tool (needs an API key)
A single-file HTML app that does the same review but calls the Anthropic API
directly from the browser. **This one requires an Anthropic API key** (entered
via a 🔑 icon, stored only in sessionStorage — never persisted, never shared).
The owner chose the Projects approach instead because Pro covers it with no key.
Kept in the repo as an alternative. Model used: `claude-sonnet-4-6`.

### 3. `crr_agent.py`  ← Python agent for the *back end* of the review
Different job from the reviewer above. This one processes a **completed CRR
Summary of Findings (.docx)**: extracts Required Corrective Actions, fills the
Voluntary Compliance Plan (VCP), generates the LOF Cover Letter (Findings vs
No Findings), and looks up superintendent + COE Monitoring Lead from the CDE
directory.

### 4. `crr_mcp_server.py`  ← exposes `crr_agent.py` to Claude Desktop as a tool
### 5. `setup_claude_desktop.bat`  ← Windows helper to wire up the MCP server
### 6. `templates/`, `samples/`, `requirements.txt`  ← supporting files for the Python agent

---

## Key decisions already made (don't re-litigate these)

1. **No shared/hardcoded API keys — ever.** Security requirement from the owner.
   The HTML tool only accepts a key in the browser session (sessionStorage).
2. **Primary path is Claude.ai Projects, not the API.** The owner has Pro, not
   API credits. The project-instructions file is the main deliverable.
3. **"Agent" = better prompting, not multi-pass tool use.** The improvement is
   the structured 4-step protocol + exact instrument requirements baked in.
4. **The owner cannot code JSON.** Any yearly-update path must be doable by
   uploading a .docx, not editing code.
5. **Two required outputs every review:** CMT Comment (prose) + Reviewer Note (internal).

## Compliance rules that are easy to get wrong (carry these forward)

- A **typed name is NEVER a wet signature.** Wet-signature-required + typed/absent = Does Not Meet.
- **Title IX Coordinator needs ALL FOUR:** name/title, address, phone, **email**.
- **ADA/Title II and Section 504 Coordinators do NOT need email** (name/title, address, phone).
- **Site & Floor Plans must be actual floor plans** (interior room layouts), not aerial photos/site maps.
- **PII must be fully redacted** on written determinations, notices, IEPs, 504 plans, enrollment data.
- **Employee demographics:** no names, ID numbers, or SSNs.
- **Addenda Process** (not reprinting) is the remedy when an already-distributed
  document is missing an element — use the exact scripted language in the instructions file.

---

## Yearly updates (how to move to 2026-27 and beyond)

The instrument is released fresh each year as a **.docx**. To update:
1. Upload the new instrument .docx to a Claude session.
2. Ask Claude to extract all sections + evidence requests + Item Instructions
   (prior extraction used Python `zipfile` to read `word/document.xml` and strip
   XML tags, because the .docx is too large/binary for the Read tool).
3. Regenerate `crr_project_instructions.txt` following the SAME structure:
   GROUP 1 (section list) → GROUP 2 (numbered evidence requests per section) →
   GROUP 3 (school + submission type) → GROUP 4 (upload) → 4-step protocol →
   Universal Rules → per-evidence-request checkboxes → two output blocks.
4. Watch for **section count/numbering changes between years** (e.g., the LOF
   skills note 2024-25 had 21 sections with CRR 13 removed and facilities at
   CRR 21; 2025-26 has 20 sections, CRR 13 = Student Financial Assistance,
   facilities at CRR 20). Don't assume numbering is stable across cycles.

---

## Current status / what's done

- ✅ `crr_project_instructions.txt` fully rewritten with exact 2025-26 instrument
  language for all 20 sections. Committed + pushed to the working branch.
- ✅ HTML reviewer bug-fixed (PDF sent as `document` block, auth headers added)
  and committed.
- ✅ Python back-end agent + MCP server in place from earlier work.

## Possible next steps (not yet done — pick up here)

- [ ] Real-world test: install the project instructions in a Claude.ai Project,
      run an actual evidence document through it, confirm the CMT Comment +
      Reviewer Note come out correctly.
- [ ] Decide whether to convert the project-instructions reviewer into a
      packaged **Skill** (the environment already shows related skills like
      `crr-evidence-reviewer`, `lof-draft-generator-2025-26`, etc. — there may
      be overlap worth consolidating).
- [ ] Confirm whether a PR is wanted for the working branch (none created yet;
      do not open one without explicit ask).

---

## How to resume in a new session

Option A — if the new session has this repo:
> "Read SESSION_HANDOFF.md on branch `claude/crr-compliance-agent-pjosJ` and continue."

Option B — if the new session does NOT have the repo:
Upload this file (and `crr_project_instructions.txt` if the work is about the
reviewer) into the chat and say:
> "This is the handoff for my CRR project. Read it and pick up from 'Possible next steps.'"
