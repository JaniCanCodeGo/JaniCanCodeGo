# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## About the user (read first)

The user is multifaceted, not a monolith. The Civil Rights Review (CRR) / Education
Equity (EE) work described below is their **day job** — it is NOT the sum of their
goals. They are deliberately branching into independent ventures, including:

- Grant research and grant writing
- An ADHD-focused practice (teaching the ADHD community to use AI to work more efficiently)
- A podcast
- A website / web-scraping projects
- Business ideas (finding and posting them)

**Keep these worlds separate.** Skills, agents, and automations built for the personal
ventures must be organized and kept distinct from the day-job CRR/EE tooling. Do not
default to a CRR/EE framing when making suggestions. When a request could belong to
either world and it is not obvious which, ask.

## What this project does

CRR (Civil Rights Review) Document Processing Agent for California K-12 education compliance. It takes a completed **Summary of Findings** `.docx` and produces two output documents:

1. **Voluntary Compliance Plan (VCP)** – table of CRR sections with required corrective actions and a 45-day deadline
2. **LOF Cover Letter** – either "Findings" or "No Findings" variant, with address block, salutation, and cc list filled in

The agent also performs live lookups against the CDE School Directory to auto-populate the school address, email, superintendent name, and COE Monitoring Lead.

## Usage

### CLI (primary)
```bash
python3 crr_agent.py path/to/Summary_of_Findings.docx
python3 crr_agent.py path/to/Summary_of_Findings.docx --output-dir /some/other/folder
```
Outputs are written to `outputs/` by default.

### MCP Server (Claude Desktop integration)
```bash
python3 crr_mcp_server.py   # runs on stdio transport
```
Use `setup_claude_desktop.bat` on Windows to auto-write `claude_desktop_config.json`. On macOS, manually add to `~/Library/Application Support/Claude/claude_desktop_config.json`.

### Install dependencies
```bash
pip install -r requirements.txt
# or individually:
pip install python-docx mcp requests beautifulsoup4
```

## Architecture

All core logic lives in `crr_agent.py`. `crr_mcp_server.py` is a thin wrapper that imports `crr_agent` and exposes two MCP tools (`process_crr_document`, `list_crr_templates`), capturing stdout to avoid corrupting the MCP stdio transport.

### Processing pipeline in `crr_agent.py`

1. **`parse_summary_of_findings(path)`** — reads the input `.docx`, extracts header metadata (school name, CDS code, review dates, coordinator/principal, district) by scanning the first 15 paragraphs and the document section header. Then walks all body elements, using paragraph style names (`Heading 1`, `Heading 2`) as state machine triggers to collect `CRR N: Title` sections and their `Required Corrective Action` paragraphs. For CRR sections flagged as "accessible facilities", it parses a table instead of paragraphs.

2. **CDE web lookups** (called from inside `parse_summary_of_findings`):
   - `lookup_school_cde(cds_code)` → school address, phone, email, administrator name from `cde.ca.gov/schooldirectory/details`
   - `lookup_superintendent(cds_code)` → district superintendent using first 7 digits of CDS + `0000000`
   - `lookup_coe_lead(county_name)` → COE Monitoring Lead from `cde.ca.gov/ta/cr/caisleads.asp`
   - All lookups degrade gracefully: if `beautifulsoup4` is missing, regex fallbacks are used; network failures print an `[Info]` notice and return empty.

3. **`get_findings(crr_sections)`** — filters to only sections that have non-empty corrective actions (excludes "None", "N/A", blank).

4. **`fill_vcp(...)`** — opens `templates/VCP_template.docx`, replaces text placeholders, removes blank table rows, then appends one row per finding.

5. **`fill_cover_letter(...)`** — opens the appropriate LOF template, does global placeholder replacements via `replace_in_doc()`, then calls `_fill_address_block()` which targets specific paragraph indices (4, 5, 6, 7, 25) to place the address block, salutation, and cc list.

### Key text-replacement mechanism

Word documents store text split across multiple `Run` objects within a paragraph. `_replace_adjacent_runs(para, old, new, max_window=6)` handles this by sliding a window across consecutive runs, concatenating their text, replacing if the old string is found, then clearing the extra runs. `replace_in_doc(doc, replacements)` applies this across all paragraphs in the body, headers, and footers.

### Templates

- `templates/VCP_template.docx` — contains a table with blank rows; first row is a header
- `templates/LOF_Findings_template.docx` — cover letter for when findings exist
- `templates/LOF_NoFindings_template.docx` — cover letter for when no findings exist
- `samples/` — example input documents for testing

The templates use bracket placeholders like `[School Name]`, `[Date]`, `[45 days after Date]`, `[Last name]`, `[initials]`. The address block (paras 4–7, 25) uses human-readable placeholder text like `"Name, Principal"` and `"City, State Zip Code"`.

### CRR 2 — Annual Public Notification: reviewer guidance (May 2026)

Per CDE OEO directive from Randi Solís Thompson (May 12, 2026): the APN is a standalone annual requirement and does not need to appear on every CTE-related flyer, publication, or document. A finding for CRR 2 is only appropriate when the school provided **nothing** for the APN (no website publication, no printed notice, nothing). Findings must **not** be issued solely because the APN was absent from individual CTE materials.

This is guidance for reviewers writing the Summary of Findings — the agent does not enforce it in code. The corrective actions in the VCP will reflect exactly what the reviewer wrote.

### CDS code conventions

California school CDS codes are 14 digits. First 2 digits = county (mapped in `CA_COUNTY_CODES`). First 7 digits + `0000000` = district CDS code used for superintendent lookup. `normalize_cds()` strips spaces/hyphens and zero-pads to 14 digits.

## Output files

Generated files land in `outputs/` (gitignored) named `{SchoolName}_VCP.docx` and `{SchoolName}_LOF_Cover_Letter.docx`, where the school name has special characters stripped and spaces replaced with underscores.
