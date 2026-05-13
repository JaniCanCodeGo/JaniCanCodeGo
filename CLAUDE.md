# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

This is a **Python CLI/MCP agent** that automates California Civil Rights Review (CRR) document generation. Given a completed CRR Summary of Findings `.docx`, it:
1. Parses the document to extract school metadata and CRR section data
2. Scrapes the California Department of Education (CDE) web directory for school address, superintendent, and COE Monitoring Lead
3. Fills two output `.docx` templates: a Voluntary Compliance Plan (VCP) and an LOF Cover Letter (Findings or No Findings variant)

## Setup

```bash
pip install -r requirements.txt
```

Dependencies: `python-docx`, `mcp`, `requests`, `beautifulsoup4`

## Running

**CLI mode:**
```bash
python3 crr_agent.py path/to/Summary_of_Findings.docx
python3 crr_agent.py path/to/Summary_of_Findings.docx --output-dir /path/to/outputs/
```

Custom templates can be specified via `--vcp-template`, `--lof-findings-template`, `--lof-no-findings-template`.

**MCP mode (Claude Desktop):**
```bash
# Configure claude_desktop_config.json (macOS: ~/Library/Application Support/Claude/, Windows: %APPDATA%\Claude\)
{
  "mcpServers": {
    "crr-agent": {
      "command": "python3",
      "args": ["/absolute/path/to/crr_mcp_server.py"]
    }
  }
}
```
Windows users can run `setup_claude_desktop.bat` to auto-configure this.

**Manual smoke test** (no formal test suite):
```bash
python3 crr_agent.py samples/Marysville_Charter_Academy_LOF_Sample.docx
```

## Architecture

There are two entry points into the same logic:
- **`crr_agent.py`** — standalone CLI; all core functions live here
- **`crr_mcp_server.py`** — thin MCP wrapper that imports `crr_agent` and exposes two tools (`process_crr_document`, `list_crr_templates`) via FastMCP over stdio

### Processing pipeline (`crr_agent.py`)

```
parse_summary_of_findings()
  ├─ Reads first 15 body paragraphs for: school_name, cds_code, review_dates, coordinator, reviewer
  ├─ Reads document section headers for: district name
  ├─ Derives county from first 2 digits of CDS code (CA_COUNTY_CODES map)
  ├─ Calls lookup_school_cde() → CDE School Directory scrape
  ├─ Calls lookup_superintendent() → CDE district page scrape (uses first 7 digits of CDS + "0000000")
  ├─ Calls lookup_coe_lead() → CDE CAIS leads page scrape
  └─ Walks doc.element.body to collect CRR sections by Heading 1/Heading 2 styles
       ├─ Heading 1 matching "CRR \d+: Title" starts a new crr_section dict
       ├─ Heading 2 "Required Corrective Action" sets state → collects body paragraphs as corrective_actions
       └─ For CRR sections with "accessible facilit" in title: reads corrective actions from table cells instead

get_findings() → filters crr_sections to those with non-placeholder corrective_actions

fill_vcp() → opens VCP_template.docx, does text replacements, then appends finding rows to doc.tables[0]
fill_cover_letter() → opens LOF_*_template.docx, does text replacements, then calls _fill_address_block()
  └─ _fill_address_block() targets specific paragraph indices (e.g. paras[4], paras[5], paras[25])
```

### Key technical details

**Word document text replacement is non-trivial.** Word splits paragraph text across multiple `Run` objects at style boundaries. `_replace_adjacent_runs()` handles this by sliding a window of up to 6 consecutive runs and matching the combined text. Any change to template `.docx` files may shift which run combinations are needed and break replacement silently.

**CDE scraping has a regex fallback.** `beautifulsoup4` is listed as a dependency but all three scraping functions (`_soup_label_value`, `_soup_all_label_values`, `lookup_superintendent`, `lookup_coe_lead`) catch `ImportError` and fall back to regex. The CDE pages use `<tr>/<th>/<td>` and `<dt>/<dd>` patterns; scraping may fail silently if the CDE site is unreachable (sandbox environments, etc.).

**MCP server must never print to stdout.** The MCP stdio transport uses stdout for protocol messages. `crr_mcp_server.py` wraps all agent calls with `contextlib.redirect_stdout` to capture `print()` output from `crr_agent.py` into a string buffer.

**Paragraph index assumptions in cover letter.** `_fill_address_block()` references `doc.paragraphs[4]`, `[5]`, `[6]`, `[7]`, and `[25]` by hard-coded index. If the LOF template structure changes, these indices must be updated.

**CDS code structure.** CDS codes are 14-digit California school identifiers: digits 1–2 = county, digits 1–7 = district, digits 1–14 = school. `lookup_superintendent()` constructs the district CDS by taking the first 7 digits and appending `0000000`.

## Templates

Templates live in `templates/` and use bracketed placeholders like `[School Name]`, `[Date]`, `[45 days after Date]`. These are replaced via `replace_in_doc()`. The VCP template contains one table; the agent removes blank rows and appends finding rows. Template `.docx` files are binary and tracked in git.

Output files are named `{SchoolName}_VCP.docx` and `{SchoolName}_LOF_Cover_Letter.docx` and saved to `outputs/` (gitignored).
