#!/usr/bin/env python3
"""
CRR Civil Rights Review – MCP Server for Claude Desktop

This exposes crr_agent.py as a tool Claude can call directly.

Setup:
  1. pip install mcp python-docx requests beautifulsoup4
  2. Add to claude_desktop_config.json:

  macOS:   ~/Library/Application Support/Claude/claude_desktop_config.json
  Windows: %APPDATA%\\Claude\\claude_desktop_config.json

  {
    "mcpServers": {
      "crr-agent": {
        "command": "python3",
        "args": ["/ABSOLUTE/PATH/TO/crr_mcp_server.py"]
      }
    }
  }

  3. Restart Claude Desktop.
  4. Ask Claude: "Process this CRR Summary of Findings: /path/to/LOF.docx"
"""

import io
import sys
import os
from contextlib import redirect_stdout
from datetime import datetime
from pathlib import Path

# MCP must use stdio transport – never write to stdout directly
from mcp.server.fastmcp import FastMCP

# Import the agent logic from the sibling file
SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))
import crr_agent as agent

mcp = FastMCP("CRR Civil Rights Review Agent")


@mcp.tool()
def process_crr_document(
    input_path: str,
    output_dir: str = "",
) -> str:
    """
    Process a completed CRR Summary of Findings (.docx) document.

    Reads the LOF, extracts all Required Corrective Actions, determines
    whether findings exist, fills the Voluntary Compliance Plan (VCP),
    and generates the correct LOF Cover Letter (Findings or No Findings).
    Also looks up the school address, superintendent, and COE Monitoring
    Lead from the CDE School Directory.

    Args:
        input_path: Absolute path to the completed Summary of Findings .docx
        output_dir: Folder where VCP and cover letter will be saved.
                    Defaults to an 'outputs/' folder next to the input file.

    Returns:
        A summary of what was found and the paths to the two generated files.
    """
    input_path = input_path.strip().strip('"').strip("'")
    p = Path(input_path)

    if not p.exists():
        return f"ERROR: File not found: {input_path}"
    if p.suffix.lower() != ".docx":
        return f"ERROR: Expected a .docx file, got: {p.suffix}"

    # Default output dir: sibling folder named 'outputs' next to the input file
    if not output_dir:
        output_dir = str(p.parent / "outputs")
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # Capture all print() output so it doesn't corrupt MCP stdio
    log = io.StringIO()
    today = datetime.now()

    try:
        with redirect_stdout(log):
            metadata, crr_sections = agent.parse_summary_of_findings(str(p))
            findings = agent.get_findings(crr_sections)
            has_findings = bool(findings)

            safe = __import__("re").sub(r"[^\w\s\-]", "", metadata.get("school_name", "School"))
            safe = safe.strip().replace(" ", "_")

            vcp_out = Path(output_dir) / f"{safe}_VCP.docx"
            lof_out = Path(output_dir) / f"{safe}_LOF_Cover_Letter.docx"

            template_dir = SCRIPT_DIR / "templates"
            agent.fill_vcp(
                str(template_dir / "VCP_template.docx"),
                str(vcp_out),
                metadata,
                findings,
                today,
            )

            lof_template = (
                str(template_dir / "LOF_Findings_template.docx")
                if has_findings
                else str(template_dir / "LOF_NoFindings_template.docx")
            )
            agent.fill_cover_letter(lof_template, str(lof_out), metadata, has_findings, today)

    except Exception as e:
        import traceback
        return f"ERROR processing document:\n{traceback.format_exc()}"

    # Build a clean summary to return to Claude
    lines = []
    lines.append("✅ CRR Documents Generated Successfully")
    lines.append("")
    lines.append(f"School:         {metadata.get('school_name', '—')}")
    lines.append(f"District:       {metadata.get('district', '—')}")
    lines.append(f"County:         {metadata.get('county', '—')}")
    lines.append(f"Principal:      {metadata.get('principal_name', '—')}")
    lines.append(f"Superintendent: {metadata.get('superintendent', '—') or '(not found on CDE)'}")
    lines.append(f"COE Lead:       {metadata.get('coe_lead', '—') or '(not found on CDE)'}")
    lines.append(f"Review Dates:   {metadata.get('review_dates', '—')}")
    lines.append(f"LOF Date:       {today.strftime('%B %d, %Y')}")
    lines.append(f"VCP Deadline:   {(today + __import__('datetime').timedelta(days=45)).strftime('%B %d, %Y')} (45 days)")
    lines.append("")

    if findings:
        lines.append(f"Findings ({len(findings)} sections):")
        for f in findings:
            lines.append(f"  • {f['crr_num']}: {f['crr_title']}")
        lines.append("")
        lines.append("Cover Letter:   LOF – Findings")
    else:
        lines.append("Findings:       None — LOF – No Findings cover letter generated")

    cde = metadata.get("cde_info", {})
    if cde:
        lines.append("")
        lines.append("CDE lookup:     " + ", ".join(cde.keys()))
    else:
        lines.append("")
        lines.append("CDE lookup:     Address/email not retrieved (CDE site unreachable or no data)")

    lines.append("")
    lines.append("Generated files:")
    lines.append(f"  VCP:          {vcp_out}")
    lines.append(f"  Cover Letter: {lof_out}")

    remaining_blanks = [
        p for p in ["[Name] (Superintendent)", "[Name] (COE Lead)"]
        if not metadata.get("superintendent") or not metadata.get("coe_lead")
    ]
    if remaining_blanks:
        lines.append("")
        lines.append("Still needs manual fill-in:")
        if not metadata.get("superintendent"):
            lines.append("  • Superintendent name in cc list")
        if not metadata.get("coe_lead"):
            lines.append("  • COE Monitoring Lead name in cc list")
        lines.append("  • School mailing address (if CDE lookup failed)")
        lines.append("  • VCP columns 2–5 (completed by the school)")

    return "\n".join(lines)


@mcp.tool()
def list_crr_templates() -> str:
    """
    List the CRR document templates available to this agent.
    Useful for confirming the agent is set up correctly.
    """
    template_dir = SCRIPT_DIR / "templates"
    sample_dir = SCRIPT_DIR / "samples"

    lines = ["CRR Agent – Available Files", ""]
    lines.append("Templates:")
    for f in sorted(template_dir.glob("*.docx")):
        lines.append(f"  {f.name}")

    lines.append("")
    lines.append("Sample documents:")
    if sample_dir.exists():
        for f in sorted(sample_dir.glob("*.docx")):
            lines.append(f"  {f.name}")
    else:
        lines.append("  (none)")

    lines.append("")
    lines.append(f"Agent root: {SCRIPT_DIR}")
    return "\n".join(lines)


if __name__ == "__main__":
    mcp.run(transport="stdio")
