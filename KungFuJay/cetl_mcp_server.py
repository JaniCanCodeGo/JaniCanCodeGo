#!/usr/bin/env python3
"""
Clear Enough To Lead — MCP Server for Claude Desktop

No API key needed. The tools build structured prompts; Claude generates the content.

Add to claude_desktop_config.json:

  macOS:   ~/Library/Application Support/Claude/claude_desktop_config.json
  Windows: %APPDATA%\\Claude\\claude_desktop_config.json

  {
    "mcpServers": {
      "clear-enough-to-lead": {
        "command": "python3",
        "args": ["/ABSOLUTE/PATH/TO/cetl_mcp_server.py"]
      }
    }
  }

Restart Claude Desktop, then ask things like:
  "Write me an email to my team about the deadline we missed."
  "Give me a fog recovery protocol — I have no idea what I was doing."
  "Build me a realistic work plan. Energy level is low, I have a 10am standup."
"""

import sys
from pathlib import Path

from mcp.server.fastmcp import FastMCP

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))
import cetl_agent as agent

mcp = FastMCP("Clear Enough To Lead")


@mcp.tool()
def generate_for_others(
    content_type: str,
    situation: str = "",
    context: str = "",
    topic: str = "",
    duration: int = 60,
) -> str:
    """
    Generate professional workplace content for use with your team, manager, or HR.

    content_type options:
      email          — workplace email for any situation
      feedback       — script for a feedback conversation with an employee
      meeting-agenda — ADHD-friendly agenda with time boxes and parking lot
      boundary       — boundary-setting script (verbal and written versions)

    Args:
        content_type: What to generate (email, feedback, meeting-agenda, boundary)
        situation: Describe the situation or what the email/meeting is about
        context: Background on people, history, or workplace dynamics
        topic: Meeting topic (for meeting-agenda)
        duration: Meeting length in minutes (for meeting-agenda, default 60)
    """
    return agent.build_for_others_prompt(
        content_type=content_type,
        situation=situation,
        context=context,
        topic=topic,
        duration=duration,
    )


@mcp.tool()
def generate_for_self(
    content_type: str,
    input_text: str = "",
    situation: str = "",
    context: str = "",
    level: str = "medium",
    symptoms: str = "",
    commitments: str = "",
) -> str:
    """
    Generate personal content for managing your own brain, body, and workday.

    content_type options:
      brain-dump     — converts your chaotic thoughts into an organized action list
      fog-recovery   — step-by-step protocol for when you've completely lost the thread
      morning-check  — realistic morning checklist based on your current symptom level
      self-advocacy  — script for requesting accommodations from boss or HR
      energy-plan    — realistic work plan based on your current energy level

    Args:
        content_type: What to generate (brain-dump, fog-recovery, morning-check, self-advocacy, energy-plan)
        input_text: Your raw brain dump or unfiltered thoughts (for brain-dump)
        situation: What you need or what's happening (for self-advocacy)
        context: Any additional context
        level: Current energy level — low, medium, or high (for morning-check, energy-plan)
        symptoms: What you're experiencing today (brain fog, hot flashes, fatigue, etc.)
        commitments: Today's known meetings or deadlines (for energy-plan)
    """
    return agent.build_for_self_prompt(
        content_type=content_type,
        input_text=input_text,
        situation=situation,
        context=context,
        level=level,
        symptoms=symptoms,
        commitments=commitments,
    )


if __name__ == "__main__":
    mcp.run(transport="stdio")
