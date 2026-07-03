#!/usr/bin/env python3
"""
Clear Enough To Lead (CETL) — MCP Server

No API key needed. Tools build structured prompts; Claude generates the content.

Add to claude_desktop_config.json:

  macOS:   ~/Library/Application Support/Claude/claude_desktop_config.json
  Windows: %APPDATA%\\Claude\\claude_desktop_config.json

  {
    "mcpServers": {
      "cetl": {
        "command": "python3",
        "args": ["/ABSOLUTE/PATH/TO/cetl_mcp_server.py"]
      }
    }
  }

Restart Claude Desktop, then ask things like:
  "Write a lesson plan for the html-basics module."
  "Give me an episode outline for a podcast episode about burnout math."
  "Write a welcome message for a new workgroup member."
"""

import sys
from pathlib import Path

from mcp.server.fastmcp import FastMCP

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))
import cetl_content_agent as agent

mcp = FastMCP("Clear Enough To Lead")


@mcp.tool()
def generate_curriculum_content(content_type: str, module: str = "", context: str = "") -> str:
    """
    Generate teaching content for the CETL "Build With Claude" curriculum.

    content_type options:
      lesson-plan         — full lesson plan for a module
      exercise            — hands-on exercise with a clear deliverable
      troubleshooting-faq — common issues Q&A for a module

    Args:
        content_type: What to generate
        module: Module name (e.g. html-basics, localhost-projects, before-apps, mini-claude-agents)
        context: Any additional detail about the learners or session
    """
    return agent.build_curriculum_prompt(content_type, module=module, context=context)


@mcp.tool()
def generate_podcast_content(content_type: str, topic: str = "", episode: str = "", context: str = "") -> str:
    """
    Generate content for the "Clear Enough To Lead" podcast.

    content_type options:
      episode-outline    — full episode structure
      show-notes         — post-episode summary and pull-quote
      interview-questions — guest interview questions

    Args:
        content_type: What to generate
        topic: Episode topic (for episode-outline)
        episode: Episode title (for show-notes)
        context: Guest info, background, or other detail
    """
    return agent.build_podcast_prompt(content_type, topic=topic, episode=episode, context=context)


@mcp.tool()
def generate_audience_content(content_type: str, context: str = "") -> str:
    """
    Generate content directed at CETL workgroup members.

    content_type options:
      welcome-message   — for a new member joining the workgroup
      encouragement-note — for a member who is stuck or feeling behind
      assignment-recap  — recap after a session or exercise

    Args:
        content_type: What to generate
        context: Situation detail (e.g. what they're stuck on, what session just happened)
    """
    return agent.build_audience_prompt(content_type, context=context)


@mcp.tool()
def generate_host_content(content_type: str, module: str = "", topic: str = "", context: str = "") -> str:
    """
    Generate Jani's own prep content for teaching sessions or podcast recording.

    content_type options:
      teaching-script      — talking points for leading a workgroup session
      podcast-intro-script — opening monologue for a podcast episode

    Args:
        content_type: What to generate
        module: Module being taught (for teaching-script)
        topic: Episode topic (for podcast-intro-script)
        context: Any additional detail
    """
    return agent.build_host_prompt(content_type, module=module, topic=topic, context=context)


@mcp.tool()
def generate_promo_content(content_type: str, platform: str = "", about: str = "", context: str = "") -> str:
    """
    Generate marketing/outreach content for CETL.

    content_type options:
      social-post — platform-specific promotional post
      newsletter  — email update for subscribers

    Args:
        content_type: What to generate
        platform: Social platform (for social-post) e.g. instagram, linkedin
        about: What the post/newsletter is about
        context: Any additional detail
    """
    return agent.build_promo_prompt(content_type, platform=platform, about=about, context=context)


if __name__ == "__main__":
    mcp.run(transport="stdio")
