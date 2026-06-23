# Installed Claude Code Skills

These are community-made skills (MIT-licensed) installed for use with Claude Code
in this repository. They load automatically when their trigger conditions are met,
or you can ask Claude to use them by name.

## stop-slop
Source: https://github.com/hardikpandya/stop-slop

Strips the robotic "AI voice" out of writing — clichés, throat-clearing openers,
passive voice, em dashes, vague filler. Makes output read like a human wrote it.

**Use it for:** grant narratives, business-idea posts, podcast show notes and
scripts, ADHD-community teaching materials, and CRR/LOF letters. Anywhere the
writing needs to sound credible and human.

How to trigger: just ask Claude to "clean this up with stop-slop" or "remove the
AI patterns from this draft."

## karpathy-llm-wiki
Source: https://github.com/Astro-Han/karpathy-llm-wiki

Builds and maintains a self-organizing knowledge base. You drop source material
(PDFs, articles, notes, pasted text) into a `raw/` folder; Claude compiles it into
an interlinked, cited wiki under `wiki/` that compounds over time. You pick the
sources and ask questions; Claude maintains the library.

**Use it for:** a living wiki of grants you've researched, civil-rights regs,
business ideas and the notes behind them, podcast topics/guests, and ADHD + AI
teaching material. A low-maintenance "second brain."

How to start: give Claude a source and say "add this to my wiki." The first
ingest creates the `raw/` and `wiki/` folders automatically. Later, ask
"what do I know about X" to query it.

---
Note: these are third-party skills, not official Anthropic skills. Review their
SKILL.md files to see exactly what they do.
