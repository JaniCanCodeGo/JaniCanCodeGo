# Meeting OS

A [Claude Code](https://claude.com/claude-code) **skill** that turns a
[Granola](https://granola.ai) meeting into an interactive, type-aware HTML
dashboard.

Point it at a meeting and it will:

1. **Pull the transcript** from Granola (via the Granola MCP connector)
2. **Detect the meeting type** — brainstorm, planning, sales, standup, 1:1,
   interview, status review, or decision
3. **Extract** decisions, action items, open questions, and key quotes —
   grounded strictly in the transcript, no fabrication
4. **Render a self-contained interactive dashboard** (tabs, a checkable
   action-item list that remembers your checkmarks via `localStorage`,
   copy-to-clipboard) tailored to the meeting type
5. **Offer follow-ups** that hook into your other connectors (draft the recap
   email, create calendar holds for action items, etc.)

The dashboard is a single `.html` file with no external dependencies — it opens
offline by double-click.

## Prerequisites

- **Claude Code** (CLI, desktop, or web)
- The **Granola MCP connector** connected to your session. The skill relies on
  the `mcp__Granola__*` tools (`list_meetings`, `get_meeting_transcript`, etc.).
  See Granola's docs for connecting it to Claude.

## Install

### Option A — as a personal skill (simplest)

Clone this repo into your Claude Code skills directory:

```bash
git clone https://github.com/JaniCanCodeGo/meetingos.git ~/.claude/skills/meetingos
```

That's it. The skill is now available in every Claude Code session as
`meetingos`.

### Option B — per-project skill

Copy the skill into a specific project so it travels with that repo:

```bash
mkdir -p your-project/.claude/skills
git clone https://github.com/JaniCanCodeGo/meetingos.git your-project/.claude/skills/meetingos
```

## Usage

In any Claude Code session (with Granola connected), just ask:

- `run meetingos on my latest meeting`
- `use meetingos on my "Q3 Planning" meeting`
- `build a dashboard from my standup this morning`

Claude will pick the meeting (or ask which one), generate the dashboard into a
`meetingos_output/` folder in your current directory, and send you the file to
open.

## What's in this repo

```
meetingos/
├── SKILL.md                          # the skill definition + logic
├── reference/
│   └── dashboard_template.html       # the dashboard layout the skill builds from
├── README.md
└── LICENSE
```

## How it works

`SKILL.md` is the instruction set Claude loads when you invoke `meetingos`. It
tells Claude how to find the meeting, classify it, extract the structured
elements, and render the dashboard by adapting `reference/dashboard_template.html`
to the detected meeting type. Everything stays grounded in the real transcript —
empty categories are shown as "None captured" rather than invented.

## Credits

Inspired by the "Meeting OS" concept popularized by Matt Paige's *30 Days of
Claude Skills* series. This is an independent, open implementation built against
the public Granola MCP connector.

## License

MIT — see [LICENSE](LICENSE).
