# CLAUDE.md

## What this project does

Clear Enough To Lead (CETL) is an AI-powered content generation agent for a middle-aged Black American woman managing a team while navigating ADHD and perimenopause. Voice: humorous, sarcastic, real — not clinical, not corporate, not a wellness brochure.

## Usage

### CLI
```bash
# Content directed at others (team, boss, HR)
python3 cetl_agent.py for-others --type email --situation "team missed the deadline again"
python3 cetl_agent.py for-others --type feedback --situation "employee keeps talking over people in meetings"
python3 cetl_agent.py for-others --type meeting-agenda --topic "Q3 retro" --duration 45
python3 cetl_agent.py for-others --type boundary --situation "manager keeps scheduling over my focus blocks"

# Content for yourself
python3 cetl_agent.py for-self --type brain-dump --input "I have 47 things to do and I forgot what all of them are"
python3 cetl_agent.py for-self --type fog-recovery
python3 cetl_agent.py for-self --type morning-check --level low --symptoms "brain fog, hot flash at 3am"
python3 cetl_agent.py for-self --type self-advocacy --situation "I need a standing desk and flexible start time"
python3 cetl_agent.py for-self --type energy-plan --level medium --symptoms "mild fog" --commitments "10am standup, 2pm 1:1s"
```

### MCP Server (Claude Desktop)
```bash
python3 cetl_mcp_server.py   # runs on stdio transport
```
Use `setup_claude_desktop.bat` on Windows to configure Claude Desktop automatically.

### Install dependencies
```bash
pip install -r requirements.txt
```

### API key
Set `ANTHROPIC_API_KEY` as an environment variable before running.

## Architecture

All core logic lives in `cetl_agent.py`. `cetl_mcp_server.py` is a thin wrapper that exposes two MCP tools:
- `generate_for_others(content_type, situation, context, topic, duration)`
- `generate_for_self(content_type, input_text, situation, context, level, symptoms, commitments)`

## Content types

### For Others (directed at team, boss, HR)
| Type | What it generates |
|------|------------------|
| `email` | Workplace email for any situation |
| `feedback` | Script for a feedback conversation with an employee |
| `meeting-agenda` | ADHD-friendly agenda with time boxes and parking lot |
| `boundary` | Boundary-setting script — verbal and written versions |

### For Self (personal use)
| Type | What it generates |
|------|------------------|
| `brain-dump` | Converts chaotic thoughts into an organized action list |
| `fog-recovery` | Step-by-step protocol for when you've lost the thread |
| `morning-check` | Realistic morning checklist based on current symptom level |
| `self-advocacy` | Script for requesting accommodations from boss or HR |
| `energy-plan` | Realistic work plan based on current energy level |

## Output files

Generated content is saved to `outputs/for_others/` and `outputs/for_self/` (gitignored), named by type and timestamp.

## The voice

The persona is baked into the system prompt in `cetl_agent.py` under `PERSONA`. It governs every piece of content generated. If the tone needs adjustment, that's the place to change it.
