# Clear Enough To Lead (CETL)

A workgroup teaching perimenopausal women with ADHD to build web projects and small Claude agents, paired with a podcast on leadership, life, and stress management. Founded and hosted by Jani.

## For the website build

Start with `data/site_manifest.json` — it's the entry point describing everything else in this repo and how to use it.

- `data/curriculum.json` — source of truth for the teaching track: module names, order, slugs, summaries, outcomes
- `data/podcast_episodes.json` — source of truth for the podcast episode list
- `content/curriculum/` — long-form lesson text, one markdown file per module slug
- `content/podcast/` — episode scripts and show notes
- `content/for_audience/` — direct messaging to workgroup members (not public site copy)
- `content/for_host/` — Jani's private prep material (not public, do not surface)
- `content/promo/` — marketing/outreach drafts

Re-fetch `site_manifest.json` before assuming the data shape — fields and tracks may be added over time.

## Generating new content

```bash
pip install -r requirements.txt
python3 cetl_content_agent.py curriculum --type lesson-plan --module html-basics
python3 cetl_content_agent.py podcast --type episode-outline --topic "the burnout math"
python3 cetl_content_agent.py audience --type welcome-message
python3 cetl_content_agent.py host --type teaching-script --module mini-claude-agents
python3 cetl_content_agent.py promo --type social-post --platform instagram --about "workgroup enrollment open"
```

## Business plan

`BUSINESS_PLAN.md` is the full researched business plan. `business_plan_agent.py` maintains it:

```bash
python3 business_plan_agent.py write   # gathers all project info + verified research into a plan-writing prompt
python3 business_plan_agent.py         # converts BUSINESS_PLAN.md into outputs/CETL_Business_Plan.docx (Word)
```

Same no-API-key pattern as the content agent: `write` builds a prompt (saved to `outputs/business_plan/`), you paste it into Claude, save the generated markdown over `BUSINESS_PLAN.md`, then rebuild the Word doc. Both steps are also MCP tools (`generate_business_plan`, `build_business_plan_docx`) in `cetl_mcp_server.py`.

No API key required — the CLI builds a prompt and saves it to `outputs/`; paste it into Claude to generate the content. The MCP server (`cetl_mcp_server.py`) does the same thing as callable tools inside Claude Desktop or claude.ai.

See `CLAUDE.md` for full architecture notes.
