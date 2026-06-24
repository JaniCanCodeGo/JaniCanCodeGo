# AI Skills Library

Vetted AI/Claude skills for my **personal ventures** — grant research & writing,
ADHD + AI coaching, the podcast, web-scraping/website projects, and business ideas.

This library is deliberately kept **separate from my day-job CRR/EE tooling**. Nothing
here is about the Civil Rights Review work.

Every skill in this repo was security-reviewed before being added (static analysis
only — see [`SECURITY-AUDIT.md`](SECURITY-AUDIT.md)). All six reviewed SAFE.

## How this is organized

### `active/` — safe to use anywhere
These are **text-only** skills (plain instructions). They cannot run code, read files,
or reach the network, so there is nothing to sandbox. Drop them into any Claude Code
`.claude/skills/` folder, or your global `~/.claude/skills/`, and use them freely.

| Skill | What it does | Good for |
|---|---|---|
| [stop-slop](active/stop-slop) | Removes the "AI voice" from writing | Grants, podcast scripts, business posts, ADHD content |
| [karpathy-llm-wiki](active/karpathy-llm-wiki) | Builds a self-maintaining knowledge base | A "second brain" across all ventures |

### `sandboxed/` — run only inside an isolated environment
These four skills **execute code**. They are SAFE, but because they run programs,
scrape the web, or can touch login cookies/API keys, the rule is: **run them only in
a throwaway sandbox** (e.g. a Claude Code on the web session, a VM, or a container),
**never wired to auto-run on my main machine.**

Each folder holds a vetted *card*, not the full source:
- `SKILL.md` — the actual skill instructions (safe to read)
- `AUDIT.md` — the security verdict and caveats
- `SOURCE.md` — where it came from and how to fetch + run it safely

| Skill | What it does | Why sandboxed |
|---|---|---|
| [skill-seekers](sandboxed/skill-seekers) | Turns docs/sites/videos into reusable skills | Scraper; optional "enhance" disables safety prompts |
| [last30days-skill](sandboxed/last30days-skill) | Recency research across social platforms | Can read browser login cookies; uses 3rd-party scrapers |
| [ui-ux-pro-max-skill](sandboxed/ui-ux-pro-max-skill) | Generates website/UI design systems | Ships a CLI; optional Gemini image-gen |
| [understand-anything](sandboxed/understand-anything) | Maps a codebase into a knowledge graph | Generates and runs analysis scripts |

## Golden rules for the sandboxed skills
1. Run them in an isolated/disposable environment, not on your daily machine.
2. Only point a scraper or analyzer at sources you trust.
3. Treat their output as research to verify — a malicious webpage could try to
   manipulate a summary.
4. Leave "permission-skipping" / auto-enhance features OFF unless you fully trust
   the input.
5. Keep your own API keys in `.env` files, never pasted into chats or committed.

---
*Skills here are third-party, MIT-licensed community projects. They are not official
Anthropic skills. Vetted on 2026-06-23.*
