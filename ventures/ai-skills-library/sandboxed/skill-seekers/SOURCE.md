# skill-seekers — source & safe use

- **Upstream:** https://github.com/yusufkaraaslan/Skill_Seekers
- **Retrieved:** 2026-06-23 (default branch `main`)
- **License:** MIT
- **What it does:** Scrapes documentation sites, GitHub repos, PDFs, videos, and ~15
  other source types and converts them into reusable AI "skills" / RAG knowledge,
  exposed via a ~40-tool MCP server.

## Good for (my ventures)
Building a custom expert: feed it grant-writing guides, funder guidelines, or ADHD
research, and Claude becomes a specialist in that material.

## How to run it SAFELY
Run inside a disposable sandbox (Claude Code on the web, a VM, or a container) — not
wired into your daily machine.

```bash
# inside a sandbox only
git clone https://github.com/yusufkaraaslan/Skill_Seekers
# follow its README; provide API keys via .env, never in chat
```

## Caveats (from the audit)
- The optional **"enhance"** feature spawns a nested AI agent with
  `--dangerously-skip-permissions`. Leave it OFF unless you fully trust the scraped
  source.
- It will fetch any URL you give it and can auto-download config JSON from the
  project's own registry (`api.skillseekersweb.com`).
- Only scrape sources you trust; treat generated skills as draft material.
