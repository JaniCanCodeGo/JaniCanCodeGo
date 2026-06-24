# understand-anything — source & safe use

- **Upstream:** https://github.com/GoodDingo/understand-anything
  (originally `Lum1104/Understand-Anything`)
- **Retrieved:** 2026-06-23 (default branch `main`)
- **License:** MIT
- **What it does:** Scans a codebase with a multi-agent pipeline, builds a JSON
  knowledge graph, and serves an interactive localhost dashboard to explore it.

## Good for (my ventures)
Understanding a codebase you didn't write — including your own CRR agent — in plain
English, without reading the Python. Also handy for any future website/scraping code.

## How to run it SAFELY
Run inside a disposable sandbox, pointed only at code you trust.

```bash
# inside a sandbox only, in the project you want to map
/understand            # in Claude Code, builds .understand-anything/knowledge-graph.json
npx vite --open        # opens the localhost-only dashboard
```

## Caveats (from the audit)
- By design it **generates and runs** Node/bash analysis scripts (in `/tmp`) and runs
  `git` — real code execution, so only use on repos you trust.
- The dashboard has no auth but is loopback-only (your machine only).
- It loads Google Fonts from the internet (minor cosmetic network call).
