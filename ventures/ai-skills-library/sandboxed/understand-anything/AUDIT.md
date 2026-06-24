# understand-anything — audit verdict

**VERDICT: ✅ SAFE** (static analysis, 2026-06-23)

- No telemetry, no exfiltration. Dashboard `fetch()` hits only the local dev server
  (`/knowledge-graph.json`, `/diff-overlay.json`); reads/writes confined to
  `<project>/.understand-anything/`.
- Only OS-command site uses `execFileSync('git', [...])` with array args (no shell).
  Dev-server middleware serves two fixed files — no path traversal.
- Local Vite server binds to **localhost** (no `0.0.0.0`), port 5173, loopback-only.
- `package.json` `prepare` runs a local `tsc` build — no network fetch, no
  `curl | bash`. No secrets, no prompt injection.
- Only outbound calls are Google Fonts (cosmetic) and an inert attribution link.

**Residual risks:** sub-agents legitimately write and execute generated Node/bash
scripts and read your source files — run only on trusted code. See `SOURCE.md`.
