# skill-seekers — audit verdict

**VERDICT: ✅ SAFE** (static analysis, 2026-06-23)

- No `eval`/`exec`/`os.system`; all `subprocess` calls use list-form args (no shell
  injection).
- No exfiltration. API keys (`ANTHROPIC`/`OPENAI`/`GOOGLE`/`VOYAGE`/`PINECONE`/
  `GITHUB_TOKEN`) are sent only to their own provider APIs. No reads of
  `~/.ssh`/`~/.aws`/cookies.
- No malicious install hooks. `pyproject.toml` uses standard console entry points; the
  one `curl | sh` is inside a printed help string (install `uv`), not auto-run.
- Only "secret" found: AWS's public docs placeholder `AKIAIOSFODNN7EXAMPLE`.
- No prompt injection in `SKILL.md`/`AGENTS.md`/`CLAUDE.md`.
- 44MB size is benign — mostly translated READMEs (11 languages) and a large CHANGELOG.

**Residual risks:** optional "enhance" disables nested-agent permission prompts; broad
scraper writes/deletes within its own temp/output dirs; auto-downloads config from a
third-party registry. See `SOURCE.md`.
