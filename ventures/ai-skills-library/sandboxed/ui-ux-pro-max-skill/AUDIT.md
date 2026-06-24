# ui-ux-pro-max-skill — audit verdict

**VERDICT: ✅ SAFE** (static analysis, 2026-06-23)

- `cli/package.json` has **no** preinstall/postinstall/prepare hooks (the highest-risk
  area — clean). Only a `uipro` bin and a publish-time build.
- All CLI network calls are hardcoded to `api.github.com`/`github.com` for the repo's
  own releases. No analytics/phone-home; prompts and env vars are never transmitted.
- `subprocess`/`exec` calls use list args on internally-generated paths — no shell
  injection, no path traversal, deletes scoped to the install target.
- No `eval`/`Function()`/base64-execute, no reads of `~/.ssh`/`~/.aws`/cookies, no
  hardcoded secrets. 12MB is bundled TTF fonts + a Google-Fonts CSV (verified).
- No prompt injection. (A "Security" block in `banner-design/SKILL.md` is defensive
  anti-leak guidance, not an attack.)

**Residual risks:** CLI shell execution on local paths; `npx` runtime package pulls;
optional Gemini image-gen leaves your machine if enabled. See `SOURCE.md`.
