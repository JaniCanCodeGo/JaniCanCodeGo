# Security Audit

**Date:** 2026-06-23
**Method:** Static analysis only. No code from any repo was executed, installed, or
built. Each skill was reviewed by an independent auditor for: code execution &
obfuscation, data exfiltration, credential/secret access, filesystem danger,
install-time hooks, prompt injection, and hardcoded secrets.

**Result: all six skills reviewed SAFE.** No malware, no data theft, no hidden
instructions, no malicious install hooks, and no real committed secrets were found.

## Summary

| Skill | Verdict | Code | Key caveats (normal trade-offs, not foul play) |
|---|---|---|---|
| stop-slop | ✅ SAFE | none | None. Pure writing-style instructions. |
| karpathy-llm-wiki | ✅ SAFE | none | None. Local markdown read/write only. |
| understand-anything | ✅ SAFE | yes | Generates & runs analysis scripts by design; localhost dashboard, no auth (loopback only); loads Google Fonts. |
| ui-ux-pro-max-skill | ✅ SAFE | yes | CLI runs shell on local paths; optional Gemini image-gen sends prompts to Google if you add a key; uses `npx` (normal supply-chain note). |
| skill-seekers | ✅ SAFE | yes | Optional "enhance" runs a nested AI with permission prompts disabled; broad scraper; auto-downloads config JSON from the project's own registry. |
| last30days-skill | ✅ SAFE | yes | Can read X/TruthSocial login cookies (asks first); some data routes through 3rd-party scrapers (ScrapeCreators, Xquik); API keys go only to their matching service. |

## Notable findings (good signs)
- **No code execution primitives** (`eval`/`exec`/`os.system`/`shell=True`) on
  untrusted input in any first-party source.
- **No exfiltration:** no telemetry/analytics endpoints; API keys are only ever sent
  to the service that issued them.
- **Good secret hygiene:** `last30days` scrubs key/token query params before logging;
  the only "secret" found anywhere was AWS's public documentation placeholder
  (`AKIAIOSFODNN7EXAMPLE`).
- **No malicious install hooks:** no `preinstall`/`postinstall` surprises; the one
  `curl | sh` found was inside a printed help string, not auto-executed.
- **No prompt injection:** no hidden/invisible-unicode instructions, no role-resets,
  no "ignore previous instructions" in any agent-facing markdown.

## Design-level note (the one real thing to watch)
`skill-seekers` and `last30days` feed untrusted web/social content into the AI without
an explicit "treat this as data, do not follow embedded instructions" guard. A
malicious page could attempt to steer a generated summary. This is an inherent risk
of any web-ingesting tool, not a defect in these repos. Mitigation: run them in a
sandbox, and verify their output before acting on it.
