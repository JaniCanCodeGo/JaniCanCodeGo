# last30days-skill — audit verdict

**VERDICT: ✅ SAFE** (static analysis, 2026-06-23)

- No `eval`/`exec`/`pickle`/base64-decode-then-run; no `curl | bash`; no telemetry
  (no posthog/sentry/mixpanel).
- Each API key flows only to the service that owns it. Outbound URLs are scrubbed of
  `key`/`token`/`secret` params before logging (good hygiene).
- Browser-cookie access is hard-scoped to X (`auth_token`,`ct0`) and TruthSocial
  (`_session_id`) only, used locally, never POSTed out. Consent required in chat first.
- One benign SessionStart hook (`check-config.sh`) reads `.env` for status display
  without executing values. No malicious install hooks. Zero runtime dependencies.
- No hardcoded secrets. No deliberate prompt-injection payload.

**Residual risks:** reads logged-in session cookies (consent-gated); routes some
queries through third-party scrapers; feeds untrusted web text to the summarizer.
See `SOURCE.md`.
