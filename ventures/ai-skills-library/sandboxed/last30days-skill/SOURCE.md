# last30days-skill — source & safe use

- **Upstream:** https://github.com/mvanhorn/last30days-skill
- **Retrieved:** 2026-06-23 (default branch `main`)
- **License:** MIT
- **What it does:** Given a topic, fans out searches across Reddit, X, YouTube, TikTok,
  Instagram, Threads, Hacker News, Polymarket, GitHub, Bluesky, TruthSocial, and web
  search; ranks by engagement; and writes a cited HTML brief.

## Good for (my ventures)
Grant-opportunity scans, business-idea trend-spotting, and seeing what the ADHD
community is actually asking about AI right now.

## How to run it SAFELY
Run inside a disposable sandbox. Reddit/HN/Polymarket/YouTube work with no keys.

```bash
# inside a sandbox only
git clone https://github.com/mvanhorn/last30days-skill
# optional API keys go in .env, never in chat
```

## Caveats (from the audit)
- It can **read your X / TruthSocial login cookies** (macOS Keychain prompt) to reach
  logged-in sources. It asks first — **decline if unsure**; keyless mode still works.
- Two data sources are **third-party scrapers** (ScrapeCreators, Xquik). If you sign up
  for their keys, you're trusting those companies with your queries.
- It ingests untrusted web content, so a malicious post could try to steer the summary.
  Verify anything important.
