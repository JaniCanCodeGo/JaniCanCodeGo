# ✦ Rez Wiz — California State Resume & SOQ Builder

A single-file, browser-only web app for writing resumes and Statements of
Qualifications (SOQs) for State of California job applications.

**Run it:** open `index.html` in any modern browser — no build step, no server,
no dependencies. (Or serve the folder: `python3 -m http.server` and visit
`http://localhost:8000/`.)

## Features

- **Resume Builder** — contact, summary, repeatable work-experience and
  education entries, skills; live preview; auto-saves to browser localStorage.
  First visit loads a fully populated sample profile (Alex Sacramento, Staff
  Services Manager I) so every page renders with content; use
  "Clear everything & start fresh" to wipe it or "Load sample resume" to
  bring it back.
- **SOQ Writer** — paste each factor exactly as written in the CalCareers
  posting, answer in STAR format, and export a correctly formatted SOQ
  (Arial, name + JC number header, factors restated in order).
- **Jobs That Might Interest You** — typing keywords such as `manager`,
  `supervisor`, `analyst`, `IT`, `accountant`, etc. anywhere in the resume
  surfaces matching State of California classifications with one-click
  searches on [CalCareers](https://www.calcareers.ca.gov/), the State's
  official jobs website.
- **Four output templates** — Capitol Classic, Golden State Modern,
  Sacramento Slate, and an ATS-safe State Standard. Export via Print → Save
  as PDF, or download a Word-compatible `.doc` file.
- **Privacy Policy & Terms** — built-in pages; the policy is written to
  CCPA/CPRA requirements and addresses children's protections
  (COPPA / CIPA-friendly: no data collection, no chat, no trackers).
- **Design** — client color palette (maroon `#8A2432`, crimson `#D2372A`,
  amber `#E68F3C`, cream `#F2ECDF`, slate `#26333D`), Impact display
  typography, animated brand stripe / floating-orb hero / sparkle motion
  (all disabled under `prefers-reduced-motion`), light + dark themes.

## Is an API key needed?

**No — the app as shipped needs no API key.** CalCareers (jobs.ca.gov /
calcareers.ca.gov) does **not** offer a public REST API or issue API keys.
Rez Wiz therefore matches keywords against a curated list of common state
classifications and deep-links into CalCareers' own search
(`…/JobSearchResults.aspx#kw=<keyword>`), which requires no credentials.

If you later want **live** job postings inside the app you would need one of:

1. **A server-side proxy that scrapes CalCareers search results** — no key,
   but fragile and subject to the site's terms of use.
2. **data.ca.gov open-data (CKAN) API** — free, no key required, for any
   CalHR/vacancy datasets published there (coverage varies).
3. **A commercial jobs API** (e.g. an aggregator that indexes government
   postings) — these do require paid API keys.

None of these are required for the current keyword-match + deep-link design.

## Files

- `index.html` — the entire app (HTML + CSS + JS, self-contained).
- All user data lives in the browser's `localStorage` under the key
  `rezwiz.v1`; nothing is transmitted anywhere.

## Disclaimer

Rez Wiz is an independent tool and is not affiliated with, sponsored by, or
endorsed by the State of California. The individual CalCareers job posting —
its duty statement, filing instructions, and SOQ requirements — is always
the authoritative source.
