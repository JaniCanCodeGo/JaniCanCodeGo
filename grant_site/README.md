# Grant Writing Studio

An agent-driven grant-writing website. One run takes an organization from
website review to a complete, downloadable grant package:

1. **Website review** — fetches the org's site, extracts mission language,
   programs, contacts, and any published EIN, and runs a heuristic
   Section 508 / WCAG 2.1 AA spot-check of the site itself.
2. **EIN lookup** — searches IRS public records (via the free ProPublica
   Nonprofit Explorer API). If no EIN exists, the agent prepares a completed
   **IRS Form SS-4 worksheet** plus a direct link to the IRS online EIN
   Assistant. (The IRS has no third-party filing API; a responsible party
   must submit it — takes minutes and is free.)
3. **Business plan & prospectus** — full plan with **documented, sourced
   pain points** for the org's focus areas, a deterministic **3-year
   forecast of success, earnings, participants, and merchandise** (every
   number traces to printed assumptions), and a sustainability story.
4. **Trademark** — knockout search for identical/sound-alike marks
   (USPTO APIs when keys are configured, always with manual deep links to
   USPTO/California SOS search), Nice class suggestions, and a complete
   **TEAS Plus application packet** with fee calculation, specimen
   checklist, and timeline. Filing is signed at teas.uspto.gov (the USPTO
   has no third-party filing API). Not legal advice.
5. **Grant narrative** — need statement, goals, evaluation plan, budget
   narrative, capacity, sustainability.

Deliverables are written as Markdown and Word (.docx) per run and are
downloadable from the results page.

## Run it

```bash
cd grant_site
pip install -r requirements.txt
python3 -m uvicorn server:app --port 8000
# open http://127.0.0.1:8000/
```

Optional API keys (environment variables):

- `USPTO_TSDR_API_KEY` — free key from developer.uspto.gov for serial-number
  status checks.
- `USPTO_TM_API_KEY` — RapidAPI key for keyword trademark search. Without it
  the agent still generates variants and manual search links.

## API

- `POST /api/runs` — start a run (JSON body; see `RunRequest` in server.py)
- `GET /api/runs/{id}` — progress + results
- `GET /api/runs/{id}/files/{name}` — download a generated document
- `DELETE /api/runs/{id}` — delete all data for a run (CCPA right-to-delete)
- `GET /api/ein/lookup?name=&state=` — standalone EIN search
- `GET /api/trademark/search?q=&goods=` — standalone trademark search

## Compliance

- **Section 508 / WCAG 2.1 AA** — see `static/accessibility.html` and
  `docs/COMPLIANCE.md` for the criterion-by-criterion mapping.
- **Privacy (CalOPPA, CCPA/CPRA)** — `static/privacy.html`; no cookies, no
  trackers, no sale/sharing, delete-run endpoint.
- **COPPA / California minors / CIPA** — `static/childrens-privacy.html`;
  the site is filter-friendly and safe for CIPA-covered school/library
  networks.

## What the agent does NOT do (by law, not by limitation)

- It does not submit EIN applications to the IRS (no API exists; the
  responsible party applies online with the prepared SS-4 worksheet).
- It does not e-file trademark applications (the USPTO requires signature
  at teas.uspto.gov; the agent prepares the complete packet).
- It does not provide legal, tax, or accounting advice.

## Tests

```bash
cd grant_site && python3 -m pytest tests/ -q
```
