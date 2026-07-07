# CDE Facilities Toolset

Civil Rights Review, Building Accessibility (CRR 20 / CRR 21).
California Department of Education, Office of Equal Opportunity,
Civil Rights and Education Equity Monitoring Unit.
Owner: Murjani McTier, Program Manager I.

Three connected tools sharing one legal framework: six accessibility
standards keyed to construction and alteration dates.

```
cde-facilities-toolset/
  references/     Shared source of truth (matrix, thresholds, boilerplate,
                  feedback letter, house style). All three tools read these.
  lof-agent/      Facilities LOF Agent: completed guide in, Missing
                  Information Report + CRR 20/21 Letter of Findings out.
  html-form/      Interactive browser version of the Facilities Review
                  Guide with Add/Remove instances and live standard and
                  compliance calculation. No network calls.
  guide/          Builder for the fillable Word guide
                  (Accessibility_Facilities_Checklist_v2.docx).
```

## Quick start

LOF agent (see `lof-agent/CLAUDE.md` for the full 6-step workflow):

```bash
cd lof-agent/scripts
npm install
node generate_lof.js sample_findings.json ../outputs
```

HTML form: open `html-form/facilities_review_form.html` in a browser.
Data autosaves to localStorage; use the export button for a JSON backup.

Word guide:

```bash
pip install python-docx
python3 guide/build_guide.py
```

## Non-negotiable rules

See `references/house-style.md`. Highlights: no em-dashes anywhere;
never invent a measurement; corrective actions always cite 2010 ADA while
violations cite the date-applicable standard; Program Access rows are
always None. / None.; quoted instrument language is verbatim; date
boundaries are re-verified against the current instrument every cycle.
