# Facilities LOF Agent

CDE Office of Equal Opportunity, Civil Rights and Education Equity
Monitoring Unit. Civil Rights Review, Building Accessibility (CRR 20 /
CRR 21).

This agent takes a completed Facilities Review Guide (PDF or DOCX) and
produces (a) a Missing Information Report back to the site and (b) the
CRR 20/21 Letter of Findings (.docx).

## Source of truth (read before every run)

- `../references/standards-matrix.md` : date-to-standard matrix, the two
  reviewer rules, Program Access suppression, boundary conditions.
- `../references/thresholds.md` : compliance thresholds by standard.
- `../references/boilerplate.md` : verbatim letter language, table format,
  fixed 18-row order. Never paraphrase quoted instrument language.
- `../references/feedback-template.md` : courteous missing-information
  letter.
- `../references/house-style.md` : hard guardrails (no em-dashes, never
  invent measurements, attribution to Murjani McTier, verify legal dates
  each cycle, stop and ask on ambiguity).

## Workflow (run in order; NEVER skip the completeness check)

### STEP 0: Intake

Collect: school name, district, cycle (CRR 20 = 2025-26, CRR 21 =
2024-25), review date, path to the completed packet. The packet is the
completed "Civil Rights Review - Building Accessibility" document (or
the JSON export from its HTML form), optionally accompanied by the
completed "Program Access, Facilities and Maintenance & Operations Staff
Interview" document, which is issued and collected separately.

Read the packet:
- PDF: `pdftotext` first; if the output is empty or garbled the document
  is scanned, use OCR (`ocrmypdf` or `tesseract`) and re-extract.
- DOCX: `pandoc`, `python-docx`, or `mammoth`. Extract content-control
  (SDT) values as well as body text; the guide stores answers in SDTs.

### STEP 1: Completeness check (ALWAYS FIRST, before any analysis)

Flag every instance of:
- Blank measurement fields
- "Unknown" entries
- Unanswered Yes/No questions
- Missing Year Built or ADA Modification Date (these are required to
  determine the applicable standard)
- Contradictions between entries
- Alterations gate inconsistencies: each section asks "Have any
  alterations or ADA modifications been made since original
  construction?" A "Yes" with an empty Alterations Log, a "No" or blank
  gate with populated log rows, or a blank gate altogether are all
  completeness flags. The Alterations Log (Element Altered / Date
  Altered) is the ONLY place modification data lives; there are no
  separate narrative modification fields.
- Missing signature block
- Skipped conditional follow-ups (for example: "Ramp leading to
  entrance? Yes" but no matching Ramps instance)

Output BOTH:
1. `Missing_Information_Report_[School].md`
2. `Feedback_to_Site_[School].docx` built from
   `../references/feedback-template.md`

Tone: courteous, never accusatory. Missing data is a completeness issue,
not a finding. No em-dashes.

### STEP 2: Determine the applicable standard per area

THE AGENT OWNS THIS STEP ENTIRELY. The site-facing documents (the
"Civil Rights Review - Building Accessibility" guide and its HTML form)
intentionally contain NO standards content: no era dropdowns, no
date-to-standard tables, no reviewer rules. The LEA records only raw
data (Year Built, ADA Modification Dates, the two-column Alterations Log
of Element Altered / Date Altered, and measurements). Determining the
standard is internal work this agent performs for the Program Reviewer.

Use the matrix in `../references/standards-matrix.md`, applied element by
element: an area's base standard comes from Year Built; each entry in the
Alterations Log moves ONLY that altered element to the alteration-date
standard (whole-room exception: a fully rebuilt space takes the
alteration-date standard entirely). If dates are ambiguous, STOP AND ASK;
never guess a standard.

Output `Standard_Determination_[School].md` for the Program Reviewer:
one entry per area and per altered element, showing the date used, the
standard selected, the matrix row that selected it, and a one-sentence
plain-language explanation (for example: "Building 100 Boys Restroom,
built 1968: Program Access, observational review only. Grab bars replaced
2018: evaluated under 2010 ADA per the element-by-element rule; only the
grab bars take the newer standard."). The Program Reviewer verifies this
document; the agent explains, the reviewer decides.

### STEP 3: Determine whether findings can be issued

- Program Access areas: both the Violation and Corrective Actions columns
  read exactly `None.` This is a HARD SUPPRESSION applied at generation
  time, even if findings are mistakenly present in the findings JSON.
- All other standards can generate findings.
- The Standard column and each Violation cite the APPLICABLE standard for
  that element (the one from Step 2, not always 2010 ADA).
- Every Required Corrective Action is ALWAYS cited to 2010 ADA,
  irrespective of construction or alteration date. Consequence: a
  deficiency under an older standard that already meets 2010 ADA requires
  NO corrective action; note it accordingly.
- Measurements exactly at a threshold boundary are compliant.
- When sub-locations in one row fall under different standards, list all
  in the Standard column and only issue findings for the sub-locations
  under a finding-eligible standard.
- NEVER invent a measurement. Anything missing belongs in the Step 1
  report, not the letter.

### STEP 4: Generate the LOF (.docx)

Build the findings JSON (schema below), then run:

```bash
cd scripts && npm install && node generate_lof.js <findings.json> <output_dir>
```

The generator enforces the fixed 18-row order, the boilerplate, the table
format from `../references/boilerplate.md`, and Program Access
suppression. Filename: `[SchoolName]_Facilities_LOF.docx`. VALIDATE the
file opens before delivering (the script re-reads the zip; also open it
with python-docx or confirm Word/LibreOffice can load it).

### STEP 5: Deliverables summary

Present: the Missing Information Report, the Feedback docx, the Standard
Determination report, the LOF docx, and a console summary: areas
reviewed, findings issued, Program Access areas, missing items count.

## Findings JSON schema

```json
{
  "school": "Example High School",
  "district": "Example Unified School District",
  "crr": 20,
  "reviewDate": "2026-04-15",
  "rows": [
    {
      "area": "Accessible Parking",
      "locations": [
        { "name": "Lot A (Main)", "constructed": "1968", "modified": "2018 (restriping)" }
      ],
      "standards": ["Program Access", "2010 ADA"],
      "violations": [
        { "text": "Access aisle measures 48 inches in width; minimum 60 inches required.", "cite": "2010 ADA Section 502.3" }
      ],
      "correctives": [
        { "text": "Restripe the access aisle to a minimum width of 60 inches.", "cite": "2010 ADA Section 502.3" }
      ]
    }
  ]
}
```

Rules the generator enforces (do not rely on the JSON being right):
- Rows are emitted in the fixed 18-row order regardless of input order;
  missing rows are emitted with `None.` / `None.`
- A row whose standards are ONLY Program Access is forced to `None.` /
  `None.` even if violations are present in the JSON.
- Corrective cites must reference 2010 ADA; the generator warns on any
  corrective cite that does not.
- Em-dash characters are rejected anywhere in the JSON text.
