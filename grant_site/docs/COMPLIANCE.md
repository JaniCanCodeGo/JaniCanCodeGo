# Compliance Mapping — Grant Writing Studio

## Section 508 (36 C.F.R. Part 1194 / WCAG 2.0 AA baseline; WCAG 2.1 AA target)

| Requirement | WCAG SC | Implementation |
|---|---|---|
| Text alternatives | 1.1.1 | No informational images; any future images require alt text (also checked on reviewed org sites) |
| Info & relationships | 1.3.1 | Semantic landmarks, labeled form controls, `caption`/`scope` on tables |
| Use of color | 1.4.1 | Status conveyed with text ("Complete", "In progress"), not color alone |
| Contrast | 1.4.3 / 1.4.11 | Palette documented in styles.css; ≥4.5:1 text, ≥3:1 UI, light & dark |
| Resize / reflow | 1.4.4 / 1.4.10 | Relative units, responsive grid, no horizontal scroll at 320px |
| Keyboard | 2.1.1 / 2.1.2 | No pointer-only interactions; standard controls only |
| Bypass blocks | 2.4.1 | Skip link on every page |
| Page titles | 2.4.2 | Unique descriptive `<title>` per page |
| Focus order/visible | 2.4.3 / 2.4.7 | Managed focus to progress/results; 3px focus outline |
| Motion | 2.3.1 / 2.2.2 | No flashing/auto-motion; `prefers-reduced-motion` honored |
| Language | 3.1.1 | `<html lang="en">` |
| Labels & errors | 3.3.1 / 3.3.2 | `role="alert"` errors, per-field labels + hints |
| Status messages | 4.1.3 | `role="status"` live region for agent progress |
| Electronic content (E205) | — | Generated .docx/.md use true heading structure |

Deployment checklist: run axe-core or ANDI, keyboard-only pass, and one
screen reader (NVDA/JAWS/VoiceOver) before release; record results in the
agency's ACR/VPAT if applicable.

## California privacy

- **CalOPPA (B&P §22575)**: conspicuous policy link in footer; categories,
  third parties, change-notice process, and DNT response disclosed.
- **CCPA/CPRA (Civ. Code §1798.100+)**: rights to know/delete/correct
  honored (delete-run endpoint); no sale/sharing; non-discrimination;
  GPC honored by design (no tracking exists).
- **Online Eraser (B&P §22581)** and **SOPIPA (§22584)**: addressed on the
  children's privacy page.

## Federal privacy & children

- **COPPA**: site not directed to children; no knowing collection under 13;
  deletion on notice.
- **CIPA (47 U.S.C. §254(h))**: obligations bind E-Rate schools/libraries,
  not websites, but the site is designed for CIPA-compliant environments:
  no harmful-to-minors content, no chat/social/DM features, no collection
  from minors, no filter-circumvention tools. Institutions deploying the
  site remain responsible for their Internet safety policy, filtering,
  notice/hearing, and FCC certifications.
- **FERPA**: no student education records are collected; UI instructs users
  not to enter student PII.

## Regulated filings — honest boundaries

- **EIN**: IRS issues EINs only through its own channels; agent prepares a
  complete SS-4 worksheet and links the online EIN Assistant. SSNs/ITINs are
  never collected by this app.
- **Trademark**: USPTO TEAS requires applicant/attorney signature; agent
  performs knockout searching, class suggestion, fee calculation, and
  prepares the full TEAS Plus packet. Output states it is not legal advice.
