# Accessibility Standards Matrix (single source of truth)

All three tools (Guide, HTML form, LOF agent) read from this file. Do not
duplicate this logic elsewhere; link here.

> VERIFY EACH CYCLE: confirm the exact date boundaries against the current
> CRR instrument before carrying them forward. Do not silently reuse.

## Standard determination by construction / alteration date

| Standard | Date built or last altered | Can generate findings? |
|---|---|---|
| Program Access / Existing Facility (Section 504) | On or before June 3, 1977 | NO: observational only, no measurable citations |
| ANSI A117.1 (1961, R1971) | June 4, 1977 to January 17, 1991 | Yes |
| UFAS (1984) | January 18, 1991 to January 26, 1992 | Yes |
| 1991 ADA / ADAAG | January 27, 1992 to September 14, 2010 | Yes |
| 1991 ADA or 2010 ADA (subrecipient's choice; UFAS also permitted) | September 15, 2010 to March 14, 2012 | Yes |
| 2010 ADA | On or after March 15, 2012 | Yes |

Transition-window language (authoritative for the letter, from the CRR 20
boilerplate): UFAS or the 1991 ADA Standard apply to facilities constructed
or altered on or after January 27, 1992 and before September 15, 2010.
Between September 15, 2010 and March 14, 2012 a subrecipient may use UFAS,
1991 ADA, or 2010 ADA. The 2010 ADA Standards apply on or after
March 15, 2012.

## The two reviewer rules (settled; do not re-litigate)

### Rule 1: element-by-element alteration analysis

Per 2010 ADA Section 202.3 (U.S. Access Board scoping guidance) and
28 CFR Section 35.151(b): only those elements or spaces altered are required
to comply with the standard in effect at the time of alteration. If a room
or space is completely altered (or built new as part of an alteration), the
entire room or space is fully subject to that standard.

Examples:
- A 1968 restroom with grab bars replaced in 2018: grab bars are evaluated
  under 2010 ADA; everything else in the room remains Program Access.
- A 1985 building re-roofed in 2020: re-roofing is not an
  accessibility-affecting alteration; the whole building stays ANSI.
- Whole-room exception: a gutted and rebuilt room is entirely subject to
  the alteration-date standard.

### Rule 2: two columns, two rules, same row

| Column | Cite |
|---|---|
| Standard / Violation | The standard in effect when the element was built or last altered (Program Access / ANSI / UFAS / 1991 ADA / 2010 ADA) |
| Corrective Action | ALWAYS 2010 ADA, irrespective of construction or alteration date |

Confirmed by the Program Manager 2026-07-07: the Standard column is assessed
from the dates provided by the LEA and cites the applicable standard. If the
applicable standard has not been met, a corrective action is issued, and the
corrective action is always cited to 2010 ADA.

Consequence: some noted deficiencies will NOT require corrective action
because the condition already meets 2010 ADA (the corrective baseline).

### Program Access: None / None

Program Access has no measurable dimensional standard. For any element
under Program Access, both the Violation and Corrective Action columns read
exactly "None." with no exceptions. The LOF generator hard-suppresses
findings on Program Access rows even if findings are mistakenly present in
the input JSON. Program Access compliance is assessed observationally and
through the 14 Facilities / Maintenance and Operations staff interview
questions, not through measurement.

## Boundary conditions

- Measurements exactly at a threshold boundary are COMPLIANT.
- When sub-locations within one LOF row fall under different standards,
  list all applicable standards in the Standard column and issue findings
  only for sub-locations under a finding-eligible standard.
- Missing Year Built or ADA Modification Date means the standard cannot be
  determined: this is a completeness issue for the Missing Information
  Report, never a guessed standard and never an invented finding.
