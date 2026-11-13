"""Application readiness check: registrations and standard attachments.

Verifies what can be verified automatically (EIN, 501(c)(3) subsection from
IRS public records) and produces a have/need checklist for everything a
funder typically requires, with direct links to complete each item.
"""

STANDARD_ATTACHMENTS = [
    ("irs_determination", "IRS 501(c)(3) determination letter",
     "Required by nearly all foundations and most government funders. "
     "Apply via Form 1023/1023-EZ at pay.gov if not yet exempt."),
    ("board_list", "Board of directors list with affiliations",
     "Names, roles, and professional affiliations of all board members."),
    ("bylaws", "Articles of incorporation and bylaws",
     "Certified copies from your Secretary of State filing."),
    ("form_990", "Most recent IRS Form 990 / 990-EZ / 990-N",
     "Downloadable from your records or ProPublica Nonprofit Explorer."),
    ("financials", "Audited financial statements or board-approved budget",
     "Audit if revenue requires it; otherwise most recent board-approved "
     "financials."),
    ("org_budget", "Current organizational operating budget", ""),
    ("letters_support", "Letters of support / MOUs from partners",
     "Request 2-4 from schools, districts, or partner organizations."),
    ("org_chart", "Organizational chart", ""),
    ("staff_bios", "Key staff bios or resumes", ""),
    ("logic_model", "Logic model (inputs → activities → outputs → outcomes)",
     "One page; reviewers expect it for program grants."),
]

REGISTRATIONS = [
    ("sam_gov", "SAM.gov registration + UEI number",
     "REQUIRED for all federal grants. Free at https://sam.gov — allow "
     "up to several weeks; renew annually."),
    ("ca_registry", "California AG Registry of Charities (current RRF-1)",
     "Required for CA charities soliciting donations: "
     "https://oag.ca.gov/charities"),
    ("sos_good_standing", "Secretary of State good standing",
     "Check https://bizfileonline.sos.ca.gov/search/business"),
    ("ftb_status", "Franchise Tax Board exempt/good standing",
     "Check https://www.ftb.ca.gov/help/business/entity-status-letter.asp"),
]


def assess_readiness(ein_result, have_items=None):
    """have_items: list of checklist keys the user says they possess."""
    have = set(have_items or [])
    items = []

    ein_ok = bool(ein_result.get("found"))
    items.append({
        "key": "ein", "label": "EIN (federal tax ID)",
        "status": "ready" if ein_ok else "action_needed",
        "detail": f"EIN {ein_result.get('ein')}" if ein_ok else
        "No EIN — file the prepared SS-4 (see the SS-4 document).",
    })

    subsection = ""
    for m in ein_result.get("matches") or []:
        if m.get("ein") == ein_result.get("ein"):
            subsection = str(m.get("subsection") or "")
    c3_known = subsection == "3" or "irs_determination" in have
    items.append({
        "key": "501c3", "label": "501(c)(3) tax-exempt status",
        "status": "ready" if c3_known else "verify",
        "detail": ("Confirmed 501(c)(3) in IRS public records." if
                   subsection == "3" else
                   "Attested by applicant." if "irs_determination" in have else
                   "Not confirmed in IRS records — verify at "
                   "https://apps.irs.gov/app/eos/ or file Form 1023/1023-EZ."),
    })

    for key, label, detail in REGISTRATIONS:
        items.append({"key": key, "label": label,
                      "status": "ready" if key in have else "verify",
                      "detail": detail})
    for key, label, detail in STANDARD_ATTACHMENTS:
        if key == "irs_determination":
            continue  # covered by the 501(c)(3) item
        items.append({"key": key, "label": label,
                      "status": "ready" if key in have else "needed",
                      "detail": detail})

    ready = sum(1 for i in items if i["status"] == "ready")
    return {"items": items, "ready": ready, "total": len(items)}


def build_readiness_document(org_name, readiness):
    by_status = {"action_needed": [], "verify": [], "needed": [], "ready": []}
    for item in readiness["items"]:
        by_status.setdefault(item["status"], []).append(item)

    def fmt(items):
        return "\n\n".join(f"• {i['label']}" +
                           (f"\n  {i['detail']}" if i["detail"] else "")
                           for i in items) or "None."

    return {
        "title": f"{org_name} — Grant Application Readiness Checklist",
        "sections": [
            ("Summary",
             f"{readiness['ready']} of {readiness['total']} items ready. "
             "Most funders will not accept an application missing the "
             "items below — complete these before the deadline."),
            ("Action needed now", fmt(by_status["action_needed"])),
            ("Verify before applying", fmt(by_status["verify"])),
            ("Documents to gather", fmt(by_status["needed"])),
            ("Ready", fmt(by_status["ready"])),
        ],
    }
