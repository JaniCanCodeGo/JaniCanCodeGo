"""Funder-format project budget: line items, budget justification, and a
request-justification (funding gap) statement.

If the user supplies line items they are used exactly; otherwise the grant
request is allocated across standard categories with printed percentages so
every number remains traceable. Personnel/fringe/travel/equipment/supplies/
contractual/other/indirect follows the SF-424A federal category layout that
most state and foundation budgets mirror.
"""

DEFAULT_ALLOCATION = [
    ("personnel", 0.55), ("travel", 0.03), ("equipment", 0.05),
    ("supplies", 0.10), ("contractual", 0.12), ("other", 0.05),
]
DEFAULT_FRINGE_PCT = 22.0
DEFAULT_INDIRECT_PCT = 10.0  # de minimis federal indirect rate

CATEGORY_JUSTIFICATIONS = {
    "personnel": "Salaries for program staff delivering direct services "
                 "(program lead and instructors), prorated to project time.",
    "fringe": "Fringe benefits (payroll taxes, health, retirement) at the "
              "stated rate applied to project personnel.",
    "travel": "Local mileage for program delivery sites and one program-"
              "related conference/training.",
    "equipment": "Durable items required for delivery (e.g., laptops, "
                 "presentation equipment) retained for program use.",
    "supplies": "Consumable program materials, curriculum, and participant "
                "supplies.",
    "contractual": "Specialized contracted services (e.g., evaluators, "
                   "guest instructors) procured per written agreements.",
    "other": "Facility use, insurance, printing, and communication costs "
             "directly attributable to the project.",
    "indirect": "Indirect costs at the stated rate (de minimis or "
                "negotiated rate) covering administration and facilities.",
}


def build_budget(grant_target, items=None, fringe_pct=None,
                 indirect_pct=None, match_amount=0.0):
    """Build a line-item budget totaling the grant request.

    items: optional dict of category -> amount (personnel, travel, equipment,
    supplies, contractual, other). Fringe and indirect are computed.
    """
    fringe_pct = DEFAULT_FRINGE_PCT if fringe_pct in (None, "") else float(fringe_pct)
    indirect_pct = (DEFAULT_INDIRECT_PCT if indirect_pct in (None, "")
                    else float(indirect_pct))
    match_amount = float(match_amount or 0)

    lines = []
    if items and any(v for v in items.values()):
        direct = {k: float(v or 0) for k, v in items.items()}
        allocated = False
    else:
        # Allocate the request across default percentages, solving so that
        # direct + fringe + indirect == grant_target.
        # total = D + fringe(personnel) + indirect*(D+fringe)
        # with D distributed by DEFAULT_ALLOCATION shares of D.
        # total = direct + fringe + indirect
        #       = D*(S + p*fringe%) * (1 + indirect%), where S is the sum of
        # the allocation shares and D the allocation base — solve for D.
        p_share = dict(DEFAULT_ALLOCATION)["personnel"]
        share_sum = sum(s for _, s in DEFAULT_ALLOCATION)
        factor = ((share_sum + fringe_pct / 100 * p_share) *
                  (1 + indirect_pct / 100))
        direct_total = grant_target / factor
        direct = {cat: round(direct_total * share, 2)
                  for cat, share in DEFAULT_ALLOCATION}
        allocated = True

    fringe = round(direct.get("personnel", 0) * fringe_pct / 100, 2)
    direct_plus_fringe = round(sum(direct.values()) + fringe, 2)
    indirect = round(direct_plus_fringe * indirect_pct / 100, 2)
    total = round(direct_plus_fringe + indirect, 2)

    for cat in ("personnel", "travel", "equipment", "supplies",
                "contractual", "other"):
        if direct.get(cat):
            lines.append({"category": cat.title(), "amount": direct[cat],
                          "justification": CATEGORY_JUSTIFICATIONS[cat]})
    lines.insert(1 if direct.get("personnel") else 0,
                 {"category": f"Fringe benefits ({fringe_pct:.0f}%)",
                  "amount": fringe,
                  "justification": CATEGORY_JUSTIFICATIONS["fringe"]})
    lines.append({"category": f"Indirect ({indirect_pct:.0f}%)",
                  "amount": indirect,
                  "justification": CATEGORY_JUSTIFICATIONS["indirect"]})

    return {
        "grant_target": grant_target, "lines": lines, "total": total,
        "match_amount": match_amount,
        "variance_from_request": round(total - grant_target, 2),
        "allocated_by_default_pcts": allocated,
        "note": ("Amounts allocated from the request using standard "
                 "percentages — replace with your actual figures before "
                 "submission." if allocated else
                 "Amounts as provided by the applicant."),
    }


def request_justification(budget, forecast):
    """Explain why the grant is needed (funding-gap framing).

    Uses the Year-1 forecast WITHOUT the grant to show the gap the award
    fills — this answers the reviewer question 'why do you need this money?'
    """
    y1 = forecast["years"][0]
    without_grant_net = y1["net"] - y1["grant_revenue"]
    if without_grant_net < 0:
        gap_text = (f"Without this award, Year-1 operations show a deficit "
                    f"of ${abs(without_grant_net):,.0f}; the requested "
                    f"${budget['grant_target']:,.0f} closes that gap and "
                    "funds the expansion described in the project plan.")
    else:
        gap_text = (f"Existing revenue sustains current operations "
                    f"(Year-1 net of ${without_grant_net:,.0f} before this "
                    "award); the requested funds are restricted to the NEW "
                    "activities in this proposal — expanded enrollment, "
                    "additional staff hours, and materials — none of which "
                    "occur without the award. Grant funds supplement and do "
                    "not supplant existing program spending.")
    if budget["match_amount"]:
        gap_text += (f" The organization commits "
                     f"${budget['match_amount']:,.0f} in match/in-kind.")
    return gap_text


def build_budget_document(org_name, budget, forecast):
    rows = [("Category", "Amount", "Justification")]
    for line in budget["lines"]:
        rows.append((line["category"], f"${line['amount']:,.2f}",
                     line["justification"]))
    rows.append(("TOTAL PROJECT BUDGET", f"${budget['total']:,.2f}", ""))
    if budget["match_amount"]:
        rows.append(("Applicant match / in-kind",
                     f"${budget['match_amount']:,.2f}",
                     "Committed by the applicant; not part of the request."))
    widths = [max(len(r[i]) for r in rows) for i in range(3)]
    table = "\n".join("  ".join(c.ljust(widths[i])
                                for i, c in enumerate(row)) for row in rows)
    sections = [
        ("Line-Item Project Budget", table + "\n\n" + budget["note"]),
        ("Why This Funding Is Needed",
         request_justification(budget, forecast)),
        ("Budget Narrative",
         "\n\n".join(f"{l['category']}: ${l['amount']:,.2f} — "
                     f"{l['justification']}" for l in budget["lines"])),
    ]
    if abs(budget["variance_from_request"]) > 1:
        sections.insert(1, (
            "Reconciliation",
            f"Line items total ${budget['total']:,.2f} vs the requested "
            f"${budget['grant_target']:,.2f} (difference "
            f"${budget['variance_from_request']:,.2f}). Adjust line items "
            "or the request so they match before submission."))
    return {"title": f"{org_name} — Project Budget & Justification",
            "sections": sections}
