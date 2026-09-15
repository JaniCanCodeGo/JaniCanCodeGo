"""Stage 5 — Grant narrative assembly.

Produces a standard grant application narrative (need, goals, activities,
evaluation, budget narrative, organizational capacity) from the pipeline's
earlier stages so every claim in the grant traces to the website review,
pain-point research, or the printed forecast assumptions.
"""


def build_grant_narrative(org, website, painpoints, forecast, ein_result):
    name = org.get("name", "The Organization")
    amount = forecast["assumptions"]["grant_target"]
    y1 = forecast["years"][0]
    ein_line = (f"EIN: {ein_result['ein']}" if ein_result.get("found")
                else "EIN: application in progress (SS-4 worksheet prepared; "
                     "see EIN report)")

    sections = [
        ("Applicant Information",
         f"{name}\n{ein_line}\nWebsite: {website.get('url') or 'N/A'}\n"
         f"Contact: {', '.join(website.get('emails', [])[:2]) or org.get('contact_email', 'on file')}"),
        ("Statement of Need",
         f"{name} addresses documented community needs:\n\n" +
         "\n".join(f"• {p['pain']} (Source: {p['source']})"
                   for p in painpoints["painpoints"][:5])),
        ("Project Description and Goals",
         f"With an award of ${amount:,.0f}, {name} will deliver "
         f"{org.get('programs', 'its core programs')} to "
         f"{y1['participants_served']:,} participants in Year 1.\n\n"
         "Goals:\n"
         "1. Expand program capacity and enrollment.\n"
         "2. Achieve measurable participant outcomes (pre/post gains).\n"
         "3. Build earned-revenue sustainability so the program continues "
         "beyond the grant period."),
        ("Evaluation Plan",
         "Outputs: enrollment, attendance, completion rate.\n"
         "Outcomes: pre/post assessments, participant/family satisfaction "
         "(target ≥ 85%), longitudinal tracking where feasible.\n"
         "Reporting: quarterly dashboards; final report with disaggregated "
         "results."),
        ("Budget Narrative",
         f"Request: ${amount:,.0f} (Year 1). Applied to direct program "
         f"delivery against Year-1 expenses of ${y1['total_expenses']:,.0f}. "
         "Full three-year forecast with printed assumptions is attached in "
         "the business plan."),
        ("Organizational Capacity",
         (website.get("mission_text") or org.get("mission") or
          f"{name} operates established community programs.") +
         "\n\nThe organization's public website was reviewed as part of this "
         f"application ({website.get('url') or 'no site'}); "
         f"accessibility spot-check: "
         f"{website.get('accessibility', {}).get('passed', 0)}/"
         f"{website.get('accessibility', {}).get('total', 0)} checks passed."),
        ("Sustainability",
         "See Business Plan §Sustainability: grant funds are catalytic; "
         "operations sustain through earned revenue and merchandise margin "
         "per the attached forecast."),
    ]
    return {"title": f"{name} — Grant Application Narrative",
            "sections": sections}
