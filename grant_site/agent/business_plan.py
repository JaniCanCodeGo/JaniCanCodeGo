"""Stage 3 — Business plan + investor/funder prospectus generation.

Combines the website review, user-supplied program data, and documented
pain-point research into a full business plan, and builds a 3-year financial
forecast (revenue, expenses, net, merchandise line) from the user's baseline
assumptions using a deterministic growth model — no invented numbers: every
figure traces back to a stated assumption that is printed with the forecast.
"""

SECTOR_PAINPOINTS = {
    # Documented sector pain points with verifiable public sources. The
    # research module surfaces these as *starting points*; each generated
    # plan prints the source so reviewers can verify and update figures.
    "education": [
        {"pain": "Persistent K-12 achievement and opportunity gaps, "
                 "especially post-pandemic learning loss in math and reading.",
         "source": "NAEP Nation's Report Card (nces.ed.gov/nationsreportcard)"},
        {"pain": "Teacher shortages and high turnover in high-poverty schools.",
         "source": "Learning Policy Institute teacher shortage research "
                   "(learningpolicyinstitute.org)"},
        {"pain": "Unequal access to enrichment, arts, and CTE programming "
                 "across district funding levels.",
         "source": "California Department of Education data (cde.ca.gov/ds)"},
    ],
    "youth": [
        {"pain": "Rising youth mental-health needs; 4 in 10 high schoolers "
                 "report persistent sadness or hopelessness.",
         "source": "CDC Youth Risk Behavior Survey (cdc.gov/yrbs)"},
        {"pain": "Shortage of affordable, safe after-school placements; "
                 "for every child in a program, several more are waiting.",
         "source": "Afterschool Alliance 'America After 3PM' (afterschoolalliance.org)"},
    ],
    "arts": [
        {"pain": "Arts education access declines with school poverty level "
                 "despite state mandates.",
         "source": "Arts Education Data Project (artseddata.org); Calif. "
                   "Prop 28 implementation reports"},
    ],
    "technology": [
        {"pain": "Digital divide: lower-income households lag in home "
                 "broadband and device access.",
         "source": "Pew Research Center internet/broadband fact sheets "
                   "(pewresearch.org)"},
        {"pain": "Underrepresentation in computing pathways and STEM careers.",
         "source": "Code.org / CSTA State of CS Education report (code.org)"},
    ],
    "health": [
        {"pain": "Care deserts and coverage gaps for low-income families.",
         "source": "KFF state health facts (kff.org)"},
    ],
    "community": [
        {"pain": "Nonprofits report demand outpacing capacity and flat or "
                 "declining individual giving.",
         "source": "Giving USA annual report; Nonprofit Finance Fund State "
                   "of the Sector survey"},
    ],
}


def research_painpoints(focus_areas, mission_text="", local_need="",
                        county=""):
    """Assemble documented pain points for the plan's needs statement.

    local_need: applicant-provided local statistics (strongest evidence —
    listed first). county: used to generate local data-source links so the
    applicant can replace national figures with local ones.
    """
    text = " ".join(focus_areas or []).lower() + " " + (mission_text or "").lower()
    hits, seen = [], set()
    for line in (local_need or "").splitlines():
        line = line.strip().lstrip("•-* ")
        if len(line) > 15:
            hits.append({"pain": line,
                         "source": "Local data provided by applicant — "
                                   "cite the underlying source in the final "
                                   "submission",
                         "sector": "local"})
            seen.add(line)
    for sector, items in SECTOR_PAINPOINTS.items():
        if sector in text or not focus_areas:
            for item in items:
                if item["pain"] not in seen:
                    seen.add(item["pain"])
                    hits.append(dict(item, sector=sector))
    if not hits:
        hits = [dict(item, sector=sector)
                for sector in ("community",)
                for item in SECTOR_PAINPOINTS[sector]]
    county_label = county.strip() if county else "your county"
    return {
        "painpoints": hits[:10],
        "method": ("Local applicant-provided data is listed first (funders "
                   "weight local need most heavily), followed by sector "
                   "pain points from named public research sources; verify "
                   "current-year figures at each source before submission."),
        "local_data_sources": [
            {"label": f"US Census QuickFacts for {county_label}",
             "url": "https://www.census.gov/quickfacts/"},
            {"label": "CDE DataQuest (CA school/district data)",
             "url": "https://dq.cde.ca.gov/dataquest/"},
            {"label": f"kidsdata.org child well-being data for {county_label}",
             "url": "https://www.kidsdata.org/"},
            {"label": "CA EDD labor market data by county",
             "url": "https://labormarketinfo.edd.ca.gov/"},
        ],
    }


def build_forecast(assumptions):
    """3-year forecast from explicit baseline assumptions.

    assumptions keys (all optional, sensible small-org defaults):
      base_revenue, revenue_growth_pct, base_expenses, expense_growth_pct,
      grant_target, merch_units, merch_price, merch_unit_cost,
      merch_growth_pct, participants, participant_growth_pct
    """
    a = {
        "base_revenue": 50000.0, "revenue_growth_pct": 15.0,
        "base_expenses": 45000.0, "expense_growth_pct": 8.0,
        "grant_target": 25000.0,
        "merch_units": 200, "merch_price": 20.0, "merch_unit_cost": 8.0,
        "merch_growth_pct": 25.0,
        "participants": 100, "participant_growth_pct": 20.0,
    }
    for key in a:
        if assumptions.get(key) not in (None, ""):
            a[key] = float(assumptions[key])

    years = []
    for i in range(3):
        rev = a["base_revenue"] * (1 + a["revenue_growth_pct"] / 100) ** i
        exp = a["base_expenses"] * (1 + a["expense_growth_pct"] / 100) ** i
        units = a["merch_units"] * (1 + a["merch_growth_pct"] / 100) ** i
        merch_rev = units * a["merch_price"]
        merch_cost = units * a["merch_unit_cost"]
        participants = a["participants"] * (1 + a["participant_growth_pct"] / 100) ** i
        years.append({
            "year": i + 1,
            "program_revenue": round(rev, 2),
            "grant_revenue": round(a["grant_target"], 2) if i == 0 else 0.0,
            "merch_units": int(round(units)),
            "merch_revenue": round(merch_rev, 2),
            "merch_cogs": round(merch_cost, 2),
            "total_revenue": round(rev + merch_rev +
                                   (a["grant_target"] if i == 0 else 0), 2),
            "total_expenses": round(exp + merch_cost, 2),
            "net": round(rev + merch_rev +
                         (a["grant_target"] if i == 0 else 0)
                         - exp - merch_cost, 2),
            "participants_served": int(round(participants)),
        })
    return {"assumptions": a, "years": years,
            "note": ("Deterministic model: every figure derives from the "
                     "printed assumptions. Year-1 grant revenue equals the "
                     "requested award; later years assume sustainability "
                     "from earned revenue growth.")}


def build_business_plan(org, website, painpoints, forecast, trademark=None):
    """Assemble the full business plan + prospectus as structured sections."""
    name = org.get("name", "The Organization")
    mission = (org.get("mission") or website.get("mission_text") or
               website.get("description") or
               "[ACTION NEEDED: no mission statement was provided or found "
               "on the website — add 2-3 sentences on who you serve, what "
               "you do, and the change you create. Reviewers score this "
               "section heavily.]")
    programs = org.get("programs") or "Core programs as described on the website."
    merch = org.get("merchandise") or "Branded merchandise supporting the mission."
    y = forecast["years"]

    sections = [
        ("Executive Summary",
         f"{name} requests ${forecast['assumptions']['grant_target']:,.0f} "
         f"to expand its programs. This plan documents the community need, "
         f"the organization's model and traction, and a three-year path to "
         f"sustainability in which earned revenue (excluding the grant) "
         f"grows from "
         f"${y[0]['total_revenue'] - y[0]['grant_revenue']:,.0f} in Year 1 "
         f"to ${y[2]['total_revenue'] - y[2]['grant_revenue']:,.0f} in "
         f"Year 3 while serving "
         f"{y[2]['participants_served']:,} people annually by Year 3."),
        ("Mission and Organization Overview", mission.strip()),
        ("Programs and Services", programs),
        ("Market Need — Documented Pain Points",
         "\n".join(f"• {p['pain']}\n  Source: {p['source']}"
                   for p in painpoints["painpoints"]) +
         f"\n\nMethodology: {painpoints['method']}"),
        ("Target Population and Reach",
         f"Year 1 reach: {y[0]['participants_served']:,} participants, "
         f"growing to {y[2]['participants_served']:,} by Year 3 "
         f"({forecast['assumptions']['participant_growth_pct']:.0f}% annual "
         "growth assumption)."),
        ("Products and Merchandise",
         f"{merch}\n\nMerchandise model: "
         f"{y[0]['merch_units']:,} units in Year 1 at "
         f"${forecast['assumptions']['merch_price']:,.2f} each "
         f"(unit cost ${forecast['assumptions']['merch_unit_cost']:,.2f}), "
         f"yielding ${y[0]['merch_revenue'] - y[0]['merch_cogs']:,.0f} gross "
         f"margin in Year 1 and "
         f"${y[2]['merch_revenue'] - y[2]['merch_cogs']:,.0f} by Year 3."),
        ("Three-Year Financial Forecast", _forecast_table(forecast)),
        ("Forecast of Success — Key Metrics",
         "\n".join([
             f"• Participants served: {y[0]['participants_served']:,} → "
             f"{y[2]['participants_served']:,}",
             f"• Total revenue: ${y[0]['total_revenue']:,.0f} → "
             f"${y[2]['total_revenue']:,.0f}",
             f"• Net position Year 3: ${y[2]['net']:,.0f}",
             "• Program outcomes tracked per logic model (enrollment, "
             "completion, pre/post assessment, satisfaction ≥ 85%).",
         ])),
        ("Sustainability Plan",
         "Grant funds are catalytic (Year 1 only in this model); ongoing "
         "operations are sustained by program revenue growing at "
         f"{forecast['assumptions']['revenue_growth_pct']:.0f}%/yr, "
         "merchandise margin, and a diversified funder pipeline."),
    ]
    if trademark:
        risk = ("No conflicting live marks surfaced in the automated search."
                if not trademark.get("api_hits")
                else f"{len(trademark['api_hits'])} potentially similar "
                     "mark(s) found — see trademark report.")
        sections.append((
            "Intellectual Property",
            f"Brand: “{trademark['query']}”. {risk} "
            "A TEAS Plus application packet has been prepared; filing and "
            "signature occur at teas.uspto.gov. " + trademark["disclaimer"]))
    return {"title": f"{name} — Business Plan & Funding Prospectus",
            "sections": sections}


def _forecast_table(forecast):
    rows = [("Line", "Year 1", "Year 2", "Year 3")]
    y = forecast["years"]

    def money(v):
        return f"${v:,.0f}"

    for label, key in [("Program revenue", "program_revenue"),
                       ("Grant revenue", "grant_revenue"),
                       ("Merchandise revenue", "merch_revenue"),
                       ("Total revenue", "total_revenue"),
                       ("Total expenses (incl. COGS)", "total_expenses"),
                       ("Net", "net")]:
        rows.append((label,) + tuple(money(yr[key]) for yr in y))
    rows.append(("Merch units",) + tuple(f"{yr['merch_units']:,}" for yr in y))
    rows.append(("Participants",) +
                tuple(f"{yr['participants_served']:,}" for yr in y))
    widths = [max(len(r[i]) for r in rows) for i in range(4)]
    lines = ["  ".join(cell.ljust(widths[i]) for i, cell in enumerate(row))
             for row in rows]
    lines.append("")
    lines.append("Assumptions: " + ", ".join(
        f"{k}={v}" for k, v in forecast["assumptions"].items()))
    lines.append(forecast["note"])
    return "\n".join(lines)
