/* Grant Writing Studio — browser-only agent engine.
   A faithful JavaScript port of the Python pipeline (grant_site/agent/*):
   forecast, pain-point research, business plan & prospectus, grant
   narrative, line-item budget + justification, readiness checklist,
   RFP extraction + response, SS-4 and TEAS Plus filing documents, and
   trademark knockout analysis. Network lookups (EIN, opportunities) are
   attempted directly from the browser and degrade gracefully to manual
   links when blocked. */

"use strict";
const Engine = (function () {

  const money = v => "$" + Math.round(v).toLocaleString("en-US");
  const money2 = v => "$" + v.toLocaleString("en-US",
    { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  /* ---------------- Pain-point research ---------------- */
  const SECTOR_PAINPOINTS = {
    education: [
      ["Persistent K-12 achievement and opportunity gaps, especially post-pandemic learning loss in math and reading.",
        "NAEP Nation's Report Card (nces.ed.gov/nationsreportcard)"],
      ["Teacher shortages and high turnover in high-poverty schools.",
        "Learning Policy Institute teacher shortage research (learningpolicyinstitute.org)"],
      ["Unequal access to enrichment, arts, and CTE programming across district funding levels.",
        "California Department of Education data (cde.ca.gov/ds)"],
    ],
    youth: [
      ["Rising youth mental-health needs; 4 in 10 high schoolers report persistent sadness or hopelessness.",
        "CDC Youth Risk Behavior Survey (cdc.gov/yrbs)"],
      ["Shortage of affordable, safe after-school placements; for every child in a program, several more are waiting.",
        "Afterschool Alliance 'America After 3PM' (afterschoolalliance.org)"],
    ],
    arts: [
      ["Arts education access declines with school poverty level despite state mandates.",
        "Arts Education Data Project (artseddata.org); Calif. Prop 28 implementation reports"],
    ],
    technology: [
      ["Digital divide: lower-income households lag in home broadband and device access.",
        "Pew Research Center internet/broadband fact sheets (pewresearch.org)"],
      ["Underrepresentation in computing pathways and STEM careers.",
        "Code.org / CSTA State of CS Education report (code.org)"],
    ],
    health: [
      ["Care deserts and coverage gaps for low-income families.", "KFF state health facts (kff.org)"],
    ],
    community: [
      ["Nonprofits report demand outpacing capacity and flat or declining individual giving.",
        "Giving USA annual report; Nonprofit Finance Fund State of the Sector survey"],
    ],
  };

  function researchPainpoints(focusAreas, localNeed, county) {
    const hits = [];
    (localNeed || "").split("\n").forEach(line => {
      line = line.trim().replace(/^[•\-*]\s*/, "");
      if (line.length > 15) {
        hits.push({ pain: line, sector: "local",
          source: "Local data provided by applicant — cite the underlying source in the final submission" });
      }
    });
    const text = (focusAreas || []).join(" ").toLowerCase();
    for (const [sector, items] of Object.entries(SECTOR_PAINPOINTS)) {
      if (text.includes(sector) || !(focusAreas || []).length) {
        for (const [pain, source] of items) {
          if (!hits.some(h => h.pain === pain)) hits.push({ pain, source, sector });
        }
      }
    }
    if (!hits.length) {
      for (const [pain, source] of SECTOR_PAINPOINTS.community) hits.push({ pain, source, sector: "community" });
    }
    const countyLabel = (county || "").trim() || "your county";
    return {
      painpoints: hits.slice(0, 10),
      method: "Local applicant-provided data is listed first (funders weight local need most heavily), followed by sector pain points from named public research sources; verify current-year figures at each source before submission.",
      localDataSources: [
        { label: "US Census QuickFacts for " + countyLabel, url: "https://www.census.gov/quickfacts/" },
        { label: "CDE DataQuest (CA school/district data)", url: "https://dq.cde.ca.gov/dataquest/" },
        { label: "kidsdata.org child well-being data for " + countyLabel, url: "https://www.kidsdata.org/" },
        { label: "CA EDD labor market data by county", url: "https://labormarketinfo.edd.ca.gov/" },
      ],
    };
  }

  /* ---------------- Forecast ---------------- */
  function buildForecast(assumptions) {
    const a = Object.assign({
      base_revenue: 50000, revenue_growth_pct: 15,
      base_expenses: 45000, expense_growth_pct: 8,
      grant_target: 25000,
      merch_units: 200, merch_price: 20, merch_unit_cost: 8, merch_growth_pct: 25,
      participants: 100, participant_growth_pct: 20,
    }, assumptions || {});
    for (const k of Object.keys(a)) {
      a[k] = Number(a[k]);
      if (!isFinite(a[k])) a[k] = 0;
    }
    const years = [];
    for (let i = 0; i < 3; i++) {
      const rev = a.base_revenue * Math.pow(1 + a.revenue_growth_pct / 100, i);
      const exp = a.base_expenses * Math.pow(1 + a.expense_growth_pct / 100, i);
      const units = a.merch_units * Math.pow(1 + a.merch_growth_pct / 100, i);
      const merchRev = units * a.merch_price, merchCost = units * a.merch_unit_cost;
      const grant = i === 0 ? a.grant_target : 0;
      years.push({
        year: i + 1,
        program_revenue: rev, grant_revenue: grant,
        merch_units: Math.round(units), merch_revenue: merchRev, merch_cogs: merchCost,
        total_revenue: rev + merchRev + grant,
        total_expenses: exp + merchCost,
        net: rev + merchRev + grant - exp - merchCost,
        participants_served: Math.round(a.participants * Math.pow(1 + a.participant_growth_pct / 100, i)),
      });
    }
    return { assumptions: a, years,
      note: "Deterministic model: every figure derives from the printed assumptions. Year-1 grant revenue equals the requested award; later years assume sustainability from earned revenue growth." };
  }

  function forecastTable(fc) {
    const y = fc.years;
    const rows = [["Line", "Year 1", "Year 2", "Year 3"]];
    const lines = [["Program revenue", "program_revenue"], ["Grant revenue", "grant_revenue"],
      ["Merchandise revenue", "merch_revenue"], ["Total revenue", "total_revenue"],
      ["Total expenses (incl. COGS)", "total_expenses"], ["Net", "net"]];
    for (const [label, key] of lines) rows.push([label, ...y.map(r => money(r[key]))]);
    rows.push(["Merch units", ...y.map(r => r.merch_units.toLocaleString())]);
    rows.push(["Participants", ...y.map(r => r.participants_served.toLocaleString())]);
    const w = [0, 1, 2, 3].map(i => Math.max(...rows.map(r => r[i].length)));
    return rows.map(r => r.map((c, i) => c.padEnd(w[i])).join("  ")).join("\n") +
      "\n\n" + fc.note;
  }

  /* ---------------- Business plan ---------------- */
  function buildBusinessPlan(org, pains, fc, tm) {
    const name = org.name || "The Organization";
    const y = fc.years, a = fc.assumptions;
    const mission = org.mission ||
      "[ACTION NEEDED: no mission statement was provided — add 2-3 sentences on who you serve, what you do, and the change you create. Reviewers score this section heavily.]";
    const sections = [
      ["Executive Summary",
        `${name} requests ${money(a.grant_target)} to expand its programs. This plan documents the community need, the organization's model and traction, and a three-year path to sustainability in which earned revenue (excluding the grant) grows from ${money(y[0].total_revenue - y[0].grant_revenue)} in Year 1 to ${money(y[2].total_revenue - y[2].grant_revenue)} in Year 3 while serving ${y[2].participants_served.toLocaleString()} people annually by Year 3.`],
      ["Mission and Organization Overview", mission],
      ["Programs and Services", org.programs || "Core programs as described by the applicant."],
      ["Market Need — Documented Pain Points",
        pains.painpoints.map(p => `• ${p.pain}\n  Source: ${p.source}`).join("\n") +
        "\n\nMethodology: " + pains.method],
      ["Target Population and Reach",
        `Year 1 reach: ${y[0].participants_served.toLocaleString()} participants, growing to ${y[2].participants_served.toLocaleString()} by Year 3 (${a.participant_growth_pct}% annual growth assumption).`],
      ["Products and Merchandise",
        (org.merchandise || "Branded merchandise supporting the mission.") +
        `\n\nMerchandise model: ${y[0].merch_units.toLocaleString()} units in Year 1 at ${money2(a.merch_price)} each (unit cost ${money2(a.merch_unit_cost)}), yielding ${money(y[0].merch_revenue - y[0].merch_cogs)} gross margin in Year 1 and ${money(y[2].merch_revenue - y[2].merch_cogs)} by Year 3.`],
      ["Three-Year Financial Forecast", forecastTable(fc)],
      ["Forecast of Success — Key Metrics",
        [`• Participants served: ${y[0].participants_served.toLocaleString()} → ${y[2].participants_served.toLocaleString()}`,
         `• Total revenue: ${money(y[0].total_revenue)} → ${money(y[2].total_revenue)}`,
         `• Net position Year 3: ${money(y[2].net)}`,
         "• Program outcomes tracked per logic model (enrollment, completion, pre/post assessment, satisfaction ≥ 85%)."].join("\n")],
      ["Sustainability Plan",
        `Grant funds are catalytic (Year 1 only in this model); ongoing operations are sustained by program revenue growing at ${a.revenue_growth_pct}%/yr, merchandise margin, and a diversified funder pipeline.`],
    ];
    if (tm) {
      sections.push(["Intellectual Property",
        `Brand: “${tm.query}”. ${tm.hitsNote} A TEAS Plus application packet has been prepared; filing and signature occur at teas.uspto.gov. ${tm.disclaimer}`]);
    }
    return { title: `${name} — Business Plan & Funding Prospectus`, sections };
  }

  /* ---------------- Budget ---------------- */
  const ALLOC = [["personnel", 0.55], ["travel", 0.03], ["equipment", 0.05],
    ["supplies", 0.10], ["contractual", 0.12], ["other", 0.05]];
  const JUST = {
    personnel: "Salaries for program staff delivering direct services (program lead and instructors), prorated to project time.",
    fringe: "Fringe benefits (payroll taxes, health, retirement) at the stated rate applied to project personnel.",
    travel: "Local mileage for program delivery sites and one program-related conference/training.",
    equipment: "Durable items required for delivery (e.g., laptops, presentation equipment) retained for program use.",
    supplies: "Consumable program materials, curriculum, and participant supplies.",
    contractual: "Specialized contracted services (e.g., evaluators, guest instructors) procured per written agreements.",
    other: "Facility use, insurance, printing, and communication costs directly attributable to the project.",
    indirect: "Indirect costs at the stated rate (de minimis or negotiated rate) covering administration and facilities.",
  };

  function buildBudget(grantTarget, items, fringePct, indirectPct, matchAmount) {
    fringePct = fringePct === "" || fringePct == null ? 22 : Number(fringePct);
    indirectPct = indirectPct === "" || indirectPct == null ? 10 : Number(indirectPct);
    matchAmount = Number(matchAmount || 0);
    let direct = {}, allocated = false;
    const hasItems = items && Object.values(items).some(v => Number(v) > 0);
    if (hasItems) {
      for (const [k, v] of Object.entries(items)) direct[k] = Number(v) || 0;
    } else {
      const pShare = 0.55, shareSum = ALLOC.reduce((s, [, v]) => s + v, 0);
      const factor = (shareSum + fringePct / 100 * pShare) * (1 + indirectPct / 100);
      const D = grantTarget / factor;
      for (const [cat, share] of ALLOC) direct[cat] = Math.round(D * share * 100) / 100;
      allocated = true;
    }
    const fringe = Math.round((direct.personnel || 0) * fringePct) / 100 * 100 / 100 * 100 / 100;
    const fringeAmt = Math.round((direct.personnel || 0) * fringePct / 100 * 100) / 100;
    const directPlusFringe = Object.values(direct).reduce((s, v) => s + v, 0) + fringeAmt;
    const indirect = Math.round(directPlusFringe * indirectPct / 100 * 100) / 100;
    const total = Math.round((directPlusFringe + indirect) * 100) / 100;
    const lines = [];
    for (const cat of ["personnel", "travel", "equipment", "supplies", "contractual", "other"]) {
      if (direct[cat]) lines.push({ category: cat[0].toUpperCase() + cat.slice(1), amount: direct[cat], justification: JUST[cat] });
    }
    lines.splice(direct.personnel ? 1 : 0, 0,
      { category: `Fringe benefits (${fringePct}%)`, amount: fringeAmt, justification: JUST.fringe });
    lines.push({ category: `Indirect (${indirectPct}%)`, amount: indirect, justification: JUST.indirect });
    return { grantTarget, lines, total, matchAmount, allocated,
      variance: Math.round((total - grantTarget) * 100) / 100,
      note: allocated
        ? "Amounts allocated from the request using standard percentages — replace with your actual figures before submission."
        : "Amounts as provided by the applicant." };
  }

  function requestJustification(budget, fc) {
    const y1 = fc.years[0];
    const withoutGrant = y1.net - y1.grant_revenue;
    let text;
    if (withoutGrant < 0) {
      text = `Without this award, Year-1 operations show a deficit of ${money(Math.abs(withoutGrant))}; the requested ${money(budget.grantTarget)} closes that gap and funds the expansion described in the project plan.`;
    } else {
      text = `Existing revenue sustains current operations (Year-1 net of ${money(withoutGrant)} before this award); the requested funds are restricted to the NEW activities in this proposal — expanded enrollment, additional staff hours, and materials — none of which occur without the award. Grant funds supplement and do not supplant existing program spending.`;
    }
    if (budget.matchAmount) text += ` The organization commits ${money(budget.matchAmount)} in match/in-kind.`;
    return text;
  }

  function buildBudgetDocument(orgName, budget, fc) {
    const rows = [["Category", "Amount", "Justification"]];
    for (const l of budget.lines) rows.push([l.category, money2(l.amount), l.justification]);
    rows.push(["TOTAL PROJECT BUDGET", money2(budget.total), ""]);
    if (budget.matchAmount) rows.push(["Applicant match / in-kind", money2(budget.matchAmount),
      "Committed by the applicant; not part of the request."]);
    const w = [0, 1].map(i => Math.max(...rows.map(r => r[i].length)));
    const table = rows.map(r => r[0].padEnd(w[0]) + "  " + r[1].padEnd(w[1]) + "  " + r[2]).join("\n");
    const sections = [
      ["Line-Item Project Budget", table + "\n\n" + budget.note],
      ["Why This Funding Is Needed", requestJustification(budget, fc)],
      ["Budget Narrative", budget.lines.map(l => `${l.category}: ${money2(l.amount)} — ${l.justification}`).join("\n\n")],
    ];
    if (Math.abs(budget.variance) > 1) {
      sections.splice(1, 0, ["Reconciliation",
        `Line items total ${money2(budget.total)} vs the requested ${money2(budget.grantTarget)} (difference ${money2(budget.variance)}). Adjust line items or the request so they match before submission.`]);
    }
    return { title: `${orgName} — Project Budget & Justification`, sections };
  }

  /* ---------------- Readiness ---------------- */
  const REGISTRATIONS = [
    ["sam_gov", "SAM.gov registration + UEI number", "REQUIRED for all federal grants. Free at https://sam.gov — allow up to several weeks; renew annually."],
    ["ca_registry", "California AG Registry of Charities (current RRF-1)", "Required for CA charities soliciting donations: https://oag.ca.gov/charities"],
    ["sos_good_standing", "Secretary of State good standing", "Check https://bizfileonline.sos.ca.gov/search/business"],
    ["ftb_status", "Franchise Tax Board exempt/good standing", "Check https://www.ftb.ca.gov/help/business/entity-status-letter.asp"],
  ];
  const ATTACHMENTS = [
    ["board_list", "Board of directors list with affiliations", "Names, roles, and professional affiliations of all board members."],
    ["bylaws", "Articles of incorporation and bylaws", "Certified copies from your Secretary of State filing."],
    ["form_990", "Most recent IRS Form 990 / 990-EZ / 990-N", "Downloadable from your records or ProPublica Nonprofit Explorer."],
    ["financials", "Audited financial statements or board-approved budget", "Audit if revenue requires it; otherwise most recent board-approved financials."],
    ["org_budget", "Current organizational operating budget", ""],
    ["letters_support", "Letters of support / MOUs from partners", "Request 2-4 from schools, districts, or partner organizations."],
    ["org_chart", "Organizational chart", ""],
    ["staff_bios", "Key staff bios or resumes", ""],
    ["logic_model", "Logic model (inputs → activities → outputs → outcomes)", "One page; reviewers expect it for program grants."],
  ];

  function assessReadiness(einResult, have) {
    have = new Set(have || []);
    const items = [];
    const einOk = !!(einResult && einResult.found);
    items.push({ key: "ein", label: "EIN (federal tax ID)",
      status: einOk ? "ready" : "action_needed",
      detail: einOk ? `EIN ${einResult.ein}` : "No EIN confirmed — file the prepared SS-4 (see the SS-4 document)." });
    items.push({ key: "501c3", label: "501(c)(3) tax-exempt status",
      status: have.has("irs_determination") ? "ready" : "verify",
      detail: have.has("irs_determination") ? "Attested by applicant."
        : "Verify at https://apps.irs.gov/app/eos/ or file Form 1023/1023-EZ." });
    for (const [key, label, detail] of REGISTRATIONS) {
      items.push({ key, label, status: have.has(key) ? "ready" : "verify", detail });
    }
    for (const [key, label, detail] of ATTACHMENTS) {
      items.push({ key, label, status: have.has(key) ? "ready" : "needed", detail });
    }
    return { items, ready: items.filter(i => i.status === "ready").length, total: items.length };
  }

  function buildReadinessDocument(orgName, r) {
    const by = { action_needed: [], verify: [], needed: [], ready: [] };
    for (const i of r.items) by[i.status].push(i);
    const fmt = arr => arr.map(i => "• " + i.label + (i.detail ? "\n  " + i.detail : "")).join("\n\n") || "None.";
    return {
      title: `${orgName} — Grant Application Readiness Checklist`,
      sections: [
        ["Summary", `${r.ready} of ${r.total} items ready. Most funders will not accept an application missing the items below — complete these before the deadline.`],
        ["Action needed now", fmt(by.action_needed)],
        ["Verify before applying", fmt(by.verify)],
        ["Documents to gather", fmt(by.needed)],
        ["Ready", fmt(by.ready)],
      ],
    };
  }

  /* ---------------- Grant narrative ---------------- */
  function buildGrantNarrative(org, pains, fc, einResult, budget) {
    const name = org.name || "The Organization";
    const amount = fc.assumptions.grant_target, y1 = fc.years[0];
    const einLine = einResult && einResult.found ? `EIN: ${einResult.ein}`
      : "EIN: application in progress (completed SS-4 document prepared)";
    return {
      title: `${name} — Grant Application Narrative`,
      sections: [
        ["Applicant Information", `${name}\n${einLine}\nContact: ${org.contact_email || "on file"}`],
        ["Statement of Need",
          `${name} addresses documented community needs:\n\n` +
          pains.painpoints.slice(0, 6).map(p => `• ${p.pain} (Source: ${p.source})`).join("\n")],
        ["Project Description and Goals",
          `With an award of ${money(amount)}, ${name} will deliver ${org.programs || "its core programs"} to ${y1.participants_served.toLocaleString()} participants in Year 1.\n\nGoals:\n1. Expand program capacity and enrollment.\n2. Achieve measurable participant outcomes (pre/post gains).\n3. Build earned-revenue sustainability so the program continues beyond the grant period.`],
        ["Evaluation Plan",
          "Outputs: enrollment, attendance, completion rate.\nOutcomes: pre/post assessments, participant/family satisfaction (target ≥ 85%), longitudinal tracking where feasible.\nReporting: quarterly dashboards; final report with disaggregated results."],
        ["Budget Narrative",
          requestJustification(budget, fc) + "\n\nSee the attached line-item Project Budget & Justification document."],
        ["Organizational Capacity",
          (org.mission || `${name} operates established community programs.`)],
        ["Sustainability",
          "See Business Plan §Sustainability: grant funds are catalytic; operations sustain through earned revenue and merchandise margin per the attached forecast."],
      ],
    };
  }

  /* ---------------- RFP ingestion ---------------- */
  const TOPIC_MAP = [
    [["need", "problem", "gap", "data", "population", "community", "statement of need", "target"], "need"],
    [["budget", "cost", "funds", "expenditure", "financial", "allocat"], "budget"],
    [["evaluat", "measure", "outcome", "metric", "assess", "success", "data collection", "report"], "evaluation"],
    [["sustain", "continu", "beyond the grant", "future funding"], "sustainability"],
    [["capacity", "experience", "qualification", "history", "staff", "organization", "track record", "board"], "capacity"],
    [["partner", "collaborat", "letter of support", "mou"], "partners"],
    [["goal", "objective", "activit", "project", "program", "plan", "implement", "timeline", "approach", "describe the"], "project"],
  ];
  function classify(q) {
    const low = q.toLowerCase();
    for (const [kws, topic] of TOPIC_MAP) if (kws.some(k => low.includes(k))) return topic;
    return "project";
  }

  function extractRfp(text) {
    text = text || "";
    const questions = [], seen = new Set();
    const qRe = /^\s*(?:\d+[.)]\s+|[A-Z][.)]\s+|[•\-*]\s+|Q\d+[:.]?\s+)?(.{15,400}?\?)\s*$/gm;
    const impRe = /^\s*(?:\d+[.)]\s+|[A-Z][.)]\s+|[•\-*]\s+)?((?:Describe|Explain|Provide|Identify|Summarize|List|Demonstrate|Detail|Discuss|Outline|Include|State|Address)\b.{10,400}?[.:])\s*$/gmi;
    for (const re of [qRe, impRe]) {
      let m;
      while ((m = re.exec(text)) !== null) {
        const q = m[1].replace(/\s+/g, " ").trim();
        if (!seen.has(q.toLowerCase())) { seen.add(q.toLowerCase()); questions.push({ question: q, topic: classify(q) }); }
      }
    }
    const limits = [];
    const limRe = /(?:no more than|not to exceed|maximum(?: of)?|limited? to|up to)\s+(\d{1,4})\s*(pages?|words?|characters?)/gi;
    let lm; while ((lm = limRe.exec(text)) !== null) limits.push({ limit: Number(lm[1]), unit: lm[2].toLowerCase().replace(/s$/, "") + "s" });
    const deadlines = [];
    const dRe = /(?:due|deadline|submitted?|received)\D{0,40}?((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}|\d{1,2}\/\d{1,2}\/\d{2,4})/gi;
    let dm; while ((dm = dRe.exec(text)) !== null) if (!deadlines.includes(dm[1])) deadlines.push(dm[1]);
    const scoring = [];
    const sRe = /^\s*(.{5,120}?)[\s.\-–—]*\(?(\d{1,3})\s*(?:points?|pts\.?|%)\)?\s*$/gmi;
    let sm; while ((sm = sRe.exec(text)) !== null) { const p = Number(sm[2]); if (p <= 100) scoring.push({ criterion: sm[1].trim(), points: p }); }
    return { questions: questions.slice(0, 30), limits: limits.slice(0, 10),
      deadlines: deadlines.slice(0, 5), scoring: scoring.slice(0, 15) };
  }

  function buildRfpResponse(rfp, ctx, funderName) {
    const route = {
      need: ctx.need, budget: ctx.budget, evaluation: ctx.evaluation,
      sustainability: ctx.sustainability, capacity: ctx.capacity,
      partners: ctx.partners || "Partnerships and letters of support are being assembled; see the readiness checklist for status.",
      project: ctx.project,
    };
    const meta = [];
    if (rfp.deadlines.length) meta.push("Deadline(s) found in RFP: " + rfp.deadlines.join("; "));
    if (rfp.limits.length) meta.push("Stated limits: " + rfp.limits.map(l => `${l.limit} ${l.unit}`).join("; "));
    if (rfp.scoring.length) meta.push("Scoring rubric found: " + rfp.scoring.map(s => `${s.criterion} (${s.points} pts)`).join("; "));
    const sections = [["About this response",
      "Answers below are mapped to the funder's own questions extracted from the RFP text. Verify each answer against the RFP before submission and trim to the stated limits." +
      (meta.length ? "\n\n" + meta.join("\n") : "")]];
    if (!rfp.questions.length) {
      sections.push(["No questions detected",
        "The parser found no application questions in the pasted RFP text. Paste the section of the RFP that lists the narrative questions (often called 'Application Questions', 'Narrative', or 'Selection Criteria') and re-run."]);
    }
    rfp.questions.forEach((q, i) => sections.push([`${i + 1}. ${q.question}`, route[q.topic] || ctx.project]));
    return { title: `${ctx.orgName} — Response to ${funderName || "Funder"} RFP`, sections };
  }

  /* ---------------- EIN / SS-4 ---------------- */
  async function lookupEin(orgName, state, knownEin) {
    const digits = (knownEin || "").replace(/\D/g, "");
    if (digits.length === 9) {
      return { found: true, ein: digits.slice(0, 2) + "-" + digits.slice(2), source: "user-provided", matches: [] };
    }
    const result = { found: false, ein: null, matches: [], error: null };
    if (!orgName) { result.error = "No organization name provided."; return result; }
    try {
      let url = "https://projects.propublica.org/nonprofits/api/v2/search.json?q=" + encodeURIComponent(orgName);
      if (state) url += "&state%5Bid%5D=" + encodeURIComponent(state.toUpperCase());
      const resp = await fetch(url);
      if (!resp.ok) throw new Error("HTTP " + resp.status);
      const orgs = (await resp.json()).organizations || [];
      result.matches = orgs.slice(0, 8).map(o => ({
        ein: String(o.ein).length === 9 ? String(o.ein).slice(0, 2) + "-" + String(o.ein).slice(2) : String(o.ein),
        name: o.name, city: o.city, state: o.state }));
      const exact = result.matches.find(m => (m.name || "").toLowerCase() === orgName.toLowerCase());
      if (exact) { result.found = true; result.ein = exact.ein; result.source = "propublica-nonprofit-explorer"; }
      else if (result.matches.length === 1) { result.found = true; result.ein = result.matches[0].ein; result.source = "propublica-nonprofit-explorer"; }
    } catch (e) {
      result.error = "Automated EIN search unavailable from the browser (" + e.message +
        "). Search manually at https://projects.propublica.org/nonprofits/ or https://apps.irs.gov/app/eos/";
    }
    return result;
  }

  function buildSs4Document(info) {
    const lines = [
      ["Line 1 — Legal name of entity", info.legal_name],
      ["Line 2 — Trade name / DBA (if different)", ""],
      ["Lines 4a–4b — Mailing address", info.mailing_address],
      ["Line 6 — County and state of principal business", [info.county, info.state].filter(Boolean).join(", ")],
      ["Line 7a — Responsible party (name)", info.responsible_party],
      ["Line 7b — Responsible party SSN/ITIN", "(entered directly with the IRS; never store this in any app)"],
      ["Line 9a — Type of entity", info.entity_type],
      ["Line 10 — Reason for applying", "Started new business / banking purpose / applying for grants"],
      ["Line 11 — Date business started/acquired", info.start_date],
      ["Line 12 — Closing month of accounting year", "December"],
      ["Line 16 — Principal activity", info.activity],
    ];
    const walkthrough = [
      "1. Go to the IRS online EIN Assistant (link below). It is open Mon–Fri, 7am–10pm Eastern, and is completely free.",
      "2. Choose your legal structure when asked (matches Line 9a below).",
      "3. Choose why you are requesting an EIN (matches Line 10).",
      "4. Enter the responsible party's name and SSN/ITIN (Lines 7a–7b). Enter this directly on the IRS screen only — this document deliberately does not contain it.",
      "5. Enter the addresses and business name exactly as shown (the IRS matches punctuation).",
      "6. Answer the activity questions using Line 16.",
      "7. Choose 'Receive letter online' — the EIN is issued immediately; save the CP 575 confirmation letter PDF the moment it appears (it cannot be re-downloaded).",
      "8. Nonprofits: after receiving the EIN, apply for tax-exempt status separately (Form 1023 or 1023-EZ).",
    ].join("\n");
    return {
      title: `${info.legal_name || "Organization"} — IRS Form SS-4 (EIN Application), Completed Worksheet`,
      sections: [
        ["How to file (about 10 minutes)",
          walkthrough + "\n\nFile at: https://www.irs.gov/businesses/small-businesses-self-employed/apply-for-an-employer-identification-number-ein-online"],
        ["Your completed SS-4, line by line",
          lines.map(([l, v]) => `${l}:\n    ${v || "(complete before filing)"}`).join("\n")],
      ],
    };
  }

  /* ---------------- Trademark / TEAS ---------------- */
  const NICE = { education: 41, training: 41, tutoring: 41, workshops: 41,
    software: 42, saas: 42, website: 42, app: 9, clothing: 25, apparel: 25,
    "t-shirt": 25, merchandise: 25, hats: 25, hoodies: 25, consulting: 35, advertising: 35,
    retail: 35, charitable: 36, fundraising: 36, grants: 36, books: 16, printed: 16,
    toys: 28, food: 43 };

  function suggestClasses(goods) {
    const t = (goods || "").toLowerCase();
    const cls = [...new Set(Object.entries(NICE)
      .filter(([k]) => new RegExp("\\b" + k.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b").test(t))
      .map(([, v]) => v))].sort((a, b) => a - b);
    return cls.length ? cls : [35];
  }

  function phoneticVariants(name) {
    const base = name.toLowerCase().replace(/[^a-z0-9 ]/g, "").trim();
    const out = new Set([base, base.replace(/ /g, ""), base.replace(/s\b/g, ""),
      base.replace(/\b(the|a|an)\b\s*/g, "").trim()]);
    for (const [a, b] of [["ph", "f"], ["c", "k"], ["k", "c"], ["z", "s"], ["s", "z"],
      ["x", "ks"], ["qu", "kw"], ["ee", "i"], ["y", "i"], ["ie", "y"], ["gh", "g"]]) {
      if (base.includes(a)) out.add(base.split(a).join(b));
    }
    return [...out].filter(Boolean).sort();
  }

  function searchTrademark(name, goods) {
    const q = encodeURIComponent(name);
    return {
      query: name,
      variants: phoneticVariants(name),
      classes: suggestClasses(goods),
      hitsNote: "Automated keyword search requires the full server version; use the manual search links to confirm.",
      disclaimer: "Automated knockout analysis only — likelihood-of-confusion analysis under 15 U.S.C. §1052(d) considers sound, appearance, meaning, and relatedness of goods. Confirm results in the USPTO search system and consider a trademark attorney before adopting or filing a mark. Not legal advice.",
      manualLinks: [
        { label: "USPTO Trademark Search (exact and similar marks)", url: "https://tmsearch.uspto.gov/search/search-information?query=" + q },
        { label: "California Secretary of State trademark search", url: "https://bizfileonline.sos.ca.gov/search/trademarks" },
        { label: "Google web search for similar concepts", url: "https://www.google.com/search?q=%22" + q + "%22" },
        { label: "Domain availability", url: "https://www.namecheap.com/domains/registration/results/?domain=" + q },
      ],
    };
  }

  function buildTeasDocument(tm, owner, goods, inUse) {
    const fee = 250 * tm.classes.length;
    const fields = [
      `Mark (standard characters): ${tm.query}`,
      "Mark type: Standard character mark (protects the wording in any font/style/color)",
      `Owner name: ${owner.name || "(complete before filing)"}`,
      `Owner entity type: ${owner.entity_type || "(complete before filing)"}`,
      `Owner email: ${owner.email || "(complete before filing)"}`,
      `Filing basis: ${inUse ? "1(a) Use in commerce — specimen required" : "1(b) Intent to use — Statement of Use filed later (additional fee per class)"}`,
      `International class(es): ${tm.classes.join(", ")}`,
      `Goods/services description: ${goods || "(complete before filing)"}`,
      "ID Manual note: TEAS Plus requires descriptions selected from the USPTO ID Manual: https://idm-tmng.uspto.gov/id-master-list-public.html",
    ].join("\n");
    const walkthrough = [
      "1. Go to the TEAS Plus initial application form (link below) and sign in with a USPTO.gov account (free to create).",
      "2. Applicant screen: enter the owner name, entity type, and address exactly as shown in the fields section.",
      "3. Mark screen: choose 'Standard Characters' and type the mark exactly as shown.",
      "4. Goods/Services screen: search the ID Manual for the description(s) below and select the matching entries in the listed class(es) — TEAS Plus requires ID Manual entries.",
      "5. Basis screen: select the filing basis shown. For use in commerce (1a), upload a specimen and enter dates of first use. For intent to use (1b), no specimen yet.",
      "6. Correspondence screen: use the owner email; USPTO sends all office actions there.",
      "7. Signature screen: the owner (or attorney) signs electronically — type /Full Name/ in the signature box.",
      "8. Pay the fee shown below by card. Save the serial number from the confirmation; track status at tsdr.uspto.gov.",
    ].join("\n");
    return {
      title: `“${tm.query}” — USPTO TEAS Plus Trademark Application, Completed & Ready to File`,
      sections: [
        ["Clearance summary",
          tm.hitsNote + " Knockout variants checked: " + tm.variants.join(", ") + ".\n\n" + tm.disclaimer],
        ["How to file (15–30 minutes)",
          walkthrough + `\n\nFile at: https://teas.uspto.gov/forms/bas/\n\nFee: $${fee} (${tm.classes.length} class(es) × $250). Verify current fees at uspto.gov/trademark/trademark-fee-information`],
        ["Your completed application fields", fields],
        ["What happens after filing",
          "• ~1 week: serial number issued and application viewable in TSDR\n• ~8-10 months: examining attorney review (office actions possible)\n• ~1 month publication for opposition\n• Registration (1a) or Notice of Allowance (1b)\n\nSignature requirement: must be personally signed at teas.uspto.gov by the applicant or an authorized attorney — the USPTO has no third-party filing API, so this packet is prepared-to-file, not auto-filed."],
      ],
    };
  }

  /* ---------------- Opportunities ---------------- */
  async function searchOpportunities(keyword) {
    const out = { keyword, federal: { ok: false, results: [] }, manualLinks: [
      { label: "Grants.gov search", url: "https://grants.gov/search-grants?query=" + encodeURIComponent(keyword) },
      { label: "California Grants Portal", url: "https://www.grants.ca.gov/grants/?fwp_search=" + encodeURIComponent(keyword) },
      { label: "Candid Foundation Directory (library access often free)", url: "https://candid.org/find-funding" },
    ] };
    try {
      const resp = await fetch("https://api.grants.gov/v1/api/search2", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ keyword, oppStatuses: "posted", rows: 10, startRecordNum: 0 }),
      });
      if (!resp.ok) throw new Error("HTTP " + resp.status);
      const hits = ((await resp.json()).data || {}).oppHits || [];
      out.federal = { ok: true, results: hits.slice(0, 10).map(h => ({
        title: h.title, agency: h.agencyName || h.agency, close_date: h.closeDate,
        url: "https://grants.gov/search-results-detail/" + h.id })) };
    } catch (e) {
      out.federal.error = e.message;
    }
    return out;
  }

  return { researchPainpoints, buildForecast, buildBusinessPlan, buildBudget,
    buildBudgetDocument, requestJustification, assessReadiness,
    buildReadinessDocument, buildGrantNarrative, extractRfp, buildRfpResponse,
    lookupEin, buildSs4Document, searchTrademark, buildTeasDocument,
    searchOpportunities };
})();
