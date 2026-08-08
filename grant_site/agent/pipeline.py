"""Full grant-writing pipeline orchestrator.

run_pipeline(payload, out_dir, progress) executes:
  1. Website review (content + Section 508 spot-check of the org's site)
  2. EIN lookup — prepares an SS-4 application worksheet if none exists
  3. Pain-point research + 3-year forecast + business plan & prospectus
  4. Trademark knockout search + TEAS Plus application packet
  5. Grant narrative assembly
and writes every deliverable to out_dir.
"""

import traceback

from . import budget as budget_mod
from . import business_plan, documents, ein, grant_writer, narrative_ai
from . import readiness, rfp, trademark, website_review


def run_pipeline(payload, out_dir, progress=None):
    """payload: dict from the web form (see server.RunRequest)."""
    def report(step, status, detail=""):
        if progress:
            progress(step, status, detail)

    org = {
        "name": payload.get("org_name", "").strip(),
        "mission": payload.get("mission", ""),
        "programs": payload.get("programs", ""),
        "merchandise": payload.get("merchandise", ""),
        "contact_email": payload.get("contact_email", ""),
    }
    results = {"org": org, "files": []}
    base = documents.safe_name(org["name"] or "organization")

    # 1. Website review
    report("website_review", "running")
    site = website_review.review_website(payload.get("website_url", ""))
    results["website_review"] = site
    report("website_review", "done",
           site["error"] or f"Reviewed {site['url']}")

    # 2. EIN
    report("ein", "running")
    ein_result = ein.lookup_ein(org["name"], payload.get("state"),
                                payload.get("ein"))
    if not ein_result["found"]:
        # Rebuild the SS-4 with everything the form gave us and emit it as
        # its own ready-to-file document.
        ein_result["application"] = ein.build_ss4_worksheet({
            "legal_name": org["name"],
            "state": payload.get("state", ""),
            "mailing_address": payload.get("mailing_address", ""),
            "responsible_party": payload.get("responsible_party", ""),
            "entity_type": payload.get("entity_type", ""),
            "start_date": payload.get("start_date", ""),
            "activity": payload.get("programs", ""),
        })
        results["files"] += documents.write_document(
            ein.build_ss4_document(ein_result["application"], org["name"]),
            out_dir, f"{base}_IRS_SS4_EIN_Application")
    results["ein"] = ein_result
    report("ein", "done",
           f"EIN {ein_result['ein']}" if ein_result["found"]
           else "No EIN found — completed SS-4 application document prepared")

    # 3. Research, forecast, business plan
    report("business_plan", "running")
    focus = [f.strip() for f in
             (payload.get("focus_areas") or "").split(",") if f.strip()]
    pains = business_plan.research_painpoints(
        focus, site.get("mission_text"),
        local_need=payload.get("local_need", ""),
        county=payload.get("county", ""))
    forecast = business_plan.build_forecast(payload.get("forecast") or {})
    results["painpoints"] = pains
    results["forecast"] = forecast

    # Line-item budget + justification
    report("budget", "running")
    budget = budget_mod.build_budget(
        forecast["assumptions"]["grant_target"],
        items=payload.get("budget_items") or None,
        fringe_pct=payload.get("fringe_pct"),
        indirect_pct=payload.get("indirect_pct"),
        match_amount=payload.get("match_amount") or 0)
    results["budget"] = budget
    budget_doc = budget_mod.build_budget_document(org["name"], budget,
                                                  forecast)
    results["files"] += documents.write_document(
        budget_doc, out_dir, f"{base}_Project_Budget_Justification")
    report("budget", "done", f"Line-item budget totals ${budget['total']:,.0f}")

    # Readiness checklist (registrations + attachments)
    report("readiness", "running")
    ready = readiness.assess_readiness(ein_result,
                                       payload.get("have_items") or [])
    results["readiness"] = ready
    results["files"] += documents.write_document(
        readiness.build_readiness_document(org["name"], ready),
        out_dir, f"{base}_Readiness_Checklist")
    report("readiness", "done",
           f"{ready['ready']}/{ready['total']} readiness items in place")

    # 4. Trademark
    tm_result = None
    mark = (payload.get("trademark_name") or org["name"]).strip()
    if mark:
        report("trademark", "running")
        tm_result = trademark.search_trademark(
            mark, payload.get("merchandise", "") + " " +
            (payload.get("programs") or ""))
        tm_result["teas_application"] = trademark.build_teas_application(
            mark,
            owner={"name": org["name"], "email": org["contact_email"],
                   "entity_type": payload.get("entity_type", "")},
            goods_description=payload.get("merchandise") or
            payload.get("programs") or "",
            use_in_commerce=bool(payload.get("mark_in_use")))
        results["files"] += documents.write_document(
            trademark.build_teas_document(tm_result["teas_application"],
                                          tm_result),
            out_dir, f"{base}_USPTO_TEAS_Trademark_Application")
        results["trademark"] = tm_result
        report("trademark", "done",
               f"{len(tm_result['api_hits'])} potential conflicts flagged"
               if tm_result["api_hits"] else
               "Knockout search complete; TEAS packet prepared")

    plan = business_plan.build_business_plan(org, site, pains, forecast,
                                             tm_result)
    results["files"] += documents.write_document(
        plan, out_dir, f"{base}_Business_Plan_Prospectus")
    report("business_plan", "done", "Business plan & prospectus written")

    # 5. Grant narrative
    report("grant", "running")
    grant = grant_writer.build_grant_narrative(org, site, pains, forecast,
                                               ein_result)
    report("grant", "done", "Grant narrative drafted")

    # 6. RFP-specific response (if an RFP was pasted)
    rfp_doc = None
    rfp_text = (payload.get("rfp_text") or "").strip()
    if rfp_text:
        report("rfp", "running")
        parsed = rfp.extract_rfp(rfp_text)
        results["rfp"] = {k: parsed[k] for k in
                          ("questions", "limits", "deadlines", "scoring")}
        sections = {h: b for h, b in grant["sections"]}
        plan_sections = {h: b for h, b in
                         business_plan.build_business_plan(
                             org, site, pains, forecast, tm_result)["sections"]}
        ctx = {
            "org": org,
            "need_text": sections.get("Statement of Need", ""),
            "project_text": sections.get("Project Description and Goals", ""),
            "budget_text": budget_mod.request_justification(budget, forecast) +
            "\n\nSee the attached line-item Project Budget & Justification.",
            "evaluation_text": sections.get("Evaluation Plan", ""),
            "capacity_text": sections.get("Organizational Capacity", ""),
            "sustainability_text": sections.get("Sustainability", ""),
        }
        rfp_doc = rfp.build_rfp_response(parsed, ctx,
                                         payload.get("funder_name", ""))
        report("rfp", "done",
               f"{len(parsed['questions'])} funder questions answered")

    # 7. Optional Claude tailoring of the prose documents
    context_notes = (
        f"Organization: {org['name']}. Programs: {org['programs']}. "
        f"Grant request: ${forecast['assumptions']['grant_target']:,.0f}. "
        f"County: {payload.get('county', '')}. "
        f"Local need data: {payload.get('local_need', '')[:2000]}")
    ai_used = narrative_ai.available()
    if ai_used:
        report("tailor", "running", "Rewriting prose with Claude")

    def maybe_tailor(doc):
        if not doc:
            return doc
        if ai_used:
            tailored = narrative_ai.tailor_document(doc, context_notes)
            if tailored:
                return tailored
        return doc

    grant = maybe_tailor(grant)
    results["files"] += documents.write_document(
        grant, out_dir, f"{base}_Grant_Narrative")
    if rfp_doc:
        rfp_doc = maybe_tailor(rfp_doc)
        results["files"] += documents.write_document(
            rfp_doc, out_dir, f"{base}_RFP_Response")
    if ai_used:
        report("tailor", "done", "Prose tailored by Claude")
    else:
        report("tailor", "done",
               "Template prose used (install `anthropic` and set "
               "ANTHROPIC_API_KEY for AI-tailored writing)")
    results["ai_tailored"] = ai_used

    # Machine-readable reports
    results["files"] += documents.write_json(
        {k: results[k] for k in
         ("website_review", "ein", "painpoints", "forecast", "budget",
          "readiness", "rfp")
         if k in results} | ({"trademark": tm_result} if tm_result else {}),
        out_dir, f"{base}_Agent_Reports")
    return results


def run_safe(payload, out_dir, progress=None):
    try:
        return {"ok": True, "results": run_pipeline(payload, out_dir, progress)}
    except Exception:
        return {"ok": False, "error": traceback.format_exc()}
