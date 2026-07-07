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

from . import business_plan, documents, ein, grant_writer, trademark
from . import website_review


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
    pains = business_plan.research_painpoints(focus, site.get("mission_text"))
    forecast = business_plan.build_forecast(payload.get("forecast") or {})
    results["painpoints"] = pains
    results["forecast"] = forecast

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
    results["files"] += documents.write_document(
        grant, out_dir, f"{base}_Grant_Narrative")
    report("grant", "done", "Grant narrative written")

    # Machine-readable reports
    results["files"] += documents.write_json(
        {k: results[k] for k in
         ("website_review", "ein", "painpoints", "forecast")
         if k in results} | ({"trademark": tm_result} if tm_result else {}),
        out_dir, f"{base}_Agent_Reports")
    return results


def run_safe(payload, out_dir, progress=None):
    try:
        return {"ok": True, "results": run_pipeline(payload, out_dir, progress)}
    except Exception:
        return {"ok": False, "error": traceback.format_exc()}
