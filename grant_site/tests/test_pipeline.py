"""Offline unit tests for the grant agent (no network required)."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from agent import (budget, business_plan, documents, ein, readiness, rfp,
                   trademark, website_review)


def test_forecast_is_deterministic_and_traceable():
    fc = business_plan.build_forecast({"base_revenue": 100000,
                                       "revenue_growth_pct": 10,
                                       "grant_target": 50000})
    assert len(fc["years"]) == 3
    assert fc["years"][0]["grant_revenue"] == 50000
    assert fc["years"][1]["grant_revenue"] == 0
    assert fc["years"][1]["program_revenue"] == 110000
    assert fc["assumptions"]["base_revenue"] == 100000


def test_painpoints_have_sources():
    res = business_plan.research_painpoints(["education", "youth"])
    assert res["painpoints"]
    assert all(p["source"] for p in res["painpoints"])


def test_business_plan_sections():
    fc = business_plan.build_forecast({})
    pains = business_plan.research_painpoints(["education"])
    plan = business_plan.build_business_plan(
        {"name": "Test Org", "mission": "Serve kids."},
        {"mission_text": "", "url": "https://example.org"}, pains, fc)
    titles = [t for t, _ in plan["sections"]]
    assert "Executive Summary" in titles
    assert "Market Need — Documented Pain Points" in titles
    assert "Three-Year Financial Forecast" in titles


def test_ein_normalize():
    assert ein.normalize_ein("12-3456789") == "12-3456789"
    assert ein.normalize_ein("123456789") == "12-3456789"
    assert ein.normalize_ein("123") is None


def test_ss4_worksheet_never_stores_ssn():
    ws = ein.build_ss4_worksheet({"legal_name": "Test Org"})
    assert "never store" in ws["lines"]["7b_responsible_party_ssn_itin"]
    assert ws["lines"]["1_legal_name"] == "Test Org"


def test_trademark_variants_and_classes():
    res = trademark.search_trademark("Kids Kode Klub",
                                     "education software t-shirts")
    assert "kids kode klub" in res["variants"]
    assert res["manual_search_links"]
    assert set(res["suggested_classes"]) >= {25, 41}


def test_teas_packet_fee_math():
    app = trademark.build_teas_application(
        "Kids Kode Klub", {"name": "Test Org"},
        "education services and apparel", use_in_commerce=True)
    assert app["fees"]["total_usd"] == 250 * len(app["international_classes"])
    assert "teas.uspto.gov" in app["file_at"]


def test_accessibility_audit_flags_missing_alt():
    html = "<html><head><title>x</title></head><body><img src='a.png'></body></html>"
    audit = website_review.audit_accessibility(html)
    by_id = {c["id"]: c for c in audit["checks"]}
    assert by_id["img-alt"]["status"] == "review"
    assert by_id["title"]["status"] == "pass"
    assert by_id["html-lang"]["status"] == "review"


def test_ss4_document_is_filled_and_has_filing_guide():
    ws = ein.build_ss4_worksheet({
        "legal_name": "Test Org", "state": "CA",
        "responsible_party": "Jane Doe",
        "mailing_address": "1 Main St, Sacramento, CA 95814",
    })
    doc = ein.build_ss4_document(ws, "Test Org")
    text = "\n".join(body for _, body in doc["sections"])
    assert "Jane Doe" in text
    assert "1 Main St" in text
    assert "EIN Assistant" in text          # filing walkthrough present
    assert "irs.gov" in text
    assert "Line 7a" in text                # line-by-line mapping


def test_teas_document_is_filled_and_has_filing_guide():
    search = trademark.search_trademark("Kids Kode Klub", "apparel education")
    app = trademark.build_teas_application(
        "Kids Kode Klub", {"name": "Test Org", "email": "a@b.org",
                           "entity_type": "Nonprofit corporation"},
        "education services and apparel")
    doc = trademark.build_teas_document(app, search)
    text = "\n".join(body for _, body in doc["sections"])
    assert "Kids Kode Klub" in text
    assert "teas.uspto.gov" in text
    assert "Signature screen" in text       # filing walkthrough present
    assert "Clearance summary" in doc["sections"][0][0]


SAMPLE_RFP = """
Community Youth Grant Program — Notice of Funding Opportunity
Applications are due September 15, 2026. Narratives are limited to
no more than 10 pages.

Application Questions:
1. Describe the community need your project addresses, using local data.
2. What are your project's goals, objectives, and key activities?
3. How will you evaluate outcomes and measure success?
4. Provide a budget narrative explaining how funds will be spent.
5. How will the program be sustained beyond the grant period?

Scoring:
Need Statement (30 points)
Project Design (30 points)
Evaluation (20 points)
Budget (20 points)
"""


def test_rfp_extraction():
    parsed = rfp.extract_rfp(SAMPLE_RFP)
    assert len(parsed["questions"]) >= 5
    topics = {q["topic"] for q in parsed["questions"]}
    assert {"need", "evaluation", "budget", "sustainability"} <= topics
    assert parsed["limits"][0] == {"limit": 10, "unit": "pages"}
    assert any("September 15, 2026" in d for d in parsed["deadlines"])
    assert {"criterion": "Need Statement", "points": 30} in [
        {"criterion": s["criterion"], "points": s["points"]}
        for s in parsed["scoring"]]


def test_rfp_response_answers_every_question():
    parsed = rfp.extract_rfp(SAMPLE_RFP)
    ctx = {"org": {"name": "Test Org"}, "need_text": "NEED",
           "project_text": "PROJECT", "budget_text": "BUDGET",
           "evaluation_text": "EVAL", "capacity_text": "CAP",
           "sustainability_text": "SUSTAIN"}
    doc = rfp.build_rfp_response(parsed, ctx, "Test Funder")
    numbered = [s for s in doc["sections"] if s[0][0].isdigit()]
    assert len(numbered) == len(parsed["questions"])
    bodies = " ".join(b for _, b in numbered)
    assert "NEED" in bodies and "EVAL" in bodies and "BUDGET" in bodies


def test_budget_auto_allocation_sums_to_request():
    b = budget.build_budget(50000)
    assert abs(b["total"] - 50000) < 1
    assert b["allocated_by_default_pcts"]
    assert any("Indirect" in l["category"] for l in b["lines"])


def test_budget_user_items_and_gap_statement():
    b = budget.build_budget(10000, items={"personnel": 6000, "supplies": 2000},
                            fringe_pct=20, indirect_pct=10, match_amount=2500)
    fringe = 6000 * 0.20
    assert b["total"] == round((8000 + fringe) * 1.10, 2)
    fc = business_plan.build_forecast({"base_revenue": 10000,
                                       "base_expenses": 60000,
                                       "grant_target": 10000})
    text = budget.request_justification(b, fc)
    assert "deficit" in text
    assert "$2,500" in text


def test_readiness_flags_missing_ein_and_counts_have_items():
    r = readiness.assess_readiness({"found": False, "matches": []},
                                   have_items=["sam_gov", "board_list"])
    by_key = {i["key"]: i for i in r["items"]}
    assert by_key["ein"]["status"] == "action_needed"
    assert by_key["sam_gov"]["status"] == "ready"
    assert by_key["board_list"]["status"] == "ready"
    assert by_key["form_990"]["status"] == "needed"


def test_local_need_leads_painpoints():
    res = business_plan.research_painpoints(
        ["education"], local_need="72% of local students qualify for "
        "free/reduced lunch (DataQuest 2025)", county="Sacramento County")
    assert "72%" in res["painpoints"][0]["pain"]
    assert res["local_data_sources"]


def test_document_writer(tmp_path):
    doc = {"title": "T", "sections": [("A", "body")]}
    files = documents.write_document(doc, str(tmp_path), "test")
    assert any(f.endswith(".md") for f in files)
    md = open(files[0], encoding="utf-8").read()
    assert "# T" in md and "## A" in md
