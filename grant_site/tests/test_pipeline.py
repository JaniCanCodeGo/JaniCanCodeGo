"""Offline unit tests for the grant agent (no network required)."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from agent import business_plan, documents, ein, trademark, website_review


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


def test_document_writer(tmp_path):
    doc = {"title": "T", "sections": [("A", "body")]}
    files = documents.write_document(doc, str(tmp_path), "test")
    assert any(f.endswith(".md") for f in files)
    md = open(files[0], encoding="utf-8").read()
    assert "# T" in md and "## A" in md
