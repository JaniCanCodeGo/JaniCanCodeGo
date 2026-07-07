"""Stage 4 — Trademark clearance search and application preparation.

Search strategy (in order):
  1. USPTO TSDR API (https://tsdrapi.uspto.gov) — exact serial/registration
     status checks; needs a free USPTO API key in env var USPTO_TSDR_API_KEY.
  2. Marker API / RapidAPI USPTO keyword search if USPTO_TM_API_KEY is set.
  3. Always: deterministic knockout analysis (identical / phonetic-similar
     name variants) plus direct deep links into the USPTO's public search
     (tmsearch.uspto.gov) and state registries for human confirmation.

Filing: the USPTO does not expose a public API for submitting TEAS trademark
applications — filing happens at https://teas.uspto.gov and must be signed by
the applicant or their attorney. The agent therefore produces a complete,
ready-to-enter TEAS Plus application packet (mark, owner, filing basis,
international classes, specimen checklist, fee calculation).
"""

import os
import re

import requests

TIMEOUT = 20
TSDR_STATUS = "https://tsdrapi.uspto.gov/ts/cd/casestatus/sn{serial}/info.json"
TESS_SEARCH = ("https://tmsearch.uspto.gov/search/search-information?"
               "query={query}")

# TEAS Plus fee per class as of 2025 (base application fee).
TEAS_PLUS_FEE_PER_CLASS = 250

NICE_CLASS_HINTS = {
    "education": 41, "training": 41, "tutoring": 41, "workshops": 41,
    "software": 42, "saas": 42, "website": 42, "app": 9,
    "clothing": 25, "apparel": 25, "t-shirt": 25, "merchandise": 25,
    "hats": 25, "consulting": 35, "advertising": 35, "retail": 35,
    "charitable": 36, "fundraising": 36, "grants": 36,
    "books": 16, "printed": 16, "toys": 28, "food": 43,
}


def _phonetic_variants(name):
    """Generate simple knockout variants (sound-alikes/spelling variants)."""
    base = re.sub(r"[^a-z0-9 ]", "", name.lower()).strip()
    variants = {base}
    swaps = [("ph", "f"), ("c", "k"), ("k", "c"), ("z", "s"), ("s", "z"),
             ("x", "ks"), ("qu", "kw"), ("ee", "i"), ("y", "i"),
             ("ie", "y"), ("gh", "g")]
    for a, b in swaps:
        if a in base:
            variants.add(base.replace(a, b))
    variants.add(base.replace(" ", ""))
    variants.add(re.sub(r"s\b", "", base))          # drop plural
    variants.add(re.sub(r"\b(the|a|an)\b\s*", "", base).strip())
    return sorted(v for v in variants if v)


def suggest_classes(goods_description):
    """Suggest Nice international classes from a goods/services description."""
    text = (goods_description or "").lower()
    classes = sorted({cls for kw, cls in NICE_CLASS_HINTS.items()
                      if kw in text})
    return classes or [35]


def search_trademark(name, goods_description=""):
    """Run a trademark knockout search for a proposed name."""
    result = {
        "query": name, "variants": _phonetic_variants(name),
        "api_hits": [], "api_source": None, "api_error": None,
        "manual_search_links": [], "suggested_classes":
            suggest_classes(goods_description),
        "disclaimer": (
            "Automated knockout search only — likelihood-of-confusion "
            "analysis under 15 U.S.C. §1052(d) considers sound, appearance, "
            "meaning, and relatedness of goods. Confirm results in the "
            "USPTO search system and consider a trademark attorney before "
            "adopting or filing a mark."),
    }

    tm_key = os.environ.get("USPTO_TM_API_KEY")
    if tm_key:
        try:
            resp = requests.get(
                "https://uspto-trademark.p.rapidapi.com/v1/trademarkSearch/"
                f"{requests.utils.quote(name)}/active",
                headers={"x-rapidapi-key": tm_key,
                         "x-rapidapi-host": "uspto-trademark.p.rapidapi.com"},
                timeout=TIMEOUT)
            if resp.ok:
                items = resp.json().get("items", [])[:15]
                result["api_hits"] = [{
                    "keyword": i.get("keyword"),
                    "serial": i.get("serial_number"),
                    "status": i.get("status_label"),
                    "owner": (i.get("owners") or [{}])[0].get("name", ""),
                    "classes": i.get("international_class_codes", []),
                } for i in items]
                result["api_source"] = "uspto-trademark (RapidAPI)"
            else:
                result["api_error"] = f"Trademark API HTTP {resp.status_code}"
        except requests.RequestException as exc:
            result["api_error"] = f"Trademark API unavailable: {exc}"
    else:
        result["api_error"] = ("No trademark search API key configured "
                               "(set USPTO_TM_API_KEY); use the manual "
                               "search links below.")

    quoted = requests.utils.quote(name)
    result["manual_search_links"] = [
        {"label": "USPTO Trademark Search (exact and similar marks)",
         "url": TESS_SEARCH.format(query=quoted)},
        {"label": "USPTO TSDR (status of a specific serial number)",
         "url": "https://tsdr.uspto.gov/"},
        {"label": "California Secretary of State trademark search",
         "url": "https://bizfileonline.sos.ca.gov/search/trademarks"},
        {"label": "Google web search for similar concepts",
         "url": f"https://www.google.com/search?q=%22{quoted}%22"},
        {"label": "Domain availability",
         "url": f"https://www.namecheap.com/domains/registration/results/"
                f"?domain={quoted}"},
    ]
    return result


def check_serial_status(serial):
    """Check a specific USPTO serial number via TSDR (needs free API key)."""
    key = os.environ.get("USPTO_TSDR_API_KEY")
    if not key:
        return {"error": "Set USPTO_TSDR_API_KEY (free at developer.uspto.gov)"}
    try:
        resp = requests.get(TSDR_STATUS.format(serial=serial),
                            headers={"USPTO-API-KEY": key}, timeout=TIMEOUT)
        resp.raise_for_status()
        return resp.json()
    except requests.RequestException as exc:
        return {"error": str(exc)}


def build_teas_application(mark, owner, goods_description,
                           classes=None, use_in_commerce=False):
    """Prepare a complete TEAS Plus application packet.

    The packet contains every field the TEAS Plus form asks for, ready to be
    entered and signed at https://teas.uspto.gov by the owner or attorney.
    """
    classes = classes or suggest_classes(goods_description)
    return {
        "file_at": "https://teas.uspto.gov/forms/bas/",
        "form": "TEAS Plus (initial application)",
        "mark": {
            "literal_element": mark,
            "mark_type": "Standard character mark (protects the wording in "
                         "any font/style/color)",
        },
        "owner": owner,
        "filing_basis": ("1(a) Use in commerce — specimen required"
                         if use_in_commerce else
                         "1(b) Intent to use — Statement of Use filed later "
                         "(additional fee per class)"),
        "international_classes": classes,
        "goods_and_services": goods_description,
        "identification_note": ("TEAS Plus requires descriptions selected "
                                "from the USPTO ID Manual: "
                                "https://idm-tmng.uspto.gov/id-master-list-public.html"),
        "specimen_checklist": [
            "Goods (e.g. merchandise): photo of the mark on the product, "
            "packaging, or a point-of-sale web page with a buy button.",
            "Services: advertising/marketing material or website screenshot "
            "showing the mark used with the services.",
        ] if use_in_commerce else [
            "No specimen needed at filing under 1(b); required with the "
            "Statement of Use."],
        "fees": {
            "per_class_usd": TEAS_PLUS_FEE_PER_CLASS,
            "classes": len(classes),
            "total_usd": TEAS_PLUS_FEE_PER_CLASS * len(classes),
            "note": "Verify current fees at uspto.gov/trademark/trademark-fee-information",
        },
        "signature": ("Must be personally signed at teas.uspto.gov by the "
                      "applicant or an authorized attorney — the USPTO has "
                      "no third-party filing API, so this packet is "
                      "prepared-to-file, not auto-filed."),
        "timeline": [
            "~1 week: serial number issued and application viewable in TSDR",
            "~8-10 months: examining attorney review (office actions possible)",
            "~1 month publication for opposition",
            "Registration (1(a)) or Notice of Allowance (1(b))",
        ],
    }
