"""Stage 2 — EIN lookup and (if none exists) EIN application preparation.

Lookup uses the ProPublica Nonprofit Explorer public API (IRS BMF/990 data,
no API key required). The IRS does not offer an API for issuing EINs, so when
no EIN is found the agent prepares a completed Form SS-4 worksheet and a
direct link to the IRS online EIN Assistant, where an authorized officer must
submit it (the IRS requires a responsible party with a valid SSN/ITIN to
apply; a third-party designee may assist but the applicant signs).
"""

import re

import requests

PROPUBLICA_SEARCH = "https://projects.propublica.org/nonprofits/api/v2/search.json"
PROPUBLICA_ORG = "https://projects.propublica.org/nonprofits/api/v2/organizations/{ein}.json"
IRS_EIN_ASSISTANT = ("https://www.irs.gov/businesses/small-businesses-self-"
                     "employed/apply-for-an-employer-identification-number-"
                     "ein-online")
TIMEOUT = 20


def normalize_ein(ein):
    digits = re.sub(r"\D", "", ein or "")
    if len(digits) != 9:
        return None
    return f"{digits[:2]}-{digits[2:]}"


def lookup_ein(org_name, state=None, known_ein=None):
    """Look up an organization's EIN.

    Returns {found, ein, matches, source, error, application} where
    `application` is a prefilled SS-4 worksheet when nothing is found.
    """
    result = {"found": False, "ein": None, "matches": [], "source": None,
              "error": None, "application": None}

    ein = normalize_ein(known_ein) if known_ein else None
    if ein:
        try:
            resp = requests.get(
                PROPUBLICA_ORG.format(ein=ein.replace("-", "")),
                timeout=TIMEOUT)
            if resp.ok:
                org = resp.json().get("organization") or {}
                result.update(found=True, ein=ein,
                              source="propublica-nonprofit-explorer")
                result["matches"] = [_match_from_org(org)]
                return result
        except requests.RequestException as exc:
            result["error"] = f"EIN verification lookup failed: {exc}"
        # Even if verification failed, trust the user-supplied EIN format.
        result.update(found=True, ein=ein, source="user-provided")
        return result

    if not org_name:
        result["error"] = "No organization name provided for EIN search."
        result["application"] = build_ss4_worksheet({})
        return result

    try:
        params = {"q": org_name}
        if state:
            params["state[id]"] = state.upper()
        resp = requests.get(PROPUBLICA_SEARCH, params=params, timeout=TIMEOUT)
        resp.raise_for_status()
        orgs = resp.json().get("organizations", [])
    except requests.RequestException as exc:
        result["error"] = (f"EIN search unavailable ({exc}). If the "
                           "organization is tax-exempt, try IRS Tax Exempt "
                           "Organization Search at "
                           "https://apps.irs.gov/app/eos/")
        orgs = []

    for org in orgs[:8]:
        result["matches"].append(_match_from_org(org))

    exact = [m for m in result["matches"]
             if m["name"].lower() == org_name.lower()]
    if exact:
        result.update(found=True, ein=exact[0]["ein"],
                      source="propublica-nonprofit-explorer")
    elif len(result["matches"]) == 1:
        result.update(found=True, ein=result["matches"][0]["ein"],
                      source="propublica-nonprofit-explorer")

    if not result["found"]:
        result["application"] = build_ss4_worksheet(
            {"legal_name": org_name, "state": state})
    return result


def _match_from_org(org):
    ein_raw = str(org.get("ein", ""))
    return {
        "ein": normalize_ein(ein_raw) or ein_raw,
        "name": org.get("name", ""),
        "city": org.get("city", ""),
        "state": org.get("state", ""),
        "ntee_code": org.get("ntee_code", ""),
        "subsection": org.get("subseccd", org.get("subsection_code", "")),
    }


def build_ss4_worksheet(info):
    """Prefilled IRS Form SS-4 worksheet for an EIN application.

    The IRS issues EINs only via its own channels (online assistant, fax, or
    mail); there is no third-party filing API. This worksheet maps our data
    onto the SS-4 line numbers so the responsible party can complete the
    online assistant in one sitting.
    """
    return {
        "form": "IRS Form SS-4 (Application for Employer Identification Number)",
        "apply_online": IRS_EIN_ASSISTANT,
        "notes": [
            "The online EIN Assistant issues the EIN immediately and is free.",
            "The responsible party must be an individual with a valid SSN or "
            "ITIN; the application must be signed/submitted by that person "
            "or with their authorization (Third Party Designee, lines 18+).",
            "Nonprofits: obtaining an EIN does NOT grant tax-exempt status; "
            "file Form 1023/1023-EZ afterwards.",
        ],
        "lines": {
            "1_legal_name": info.get("legal_name", ""),
            "2_trade_name_dba": info.get("dba", ""),
            "3_care_of": info.get("care_of", ""),
            "4a_4b_mailing_address": info.get("mailing_address", ""),
            "5a_5b_street_address": info.get("street_address", ""),
            "6_county_state": info.get("state", ""),
            "7a_responsible_party": info.get("responsible_party", ""),
            "7b_responsible_party_ssn_itin": "(entered directly with the IRS; "
                                             "never store this in the app)",
            "9a_type_of_entity": info.get("entity_type",
                                          "e.g. Corporation, Nonprofit, LLC"),
            "10_reason_for_applying": info.get(
                "reason", "Started new business / banking purpose / hired "
                          "employees / applying for grants"),
            "11_date_business_started": info.get("start_date", ""),
            "12_closing_month": info.get("fiscal_year_end", "December"),
            "16_principal_activity": info.get("activity", ""),
        },
    }
