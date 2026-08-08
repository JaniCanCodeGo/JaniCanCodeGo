"""Grant opportunity search.

Sources:
  - Grants.gov Search2 public API (federal; no key required)
  - California Grants Portal dataset on data.ca.gov (CKAN datastore API)
Both degrade gracefully to a manual search link on any failure.
"""

import requests

TIMEOUT = 25
GRANTS_GOV_API = "https://api.grants.gov/v1/api/search2"
CA_CKAN_SQL = "https://data.ca.gov/api/3/action/datastore_search"
# Resource ID for the "California Grants Portal" dataset on data.ca.gov.
CA_GRANTS_RESOURCE = "111c8c88-21f6-453c-ae2c-b4785a0624f5"


def search_opportunities(keyword, limit=10):
    """Search federal + California grant opportunities for a keyword."""
    return {
        "keyword": keyword,
        "federal": _search_grants_gov(keyword, limit),
        "california": _search_ca_portal(keyword, limit),
        "manual_links": [
            {"label": "Grants.gov search",
             "url": f"https://grants.gov/search-grants?query={requests.utils.quote(keyword)}"},
            {"label": "California Grants Portal",
             "url": f"https://www.grants.ca.gov/grants/?fwp_search={requests.utils.quote(keyword)}"},
            {"label": "Candid Foundation Directory (library access often free)",
             "url": "https://candid.org/find-funding"},
        ],
    }


def _search_grants_gov(keyword, limit):
    try:
        resp = requests.post(GRANTS_GOV_API, json={
            "keyword": keyword, "oppStatuses": "posted",
            "rows": limit, "startRecordNum": 0,
        }, timeout=TIMEOUT)
        resp.raise_for_status()
        hits = (resp.json().get("data") or {}).get("oppHits", [])
        return {"ok": True, "results": [{
            "number": h.get("number"),
            "title": h.get("title"),
            "agency": h.get("agencyName") or h.get("agency"),
            "open_date": h.get("openDate"),
            "close_date": h.get("closeDate"),
            "url": f"https://grants.gov/search-results-detail/{h.get('id')}",
        } for h in hits[:limit]]}
    except requests.RequestException as exc:
        return {"ok": False, "error": str(exc), "results": []}


def _search_ca_portal(keyword, limit):
    try:
        resp = requests.get(CA_CKAN_SQL, params={
            "resource_id": CA_GRANTS_RESOURCE, "q": keyword, "limit": limit,
        }, timeout=TIMEOUT)
        resp.raise_for_status()
        records = (resp.json().get("result") or {}).get("records", [])
        results = []
        for r in records[:limit]:
            results.append({
                "title": r.get("Title") or r.get("title"),
                "agency": r.get("AgencyDept") or r.get("agency"),
                "close_date": r.get("ApplicationDeadline") or r.get("CloseDate"),
                "amount": r.get("EstAvailFunds") or r.get("EstAmounts"),
                "url": r.get("GrantURL") or r.get("url") or
                "https://www.grants.ca.gov/",
            })
        return {"ok": True, "results": results}
    except requests.RequestException as exc:
        return {"ok": False, "error": str(exc), "results": []}
