"""Stage 1 — Website review.

Fetches the organization's website and extracts the raw material the rest of
the pipeline needs (mission language, programs, contact info), plus a quick
Section 508 / WCAG 2.1 AA spot-check of the org's own site so the grant can
truthfully describe the applicant's accessibility posture.
"""

import re
from urllib.parse import urljoin, urlparse

import requests

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

USER_AGENT = "GrantStudioAgent/1.0 (+grant application research; contact site owner)"
TIMEOUT = 20

EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
PHONE_RE = re.compile(r"\(?\b\d{3}\)?[-.\s]\d{3}[-.\s]\d{4}\b")
EIN_RE = re.compile(r"\b(\d{2})-(\d{7})\b")


def _normalize_url(url):
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    return url


def review_website(url):
    """Fetch and analyze an organization website.

    Returns a dict with keys: url, ok, title, description, headings,
    mission_text, emails, phones, ein_candidates, social_links,
    accessibility (spot-check results), pages_scanned, error.
    """
    result = {
        "url": None, "ok": False, "title": "", "description": "",
        "headings": [], "mission_text": "", "emails": [], "phones": [],
        "ein_candidates": [], "social_links": [], "accessibility": {},
        "pages_scanned": [], "error": None,
    }
    if not url or not url.strip():
        result["error"] = "No website URL provided."
        return result

    url = _normalize_url(url.strip())
    result["url"] = url
    try:
        resp = requests.get(url, headers={"User-Agent": USER_AGENT},
                            timeout=TIMEOUT)
        resp.raise_for_status()
    except requests.RequestException as exc:
        result["error"] = f"Could not fetch website: {exc}"
        return result

    html = resp.text
    result["ok"] = True
    result["pages_scanned"].append(url)
    result["emails"] = sorted(set(EMAIL_RE.findall(html)))[:10]
    result["phones"] = sorted(set(PHONE_RE.findall(html)))[:5]
    result["ein_candidates"] = sorted(
        {f"{a}-{b}" for a, b in EIN_RE.findall(html)})[:5]

    if HAS_BS4:
        _analyze_soup(html, url, result)
    else:
        _analyze_regex(html, result)

    result["accessibility"] = audit_accessibility(html)
    return result


def _analyze_soup(html, base_url, result):
    soup = BeautifulSoup(html, "html.parser")
    if soup.title and soup.title.string:
        result["title"] = soup.title.string.strip()
    meta = soup.find("meta", attrs={"name": "description"})
    if meta and meta.get("content"):
        result["description"] = meta["content"].strip()

    for tag in soup.find_all(["h1", "h2", "h3"]):
        text = tag.get_text(" ", strip=True)
        if text:
            result["headings"].append({"level": tag.name, "text": text[:200]})
    result["headings"] = result["headings"][:40]

    # Mission language: paragraphs near words like mission/vision/about.
    mission_parts = []
    for p in soup.find_all("p"):
        text = p.get_text(" ", strip=True)
        if len(text) > 60 and re.search(
                r"\b(mission|vision|we believe|our goal|founded|serve|serving|"
                r"community|empower)\b", text, re.I):
            mission_parts.append(text)
        if len(mission_parts) >= 5:
            break
    result["mission_text"] = "\n\n".join(mission_parts)[:4000]

    for a in soup.find_all("a", href=True):
        href = urljoin(base_url, a["href"])
        host = urlparse(href).netloc.lower()
        if any(s in host for s in ("facebook.", "instagram.", "linkedin.",
                                   "twitter.", "x.com", "youtube.",
                                   "tiktok.")):
            result["social_links"].append(href)
    result["social_links"] = sorted(set(result["social_links"]))[:10]


def _analyze_regex(html, result):
    m = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
    if m:
        result["title"] = re.sub(r"\s+", " ", m.group(1)).strip()
    m = re.search(
        r'<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']*)',
        html, re.I)
    if m:
        result["description"] = m.group(1).strip()
    for level, text in re.findall(r"<(h[1-3])[^>]*>(.*?)</\1>",
                                  html, re.I | re.S)[:40]:
        clean = re.sub(r"<[^>]+>", " ", text)
        clean = re.sub(r"\s+", " ", clean).strip()
        if clean:
            result["headings"].append({"level": level.lower(),
                                       "text": clean[:200]})


def audit_accessibility(html):
    """Lightweight Section 508 / WCAG 2.1 AA spot-check of raw HTML.

    This is a heuristic pre-scan, not a substitute for a full audit with
    axe-core / ANDI and manual testing.
    """
    checks = []

    def check(cid, wcag, passed, detail):
        checks.append({"id": cid, "wcag": wcag,
                       "status": "pass" if passed else "review",
                       "detail": detail})

    imgs = re.findall(r"<img\b[^>]*>", html, re.I)
    missing_alt = [i for i in imgs if not re.search(r"\balt\s*=", i, re.I)]
    check("img-alt", "1.1.1", not missing_alt,
          f"{len(missing_alt)} of {len(imgs)} <img> tags missing alt attributes.")

    check("html-lang", "3.1.1",
          bool(re.search(r"<html[^>]+lang\s*=", html, re.I)),
          "Document language attribute (<html lang>).")

    inputs = re.findall(r"<input\b[^>]*>", html, re.I)
    labelable = [i for i in inputs if not re.search(
        r'type\s*=\s*["\']?(hidden|submit|button|image)', i, re.I)]
    labels = len(re.findall(r"<label\b", html, re.I))
    aria_labels = len([i for i in labelable
                       if re.search(r"aria-label(ledby)?\s*=", i, re.I)])
    check("form-labels", "1.3.1/4.1.2",
          not labelable or (labels + aria_labels) >= len(labelable),
          f"{len(labelable)} form inputs vs {labels} <label> + "
          f"{aria_labels} aria-label(s).")

    check("skip-link", "2.4.1",
          bool(re.search(r'href\s*=\s*["\']#(main|content|skip)', html, re.I)),
          "Skip-to-content link present.")

    check("headings", "1.3.1/2.4.6",
          bool(re.search(r"<h1\b", html, re.I)),
          "At least one <h1> heading.")

    check("title", "2.4.2",
          bool(re.search(r"<title[^>]*>\s*\S", html, re.I | re.S)),
          "Page has a non-empty <title>.")

    passed = sum(1 for c in checks if c["status"] == "pass")
    return {"checks": checks, "passed": passed, "total": len(checks),
            "note": ("Heuristic spot-check only. A conformant Section 508 "
                     "audit requires automated tooling (axe-core, ANDI) plus "
                     "manual keyboard and screen-reader testing.")}
