import urllib.request, re
from bs4 import BeautifulSoup

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml",
}

def fetch(url, label=""):
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as r:
            return r.read().decode("utf-8", errors="ignore")
    except Exception as e:
        print(f"  FAILED ({label}): {e}")
        return ""

# Test 1: CDE school detail for Marysville Charter Academy (CDS 57727390130254)
print("=== CDE SCHOOL DETAIL FIELDS ===")
html = fetch("https://www.cde.ca.gov/schooldirectory/details?cdscode=57727390130254", "school detail")
if html:
    soup = BeautifulSoup(html, "html.parser")
    for row in soup.find_all("tr"):
        cells = row.find_all(["th", "td"])
        if len(cells) >= 2:
            lbl = cells[0].get_text(" ", strip=True)
            val = cells[1].get_text(" ", strip=True)
            if lbl and val:
                print(f"  [{lbl}] => [{val[:80]}]")
else:
    # Try name search
    print("  Trying name search...")
    import urllib.parse
    q = urllib.parse.quote_plus("Marysville Charter Academy for the Arts")
    html2 = fetch(f"https://www.cde.ca.gov/schooldirectory/results?searchtext={q}&searchtype=S", "name search")
    if html2:
        soup2 = BeautifulSoup(html2, "html.parser")
        for a in soup2.find_all("a", href=True)[:10]:
            if "details" in a["href"]:
                print(f"  Link: {a.get_text(strip=True)} => {a['href']}")

print()

# Test 2: COE leads page structure
print("=== COE LEADS PAGE ===")
html3 = fetch("https://www.cde.ca.gov/ta/cr/caisleads.asp", "COE leads")
if html3:
    soup3 = BeautifulSoup(html3, "html.parser")
    for i, row in enumerate(soup3.find_all("tr")[:50]):
        cells = row.find_all(["th", "td"])
        row_text = " | ".join(c.get_text(" ", strip=True) for c in cells)
        if row_text.strip():
            print(f"  Row {i:02d}: {row_text[:140]}")
else:
    print("  Page not reachable")
