"""
Given a lead's company name / domain, tries to find a real contact email on
the company's own website. Falls back to a domain guess only when the
source data didn't already include one.
"""
import re
import requests
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0"}
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
CANDIDATE_PAGES = ["", "/contact", "/contact-us", "/about", "/about-us", "/team"]


def guess_domain_from_company(company_name):
    # Crude fallback only. Prefer wiring through a real domain from the
    # source record (e.g. the job posting's own site) wherever one exists -
    # this guess is a last resort and will often be wrong.
    slug = re.sub(r"[^a-z0-9]", "", (company_name or "").lower())
    return f"{slug}.com" if slug else None


def find_email_on_site(domain):
    if not domain:
        return None
    for path in CANDIDATE_PAGES:
        url = f"https://{domain}{path}"
        try:
            r = requests.get(url, headers=HEADERS, timeout=15)
            if r.status_code != 200:
                continue
            soup = BeautifulSoup(r.text, "html.parser")
            matches = EMAIL_RE.findall(soup.get_text(" "))
            for a in soup.find_all("a", href=True):
                if a["href"].startswith("mailto:"):
                    matches.append(a["href"].replace("mailto:", "").split("?")[0])
            matches = [m for m in matches if not m.lower().endswith((".png", ".jpg", ".svg", ".gif"))]
            if matches:
                preferred = [
                    m for m in matches
                    if any(p in m.lower() for p in ["hello", "info", "contact", "ops", "hiring", "careers"])
                ]
                return (preferred or matches)[0]
        except Exception:
            continue
    return None


def enrich_lead(lead):
    domain = lead.get("domain") or guess_domain_from_company(lead.get("company"))
    email = find_email_on_site(domain)
    lead["domain"] = domain
    lead["email"] = email
    return lead
