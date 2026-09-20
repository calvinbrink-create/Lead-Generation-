"""
Pulls UK BPO-relevant notices from the two official OCDS procurement APIs:
Contracts Finder (below-threshold, England-wide) and Find a Tender
(above-threshold, UK-wide). Both are public, no API key, no login.

NOTE: government API response envelopes occasionally change shape. Verify
the actual JSON structure against a live response before relying on this
in production, and adjust the field paths below if they've drifted.
"""
import requests
from config import TENDER_KEYWORDS

HEADERS = {"User-Agent": "Mozilla/5.0"}
CONTRACTS_FINDER_SEARCH = "https://www.contractsfinder.service.gov.uk/Published/Notices/OCDS/Search"
FIND_A_TENDER_FEED = "https://www.find-tender.service.gov.uk/api/1.0/ocdsReleasePackages"


def matches_keywords(text):
    t = (text or "").lower()
    return any(k in t for k in TENDER_KEYWORDS)


def fetch_contracts_finder(keyword_query="call centre OR back office OR outsourcing"):
    params = {"keyword": keyword_query, "size": 100, "stages": "tender"}
    try:
        r = requests.get(CONTRACTS_FINDER_SEARCH, params=params, headers=HEADERS, timeout=30)
        if r.status_code != 200:
            return []
        data = r.json()
    except Exception:
        return []
    out = []
    for release in data.get("releases", []):
        tender = release.get("tender", {})
        title = tender.get("title", "")
        desc = tender.get("description", "")
        if matches_keywords(title) or matches_keywords(desc):
            buyer = release.get("buyer", {})
            out.append({
                "source": "uk_tender_cf",
                "source_ref": release.get("ocid"),
                "company": buyer.get("name"),
                "title": title,
                "location": "UK",
                "url": f"https://www.contractsfinder.service.gov.uk/Notice/{release.get('id', '')}",
                "posted_date": release.get("date"),
            })
    return out


def fetch_find_a_tender():
    try:
        r = requests.get(FIND_A_TENDER_FEED, headers=HEADERS, timeout=30)
        if r.status_code != 200:
            return []
        data = r.json()
    except Exception:
        return []
    out = []
    packages = data.get("releasePackages") or data.get("releases") or []
    for pkg in packages:
        releases = pkg.get("releases", [pkg]) if isinstance(pkg, dict) else []
        for release in releases:
            tender = release.get("tender", {})
            title = tender.get("title", "")
            desc = tender.get("description", "")
            if matches_keywords(title) or matches_keywords(desc):
                buyer = release.get("buyer", {})
                docs = tender.get("documents") or []
                out.append({
                    "source": "uk_tender_fts",
                    "source_ref": release.get("ocid"),
                    "company": buyer.get("name"),
                    "title": title,
                    "location": "UK",
                    "url": docs[0].get("url", "") if docs else "",
                    "posted_date": release.get("date"),
                })
    return out


def fetch_all():
    return fetch_contracts_finder() + fetch_find_a_tender()
