"""
Fetches open roles from public, keyless ATS JSON APIs and filters to
BPO-relevant roles (support, scheduling, back-office, data entry, etc.)
Each open role at a company is itself a lead - it means someone there is
currently paying to staff exactly what Auto BPO replaces.
"""
import json
import requests
from pathlib import Path
from config import ROLE_KEYWORDS

SLUG_FILE = Path(__file__).parent / "company_slugs.json"
HEADERS = {"User-Agent": "Mozilla/5.0"}


def load_slugs():
    data = json.loads(SLUG_FILE.read_text())
    data.pop("note", None)
    return data


def matches_keywords(title):
    t = (title or "").lower()
    return any(k in t for k in ROLE_KEYWORDS)


def fetch_greenhouse(slug):
    url = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=false"
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code != 200:
            return []
        jobs = r.json().get("jobs", [])
    except Exception:
        return []
    out = []
    for j in jobs:
        if matches_keywords(j.get("title", "")):
            out.append({
                "source": "greenhouse",
                "source_ref": str(j["id"]),
                "company": slug,
                "title": j.get("title"),
                "location": (j.get("location") or {}).get("name"),
                "url": j.get("absolute_url"),
                "posted_date": j.get("updated_at"),
            })
    return out


def fetch_lever(slug):
    url = f"https://api.lever.co/v0/postings/{slug}?mode=json"
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code != 200:
            return []
        jobs = r.json()
    except Exception:
        return []
    out = []
    for j in jobs:
        title = j.get("text", "")
        if matches_keywords(title):
            cats = j.get("categories", {})
            out.append({
                "source": "lever",
                "source_ref": j.get("id"),
                "company": slug,
                "title": title,
                "location": cats.get("location"),
                "url": j.get("hostedUrl"),
                "posted_date": j.get("createdAt"),
            })
    return out


def fetch_ashby(slug):
    url = f"https://api.ashbyhq.com/posting-api/job-board/{slug}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code != 200:
            return []
        jobs = r.json().get("jobs", [])
    except Exception:
        return []
    out = []
    for j in jobs:
        title = j.get("title", "")
        if matches_keywords(title):
            out.append({
                "source": "ashby",
                "source_ref": j.get("id"),
                "company": slug,
                "title": title,
                "location": j.get("location"),
                "url": j.get("jobUrl") or j.get("applyUrl"),
                "posted_date": j.get("publishedAt"),
            })
    return out


def fetch_all():
    slugs = load_slugs()
    results = []
    for slug in slugs.get("greenhouse", []):
        results.extend(fetch_greenhouse(slug))
    for slug in slugs.get("lever", []):
        results.extend(fetch_lever(slug))
    for slug in slugs.get("ashby", []):
        results.extend(fetch_ashby(slug))
    return results
