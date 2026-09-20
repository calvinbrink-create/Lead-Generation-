"""
Grows sources/company_slugs.json by searching DuckDuckGo's HTML results for
pages published on each ATS's public board domain. Free, no API key.

Run this weekly, separate from the daily discover+outreach job - it just
maintains the target list that ats_boards.py polls every day.
"""
import json
import re
import time
import requests
from pathlib import Path

SLUG_FILE = Path(__file__).parent / "company_slugs.json"
DDG_HTML = "https://html.duckduckgo.com/html/"
HEADERS = {"User-Agent": "Mozilla/5.0"}

PATTERNS = {
    "greenhouse": re.compile(r"boards\.greenhouse\.io/([a-zA-Z0-9\-_]+)"),
    "lever": re.compile(r"jobs\.lever\.co/([a-zA-Z0-9\-_]+)"),
    "ashby": re.compile(r"jobs\.ashbyhq\.com/([a-zA-Z0-9\-_]+)"),
}

# TODO: expand this list. More/narrower queries = more coverage, but be
# mindful of DuckDuckGo rate limits - the 2s sleep below is deliberate.
SEED_QUERIES = [
    "site:boards.greenhouse.io customer support",
    "site:jobs.lever.co appointment setter",
    "site:jobs.ashbyhq.com call center",
    "site:boards.greenhouse.io back office",
    "site:jobs.lever.co scheduling coordinator",
    "site:jobs.ashbyhq.com customer service representative",
    "site:boards.greenhouse.io data entry",
]

def search_ddg(query):
    r = requests.post(DDG_HTML, data={"q": query}, headers=HEADERS, timeout=20)
    return r.text

def extract_slugs(html):
    return {source: sorted(set(pattern.findall(html))) for source, pattern in PATTERNS.items()}

def grow_slug_file():
    data = json.loads(SLUG_FILE.read_text()) if SLUG_FILE.exists() else {
        "greenhouse": [], "lever": [], "ashby": []
    }
    for source in ("greenhouse", "lever", "ashby"):
        data.setdefault(source, [])

    for q in SEED_QUERIES:
        try:
            html = search_ddg(q)
        except Exception as e:
            print(f"query failed: {q} ({e})")
            continue
        found = extract_slugs(html)
        for source, slugs in found.items():
            for s in slugs:
                if s not in data[source]:
                    data[source].append(s)
        time.sleep(2)  # be polite to DuckDuckGo

    data.pop("note", None)
    SLUG_FILE.write_text(json.dumps(data, indent=2))
    return data

if __name__ == "__main__":
    result = grow_slug_file()
    for k, v in result.items():
        print(k, len(v))
