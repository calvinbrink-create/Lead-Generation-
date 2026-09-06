#!/usr/bin/env python3
"""
Growth Supply House — Automated Lead Generator
Multi-engine search (Google / Bing / DuckDuckGo) -> site scrape -> score -> Excel

Free-stack, no paid APIs. Runs via GitHub Actions on a daily schedule.
"""

import argparse
import concurrent.futures
import random
import re
import sqlite3
import urllib.parse
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

NICHES = ["plumber", "HVAC contractor", "roofing contractor", "roofer", "plumbing company"]

CITIES = [
    "Dallas TX", "Houston TX", "San Antonio TX", "Austin TX", "Phoenix AZ",
    "Atlanta GA", "Charlotte NC", "Raleigh NC", "Tampa FL", "Orlando FL",
    "Jacksonville FL", "Columbus OH", "Indianapolis IN", "Nashville TN",
    "Memphis TN", "Louisville KY", "Kansas City MO", "Denver CO",
    "Oklahoma City OK", "Tulsa OK", "Fort Worth TX", "El Paso TX",
    "Sacramento CA", "Fresno CA", "Las Vegas NV", "Albuquerque NM",
    "Tucson AZ", "Mesa AZ", "Omaha NE", "Wichita KS", "Baton Rouge LA",
]

TARGET_LEADS = 200
MAX_WORKERS_SEARCH = 8
MAX_WORKERS_SCRAPE = 12
REQUEST_TIMEOUT = 10
MIN_DELAY = 1.0
MAX_DELAY = 2.5

DB_PATH = Path("seen_domains.sqlite3")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0 Safari/537.36 Edg/123.0",
]

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
BAD_EMAIL_HINTS = ("example.com", "yourdomain", "sentry.io", "wixpress.com", ".png", ".jpg", ".gif")
OWNER_HINTS = ("owner", "founder", "president", "ceo", "president/owner", "operations manager")

SKIP_DOMAINS = (
    "facebook.com", "yelp.com", "angi.com", "homeadvisor.com", "bbb.org",
    "yellowpages.com", "thumbtack.com", "instagram.com", "linkedin.com",
    "youtube.com", "pinterest.com", "mapquest.com", "google.com", "bing.com",
)


def rand_headers():
    return {"User-Agent": random.choice(USER_AGENTS), "Accept-Language": "en-US,en;q=0.9"}


def polite_sleep():
    import time
    time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))


def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("CREATE TABLE IF NOT EXISTS seen (domain TEXT PRIMARY KEY, first_seen TEXT)")
    conn.commit()
    return conn


def already_seen(conn, domain):
    return conn.execute("SELECT 1 FROM seen WHERE domain = ?", (domain,)).fetchone() is not None


def mark_seen(conn, domain):
    conn.execute(
        "INSERT OR IGNORE INTO seen (domain, first_seen) VALUES (?, ?)",
        (domain, datetime.utcnow().isoformat()),
    )
    conn.commit()


def search_duckduckgo(query, max_results=10):
    url = "https://html.duckduckgo.com/html/"
    try:
        r = requests.post(url, data={"q": query}, headers=rand_headers(), timeout=REQUEST_TIMEOUT)
        soup = BeautifulSoup(r.text, "html.parser")
        return [a.get("href") for a in soup.select("a.result__a")[:max_results] if a.get("href")]
    except Exception as e:
        print(f"  [ddg] error for '{query}': {e}")
        return []


def search_bing(query, max_results=10):
    url = "https://www.bing.com/search"
    try:
        r = requests.get(url, params={"q": query, "count": max_results}, headers=rand_headers(), timeout=REQUEST_TIMEOUT)
        soup = BeautifulSoup(r.text, "html.parser")
        return [li.get("href") for li in soup.select("li.b_algo h2 a")[:max_results] if li.get("href")]
    except Exception as e:
        print(f"  [bing] error for '{query}': {e}")
        return []


def search_google(query, max_results=10):
    url = "https://www.google.com/search"
    try:
        r = requests.get(url, params={"q": query, "num": max_results}, headers=rand_headers(), timeout=REQUEST_TIMEOUT)
        if "captcha" in r.text.lower() or r.status_code != 200:
            return []
        soup = BeautifulSoup(r.text, "html.parser")
        links = []
        for a in soup.select("a"):
            href = a.get("href", "")
            if href.startswith("/url?q="):
                real = urllib.parse.parse_qs(urllib.parse.urlparse(href).query).get("q")
                if real:
                    links.append(real[0])
            elif href.startswith("http") and "google.com" not in href:
                links.append(href)
        return links[:max_results]
    except Exception as e:
        print(f"  [google] error for '{query}': {e}")
        return []


SEARCH_ENGINES = [search_duckduckgo, search_bing, search_google]


def clean_domain(url):
    try:
        return urllib.parse.urlparse(url).netloc.lower().replace("www.", "")
    except Exception:
        return None


def run_search_task(query):
    results = []
    for engine in SEARCH_ENGINES:
        results.extend(engine(query))
        polite_sleep()
    seen_domains = set()
    candidates = []
    for url in results:
        domain = clean_domain(url)
        if not domain or domain in seen_domains:
            continue
        if any(skip in domain for skip in SKIP_DOMAINS):
            continue
        seen_domains.add(domain)
        candidates.append((domain, url))
    return candidates


@dataclass
class Lead:
    business: str = ""
    domain: str = ""
    url: str = ""
    owner_name: str = ""
    email: str = ""
    phone: str = ""
    score: int = 0
    city_query: str = ""
    engine_source: str = ""
    date_found: str = field(default_factory=lambda: datetime.utcnow().strftime("%Y-%m-%d"))


PHONE_RE = re.compile(r"(\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4})")


def extract_emails(text):
    return [e for e in EMAIL_RE.findall(text) if not any(bad in e.lower() for bad in BAD_EMAIL_HINTS)]


def guess_owner_name(soup):
    text = soup.get_text(" ", strip=True)
    for hint in OWNER_HINTS:
        idx = text.lower().find(hint)
        if idx != -1:
            return text[max(0, idx - 60):idx + 20].strip()
    return ""


def score_site(soup, r):
    score = 0
    if r.url.startswith("https"):
        score += 15
    if soup.find("meta", attrs={"name": "viewport"}):
        score += 15
    title = soup.find("title")
    if title and len(title.get_text(strip=True)) > 10:
        score += 15
    if soup.find(string=re.compile(r"contact", re.I)):
        score += 15
    if PHONE_RE.search(soup.get_text(" ")):
        score += 15
    meta_desc = soup.find("meta", attrs={"name": "description"})
    if meta_desc and meta_desc.get("content"):
        score += 15
    if soup.find("form"):
        score += 10
    return min(score, 100)


def scrape_site(domain, url, city_query, engine_source):
    full_url = url if url.startswith("http") else f"https://{domain}"
    try:
        r = requests.get(full_url, headers=rand_headers(), timeout=REQUEST_TIMEOUT)
        soup = BeautifulSoup(r.text, "html.parser")

        emails = extract_emails(r.text)
        contact_link = soup.find("a", href=re.compile(r"contact", re.I))
        if contact_link and contact_link.get("href"):
            contact_url = urllib.parse.urljoin(full_url, contact_link["href"])
            try:
                r2 = requests.get(contact_url, headers=rand_headers(), timeout=REQUEST_TIMEOUT)
                emails.extend(extract_emails(r2.text))
                soup2 = BeautifulSoup(r2.text, "html.parser")
                owner = guess_owner_name(soup2) or guess_owner_name(soup)
            except Exception:
                owner = guess_owner_name(soup)
        else:
            owner = guess_owner_name(soup)

        emails = list(dict.fromkeys(emails))
        phone_match = PHONE_RE.search(soup.get_text(" "))
        title_tag = soup.find("title")
        business_name = title_tag.get_text(strip=True)[:80] if title_tag else domain

        if not emails:
            return None

        return Lead(
            business=business_name, domain=domain, url=full_url,
            owner_name=owner[:60], email=emails[0],
            phone=phone_match.group(0) if phone_match else "",
            score=score_site(soup, r), city_query=city_query, engine_source=engine_source,
        )
    except Exception:
        return None


def write_excel(leads, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    headers = ["Business", "Domain", "URL", "Owner/Contact", "Email", "Phone",
               "Score", "Search Query", "Source Engine(s)", "Date Found"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")

    for lead in sorted(leads, key=lambda l: l.score, reverse=True):
        ws.append([lead.business, lead.domain, lead.url, lead.owner_name, lead.email,
                   lead.phone, lead.score, lead.city_query, lead.engine_source, lead.date_found])

    widths = [30, 22, 35, 30, 28, 16, 8, 22, 16, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    wb.save(path)


def main(target_leads):
    conn = init_db()
    queries = [f"{niche} {city}" for niche in NICHES for city in CITIES]
    random.shuffle(queries)

    print(f"[+] {len(queries)} search queries queued across {len(SEARCH_ENGINES)} engines")
    print(f"[+] Target: {target_leads} qualified leads")

    all_candidates = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS_SEARCH) as pool:
        futures = {pool.submit(run_search_task, q): q for q in queries}
        for future in concurrent.futures.as_completed(futures):
            query = futures[future]
            try:
                results = future.result()
            except Exception as e:
                print(f"  [search] failed for '{query}': {e}")
                continue
            for domain, url in results:
                if not already_seen(conn, domain):
                    all_candidates.append((domain, url, query))
            print(f"  [search] '{query}' -> {len(results)} raw candidates (total pool: {len(all_candidates)})")
            if len(all_candidates) >= target_leads * 4:
                break

    print(f"[+] {len(all_candidates)} unique candidate domains to scrape")

    leads = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS_SCRAPE) as pool:
        futures = {
            pool.submit(scrape_site, domain, url, query, "multi"): domain
            for domain, url, query in all_candidates
        }
        for future in concurrent.futures.as_completed(futures):
            domain = futures[future]
            lead = future.result()
            mark_seen(conn, domain)
            if lead:
                leads.append(lead)
                print(f"  [scrape] QUALIFIED: {domain} (score {lead.score}, email {lead.email})")
            if len(leads) >= target_leads:
                break

    conn.close()

    out_path = OUTPUT_DIR / f"leads_{datetime.utcnow().strftime('%Y-%m-%d')}.xlsx"
    write_excel(leads, out_path)
    print(f"\n[+] Done. {len(leads)} qualified leads written to {out_path}")
    if len(leads) < target_leads:
        print(f"[!] Fell short of target by {target_leads - len(leads)}. Add more cities/niches to CONFIG.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--target", type=int, default=TARGET_LEADS)
    args = parser.parse_args()
    main(args.target)
