#!/usr/bin/env python3
"""
Growth Supply House - Automated Lead Generator

Continuous multi-engine search -> site scrape -> niche verification -> Excel.

Designed to run every few minutes on GitHub Actions. Each run takes a rotating
slice of the query space (least-recently-searched first), verifies that every
candidate site is genuinely a business in the target niche, records every domain
it has visited so it never revisits one, and appends new qualified leads to a
cumulative workbook.

Free-stack, no paid APIs.
"""

import argparse
import concurrent.futures
import random
import re
import sqlite3
import threading
import time
import urllib.parse
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

# ---------------------------------------------------------------------------
# NICHE DEFINITIONS
#
# A candidate site is only kept if it reads like an actual operator in one of
# these trades. "strong" terms identify the trade; "supporting" terms are the
# vocabulary a real service business uses. Both thresholds must be met.
# ---------------------------------------------------------------------------

NICHE_PROFILES = {
    "Plumbing": {
        "queries": [
            "plumber", "plumbing company", "emergency plumber",
            "drain cleaning service", "water heater repair", "sewer repair",
        ],
        "strong": ("plumber", "plumbing", "rooter"),
        "supporting": (
            "drain", "sewer", "water heater", "repipe", "leak", "faucet",
            "toilet", "sump pump", "backflow", "tankless", "clog", "pipe",
        ),
    },
    "HVAC": {
        "queries": [
            "HVAC contractor", "air conditioning repair", "heating and cooling company",
            "furnace repair", "AC installation", "heat pump installation",
        ],
        "strong": ("hvac", "air conditioning", "heating and cooling", "heating & cooling"),
        "supporting": (
            "furnace", "heat pump", "air handler", "ductwork", "duct", "ac repair",
            "air conditioner", "thermostat", "refrigerant", "condenser", "seer",
            "boiler", "mini split", "air quality",
        ),
    },
    "Roofing": {
        "queries": [
            "roofing contractor", "roofer", "roof repair company",
            "roof replacement", "residential roofing company", "metal roofing contractor",
        ],
        "strong": ("roofing", "roofer", "roof repair", "roof replacement"),
        "supporting": (
            "shingle", "asphalt shingle", "metal roof", "flat roof", "gutter",
            "flashing", "underlayment", "soffit", "fascia", "re-roof", "reroof",
            "storm damage", "hail damage", "tpo", "torch down",
        ),
    },
}

MIN_STRONG_HITS = 1      # must clearly be the trade
MIN_SUPPORT_HITS = 2     # must use the trade's working vocabulary

CITIES = [
    "Birmingham AL", "Montgomery AL", "Huntsville AL", "Mobile AL", "Tuscaloosa AL",
    "Anchorage AK", "Fairbanks AK", "Juneau AK", "Phoenix AZ", "Tucson AZ", "Mesa AZ",
    "Chandler AZ", "Scottsdale AZ", "Glendale AZ", "Gilbert AZ", "Tempe AZ", "Peoria AZ",
    "Surprise AZ", "Little Rock AR", "Fort Smith AR", "Fayetteville AR", "Springdale AR",
    "Jonesboro AR", "Los Angeles CA", "San Diego CA", "San Jose CA", "San Francisco CA",
    "Fresno CA", "Sacramento CA", "Long Beach CA", "Oakland CA", "Bakersfield CA",
    "Anaheim CA", "Santa Ana CA", "Riverside CA", "Stockton CA", "Irvine CA", "Chula Vista CA",
    "Fremont CA", "Modesto CA", "Fontana CA", "Moreno Valley CA", "Santa Clarita CA",
    "Oxnard CA", "Rancho Cucamonga CA", "Ontario CA", "Elk Grove CA", "Corona CA",
    "Salinas CA", "Palmdale CA", "Escondido CA", "Sunnyvale CA", "Torrance CA", "Denver CO",
    "Colorado Springs CO", "Aurora CO", "Fort Collins CO", "Lakewood CO", "Thornton CO",
    "Arvada CO", "Westminster CO", "Pueblo CO", "Boulder CO", "Bridgeport CT", "New Haven CT",
    "Stamford CT", "Hartford CT", "Waterbury CT", "Norwalk CT", "Danbury CT", "Wilmington DE",
    "Dover DE", "Newark DE", "Jacksonville FL", "Miami FL", "Tampa FL", "Orlando FL",
    "St Petersburg FL", "Hialeah FL", "Port St Lucie FL", "Cape Coral FL", "Tallahassee FL",
    "Fort Lauderdale FL", "Pembroke Pines FL", "Hollywood FL", "Gainesville FL", "Miramar FL",
    "Coral Springs FL", "Clearwater FL", "Palm Bay FL", "West Palm Beach FL", "Lakeland FL",
    "Pompano Beach FL", "Sarasota FL", "Naples FL", "Ocala FL", "Kissimmee FL", "Atlanta GA",
    "Augusta GA", "Columbus GA", "Macon GA", "Savannah GA", "Athens GA", "Sandy Springs GA",
    "Roswell GA", "Marietta GA", "Alpharetta GA", "Honolulu HI", "Hilo HI", "Kailua HI",
    "Boise ID", "Meridian ID", "Nampa ID", "Idaho Falls ID", "Pocatello ID", "Chicago IL",
    "Aurora IL", "Joliet IL", "Naperville IL", "Rockford IL", "Springfield IL", "Elgin IL",
    "Peoria IL", "Champaign IL", "Waukegan IL", "Indianapolis IN", "Fort Wayne IN",
    "Evansville IN", "South Bend IN", "Carmel IN", "Fishers IN", "Bloomington IN",
    "Hammond IN", "Gary IN", "Lafayette IN", "Des Moines IA", "Cedar Rapids IA",
    "Davenport IA", "Sioux City IA", "Iowa City IA", "Ankeny IA", "Wichita KS",
    "Overland Park KS", "Kansas City KS", "Olathe KS", "Topeka KS", "Lawrence KS",
    "Louisville KY", "Lexington KY", "Bowling Green KY", "Owensboro KY", "Covington KY",
    "New Orleans LA", "Baton Rouge LA", "Shreveport LA", "Lafayette LA", "Lake Charles LA",
    "Metairie LA", "Portland ME", "Lewiston ME", "Bangor ME", "Baltimore MD", "Columbia MD",
    "Germantown MD", "Silver Spring MD", "Waldorf MD", "Frederick MD", "Rockville MD",
    "Annapolis MD", "Boston MA", "Worcester MA", "Springfield MA", "Cambridge MA", "Lowell MA",
    "Brockton MA", "Quincy MA", "New Bedford MA", "Detroit MI", "Grand Rapids MI", "Warren MI",
    "Sterling Heights MI", "Ann Arbor MI", "Lansing MI", "Flint MI", "Dearborn MI",
    "Livonia MI", "Troy MI", "Minneapolis MN", "St Paul MN", "Rochester MN", "Duluth MN",
    "Bloomington MN", "Brooklyn Park MN", "Plymouth MN", "Woodbury MN", "Jackson MS",
    "Gulfport MS", "Southaven MS", "Biloxi MS", "Hattiesburg MS", "Kansas City MO",
    "St Louis MO", "Springfield MO", "Columbia MO", "Independence MO", "Lees Summit MO",
    "OFallon MO", "St Charles MO", "Billings MT", "Missoula MT", "Great Falls MT",
    "Bozeman MT", "Omaha NE", "Lincoln NE", "Bellevue NE", "Grand Island NE", "Las Vegas NV",
    "Henderson NV", "Reno NV", "North Las Vegas NV", "Sparks NV", "Carson City NV",
    "Manchester NH", "Nashua NH", "Concord NH", "Newark NJ", "Jersey City NJ", "Paterson NJ",
    "Elizabeth NJ", "Edison NJ", "Woodbridge NJ", "Trenton NJ", "Camden NJ", "Toms River NJ",
    "Albuquerque NM", "Las Cruces NM", "Rio Rancho NM", "Santa Fe NM", "New York NY",
    "Buffalo NY", "Rochester NY", "Yonkers NY", "Syracuse NY", "Albany NY", "New Rochelle NY",
    "Mount Vernon NY", "Schenectady NY", "Utica NY", "Brooklyn NY", "Queens NY", "Bronx NY",
    "Staten Island NY", "Charlotte NC", "Raleigh NC", "Greensboro NC", "Durham NC",
    "Winston Salem NC", "Fayetteville NC", "Cary NC", "Wilmington NC", "High Point NC",
    "Concord NC", "Asheville NC", "Fargo ND", "Bismarck ND", "Grand Forks ND", "Columbus OH",
    "Cleveland OH", "Cincinnati OH", "Toledo OH", "Akron OH", "Dayton OH", "Parma OH",
    "Canton OH", "Youngstown OH", "Lorain OH", "Oklahoma City OK", "Tulsa OK", "Norman OK",
    "Broken Arrow OK", "Edmond OK", "Lawton OK", "Portland OR", "Salem OR", "Eugene OR",
    "Gresham OR", "Hillsboro OR", "Beaverton OR", "Bend OR", "Medford OR", "Philadelphia PA",
    "Pittsburgh PA", "Allentown PA", "Erie PA", "Reading PA", "Scranton PA", "Bethlehem PA",
    "Lancaster PA", "Harrisburg PA", "York PA", "Providence RI", "Warwick RI", "Cranston RI",
    "Pawtucket RI", "Charleston SC", "Columbia SC", "North Charleston SC", "Mount Pleasant SC",
    "Rock Hill SC", "Greenville SC", "Summerville SC", "Myrtle Beach SC", "Sioux Falls SD",
    "Rapid City SD", "Aberdeen SD", "Nashville TN", "Memphis TN", "Knoxville TN",
    "Chattanooga TN", "Clarksville TN", "Murfreesboro TN", "Franklin TN", "Johnson City TN",
    "Houston TX", "San Antonio TX", "Dallas TX", "Austin TX", "Fort Worth TX", "El Paso TX",
    "Arlington TX", "Corpus Christi TX", "Plano TX", "Laredo TX", "Lubbock TX", "Garland TX",
    "Irving TX", "Amarillo TX", "Grand Prairie TX", "Brownsville TX", "McKinney TX",
    "Frisco TX", "Pasadena TX", "Killeen TX", "McAllen TX", "Mesquite TX", "Midland TX",
    "Denton TX", "Waco TX", "Carrollton TX", "Round Rock TX", "Abilene TX", "Pearland TX",
    "Richardson TX", "College Station TX", "Sugar Land TX", "Odessa TX", "Tyler TX",
    "Beaumont TX", "Katy TX", "The Woodlands TX", "San Marcos TX", "Conroe TX",
    "New Braunfels TX", "Salt Lake City UT", "West Valley City UT", "Provo UT",
    "West Jordan UT", "Orem UT", "Sandy UT", "Ogden UT", "St George UT", "Lehi UT",
    "Burlington VT", "Rutland VT", "Virginia Beach VA", "Norfolk VA", "Chesapeake VA",
    "Richmond VA", "Newport News VA", "Alexandria VA", "Hampton VA", "Roanoke VA",
    "Arlington VA", "Fredericksburg VA", "Seattle WA", "Spokane WA", "Tacoma WA",
    "Vancouver WA", "Bellevue WA", "Kent WA", "Everett WA", "Renton WA", "Federal Way WA",
    "Yakima WA", "Bellingham WA", "Olympia WA", "Charleston WV", "Huntington WV",
    "Morgantown WV", "Milwaukee WI", "Madison WI", "Green Bay WI", "Kenosha WI", "Racine WI",
    "Appleton WI", "Waukesha WI", "Oshkosh WI", "Cheyenne WY", "Casper WY", "Laramie WY",
    "Washington DC"]

# ---------------------------------------------------------------------------
# EXCLUSIONS
# ---------------------------------------------------------------------------

# Adult / explicit content. Any hit on domain or page text discards the site
# outright before it is ever considered as a lead.
#
# Split by specificity on purpose: several obvious-looking terms are substrings
# of ordinary words a trades business legitimately uses -- "sex" in Essex and
# Sussex, "strip" in weatherstripping -- so those are matched with care rather
# than as bare substrings.
ADULT_DOMAIN_STRONG = (
    "porn", "xxx", "escort", "cam4", "chaturbate", "onlyfans", "hentai",
    "nsfw", "fetish", "erotic", "milf", "bdsm", "hookup", "swinger", "brothel",
    "playboy", "redtube", "xvideos", "xhamster", "pornhub", "brazzers",
    "camsoda", "stripchat", "bongacams", "livejasmin", "camgirl", "sexcam",
)
ADULT_DOMAIN_AMBIGUOUS = ("sex", "adult", "nude", "strip")

# Ordinary words that legitimately contain an ambiguous term above. A match that
# sits entirely inside one of these is not treated as a hit.
AMBIGUOUS_SAFE_WORDS = (
    "essex", "sussex", "middlesex", "wessex", "unisex", "sexton",
    "weatherstrip", "weatherstripping", "stripe", "stripes", "striping",
    "stripping", "stripped", "airstrip", "stripmall", "adulting",
    "adultday", "adultcare", "nudge",
)

ADULT_TEXT_TERMS = (
    "porn", "xxx video", "adult video", "escort service", "live sex", "sex cam",
    "webcam girls", "nude photos", "explicit content", "adults only", "18+ only",
    "hookup site", "onlyfans", "fetish", "hardcore video", "cam girls",
)

# Lead directories, marketplaces, franchises-of-listings, manufacturers, news.
# These pass a naive niche keyword test but are not businesses you can sell to.
SKIP_DOMAINS = (
    "facebook.com", "yelp.com", "angi.com", "angieslist.com", "homeadvisor.com",
    "bbb.org", "yellowpages.com", "thumbtack.com", "instagram.com", "linkedin.com",
    "youtube.com", "pinterest.com", "mapquest.com", "google.com", "bing.com",
    "duckduckgo.com", "twitter.com", "x.com", "tiktok.com", "reddit.com",
    "expertise.com", "buildzoom.com", "todayshomeowner.com", "porch.com",
    "houzz.com", "nextdoor.com", "manta.com", "chamberofcommerce.com",
    "networx.com", "modernize.com", "homeguide.com", "thumbtack.com",
    "bark.com", "craigslist.org", "indeed.com", "glassdoor.com", "ziprecruiter.com",
    "wikipedia.org", "amazon.com", "ebay.com", "homedepot.com", "lowes.com",
    "menards.com", "walmart.com", "tripadvisor.com", "apartments.com",
    "zillow.com", "realtor.com", "redfin.com", "trulia.com", "citysearch.com",
    "superpages.com", "local.com", "merchantcircle.com", "hotfrog.com",
    "dexknows.com", "whitepages.com", "spokeo.com", "trustpilot.com",
    "consumeraffairs.com", "sitejabber.com", "birdeye.com", "opendoor.com",
    "carrier.com", "trane.com", "lennox.com", "rheem.com", "goodmanmfg.com",
    "bradfordwhite.com", "gaf.com", "owenscorning.com", "certainteed.com",
    "iko.com", "malarkeyroofing.com", "tamko.com", "atlasroofing.com",
    "cedur.com", "davinciroofscapes.com", "moen.com", "kohler.com", "delta.com",
    "energy.gov", "epa.gov", "osha.gov", "irs.gov", "usa.gov",
    "forbes.com", "nytimes.com", "cnn.com", "bloomberg.com", "businessinsider.com",
    "quora.com", "medium.com", "wordpress.com", "blogspot.com", "wixsite.com",
    "godaddysites.com", "squarespace.com", "shopify.com", "etsy.com",
)

# Language typical of a directory/marketplace page rather than an operator.
DIRECTORY_TEXT_MARKERS = (
    "compare quotes", "get matched", "find a pro", "find pros", "top 10 best",
    "best of", "browse listings", "claim your listing", "claim this business",
    "add your business", "list your business", "advertise with us",
    "free estimates from", "matched with", "get up to 4 quotes",
    "screened and approved", "verified pros", "read reviews of",
    "directory of", "business listings", "leads for contractors",
    "join our network", "become a pro", "pro network", "franchise opportunit",
)
DIRECTORY_MARKER_LIMIT = 2   # this many markers = treat as a directory

# ---------------------------------------------------------------------------
# RUNTIME
# ---------------------------------------------------------------------------

QUERIES_PER_RUN = 12         # rotating slice worked by each run
MAX_WORKERS_SEARCH = 6
MAX_WORKERS_SCRAPE = 10
REQUEST_TIMEOUT = 12
SEARCH_RETRIES = 2
MIN_DELAY = 0.8
MAX_DELAY = 2.0
DEFAULT_MAX_SECONDS = 420    # stay well inside the workflow's schedule interval

DB_PATH = Path("seen_domains.sqlite3")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)
MASTER_XLSX = OUTPUT_DIR / "leads_master.xlsx"

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0 Safari/537.36 Edg/123.0",
]

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
PHONE_RE = re.compile(r"(\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4})")

# Asset filenames like "flags-sprite@2x.webp" match the email shape. Reject any
# address whose domain half ends in a file extension.
ASSET_SUFFIXES = (
    ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".ico", ".bmp", ".avif",
    ".css", ".js", ".json", ".woff", ".woff2", ".ttf", ".eot", ".mp4", ".webm",
    ".pdf", ".zip", ".xml",
)
BAD_EMAIL_HINTS = (
    "example.com", "example.org", "yourdomain", "domain.com", "email.com",
    "sentry.io", "wixpress.com", "wix.com", "squarespace.com", "godaddy.com",
    "sentry-next", "test.com", "yoursite", "company.com", "acme.com",
)
LOW_VALUE_LOCALPARTS = ("noreply", "no-reply", "donotreply", "postmaster", "abuse", "webmaster")
PREFERRED_LOCALPARTS = ("info", "contact", "office", "sales", "service", "admin", "hello", "estimates")

OWNER_HINTS = (
    "owner", "founder", "president", "ceo", "co-owner", "proprietor",
    "general manager", "operations manager", "managing partner",
)
NAME_RE = re.compile(r"\b([A-Z][a-z]{1,15})\s+([A-Z][a-z]{1,20})\b")

CONTACT_PATH_RE = re.compile(r"(contact|about|our-team|meet-the-team|staff|who-we-are)", re.I)


def rand_headers():
    return {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept-Language": "en-US,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }


def polite_sleep():
    time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))


def utcnow():
    return datetime.now(timezone.utc)


# ---------------------------------------------------------------------------
# STATE  (all access from the main thread only)
# ---------------------------------------------------------------------------

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS seen (
        domain TEXT PRIMARY KEY,
        first_seen TEXT,
        status TEXT,
        niche TEXT,
        query TEXT
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS queries (
        q TEXT PRIMARY KEY,
        niche TEXT,
        last_run TEXT,
        runs INTEGER DEFAULT 0,
        found INTEGER DEFAULT 0
    )""")
    # Older installs used a two-column seen table; bring it forward.
    cols = {r[1] for r in conn.execute("PRAGMA table_info(seen)")}
    for col in ("status", "niche", "query"):
        if col not in cols:
            conn.execute(f"ALTER TABLE seen ADD COLUMN {col} TEXT")
    # Rows written before status tracking carry no verdict. Clear them so the
    # new niche/adult/directory filters get to judge those domains properly
    # instead of them being skipped forever as "already seen".
    conn.execute("DELETE FROM seen WHERE status IS NULL OR status = ''")
    conn.commit()
    return conn


def sync_query_catalog(conn):
    """Make sure every niche x city pair exists in the rotation table."""
    rows = []
    for niche, profile in NICHE_PROFILES.items():
        for phrase in profile["queries"]:
            for city in CITIES:
                rows.append((f"{phrase} {city}", niche))
    conn.executemany("INSERT OR IGNORE INTO queries (q, niche) VALUES (?, ?)", rows)
    conn.commit()
    return len(rows)


def pick_queries(conn, limit):
    """Least-recently-searched first, so coverage rotates across the country."""
    cur = conn.execute(
        "SELECT q, niche FROM queries ORDER BY COALESCE(last_run, '') ASC, RANDOM() LIMIT ?",
        (limit,),
    )
    return cur.fetchall()


def mark_query_run(conn, q, found):
    conn.execute(
        "UPDATE queries SET last_run = ?, runs = runs + 1, found = found + ? WHERE q = ?",
        (utcnow().isoformat(), found, q),
    )


def already_seen(conn, domain):
    return conn.execute("SELECT 1 FROM seen WHERE domain = ?", (domain,)).fetchone() is not None


def mark_seen(conn, domain, status, niche="", query=""):
    conn.execute(
        """INSERT INTO seen (domain, first_seen, status, niche, query) VALUES (?, ?, ?, ?, ?)
           ON CONFLICT(domain) DO UPDATE SET status = excluded.status, niche = excluded.niche""",
        (domain, utcnow().isoformat(), status, niche, query),
    )


# ---------------------------------------------------------------------------
# VERIFICATION
# ---------------------------------------------------------------------------

def _ambiguous_hit(haystack, term):
    """True if `term` appears somewhere that is not inside an ordinary word."""
    start = haystack.find(term)
    while start != -1:
        end = start + len(term)
        benign = False
        for safe in AMBIGUOUS_SAFE_WORDS:
            if term not in safe:
                continue
            probe = haystack.rfind(safe, max(0, start - len(safe)), end + len(safe))
            if probe != -1 and probe <= start and probe + len(safe) >= end:
                benign = True
                break
        if not benign:
            return True
        start = haystack.find(term, start + 1)
    return False


def is_adult_domain(domain):
    d = domain.lower()
    if any(term in d for term in ADULT_DOMAIN_STRONG):
        return True
    return any(_ambiguous_hit(d, term) for term in ADULT_DOMAIN_AMBIGUOUS)


def is_adult_content(domain, text):
    if is_adult_domain(domain):
        return True
    low = text[:200_000].lower()
    return any(term in low for term in ADULT_TEXT_TERMS)


def looks_like_directory(text):
    low = text[:200_000].lower()
    hits = sum(1 for m in DIRECTORY_TEXT_MARKERS if m in low)
    return hits >= DIRECTORY_MARKER_LIMIT, hits


def classify_niche(text, expected=None):
    """Return (niche, strong_hits, support_hits) or (None, 0, 0) if it is not
    convincingly a business in one of our trades."""
    low = text[:200_000].lower()
    best = (None, 0, 0)
    order = [expected] + [n for n in NICHE_PROFILES if n != expected] if expected else list(NICHE_PROFILES)
    for niche in order:
        if niche not in NICHE_PROFILES:
            continue
        profile = NICHE_PROFILES[niche]
        strong = sum(1 for t in profile["strong"] if t in low)
        support = sum(1 for t in profile["supporting"] if t in low)
        if strong >= MIN_STRONG_HITS and support >= MIN_SUPPORT_HITS:
            if (strong, support) > (best[1], best[2]):
                best = (niche, strong, support)
    return best


# ---------------------------------------------------------------------------
# EXTRACTION
# ---------------------------------------------------------------------------

def extract_emails(text, site_domain=""):
    """Return plausible addresses, best first."""
    out = []
    for raw in EMAIL_RE.findall(text):
        e = raw.strip().strip(".").lower()
        local, _, host = e.partition("@")
        if not host or host.endswith(ASSET_SUFFIXES):
            continue
        if any(bad in e for bad in BAD_EMAIL_HINTS):
            continue
        if len(local) > 40 or len(host) > 60:
            continue
        if any(ch.isdigit() for ch in host.split(".")[0]) and host.split(".")[0].endswith("x"):
            continue  # "@2x.webp"-style asset leftovers
        if e not in out:
            out.append(e)

    root = site_domain.lower().replace("www.", "")

    def rank(e):
        local, _, host = e.partition("@")
        same_domain = root and (host == root or host.endswith("." + root) or root.endswith("." + host))
        preferred = any(local.startswith(p) for p in PREFERRED_LOCALPARTS)
        low_value = any(p in local for p in LOW_VALUE_LOCALPARTS)
        return (not same_domain, low_value, not preferred, len(e))

    return sorted(out, key=rank)


def guess_owner_name(soup):
    """Look for a proper name sitting next to an ownership title."""
    text = soup.get_text(" ", strip=True)
    low = text.lower()
    for hint in OWNER_HINTS:
        idx = low.find(hint)
        while idx != -1:
            window = text[max(0, idx - 70): idx + len(hint) + 40]
            names = NAME_RE.findall(window)
            for first, last in names:
                if first.lower() in ("the", "our", "meet", "contact", "about", "we", "us"):
                    continue
                return f"{first} {last}"
            idx = low.find(hint, idx + 1)
    return ""


def score_site(soup, resp, strong, support):
    score = 0
    if resp.url.startswith("https"):
        score += 12
    if soup.find("meta", attrs={"name": "viewport"}):
        score += 12
    title = soup.find("title")
    if title and len(title.get_text(strip=True)) > 10:
        score += 12
    if soup.find(string=re.compile(r"contact", re.I)):
        score += 10
    if PHONE_RE.search(soup.get_text(" ")):
        score += 14
    meta_desc = soup.find("meta", attrs={"name": "description"})
    if meta_desc and meta_desc.get("content"):
        score += 10
    if soup.find("form"):
        score += 10
    score += min(strong * 6, 12)      # depth of trade match
    score += min(support * 2, 8)
    return min(score, 100)


def business_name(soup, domain):
    for tag in (soup.find("meta", attrs={"property": "og:site_name"}),
                soup.find("meta", attrs={"property": "og:title"})):
        if tag and tag.get("content", "").strip():
            return tag["content"].strip()[:80]
    title = soup.find("title")
    if title:
        name = re.split(r"\s+[|\-–—]\s+", title.get_text(strip=True))[0]
        if len(name) > 3:
            return name[:80]
        return title.get_text(strip=True)[:80]
    return domain


# ---------------------------------------------------------------------------
# SEARCH
# ---------------------------------------------------------------------------

def _get(url, **kw):
    kw.setdefault("headers", rand_headers())
    kw.setdefault("timeout", REQUEST_TIMEOUT)
    return requests.get(url, **kw)


def search_duckduckgo(query, max_results=12):
    last = None
    for attempt in range(SEARCH_RETRIES + 1):
        try:
            r = requests.post(
                "https://html.duckduckgo.com/html/",
                data={"q": query}, headers=rand_headers(), timeout=REQUEST_TIMEOUT,
            )
            soup = BeautifulSoup(r.text, "html.parser")
            hits = [a.get("href") for a in soup.select("a.result__a")[:max_results] if a.get("href")]
            if hits:
                return hits
            last = "no results"
        except Exception as e:
            last = type(e).__name__
        if attempt < SEARCH_RETRIES:
            time.sleep(1.5 * (attempt + 1))
    print(f"  [ddg] {query!r}: {last}")
    return []


def search_bing(query, max_results=12):
    last = None
    for attempt in range(SEARCH_RETRIES + 1):
        try:
            r = _get("https://www.bing.com/search", params={"q": query, "count": max_results})
            soup = BeautifulSoup(r.text, "html.parser")
            hits = [a.get("href") for a in soup.select("li.b_algo h2 a")[:max_results] if a.get("href")]
            if hits:
                return hits
            last = "no results"
        except Exception as e:
            last = type(e).__name__
        if attempt < SEARCH_RETRIES:
            time.sleep(1.5 * (attempt + 1))
    print(f"  [bing] {query!r}: {last}")
    return []


def search_mojeek(query, max_results=12):
    """Independent index, tolerant of automated clients - a useful third leg."""
    try:
        r = _get("https://www.mojeek.com/search", params={"q": query})
        soup = BeautifulSoup(r.text, "html.parser")
        return [a.get("href") for a in soup.select("ul.results-standard li h2 a")[:max_results] if a.get("href")]
    except Exception as e:
        print(f"  [mojeek] {query!r}: {type(e).__name__}")
        return []


def search_google(query, max_results=12):
    try:
        r = _get("https://www.google.com/search", params={"q": query, "num": max_results})
        if r.status_code != 200 or "captcha" in r.text.lower():
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
        print(f"  [google] {query!r}: {type(e).__name__}")
        return []


SEARCH_ENGINES = [
    ("ddg", search_duckduckgo),
    ("bing", search_bing),
    ("mojeek", search_mojeek),
    ("google", search_google),
]


def clean_domain(url):
    try:
        netloc = urllib.parse.urlparse(url).netloc.lower()
        if netloc.startswith("www."):
            netloc = netloc[4:]
        return netloc or None
    except Exception:
        return None


def is_blocked_domain(domain):
    """Exact host or true subdomain only. A loose substring test is unsafe here:
    "bing.com" is a substring of "plumbing.com"."""
    d = domain.lower()
    if is_adult_domain(d):
        return "adult"
    for skip in SKIP_DOMAINS:
        if d == skip or d.endswith("." + skip):
            return "directory"
    return None


def run_search_task(query):
    """Returns (query, [(domain, url, engine), ...]) - no DB access, thread safe."""
    found, order = {}, []
    for name, engine in SEARCH_ENGINES:
        for url in engine(query):
            domain = clean_domain(url)
            if not domain or domain in found:
                continue
            found[domain] = (domain, url, name)
            order.append(domain)
        polite_sleep()
    return query, [found[d] for d in order]


@dataclass
class Lead:
    business: str = ""
    niche: str = ""
    domain: str = ""
    url: str = ""
    owner_name: str = ""
    email: str = ""
    phone: str = ""
    city: str = ""
    score: int = 0
    query: str = ""
    engine: str = ""
    date_found: str = field(default_factory=lambda: utcnow().strftime("%Y-%m-%d"))


@dataclass
class ScrapeResult:
    domain: str
    status: str            # qualified | adult | directory | off_niche | no_email | unreachable
    lead: "Lead | None" = None
    niche: str = ""


def scrape_site(domain, url, query, engine, expected_niche, city):
    """Fetch the site, prove it is a real operator in the niche, pull contacts."""
    full_url = url if url.startswith("http") else f"https://{domain}"
    try:
        r = _get(full_url, allow_redirects=True)
        if r.status_code >= 400 or not r.text:
            return ScrapeResult(domain, "unreachable")
        home_html = r.text
        soup = BeautifulSoup(home_html, "html.parser")
    except Exception:
        return ScrapeResult(domain, "unreachable")

    page_text = soup.get_text(" ", strip=True)

    if is_adult_content(domain, page_text) or is_adult_content(domain, home_html):
        return ScrapeResult(domain, "adult")

    is_dir, _ = looks_like_directory(page_text)
    if is_dir:
        return ScrapeResult(domain, "directory")

    # Pull the contact/about page early: it carries both the contact details and
    # more of the trade vocabulary we verify against.
    combined_text = page_text
    combined_html = home_html
    contact_soup = None
    link = soup.find("a", href=CONTACT_PATH_RE)
    if link and link.get("href"):
        try:
            contact_url = urllib.parse.urljoin(r.url, link["href"])
            if clean_domain(contact_url) == domain:
                r2 = _get(contact_url)
                if r2.status_code < 400 and r2.text:
                    contact_soup = BeautifulSoup(r2.text, "html.parser")
                    combined_html += "\n" + r2.text
                    combined_text += " " + contact_soup.get_text(" ", strip=True)
        except Exception:
            pass

    niche, strong, support = classify_niche(combined_text, expected_niche)
    if not niche:
        return ScrapeResult(domain, "off_niche")

    emails = extract_emails(combined_html, domain)
    if not emails:
        return ScrapeResult(domain, "no_email", niche=niche)

    phone = PHONE_RE.search(combined_text)
    owner = (guess_owner_name(contact_soup) if contact_soup else "") or guess_owner_name(soup)

    lead = Lead(
        business=business_name(soup, domain),
        niche=niche,
        domain=domain,
        url=r.url,
        owner_name=owner[:60],
        email=emails[0],
        phone=phone.group(0) if phone else "",
        city=city,
        score=score_site(soup, r, strong, support),
        query=query,
        engine=engine,
    )
    return ScrapeResult(domain, "qualified", lead=lead, niche=niche)


# ---------------------------------------------------------------------------
# OUTPUT
# ---------------------------------------------------------------------------

HEADERS = ["Business", "Niche", "Domain", "URL", "Owner/Contact", "Email", "Phone",
           "City", "Score", "Search Query", "Engine", "Date Found"]
COL_WIDTHS = [32, 10, 24, 38, 24, 30, 16, 20, 7, 30, 9, 12]


def _style_header(ws):
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    for i, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"


def lead_row(lead):
    return [lead.business, lead.niche, lead.domain, lead.url, lead.owner_name,
            lead.email, lead.phone, lead.city, lead.score, lead.query,
            lead.engine, lead.date_found]


def _open_book(path):
    """Open an existing workbook only if its columns still match. A workbook
    written by an older column layout is replaced rather than appended to,
    which would silently misalign every field."""
    if path.exists():
        try:
            wb = load_workbook(path)
            ws = wb.active
            header = [c.value for c in ws[1]] if ws.max_row >= 1 else []
            # Trailing columns added by the outreach stage (Audit PDF, Audit
            # Status, Emailed At, Email Status) are expected. Requiring an
            # exact match here wiped every existing row each time outreach.py
            # had extended the sheet.
            if header[:len(HEADERS)] == HEADERS:
                existing = {str(r[2]).lower() for r in ws.iter_rows(min_row=2, values_only=True) if r[2]}
                return wb, ws, existing
            print(f"  [xlsx] {path.name} uses an old column layout - rewriting")
        except Exception as e:
            print(f"  [xlsx] {path.name} unreadable ({type(e).__name__}) - rewriting")
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    ws.append(HEADERS)
    _style_header(ws)
    return wb, ws, set()


def _append(path, leads):
    wb, ws, existing = _open_book(path)
    added = 0
    for lead in sorted(leads, key=lambda l: l.score, reverse=True):
        key = lead.domain.lower()
        if key in existing:
            continue
        ws.append(lead_row(lead))
        existing.add(key)
        added += 1
    if added or not path.exists():
        wb.save(path)
    return added, ws.max_row - 1


def append_to_master(leads):
    """Append new leads to the cumulative workbook, skipping domains already in it."""
    return _append(MASTER_XLSX, leads)


def write_daily(leads, path):
    """One workbook per day holding that day's additions."""
    _append(path, leads)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

_CITY_SUFFIXES = sorted(CITIES, key=len, reverse=True)


def city_from_query(q):
    for city in _CITY_SUFFIXES:
        if q.endswith(" " + city):
            return city
    return ""


def main(n_queries, max_seconds, fail_on_empty):
    started = time.monotonic()
    deadline = started + max_seconds
    search_deadline = started + max_seconds * 0.55

    conn = init_db()
    total_queries = sync_query_catalog(conn)
    batch = pick_queries(conn, n_queries)

    already = conn.execute("SELECT COUNT(*) FROM seen").fetchone()[0]
    print(f"[+] catalog: {total_queries} queries across {len(CITIES)} US cities, "
          f"{len(NICHE_PROFILES)} niches")
    print(f"[+] this run: {len(batch)} queries | {already} domains already visited "
          f"| budget {max_seconds}s")
    for q, niche in batch:
        print(f"      - [{niche}] {q}")

    # ---- search -----------------------------------------------------------
    candidates, engines_ok = [], 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS_SEARCH) as pool:
        futures = {pool.submit(run_search_task, q): (q, niche) for q, niche in batch}
        for future in concurrent.futures.as_completed(futures):
            q, niche = futures[future]
            try:
                _, results = future.result()
            except Exception as e:
                print(f"  [search] {q!r} failed: {type(e).__name__}")
                results = []
            engines_ok += 1 if results else 0

            kept = 0
            for domain, url, engine in results:
                blocked = is_blocked_domain(domain)
                if blocked:
                    if not already_seen(conn, domain):
                        mark_seen(conn, domain, blocked, query=q)
                    continue
                if already_seen(conn, domain):
                    continue
                candidates.append((domain, url, engine, q, niche, city_from_query(q)))
                mark_seen(conn, domain, "queued", niche, q)   # claim it now
                kept += 1

            mark_query_run(conn, q, kept)
            conn.commit()
            print(f"  [search] {q!r} -> {len(results)} results, {kept} new (pool {len(candidates)})")
            if time.monotonic() > search_deadline:
                print("  [search] time budget reached, moving to scrape")
                break

    print(f"[+] {len(candidates)} new domains to inspect")

    # ---- scrape -----------------------------------------------------------
    leads, tally = [], {}
    if candidates:
        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS_SCRAPE) as pool:
            futures = {
                pool.submit(scrape_site, d, u, q, e, niche, city): d
                for d, u, e, q, niche, city in candidates
            }
            for future in concurrent.futures.as_completed(futures):
                domain = futures[future]
                try:
                    res = future.result()
                except Exception:
                    res = ScrapeResult(domain, "unreachable")
                tally[res.status] = tally.get(res.status, 0) + 1
                mark_seen(conn, domain, res.status, res.niche)
                if res.lead:
                    leads.append(res.lead)
                    print(f"  [keep] {res.lead.niche:<8} {domain} "
                          f"(score {res.lead.score}, {res.lead.email})")
                if time.monotonic() > deadline:
                    print("  [scrape] time budget reached, wrapping up")
                    break
        conn.commit()

    # ---- output -----------------------------------------------------------
    added = master_total = 0
    if leads:
        added, master_total = append_to_master(leads)
        write_daily(leads, OUTPUT_DIR / f"leads_{utcnow().strftime('%Y-%m-%d')}.xlsx")

    conn.commit()
    kept_total = conn.execute("SELECT COUNT(*) FROM seen WHERE status='qualified'").fetchone()[0]
    conn.close()

    elapsed = time.monotonic() - started
    breakdown = ", ".join(f"{k}={v}" for k, v in sorted(tally.items())) or "nothing inspected"
    print(f"\n[=] inspected: {breakdown}")
    print(f"[=] new leads this run: {added} | master workbook: {master_total} rows "
          f"| qualified all-time: {kept_total}")
    print(f"[=] elapsed {elapsed:.0f}s")

    if not candidates and engines_ok == 0:
        msg = "[!] every search engine returned nothing - likely rate limited or blocked"
        print(msg)
        if fail_on_empty:
            raise SystemExit(msg)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Continuous niche lead generator")
    parser.add_argument("--queries", type=int, default=QUERIES_PER_RUN,
                        help="how many search queries to work this run")
    parser.add_argument("--max-seconds", type=int, default=DEFAULT_MAX_SECONDS,
                        help="wall-clock budget so runs never overlap the schedule")
    parser.add_argument("--fail-on-empty", action="store_true",
                        help="exit non-zero if every search engine came back empty")
    args = parser.parse_args()
    main(args.queries, args.max_seconds, args.fail_on_empty)
