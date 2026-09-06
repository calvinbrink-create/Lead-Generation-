# Lead Generation Pipeline

Automated daily lead generator for home-services businesses. Searches multiple
engines, scrapes candidate business sites, extracts contact details, scores each
site, and writes a dated Excel file of qualified leads.

Runs unattended on GitHub Actions. No paid APIs, no external services.

## How it works

1. **Search** — builds a query for every `NICHE` x `CITY` pair and runs each
   across DuckDuckGo, Bing and Google concurrently. Aggregator and social
   domains (Yelp, Angi, Facebook, ...) are filtered out.
2. **Dedupe** — `seen_domains.sqlite3` tracks every domain ever scraped, so a
   business is only ever reported once across all runs.
3. **Scrape** — fetches each candidate site plus its contact page, pulling
   email, phone and a best-guess owner/contact name.
4. **Score** — 0-100 on HTTPS, mobile viewport, title, contact info, phone,
   meta description and presence of a form.
5. **Export** — writes `output/leads_<YYYY-MM-DD>.xlsx`, sorted by score
   descending. Sites with no discoverable email are dropped.

## Schedule

The workflow runs daily at **11:00 UTC** (`0 11 * * *`) and can also be started
by hand from the Actions tab via **Run workflow**.

Each run commits the new Excel file and the updated dedupe database back to the
repository, and uploads the spreadsheet as a workflow artifact with 30-day
retention.

> Scheduled workflows only fire from the repository's **default branch**.

## Tuning output volume

The daily target is 200 leads (`--target` in the workflow, `TARGET_LEADS` in the
script). A run finishing below target is a data-availability outcome, not a
failure — the fix is a wider net, not a code change:

- add entries to `CITIES` and `NICHES` in `leadgen.py`

Free-tier scraping typically finds usable contact details on 30-50% of sites
visited, which is why the pipeline over-fetches roughly 4x the target pool.

## Local run

```bash
pip install -r requirements.txt
python leadgen.py --target 25
```

## Files

| Path | Purpose |
| --- | --- |
| `leadgen.py` | The pipeline: search, scrape, score, export |
| `requirements.txt` | `requests`, `beautifulsoup4`, `openpyxl` |
| `.github/workflows/daily-leadgen.yml` | Daily schedule + manual dispatch |
| `output/` | Dated Excel exports, committed by each run |
| `seen_domains.sqlite3` | Cross-run dedupe state (created on first run) |

## Notes

- Google is the least reliable of the three engines and will intermittently
  return nothing when it serves a CAPTCHA. This is handled gracefully; Bing and
  DuckDuckGo carry the run.
- Lead exports contain business contact details. Keep this repository
  **private** if that data shouldn't be public.
