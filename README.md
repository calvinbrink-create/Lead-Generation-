# Lead Generation Pipeline

Continuously hunts US home-services businesses, verifies each site is genuinely
in one of the target trades, pulls contact details, and appends new leads to a
cumulative Excel workbook.

Runs unattended on GitHub Actions every ~10 minutes. No paid APIs.

## Niches

Only these trades are kept. Everything else is discarded:

| Niche | Verified by |
| --- | --- |
| **Plumbing** | plumber / plumbing / rooter + drain, sewer, water heater, repipe, leak... |
| **HVAC** | hvac / air conditioning / heating and cooling + furnace, heat pump, duct... |
| **Roofing** | roofing / roofer / roof repair + shingle, metal roof, gutter, flashing... |

A site must show at least one trade term **and** two supporting vocabulary terms
before it counts. Mentioning "we hired a plumber once" is not enough.

## Coverage

**411 cities across all 50 states + DC**, times 18 query phrasings = **7,398
distinct searches**. Each run takes the 12 least-recently-searched queries, so
coverage rotates continuously across the whole country rather than hammering the
same cities.

## What gets thrown away

- **Adult/explicit sites** — matched on domain and page text. Ambiguous terms are
  matched carefully so real businesses survive: `essexplumbing.com`,
  `sussexhvac.com` and `weatherstrippingpros.com` are not false-flagged.
- **Directories and marketplaces** — Yelp, Angi, Thumbtack, HomeAdvisor,
  Expertise, BuildZoom, Porch, Houzz and ~100 more, plus any page using
  directory language ("compare quotes", "get matched", "top 10 best").
- **Manufacturers and suppliers** — Carrier, Trane, GAF, Owens Corning, Kohler...
- **Sites with no discoverable email.**
- **Anything already visited** — every domain is recorded with its verdict, so no
  domain is ever inspected twice.

## Output

| File | Contents |
| --- | --- |
| `output/leads_master.xlsx` | **The deliverable.** Every lead ever found, accumulating. |
| `output/leads_<date>.xlsx` | That day's additions only. |
| `seen_domains.sqlite3` | Every domain visited + verdict, and query rotation state. |

Columns: Business, Niche, Domain, URL, Owner/Contact, Email, Phone, City, Score,
Search Query, Engine, Date Found.

Emails are ranked, not taken at random: an address on the business's own domain
beats an off-domain one, `info@`/`contact@`/`office@` beat a personal alias, and
`noreply@` sinks to the bottom. Asset filenames that look like addresses
(`flags-sprite@2x.webp`) are rejected.

## Schedule

`*/10 * * * *` — every 10 minutes, plus **Run workflow** on the Actions tab.

GitHub's minimum for scheduled workflows is 5 minutes and schedules are
best-effort under load, so in practice expect every 10-20 minutes. Each run is
budgeted to 420s so it finishes before the next is due, and a `concurrency` group
queues runs rather than overlapping them. A run that finds nothing commits
nothing.

## Tuning

- **More/other trades** — add to `NICHE_PROFILES` in `leadgen.py`. Each entry needs
  `queries` (search phrasings), `strong` (identifies the trade) and `supporting`
  (working vocabulary).
- **Stricter or looser matching** — `MIN_STRONG_HITS` / `MIN_SUPPORT_HITS`.
- **More ground per run** — raise `--queries` and `--max-seconds` in the workflow,
  keeping the budget under the cron interval.
- **More cities** — extend `CITIES`.

## Local run

```bash
pip install -r requirements.txt
python leadgen.py --queries 5 --max-seconds 120
```

`--fail-on-empty` exits non-zero when every search engine returns nothing, which
is the signal that engines are blocking rather than that leads ran out.

## Known limits

- **Search engines are the bottleneck**, not the query list. They rate-limit
  datacenter IPs, and GitHub's runners are datacenter IPs. Each engine gets
  retries with backoff and four engines are tried (DuckDuckGo, Bing, Mojeek,
  Google), but some runs will return little or nothing. That is throttling, not
  a bug, and adding cities will not fix it.
- **Free-tier contact extraction** finds a usable email on roughly a third to a
  half of sites visited.
- Lead exports contain business contact details. Keep this repository **private**
  if that data should not be public.
