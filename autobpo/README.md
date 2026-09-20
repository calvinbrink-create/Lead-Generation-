# AutoBPO Pipeline

Two scheduled jobs, no manual steps once secrets are set:

1. **Daily discovery + outreach** (`main_discover.py`, `.github/workflows/autobpo-discover.yml`)
   — pulls BPO-relevant leads from public, keyless APIs (Greenhouse/Lever/Ashby
   job boards + UK Contracts Finder/Find a Tender procurement notices), scores
   them, finds a contact email on the company's own site, and sends the first
   cold email.
2. **Inbox + invoicing, every 2 hours** (`main_inbox.py`, `.github/workflows/autobpo-inbox.yml`)
   — checks the same inbox for replies, classifies them by keyword, and for
   anyone who replied positively, creates and sends a PayPal invoice
   automatically via the PayPal Invoicing API.

Everything is free-tier: no paid APIs, no paid platforms. Lives as a subfolder
of this repo rather than its own repo, sharing the same sales@growthsupplyhouse.com
Gmail credentials this repo's root-level `outreach.py` already uses.

## One-time setup (Calvin does this — can't be automated)

- **SMTP/IMAP**: already done. This reuses the repo's existing `SMTP_USER` /
  `SMTP_PASS` Actions secrets (the same sales@growthsupplyhouse.com Gmail app
  password already wired for the root `outreach.py`). No new Gmail setup
  needed unless a different inbox should be used for AutoBPO specifically.
- **PayPal REST app**: developer.paypal.com → Apps & Credentials → Create App.
  Add the Client ID and Secret as repo secrets `PAYPAL_CLIENT_ID` /
  `PAYPAL_SECRET` (Settings → Secrets and variables → Actions → **Secrets**
  tab). Test everything with the default sandbox mode first (sandbox invoices
  go nowhere real). Switch to live only by adding a repo **variable**
  (Actions → **Variables** tab, not Secrets) named `PAYPAL_MODE` set to
  `live`, once a real invoice has been sent and paid successfully in sandbox.
- **Real pricing**: `config.py` currently has placeholder Auto BPO tiers and
  prices. Edit `PRICING_TIERS` and `DEFAULT_TIER` before switching to live —
  this is the one thing here that was invented and needs a human decision.
- **Seed the company list**: the starter `sources/company_slugs.json` only
  has three throwaway example companies (Stripe, Figma, etc. — not real
  prospects). Run `python3 sources/discover_slugs.py` once locally (or as a
  manual Actions run) to populate it with real small/mid-size US/UK companies
  before the first live outreach run, and re-run it weekly to keep growing
  coverage.

## Known gaps to be upfront about

- Reply classification is keyword-based, not a real language model — it will
  misfire on ambiguous replies ("maybe next quarter" won't match either list
  and stays "neutral," which is safe; a sarcastic "yeah right, sounds great"
  would wrongly fire positive). Spot-check the `replies` table periodically,
  especially early on.
- Auto-invoicing fires the moment a reply matches a positive keyword, with no
  confirmation step and no verification that the reply-sender actually has
  authority to agree to a contract. That is what "fully automated" requires,
  but it also means a wrong keyword match turns directly into a real PayPal
  invoice in someone's inbox once `PAYPAL_MODE=live`. Worth watching closely
  during the first few weeks.
- Email enrichment guesses a domain from the company name when the source
  data doesn't already include one, and that guess will often be wrong for
  anything but a simple, exact-match name.
- `sources/uk_procurement.py` is marked in-file as needing verification
  against a live Find a Tender response — government API response envelopes
  occasionally change shape.

## Local run

```bash
cd autobpo
pip install -r requirements.txt
cp .env.example .env   # fill in real values, never commit this file
python3 -m py_compile *.py sources/*.py   # sanity check
python main_discover.py
python main_inbox.py
```

## Database

`output/autobpo.db` (SQLite) tracks `leads`, `outreach_log`, `replies`, and
`invoices`. Both workflows commit it back to the repo after each run, same
pattern this repo's root-level `outreach.yml` already uses for its own state.
