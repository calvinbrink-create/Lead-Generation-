import os

# ---- SMTP (outbound cold email) ----
# Reuses this repo's existing SMTP_USER / SMTP_PASS GitHub Actions secrets
# (already wired for the sales@growthsupplyhouse.com inbox used by
# outreach.py at the repo root) - no separate credentials needed for AutoBPO.
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "")       # your sending address
SMTP_PASS = os.environ.get("SMTP_PASS", "")       # app password, NOT your login password
FROM_NAME = os.environ.get("FROM_NAME", "Growth Supply House")

# ---- IMAP (reply detection - same inbox you send from) ----
IMAP_HOST = os.environ.get("IMAP_HOST", "imap.gmail.com")
IMAP_USER = os.environ.get("IMAP_USER", SMTP_USER)
IMAP_PASS = os.environ.get("IMAP_PASS", SMTP_PASS)

# ---- PayPal (Invoicing API v2) ----
# Free to create at developer.paypal.com - it's a REST app, not a paid
# platform. PayPal only takes its normal transaction fee when an invoice
# actually gets paid.
PAYPAL_CLIENT_ID = os.environ.get("PAYPAL_CLIENT_ID", "")
PAYPAL_SECRET = os.environ.get("PAYPAL_SECRET", "")
PAYPAL_MODE = os.environ.get("PAYPAL_MODE", "sandbox")  # "sandbox" while testing, "live" to go live
PAYPAL_BASE = (
    "https://api-m.sandbox.paypal.com" if PAYPAL_MODE == "sandbox"
    else "https://api-m.paypal.com"
)

# ---- TODO: put your real Auto BPO pricing here before going live ----
# Each tier needs a name, a flat monthly USD price, and a one-line label
# that becomes the invoice line item text. These numbers are placeholders.
PRICING_TIERS = {
    "starter": {
        "label": "Auto BPO - Starter (single process, business-hours coverage)",
        "price": 499.00,
    },
    "growth": {
        "label": "Auto BPO - Growth (multi-process, extended-hours coverage)",
        "price": 999.00,
    },
    "scale": {
        "label": "Auto BPO - Scale (full desk coverage, 24/5)",
        "price": 1799.00,
    },
}
DEFAULT_TIER = "starter"  # used when a reply agrees but doesn't name a tier

# ---- Role keywords: which open job postings count as a BPO buying signal ----
ROLE_KEYWORDS = [
    "customer support", "customer service", "support representative",
    "appointment setter", "scheduling coordinator", "call center", "call centre",
    "contact center", "contact centre", "virtual assistant", "back office",
    "data entry", "accounts receivable", "order processing", "help desk",
]

# ---- Tender keywords: which UK procurement notices are BPO-relevant ----
TENDER_KEYWORDS = [
    "call centre", "call center", "contact centre", "contact center",
    "back office", "outsourc", "customer service", "telephony", "bpo",
]

# ---- Reply classification (kept as free keyword rules - no paid LLM calls) ----
POSITIVE_REPLY_KEYWORDS = [
    "sounds good", "let's do it", "lets do it", "interested", "yes please",
    "go ahead", "sign me up", "send the invoice", "sounds great", "let's proceed",
]
NEGATIVE_REPLY_KEYWORDS = [
    "not interested", "unsubscribe", "remove me", "no thanks", "stop emailing",
]

DB_PATH = os.environ.get("DB_PATH", "output/autobpo.db")
