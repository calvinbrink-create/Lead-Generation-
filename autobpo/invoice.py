"""
Creates and sends a PayPal invoice for any lead marked replied_positive that
doesn't already have one. Requires a PayPal REST app (free to create at
developer.paypal.com - you only pay PayPal's normal transaction fee when an
invoice actually gets paid) with PAYPAL_CLIENT_ID / PAYPAL_SECRET set.

Uses PayPal's Invoicing API v2: create draft -> send.
https://developer.paypal.com/api/invoicing
"""
import requests
from config import PAYPAL_CLIENT_ID, PAYPAL_SECRET, PAYPAL_BASE, PRICING_TIERS, DEFAULT_TIER, FROM_NAME
from db import get_conn


def get_access_token():
    r = requests.post(
        f"{PAYPAL_BASE}/v1/oauth2/token",
        auth=(PAYPAL_CLIENT_ID, PAYPAL_SECRET),
        data={"grant_type": "client_credentials"},
        headers={"Accept": "application/json"},
        timeout=20,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def create_and_send_invoice(token, recipient_email, tier_key):
    tier = PRICING_TIERS.get(tier_key, PRICING_TIERS[DEFAULT_TIER])
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {
        "detail": {"currency_code": "USD", "note": "Thanks for your business"},
        "invoicer": {"name": {"given_name": FROM_NAME}},
        "primary_recipients": [{"billing_info": {"email_address": recipient_email}}],
        "items": [{
            "name": tier["label"],
            "quantity": "1",
            "unit_amount": {"currency_code": "USD", "value": f"{tier['price']:.2f}"},
        }],
    }
    r = requests.post(f"{PAYPAL_BASE}/v2/invoicing/invoices", json=payload, headers=headers, timeout=20)
    r.raise_for_status()
    body = r.json() if r.content else {}
    invoice_id = body.get("id") or r.headers.get("Location", "").rstrip("/").split("/")[-1]

    send_r = requests.post(
        f"{PAYPAL_BASE}/v2/invoicing/invoices/{invoice_id}/send",
        headers=headers, json={"send_to_recipient": True}, timeout=20,
    )
    send_r.raise_for_status()
    return invoice_id, tier["price"]


def run_invoicing(tier_key=None):
    token = get_access_token()
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT l.* FROM leads l
            LEFT JOIN invoices i ON i.lead_id = l.id
            WHERE l.status='replied_positive' AND i.id IS NULL
        """).fetchall()
        for row in rows:
            lead = dict(row)
            try:
                invoice_id, amount = create_and_send_invoice(
                    token, lead["email"], tier_key or DEFAULT_TIER
                )
                conn.execute(
                    "INSERT INTO invoices (lead_id, paypal_invoice_id, tier, amount) VALUES (?,?,?,?)",
                    (lead["id"], invoice_id, tier_key or DEFAULT_TIER, amount),
                )
                conn.execute("UPDATE leads SET status='invoiced' WHERE id=?", (lead["id"],))
            except Exception as e:
                print(f"invoice failed for lead {lead['id']}: {e}")
