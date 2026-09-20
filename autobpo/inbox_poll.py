"""
Polls the same inbox outreach sends from, matches replies back to a lead by
sender email, and classifies each as positive/negative/neutral by keyword.
Runs more often than the daily discovery job - e.g. every 1-2 hours -
since a reply sitting unread for a day is a lead going cold.
"""
import imaplib
import email
from config import (
    IMAP_HOST, IMAP_USER, IMAP_PASS,
    POSITIVE_REPLY_KEYWORDS, NEGATIVE_REPLY_KEYWORDS,
)
from db import get_conn


def classify(body):
    text = body.lower()
    if any(k in text for k in NEGATIVE_REPLY_KEYWORDS):
        return "negative"
    if any(k in text for k in POSITIVE_REPLY_KEYWORDS):
        return "positive"
    return "neutral"


def get_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                payload = part.get_payload(decode=True)
                return payload.decode(errors="ignore") if payload else ""
        return ""
    payload = msg.get_payload(decode=True)
    return payload.decode(errors="ignore") if payload else ""


def poll_inbox():
    conn_imap = imaplib.IMAP4_SSL(IMAP_HOST)
    conn_imap.login(IMAP_USER, IMAP_PASS)
    conn_imap.select("INBOX")
    status, data = conn_imap.search(None, "UNSEEN")
    ids = data[0].split() if data and data[0] else []

    with get_conn() as conn:
        for msg_id in ids:
            _, msg_data = conn_imap.fetch(msg_id, "(RFC822)")
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)
            from_addr = email.utils.parseaddr(msg.get("From"))[1]
            body = get_body(msg)
            classification = classify(body)

            lead = conn.execute(
                "SELECT * FROM leads WHERE email=? ORDER BY id DESC LIMIT 1", (from_addr,)
            ).fetchone()
            if not lead:
                continue

            conn.execute(
                "INSERT INTO replies (lead_id, body, classification) VALUES (?,?,?)",
                (lead["id"], body, classification),
            )
            if classification == "positive":
                conn.execute("UPDATE leads SET status='replied_positive' WHERE id=?", (lead["id"],))
            elif classification == "negative":
                conn.execute("UPDATE leads SET status='replied_negative' WHERE id=?", (lead["id"],))

    conn_imap.logout()
