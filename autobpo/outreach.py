"""
Sends the initial cold email for a scored, enriched lead over SMTP, and
logs it so the daily run never contacts the same lead twice.
"""
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from config import SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, FROM_NAME
from db import get_conn

JOB_TEMPLATE = """Hi,

I noticed you're hiring for {title} - that role usually means someone on \
your team is stuck triaging tickets/calls/scheduling instead of doing \
higher-value work.

We run that whole desk for companies like yours as a fully outsourced, \
managed process - no new hire, no ramp time, live within a couple of weeks. \
Happy to send over how it'd work for a role like this one if useful.

Best,
{from_name}
"""

TENDER_TEMPLATE = """Hi,

I saw your notice for {title} - we run outsourced back-office / contact \
centre operations for organisations with exactly this kind of requirement, \
and can put together a response or a call whenever's useful.

Best,
{from_name}
"""


def build_message(lead):
    if lead["source"].startswith("uk_tender"):
        body = TENDER_TEMPLATE.format(title=lead["title"], from_name=FROM_NAME)
        subject = f"Re: {lead['title']}"
    else:
        body = JOB_TEMPLATE.format(title=lead["title"], from_name=FROM_NAME)
        subject = f"Quick note on your {lead['title']} opening"
    return subject, body


def send_email(to_email, subject, body):
    msg = MIMEMultipart()
    msg["From"] = f"{FROM_NAME} <{SMTP_USER}>"
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.sendmail(SMTP_USER, to_email, msg.as_string())


def run_outreach():
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM leads WHERE status='enriched' AND email IS NOT NULL"
        ).fetchall()
        for row in rows:
            lead = dict(row)
            subject, body = build_message(lead)
            try:
                send_email(lead["email"], subject, body)
                conn.execute(
                    "INSERT INTO outreach_log (lead_id, subject, to_email) VALUES (?,?,?)",
                    (lead["id"], subject, lead["email"]),
                )
                conn.execute("UPDATE leads SET status='contacted' WHERE id=?", (lead["id"],))
            except Exception as e:
                print(f"send failed for lead {lead['id']}: {e}")
