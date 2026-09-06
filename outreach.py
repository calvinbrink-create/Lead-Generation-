#!/usr/bin/env python3
"""
Growth Supply House - Audit + Outreach

Two stages over the leads leadgen.py collects:

  audit  - run each captured website through the Growth Supply House audit tool
           and save the returned PDF, recording it against the lead
  send    - email the owner a personalised note with their audit attached

Nothing is sent unless every check passes: the PDF must exist, be a real PDF,
belong to that exact business, and the address must not already have been
contacted or opted out. Sending is a dry run unless OUTREACH_LIVE=1.
"""

import argparse
import mimetypes
import os
import re
import smtplib
import sqlite3
import ssl
import sys
import time
import urllib.parse
from dataclasses import dataclass
from datetime import datetime, timezone, date
from email.message import EmailMessage
from pathlib import Path

import requests
from openpyxl import load_workbook

DB_PATH = Path("seen_domains.sqlite3")
OUTPUT_DIR = Path("output")
AUDIT_DIR = Path("audits")
MASTER_XLSX = OUTPUT_DIR / "leads_master.xlsx"

# --- audit tool -------------------------------------------------------------
# The endpoint is configurable so the exact shape of the audit service can be
# set without editing code. AUDIT_METHOD is GET or POST; AUDIT_FIELD is the
# parameter carrying the website being audited.
AUDIT_ENDPOINT = os.environ.get("AUDIT_ENDPOINT", "https://audit.growthsupplyhouse.com/api/audit")
AUDIT_METHOD = os.environ.get("AUDIT_METHOD", "POST").upper()
AUDIT_FIELD = os.environ.get("AUDIT_FIELD", "url")
AUDIT_NAME_FIELD = os.environ.get("AUDIT_NAME_FIELD", "business")
AUDIT_TIMEOUT = int(os.environ.get("AUDIT_TIMEOUT", "90"))
AUDIT_POLL_SECONDS = int(os.environ.get("AUDIT_POLL_SECONDS", "5"))
AUDIT_POLL_ATTEMPTS = int(os.environ.get("AUDIT_POLL_ATTEMPTS", "12"))

# --- mail -------------------------------------------------------------------
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "")          # e.g. sales@growthsupplyhouse.com
SMTP_PASS = os.environ.get("SMTP_PASS", "")          # Gmail app password
FROM_NAME = os.environ.get("FROM_NAME", "Growth Supply House")
REPLY_TO = os.environ.get("REPLY_TO", SMTP_USER)
POSTAL_ADDRESS = os.environ.get("POSTAL_ADDRESS", "")
SITE_URL = os.environ.get("SITE_URL", "https://www.growthsupplyhouse.com")
DAILY_SEND_CAP = int(os.environ.get("DAILY_SEND_CAP", "4"))
LIVE = os.environ.get("OUTREACH_LIVE", "") == "1"

FREE_MAIL_HOSTS = {
    "gmail.com", "googlemail.com", "yahoo.com", "ymail.com", "outlook.com",
    "hotmail.com", "live.com", "msn.com", "aol.com", "icloud.com", "me.com",
    "comcast.net", "att.net", "verizon.net", "sbcglobal.net", "bellsouth.net",
    "cox.net", "charter.net", "protonmail.com", "proton.me", "gmx.com",
}

STOPWORDS = {
    "the", "and", "llc", "inc", "co", "company", "corp", "group", "services",
    "service", "solutions", "heating", "cooling", "air", "plumbing", "roofing",
    "plumbers", "roofers", "hvac", "contractors", "contractor", "of", "your",
    "best", "home", "us", "usa", "pros", "pro", "team", "sons", "son",
}

EMAIL_OK_RE = re.compile(r"^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$")


def utcnow():
    return datetime.now(timezone.utc)


def log(msg):
    print(msg, flush=True)


# ---------------------------------------------------------------------------
# STATE
# ---------------------------------------------------------------------------

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS outreach (
        domain       TEXT PRIMARY KEY,
        business     TEXT,
        email        TEXT,
        website      TEXT,
        audit_pdf    TEXT,
        audit_status TEXT,
        audit_at     TEXT,
        sent_at      TEXT,
        send_status  TEXT,
        message_id   TEXT
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS suppression (
        email    TEXT PRIMARY KEY,
        reason   TEXT,
        added_at TEXT
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS send_log (
        id        INTEGER PRIMARY KEY AUTOINCREMENT,
        day       TEXT,
        email     TEXT,
        domain    TEXT,
        sent_at   TEXT
    )""")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_send_log_day ON send_log(day)")
    conn.commit()
    return conn


def is_suppressed(conn, email):
    return conn.execute(
        "SELECT 1 FROM suppression WHERE email = ?", (email.lower(),)
    ).fetchone() is not None


def suppress(conn, email, reason):
    conn.execute(
        "INSERT OR IGNORE INTO suppression (email, reason, added_at) VALUES (?,?,?)",
        (email.lower(), reason, utcnow().isoformat()),
    )
    conn.commit()


def already_contacted(conn, domain, email):
    """A send is blocked if this domain OR this address has ever been mailed.
    Two leads can share an address; the person still only hears from us once."""
    row = conn.execute(
        "SELECT 1 FROM outreach WHERE (domain = ? OR email = ?) AND sent_at IS NOT NULL",
        (domain.lower(), email.lower()),
    ).fetchone()
    if row:
        return True
    return conn.execute(
        "SELECT 1 FROM send_log WHERE email = ? OR domain = ?",
        (email.lower(), domain.lower()),
    ).fetchone() is not None


def sent_today(conn):
    return conn.execute(
        "SELECT COUNT(*) FROM send_log WHERE day = ?", (date.today().isoformat(),)
    ).fetchone()[0]


def record_send(conn, domain, email, message_id, status):
    now = utcnow().isoformat()
    conn.execute(
        """UPDATE outreach SET sent_at = ?, send_status = ?, message_id = ?
           WHERE domain = ?""",
        (now, status, message_id, domain.lower()),
    )
    if status == "sent":
        conn.execute(
            "INSERT INTO send_log (day, email, domain, sent_at) VALUES (?,?,?,?)",
            (date.today().isoformat(), email.lower(), domain.lower(), now),
        )
    conn.commit()


# ---------------------------------------------------------------------------
# WORKBOOK
# ---------------------------------------------------------------------------

AUDIT_COLUMNS = ["Audit PDF", "Audit Status", "Emailed At", "Email Status"]


@dataclass
class LeadRow:
    row: int
    business: str
    niche: str
    domain: str
    website: str
    owner: str
    email: str
    phone: str
    city: str


def _header_map(ws):
    return {str(c.value).strip(): c.column for c in ws[1] if c.value}


def ensure_audit_columns(ws):
    """Add the audit/outreach columns to the workbook if they are not there."""
    headers = _header_map(ws)
    added = []
    for name in AUDIT_COLUMNS:
        if name not in headers:
            col = ws.max_column + 1
            ws.cell(row=1, column=col, value=name)
            ws.cell(row=1, column=col).font = ws.cell(row=1, column=1).font.copy()
            ws.cell(row=1, column=col).fill = ws.cell(row=1, column=1).fill.copy()
            ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 34
            headers[name] = col
            added.append(name)
    return headers, added


def read_leads(ws, headers):
    out = []
    need = ("Business", "Domain", "URL", "Email")
    if not all(k in headers for k in need):
        raise SystemExit(f"workbook is missing required columns: {need}")
    for r in range(2, ws.max_row + 1):
        get = lambda k: (ws.cell(row=r, column=headers[k]).value if k in headers else "") or ""
        domain = str(get("Domain")).strip().lower()
        if not domain:
            continue
        out.append(LeadRow(
            row=r,
            business=str(get("Business")).strip(),
            niche=str(get("Niche")).strip(),
            domain=domain,
            website=str(get("URL")).strip(),
            owner=str(get("Owner/Contact")).strip(),
            email=str(get("Email")).strip().lower(),
            phone=str(get("Phone")).strip(),
            city=str(get("City")).strip(),
        ))
    return out


# ---------------------------------------------------------------------------
# VERIFICATION
# ---------------------------------------------------------------------------

def domain_root(host):
    """Strip www and the public suffix down to the registrable-ish label."""
    host = (host or "").lower().strip()
    if host.startswith("www."):
        host = host[4:]
    parts = [p for p in host.split(".") if p]
    if len(parts) >= 3 and parts[-2] in ("co", "com", "net", "org", "gov", "ac"):
        return parts[-3]
    return parts[-2] if len(parts) >= 2 else (parts[0] if parts else "")


def tokens(text):
    raw = re.split(r"[^a-z0-9]+", (text or "").lower())
    return {t for t in raw if len(t) > 2 and t not in STOPWORDS}


def split_concatenated(label, vocab):
    """Domains are written without separators. Find which vocabulary words the
    label actually contains, so 'mesquiteheatingandair' matches 'Mesquite'."""
    return {w for w in vocab if len(w) > 3 and w in label}


def business_email_match(business, domain, email):
    """Return (ok, confidence, why). Guards against mailing an address that
    belongs to a different company than the lead row claims."""
    email = (email or "").lower().strip()
    if not EMAIL_OK_RE.match(email):
        return False, "invalid", "not a valid address"

    host = email.split("@", 1)[1]
    site_root = domain_root(domain)
    host_root = domain_root(host)

    if not site_root:
        return False, "invalid", "lead has no usable domain"

    # Best case: the address lives on the business's own domain.
    if host == domain.lower() or host.endswith("." + domain.lower()) or host_root == site_root:
        return True, "own-domain", f"{host} is the business domain"

    vocab = tokens(business) | tokens(domain.replace(".", " "))
    local = email.split("@", 1)[0]

    if host_root in FREE_MAIL_HOSTS or host in FREE_MAIL_HOSTS:
        hits = tokens(local) & vocab
        hits |= split_concatenated(re.sub(r"[^a-z0-9]", "", local), vocab)
        hits |= split_concatenated(re.sub(r"[^a-z0-9]", "", local), {site_root})
        if hits:
            return True, "freemail-named", f"free mailbox naming the business ({', '.join(sorted(hits))})"
        return False, "freemail-unnamed", (
            f"free mailbox '{local}@{host}' with nothing tying it to "
            f"'{business or domain}'"
        )

    # A different company's domain entirely - this is the case that produces
    # embarrassing sends, so it is rejected rather than guessed at.
    shared = split_concatenated(host_root, vocab) | (tokens(host_root) & vocab)
    if shared or site_root in host_root or host_root in site_root:
        return True, "related-domain", f"{host} shares naming with {domain}"
    return False, "mismatch", f"{host} belongs to a different domain than {domain}"


def verify_pdf(path, min_bytes=1024):
    """The attachment must exist, be non-trivial, and actually be a PDF."""
    if not path:
        return False, "no audit PDF recorded"
    p = Path(path)
    if not p.exists():
        return False, f"audit PDF missing on disk: {p}"
    size = p.stat().st_size
    if size < min_bytes:
        return False, f"audit PDF is only {size} bytes"
    try:
        with p.open("rb") as fh:
            if fh.read(5) != b"%PDF-":
                return False, "file is not a PDF (bad magic bytes)"
            fh.seek(-1400 if size > 1400 else -size, os.SEEK_END)
            if b"%%EOF" not in fh.read():
                return False, "PDF looks truncated (no %%EOF)"
    except Exception as e:
        return False, f"could not read PDF: {type(e).__name__}"
    return True, f"{size} bytes"


def audit_filename(domain):
    safe = re.sub(r"[^a-z0-9.-]+", "_", domain.lower())
    return AUDIT_DIR / f"audit_{safe}.pdf"


def verify_pdf_belongs_to(path, domain):
    """The saved audit is named for its domain; confirm the file we are about
    to attach is the one generated for this exact lead."""
    expected = audit_filename(domain).name
    return Path(path).name == expected, expected


# ---------------------------------------------------------------------------
# AUDIT
# ---------------------------------------------------------------------------

def _looks_like_pdf(resp):
    ctype = resp.headers.get("content-type", "").lower()
    return "pdf" in ctype or resp.content[:5] == b"%PDF-"


def fetch_audit(website, business):
    """Run one site through the audit tool and return PDF bytes.

    Handles the three shapes an audit service normally takes: the PDF straight
    back in the response, a JSON body carrying a link to it, or a job that has
    to be polled. Endpoint and field names come from env vars so the exact
    contract can be set without a code change.
    """
    payload = {AUDIT_FIELD: website}
    if AUDIT_NAME_FIELD:
        payload[AUDIT_NAME_FIELD] = business
    headers = {"Accept": "application/pdf, application/json;q=0.9, */*;q=0.5"}

    if AUDIT_METHOD == "GET":
        r = requests.get(AUDIT_ENDPOINT, params=payload, headers=headers, timeout=AUDIT_TIMEOUT)
    else:
        r = requests.post(AUDIT_ENDPOINT, json=payload, headers=headers, timeout=AUDIT_TIMEOUT)
    r.raise_for_status()

    if _looks_like_pdf(r):
        return r.content

    # JSON: either a direct link, or a job to poll.
    try:
        data = r.json()
    except ValueError:
        raise RuntimeError(
            f"audit endpoint returned {r.headers.get('content-type', 'unknown')} "
            f"({len(r.content)} bytes), neither PDF nor JSON"
        )

    link = None
    for key in ("pdf_url", "pdfUrl", "url", "download_url", "downloadUrl",
                "report_url", "reportUrl", "file", "link", "result"):
        val = data.get(key) if isinstance(data, dict) else None
        if isinstance(val, str) and val.startswith("http"):
            link = val
            break

    if link is None and isinstance(data, dict):
        job = data.get("id") or data.get("job_id") or data.get("jobId")
        status_url = data.get("status_url") or data.get("statusUrl")
        if job or status_url:
            poll = status_url or urllib.parse.urljoin(AUDIT_ENDPOINT.rstrip("/") + "/", str(job))
            for _ in range(AUDIT_POLL_ATTEMPTS):
                time.sleep(AUDIT_POLL_SECONDS)
                pr = requests.get(poll, headers=headers, timeout=AUDIT_TIMEOUT)
                if _looks_like_pdf(pr):
                    return pr.content
                try:
                    pdata = pr.json()
                except ValueError:
                    continue
                for key in ("pdf_url", "pdfUrl", "url", "download_url", "reportUrl"):
                    if isinstance(pdata.get(key), str) and pdata[key].startswith("http"):
                        link = pdata[key]
                        break
                if link:
                    break

    if not link:
        raise RuntimeError(f"no PDF or PDF link in audit response: {str(data)[:200]}")

    pr = requests.get(link, headers=headers, timeout=AUDIT_TIMEOUT)
    pr.raise_for_status()
    if not _looks_like_pdf(pr):
        raise RuntimeError(f"audit link {link} did not return a PDF")
    return pr.content


def run_audits(conn, leads, headers, ws, limit):
    AUDIT_DIR.mkdir(exist_ok=True)
    done = skipped = failed = 0

    for lead in leads:
        if done + failed >= limit:
            break
        row = conn.execute(
            "SELECT audit_pdf, audit_status FROM outreach WHERE domain = ?", (lead.domain,)
        ).fetchone()

        conn.execute(
            """INSERT INTO outreach (domain, business, email, website)
               VALUES (?,?,?,?)
               ON CONFLICT(domain) DO UPDATE SET
                   business = excluded.business,
                   email    = excluded.email,
                   website  = excluded.website""",
            (lead.domain, lead.business, lead.email, lead.website),
        )

        path = audit_filename(lead.domain)
        if row and row[1] == "ok":
            ok, _ = verify_pdf(row[0])
            if ok:
                skipped += 1
                _write_audit_cells(ws, headers, lead.row, row[0], "ok")
                continue

        log(f"  [audit] {lead.domain} ...")
        try:
            pdf = fetch_audit(lead.website or f"https://{lead.domain}", lead.business)
            path.write_bytes(pdf)
            ok, detail = verify_pdf(path)
            if not ok:
                raise RuntimeError(detail)
            conn.execute(
                "UPDATE outreach SET audit_pdf=?, audit_status='ok', audit_at=? WHERE domain=?",
                (str(path), utcnow().isoformat(), lead.domain),
            )
            _write_audit_cells(ws, headers, lead.row, str(path), "ok")
            log(f"           saved {path.name} ({detail})")
            done += 1
        except Exception as e:
            reason = f"{type(e).__name__}: {e}"[:180]
            if path.exists() and path.stat().st_size < 1024:
                path.unlink()
            conn.execute(
                "UPDATE outreach SET audit_status=?, audit_at=? WHERE domain=?",
                (f"failed: {reason}", utcnow().isoformat(), lead.domain),
            )
            _write_audit_cells(ws, headers, lead.row, "", f"failed: {reason}")
            log(f"           FAILED {reason}")
            failed += 1
        conn.commit()

    conn.commit()
    return done, skipped, failed


def _write_audit_cells(ws, headers, row, pdf_path, status):
    if "Audit PDF" in headers:
        ws.cell(row=row, column=headers["Audit PDF"], value=pdf_path)
    if "Audit Status" in headers:
        ws.cell(row=row, column=headers["Audit Status"], value=status)


# ---------------------------------------------------------------------------
# EMAIL
# ---------------------------------------------------------------------------

def first_name(owner):
    owner = (owner or "").strip()
    if not owner:
        return ""
    part = re.split(r"[,\s]+", owner)[0]
    return part if re.match(r"^[A-Z][a-z]{1,20}$", part) else ""


def build_email(lead):
    who = first_name(lead.owner)
    greeting = f"Hi {who}," if who else "Hi there,"
    biz = lead.business or lead.domain

    subject = f"A quick website idea for {biz}"[:120]
    body = f"""{greeting}

I reviewed the current {biz} website and put together a free redesign concept
focused on clearer services, stronger trust signals, and an easier path to
contact you.

I've attached the personalised audit PDF for {biz} so you can see what I
noticed. If it is useful, would you be open to a quick look?

You can view our website here: {SITE_URL}

If you do not want to receive further emails from Growth Supply House, reply
"unsubscribe" and I will remove you from future outreach.

Best,
Growth Supply House
{SMTP_USER}"""
    if POSTAL_ADDRESS:
        body += f"\n{POSTAL_ADDRESS}"
    return subject, body


def send_email(to_addr, subject, body, attachment):
    msg = EmailMessage()
    msg["From"] = f"{FROM_NAME} <{SMTP_USER}>"
    msg["To"] = to_addr
    msg["Subject"] = subject
    if REPLY_TO:
        msg["Reply-To"] = REPLY_TO
    msg["List-Unsubscribe"] = f"<mailto:{SMTP_USER}?subject=unsubscribe>"
    msg.set_content(body)

    data = Path(attachment).read_bytes()
    msg.add_attachment(data, maintype="application", subtype="pdf",
                       filename=Path(attachment).name)

    ctx = ssl.create_default_context()
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as s:
        s.ehlo()
        s.starttls(context=ctx)
        s.login(SMTP_USER, SMTP_PASS)
        s.send_message(msg)
    return msg["Message-ID"] or ""


def send_batch(conn, leads, headers, ws, cap):
    already = sent_today(conn)
    remaining = max(0, cap - already)
    log(f"[+] daily cap {cap}, already sent today {already}, room for {remaining}")
    if remaining == 0:
        return 0, []

    sent, rejected = 0, []
    for lead in leads:
        if sent >= remaining:
            break

        row = conn.execute(
            "SELECT audit_pdf, audit_status FROM outreach WHERE domain = ?", (lead.domain,)
        ).fetchone()
        pdf_path = row[0] if row else ""

        # --- every gate must pass -----------------------------------------
        if not lead.email:
            rejected.append((lead.domain, "no email captured")); continue
        if is_suppressed(conn, lead.email):
            rejected.append((lead.domain, "address opted out")); continue
        if already_contacted(conn, lead.domain, lead.email):
            continue                                    # silent: normal steady state
        ok, detail = verify_pdf(pdf_path)
        if not ok:
            rejected.append((lead.domain, f"audit: {detail}")); continue
        belongs, expected = verify_pdf_belongs_to(pdf_path, lead.domain)
        if not belongs:
            rejected.append((lead.domain, f"PDF is not this lead's (expected {expected})")); continue
        match_ok, confidence, why = business_email_match(lead.business, lead.domain, lead.email)
        if not match_ok:
            rejected.append((lead.domain, f"email/business mismatch - {why}")); continue

        subject, body = build_email(lead)
        log(f"  [send] {lead.email} <- {lead.business or lead.domain} "
            f"[{confidence}: {why}] + {Path(pdf_path).name}")

        if not LIVE:
            rejected.append((lead.domain, "DRY RUN - would have sent"))
            continue

        try:
            mid = send_email(lead.email, subject, body, pdf_path)
            record_send(conn, lead.domain, lead.email, mid, "sent")
            if "Emailed At" in headers:
                ws.cell(row=lead.row, column=headers["Emailed At"], value=utcnow().strftime("%Y-%m-%d %H:%M"))
            if "Email Status" in headers:
                ws.cell(row=lead.row, column=headers["Email Status"], value=f"sent ({confidence})")
            sent += 1
            time.sleep(random_gap())
        except Exception as e:
            reason = f"{type(e).__name__}: {e}"[:180]
            record_send(conn, lead.domain, lead.email, "", f"failed: {reason}")
            if "Email Status" in headers:
                ws.cell(row=lead.row, column=headers["Email Status"], value=f"failed: {reason}")
            rejected.append((lead.domain, f"send failed - {reason}"))
    return sent, rejected


def random_gap():
    import random
    return random.uniform(20, 60)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def load_workbook_or_exit():
    if not MASTER_XLSX.exists():
        raise SystemExit(f"{MASTER_XLSX} does not exist yet - run leadgen.py first")
    wb = load_workbook(MASTER_XLSX)
    ws = wb.active
    headers, added = ensure_audit_columns(ws)
    if added:
        log(f"[+] added workbook columns: {', '.join(added)}")
    return wb, ws, headers


def cmd_audit(args):
    conn = init_db()
    wb, ws, headers = load_workbook_or_exit()
    leads = read_leads(ws, headers)
    log(f"[+] {len(leads)} leads in workbook; auditing up to {args.limit}")
    done, skipped, failed = run_audits(conn, leads, headers, ws, args.limit)
    wb.save(MASTER_XLSX)
    conn.close()
    log(f"\n[=] audits: {done} new, {skipped} already had one, {failed} failed")
    if failed and not done:
        raise SystemExit("every audit attempt failed - check AUDIT_ENDPOINT")


def cmd_send(args):
    conn = init_db()
    wb, ws, headers = load_workbook_or_exit()
    leads = read_leads(ws, headers)

    if not LIVE:
        log("[!] DRY RUN - set OUTREACH_LIVE=1 to actually send")
    else:
        missing = [n for n, v in (("SMTP_USER", SMTP_USER), ("SMTP_PASS", SMTP_PASS)) if not v]
        if missing:
            raise SystemExit(f"OUTREACH_LIVE=1 but {', '.join(missing)} not set")
        if not POSTAL_ADDRESS:
            log("[!] POSTAL_ADDRESS is empty - CAN-SPAM requires a physical mailing "
                "address in commercial email")

    sent, rejected = send_batch(conn, leads, headers, ws, args.cap)
    wb.save(MASTER_XLSX)

    if rejected:
        log("\n[=] not sent:")
        for domain, why in rejected[:40]:
            log(f"      {domain:<34} {why}")
    total = conn.execute("SELECT COUNT(*) FROM send_log").fetchone()[0]
    today = sent_today(conn)
    conn.close()
    log(f"\n[=] sent this run: {sent} | sent today: {today} | all time: {total}")


def cmd_suppress(args):
    conn = init_db()
    for addr in args.emails:
        suppress(conn, addr, args.reason)
        log(f"  suppressed {addr}")
    conn.close()


def cmd_status(args):
    conn = init_db()

    def one(sql, *p):
        return conn.execute(sql, p).fetchone()[0]

    audits_ok = one("SELECT COUNT(*) FROM outreach WHERE audit_status = 'ok'")
    audits_bad = one("SELECT COUNT(*) FROM outreach WHERE audit_status LIKE 'failed%'")
    total_sent = one("SELECT COUNT(*) FROM send_log")
    suppressed = one("SELECT COUNT(*) FROM suppression")
    log(f"  leads with audits : {audits_ok}")
    log(f"  audits failed     : {audits_bad}")
    log(f"  emails sent total : {total_sent}")
    log(f"  emails sent today : {sent_today(conn)}")
    log(f"  suppressed        : {suppressed}")
    conn.close()


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Growth Supply House audit + outreach")
    sub = ap.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("audit", help="run captured sites through the audit tool")
    a.add_argument("--limit", type=int, default=20)
    a.set_defaults(func=cmd_audit)

    s = sub.add_parser("send", help="email owners their audit")
    s.add_argument("--cap", type=int, default=DAILY_SEND_CAP)
    s.set_defaults(func=cmd_send)

    u = sub.add_parser("suppress", help="add addresses to the do-not-contact list")
    u.add_argument("emails", nargs="+")
    u.add_argument("--reason", default="manual")
    u.set_defaults(func=cmd_suppress)

    st = sub.add_parser("status", help="show counts")
    st.set_defaults(func=cmd_status)

    args = ap.parse_args()
    args.func(args)
