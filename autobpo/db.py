import sqlite3
import os
from contextlib import contextmanager
from config import DB_PATH

SCHEMA = """
CREATE TABLE IF NOT EXISTS leads (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source TEXT NOT NULL,            -- 'greenhouse','lever','ashby','uk_tender_cf','uk_tender_fts'
    source_ref TEXT NOT NULL,        -- job id / notice ocid, unique per source
    company TEXT,
    title TEXT,
    location TEXT,
    url TEXT,
    posted_date TEXT,
    score INTEGER DEFAULT 0,
    domain TEXT,
    email TEXT,
    status TEXT DEFAULT 'new',       -- new -> enriched -> contacted -> replied_positive/replied_negative -> invoiced
    created_at TEXT DEFAULT (datetime('now')),
    UNIQUE(source, source_ref)
);

CREATE TABLE IF NOT EXISTS outreach_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lead_id INTEGER NOT NULL REFERENCES leads(id),
    sent_at TEXT DEFAULT (datetime('now')),
    subject TEXT,
    to_email TEXT
);

CREATE TABLE IF NOT EXISTS replies (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lead_id INTEGER NOT NULL REFERENCES leads(id),
    received_at TEXT DEFAULT (datetime('now')),
    body TEXT,
    classification TEXT   -- 'positive','negative','neutral'
);

CREATE TABLE IF NOT EXISTS invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lead_id INTEGER NOT NULL UNIQUE REFERENCES leads(id),
    paypal_invoice_id TEXT,
    tier TEXT,
    amount REAL,
    status TEXT DEFAULT 'sent',
    created_at TEXT DEFAULT (datetime('now'))
);
"""

@contextmanager
def get_conn():
    os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()

def init_db():
    with get_conn() as conn:
        conn.executescript(SCHEMA)
