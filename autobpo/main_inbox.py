"""
Frequent job (every 1-2 hours): check replies, invoice anyone who said yes.
Run via .github/workflows/autobpo-inbox.yml.
"""
from db import init_db
from inbox_poll import poll_inbox
from invoice import run_invoicing


def main():
    init_db()
    poll_inbox()
    run_invoicing()


if __name__ == "__main__":
    main()
