"""
Daily job: discover -> score -> enrich -> outreach.
Run via .github/workflows/autobpo-discover.yml.
"""
from db import init_db, get_conn
from sources import ats_boards, uk_procurement
from score import score_job_lead, score_tender_lead
from enrich import enrich_lead
from outreach import run_outreach


def upsert_lead(conn, lead, score):
    conn.execute("""
        INSERT INTO leads (source, source_ref, company, title, location, url, posted_date, score, status)
        VALUES (?,?,?,?,?,?,?,?, 'new')
        ON CONFLICT(source, source_ref) DO NOTHING
    """, (lead["source"], lead["source_ref"], lead["company"], lead["title"],
          lead["location"], lead["url"], lead["posted_date"], score))


def main():
    init_db()
    job_leads = ats_boards.fetch_all()
    tender_leads = uk_procurement.fetch_all()

    by_company = {}
    for l in job_leads:
        by_company.setdefault(l["company"], []).append(l)

    with get_conn() as conn:
        for company, leads in by_company.items():
            for l in leads:
                upsert_lead(conn, l, score_job_lead(l, leads))
        for l in tender_leads:
            upsert_lead(conn, l, score_tender_lead(l))

        new_leads = conn.execute("SELECT * FROM leads WHERE status='new'").fetchall()
        new_leads = [dict(r) for r in new_leads]

    for lead in new_leads:
        enriched = enrich_lead(lead)
        if enriched.get("email"):
            with get_conn() as conn:
                conn.execute(
                    "UPDATE leads SET domain=?, email=?, status='enriched' WHERE id=?",
                    (enriched["domain"], enriched["email"], lead["id"]),
                )

    run_outreach()
    print(f"discovered {len(job_leads)} job leads, {len(tender_leads)} tender leads")


if __name__ == "__main__":
    main()
