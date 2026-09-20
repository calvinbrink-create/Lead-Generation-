"""
Simple pain-signal scoring using only what's visible in the posting/tender
record itself - no paid enrichment data.
"""
from datetime import datetime, timezone


def score_job_lead(lead, all_leads_for_company):
    score = 0
    # multiple open BPO-relevant roles at once = real operational pain
    score += min(len(all_leads_for_company), 5) * 10
    # role open a while = struggling to fill it internally
    posted = lead.get("posted_date")
    if posted:
        try:
            posted_dt = datetime.fromisoformat(posted.replace("Z", "+00:00"))
            days_open = (datetime.now(timezone.utc) - posted_dt).days
            if days_open > 45:
                score += 20
            elif days_open > 21:
                score += 10
        except Exception:
            pass
    title = (lead.get("title") or "").lower()
    if any(w in title for w in ["urgent", "immediate", "asap"]):
        score += 15
    return score


def score_tender_lead(lead):
    # any published tender for this exact work is already a strong,
    # qualified, budget-confirmed signal
    return 80
