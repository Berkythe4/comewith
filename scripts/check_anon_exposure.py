"""Ask prod, as the public, what it will hand over.

Run this after ANY migration that touches a policy, a grant, or a view:

    python scripts/check_anon_exposure.py

WHY THIS EXISTS. On 2026-08-20 the ledger was publicly readable - 29 expense
rows with payee names and amounts, 59 ticketing rows, 16 donations - and every
check in this repo said it was fine:

  * post_apply.sql checks GRANTS. The financial VIEWS are anon-revoked and pass.
    The underlying TABLES carry an anon grant from 013 default privileges and
    rely on RLS, so a grant check has nothing to say about them.
  * The REST spot-checks were run with an EMPTY apikey, because .env has no
    SUPABASE_ANON_KEY - the variable is SUPABASE_PROD_PUBLISHABLE_KEY. An empty
    key answers 401 for everything, public or not, so everything looked blocked.

Both failure modes have the same shape: checking a PROXY for the invariant
instead of the invariant. The invariant is "what comes back in the body", so
that is what this reads. A 200 with rows in it is a leak whatever the grants say.
"""
import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path

# Tables the public site legitimately reads, and the views that ARE the public
# feed. Everything else in public must come back empty.
PUBLIC_OK = {
    "v_public_events", "v_public_recap", "v_public_artists", "v_public_event_photos",
    "v_public_impact_report", "v_public_survey", "v_artist_gigs", "v_artist_content",
    "site_content", "module_registry",
    # 207 - the link-in-bio pages. Both are meant to answer 200; what keeps a
    # DRAFT page out of them is the is_published filter inside the view, not the
    # grant, so seeing 200 here is the correct result and not a leak.
    "v_public_link_pages", "v_public_link_items",
}

# THESE TWO LISTS ARE THE WHOLE SWEEP. Nothing is discovered from the schema -
# an object that is named in neither list is never requested at all, and the
# "Nothing is exposed" line at the end says nothing whatsoever about it. This
# comment used to claim the opposite, which is the same trap as the financial
# views in LEARNINGS §51: a confident report over an object it never touched.
# ADD EVERY NEW TABLE AND VIEW TO ONE OF THESE LISTS IN THE SAME MIGRATION THAT
# CREATES IT.
#
# The ones worth naming explicitly, so a failure reads as a sentence rather than
# a table name.
MUST_BE_EMPTY = [
    ("expenses", "the expense ledger - payees, amounts, dates"),
    ("income", "the income ledger"),
    ("ticketing", "who bought tickets"),
    ("sponsorships", "sponsor money"),
    ("third_party_donations", "donor names and amounts"),
    ("budget_lines", "forecasts"),
    ("events", "unpublished events"),
    ("actors", "the contact graph"),
    ("guests", "attendee names and emails"),
    ("subscribers", "the mailing list"),
    ("profiles", "staff accounts"),
    ("data_health_runs", "the audit log"),
    ("data_health_waivers", "the audit log"),
    ("capital_contributions", "what Keith has put in"),
    ("event_photos", "unpublished photos"),
    ("conversations", "email threads"),
    ("conversation_messages", "email bodies"),
    ("audit_log", "everything anybody has changed"),
    # Revoked in 186 - internal, and nothing public reads them.
    ("v_equipment_roi", "equipment purchase prices and revenue per item"),
    ("v_mailing_list_health", "how big the mailing list is"),
    ("v_metric_prior", "the internal KPI scoreboard"),
    # Revoked in 187. 186 left this one granted believing tools/visualizer.html
    # read it anonymously; it does not - it loads /staging/guard.js and reads
    # with an admin session, and its other two sources answer [] / 401 to anon
    # anyway, so the tool never worked signed-out.
    ("v_kpi_targets_current", "every KPI target we have set"),
    # Invoicing (188/189). invoice_settings is the sharpest of these: it holds
    # the Bluevine routing and account numbers, which is exactly why they are
    # NOT in site_content. The client-facing invoice page reads through the
    # invoice-doc edge function on the service role, matched on public_token, so
    # none of these needs an anon grant for the feature to work.
    ("invoices", "who was billed what"),
    ("invoice_lines", "what each client was charged for"),
    ("invoice_payments", "what has been paid and how"),
    ("invoice_settings", "the PayPal handle and the Bluevine account number"),
    ("invoice_counters", "how many invoices have been raised"),
    ("v_invoices_list", "the whole receivables book"),
    ("v_invoice_totals", "invoice totals and balances"),
    ("v_invoice_line_calc", "invoice line detail"),
    ("v_income_invoiced", "which income is billed"),
    ("invoice_events", "who was chased, when, and what they paid"),
    # --- planning (197-199) ---
    ("plan_versions", "what we forecast and when we said it"),
    ("plan_offerings", "the unit economics of everything we sell"),
    ("plan_offering_lines", "prices and costs per unit"),
    ("plan_volumes", "how many events we intend to run"),
    ("plan_overrides", "hand-set forecast figures"),
    ("v_plan_offering_unit", "margin per party, per gig, per rental"),
    ("v_plan_monthly", "the whole forward forecast"),
    ("v_plan_vs_actual", "forecast against actuals"),
    ("v_event_contribution", "what every event actually contributed"),
    # --- link-in-bio pages (207) ---
    # The tables hold DRAFT pages - a slug and a set of links Keith has not
    # published yet. v_link_click_stats is how each link is performing.
    # --- venue normalisation (208) ---
    ("venue_aliases", "which venue spellings map to which room"),
    ("v_venue_name_review", "the venue names still awaiting a ruling"),
    ("v_venue_link_health", "how much of the event history is linked"),
    # --- link-in-bio pages (207) ---
    ("link_pages", "unpublished link-in-bio pages"),
    ("link_items", "links on unpublished pages, including scheduled ones"),
    ("v_link_click_stats", "how many people click each link"),
]


def env(name):
    p = Path(__file__).resolve().parent.parent / ".env"
    if not p.exists():
        return None
    for raw in p.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if line.startswith(name + "="):
            return line.split("=", 1)[1].strip().strip("'\"")
    return None


def get(url, key, path):
    req = urllib.request.Request(
        url + "/rest/v1/" + path,
        headers={"apikey": key, "Authorization": "Bearer " + key},
    )
    try:
        with urllib.request.urlopen(req, timeout=25) as r:
            return r.status, json.loads(r.read().decode("utf-8") or "null")
    except urllib.error.HTTPError as e:
        return e.code, None
    except Exception as e:  # noqa: BLE001
        return 0, str(e)


def main():
    url = env("SUPABASE_PROD_URL") or env("SUPABASE_URL")
    key = env("SUPABASE_PROD_PUBLISHABLE_KEY") or env("SUPABASE_PUBLISHABLE_KEY")
    if not url or not key:
        print("FAIL  no prod URL / publishable key in .env")
        return 1

    # Prove the key WORKS before trusting a single 401. This is the check whose
    # absence made every previous sweep meaningless.
    status, body = get(url, key, "v_public_events?select=name&limit=1")
    if status != 200 or not isinstance(body, list):
        print("FAIL  the publishable key does not read a known-public view "
              "(status %s) - every 'blocked' below would be meaningless" % status)
        return 1
    print("PASS  the key is live (v_public_events answers 200), so a 401 means something")

    fails = 0
    for table, what in MUST_BE_EMPTY:
        status, body = get(url, key, table + "?select=*&limit=3")
        if status == 401 or status == 403:
            print("PASS  %-24s blocked (%s)" % (table, status))
        elif isinstance(body, list) and not body:
            print("PASS  %-24s readable but RLS returns nothing" % table)
        elif isinstance(body, list):
            print("FAIL  %-24s LEAKS %d row(s) to the public - %s" % (table, len(body), what))
            fails += 1
        else:
            print("PASS  %-24s no rows (%s)" % (table, status))

    for v in sorted(PUBLIC_OK):
        status, body = get(url, key, v + "?select=*&limit=1")
        if status == 200:
            print("PASS  %-24s public, as intended" % v)
        else:
            print("WARN  %-24s is meant to be public but answers %s" % (v, status))

    print("\n" + ("%d LEAK(S)" % fails if fails else "Nothing is exposed that should not be."))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
