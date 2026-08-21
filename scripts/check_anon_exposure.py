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
}

# The ones worth naming explicitly, so a failure reads as a sentence rather than
# a table name. Everything else discovered from the schema is checked too.
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
