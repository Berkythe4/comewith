"""The five financial views must return anon 401. Verify it, end to end.

Why this exists separately from check_anon_exposure.py: that script discovers
objects from the schema and asserts "nothing comes back that shouldn't". These
five are named in CLAUDE.md and MERGE_ROUTINE.md as the invariant every close
must check by name - and on 2026-08-25 they were found NOT to be in the sweep's
output at all, so the close was verifying an invariant nothing actually tested.

The two ways this check has been got wrong before (LEARNINGS SS37) both apply:

  * There is no SUPABASE_ANON_KEY in .env. The variable is
    SUPABASE_PROD_PUBLISHABLE_KEY. An empty apikey answers 401 for EVERYTHING,
    public or not - so a 401 proves nothing unless the key is known to work.
    This reads a known-public view FIRST and refuses to continue otherwise.
  * 401 is the right thing to look for on a VIEW (they are anon-revoked), but
    NOT on a table. These five are views, so 401 is the assertion.

    SBP_REF is not needed - this is the public REST endpoint, not the
    Management API.

    python scripts/check_financial_views.py
"""
import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path

# Decision E1. Revoked from anon deliberately; 016/017 silently re-granted them
# and 019 had to fix it. That regression is the reason this runs every close.
FINANCIAL_VIEWS = [
    ("v_event_summary", "per-event gross, net and attendance"),
    ("v_kpi_event_financials", "the money behind every event KPI"),
    ("v_kpi_parties", "Come With Parties revenue"),
    ("v_kpi_dance_infusion", "Dance Infusion revenue"),
    ("v_kpi_dashboard", "every KPI card, targets included"),
]
CANARY = "v_public_events"   # public by design - proves the key actually works


def env(name):
    v = os.environ.get(name)
    if v:
        return v
    envfile = Path(__file__).resolve().parent.parent / ".env"
    if envfile.exists():
        for line in envfile.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, _, val = line.partition("=")
            if k.strip() == name:
                return val.strip().strip('"').strip("'")
    return None


def get(url, key, obj):
    req = urllib.request.Request(
        f"{url}/rest/v1/{obj}?select=*&limit=1",
        headers={"apikey": key, "Authorization": f"Bearer {key}"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.status, r.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace")


def main():
    url = (env("SUPABASE_PROD_URL") or env("SUPABASE_URL") or "").rstrip("/")
    key = env("SUPABASE_PROD_PUBLISHABLE_KEY") or env("SUPABASE_PUBLISHABLE_KEY")
    if not url or not key:
        print("No prod URL / publishable key in .env "
              "(SUPABASE_PROD_URL + SUPABASE_PROD_PUBLISHABLE_KEY). Cannot verify.")
        return 2

    # Prove the key works before trusting a single 401.
    status, body = get(url, key, CANARY)
    if status != 200:
        print(f"ABORT  the publishable key does not work - {CANARY} returned {status}, not 200.")
        print("       Every 401 below would be meaningless. Fix the key first.")
        return 2
    print(f"OK     key works ({CANARY} -> 200, {len(json.loads(body))} row(s))\n")

    bad = 0
    for name, what in FINANCIAL_VIEWS:
        status, body = get(url, key, name)
        if status == 401:
            print(f"PASS  {name:<26} anon 401 ({what})")
        else:
            bad += 1
            preview = body[:160].replace("\n", " ")
            print(f"FAIL  {name:<26} anon {status} - EXPOSED: {what}")
            print(f"      body: {preview}")

    if bad:
        print(f"\n{bad} financial view(s) READABLE BY ANON. This is the 016/017 "
              "regression shape - look for a blanket grant. STOP and fix before shipping.")
        return 1
    print("\nAll five financial views are anon-revoked, verified through PostgREST.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
