"""Run all Supabase Phase 0 migrations against the staging project.

Reads credentials from environment variables (never written to disk):
  SBP_PAT  - Supabase Personal Access Token (sbp_...)
  SBP_REF  - Project ref (the subdomain piece of the dashboard URL)

Posts each SQL file to the Management API /database/query endpoint as a
single request (Postgres wraps multi-statement requests in an implicit
transaction). Stops on the first failure and prints the exact error.
After all 12 succeed, runs the four Step D sanity checks from
PHASE0_README.md and prints the results.

Uses only the Python stdlib so no pip install is required.
"""

import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path


def load_dotenv():
    """Read .env in the script directory and populate os.environ
    for any keys not already set. Mirrors db.py."""
    env_path = Path(__file__).parent / ".env"
    if not env_path.exists():
        return
    for raw in env_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


load_dotenv()


def normalize_ref(raw):
    if not raw:
        return raw
    s = raw.strip().rstrip("/").replace("https://", "").replace("http://", "")
    if ".supabase.co" in s:
        s = s.split(".supabase.co")[0]
    elif "supabase.com/dashboard/project/" in s:
        s = s.split("supabase.com/dashboard/project/")[1].split("/")[0]
    return s


PAT = os.environ.get("SBP_PAT")
REF = normalize_ref(os.environ.get("SBP_REF"))

if not PAT or not REF:
    print(
        "ERROR: Set SBP_PAT and SBP_REF environment variables before running.\n"
        "  set SBP_PAT=sbp_xxxxxxxx\n"
        "  set SBP_REF=your-project-ref",
        file=sys.stderr,
    )
    sys.exit(1)

URL = f"https://api.supabase.com/v1/projects/{REF}/database/query"
HEADERS = {
    "Authorization": f"Bearer {PAT}",
    "Content-Type": "application/json",
    # Cloudflare in front of the Management API rejects urllib's default UA
    "User-Agent": "comewith-run-migrations/1.0 (+https://comewith.org)",
}

MIGRATIONS_DIR = Path(__file__).parent / "supabase" / "migrations"
MIGRATIONS = [
    "001_extensions.sql",
    "002_profiles.sql",
    "003_clients_contractors.sql",
    "004_inquiries_agreements.sql",
    "005_financials.sql",
    "006_equipment.sql",
    "007_events.sql",
    "008_artists.sql",
    "009_mailing_list.sql",
    "010_automation_audit_photos.sql",
    "011_views.sql",
    "012_storage.sql",
    "013_grants.sql",
    "014_cron.sql",
]


def run_sql(sql: str, label: str):
    """POST one SQL string to the Management API. Exit on error."""
    body = json.dumps({"query": sql}).encode("utf-8")
    req = urllib.request.Request(URL, data=body, headers=HEADERS, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=180) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else None
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace")
        print(f"    FAIL on {label}", flush=True)
        print(f"    HTTP {e.code} {e.reason}", flush=True)
        print(f"    Body: {err_body}", flush=True)
        sys.exit(1)
    except urllib.error.URLError as e:
        print(f"    FAIL on {label}: network error: {e.reason}", flush=True)
        sys.exit(1)


def main():
    print(f"Project ref:    {REF}", flush=True)
    print(f"Migrations dir: {MIGRATIONS_DIR}", flush=True)
    print(flush=True)

    for fname in MIGRATIONS:
        path = MIGRATIONS_DIR / fname
        if not path.exists():
            print(f"ERROR: missing migration file {path}", file=sys.stderr)
            sys.exit(1)
        sql = path.read_text(encoding="utf-8")
        size_kb = len(sql.encode("utf-8")) / 1024
        print(f"==> {fname} ({size_kb:.1f} KB)", flush=True)
        run_sql(sql, fname)
        print("    OK", flush=True)

    print(flush=True)
    print("ALL 12 MIGRATIONS APPLIED SUCCESSFULLY", flush=True)
    print(flush=True)
    print("=" * 70, flush=True)
    print("STEP D SANITY CHECKS", flush=True)
    print("=" * 70, flush=True)

    checks = [
        (
            "1. Tables in public schema (expect 20+)",
            "select table_name from information_schema.tables "
            "where table_schema = 'public' order by table_name;",
        ),
        (
            "2. Total count of public tables",
            "select count(*) as table_count from information_schema.tables "
            "where table_schema = 'public';",
        ),
        (
            "3. Tables with RLS disabled (expect ZERO rows)",
            "select tablename, rowsecurity from pg_tables "
            "where schemaname = 'public' and rowsecurity = false;",
        ),
        (
            "4. Storage buckets (expect 6: agreements, artist-photos, "
            "equipment-photos, event-photos, receipts, sponsor-logos)",
            "select id, public from storage.buckets order by id;",
        ),
    ]

    for name, sql in checks:
        print(f"\n--- {name} ---", flush=True)
        result = run_sql(sql, name)
        print(json.dumps(result, indent=2, default=str), flush=True)

    print("\nDone.", flush=True)


if __name__ == "__main__":
    main()
