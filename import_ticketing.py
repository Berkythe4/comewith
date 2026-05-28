"""Import ticketing CSVs into Supabase staging.

Currently supports:
  - resident_advisor: the RA export format
    (Authorisation, Barcode, Billing name, Date purchased, Email,
     Marketing Opt In, Order number, Quantity, Price, Ticket type)

To add Zeffy / another source, register another adapter in `ADAPTERS`
that returns the same canonical shape from a CSV row.

Idempotent: existing ticketing rows are matched by source + external_id
and skipped on re-run.

Usage:
  python import_ticketing.py <event-slug> <source> <path-to-csv>

Example:
  python import_ticketing.py dance-infusion-2 resident_advisor "events/dance-infusion-2/20260509-DanceInfusion#2-list (4).csv"
"""

import csv
import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path

ENV_PATH = Path(__file__).parent / ".env"
if ENV_PATH.exists():
    for raw in ENV_PATH.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        k, v = k.strip(), v.strip().strip('"').strip("'")
        if k and k not in os.environ:
            os.environ[k] = v

PAT = os.environ.get("SBP_PAT")
RAW_REF = os.environ.get("SBP_REF", "").strip().rstrip("/").replace("https://", "").replace("http://", "")
REF = RAW_REF.split(".supabase.co")[0] if ".supabase.co" in RAW_REF else RAW_REF

if not PAT or not REF:
    print("ERROR: SBP_PAT and SBP_REF must be set", file=sys.stderr)
    sys.exit(1)

URL = f"https://api.supabase.com/v1/projects/{REF}/database/query"
HEADERS = {
    "Authorization": f"Bearer {PAT}",
    "Content-Type": "application/json",
    "User-Agent": "comewith-import-ticketing/1.0",
}


def sql(query: str):
    body = json.dumps({"query": query}).encode("utf-8")
    req = urllib.request.Request(URL, data=body, headers=HEADERS, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else None
    except urllib.error.HTTPError as e:
        print(f"HTTP {e.code}: {e.read().decode()}", file=sys.stderr)
        sys.exit(1)


def esc(s):
    if s is None or s == "":
        return "NULL"
    return "'" + str(s).replace("'", "''") + "'"


# ----- Adapters: each takes a CSV row dict, returns canonical shape -----
# Canonical: { external_id, ticket_type, amount_paid, full_name, email, purchased_at }

def ra_adapter(row):
    purchased = (row.get("Date purchased") or "").strip()
    # RA gives "2026-04-21 09:41 " — Postgres timestamptz accepts that
    return {
        "external_id": (row.get("Order number") or "").strip(),
        "ticket_type": (row.get("Ticket type") or "General admission").strip(),
        "amount_paid": float(row.get("Price") or 0),
        "full_name": (row.get("Billing name") or "").strip() or None,
        "email": (row.get("Email") or "").strip().lower() or None,
        "purchased_at": purchased or None,
    }


ADAPTERS = {
    "resident_advisor": ra_adapter,
}


def main():
    if len(sys.argv) != 4:
        print(__doc__, file=sys.stderr)
        sys.exit(1)

    slug, source, path = sys.argv[1], sys.argv[2], sys.argv[3]
    adapter = ADAPTERS.get(source)
    if not adapter:
        print(f"Unknown source '{source}'. Known: {list(ADAPTERS.keys())}", file=sys.stderr)
        sys.exit(1)

    event_row = sql(f"select id from public.events where slug = {esc(slug)} limit 1")
    if not event_row:
        print(f"Event '{slug}' not found", file=sys.stderr)
        sys.exit(1)
    event_id = event_row[0]["id"]
    print(f"Importing into event_id={event_id} (slug={slug}), source={source}")

    with open(path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = [adapter(r) for r in reader]

    print(f"Parsed {len(rows)} CSV rows")

    inserted = 0
    skipped = 0
    for r in rows:
        if not r["external_id"]:
            skipped += 1
            continue

        # Upsert guest by lowercased email (if email present)
        guest_id_clause = "NULL"
        if r["email"]:
            sql(
                f"""insert into public.guests (full_name, email)
                values ({esc(r["full_name"])}, {esc(r["email"])})
                on conflict do nothing"""
            )
            guest_row = sql(f"select id from public.guests where lower(email) = lower({esc(r['email'])}) limit 1")
            if guest_row:
                guest_id_clause = esc(guest_row[0]["id"])

        # Skip if ticketing row already exists
        existing = sql(
            f"""select id from public.ticketing
            where event_id = {esc(event_id)} and source = {esc(source)} and external_id = {esc(r['external_id'])}
            limit 1"""
        )
        if existing:
            skipped += 1
            continue

        sql(
            f"""insert into public.ticketing (event_id, guest_id, ticket_type, amount_paid, source, external_id, purchased_at)
            values ({esc(event_id)}, {guest_id_clause}, {esc(r["ticket_type"])}, {r["amount_paid"]}, {esc(source)}, {esc(r["external_id"])}, {esc(r["purchased_at"])})"""
        )
        inserted += 1

    print(f"\nDone. Inserted={inserted}, Skipped (already present or no external_id)={skipped}")

    summary = sql(
        f"""select source, count(*) as tickets, sum(amount_paid)::numeric(10,2) as revenue
        from public.ticketing where event_id = {esc(event_id)} group by source order by source"""
    )
    print("\nEvent ticketing summary:")
    print(json.dumps(summary, indent=2, default=str))


if __name__ == "__main__":
    main()
