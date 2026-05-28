"""Seed Dance Infusion #2 data into Supabase staging.

Reads events/dance-infusion-2/dance_infusion.json and inserts:
  - 1 venue (Signal NYC)
  - 1 event (Dance Infusion #2, status=completed)
  - 9 sponsors + 9 sponsorships
  - 5 artists + 5 artist_bookings
  - 5 raffle_prizes
  - 4 expenses (linked to event)

Idempotent: re-running upserts where unique constraints allow,
otherwise skips. Uses db.py's auth + project ref.

Usage:
  python seed_dance_infusion_2.py
"""

import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path

# Load .env (mirrors db.py)
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
RAW_REF = os.environ.get("SBP_REF", "").strip().rstrip("/")
RAW_REF = RAW_REF.replace("https://", "").replace("http://", "")
REF = RAW_REF.split(".supabase.co")[0] if ".supabase.co" in RAW_REF else RAW_REF

if not PAT or not REF:
    print("ERROR: SBP_PAT and SBP_REF must be set in .env or env.", file=sys.stderr)
    sys.exit(1)

URL = f"https://api.supabase.com/v1/projects/{REF}/database/query"
HEADERS = {
    "Authorization": f"Bearer {PAT}",
    "Content-Type": "application/json",
    "User-Agent": "comewith-seed/1.0",
}


def sql(query: str, label: str = ""):
    body = json.dumps({"query": query}).encode("utf-8")
    req = urllib.request.Request(URL, data=body, headers=HEADERS, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else None
    except urllib.error.HTTPError as e:
        print(f"FAIL {label}: HTTP {e.code} — {e.read().decode()}", file=sys.stderr)
        sys.exit(1)


def esc(s):
    """Escape a SQL string literal."""
    if s is None:
        return "NULL"
    return "'" + str(s).replace("'", "''") + "'"


def main():
    data = json.loads((Path(__file__).parent / "events" / "dance-infusion-2" / "dance_infusion.json").read_text(encoding="utf-8"))
    cfg = data["config"]

    # --- 1. Venue ---
    print("Seeding venue: Signal NYC")
    sql(
        """insert into public.venues (name, city, state)
        values ('Signal NYC', 'New York', 'NY')
        on conflict do nothing""",
        "venue",
    )
    venue_row = sql("select id from public.venues where name = 'Signal NYC' limit 1")
    venue_id = venue_row[0]["id"]

    # --- 2. Event ---
    print("Seeding event: Dance Infusion #2")
    sql(
        f"""insert into public.events (slug, name, series, event_date, venue_id, status, bar_minimum, description, total_attendance)
        values (
          'dance-infusion-2',
          'Dance Infusion #2',
          'Dance Infusion',
          '2026-05-09',
          {esc(venue_id)},
          'completed',
          {cfg["bar_minimum"]},
          'Brunch dance party benefiting the MS Society at Signal NYC. May 9, 2026.',
          {cfg.get("standard_tickets_locked", 0) + cfg.get("brunch_tickets_locked", 0)}
        )
        on conflict (slug) do update set
          status = excluded.status,
          venue_id = excluded.venue_id,
          bar_minimum = excluded.bar_minimum,
          total_attendance = excluded.total_attendance""",
        "event",
    )
    event_row = sql("select id from public.events where slug = 'dance-infusion-2' limit 1")
    event_id = event_row[0]["id"]

    # --- 3. Sponsors + sponsorships ---
    print(f"Seeding {len(data['sponsors'])} sponsors")
    for s in data["sponsors"]:
        notes = s.get("notes", "")
        sql(
            f"""insert into public.sponsors (name, notes)
            values ({esc(s["name"])}, {esc(notes)})
            on conflict do nothing""",
            f"sponsor-{s['name']}",
        )
        sponsor_row = sql(f"select id from public.sponsors where name = {esc(s['name'])} limit 1")
        sponsor_id = sponsor_row[0]["id"]
        sql(
            f"""insert into public.sponsorships (sponsor_id, event_id, tier, cash_amount, drink_tickets, entry_tickets, status)
            values ({esc(sponsor_id)}, {esc(event_id)}, {esc(s["tier"])}, {s["cash"]}, {s["drink_tickets"]}, {s["entry_tickets"]}, 'paid')
            on conflict (sponsor_id, event_id) do update set
              cash_amount = excluded.cash_amount,
              drink_tickets = excluded.drink_tickets,
              entry_tickets = excluded.entry_tickets,
              tier = excluded.tier,
              status = excluded.status""",
            f"sponsorship-{s['name']}",
        )

    # --- 4. Artists + bookings ---
    print(f"Seeding {len(data['dj_allocations'])} artists")
    for dj in data["dj_allocations"]:
        sql(
            f"""insert into public.artists (stage_name, status)
            values ({esc(dj["name"])}, 'active')
            on conflict do nothing""",
            f"artist-{dj['name']}",
        )
        artist_row = sql(f"select id from public.artists where lower(stage_name) = lower({esc(dj['name'])}) limit 1")
        artist_id = artist_row[0]["id"]
        sql(
            f"""insert into public.artist_bookings (artist_id, event_id, role, paid)
            values ({esc(artist_id)}, {esc(event_id)}, 'DJ', false)
            on conflict (artist_id, event_id) do nothing""",
            f"booking-{dj['name']}",
        )

    # --- 5. Raffle prizes ---
    prizes = data.get("raffle_prizes", [])
    print(f"Seeding {len(prizes)} raffle prizes")
    for p in prizes:
        name = p.get("name") or p.get("prize") or p.get("description", "Prize")
        value = p.get("value") or p.get("estimated_value")
        donor = p.get("donor") or p.get("donor_name")
        value_sql = f"{value}" if isinstance(value, (int, float)) else "NULL"
        sql(
            f"""insert into public.raffle_prizes (event_id, prize_name, donor_name, estimated_value)
            select {esc(event_id)}, {esc(name)}, {esc(donor)}, {value_sql}
            where not exists (
              select 1 from public.raffle_prizes
              where event_id = {esc(event_id)} and prize_name = {esc(name)}
            )""",
            f"raffle-{name}",
        )

    # --- 6. Expenses ---
    expenses = data.get("expenses", [])
    print(f"Seeding {len(expenses)} expenses (linked to event)")
    for e in expenses:
        name = e.get("name") or e.get("description", "Expense")
        amount = e.get("amount", 0)
        vendor = e.get("vendor", "")
        date = e.get("date") or "2026-05-09"
        sql(
            f"""insert into public.expenses (event_id, date, amount, vendor, description, category)
            select {esc(event_id)}, {esc(date)}, {amount}, {esc(vendor)}, {esc(name)}, 'event'
            where not exists (
              select 1 from public.expenses
              where event_id = {esc(event_id)} and description = {esc(name)} and amount = {amount}
            )""",
            f"expense-{name}",
        )

    print("\nSummary:")
    counts = sql(
        """select
          (select count(*) from public.events where slug='dance-infusion-2') as events,
          (select count(*) from public.sponsorships where event_id = (select id from public.events where slug='dance-infusion-2')) as sponsorships,
          (select count(*) from public.artist_bookings where event_id = (select id from public.events where slug='dance-infusion-2')) as artist_bookings,
          (select count(*) from public.raffle_prizes where event_id = (select id from public.events where slug='dance-infusion-2')) as raffle_prizes,
          (select count(*) from public.expenses where event_id = (select id from public.events where slug='dance-infusion-2')) as event_expenses"""
    )
    print(json.dumps(counts, indent=2))


if __name__ == "__main__":
    main()
