# DI Data Load — prod load log (2026-06-02)

Loaded the reconciled DI data into the live actor/event model on **prod** (`yaytdosxfhcqatmhctzk`)
via the Management API. Authoritative source: `DI_RECONCILED_NUMBERS.md` (reconciled WINS over raw).
All inserted rows tagged `[LOAD 2026-06-02]` in notes/description for reversibility.

## What loaded (+ number + source)
**Migration 029** (schema, file committed): `sponsorships.sponsor_id` → nullable; `actor_roles.role`
+ `donor`. (Needed to attach sponsorships to actors + model Keith as donor.)

**G1 — DI#1 duplicate resolved** (reference assumption was inverted; resolved per reference):
- Restored canonical **"Dance Infusion #1"** (`b65d425a`): date→2025-09-08, type+series=Dance Infusion.
- Moved the real income row from "Dance Infusion MS" → DI#1, set to reconciled **$1,140** (gross $1,142.50 − $2.50 RA fees).
- Deleted junk ticketing row ($50 / qty 5, test leftover) on the shell.
- **Keith founder donation $1,800** (third_party_donations, donor "Keith Berkman") — covers DI#1 costs, counts in raised.
- **Expense $1,800** (founder-paid, Production).
- Soft-deleted the **"Dance Infusion MS"** duplicate (`49d7dd65`).
- Result (v_kpi_dance_infusion): **net_pl $1,140 · total_raised $2,940 · 39% to mission** ✓ (reference: 39%).

**G2 — DI#2 financials** (`ff2b1917`):
- Expenses by group: **Venue $5,492 · Production $1,000 · Marketing $65.33** (Talent $0). = $6,557.33.
- Donations: **Crossroads $130** + **Keith founder $162.44** (third_party_donations).
- Consolidated **ticket+bar income $3,039.89** (per-line detail stays in `dance_infusion.json`).
- Result (with G5 sponsor cash): **net_pl $3,000 · total_raised $9,557.33 · 31% to mission** ✓.

**G3 — actors + DI#2 participants:** reused 4 contractor-actors (Sauci Soni, Kloud9, Kristen London, KRNeY; added dj/artist roles); created **Keith Berkman (Berky)** (roles dj+sponsor+donor+team — Berky = Keith). 5 `event_participants` (dj) on DI#2.

**G4 — Crossroads Café:** ONE actor, roles **vendor + sponsor** (the overlap case); sponsorship→DI#2 (catering + $130). Distinct from the showcase event.

**G5 — sponsors/raffle:** actors + sponsorships→DI#2 — AOM $2,500, Adam Cohen $1,500, Pulse $500, Theresa $500, Patrick $500, Francis $325, Keith $400, BeWell $0; raffle donors Bella Laser / Freemind / Brooklyn Cyclones (cash 0). Sponsor cash total **$6,225** ✓.

**G6 — Crossroads Artist Showcase** (`c32a537f`): annotated as the as-01-2026-01 content showcase (111 RA views, 4 RSVPs, RSVP≠attendance, free, graded on content). Kept distinct from Crossroads-the-vendor.

## Verification (G7) — all pass
- Money: DI#1 **39%**, DI#2 **31%** to mission (= 1 − cost_to_raise; net_pl = donated).
- Role overlap: Crossroads Café {sponsor,vendor}; Keith {dj,donor,sponsor,team}; Sauci {artist,contractor,dj,host}.
- DI#2 participants = 5; sponsor_cash = $6,225.
- No duplicate actors (17 active = 5 original + 12 new).
- anon-401 on all 5 financial views holds.

## Conflicts / assumptions logged
- **DI#1 inversion:** "Dance Infusion #1" was the soft-deleted shell; "Dance Infusion MS" was the live data. Resolved per reference (canonical = "Dance Infusion #1").
- **Founder $ = donation** attributed to Keith-the-actor (donor), counted in total_raised (per Keith).
- **% to mission not a native field** — represented as `1 − cost_to_raise` (works because net_pl = donated). Headline figures also in `events.notes`.
- **third_party_donations has no actor FK** — Keith/Crossroads donations attributed by `donor_name` (text), not linked to the actor row. Future enhancement.
- **DI#1 ticket income consolidated to $1,140** (reconciled), DI#2 ticket+bar consolidated to one income row — per-line detail remains in the JSON.

## Needs Keith's eyes
- **"19th & 7th Productions"** (existing contractor actor) may be Keith's org — NOT merged into Keith Berkman (Berky). Decide in actor-inspector.
- Confirm the DJ↔contractor matches (Sauci/Kloud9/Kristen/KRNeY) and Keith=Berky in the inspector.
- Yankees-hats raffle donor still unidentified (not loaded).

## Rollback
- Cleanly reversible (tagged rows): `delete from event_participants/sponsorships where notes like '%[LOAD 2026-06-02]%';`
  `delete from actors where notes like '%[LOAD 2026-06-02]%';` (cascades roles/links); 
  `delete from expenses where description like '%[LOAD 2026-06-02]%';`
  `delete from third_party_donations where notes like '%[LOAD 2026-06-02]%';`
  `delete from income where description like '%[LOAD 2026-06-02]%';` (DI#2 consolidated only).
- DI#1 dup-merge (messier — modified existing rows): move the income row back to `49d7dd65` (amount 1142.50),
  restore "Dance Infusion MS" (deleted_at=null), re-soft-delete "Dance Infusion #1" + restore its series/date.
  The deleted junk ticketing row is not restored (it was test junk).
- Migration 029 DOWN: see the file header.
