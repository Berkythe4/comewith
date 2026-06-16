# Historical Backfill — Inventory + Dry Run (Phase 3, persists nothing)

**Date:** 2026-06-16 · **Source:** `events/` (per-event folders, data-load logs, reconciled numbers) · **Prod:** `yaytdosxfhcqatmhctzk`

## Inventory of `events/`
| Folder | Event (DB) | Source data | Form |
|---|---|---|---|
| `dance-infusion/di-01-2025-09` | Dance Infusion #1 (`b65d425a`) | RA ticket export (42 rows), 3 RA snips, `dance_infusion_di1.json`, reconciled numbers | csv + jpg + json |
| `dance-infusion/di-02-2026-05` | Dance Infusion #2 (`ff2b1917`) | full finance (ledger, audit tables, door list, transactions), updates log, impact brief | xlsx + csv + md |
| `artist-showcase/as-01-2026-01` | Crossroads Café Artist Showcase (`c32a537f`) | RSVP list (4 rows: Keith, Liz) | csv |
| (updates log) | DI Artist Showcase — Kristen London & 32LVS (`1ebbec24`) | DanceInfusion_Updates_Log (Apr 25/27 posts) | md |

**Key source facts:** the `*-list.csv` files are **RA ticket/RSVP attendee exports** (Billing name, Email, Price, Ticket type) — *customers, not the people-you-work-with layer*. They are **not** actor sources. `DI_DATA_LOAD_LOG.md` (2026-06-02) already loaded the actor/money model; this backfill fills the remaining **people-links**.

## Financial state per event (drives the money action — checked against the DB)
| Event | DB: tickets/income/exp/spon/don | Spreadsheet has | **Money action** |
|---|---|---|---|
| Dance Infusion #1 | 0 / 1 / 1 / 0 / 1 | RA tickets ($1,142.50 = the reconciled $1,140 income; founder $1,800 = the expense+donation) | **UNTOUCHED** — reconciled-complete; ticket data is the *same money* as the income row (double-count risk) |
| Dance Infusion #2 | 0 / 2 / 3 / 12 / 2 | full finance | **UNTOUCHED** — complete in DB |
| Crossroads Café Artist Showcase | 0 / 1 / 0 / 0 / 0 | RSVP-only, free (no P&L) | **SKIP+FLAG** — has a placeholder income row; free content event |
| DI Artist Showcase (KL & 32LVS) | 0 / 1 / 0 / 0 / 0 | none (content) | **SKIP+FLAG** — placeholder income row present |
| Maxwell House 4/20 | 0 / 1 / 0 / 0 / 0 | none in `events/` | **SKIP+FLAG** — placeholder income row present |
| Knicks G5 / Come With 7-11 | has ticket/none | none in `events/` | **UNTOUCHED / n/a** |

> **Net money conclusion:** there is **NO event that is empty-in-DB AND has spreadsheet money** → **no POPULATE case applies.** DI#1 (the prompt's POPULATE candidate) already holds its reconciled money — populating tickets would **double-count**. Phase 4 writes **NO money** (ticketing/income/expenses/sponsorships/donations). Backfill = **actors + people-links only.**

## Actors to CREATE (new — none match the 20 existing actors)
| Name | kind | role | Source |
|---|---|---|---|
| 32LVS | person | artist | Updates log ("Duo DJ, artist showcase series"); event name |
| Gavin (Signal) | person | venue_contact | Updates log: "Technical rider submitted to Signal (Gavin); Signal records all 5 sets" → **sound** |
| Sara (Signal) | person | venue_contact | Updates log: "Sara (Signal) IG collab / Meet-the-event series" → **other (marketing/collab)** |

## Actors to LINK (existing — deduped, not recreated)
| Actor (exists) | Link |
|---|---|
| Keith Berkman (Berky) | participant on **DI#1** (dj) — DI#1 was a solo run "just Keith" (reconciled doc) |
| Kristen London | participant on **DI Artist Showcase** (artist) |

## People-links to CREATE
- **event_participants:** DI#1 → Keith (dj); DI Artist Showcase (`1ebbec24`) → Kristen London (artist) + 32LVS (artist).
- **venue_contacts (Signal `e417b79d`):** Gavin → `sound` (primary); Sara → `other` (note: IG/marketing collab).
- (Signal already has `actor_id` link? No — created on demand when contracting; left as-is.)

## ⚠ FLAGGED — uncertain, NOT executed (need Keith)
- **Crossroads Café Artist Showcase performer roster** — undocumented; the only list is RSVP attendees. No participant links created.
- **Rich Klein** — "1:1 done" but affiliation unclear (Signal? MS Society board?). Not linked.
- **Sara's exact function** — added as `other` (clearly a Signal contact); refine to booking/marketing if Keith knows.
- **DI#1 42 ticket-buyers** — attendees, deliberately NOT backfilled as actors.
- **Yankees-hats raffle donor** — still unidentified (per data-load log); not loaded.
- **Maxwell House 4/20 / Come With 7-11 / Knicks G5** — no documented roster in `events/`; left as-is.

## Phase 4 plan (safe subset)
Backup financial/ticket tables → create the 3 actors (deduped by name) → link Keith/Kristen → create the 5 people-links (3 participants + 2 venue contacts) → **zero money writes**. Idempotent (name dedup; unique constraints on participant/contact). Tagged `[BACKFILL 2026-06-16]`. Then verify DI#2 money counts unchanged.
