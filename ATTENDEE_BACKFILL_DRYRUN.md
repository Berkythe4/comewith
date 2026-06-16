# Attendee + Mailing Backfill — Inventory + Dry Run (Phase 1, persists nothing)

**Date:** 2026-06-16 · **Prod:** `yaytdosxfhcqatmhctzk`

## Sources found (attendee lists)
| Source file | Event | Rows | Fields | Email? | Opt-out flag? | Amount? |
|---|---|---|---|---|---|---|
| `di-01-2025-09/data/20250906-DanceInfusionMS-list.csv` | Dance Infusion #1 | 42 | RA export | ✅ | "Marketing Opt In" = `Opt-in`/blank (no explicit opt-out value) | ✅ Price |
| `artist-showcase/as-01-2026-01/…ClubCafepresents…-list.csv` | Crossroads Café Artist Showcase | 4 | RA export | ✅ | same | ✅ (free, $0) |
| `di-02-2026-05/source-files/20260509-DanceInfusion#2-list (4).csv` | Dance Infusion #2 | 20 | RA export | ✅ | same | ✅ Price |
| `di-02-2026-05/reports/DanceInfusion_DoorList.xlsx` | DI#2 door list | 81 | Name/Tickets/Type/Source | ❌ **no email** | — | drink tix only |

## Dry-run plan (persist nothing) — dedupe by email
| Metric | Count |
|---|---|
| Attendee rows (RA exports) | **66** |
| **Unique guests after dedupe (email)** | **45** |
| No-email rows (RA) | 0 |
| Malformed emails | 0 |
| Multi-event guests | **2** — Claudia (DI#1 + DI#2), Liz McQuillan (Crossroads + DI#1) |
| Explicitly opted in (any event) | 25 |
| Never explicitly opted in (cold) | 20 |
| Explicit opt-outs | **0** |
| **→ Would subscribe (locked rule: all w/ email except opt-outs)** | **45** |
| Excluded as opt-out | 0 |
| Cannot subscribe (no email) | 0 |

## ⚠ Deliverability exposure (cold = subscribed but never ticked RA opt-in)
Per the locked rule we subscribe everyone with an email (no source has an explicit opt-out value). These addresses **bought a ticket but did not tick RA's marketing box** — cold from a consent/deliverability standpoint:
| Event | Cold / unique |
|---|---|
| Dance Infusion #1 | **15 / 28** |
| Crossroads Café Artist Showcase | 1 / 3 |
| Dance Infusion #2 | **14 / 16** |
| **Total cold-only unique guests subscribed** | **20 / 45** |

> Keith decides with these numbers. If you'd rather only subscribe the 25 explicit opt-ins and hold the 20 cold ones, say so — the backfill tags source/segment so a later unsubscribe of the cold set is a one-liner.

## Flagged — NOT imported (ambiguous / unsubscribable)
- **DI#2 door list (81 names, no emails)** — door/comp attendees; names overlap the RA export with variant spellings (e.g. "Brunidge" vs "Brundige") and there's **no email to dedupe on**. Auto-importing would create duplicate guests and unsubscribable rows. Flagged for Keith; the RA email exports are the authoritative subscribe source.

## Plan for the persisting phases
- **Schema (037, additive):** `guest_event_attendance` (guest↔event link carrying `amount_spent`) + `v_guest_stats` view. **Deliberately NOT writing `ticketing` rows** — that would inflate `v_event_summary.ticket_revenue` and double-count DI#1's already-reconciled income. Guest "total spent" comes from the attendance link, keeping event financials untouched.
- **Guests:** 45, deduped by email, linked to their event(s) (47 guest-event links). `opted_in_mailing=true` for all (none opted out).
- **Subscribers:** 45 (`status=subscribed`), deduped by email, `guest_id` linked, per-event `subscriber_segments`.
- **Idempotent**, tagged `[ATTENDEE BACKFILL 2026-06-16]`. Backup of guests/subscribers/ticketing first.
