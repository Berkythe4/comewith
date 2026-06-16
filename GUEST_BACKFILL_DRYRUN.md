# Guest Module — Ledger Import Dry Run (Phase 1, persists nothing)

**Date:** 2026-06-16 · **Source:** `events/dance-infusion/di-02-2026-05/data/dance_infusion_ledger.csv` (113 rows) · **Prod:** `yaytdosxfhcqatmhctzk`

## Guest import (people-only — the $19,114.33 of money is NEVER written)
| Metric | Count |
|---|---|
| Ledger rows | 113 |
| Distinct valid emails | 77 |
| Overlap w/ existing 45 guests (**skip**) | 25 |
| **NET-NEW guests** | **52** |
| Net-new to subscribe (blank/True) | **52** |
| Net-new opt-outs (False) | 0 |
| **Resulting guest list** | **45 → 97** |

Each net-new guest: a `guests` row (dedupe by email) + a DI#2 `guest_event_attendance` link carrying `amount_spent` **(guest stat only — NOT a `ticketing`/`income` write)**.

## ⚠ Consent correction (the 11 `ra_marketing_opt_in=False`)
All 11 explicit opt-outs are **already existing guests/subscribers** — they were subscribed by the earlier RA-export backfill (which had them blank). Honoring consent, this sprint **unsubscribes** those 11 (`subscribers.status='unsubscribed'`, `guests.opted_in_mailing=false`) — they keep full stat tracking as guests. Net subscriber math: 45 + 52 − 11 = **86 subscribed** (97 subscriber rows; 11 unsubscribed). Reversible; backed up.
Affected: Brian Levin, Chad Hernandez, Claudia, Cory Thompson, Dana Miele, Emily Brennan, Grace Aniolek, Kyle Hilland, Michael Stevens, Sammy Smith, Zachary Gorfinkel.

## Guest↔actor graduation (donor/sponsor/vendor/dj/staff → actor, deduped)
| Bucket | Count | Action |
|---|---|---|
| **LINK** existing actor | 13 | add ledger role(s) to the existing actor; link the guest (`guests.actor_id`). No new actor. |
| **CREATE** new actor | 15 | clean relationship-people (full name + email, no name collision, not a payout artifact). |
| **FLAG** (no auto action) | 15 | possible duplicates + artifacts — Keith decides; goes to `GUEST_ACTOR_AUDIT.md`. |

**LINK (variants correctly resolved, no duplicate created):** Liz McQuillan (`Elizabeth Mcquillan`), Keith Berkman (Berky) (`Keith Berkman`), DJ Sauci Soni (`Sauci Soni`), Theresa Berkman, Adam Cohen, Patrick Savery, Francis Berkman, KRNeY, Kloud9, Kristen London, AOM Infusion, Pulse Devices, BeWell.

**CREATE (15):** Kendall Leary, Angela Tabone, Tarik Hajjam, Amelia Allen, Tanya Raine-Roosevelt, Ani Kanburiyan, Zachary Storey, Mariia Vysotska, Andriy Kashyrskyy, Victoriarose Vargas, Jennifer Alderman, Ethan Pollak, Aysha Khawaja, Laura Kennis, Jennifer Taveras (roles: donor/sponsor/staff per ledger).

**FLAG (15 — NOT created/merged):** `Crossroads Cafe`/`Crossroads Cafe (Pastries)` (≈ Crossroads Café), `Signal NYC (175 Morgan Ave LLC)` (the venue), `Patrick Savery Sr.` (≈ Patrick Savery — likely different person), `Theresa McGuinness` (≈ Theresa Berkman by surname only — likely different), `Lisa Meyer-Savery`, `Martin P. Salas`, `Oheyitsamanda`, `Anonymous Donor (Crossroads)`, `Amanda Brunidge` (no email; ≈ guest Amanda Brundige), `Tulay Sencar` (no email), `Angela Tabone` (no-email donor line), `Postnet NY143 (Fliers)`, `Resident Advisor (RA) Payout`, `Zeffy Pending Payout` — payout/financial artifacts and accent/nickname collisions. **Flag-don't-guess.**

## Plan for persisting phases
- Schema 038 (additive): `guests.actor_id` (graduation link) + `v_event_attendance_kpi` + `v_guest_kpis`.
- Backup `guests/subscribers/guest_event_attendance` first. Idempotent, tagged `[GUEST LEDGER 2026-06-16]`.
- **Money: none written.** DI#2 financials untouched (verified before/after). Reconciliation (Phase 6) reads only.
