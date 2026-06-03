# Dance Infusion — Reconciled Numbers (LOAD AUDIT REFERENCE)

**Purpose:** authoritative figures for loading DI data into the actor/event model. These were reconciled and locked in working sessions (May 29 – June 1). **Where the raw repo exports (RA reports, financial XLSX, JSON) disagree with these, THESE WIN** — the raw exports predate the corrections. Use the raw exports for operational detail (per-ticket, attendance, participants, lineup) that doesn't conflict with the headline figures below.

---

## The money model (how every figure is defined)

```
total_raised = ticket_revenue + other_income(bar/vendor/merch)
             + donations + sponsor_cash + founder_contribution
net_to_mission (donated) = what reached the National MS Society
% to mission = donated / total_raised
```
- **PUBLIC metric is "% to the mission"** (positive framing). Internal can speak expense ratio.
- "Raised" includes founder out-of-pocket contributions (money is fungible; it made the raise possible).
- DI #2 had **no money left over** — total raised = expenses + donation exactly.

---

## DI #1 — Sept 6, **2025** (folder: di-01-2025-09)

```
total_raised:        $2,940   (= $1,800 expense + $1,140 donated)
donated_to_ms:       $1,140
expenses:            $1,800   (Keith paid PERSONALLY / out of pocket)
% to mission:        39%
RA tickets sold:     42
RA ticket revenue:   $1,142.50   (NOTE: this sits in the income/ticket
                                  data; it is NOT separate from the
                                  $1,140 donated — see reconciliation note)
RA event views:      990
attendance:          UNKNOWN — RA RSVP ≠ attendance; do NOT report RSVP
                     as attendance. Record "42 RA tickets" in notes,
                     leave attendance null unless a true count exists.
```
**Context:** solo-run (just Keith), 5 hrs (1–6pm), off-peak. Proof of concept. NOT measured against any efficiency target (the ≤40%-expense commitment begins DI #3).

**Reconciliation note for DI #1:** the RA ticket revenue ($1,142.50) and the "$1,140 donated" are the same money stream (ticket sales → donated to MS). The $1,800 expense was covered personally. Don't double-count ticket revenue as both income AND a separate donation.

### ⚠ DI #1 DUPLICATE — resolve during load
There are TWO DI#1-ish events in prod:
- **"Dance Infusion MS"** — carries the real $1,142.50 (the canonical DI#1)
- **"Dance Infusion #1"** — empty shell, $0
**Action:** merge into ONE canonical DI#1 event (use "Dance Infusion #1" as the name, "Dance Infusion MS" as the data source), apply the reconciled numbers above, soft-delete the duplicate. Confirm with Keith which name survives before deleting.

---

## DI #2 — May 9, **2026** (folder: di-02-2026-05)

```
total_raised:        $9,557   (= $6,557 expense + $3,000 donated)
donated_to_ms:       $3,000
expenses:            $6,557
  ├── Venue:         $5,492
  ├── Production:    $1,000
  ├── Talent:        $0
  └── Marketing:     $65   (actual $65.33; public rounds to $65)
% to mission:        31%
```
**Total raised composition:** ~$9,264.89 gross through event accounts + $130 Crossroads Café direct donation + ~$162 founder contribution covering remaining costs = $9,557. (The audit carries one line: "Total raised includes a founder contribution that covered remaining event costs.")

**Context:** 7 hrs (3–10pm), prime time, full production, sponsors, multiple DJs.

**Participants to load** (from repo data / known): DJs include Berky, KRNeY, Kloud9, Kristen London, Sauci Soni (confirm against artist_bookings / repo lineup data). Load as actors + event_participants with appropriate roles.

**Sponsors/vendors to load** (from the impact-report brief gratitude list — confirm against repo data):
- Crossroads Café = **vendor + sponsor** (catered + $130 direct donation) — the role-overlap test case
- Sponsors: AOM Infusion, Adam Cohen, Pulse Devices, Theresa Berkman, Patrick Savery, Francis Berkman, BeWell (+ Keith)
- Raffle donors: Bella Laser, Freemind, Brooklyn Cyclones, (Yankees-hats donor TBD)
- Load as actors with sponsor/vendor roles + sponsorships linked to the DI#2 event.

---

## TRAJECTORY framing (critical — affects how comparisons are stored/shown)

DI #1 → DI #2 is **GROWTH (absolute), NOT efficiency improvement**:
- Donated: $1,140 → $3,000 (2.6×)
- Total raised: $2,940 → $9,557 (~3.3×)
- % to mission went 39% → 31% — this is EXPECTED (DI#2 scaled up: prime-time, full production). Do NOT frame as efficiency improvement. The 60%-to-mission (≤40% expense) target is a FORWARD commitment from DI #3 onward.

---

## CROSSROADS ARTIST SHOWCASE — SEPARATE, not DI

The **artist-showcase** event (folder: artist-showcase/as-01-2026-01, "Club Cafe presents," Jan 17 2026) is a SEPARATE content event — NOT part of Dance Infusion. It was free, RA used for RSVP only (RSVP ≠ attendance). Load it as a `showcase`-type event (content event, graded on views/follows, not P&L). Do NOT conflate with Crossroads Café-the-vendor at DI#2 (same name, different things: one is an event, one is an actor/vendor).

---

## Load discipline

- Reconciled numbers above WIN over raw exports for headline financials.
- Raw exports supply operational detail (per-ticket, lineup, dates) where not conflicting.
- attendance: never report RA RSVP as attendance.
- Resolve the DI#1 duplicate (confirm canonical name with Keith first).
- Crossroads-the-showcase ≠ Crossroads-the-vendor.
- After load: re-run the test-checklist against the now-populated model.
