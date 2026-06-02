# Carryover — 2026-06-02 (session close)

Session closed cleanly. For pickup: read this, then `LEARNINGS.md` (the why), `ROADMAP.md`
(parked + backlog), `CLAUDE.md` (conventions). Close ritual: `SESSION_CLOSE_PROMPTS.md`.
DI-content running log lives separately at
`events/dance-infusion/di-02-2026-05/DanceInfusion_Updates_Log.md`.

## State summary

- **Prod:** Supabase `yaytdosxfhcqatmhctzk`; live at **comewith.org** (Netlify, auto-deploy from `master`).
- **Migrations applied:** 015–022 (**highest = 022**); **unchanged this session — files only, no DB changes.**
- **Roles:** master_admin (Berky), sub_admin (Liz), customer — via `public.is_admin()`.
- **Financial views:** anon-**revoked**, verified **401** this close — v_event_summary, v_kpi_event_financials, v_kpi_parties, v_kpi_dance_infusion, v_kpi_dashboard.
- **Latest LEARNINGS §:** 10
- **Git:** DI impact-report work merged + pushed to `origin/master` (through `11e29c8`). **This close-out's commits (folder rename + ritual docs) are LOCAL, unpushed** by request.
- **New this session:** `/staging/` admin-gated review area (LEARNINGS §10); DI #2 impact report + public audit live behind the gate.

## Tomorrow's default

**CWF (Come With Fitness) BRD** remains the calendar priority — soft **June 8**, hard **June 15**.
Come With stays **maintenance-only**; HARD RULE: nothing Come With Fitness in the Come With
dashboard / schema / pages until the BRD ships and there's an explicit go (LEARNINGS §5).

**For the DI impact-report thread specifically:** the gate before it can post publicly is the
**CONSENT SWEEP** (see Parked / next).

## This session shipped (2026-06-02) — files only, no migrations

- DI #2 **impact report** (`…/di-02-2026-05/reports/impact-report.html`) + **public audit**
  (`public-audit.html`), live from `dance_infusion.json` + `dance_infusion_di1.json`; PDF export.
- Money framing locked: **"% to the mission"** public / expense-ratio internal (LEARNINGS §8);
  **$9,557** total-raised reconciliation incl. ~$162 founder contribution (LEARNINGS §9).
- Reusable **`/staging/` admin gate** reusing the dashboard's Supabase auth — `guard.js` + 2-line include + hub (LEARNINGS §10).
- DI #1 data **confirmed** ($2,940 raised / $1,140 donated / 39% to mission); folder **renamed** `di-01-2024-09 → di-01-2025-09` (+ fetch path + series_summary fixed).
- Removed duplicate "Partners" block from the report.
- All merged to `origin/master` behind the gate; **homepage untouched / still public.**

## Parked / next

- **CONSENT SWEEP** (gates the impact report going public): sponsors, team, artists, raffle donors — and **identify the Yankees-hats donor**.
- **Placeholders to fill:** human-moment quote, hero/inline photos, social reach numbers, what's-next copy.
- **To publish the report:** remove the 2 guard lines from `impact-report.html` + `public-audit.html`.
- **Dashboard:** set internal `di.cost_to_raise` target to 40% expense / 60% to mission so dashboard + audit reconcile (LEARNINGS §8).
- Still parked (unchanged): **event-model redesign** (design-first; the big one), DI dashboard `event_date` fix (2026-09-08 → 2025-09-06), flywheel redesign, roadmap/timeline tool, smaller backlog — all in `ROADMAP.md`.

## How to verify (quick)

- Anon REST GET each financial view → expect **401** (`SESSION_CLOSE_PROMPTS.md` step 1).
- `/staging/` (logged out) → redirects to `/dashboard.html`; signed-in master_admin sees the report.
- `git log --oneline` + highest number in `supabase/migrations/` = what's on prod.
