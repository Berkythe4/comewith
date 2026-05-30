# Carryover — 2026-05-29 (session close)

Session closed cleanly. For pickup: read this, then `LEARNINGS.md` (the why behind
decisions), `ROADMAP.md` (parked + backlog), and `CLAUDE.md` (how-to-work conventions).
The close ritual itself is in `SESSION_CLOSE_PROMPTS.md`.

## State summary

- **Prod:** Supabase `yaytdosxfhcqatmhctzk`; live at **comewith.org** (Netlify, auto-deploy from `master`).
- **Migrations applied:** 015–022 (**highest = 022**); `supabase/migrations/` matches prod.
- **Roles:** master_admin (Berky), sub_admin (Liz), customer — via `public.is_admin()`. Email+password and magic-link auth both live.
- **Financial views:** anon-**revoked**, verified **401** — v_event_summary, v_kpi_event_financials, v_kpi_parties, v_kpi_dance_infusion, v_kpi_dashboard.
- **Latest LEARNINGS §:** 7
- **Homepage:** placeholder; admin dashboard `/dashboard.html`, default tab = **Strategy**.
- **Git:** all shipped work pushed to `origin/master`.

## Tomorrow's default

**CWF (Come With Fitness) BRD** — soft deadline **June 8**, hard **June 15**; Martin reviewing
CWF late next week. Come With is **maintenance-only** until the BRD ships. **HARD RULE:** nothing
Come With Fitness in the Come With dashboard/schema/pages until the BRD is done and there's an
explicit go decision (LEARNINGS §5).

## This session shipped (2026-05-29)

- Auth: Berky master_admin, Liz sub_admin; email+password auth (magic-link fallback).
- Real data on prod; homepage → placeholder.
- KPI layer (migrations **015–022**): tables, views, RLS, anon-revoked financial views.
- Strategy tab: KPI cards, formula hover-tooltips, flywheel (placeholder design).
- Entry forms: Log Event, Log Numbers, Edit Target (+ create-new-metric, retire-metric).
- `feedback_log` → Notes tab; event edit + soft-delete; Add Sponsor UI; per-event Money panel; Income delete.
- **Money model fixed (022):** net P&L now includes ticket revenue. Canonical revenue definition — LEARNINGS §1.
- Income↔events reconciled: 3 junk rows soft-deleted; 5 real events created + existing income linked (no double-count) — placeholder series, LEARNINGS §7.

## Parked / next

The big one is the **event-model redesign** (events are multi-axis — TYPE / CONTENT / SIGNATURE
plus relational links; design **before** migrating). DI #1/#2 backfill (and fix the "Dance Infusion
MS" date to 2025-09-06), flywheel redesign, the roadmap/timeline tool, and the smaller backlog all
live in `ROADMAP.md` (Parked + Backlog sections). **When Come With resumes: the event-model design
session comes first** — the backfill, flywheel, and series reassignment all depend on it.

## How to verify (quick)

- Anon REST GET each financial view → expect **401** (see `SESSION_CLOSE_PROMPTS.md` step 1).
- `git log --oneline` + highest number in `supabase/migrations/` = what's on prod.
