# Carryover — 2026-06-02 (data-load session close)

Pickup order: this → `LEARNINGS.md` → `ROADMAP.md` → `CLAUDE.md`. Ritual: `SESSION_CLOSE_PROMPTS.md`.
DI data load detail: `events/dance-infusion/DI_DATA_LOAD_LOG.md`.

## ⛔ PRIORITY CONTEXT
**Come With is MAINTENANCE-ONLY.** The **CWF (Come With Fitness) BRD is project #1 — due JUNE 15, 2026.**
Nothing Come With Fitness in this repo (dashboard / schema / pages) until the BRD ships **and** there's
an explicit go (LEARNINGS §5).

## State summary
- **Prod:** Supabase `yaytdosxfhcqatmhctzk`; live at comewith.org (Netlify auto-deploy from `master`).
- **Migrations: through 029 APPLIED to prod** — 023–028 (data architecture) + **029** (sponsorships
  actor FK + `donor` role). Applied via the Management API (not the CLI migration system); the
  migration **files** are the tracked source of truth. **origin/master = `c7ca237`**, so the
  023–028/029 files for the *latest* commits are partly **held** (see Git).
- **Data model POPULATED** with reconciled DI data — DI#1 **39%** to mission, DI#2 **31%**; 17 actors
  (role-overlap working: Keith = dj+donor+sponsor+team; Crossroads = vendor+sponsor); 5 DI#2 DJ
  participants; 12 sponsorships ($6,225 cash). No duplicate actors.
- **Roles:** master_admin / sub_admin / customer via `public.is_admin()`; new `donor` role on `actors`.
- **Financial views:** anon-revoked, verified **401**. ⚠ **NOT revoked from `authenticated`** — that's
  the GATED BLOCKER before any customer/external login (ROADMAP).
- **Latest LEARNINGS §:** 14.
- **Git:** **3 commits held** (push held per Keith): `261797d` (029 + DI load log), `5cbb51e` (roadmap
  backlog), + this close-out commit. Branches `fix-lognumbers-optgroups`, `docs/roadmap-reconcile`
  pushed but unmerged.
- **Tools:** `/tools/actor-inspector.html` · `/tools/test-checklist.html` · `/tools/visualizer.html`
  deployed on comewith.org, admin-gated via the staging guard.

## Tomorrow's default
**CWF BRD (June 15).** Come With stays maintenance-only.

## This session shipped (2026-06-16 — Event Hub, Sprint 1 of the module series)
Built the **Event Detail Hub** in `dashboard.html` (additive; the 14 existing tabs untouched). Reached
via an **Open** button on each Events row → a new `#panel-eventhub` with header (name + facts + **stage
stepper** + Edit core) and sections **Overview · People · Tasks · Money · Equipment · Contracts · Files**.
- **People** = `event_participants` (add existing/new actor, role, fee; **one-click Fee→expense**, manual — not auto-posted).
- **Tasks** = `tasks`+`task_assignments` (add, assign owner/doer/reviewer, inline status, soft-delete).
- **Money** = reuses the existing per-event Money panel.
- **Equipment** = `equipment_usage` — **the previously-unwritten event-attach path now writes** (feeds the day-of generator + ROI).
- **Contracts** = new `contracts` table (canonical; legacy `agreements` ignored), kinds incl. vendor/sponsor; status + mark-paid.
- **Files** = `files` table; documents in the private `agreements` bucket (signed-URL download).
- **Stage stepper** (idea→…→reported) updates `events.stage`, distinct from public `status`.
- **Day-of generator** = prominent button → `generate_day_of_tasks(p_event_id)`; idempotent.
- **Schema add:** migration **031_v_actor_full.sql** APPLIED to prod — `v_actor_full` (actors + roles[],
  `security_invoker=true`, anon-revoked). Establishes the **actor_*_details pattern** (`docs/ACTOR_DETAILS_PATTERN.md`)
  the Artist/Vendor sprints follow. No `actor_*_details` tables built yet (by design).
- **Locked decisions honored:** internal-admin only (no RLS/external-login changes), `type` enum as the
  operational axis (`series`/KPI views untouched), full cohesive design (template for the other modules).
- **Tests:** `tests/event_hub_datalayer_test.sql` — every hub write + the day-of RPC, run against prod
  **in a rolled-back transaction** (zero rows persisted; counts verified unchanged). All green.
- **Live-drive:** `EVENT_HUB_LIVE_DRIVE.md` (UI walkthrough).
- **Deferred to follow-ups:** linking an uploaded file as a contract's `document_id` (Files + Contracts
  ship separately today); a global cross-event task board; contract signing flow.

## This session shipped (2026-06-02 — data load)
Applied 023–028 + 029 to prod; populated the model with reconciled DI#1/#2 data; resolved the DI#1
duplicate (canonical "Dance Infusion #1"); proved role-overlap on real data; anon-401 held throughout.

## Open threads — needs Keith's eyes (in actor-inspector)
- **"19th & 7th Productions"** (existing contractor actor) — merge into Keith Berkman (Berky), or keep separate?
- Confirm DJ↔contractor matches + Keith = Berky.
- **Yankees-hats raffle donor** — unidentified, not loaded.
- **Held commits** (`261797d`, `5cbb51e`, close-out) — push pending Keith's go.

## Gated blocker
**Financial-view security fix** (revoke from `authenticated`, re-issue `security_invoker`) BEFORE any
customer/external login — covers existing customer-role logins too.

## How to verify
- Anon REST GET each of the 5 financial views → **401**.
- `v_kpi_dance_infusion` → DI#1 **39%**, DI#2 **31%** (% to mission = 1 − cost_to_raise).
