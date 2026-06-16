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

## This session shipped (2026-06-16 — Attendee + mailing backfill with lifetime guest stats)
Migration 037 applied to prod (additive). Backfilled attendees → guests → subscribers, money untouched.
- **Sources:** 3 RA ticket exports (DI#1 42, Crossroads showcase 4, DI#2 20) = 66 rows → **45 unique guests** (dedupe by email). 0 no-email, 0 malformed, **0 explicit opt-outs**. The DI#2 door-list xlsx (81 names, **no emails**, fuzzy spellings overlapping RA) was **flagged, not imported** (no dedupe key → would dupe).
- **Rule applied:** subscribe everyone with email except explicit opt-outs → **all 45 subscribed** (none opted out). **Deliverability exposure flagged:** 20 cold-only guests (bought a ticket, never ticked RA marketing) — DI#1 15/28, Crossroads 1/3, DI#2 14/16. Tagged by source/segment so a cold-set unsubscribe is a one-liner if Keith wants.
- **Schema 037:** `guest_event_attendance` (additive guest↔event link carrying `amount_spent`) + `v_guest_stats` view (events_attended count+list, total_spent, first/last seen, subscribed). **Deliberately did NOT write `ticketing` rows** — that feeds `v_event_summary.ticket_revenue` and would double-count DI#1's reconciled income. Off-prompt: guest "total spent" comes from the link, event financials untouched.
- **Result:** guests 45, subscribers 45 (all subscribed), subscriber_segments 47 (per-event), attendance 47. **No dup guests/subscribers** (dedupe by email). Multi-event proven: Claudia (DI#1+DI#2 $65.70), Liz McQuillan (Crossroads+DI#1 $70). **Money untouched: ticketing=3, income=6 unchanged.** Backups in `backups/preattendee_2026-06-16_*.json`. Idempotent, tagged `[ATTENDEE BACKFILL 2026-06-16]`.
- **Surfacing:** Subscribers tab now shows real list + **Events + Total spent** columns (joined from `v_guest_stats`). Campaigns tab can send to the 45.
- **Flagged for Keith:** the 20 cold subscriptions (deliverability); the 81-name door list (no emails). See `ATTENDEE_BACKFILL_DRYRUN.md`.
- Deployed to master (Netlify). Existing tabs untouched except Subscribers enrichment.

## This session shipped (2026-06-16 — Sprint 4: chain fix + end-to-end verify + actors-only backfill)
Theme = interconnection (prove the links, not the nodes). Migrations 035 + 036 applied to prod (additive).
- **Venue-save bug = display/refetch**, not a save failure (DI#2.venue_id was correctly = Signal). The hub showed the name via a PostgREST FK-embed (`venue:venues(name)`) that returned null client-side; rest of `hubLoadEvent` already fetched owner explicitly. **Fix:** fetch venue explicitly by id (drop the embed). Now persists AND displays.
- **One gear task, not N:** migration 035 generator emits a single "Load / test / setup gear (see equipment sheet)" + new **`v_event_equipment_sheet`** view; Equipment tab has a printable **Load-in sheet**.
- **Venue-as-counterparty (A-a):** `venues.actor_id` (additive FK). Contract form preselects the venue's linked actor; "Make '<venue>' contractable" creates+links an org actor inline.
- **Delete-resurrection fixed:** generator suppression now matches a title that exists for the event INCLUDING soft-deleted → a deleted task is never re-added. Idempotency + future-only edits preserved.
- **036:** seeded a "Finalize & sign venue contract" outreach template (venue:booking) so the chain includes a contract task.
- **End-to-end verify (Phase 2 gate, rolled back) on real DI#2:** venue→contacts surface→one gear task→rider→sound, load-in→booking, contract→booking→venue contractable→deleted stays deleted. All green.
- **Backfill (Phase 4 — only persisting step):** actors + people-links ONLY, no money. Dry run found NO empty-in-DB+spreadsheet-money event → no POPULATE (DI#1 money already reconciled in the 2026-06-02 load; the `*-list.csv` files are attendee exports, not actor sources). **Created 3 actors** (32LVS, Gavin/Sara of Signal), **linked existing** (Keith→DI#1 dj solo-run, Kristen London→DI Artist Showcase), **5 people-links** (3 participants + Gavin=sound/Sara=other on Signal). Tagged `[BACKFILL 2026-06-16]`; idempotent; **DI#2 money counts identical before/after** (backup in `backups/prebackfill_2026-06-16_*.json`). Actors 20→23, no dups. Interconnection proven: backfilled Gavin auto-assigns DI#2's rider.
- **Flagged for Keith (not guessed):** Crossroads Café Artist Showcase roster; Rich Klein affiliation; Sara's exact function; the 42 DI#1 ticket-buyers (attendees). See `BACKFILL_DRYRUN.md`.
- **Off-prompt UX:** explicit-fetch over embed (matches owner pattern); load-in sheet as a print modal; venue-contractable affordance inline in the contract form; backfill deliberately conservative (people-only) because the DB already holds the money.
- Deployed to master (Netlify). Existing 16 tabs / prior hub work untouched.

## This session shipped (2026-06-16 — Sprint 3: Venue/contact matrix (3a) + conditional workflows & template editor (3b))
Resumed cleanly after a mid-3b.3 network drop (template-editor JS was the only unfinished piece — nav/panel/loadTab branch existed but `loadTemplates` et al. were missing; finished them). Both migrations were already applied to prod; re-verified objects + re-ran both test suites green before deploy.
- **Migration 033 (applied):** `venue_contacts` + `vendor_contacts` link tables (contacts are ACTORS — no parallel people table); `v_venue_contacts` / `v_vendor_contacts` views (`security_invoker`, anon-revoked) carrying a `last_event_with` recency column = the SEAM for future frequency ranking (v1 only ORDERS by it — no ML).
- **Migration 034 (applied):** `events.cw_providing_gear`; `task_templates.gear_applicability/target_function/sort_order` + unique(event_type,title); 10 seeded outreach templates (party+DI); `generate_day_of_tasks` rewritten — gear-branch + all-phase generation with due_date offsets + outreach AUTO-ASSIGN via the matrix (degrades to unassigned-with-hint).
- **3a UI:** new **Venues** tab (Venues|Vendors toggle, list→detail, contact matrix add/edit/remove, "last event here"); hub Overview venue box (set/change venue inline + ordered "contacts here" + one-click **involve**).
- **3b UI:** gear checkbox in Edit Event + hub "Gear" fact; "Generate task checklist" callout reflects gear mode; **Templates** tab editor (group by type→phase, add/edit/remove/reorder, gear applicability + target function; copy says **future-only**); **assign-task picker now grouped** (This event's people · Your team always · Venue contacts) — replaces show-everyone.
- **Off-prompt UX:** Venues + Templates as their own sidebar tabs (sanctioned surfaces); contact ordering primary→recency→last-touch; template reorder renumbers the group sequentially (clean given default sort_order=0); involve-contact one-click from the lookup.
- **Tests:** `tests/contact_matrix_test.sql` (3a gate) + `tests/conditional_workflows_test.sql` (3b) — both green, rolled back, zero persisted. JS parse-clean (170 KB module). 16 panels intact; existing 14 tabs + hub sprints 1&2 untouched (additive only). Deployed to master (Netlify).
- **Deferrals:** contract signing flow; global cross-event task board; frequency-weighted contact ranking (seam left in `last_event_with`).

## This session shipped (2026-06-16 — Event Hub Sprint 2: UX pass + Money bug + IG KPI)
Fixed the Money-section bug and closed the hub's UX gaps. **Migration 032 APPLIED to prod** (additive):
`event_participants.roles text[]` (backfilled, one-per-actor unique index), `generate_day_of_tasks`
repointed to read `roles[]` overlap, and **audit triggers on tasks/contracts/event_participants** (reuse
`audit_trigger_function`) so status/role transitions land in `audit_log` (answers the reporting Q10).
- **Money bug root cause:** the hub Money section only launched a modal and never listed line items —
  so adds showed in Overview's rollup but "vanished" in the Money tab. **Fix:** Money renders line items
  **inline** in the hub, refetching on every write; shared `loadMoneyData`/`moneySectionsHTML`/`moneyMutate`
  back both the hub-inline section and the Events-tab modal (one source of truth). Verified all 5 child types.
- **Multi-role participants:** one row per person, `roles[]` multi-select chips (custom roles allowed),
  `role` kept = `roles[1]` for back-compat. Day-of generator now catches a secondary `dj`/performer.
- **People UX:** bulk "Add people" (search + multi-select existing + stage new + batch roles/fee in one
  insert); full participant **edit** (roles/fee/bill_order/set times/contractor); fee→expense unchanged.
- **Equipment:** multi-select assign (batch purpose/dates) + edit/remove.
- **Contracts:** edit + inline document upload → `files(subject_type='contract')` → `contracts.document_id`,
  signed-URL download. (No signing flow yet — still deferred.)
- **IG KPI:** "Log IG" one form / 3 accounts (comewith/berky/danceinfusion) → upserts today's
  `metric_snapshots`; live ▲/▼ delta vs last; entry points on **Strategy toolbar** AND the **hub header**.
- **Off-prompt UX choices:** batch-default roles/fee then per-person edit (vs per-person-at-add); IG quick
  action placed in the hub **header** (always one click from landing) rather than buried in Overview;
  Overview "Open Money" now jumps to the inline section (no modal). Multi-role stored as `roles[]` (not a
  child table) — simplest additive path given how the hub + generator read roles.
- **Tests:** `tests/event_hub_sprint2_test.sql` (rolled-back, all green). **14 existing tabs untouched**
  except the Strategy toolbar's new "Log IG" button. Deployed to master (Netlify).
- **Note:** prod now has Keith's own sprint-1 live-test rows (1 contract, 1 equipment_usage, 1 task) — real, preserved.

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
