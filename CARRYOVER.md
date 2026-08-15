# Carryover — 2026-08-15 (radio discovery audit close · ran on the DESKTOP)

Pickup order: this → `DEV_DOCS/claude-memory/MEMORY.md` → `LEARNINGS.md` → `ROADMAP.md` → `CLAUDE.md`.
Ritual: `SESSION_CLOSE_PROMPTS.md`. DI data load detail: `events/dance-infusion/DI_DATA_LOAD_LOG.md`.

## 👉 If you are the LAPTOP, start here

Nothing is lost — but two things about this repo aren't obvious from a fresh checkout:

1. **Claude memory doesn't sync between machines.** The desktop's 60 memory files
   are snapshotted into **`DEV_DOCS/claude-memory/`** (index: `MEMORY.md`). Read the
   index; open individual files as needed. Re-snapshot at every close.
2. **The radio window fix is MERGED AND LIVE** (2026-08-15, `88b2153`) — dashboard on
   Netlify, `pull-dice` v5 + `pull-ticketmaster` v7 on prod. Just `git pull`.
   **It has not been exercised yet** — see "Parked / next" item 1 for the one click
   that proves it.

## ⛔ PRIORITY CONTEXT (⚠ dated 2026-06-02 — unverified, confirm with Keith)
**Come With was set MAINTENANCE-ONLY** while the **CWF (Come With Fitness) BRD** ran as
project #1, due **June 15, 2026** — two months past. Actual work since has been steady
Come With radio/dashboard building, so this framing is stale. What still stands, and is
a hard rule either way: **nothing Come With Fitness in this repo** (dashboard / schema /
pages) — LEARNINGS §5 and CLAUDE.md "Scope".

## State summary (verified against prod 2026-08-15)
- **Prod:** Supabase `yaytdosxfhcqatmhctzk`; live at comewith.org (Netlify auto-deploy from `master`).
- **Migrations: files through `137_show_counter_wording.sql`.** Applied via the Management
  API (`db.py`, `SBP_PAT` in `.env`), not the CLI — the CLI is linked to **staging**, so
  always pass the prod ref explicitly. The migration **files** are the tracked source of truth.
- **Financial views:** all five re-verified anon **401** on 2026-08-15 (`v_event_summary`,
  `v_kpi_event_financials`, `v_kpi_parties`, `v_kpi_dance_infusion`, `v_kpi_dashboard`).
  ⚠ Still **NOT revoked from `authenticated`** — the GATED BLOCKER before any customer/external login.
- **Roles:** master_admin / sub_admin / customer via `public.is_admin()`; `donor` + `staff` on `actors`.
- **Latest LEARNINGS §:** 15.
- **Git:** `master` = `88b2153`, pushed. `radio/window-by-lineup` **merged**. Older unmerged
  branches: `fix-lognumbers-optgroups`, `docs/roadmap-reconcile`, `event-hub-sprint-1`.
- **Edge-function deploys go through `scripts/deploy_edge_function.py`**, not the CLI. CLI
  2.101.0 rejects the newer `sbp_v0_…` PAT format outright ("Invalid access token format")
  *and* is linked to staging; the Management API takes the same token fine. The script
  preserves each function's existing `verify_jwt` rather than resetting it.
- **Radio discovery pool (prod, as of the 2026-08-14 pull):** 1,122 future RA events,
  159 DICE, 27 TM; 2,061 RA artists with a SoundCloud link. **No cron pulls any of this** —
  `cron.job` runs only publish/retention/YouTube. The pool is as fresh as the last manual
  "↻ Pull shows".
- **Docs freshness:** `ROADMAP.md` is reconciled to **2026-06-02** and predates the whole
  radio build; treat CARRYOVER + `DEV_DOCS/claude-memory/` as the true state.
- **Tools:** `/tools/actor-inspector.html` · `/tools/test-checklist.html` · `/tools/visualizer.html`
  deployed on comewith.org, admin-gated via the staging guard.

## Tomorrow's default
**Click "↻ Pull shows", then "↻ Refresh music & data"** on the Radio panel (items 1–2). The
fix is live but unexercised, and the pool is still the 2026-08-14 pull with 677 unscanned
artists in it.

## Parked / next
1. **Exercise the deployed fix — nobody has run a pull through it yet.** Hit "↻ Pull shows"
   and check: the toast reports DICE, TM and RA counts (and names any source that didn't
   answer); DICE now reaches past week 1; Ticketmaster returns Brooklyn/Queens venues. Then
   confirm the artist count for a **future** start date jumps — that's the 77. Couldn't be
   done from here: invoking the pulls needs an admin JWT or the service-role key, and neither
   is on this machine (`.env` carries the publishable key only).
2. **Scan the 8/18 window.** 677 of 825 in-window artists with a SoundCloud link have never
   been read (148 scanned → 89 producers). "↻ Refresh music & data" on the Radio panel. It
   only reads artists with no cache row at all, so it's safe to re-run.
3. **`dj-station` caps its artist query at `.limit(160)`** with no notice, against a window
   holding ~879 artists — a DJ with a broad genre filter is silently seeing a fraction of the
   crate. Found during this audit, **not fixed**, needs a paging/notice decision.
4. **Scan cache never re-reads.** `raScan` skips anyone with a cache row forever; a producer
   who released tracks since their scan shows the stale catalog until "↻ re-read all" (which
   nukes and re-reads the whole window). Only 2 in-window artists are >30 days stale today —
   not yet urgent, will be.
5. **Financial views still readable by `authenticated`** — the standing GATED BLOCKER before
   any customer/external login.

## This session shipped (2026-08-15 — Radio discovery audit: the window was filtering on the wrong date)
Full audit of shows → producers → tracks, run against prod. **No migration, no prod writes.**
**MERGED + DEPLOYED** (`88b2153`): dashboard live on Netlify (verified — `raWindowPool` is in
the served file), `pull-dice` **v5** and `pull-ticketmaster` **v7** live on prod (verified by
reading the deployed source back). Financial views re-checked anon **401** after the deploy.
- **The bug:** `ra_artists` collapses each artist to ONE row carrying `next_event_date` — their
  *soonest* show. The radio window filtered on that column, so an artist playing the 16th **and**
  the 25th vanished from a window starting the 18th. Keith's 8/18 + 4-week filter was hiding
  **77 artists, 70 with a SoundCloud link.** Fix indexes every date from `ra_events.lineup`;
  simulated on prod it recovers exactly those 77 (866 → 943), and re-points each artist at the
  show they play *in* the window (date, venue, that bill's genres).
- **DICE was 7 days deep, not 90.** It detail-fetched the first 160 candidates in tag order and
  saved 159 — the cap was binding exactly, so weeks 2–4 had **zero** DICE and nothing said so.
  Now date-filters before spending a fetch, takes soonest-first, caps at 240, reports
  `dropped_over_cap` + `last_date`.
- **Ticketmaster was Manhattan-only.** `city: "New York"` is a literal match at TM's end: 27
  future shows, 11 venues, nothing in Brooklyn or Queens. Now queries all five boroughs.
- **"Pull shows" swallowed TM/DICE failures**, so an outage and a real zero looked identical.
  It now names a source that didn't answer.
- Also confirmed **clean**: RA coverage itself (582 in-window events, listings to 11/10, taper
  looks like genuine listing decay, not truncation), and all five financial views at anon 401.
- **Docs:** `DEV_DOCS/claude-memory/` snapshot added so the laptop can read project state;
  `CLAUDE.md` gained a start-of-session / close-of-session section; `SESSION_CLOSE_PROMPTS.md`
  gained an "every close" block (re-snapshot memory, name the machine, don't merge to master
  to tidy up).

## This session shipped (2026-06-16 — Guest KPI fix: returning-match + filter-aware cards + mission/business split + ROADMAP)
Diagnose-first, then fix. Migration 040 applied to prod (additive views). Money untouched throughout.
- **Diagnosis (read-only gate):** DI#1 "28" is CORRECT — 42 RA rows = 42 tickets but 28 distinct buyers (10 multi-buys → 14 anonymous extra seats; RA captures only the buyer). The returning "1" had TWO bugs: (a) the sprint-6 ledger import skipped the 25 overlap emails → DI#1 people who also came to DI#2 got NO DI#2 attendance link (27/28 DI#1 guests lacked one); (b) returning was guest_id/email-exact, missing different-email/nickname same-person (Ethan Pollak, Liz↔Elizabeth).
- **Fix A (recovery, additive, money-free):** added the missing DI#2 `guest_event_attendance` links for existing guests in the ledger → DI#2 attendance 68→77 (+9). Backup `backups/prekpifix_2026-06-16_*.json`.
- **Fix B (migration 040):** `v_event_attendance_kpi` rebuilt with PERSON-identity matching (normalized full name + nickname canon, email fallback) — a KPI calc, not a record merge. **DI#2 returning 1 → 12** (43% of DI#1's 28; honestly below Keith's 50–75% gut because DI#1 has only 28 *identified* buyers, not 42 attendees, and some returnees used unmatchable identities). No false merges.
- **Filter-aware cards:** Guests-tab KPI cards now recompute for the active filtered set (event/returning/subscribed), immediate.
- **Mission/business split (`v_guest_spend_split`, type-driven):** DI = **mission** (MS-Society), CW Parties = **business**, else other. Surfaced in list + profile + cards, labeled "DI / Mission" vs "CW / Revenue" — never conflated.
- **Verify:** DI#1 $2,940 / DI#2 $9,557.33 financials identical before/after; guests still 97 (recovery added links, not guests); attendance 99→108.
- **ROADMAP.md reconciled** to true prod state: DONE (module series 030–040), QUEUED (audit cleanup → Artist → Vendor → Sponsor repoint → Equipment module), DECISIONS-WAITING (cold subs, door list, audit merges), GATED BLOCKER kept (financial-view lockdown before external login), pointer line kept.
- Deployed to master (Netlify). Only guest/KPI surfacing touched.

## This session shipped (2026-06-16 — Guest module: ledger import + actor graduation + repeat KPI + reconcile + audit)
Migrations 038 + 039 applied to prod (additive). DI#2 ledger people imported, money NEVER written.
- **Import:** ledger 113 rows → 77 emails → 25 overlap skipped → **52 net-new guests** + DI#2 `guest_event_attendance` (amount_spent = guest stat only, NO ticketing/income). Guests **45→97**.
- **Consent:** 0 net-new opt-outs; the **11 `ra_marketing_opt_in=False` were already-subscribed guests → UNSUBSCRIBED** (consent correction). Subscribers = 97 rows, **86 subscribed / 11 unsubscribed**.
- **Actor graduation:** `guests.actor_id` (mig 038). 13 LINK to existing actors (variants resolved: Elizabeth→Liz, Sauci→DJ Sauci Soni, Keith→Berky — no dups), **15 CREATE** clean donor/sponsor/staff, **15 FLAG** (fuzzy dups + payout artifacts — not created/merged). Actors **23→38**; 15 guests graduated. Mig 039 added `'staff'` to actor_roles enum (additive, like 029's 'donor').
- **Money untouched (proven):** DI#2 gross $9,557.33 / exp $6,557.33 / sponsor $6,225 / donations $292.44 — identical before/after. Dedupe: 0 dup guests/subs/actors. Backups in `backups/preguest_2026-06-16_*.json`.
- **Guest module (new Guests tab):** v_guest_stats list + filters (search, event, **subscribed-only = mailing list**, returning-only) + per-guest profile (history, spend, subscribed, actor link). KPI strip from `v_guest_kpis`; per-event new/returning from `v_event_attendance_kpi`.
- **Returning KPI:** DI#1 28 new/0 ret; Crossroads 2 new/1 ret (Liz); DI#2 67 new/1 ret (Claudia). 97 guests, 86 subscribed, 2 repeat, avg spend $58.87.
- **Reconciliation (read-only, NOT a double-count):** ledger $19,114.33 = gross mixed activity — sponsorship $9,950 + expense $4,750.33 + ticket $1,513 + payout $1,291 + zeffe_pkg $860 + donation $750 + comp $0. Differs from DB's reconciled net by scope (expenses/payouts/comps + in-kind sponsorship + DIY donation stream). Soft flag: ledger shows a $750 individual-donation stream + $9,950 sponsorship (incl in-kind) the DB headline consolidates differently — review if MS-facing granular totals are needed. No financial data changed.
- **Audit (`GUEST_ACTOR_AUDIT.md`, flag-only, nothing merged):** 6 same-human guest↔actor unlinked (Adam Cohen, Ethan Pollak, Francis/Theresa Berkman, Liz, Patrick — actor has no email so email-link didn't connect); 8 possible-dup/variant (Crossroads Café accent, Teri↔Theresa, etc. — 3 are likely *different* people, don't merge); Henry low-quality name; 5 payout artifacts excluded.
- Deployed to master (Netlify). Existing tabs untouched except Subscribers (prior) + new Guests tab.

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
