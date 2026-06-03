# Come With — Data Architecture Concept

**For:** handoff to Claude Code (introspect-and-implement, not literal final SQL)
**Goal:** an agile data model serving TWO purposes on one foundation:
1. **Analytical** — a visualizer that pulls any data point / KPI over time, slices by any entity, and lets new formulas be defined without schema changes.
2. **Operational (workflow)** — the dashboard as a working tool: create an event, assign actors, manage tasks (assignable to actors), track contracts, and manage budget (planned vs. actual). Equipment, artists, guests, sponsors, vendors, contractors, content, contracts, tasks, and money all link to events.

The shared actor + event model is what lets ONE system do both. The actor model also means **roles overlap freely** — the same actor can be customer + sponsor + vendor + artist simultaneously (e.g. Crossroads Café = vendor + sponsor; a DJ = artist + dj + customer). Overlap is the normal case, not an edge case.
**Status:** concept / decisions doc. Claude Code should introspect the live schema first, reconcile against what already exists, and propose migration SQL before applying. **Do not greenfield — evolve the existing schema.**

---

## 0. The core tension, and how this resolves it

Two requirements pull against each other:

- **"Super agile data"** → wants flexibility: add metrics, link anything, invent new analyses without migrations.
- **"Link everything cleanly, compute real KPIs"** → wants integrity: foreign keys, consistent rollups, no garbage.

Pure flexibility (everything is a key-value blob) makes KPIs uncomputable and data untrustworthy. Pure rigidity (a column for everything) means a migration every time you have a new idea. **The resolution is three layers, each with a different rigidity:**

```
LAYER 3 — DEFINITION   (config-as-data: flexible)
   metric & formula definitions; KPIs are ROWS, not code.
   Add a metric / new formula = insert a row. No migration.
        ▲ reads from
LAYER 2 — MEASUREMENT   (uniform shape: semi-flexible)
   every data point — hand-entered OR computed — resolves to
   ONE shape the visualizer reads. Time-series + cross-entity
   both fall out of this uniform stream.
        ▲ derives from
LAYER 1 — ENTITY & RELATIONSHIP CORE   (structured: rigid, FK-enforced)
   the source of truth. Actors, events, the links between them,
   money, equipment, content. Real referential integrity here.
```

The trick: **rigidity at the bottom (truth), flexibility at the top (analysis).** You get clean data AND an endlessly explorable visualizer.

---

## 1. LAYER 1 — Entity & Relationship Core

### 1.1 The Actor model — the big unlock ("the whole web")

**Problem today:** `artists`, `contractors`, `clients` are separate tables. A DJ who is also a customer, an artist who also gets booked as a contractor, a painter you engage for a show — none of these can be represented without duplicating the person across tables. This is almost certainly the "everything wasn't linked right" issue.

**Fix: one table for every person and org. Roles are relationships, not table membership.**

```
actors                          ← every person OR org, ONE row each
  id (uuid)
  kind            person | org
  display_name    "Kloud9", "AOM Infusion", "Kristen London"
  legal_name      (optional)
  email, phone, instagram, website, notes
  created_at, updated_at, deleted_at

actor_roles                     ← STANDING roles (who someone IS to us)
  id
  actor_id  → actors
  role      artist | dj | contractor | customer | sponsor |
            team | performer | painter | vendor | venue_contact ...
  context   (optional notes / since-date)
  active

  → Kloud9 = one actor, with actor_roles {artist, dj}.
    If Kloud9 books an equipment rental, ADD role {customer} —
    same actor, no duplication. The whole web, one identity each.
```

**Standing role vs. event role:** `actor_roles` is "is in our roster as X." What someone DID at a specific event lives on the event link (1.3) — e.g. the same artist "DJ'd at DI#2" vs. "painted at the showcase." Don't conflate the two.

**Migration note for Claude Code:** the existing `artists`, `contractors`, `clients`, `sponsors` rows must be merged into `actors` + `actor_roles` (dedupe by email/name where the same person appears in multiple tables — e.g. someone who's both an artist and a client). Old tables can become **views** over `actors` during transition so nothing downstream breaks immediately, then be retired. Propose the dedupe/merge plan before running it.

### 1.2 Events as the hub

`events` largely exists. Confirm/extend it to carry the **multi-axis** model from the prior design session:

```
events
  id, slug, name, event_date, doors_time, end_time
  type            party | dance_infusion | production | showcase
                  ← replaces flat 'series'; the canonical category
  status          draft | confirmed | completed | cancelled
  venue_id        → venues
  capacity
  is_content_event  bool   ← TRUE for showcases (graded on content,
                             not P&L); see 1.5
  notes
  created_at, updated_at, deleted_at
```

- `type` is the canonical axis (Party / Dance Infusion / Production / Showcase). Dance Infusion is a Come With production with a fundraising P&L lens — same metrics as Parties **plus** total_raised + cost-to-raise (per prior session). Production = work-for-hire (e.g. Maxwell House). Showcase = content-first.
- **Signature tags** (e.g. "booth-to-wall, crowd-connected") are NOT a type — they're tags. See 1.6.

`venues` exists. Keep. (id, name, address, capacity, contact actor_id.)

### 1.3 event_participants — people ↔ events (the missing link)

This is the table that lets you "link artists / performers / crew to shows," and it handles the painters/dancers/performers you mentioned because **role is just a value**:

```
event_participants
  id
  event_id    → events
  actor_id    → actors
  role        headliner | dj | opener | painter | dancer |
              performer | host | crew | photographer | producer ...
  bill_order  (int, for lineup ordering)
  set_start, set_end   (optional)
  fee         numeric (what you paid them — feeds expenses/P&L)
  is_contractor bool   (paid crew vs. featured artist — the
                        "lineup vs. crew" distinction from before)
  notes
```

- "Who played DI#2" = query event_participants where event_id = X.
- A painter at a showcase = same table, role='painter'. Endlessly extensible without schema change.
- `fee` is the bridge to money: a participant's fee is an event expense (see 1.4) — decide whether fee here auto-creates an expense row or is reconciled (open decision Q3).

### 1.4 The money layer (largely exists — keep, ensure consistent linkage)

These exist and the prior session fixed the rollup math. **The rule: every money row links to an event_id (nullable for non-event income).**

```
ticketing            event_id, actor_id(buyer, optional), tier,
                     quantity, amount_paid (line total)
income               event_id, amount, category (bar/vendor/merch),
                     deleted_at        ← "other income"
expenses             event_id, amount, category, deleted_at
third_party_donations event_id, donor (actor_id?), amount
sponsorships         event_id, sponsor actor_id, cash_amount,
                     in_kind_value, tier, status
```

**Canonical money definitions (locked prior session — preserve exactly):**
```
gross_revenue = ticket_revenue + other_income + donations + sponsor_cash
net_pl        = gross_revenue − total_expenses          (ALL events)
total_raised  = gross_revenue + sponsor_in_kind          (DI framing)
% to mission  = donated_to_mission / total_raised        (DI public metric)
```
Internal speaks "expense ratio"; public speaks "% to mission." Keep that split.

**Sponsorships now point at `actors`** (sponsor = an org/person actor with role 'sponsor'), replacing a standalone sponsors table — same party-model unification.

### 1.5 Equipment ↔ events

`equipment_inventory` + `equipment_usage` exist. `equipment_usage` is the link table:

```
equipment_usage (≈ exists)
  event_id    → events
  equipment_id → equipment_inventory
  role/purpose
  revenue_attributed   (for equipment ROI — v_equipment_roi exists)
```
Gap noted in prior session: **nothing writes equipment_usage yet.** The Log Event flow (and the per-event Money/Links panel) should let you attach gear used at an event. This feeds equipment ROI per event and per event-type.

### 1.6 Content & signature tags

**Content is its own axis** (prior session): a showcase event exists to *produce content*; any party/DI event *may also* produce content. So content is an entity that can reference an event, graded on views/follows — not a P&L line.

```
content_series (exists)   Backyard Sessions, Live at Dance Infusion, ...

content_items             ← NEW: individual pieces (a video, reel)
  id
  series_id   → content_series
  event_id    → events (nullable — content can be standalone OR
                from an event)
  title, platform (youtube|instagram|...), url, published_at

  → graded via metric_snapshots (subject = content_item):
    views, follows, watch_time. NOT P&L.
```

**Signature tags** — a generic tagging system so "booth-to-wall" (and future signatures) can mark any event or content, and the visualizer can slice by it:

```
tags            id, name ("booth-to-wall"), kind
taggables       tag_id, subject_type (event|content_item|actor),
                subject_id
  → many-to-many, applies to anything. Slice KPIs by signature.
```

---

## 2. LAYER 2 — The Measurement Layer (the visualizer's fuel)

**This is what makes "pull in whatever data I want" work.** Every measurable thing — whether hand-entered or computed — resolves to ONE uniform shape:

```
A DATA POINT =
  ( metric_key, subject_type, subject_id, captured_at, value, dims )
```

Two sources flow into this one shape:

**(a) Snapshots — manual / external readings** (`metric_snapshots` already exists):
- IG followers (per account), YT subs, watch time, per-series content views.
- subject = an actor, a content_series/item, or global. captured_on = the date.

**(b) Derived metrics — computed from Layer 1** (views):
- net_pl, total_raised, % to mission, sell-through, attendance, sponsor_count, ticket_revenue, donations, equipment ROI, headcount by role, etc.
- subject = an event (or rolled up by type/venue/tag).

**The unifying abstraction — one view the visualizer reads:**

```
v_data_points  (UNION of snapshot + derived, same columns)
  metric_key      'instagram.followers' | 'event.net_pl' | ...
  subject_type    actor | event | content_item | content_series | global
  subject_id
  captured_at     date
  value           numeric
  dims            jsonb   (event_type, venue, series, tags... for slicing)
```

Why this is the whole game:
- **Time-series** = filter metric_key, group by captured_at.
- **Cross-entity** = filter metric_key + subject_type, compare subject_ids.
- **Slice** = filter on dims (event_type='dance_infusion', tag='booth-to-wall').
- The visualizer never cares whether a number was typed or computed. Uniform stream.

**Scale note** (you're hoping this gets big): at hundreds→thousands of events, computing `v_data_points` live may slow down. Use a **materialized view refreshed by pg_cron** (you already have pg_cron from migration 014) for the derived side, with the snapshot side read live. Decide refresh cadence (nightly is fine for most KPIs).

---

## 3. LAYER 3 — Definition Layer (metrics & formulas as data)

So KPIs and **new formulas** are rows, not code:

```
metric_definitions
  metric_key       unique  'event.net_pl', 'di.cost_to_raise', ...
  label            "Net P&L", "Cost to raise $1"
  unit             $ | % | hrs | count | ''
  category         content | parties | audience | dance_infusion | ...
  comparison       gte | lte         (higher-better vs lower-better)
  source_type      snapshot | derived | formula
  source_ref       for snapshot: the snapshot metric_key
                   for derived:  the view/column it reads
                   for formula:  an expression (see below)
  active

kpi_targets (exists)  versioned targets per metric_key. Keep.
```

**Formulas — "analyze data in new and exciting ways."** A formula metric composes other data points. Example: `revenue_per_attendee = event.total_income / event.attendance`. THREE tiers of how far to build this — **recommendation: start at Tier 2**, per your "is there a simpler way / don't pre-bundle" principles:

```
TIER 1 (simplest)  Each derived metric = a hand-written view.
                   New formula = Claude Code adds a view. Bulletproof,
                   least agile. No runtime engine.

TIER 2 (recommended start)  metric_definitions of source_type='formula'
                   hold a SIMPLE expression over other metric_keys
                   (+ − × ÷, parentheses). A small, safe evaluator
                   computes them. New formula = insert a row in the UI.
                   Agile, moderate build, no arbitrary code execution.

TIER 3 (later, only if needed)  Full expression language —
                   conditionals, aggregations, windowing. Maximum
                   power, real build + maintenance. Wait until Tier 2
                   genuinely limits you. Don't pre-build.
```

The visualizer's "create a new metric" UI writes a Tier-2 `metric_definitions` row. That's the agility you asked for, bounded so it can't break integrity.

---

## 4. The Visualizer — data contract

The visualizer needs only three things from the DB:

1. **`metric_definitions`** — the menu of available metrics/KPIs (incl. user-created formulas).
2. **`v_data_points`** — the uniform stream. Query: pick metric_key(s), optionally filter subject_type/subject_id/dims, get (date, value) series.
3. **`kpi_targets`** — for target lines/markers on charts.

With that, the visualizer can: plot any metric over time; compare any metric across entities (artists, events, venues, content); overlay targets; and render user-defined formula metrics — all without schema changes, because new metrics are rows and the stream shape is fixed.

---

## 4.5 OPERATIONAL LAYER — workflow, contracts, files, touchpoints

The dashboard isn't only a record of what happened — it's where you *run* events. This layer sits on the same actor + event core. Everything below is structured (Layer 1 rigidity) because it's operational truth, but it feeds the analytical layers too (a task completion rate, a planned-vs-actual variance, are themselves KPIs).

### 4.5.1 Event lifecycle (status is a stage, not a flag)

Events move through stages; workflow needs the lifecycle, not just a status string:

```
events.stage   idea | planning | confirmed | live | wrapped | reported
  (distinct from status=cancelled, which can happen at any stage)
events.owner_actor_id  → actors   ← who OWNS this event (you / Liz),
                                     distinct from task assignees
```

### 4.5.2 Tasks & assignments (live in THIS dashboard)

Tasks live in Come With so assignees can manage them directly. Jennifer link is a later, simple export/import for Keith's own tasks only — NOT a live two-system sync now (keep it simple; design the bridge later).

```
tasks
  id
  event_id      → events (nullable — some tasks aren't event-bound)
  title, description
  status        todo | doing | blocked | done
  priority      (optional)
  due_date
  effort, reward (optional — mirrors Keith's scoring habit)
  created_by    → actors
  source        manual | template | jennifer_import
  created_at, updated_at, deleted_at

task_assignments              ← tasks ↔ actors (many assignees)
  task_id   → tasks
  actor_id  → actors
  role      owner | doer | reviewer
```

- Assigning a task to an actor uses the SAME actors table — assignees are just actors with a login (or not). Team members (Liz, Kendall, etc.) are actors with role 'team' + a profiles/auth link.
- **Jennifer bridge (deferred, keep simple):** an export of Keith's own tasks to Jennifer and an import back — `tasks.source='jennifer_import'` marks synced ones. NOT a real-time integration. Design properly later.

### 4.5.3 Task templates — recurring checklists per event-type

The biggest workflow time-saver: each event TYPE has a default task list, so a new DI / party / showcase doesn't get rebuilt from scratch.

```
task_templates
  id
  event_type    party | dance_infusion | production | showcase
  title, default_offset_days  (e.g. "book venue" = T-60,
                               "confirm DJs" = T-30)
  default_role  (which kind of actor usually owns it)
  phase         planning | promo | day_of | wrap

  → Creating an event of type X spawns its template tasks,
    auto-dated from the event_date by offset.
```

**Equipment-driven DAY-OF checklist (Keith's ask):** the day-of checklist is partly *derived*, not static. Once equipment and participants are assigned to an event, generate day-of tasks from them:

```
- For each event_equipment row → a day-of "load/test/setup [gear]" task
- For each event_participant (dj/performer) → a "soundcheck / confirm
  arrival [name]" task
- Plus the fixed day_of template tasks (doors, float, etc.)
```

So the day-of checklist = fixed template tasks + generated-from-assignments tasks. This is the one genuinely dynamic template piece — model it as a generation step that runs when assignments change or on demand, not as static rows.

### 4.5.4 Contracts / agreements (first-class, tied to actors + events + budget)

You already have agreement docs (Events Services Agreement, Contractor agreements as .docx; the existing agreements workflow). Structure them:

```
contracts
  id
  event_id      → events (nullable for master/standing contracts)
  actor_id      → actors  (the counterparty — vendor, contractor,
                           client, artist)
  kind          event_services | contractor | vendor | rental | sponsor
  fee           numeric   ← the obligation
  status        draft | sent | viewed | signed | countersigned | void
  sent_at, signed_at
  document_id   → files (the actual PDF/docx; see 4.5.5)
  notes
```

- **Tied to budget:** a contract's `fee` is a *planned* expense (or planned income, for client contracts) — it flows into the budget layer (4.5.6) as an obligation before it's an actual. When paid, it reconciles to an `expenses`/`income` row.
- The same actor model means a contract counterparty is just an actor — a vendor who's also a sponsor has one identity, possibly multiple contracts.

### 4.5.5 Files / attachments (general, on the existing storage buckets)

You already have Supabase storage buckets. A single attachments table links any file to any entity:

```
files
  id
  bucket, path           (the existing storage buckets)
  filename, mime, size
  subject_type   event | actor | contract | content_item | task
  subject_id
  kind           contract | rider | stage_plot | invoice | receipt |
                 photo | other
  uploaded_by    → actors
  created_at
```

Covers contracts (4.5.4 document_id), receipts/invoices for money, stage plots/riders for events, photos for content. One table, polymorphic subject — same pattern as tags/taggables.

### 4.5.6 Budget — planned vs. actual (per event, per type, overall)

Your money tables (income/expenses/etc.) are ACTUALS. Add a planning layer:

```
budget_lines
  id
  event_id      → events (nullable — a line can be type-level or global)
  scope         event | event_type | overall
  category      venue | production | talent | marketing | ... (match
                the public-audit cost groups)
  planned_amount  numeric
  direction     expense | income
  contract_id   → contracts (optional — a planned line backed by a
                  contract obligation)

  → ACTUAL comes from the existing money tables (expenses/income/
    ticketing/etc.) summed per event+category.
  → VARIANCE = planned − actual, computed in a view.
```

The three levels you asked for:
- **Per event:** budget_lines where event_id = X vs. actuals for X.
- **Per event-type:** roll up across all events of a type (planned templates + actuals).
- **Overall:** all events + non-event lines.

These variances are themselves data points → they flow into `v_data_points` (Layer 2) and become visualizer KPIs ("planned vs actual by event-type over time").

### 4.5.7 Touchpoints — CRM-lite (cheap on the actor model)

Logging interactions with sponsors/vendors/clients is nearly free once actors exist:

```
touchpoints
  id
  actor_id    → actors
  event_id    → events (nullable)
  kind        email | call | meeting | dm | note
  summary     text
  occurred_at
  logged_by   → actors

  → "last contacted AOM Infusion 3 weeks ago re DI#3 sponsorship."
    Becomes sponsor-pipeline / relationship tracking over time.
```

Keep it lean — a log, not a full CRM. The actor model means it can grow into pipeline tracking later without restructuring.

### 4.5.8 Content publish/embed state (the website surfacing)

`content_items` (1.6) tracks content as data; add the state for surfacing it on the public site:

```
content_items  (extend)
  + publish_status   draft | review | published
  + embed_on         jsonb / link table — where on the site it shows
                     (homepage, event page, DI report, artist page)
  + featured         bool
```

So content flows: produced → graded (views/follows via snapshots) → published/embedded on the website. The "content embedded in the website" you mentioned = the published subset, surfaced via embed_on.

---



**Claude Code: introspect live schema first; this is my best understanding, verify it.**

```
KEEP (exists, fits the model):
  events, venues, equipment_inventory, equipment_usage,
  ticketing, income, expenses, third_party_donations,
  sponsorships, guests, metric_snapshots, kpi_targets,
  content_series, the v_kpi_* views, profiles, is_admin()

NEW — analytical (the gaps that caused "not linked right"):
  actors                 ← unify people+orgs (THE big one)
  actor_roles            ← standing roles, kills duplication
  event_participants     ← people↔events (who played/painted/crewed)
  content_items          ← individual content pieces ↔ events/series
  tags + taggables       ← signature tagging, arbitrary slicing
  metric_definitions     ← KPIs/formulas as data
  v_data_points          ← the unified measurement stream
  (materialized rollup)  ← scale, pg_cron-refreshed

NEW — operational/workflow (Section 4.5):
  events.stage + owner_actor_id   ← lifecycle + ownership
  tasks + task_assignments        ← tasks, assignable to actors
  task_templates                  ← recurring checklists per type
                                    (+ equipment-driven day-of gen)
  contracts                       ← agreements: actor+event+fee+status
  files                           ← attachments on existing buckets
  budget_lines                    ← planned vs actual (3 levels)
  touchpoints                     ← CRM-lite interaction log
  content_items.publish_status/embed_on ← website surfacing

EVOLVE (reconcile into the new model):
  artists / contractors / clients / sponsors tables
     → merge into actors + actor_roles (dedupe shared people);
       keep as views during transition, then retire
  events.series (flat)  → events.type (canonical 4-value axis)
  sponsorships.sponsor  → point at actors
```

---

## 6. Migration sequencing (additive, low-risk, introspect-first)

```
PHASE A — Actors core (no data loss)
  1. Create actors + actor_roles.
  2. Backfill from artists/contractors/clients/sponsors, deduping
     shared identities. Show Keith the dedupe plan before running.
  3. Repoint sponsorships.sponsor → actors. Keep old tables as
     views so nothing breaks. Verify.

PHASE B — Event links
  4. Create event_participants. Backfill known lineups where data
     exists (e.g. DI#2's DJs).
  5. Add events.type (migrate from series), keep is_content_event.
  6. Ensure equipment_usage is writable from the Log Event / Money
     panel (it isn't yet).

PHASE C — Content & tags
  7. content_items (+ publish_status/embed_on); tags + taggables.
     Add 'booth-to-wall' tag.

PHASE C2 — Operational core (workflow)
  7a. events.stage + owner_actor_id.
  7b. contracts + files (link to existing storage buckets).
  7c. tasks + task_assignments.
  7d. task_templates per event-type; the equipment/participant-
      driven day-of checklist GENERATOR (runs on assignment change
      or on demand — not static rows).
  7e. budget_lines (planned) + variance view vs. existing actuals.
  7f. touchpoints (CRM-lite).

PHASE D — Measurement & definitions
  8. metric_definitions; seed it from existing kpi_targets keys +
     the derived v_kpi_* metrics.
  9. Build v_data_points (UNION snapshot + derived). Materialize +
     pg_cron refresh if perf needs it.
  10. (Tier 2) formula evaluator for source_type='formula'.

PHASE E — Visualizer
  11. UI reads metric_definitions + v_data_points + kpi_targets.
      Time-series, cross-entity, slice-by-dim, user-defined formulas.

Gates: RLS on every new table via is_admin() (roles are
master_admin/sub_admin — NOT 'admin'). Never blanket-grant anon
(caused the 016/017 regression). Re-assert anon revoke on any
replaced view. Introspect + show SQL before each apply.
```

---

## 7. Open decisions for Keith (resolve before/when implementing)

- **Q1 — Actor dedupe rule:** when merging artists/contractors/clients, what's the match key for "same person"? Email? Name? Manual review of ambiguous ones?
- **Q2 — Guests as actors?** Are ticket buyers / RSVPs full `actors`, or a lighter `guests` table (privacy + volume — thousands of attendees you won't otherwise track)? Lean: keep `guests` separate/light, promote to actor only if they recur (a guest who becomes a sponsor).
- **Q3 — Participant fee ↔ expense:** does a participant's `fee` auto-create an expense row, or are they reconciled? (Auto = less double-entry; reconciled = more control.)
- **Q4 — Attendance vs RSVP** (from prior session): separate fields. Free events have soft RSVP counts; ticketed have hard counts. Confirm how each event type captures headcount.
- **Q5 — Formula tier:** confirm Tier 2 (simple arithmetic composition) as the start, Tier 3 deferred.
- **Q6 — Refresh cadence** for the materialized data-point rollup (nightly? on-write?).
- **Q7 — Task assignees & logins:** team members are actors with auth logins so they can manage their tasks. Confirm: do external actors (a contracted DJ) get a login to see their tasks, or are tasks internal-only for now? (Affects RLS — non-admin actors seeing only their own tasks is a new access pattern beyond master/sub_admin.)
- **Q8 — Day-of checklist generation:** regenerate on every assignment change (always current, may churn) or generate once on demand near the event (stable, manual)? Lean: on-demand "generate day-of checklist" button.
- **Q9 — Contract → actual reconciliation:** when a contract fee is paid, does it auto-create the expense/income row, or do you mark it paid and reconcile manually? (Mirrors Q3 for participant fees.)
- **Q10 — Jennifer bridge:** confirm export/import of Keith's tasks only (not team tasks, not live sync) as the v1. The deeper integration is its own design session.

---

## 8. Handoff notes for Claude Code

- This is a **concept**, not final SQL. Introspect the live prod schema (project yaytdosxfhcqatmhctzk) first and reconcile against Section 5 — correct anything I have wrong.
- **Additive and reversible.** Old tables become views during transition; nothing downstream breaks in one step. No greenfield rebuild.
- **Respect the locked money model** (Section 1.4) — don't redefine net_pl / total_raised / % to mission.
- **RLS + anon discipline** (Section 6 gates) — the same rules that bit us before. Note Q7: actor-scoped task access (an actor seeing only their own tasks) is a NEW access pattern beyond master/sub_admin — design its RLS carefully if enabled.
- **Show the plan + SQL before applying each phase.** Phase A's actor dedupe especially needs Keith's eyes before it runs.
- Sequence respects the project rule that this is Come With (no Come With Fitness anywhere).
- **Scope is large — Keith approved an "all in one build," but build it PHASE BY PHASE with a checkpoint after each, not one giant commit.** Each phase = its own reviewable commit, push held. Phases are ordered so the foundation (actors → events → links) lands before what depends on it (workflow, measurement, visualizer). Do NOT build the visualizer (Phase E) before the measurement layer (Phase D) exists.
- **Verify after each phase:** the 5 financial views still anon-401, and existing dashboard functions still work. The actor merge (Phase A) and any view replacement are the high-risk steps.
- Start by proposing **Phase A only** with its SQL for review. Don't run the whole sequence unattended — the actor dedupe needs Keith's eyes before it executes.
