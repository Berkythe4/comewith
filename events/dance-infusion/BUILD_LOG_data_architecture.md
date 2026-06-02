# Build Log — Come With Data Architecture (A→E)

**Started:** 2026-06-02 · **Spec:** `ComeWith_Data_Architecture_Concept.md` · **Status:** in progress

This is the running decision/assumption/question log for the data-architecture build. Keith
reviews after. **Nothing is applied to prod or staging by this build — it produces reviewable
migration files + UI committed per phase, push held.** Apply happens on Keith's go.

---

## 0. How this build is being run (reconciling the prompt vs. the spec)

- The prompt says "build all phases, no stops, no approval between phases." The spec (Section 8)
  says "phase by phase, **each its own reviewable commit, push held**," "**show SQL before applying —
  don't run the whole sequence unattended; the actor dedupe needs Keith's eyes before it executes**,"
  and "start by proposing Phase A only."
- **Resolution:** I build each phase as migration SQL + UI, **commit per phase (the reviewable
  checkpoint), push held, and DO NOT apply** to prod or staging. So nothing executes against
  money/auth/prod unattended; the held commits + this log are the "show SQL before apply" surface.
  I proceed through phases without pausing for approval (per the prompt), but **apply is gated on
  Keith's review** (per the spec + CLAUDE.md). Dedupe + external-actor RLS especially need his eyes
  before apply.
- **Apply target (ASSUMPTION A1):** migration **files only**, committed locally, push held. Not
  applied to prod (`yaytdosxfhcqatmhctzk`) or staging (`qjosjafbizxdtkisyrqm`). The per-phase
  "verify 5 views anon-401 / dashboard still works" is therefore: prod is **untouched** so the
  invariant holds trivially; I additionally verify each migration *re-asserts* anon (and, for the
  new external tier, `authenticated`) revokes on any sensitive view. **Confirm the apply target.**
- **Standing-priority flag (A2):** `LEARNINGS §5` + `CARRYOVER` + `ROADMAP` say Come With is
  *maintenance-only until the CWF BRD ships* and the event-model redesign is *parked, design-first*.
  This build is that redesign + more. Keith directed it explicitly; noting the conflict, not blocking.

## Keith's confirmations (applied directly, per the prompt)

- **Q1 dedupe:** best-effort match (email then name); don't agonize. Provenance preserved +
  relationship-inspection UI so Keith fixes links/merges by hand after.
- **Q3 / Q9 fees & contracts:** contract/participant fee → **mark paid + reconcile manually**; do
  NOT auto-create `expenses`/`income` actuals.
- **Q5 formula tier:** **Tier 2** (safe arithmetic over metric_keys). Tier 3 deferred.
- **Q6 rollup cadence:** **nightly** materialized refresh via pg_cron (014).
- **Q7 external actors get logins:** YES — they see ONLY their own data. Highest-risk; see §RLS.
- **Q8 day-of checklist:** **on-demand "generate" button**, not auto-churn.

## Open / deferred (not resolved by the prompt — flagged, sensible default taken)

- **Q2 guests-as-actors:** keep `guests` separate/light (per the doc's lean); promote to actor only
  on recurrence. (Default taken.)
- **Q4 attendance vs RSVP:** events already has `total_attendance`; ticketing has `attended` bool.
  Keep both; no change this build. (Default taken.)
- **Q10 Jennifer bridge:** deferred entirely (out of scope this build).

---

## 1. Section-5 reconciliation (introspected from migrations 002–022; live prod drift NOT checked — no prod DB creds here, flag)

**Corrections / clarifications to the doc:**

- **`events`** (007): has `series` (text), `status` ∈ {planning, announced, on_sale, sold_out,
  completed, cancelled}, `total_attendance`, `bar_minimum`, `ticket_url`, `hero_image_path`.
  **Does NOT have** `type`, `stage`, `owner_actor_id`, `is_content_event`, `capacity` (capacity is
  on `venues`). Doc's proposed `status` values (draft/confirmed/…) differ from live — **keep live
  `status`; ADD `stage` + `owner_actor_id` + `type` + `is_content_event` (Phase B/C2).**
- **People tables:** `clients` (007? no → 003: id, **user_id→profiles**, full_name, email, phone,
  company, address, source), `contractors` (003: full_name, stage_name, email, phone, role,
  hourly/day_rate, payment_terms, tax_form_on_file), `artists` (008: stage_name, legal_name, bio,
  genres[], signature_color, rate, contact_email/phone, social_links jsonb, status), `sponsors`
  (007: name, contact_*, website, logo_path). **`clients.user_id` is the existing person↔login
  link** — the actor-login pattern mirrors it.
- **`artist_bookings` already exists** (008: artist×event×role×fee×paid) — this is the de-facto
  `event_participants` for artists. **Phase B backfills `event_participants` from it** (don't
  reinvent). Doc's `event_participants.is_contractor` maps to the lineup/crew split.
- **`sponsorships`** (007): `sponsor_id→sponsors`, event_id, tier, cash_amount, drink/entry_tickets,
  in_kind_value, status. **Phase A adds `actor_id→actors`, backfills, keeps `sponsor_id` during
  transition.**
- **Helpers:** `public.is_admin()` (master+sub) and `public.is_master_admin()` both exist.
  `profiles.role` ∈ {master_admin, sub_admin, customer}.
- **⚠ Grants (013):** blanket `grant all on all tables ... to anon, authenticated, service_role` +
  `alter default privileges ... to anon, authenticated`. **RLS is the real gate on tables.** **Views
  are NOT row-gated** — they rely on grants. `019` revoked the 5 financial views **from `anon` only**.
- **`guests`, `ticketing`, `income`, `expenses`, `third_party_donations`, `equipment_usage`,
  `metric_snapshots`, `kpi_targets`, `content_series`, `dashboard_prefs`, `feedback_log`** — exist
  as the doc says. `equipment_usage` exists but **nothing writes it** (Phase B gap, confirmed).

---

## 2. ⚠ THE ROLE-ACCESS MATRIX — and the external-actor leak (highest-risk)

**The problem (found at introspection):** financial/KPI views are revoked from `anon` but **NOT from
`authenticated`**. External actors log in as `authenticated`. So, naively, an external actor would
read `v_kpi_*` just like an admin. **This must be closed BEFORE external logins are enabled (Phase
C2/Q7).**

**Identity link:** `actors.user_id → profiles.id` (added Phase A, nullable). An external actor's
`auth.uid()` = their `profiles.id` = `actors.user_id`. Their actor_id is found via that link. Tasks
and participation are scoped to that actor_id.

**Target access matrix (to be enforced; negative tests in the checklist):**

| Object | anon | actor (authenticated, non-admin) | sub_admin | master_admin |
|---|---|---|---|---|
| `actors`, `actor_roles` | ✗ | **own row only** (read) | manage | manage |
| `events` | non-cancelled public read | **only events they're a participant on** (limited cols) | manage | manage |
| `event_participants` | ✗ | **own rows only** (read; their role/fee) | manage | manage |
| `tasks`, `task_assignments` | ✗ | **only tasks assigned to them** (read + update status) | manage | manage |
| `contracts` | ✗ | **own contracts only** (read) | manage | manage |
| `sponsorships`, `budget_lines`, `income`, `expenses`, `ticketing`, `third_party_donations` | ✗ | **✗ (never)** | manage | manage |
| financial/KPI **views** (`v_kpi_*`, `v_event_summary`, `v_data_points` if financial) | ✗ (401) | **✗ (must revoke from `authenticated`)** | read | read |
| `dashboard_prefs`, `feedback_log`, `metric_definitions`, admin UIs | ✗ | ✗ | per existing | per existing |

**Mitigation decisions (Phase C2 prerequisites — NOT enabled until done + tested):**
1. **Default-deny everywhere.** Every new table: RLS on, `is_admin()` manage policy; actor-self
   policies added ONLY where the matrix grants self-access, keyed to `actors.user_id = auth.uid()`.
2. **Financial views:** before any external login exists, **revoke SELECT from `authenticated`** on
   all financial views (not just anon), and have the admin dashboard read them via the existing
   admin session (which is authenticated+admin) through **`security_invoker` views over RLS-gated
   base tables** OR a `security definer` RPC that checks `is_admin()`. Chosen: re-issue the financial
   views as `security_invoker = on` so they inherit the caller's base-table RLS (admins pass
   `is_admin()`, actors get zero rows) — plus an explicit `revoke ... from anon, authenticated` and
   admin-only grant. **This is the single most important safety step; it ships in C2 with negative
   tests and is NOT applied until Keith reviews.**
3. **No actor login is provisioned** (no `profiles` row linked to an actor with a non-admin role)
   until the matrix above is implemented and the negative tests pass on staging.

**Negative tests (must pass before external logins go live — in the test checklist):**
- Log in as a non-admin actor → `select * from v_kpi_dashboard` returns **0 rows / denied**.
- → cannot read another actor's `event_participants.fee`, any `sponsorships`, `budget_lines`,
  `income`, `expenses`, or any event's P&L.
- → can read only their own `actors`/`actor_roles`/`event_participants`/`contracts` and tasks
  assigned to them; can update only their own task status.
- anon → all 5 financial views still 401.

---

## 3. Phase-by-phase decisions / deviations

### Phase A — actors core (migration 023) — IN THIS COMMIT
- Create `actors`, `actor_roles`, and `actor_source_links` (provenance: which legacy row each actor
  came from — powers the inspection UI + reversibility).
- `actors.user_id → profiles(id)` nullable (the login link; no actor login enabled yet).
- Backfill actors from `artists`, `contractors`, `clients`, `sponsors`; dedupe by
  `coalesce(lower(email), lower(name))` (best-effort, Q1). One actor per match key; `actor_roles`
  gets {artist|contractor|customer|sponsor} per contributing source. `clients.user_id` carried to
  `actors.user_id`.
- Add `sponsorships.actor_id → actors`; backfill via source links; **keep `sponsor_id`** (transition).
- **DEVIATION from doc Phase A step 3 ("old tables → views"):** NOT converting `artists/contractors/
  clients/sponsors` to views in Phase A. The dashboard still **writes** to those tables (e.g. Add
  Sponsor); replacing them with read-only views would break writes — violating "nothing downstream
  breaks." So Phase A is **purely additive** (actors populated in parallel, provenance kept); the
  view-cutover happens only AFTER the dashboard is repointed to actors, in a later step. More
  conservative + reversible than the literal doc step. **Logged for Keith.**
- RLS: `is_admin()` manage on all three new tables. (Actor-self read policy deferred to C2 with the
  financial-view lockdown, so access isn't opened prematurely.)
- **Relationship-inspection UI:** admin page listing actors → their roles → their event links, with
  edit (reassign role, merge duplicates, fix links). Keith's safety net for the dedupe.

(Phases B, C, C2, D, E logged as they're built.)

---

## 4. Verifications per phase
- **Phase A:** migration is additive (no drops, no view replacement); does not touch financial views;
  re-asserts nothing needed (no sensitive view created). Anon-401 invariant unaffected (prod
  untouched). [to record on each commit]
