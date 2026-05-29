# Come With — KPI Dashboard Build: Handoff Brief

**For:** the chat that knows the Come With website/dashboard internals, so it can write precise Claude Code prompts.
**Goal:** add a Strategy/KPI section to the admin dashboard, backed by real Supabase data, that updates as data is entered. Keith and Liz will **only ever use the dashboard** — they never open Supabase. All schema and data work is Claude Code's job.

---

## 1. Project facts

- **Supabase project:** `yaytdosxfhcqatmhctzk`
- **Repo:** GitHub `Berkythe4/comewith`, branch `master`
- **Hosting:** Netlify (auto-deploy from `master`)
- **Migration status:** v2 promoted canonical on prod (May 28). Apply new migrations through Claude Code / Supabase SQL editor.
- **Hard rule:** Keith & Liz operate entirely from the dashboard. Every read *and every write* (logging data, editing targets) must have a dashboard UI. No Supabase table-editor usage by them.

---

## 2. What already exists (relevant tables)

- `events` — id, slug, name, **series**, event_date, doors_time, end_time, venue_id, status, bar_minimum, ticket_url, **total_attendance**, … (we are adding `capacity`)
- `income` — date, event_id, **amount**, category, payment_method, deleted_at, …
- `ticketing` — event_id, guest_id, ticket_type, **amount_paid**, …
- `third_party_donations` — event_id, donor_name, **amount**, date, …
- `expenses`, `sponsorships`, `sponsors`, `guests`, `equipment_usage` — present, columns not yet confirmed
- `subscribers` / `subscriber_segments` — email list (status, confirmed_at, unsubscribed_at)
- Existing views: `v_event_summary`, `v_equipment_roi`, `v_artist_history`, `v_sponsor_history`, `v_mailing_list_health`
- `events.series` is the field that separates **Dance Infusion** vs **Come With Parties** vs **Artist Showcase**.

`agreements` carries Come With services/rental revenue (`subtotal`, `total_amount`, `deposit_amount`) — useful later, not required for the four KPI workstreams.

---

## 3. What we're adding (`comewith_kpi_layer.sql`)

Additive only — no existing table is altered except an additive `events.capacity` column.

| Object | Type | Purpose |
|---|---|---|
| `events.capacity` | column | sell-through % denominator |
| `content_series` | table | YouTube series identity (Backyard Sessions, etc.) |
| `metric_snapshots` | table | periodic readings: IG followers, YT subs/watch time/views, per-series views |
| `kpi_targets` | table | **editable targets**, versioned by `effective_date`, `comparison` = gte/lte |
| `v_kpi_targets_current` | view | latest active target per metric |
| `v_metric_latest` / `v_metric_prior` | views | current + prior reading (trend arrows) |
| `v_kpi_event_financials` | view | per-event rollup: tickets, revenue, income, expenses, donations, sponsor value |
| `v_kpi_parties` | view | sell-through %, net P&L per party |
| `v_kpi_dance_infusion` | view | total raised, cost-to-raise-a-dollar, attendance |
| `v_kpi_dashboard` | view | unified feed for snapshot metrics: current, prior, target, comparison, unit |

Seed data: the 12 starting targets + 3 content series.

---

## 4. The data contract (what the dashboard reads)

**Snapshot-style metrics** (followers, subs, watch time, engagement) → read `v_kpi_dashboard`:

```
metric_key, workstream, label, current_value, prior_value,
target_value, comparison, unit, as_of
```

Render per card:
- value = `current_value`
- trend = sign of (`current_value` − `prior_value`), flipped when `comparison = 'lte'` (down is good for cost-to-raise)
- progress bar = `current_value / target_value` (invert for `lte`)

**Event-derived metrics** (sell-through, net P&L, cost-to-raise, attendance) → read the per-event views (`v_kpi_parties`, `v_kpi_dance_infusion`) for the most recent event, and compare to the previous event for "growth per event."

Colour mapping (already used in the visual): content = amber `#E0A24E`, parties = coral `#D9744A`, audience = teal `#5DCAA5`, dance_infusion = purple `#B9A4D9`.

---

## 5. Dashboard work needed

The **visual** is already designed — see `ComeWith_Strategy_Dashboard.html` (flywheel + KPI cards + per-series rows, dark Come With theme). Two build tasks:

**A. Read-only KPI section**
Wire the existing visual to live data: card values, trends, and progress bars come from `v_kpi_dashboard` + the per-event views via the Supabase JS client. Flywheel stays static (it's the mental model).

**B. Entry forms** (because they never touch Supabase)
1. **Log Event** → writes `events` (+ `capacity`), then `ticketing` / `income` / `expenses` / `third_party_donations` / `sponsorships` rows for that event.
2. **Log Numbers** (weekly) → upserts `metric_snapshots` (metric_key, value, captured_on, optional series_id). Pre-list the metric_keys so it's a dropdown, not free text.
3. **Edit Target** → inserts a new `kpi_targets` row (new `effective_date`, `active = true`) — never updates in place, so history is kept. Progress bars recompute automatically.

---

## 6. VERIFY checklist for Claude Code (introspect live schema first)

The SQL has `[VERIFY]` flags where it references things not yet confirmed. Before applying, run and reconcile:

```sql
-- exact series values to fix the WHERE clauses in v_kpi_parties / v_kpi_dance_infusion
select distinct series from public.events;

-- expenses + sponsorships columns used by v_kpi_event_financials
select table_name, column_name, data_type from information_schema.columns
where table_schema='public' and table_name in ('expenses','sponsorships')
order by table_name, ordinal_position;

-- compare to existing rollup so we extend rather than duplicate
select definition from pg_views where viewname='v_event_summary';

-- confirm the admin RLS convention used elsewhere
select definition from pg_policies where schemaname='public' limit 5;
```

Then: fix `series` matching, confirm `expenses.event_id/amount/deleted_at` and `sponsorships.event_id/amount`, decide reuse-or-extend `v_event_summary`, and fill in the RLS policy bodies to match the project's admin check.

---

## 7. DI #1 baseline to gather (so "growth per event" has a starting point)

Numbers Keith provides; Claude Code inserts them against the DI #1 event:
- capacity + total attendance
- tickets sold + ticket revenue
- total raised (donations + sponsor value)
- total event expenses
- Instagram followers and YouTube subscribers as of DI #1 date (→ `metric_snapshots` baseline rows)

---

## 8. Suggested Claude Code prompts

1. *"Introspect the public schema for `expenses`, `sponsorships`, and the `v_event_summary` definition, and list the distinct values of `events.series`. Then open `comewith_kpi_layer.sql` and fix every `[VERIFY]` reference to match the real columns and series values. Show me the diff before applying."*
2. *"Apply `comewith_kpi_layer.sql` as a new migration to project `yaytdosxfhcqatmhctzk`, matching the RLS admin-policy convention already used on existing tables. Confirm all objects created."*
3. *"In the admin dashboard, add a Strategy section that reads `v_kpi_dashboard` and the per-event views, rendering the cards in `ComeWith_Strategy_Dashboard.html` with live values, trend arrows (invert for `comparison='lte'`), and progress bars."*
4. *"Build three dashboard forms — Log Event, Log Numbers, Edit Target — writing to the operational tables, `metric_snapshots`, and `kpi_targets` respectively. Edit Target must insert a new versioned row, never update in place."*
5. *"Backfill DI #1 from these numbers: [Keith's list]. Insert ticketing/donation/expense rows against the DI #1 event and two baseline `metric_snapshots` rows."*

---

## 9. Design principles to preserve

- Targets are **always** editable from the UI, versioned by date — never hard-coded.
- Snapshot metrics live in **one** `metric_snapshots` table (extensible to API sources later via the `source` column — integrations are a later phase, not now).
- Event KPIs are **views** over existing operational tables — no double entry.
- Keith & Liz see only the dashboard. If a number needs entering, there is a form for it.
