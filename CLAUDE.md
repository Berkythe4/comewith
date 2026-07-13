# Come With — Project Conventions

Operational conventions for this repo. These are binding — follow them exactly.
Broader migration history and architecture live in `ROADMAP.md`; phase/status
notes live in Claude memory.

## Database / Supabase migrations

- **Project:** prod is `yaytdosxfhcqatmhctzk` (`comewith-prod`). The CLI is linked
  to staging; do not assume the link points at prod. Migrations live in
  `supabase/migrations/NNN_name.sql`, numbered in sequence (… 015–019 and up).
- **Introspect before apply.** Confirm live columns / view definitions / policies
  against prod, reconcile any `[VERIFY]` refs, and show the SQL/diff for review
  before applying anything to prod.
- **Roles:** `master_admin` / `sub_admin` / `customer`. There is **no `admin`
  role.** RLS uses the helper `public.is_admin()` (= `role in ('master_admin',
  'sub_admin')`). New admin-only tables: `for all using (public.is_admin())`.
- **NEVER use a blanket `grant ... to anon` in a migration.** Specifically, do not
  write `grant all on all tables in schema public to anon` (or to `authenticated`).
  `013_grants.sql`'s `ALTER DEFAULT PRIVILEGES` already grants the right
  privileges to new tables automatically. A broad grant silently re-grants SELECT
  on **all views too**, re-exposing financial views that were deliberately revoked
  from `anon` — this caused the **016/017 regression that 019 had to fix**. If you
  ever must re-grant, immediately re-assert every prior `revoke … from anon`, and
  verify anon access in the post-apply check (financial views must return 401).
- **Financial views are anon-revoked by design** (decision E1): `v_event_summary`,
  `v_kpi_event_financials`, `v_kpi_parties`, `v_kpi_dance_infusion`,
  `v_kpi_dashboard`. Keep them revoked. Verify with an anon REST GET → expect 401.
- **Apply discipline:** apply additively, verify on prod (objects, RLS has a real
  policy — never RLS-enabled-with-no-policy, admin can read/write, anon blocked),
  then commit the migration file so tracked history matches prod.

## Series contract (events.series)

`events.series` is free text. KPI views match it **exactly**. The Log Event form
MUST write `series = 'Come With Parties'` for parties and `series = 'Dance Infusion'`
for DI events, or those KPIs read empty. `'Come With Production'` is services, not
parties.

## Mailing segments (brand delineation)

Two-level segments on `subscriber_segments`, established 2026-07-13:
- **Brand rollups** (what campaigns target): `come_with`, `dance_infusion`.
  A subscriber can hold both. Unsubscribe stays **global** (one master list).
- **Per-event segments** (cohort history): the event slug or event code,
  e.g. `come-with-7-11`, `di-02-2026-05`.

Every event import MUST add BOTH the event segment AND the matching brand
segment. Public signup widgets pass the brand segment (`come_with` on the
homepage; DI pages must pass `dance_infusion`). Never re-subscribe an
unsubscribed email during an import (e.g. `chaddercheesy@gmail.com`).

## Scope

- This codebase is **Come With only**. Do **not** add anything Come With Fitness
  (CWF) anywhere — not in the dashboard, schema, or pages.
