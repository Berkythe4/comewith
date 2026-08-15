---
name: project_public_events
description: "Public data-driven events list on comewith.org — migration 030, v_public_events view, dashboard toggles, homepage fetch"
metadata: 
  node_type: memory
  type: project
  originSessionId: 90f54be3-1f54-4bc6-ba07-0bdfee57a183
---

Shipped 2026-06-13. The comewith.org homepage Events panel is now data-driven:
admin toggles an event public in the dashboard and it appears on the site.

- **Migration 030** (`supabase/migrations/030_public_events.sql`, applied to prod via
  Management API): adds `events.is_public` (bool default false) + `events.ticket_label`
  (`ticket_url` already existed from 007). Creates **`v_public_events`** exposing ONLY
  `name, event_date, venue_name, ticket_url, ticket_label` for is_public + future-dated
  (`event_date >= current_date`) + non-cancelled + non-deleted events. Anon granted
  SELECT on that view only.
- **Security fix folded in:** anon could previously read the ENTIRE events table on prod
  (013's blanket grant + a "Public can read non-cancelled events" RLS policy; 019 only
  revoked the financial views). 030 revokes anon on the events table and drops that
  policy. Verified anon → events table now 401, v_public_events 200/[], 5 financial
  views still 401. See [[feedback_no_broad_anon_grants]].
- **Rollback:** `supabase/rollback/030_public_events.down.sql` (kept OUT of migrations/
  so the CLI doesn't parse it as a 2nd forward "030"). DOWN drops view + 2 columns but
  KEEPS the security re-lock (Keith's call). Page rollback = `git revert` the page
  commit + push (Netlify redeploys).
- **Frontend:** dashboard.html Log/Edit Event forms got a "Show on public site" toggle +
  ticket URL + button label; index.html Events panel fetches the view (soonest-first,
  http(s)-only ticket buttons, graceful empty state). Applied via [[feedback_prod_migration_apply]].
