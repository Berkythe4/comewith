---
name: project-phase-10-status
description: Phase 10 (pg_cron automation) closed 2026-05-28. Two cron jobs running on staging — nightly MV refresh + 365-day audit log retention. Scheduled campaign sends deferred pending cron→Edge-Function auth design.
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

Phase 10 closed 2026-05-28. Migration `014_cron.sql` adds two
schedules via pg_cron (enabled since Phase 0). Two sprint commits.

## What ships

| Job | Schedule | What |
|---|---|---|
| `refresh-materialized-views` | 03:00 UTC daily | REFRESH MATERIALIZED VIEW CONCURRENTLY for mv_cross_event_kpis, mv_repeat_sponsors, mv_top_artists |
| `audit-log-retention` | 04:00 UTC daily | DELETE audit_log rows older than 365 days |

Both registered in `public.automation_jobs` for dashboard visibility.

## What did NOT ship (deferred)

### Scheduled campaign sends
The `mailing_campaigns.scheduled_for` column already exists. A
cron job picking up due rows and invoking send-campaign would
let the admin queue sends ahead of time.

Blocker: `send-campaign` verifies an admin JWT. pg_cron + pg_net
can POST to the function URL, but constructing or fetching the
admin JWT in SQL is awkward. Two clean paths forward (pick one
when wiring this):

1. **Cron-secret header on send-campaign.** Add an alternate auth
   path: if request has `X-Cron-Secret: <env>` and it matches
   `CRON_SECRET` Edge Function secret, skip the JWT check. Then
   cron's pg_net call includes the header.
2. **Vault + service_role JWT.** Store the project's service_role
   key in `vault.secrets`, read it in the cron SQL, include it in
   the pg_net Authorization header. send-campaign already trusts
   service_role implicitly.

Either is ~15 min of work. Punted because (a) Berky can click
Send manually for the foreseeable future, and (b) adding the
auth code path needs a security review.

### Audit log archival to Storage
The retention job DELETEs rows. A more conservative version
would copy old rows to a Storage bucket as JSON before delete.
Worth doing if you ever need to satisfy a compliance request
that pre-dates 365 days. Skipped for now.

### Automation runs tracking
Schema has `automation_runs` (010_automation_audit_photos.sql)
intended to log each cron execution. pg_cron natively logs to
`cron.job_run_details`, so right now nothing writes to
automation_runs. A small Edge Function or pg_cron-callable
SQL fn could mirror executions into it. Phase 11+ polish.

## Hard-coded values to revisit in Phase 11
- None directly. The cron jobs are timezone-agnostic (UTC) and
  reference the same DB they run in.

## Open for Phase 11
- Full prod migration sequence — run 001-014 on `comewith-prod`
- Replace all `http://localhost:8765` in Edge Function code +
  Auth Site URL with `https://comewith.org`
- Cut over comewith.org DNS / Netlify to serve dashboard-v2,
  index-v2, customer_portal-v2, sign.html, confirm.html,
  unsubscribe.html, events/dance-infusion-2/index-v2.html
  (latter still blocked by the DI hub publish gate)
- Disable Apps Script triggers and archive the existing
  `appsscript.gs` flow

## Time tracking
Estimated ~5 min wall-clock, came in around 4 min. Just a
migration + register + roadmap update.

Related: [[project-phase-9-status]]
