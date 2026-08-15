---
name: project-phase-9-status
description: "Phase 9 (Resend broadcasts + webhooks) closed 2026-05-28. send-campaign + resend-webhook Edge Functions + Campaigns admin tab. Direct send from Supabase, one source of truth."
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

Phase 9 closed 2026-05-28. Four sprint commits.

## What ships

### Edge Functions
- `send-campaign` (admin JWT required). Takes {campaign_id},
  filters subscribers by segment_filter + status=subscribed,
  sends individually via Resend with per-recipient unsubscribe
  link, logs `mailing_events` rows with event_type='sent' +
  resend_event_id. Updates campaign status sending → sent
  (or failed) with recipient_count + sent_at.
- `resend-webhook` (public, --no-verify-jwt). Receives Resend
  delivery events. Optional svix signature verification when
  RESEND_WEBHOOK_SECRET is set. Strips "email." prefix from
  event type, attributes to original sent event via
  resend_event_id, inserts mailing_events. For bounced or
  complained, also flips subscribers.status so future sends
  skip them.

### Dashboard
- Campaigns admin tab in dashboard-v2 listing all campaigns
  with status badges and recipient_count
- "+ Draft campaign" modal: name / subject / preview_text /
  segment_filter / body_html
- Send button on each draft row with a confirm() guard.
  Toast shows "Sent to X of Y subscribers" on success.

## Architecture choice (per user pick at scoping)
Direct send from Supabase. NOT Resend Audiences sync. Means:
- Subscribers live only in Supabase (master_list_architecture)
- send-campaign iterates them and sends individually
- No syncing logic to maintain between Supabase and Resend

Tradeoff: can't use Resend's broadcast UI / analytics. The
mailing_events table + Campaigns admin tab are the analytics
surface instead.

## Required out-of-band setup (user task)
The Resend webhook endpoint exists but Resend doesn't know to
POST to it yet. To wire up:

1. Resend dashboard → Webhooks → Add endpoint
2. URL: `https://qjosjafbizxdtkisyrqm.supabase.co/functions/v1/resend-webhook`
3. Select events: `email.delivered`, `email.bounced`, `email.complained`
   (optionally `email.opened`, `email.clicked` if you want engagement)
4. Save → Resend shows you the signing secret (whsec_…)
5. From PowerShell: `supabase secrets set RESEND_WEBHOOK_SECRET=whsec_xxx`

Until step 5, the function accepts unsigned requests (logs an
acceptance) — so it'll start working as soon as step 1-4 are done,
just without signature verification. Production cutover MUST do step 5.

## Hard-coded values to revisit in Phase 11
- send-campaign UNSUB_BASE = `http://localhost:8765/unsubscribe.html`
- Same TODO pattern as send-agreement / subscribe.

## Open for Phase 10
- pg_cron schedules for:
  - Nightly materialized-view refresh (mv_cross_event_kpis,
    mv_repeat_sponsors, mv_top_artists)
  - Scheduled campaign sends (mailing_campaigns.scheduled_for
    column already exists; need a cron job that picks up
    due rows and invokes send-campaign for each)
  - Audit log retention (delete rows older than N days)

## Time tracking
Estimated ~10 min wall-clock, came in around ~10 min. Pattern
is well-trodden by now.

Related: [[project-phase-8-status]], [[project-mailing-list-architecture]]
