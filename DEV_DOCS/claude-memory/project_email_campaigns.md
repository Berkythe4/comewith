---
name: project_email_campaigns
description: Email-blast system (Resend + campaigns) is fully deployed; how it works and the one external step (verify comewith.org in Resend)
metadata: 
  node_type: memory
  type: project
  originSessionId: 23f44bb5-a672-44a4-9c2e-b8eac9975d80
---

**Email/campaign stack is LIVE on prod (evaluated 2026-06-24, commit 8d56a09).** Edge functions deployed+ACTIVE: `subscribe`, `confirm-subscription`, `unsubscribe`, `send-campaign`, `resend-webhook`. Secrets all set: `RESEND_API_KEY`, `RESEND_WEBHOOK_SECRET`, `SITE_URL`, plus the Supabase ones. From = `Come With <berky@comewith.org>`.

**Flow:** Campaigns tab → "+ Draft campaign" writes `mailing_campaigns` (status=draft). Send button invokes `send-campaign` → filters `subscribers` (status=subscribed) by `segment_filter` (joins `subscriber_segments`), sends individually via Resend with per-recipient unsubscribe links, logs a `mailing_events` row per send, sets campaign status sent/recipient_count/sent_at. `resend-webhook` (svix-signed) ingests delivered/bounced/opened/clicked and flips subscriber status on bounce/complaint. Dashboard adds: segment "Send to" dropdown with live subscribed counts, a Preview (sandboxed iframe), a send-confirm showing audience size, a **Test** button (send-campaign accepts `{campaign_id, test_email}` → single [TEST] email, no list send/status change/logging; deployed to prod 2026-06-24 via `supabase functions deploy send-campaign --project-ref yaytdosxfhcqatmhctzk` with SUPABASE_ACCESS_TOKEN=SBP_PAT — Docker NOT needed), and a **Stats** button aggregating `mailing_events` per campaign (Sent/Delivered/Opened/Clicked/Bounced/Complained/Failed, unique people + rates). Bounces/complaints auto-flip subscriber status via resend-webhook; unsubscribes via the per-recipient link → `unsubscribe` fn. Both auto-exclude from future sends (send only targets status=subscribed).

**Data (2026-06-24):** 86 subscribed / 11 unsubscribed (97 total). Segments: `di-02-2026-05` (68), `di-01-2025-09` (28), `as-01-2026-01` (3). 0 campaigns sent yet.

**The ONE external setup step that can't be done from code — verify before first blast:** comewith.org must be a **verified sending domain in the Resend dashboard** (SPF/DKIM/DMARC DNS records) or mail from berky@comewith.org silently fails / spam-folders. Also confirm the Resend **webhook endpoint** points to the `resend-webhook` function URL with signing secret = `RESEND_WEBHOOK_SECRET` (secret is set, but the endpoint registration lives in Resend). Always Preview + send a test to yourself first. NOTE: send-campaign has no built-in "test send to one address" — adding that needs an Edge function change + redeploy (CLI is linked to staging; prod functions deploy separately).

**Lesson:** after DROPping a table (e.g. 047 dropped `sponsors`), grep the dashboard for PostgREST embeds of it (`sponsor:sponsors(...)`) — a dropped-table embed 400s the whole query and silently empties the tab. This bit the Sponsorships/Money/Customers panels. See [[project_actor_model_and_equipment]].
