---
name: project_resend_engagement_tracking
description: "Resend open/click tracking gotcha + how DI#2 campaign opens were recovered from a CSV export; webhook must subscribe to email.opened/clicked AND domain tracking must be ON"
metadata: 
  node_type: memory
  type: project
  originSessionId: 6ae7a3fb-5f8f-4348-8f22-cc711985cf02
---

The DI#2 campaign (sent 2026-06-29, campaign id `92cde9c5-0734-4e8c-b060-83c3e7d8c990`)
showed **0 opens/clicks** in the dashboard even though people opened it. Root cause was
**two-part, both on the Resend side, not our DB**:

1. **Resend open/click TRACKING is opt-in per sending domain and off by default.** Toggles
   live at Resend → Domains → click `comewith.org` → Configuration (NOT under "Settings").
2. **The webhook must be SUBSCRIBED to `email.opened` + `email.clicked`.** Ours was only
   getting delivered/bounced. Fixed 2026-07-07 by subscribing the endpoint to **all events**.
   Endpoint: `https://yaytdosxfhcqatmhctzk.supabase.co/functions/v1/resend-webhook`.

Webhook plumbing itself was always fine — `resend-webhook` inserts ANY event_type into
`mailing_events` and passes svix signature verification (RESEND_WEBHOOK_SECRET set). Both
`delivered` and `bounced` flowed in correctly. The `campaignStats()` hint in dashboard.html
was reworded to distinguish "tracking off" (delivered arriving, 0 opens+clicks = `noEngage`)
from "webhook down" (nothing arriving).

**Backfill method (works without the API key):** Resend dashboard → export sent emails to CSV
(cols: `id`, `last_event`, `to`, `sent_at`, …). `last_event` = furthest state reached
(delivered→opened→clicked). Join CSV `id` → `mailing_events.resend_event_id`. last_event
`opened`/`clicked` → opened; `clicked` → clicked. Backfilled into `mailing_events` tagged
`metadata.source = resend_export_backfill_20260707` (timestamps approximated to sent_at — export
lacks event timestamps; delete by that tag to revert).

**DI#2 TRUE numbers (88 emails = 86 subscribers + 2 CC): 87 delivered, 38 opened (~44%),
2 clicked, 1 bounced.** Among subscribers alone: 36 opened, **0 clicked**. The only 2 clicks were
the CC'd National MS Society partners (`dana.miele@` + `rich.klein@nmss.org`).

**CC TRACKING GAP (found + fixed 2026-07-07, commit ad1c2ea):** `send-campaign` did NOT log CC
sends to `mailing_events`, so CC engagement was invisible (that's why the 2 NMSS clicks didn't
show). Fixed: CC sends now insert a `sent` row (subscriber_id null, `metadata.email`) so the
webhook can attribute their events. `campaignStats()` now counts uniqueness by subscriber_id OR
email (handles `metadata.email` from our send + `metadata.to` from the webhook). Survey CTA in
`send-campaign` (`surveyCta`) was also enlarged.

**Click count CONFIRMED = 2 for DI#2** (both NMSS partners). Resolved 2026-07-07 via Resend
dashboard exports filtered by event: the clicked-filter export shows only 2 DI#2 rows; the
"more than 2" Keith saw was the ACCOUNT-WIDE clicked total (14 emails in the window — the other
12 are agreements/subscription-confirms/artist emails/the test). NOTE: Resend's log filters +
exports are all `last_event`-based (each email in exactly one bucket), so a click hidden by a
later re-open can't be surfaced by any export/API — only Resend's raw aggregate metric would show
it. Going forward the webhook captures every `email.clicked` so this is moot. Webhook chain
VERIFIED live 2026-07-07 (test send → sent/delivered/opened/clicked all landed in real time).

**0 survey responses is REAL**, not a broken link — survey path verified live (survey.html 200,
survey-get returns the DI#2 survey; a live test submission wrote 6 answers, since cleaned up).
Subscribers opened but didn't click through; even the 2 NMSS clicks didn't submit. CTA/conversion
problem. **Certainty locked in:** migration **075 applied to prod** adds an audit trigger on
`survey_responses` (INSERT/UPDATE/DELETE → audit_log, master-only) — before this, survey tables
weren't audited so an insert-then-delete would've been untraceable. `survey_responses` is back to
0 rows (clean). See [[project_survey_system]] and [[project_email_campaigns]].

Retention note: Resend keeps email data ~1d (Free)/3d (Pro) historically, reportedly ~30d now.
The CSV export is the durable record — always export before the window lapses.
