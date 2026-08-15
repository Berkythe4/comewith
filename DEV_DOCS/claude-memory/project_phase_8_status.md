---
name: project-phase-8-status
description: "Phase 8 (mailing list) closed 2026-05-28. Subscribe widget + 3 Edge Functions + confirm/unsubscribe pages + Subscribers admin tab. Master-list architecture (per-segment tagging, global status)."
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

Phase 8 closed 2026-05-28. Four sprint commits.

## What ships

### Edge Functions
- `subscribe` — anon. Upserts subscriber + adds segment + sends
  double-opt-in confirm email (only for new/re-sub flows)
- `confirm-subscription` — anon. Token → status=subscribed +
  confirmed_at. Handles already-confirmed gracefully
- `unsubscribe` — anon. Token → status=unsubscribed +
  unsubscribed_at. Global per master-list architecture

### Frontend
- Subscribe widget on `index-v2.html` (top card, segment='come_with')
- `confirm.html` — auto-confirms on load (it's a "click from email" page)
- `unsubscribe.html` — requires button click (defends against link-scanner
  auto-unsubs)
- Subscribers admin tab in dashboard-v2 with email/name/status/segments/
  source/confirmed_at columns

### Schema
- Reuses existing `subscribers` + `subscriber_segments` tables from
  009_mailing_list.sql. No new migrations.

## Architectural notes
The master-list architecture (per [[project-mailing-list-architecture]])
is enforced by the function logic, not the schema:
- subscribe() picks the existing subscriber row by email if any; new
  segments just add subscriber_segments rows
- unsubscribe() flips status globally — there's no per-segment unsub UX
- Re-subscribing after unsubscribed flips back to pending + fresh confirm

## What did NOT ship (deferred)
- Subscribe widget on DI hub page (the hub is also gated until
  user approves; see [[project-di-hub-publish-gate]])
- Subscribe widget on other future pages — each just calls subscribe
  with its own segment
- Per-segment broadcast send (that's Phase 9 — Resend broadcasts)
- Resend webhook handler for delivery/bounce/complaint events
  (also Phase 9)

## Hard-coded values to revisit in Phase 11
- subscribe Edge Function CONFIRM_BASE = `http://localhost:8765/confirm.html`
- Same TODO pattern as send-agreement (SIGN_BASE_URL)

## Open for Phase 9
- Resend Audiences sync — keep Supabase subscribers + Resend Audiences
  in sync so Resend's broadcast tools work
- mailing_campaigns admin UI in dashboard (draft / schedule / send)
- Resend webhook → mailing_events table for delivered/bounced/complained
- Segment-filtered broadcast sends

## Time tracking
Estimated ~10 min wall-clock, came in around 8 min. The 4 sprints
each followed the established Edge-Function + supabase-js pattern
from Phases 5-7, so very little new ground.

Related: [[project-mailing-list-architecture]],
[[project-phase-7-status]], [[feedback-time-estimates]]
