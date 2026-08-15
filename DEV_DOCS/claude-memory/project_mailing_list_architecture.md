---
name: project-mailing-list-architecture
description: One master mailing list with per-segment tagging — never separate lists per brand. User stated 2026-05-28 at Phase 8 kickoff.
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

Architectural decision: there is ONE master mailing list. Subscribers
are tagged with one or more segments (e.g., "come_with",
"dance_infusion", "dance_infusion_2", "vip"). Status (subscribed /
unsubscribed / bounced / complained) is **per-subscriber, not
per-segment**.

This matches the existing schema (006_mailing_list.sql / 009):
- `public.subscribers` — one row per email, holds status
- `public.subscriber_segments` — many-to-many join (subscriber × segment text)

**Why:** User explicitly stated: "one master mailing list which check
subscribed/unsubscribed for dance infusion, come with, and any other
lists we create." The data-lake goal benefits from unified subscriber
records that cross-reference with clients/guests/ticketing.

**How to apply:**
- ANY new subscribe form (Come With main, Dance Infusion hub, future
  brands/events) writes to the SAME subscribers table with a different
  `segment` tag in subscriber_segments
- Unsubscribe flips `subscribers.status = 'unsubscribed'` globally —
  one click removes the person from ALL segments. There is no
  "unsubscribe from this segment only" UX
- Broadcast sends filter by segment but always exclude
  `status != 'subscribed'`
- Re-subscribing requires a fresh confirm flow (status flips back to
  'pending' → 'subscribed' after click)
- Segments are free-text strings (no enum) — coordinated by convention
  in the frontend. Keep a short list (e.g., `come_with`, `dance_infusion`,
  `vip`) and document new ones as they're added

**Update 2026-07-13 (two-level convention, now in CLAUDE.md and live in
prod):** brand rollups `come_with` / `dance_infusion` are what campaigns
target; per-event segments (`di-02-2026-05`, `come-with-7-11`, …) keep
cohort history. Backfilled: di-* → dance_infusion (83 subscribed),
come-with-* + as-01 → come_with (32 subscribed); 8 people in both.
Every event import must add BOTH segments. Campaign picker sends to one
segment at a time, so brand rollups are what make "email the whole DI
list" a single send. Partiful has NO official API and its CSV export has
no emails — manual CSV stays. Related: [[cw711-import]].
