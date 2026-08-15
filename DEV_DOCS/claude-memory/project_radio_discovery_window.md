---
name: project-radio-discovery-window
description: "Radio window filtered on ra_artists.next_event_date (their SOONEST show) and hid 77 artists in the 8/18+4wk window; fix indexes ra_events.lineup — on branch radio/window-by-lineup, unmerged/undeployed as of 2026-08-15"
metadata: 
  node_type: memory
  type: project
  originSessionId: 6264aafb-c0a6-4e5c-bdca-a6f2c2a39988
  modified: 2026-08-15T20:54:08.357Z
---

Audited 2026-08-15 against prod for the **2026-08-18 + 4 weeks** window.

`ra_artists` holds one row per artist carrying `next_event_date` — their **soonest**
show, i.e. a summary of the pull, not a fact about the artist. The radio window
filtered on it, so anyone playing just *before* the start date **and again** inside it
fell out: **77 artists, 70 with a SoundCloud link.** Fix builds the window from every
date in `ra_events.lineup` and re-points each artist at the show they play *inside*
it; one shared `raWindowPool()` now feeds the list, counts, venue filter and all four
scan/match passes. Simulated on prod: 866 → 943 (exactly the 77).

Same audit, same family of silent shortfall:
- **DICE** detail-fetched the first 160 candidates in tag order and saved 159 — cap
  binding exactly, so a 90-day request returned **7 days** (weeks 2–4 had none).
- **Ticketmaster** `city` is a literal match: "New York" = Manhattan only, nothing in
  Brooklyn/Queens.
- "↻ Pull shows" swallowed both sources' failures, so an outage looked like a zero.

**Status: MERGED + DEPLOYED 2026-08-15** (`88b2153`). Dashboard live on Netlify;
`pull-dice` v5 + `pull-ticketmaster` v7 live on prod, both verified by reading the
deployed source back. **Not yet exercised** — no pull has run through it, because
invoking the pulls needs an admin JWT or the service-role key and neither is on the
desktop. One "↻ Pull shows" proves DICE past week 1, TM in Brooklyn, and the 77 back.

**Also found, NOT fixed:** `dj-station` caps its artist query at `.limit(160)` with no
notice against an ~879-artist window; the scan cache never re-reads a cached profile;
**no cron pulls the market at all** — the pool is only as fresh as the last manual
click. Full write-up: LEARNINGS §15 + `reviews/session_2026-08-15.md`.
Related: [[project-two-machine-handoff]], [[project-radio-episode-planning]].
