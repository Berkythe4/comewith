---
name: project-radio-discovery-window
description: "Radio window filtered on ra_artists.next_event_date (SOONEST show) and hid 77 artists; rebuilt on ra_events.lineup — MERGED + DEPLOYED 2026-08-15. Follow-up same day: the dashboard was also silently truncating at PostgREST max_rows=1000, dj-station had the same bug plus limit(160), and the window can now START in the future"
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

**Status: MERGED + DEPLOYED + EXERCISED 2026-08-15.** Keith ran a real "↻ Pull shows"
at 17:13: Ticketmaster 25 → 40 events across 10 → 16 venues with Brooklyn present, DICE
124 → 240 reaching 9/07 instead of 8/21. Confirmed fixed.

**Follow-up session the same day (laptop) — the same failure class, three more places:**
- **PostgREST `max_rows` is 1000 on prod and truncates SILENTLY.** Every radio load was
  past it (1,327 future events / 1,594 future artists / 1,956 cache rows) and none was
  ordered, so *which* thousand arrived was arbitrary. Worse the further out you looked,
  and it made scanned artists read as unscanned. All page through `sbAll(build, pk)` now,
  ordered by PRIMARY KEY so a tie can't straddle a page boundary. **LEARNINGS §18.**
- **`dj-station` had the identical next_event_date bug** plus the silent `.limit(160)`.
  Rebuilt on `ra_events.lineup` mirroring `raWindowPool()`; paged, and reports
  `scope.capped` + `pool_total` if the 1,500 safety stop binds. **v9 on prod.**
- **The window can now START in the future** — `dj_search_params.start` (date field on
  the episode form), and all three pulls take `from`/`to`, clamp at 180 days not 90/120,
  and echo the window they pulled. DICE's detail budget is aimed at the radio window
  rather than a flat 90 days, scales with it, ceiling 400 → 600. `pull-dice` v6,
  `pull-ticketmaster` v8, `pull-ra-market` v15.
- **Caught before deploy:** all three pulls delete their own source's rows bounded only
  at the bottom. Aiming a pull at a 4-week window would have deleted every TM show beyond
  it. Bounded at both ends now.
- **Not compile-checked** — the laptop has no deno/node. Verified by bundle read-back on
  prod; nobody has clicked it in a browser yet.

**Still NOT fixed:** the scan cache never re-reads a cached profile; **no cron pulls the
market at all** — the pool is only as fresh as the last manual click.
Full write-up: LEARNINGS §15 + §18, `reviews/session_2026-08-15.md` and
`reviews/session_2026-08-15-radio-future-window.md`.
Related: [[project-two-machine-handoff]], [[project-radio-episode-planning]].
