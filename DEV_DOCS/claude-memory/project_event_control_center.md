---
name: project-event-control-center
description: "Event Control Center (control.html) — 222-223 built it 2026-10-29; 224-228 (2026-11-04/05) loaded DI #3 real data and REBUILT it (Where we stand, Chase list, RSVPs, grouped guest list); one eventRoster() shared with the hub Customers tab; render harness exists"
metadata:
  node_type: memory
  type: project
  originSessionId: fab1769f-05ec-4f5c-8644-298732a09006
  modified: 2026-10-29T20:23:46.489Z
---

control.html + assets/control-center.js, built 2026-10-29 from a claude.ai prototype brief. Keith said "just do it, don't stop", so Phase 1 approval was skipped.
222 (schema, additive) and 223 (DI #3 seed: 2 tiers, 3 Champions pending, 3 named Savery comps, 5 blocker tasks, 4 planned costs) are applied to prod.

**Why:** DI #3 is on 2026-11-14 at Signal, and DI #2 ran on spreadsheets plus dance_infusion.json.

**How to apply:** it has only been verified headless, against an in-memory fake loaded with prod rows. The first real login, real check-in and real import on prod are still to do.
Open questions for Keith:
- the drink-ticket $ value toward the bar minimum
- whether sponsor comps credit $25 each to MS (`comp_mission_value`; DI #2 did)
- DI #3's `doors_time` reads 15:00, copied from DI #1
- DI #1's date reads 2025-09-08, not Sep 6
- a duplicate "Dance Infusion MS" event row

See [[project-event-hub-and-di2-money]], [[project-event-import-tool]].

**Update 2026-11-05 (224-228):** DI #3 now runs on real data: Zeffy, Partiful RSVPs (`event_rsvps`), sponsor fixes, Signal deposit, Keith's $1,500 contingent advance (`owner_advances`), lineup + Henry as producer, and every comp under its sponsor. Keith found the old page underwhelming, so it was rebuilt: hero, Where we stand, Chase list, Run of show, grouped Guest list and an RSVPs tab. `eventRoster()` in control-center.js is now the ONLY roster logic, used by both control.html and the hub Customers tab. To see the page locally, run `scripts/control_harness/render.py` (headless Edge can't go under ~477px, so the phone shot uses an iframe). Keith has not yet clicked through it on prod. Open items: JAS, open comps, Signal balance date, drink-ticket value. See [[project-event-hub-and-di2-money]].
