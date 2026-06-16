# Event Hub — Live-Drive Checklist (Sprint 1)

Walk this in the browser on `dashboard.html` (signed in as a master_admin/sub_admin against
prod `yaytdosxfhcqatmhctzk`). The data layer is already proven by
`tests/event_hub_datalayer_test.sql` (green, rolled back — see below); this is the **UI**
verification. Each line is one observable behavior.

## Open & shell
- [ ] **Events tab → an event row now has an "Open" button** (red, first action) alongside Money / Edit / Delete.
- [ ] Click **Open** → the main area swaps to the **Event Hub**; the page title becomes the event name; **Events** stays highlighted in the sidebar; a **"← Back to events"** link sits at the top.
- [ ] **"← Back to events"** returns to the Events table.
- [ ] Hub header shows: event **name** (Bebas Neue), and facts row — **Date · Type · Series · Venue · Status (badge) · Owner**.
- [ ] Section sub-nav reads **Overview · People · Tasks · Money · Equipment · Contracts · Files**, each with a count chip; clicking switches sections.
- [ ] Visual identity matches the dashboard (cream bg + halftone dots, red accents, Bebas/Inconsolata). The 14 existing tabs look and work exactly as before.

## Stage stepper
- [ ] Header shows a 6-step stepper: **idea → planning → confirmed → live → wrapped → reported**, current stage in red, prior steps filled.
- [ ] Click a different stage → toast "Stage → …", the stepper updates, and it's distinct from the public **Status** badge (status unchanged).

## Overview
- [ ] Summary cards show **Gross revenue, Net P&L (red if negative, green if positive), Ticket revenue, Tickets sold, Capacity, Sell-through %, Attendance, Expenses**; a **Total raised** card appears only for `dance_infusion` events.
- [ ] A prominent **"Day-of task generator"** callout with explanation + **Generate day-of tasks** button. If no people/equipment assigned, a gentle amber hint appears but the button still works.
- [ ] An **Open Money panel** button opens the existing per-event Money modal.

## People
- [ ] **+ Add participant** → modal. "Pick an existing actor" lists actors (with their roles); switching to "Create a new actor" reveals name+kind.
- [ ] Add one with a role (e.g. `dj`) and a **Fee** → it appears in the list with the fee shown.
- [ ] On a participant with a fee, **"Fee → expense"** asks to confirm, then creates an expense (check the Expenses tab / Money panel) — **one-time, manual** (run-once warning shown).
- [ ] **Remove** deletes the participant row (not the actor).

## Tasks
- [ ] **+ Add task** → modal with title/description/priority/due/effort/reward + optional assignee & role (owner/doer/reviewer). Saving shows the task with its assignee.
- [ ] Inline **status dropdown** (todo/doing/blocked/done) persists on change (toast).
- [ ] **Assign** adds another assignee. **Delete** removes the task (soft-delete).

## Equipment  *(the previously-unwritten path)*
- [ ] **+ Assign equipment** → pick from inventory, set **purpose** (own_event/rental/other), dates, optional revenue → row appears.
- [ ] **Remove** clears the assignment.
- [ ] After assigning equipment + a DJ/performer, **Generate day-of tasks** (Overview) creates load/setup + soundcheck + template tasks; running it again does **not** duplicate them.

## Contracts  *(new `contracts` table — canonical; legacy `agreements` ignored)*
- [ ] **+ Add contract** → actor + kind (incl. **vendor**/**sponsor**) + fee + status → row appears.
- [ ] Inline status dropdown persists; **mark paid** toggles to "✓ paid".
- [ ] **Delete** removes it.

## Files
- [ ] Pick a file + kind + **Upload** → it lists (stored in the private `agreements` bucket).
- [ ] **Download** opens a working signed URL. **Delete** removes both the storage object and the row.

## Cross-checks
- [ ] **Edit core fields** (header) opens the existing Edit Event modal; saving refreshes the hub header.
- [ ] Money panel changes from inside the hub refresh the Overview cards.
- [ ] Nothing on the original 14 tabs changed.

---

## Automated data-layer proof (already green)

`tests/event_hub_datalayer_test.sql` exercised, against prod, **inside a transaction it rolls
back**: add participant · fee→expense · task+assignment+status · equipment_usage write · stage
update · `generate_day_of_tasks` run twice (idempotent: 3 rows → 3 rows) · contract · file row ·
`v_actor_full` roles. Result: **all assertions passed; zero rows persisted** (verified counts
unchanged afterward). Re-run any time:

```
POST https://api.supabase.com/v1/projects/yaytdosxfhcqatmhctzk/database/query
  Authorization: Bearer $SBP_PAT
  body: {"query": <contents of tests/event_hub_datalayer_test.sql>}
Expect an ERROR message beginning "TEST_RESULTS_OK { … }".
```
