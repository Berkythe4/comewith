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

---

# Sprint 2 — UX pass + Money fix + IG KPI (live-drive)

Walk on `dashboard.html` signed in as admin.

## Money bug (was: ticket added shows in Overview, not in Money)
- [ ] Open an event → **Money** section. It now **lists line items inline** (Tickets / Other income / Expenses / Donations / Sponsorships) — not just a button.
- [ ] Add a ticket tier (Tier + Qty + Unit $ → **Add**). It appears in the list **immediately**; **Overview** gross/ticket cards reflect it; the **Events** tab bar/totals update. Repeat for income, expense, donation, sponsorship — each shows instantly.
- [ ] Delete a line → it disappears immediately and Overview updates.
- [ ] The **Events-tab Money button** (modal) still works exactly as before.

## People — bulk + multi-role + edit
- [ ] **+ Add people** → search, tick **several** actors at once; type a name + **add** to stage brand-new people; pick **roles** (multi-select chips, add a custom one) + a batch fee → **Add selected** creates them all in one go.
- [ ] Each person shows their **roles as chips**; multiple roles on **one row** (not duplicate rows).
- [ ] **Edit** a person → change roles (chips), fee, bill order, set times, contractor flag — no remove-and-re-add. Saves and reflects instantly.
- [ ] **Fee → expense** still one click (manual, run-once-warned). **Remove** removes the participation only.

## Equipment — multi-select + edit
- [ ] **+ Assign equipment** → tick **several** inventory items, set purpose + dates **once** → **Assign selected** creates all assignments.
- [ ] **Edit** an assignment (purpose/dates/revenue); **Remove** clears it. Immediate reflection.

## Contracts — edit + document
- [ ] **+ Add contract** → optionally attach a **document** inline (uploads to the private bucket, links as the contract's doc).
- [ ] **Edit** a contract (actor/kind/fee/status/notes) and replace its doc. **Doc** button downloads it via a signed URL. Inline **status** + **mark paid** still work.

## IG followers KPI
- [ ] **Strategy tab → "Log IG"** (and **event hub header → "Log IG followers"**) → one form, **3 accounts** (Come With / Berky / Dance Infusion), each showing last value; type a number to see the **▲/▼ delta**; **Save all** writes today's snapshots in one action.

## Automated data-layer proof (green, rolled back)
`tests/event_hub_sprint2_test.sql` — multi-role + role=roles[1], one-per-actor unique enforced, bulk people, multi-equipment, money inserts, contract edit + document_id wiring, **day-of generator reading a secondary `dj` from `roles[]`**, IG 3-account upsert, and audit_log capturing the hub tables. All assertions pass; zero rows persisted (counts verified unchanged).

---

# Sprint 3 — Venue/contact matrix (3a) + conditional workflows & template editor (3b)

## 3a — Venues + contact matrix + "last time" lookup
- [ ] New **Venues** tab → list of venues (Add / Edit / Archive). Toggle **Venues | Vendors** at top.
- [ ] **Open** a venue → detail with its fields, **"Last event here: …"**, and the **contact matrix** (people you deal with) — Add / Edit / Remove a contact, each with a **function** (booking / sound / day_of / gm / security) and a **primary** star.
- [ ] Contacts are actors (pick existing or create new inline) — no parallel people table. Adding a venue contact tags them `venue_contact` in the actor model.
- [ ] **Vendors** toggle → vendors are actors with the `vendor` role; each has its own contact matrix (`vendor_contacts`).
- [ ] On an **event hub → Overview**, a **Venue** box: set/change the venue (pick or create inline). Once set, it surfaces **"Contacts here:"** (ordered primary-first, then most-recently-involved) with one-click **involve** (adds them as a participant), plus **"last event here"**.

## 3b — Conditional workflows, outreach auto-assign, template editor, assign fix
- [ ] **Edit event → Logistics → "Come With is providing the gear"** checkbox; the hub header shows **Gear: CW providing / Venue·other**.
- [ ] **Generate task checklist** (hub Overview) branches on the gear flag: gear-on → load-in / soundcheck / breakdown / return; gear-off → "confirm house gear with venue". Idempotent (re-run never duplicates).
- [ ] **Outreach tasks auto-assign via the matrix:** e.g. "Send rider to sound contact" lands on the venue's **sound** contact; "Confirm load-in time with venue" on the **booking** contact. If no contact is on file for that function, the task generates **unassigned with a "assign a … contact" hint** (prompting you to grow the matrix).
- [ ] **Templates tab** = the standard-workflow editor: grouped by event type → phase; **add / edit / remove / reorder (↑↓)**; set **offset**, **phase**, **gear applicability** (gear / no_gear / both), and **auto-assign target** (venue:sound, venue:booking, vendor, …). Copy states edits are **future-only**.
- [ ] **Editing a template does NOT rewrite tasks already on an event** — only future generation changes. (Proven by test.)
- [ ] **Assign-task picker is grouped:** *This event's people* · *Your team* (always assignable) · *Venue contacts* — no longer "everyone".

## Automated data-layer proof (green, rolled back)
- `tests/contact_matrix_test.sql` (3a gate) — venue CRUD, venue+vendor contact links, one-per-function enforced, set-venue-on-event, `v_venue_contacts` lookup with `last_event_with` recency seam. All green; zero persisted.
- `tests/conditional_workflows_test.sql` (3b) — gear-off vs gear-on task sets, rider offset = T−14, **outreach auto-assigned to the sound contact**, booking outreach **degrades to unassigned + hint**, idempotent re-run, and **future-only** (editing a template leaves an already-generated task's due_date unchanged). All green; zero persisted (template offset rolled back).
