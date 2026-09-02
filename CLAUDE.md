# Come With — Project Conventions

Operational conventions for this repo. These are binding — follow them exactly.
Broader migration history and architecture live in `ROADMAP.md`.

## Start of session (read these, in this order)

Keith works from **three machines** now (the desktop, `C:\Users\keith\comewith`,
and Henry's). Claude Code's memory is per-machine and does **not** sync, so
everything a session needs to resume lives **in the repo**:

1. **`CARRYOVER.md`** — where the last session left off, and what's next. Start here.
2. **`DEV_DOCS/claude-memory/MEMORY.md`** — index of the desktop's Claude memory,
   snapshotted into the repo so either machine can read it. Open individual files
   from there as needed; they're background, and reflect what was true when written.
3. **`LEARNINGS.md`** — numbered, append-only design decisions with rationale.
4. **`ROADMAP.md`** — architecture + phase history (older; CARRYOVER wins on state).
5. **`reviews/session_YYYY-MM-DD.md`** — the narrative of a given session.

Then `git log --oneline -15` and `git branch -a` — work is sometimes parked on a
branch specifically so it does **not** auto-deploy (see below).

## End of session (the MERGE ROUTINE)

**`MERGE_ROUTINE.md` is the ritual** (called the "session close" until 2026-08-15;
renamed because three machines ship into this repo now and every close is also a
merge). Quick / Standard / Full by session size. Run it at the end of any session
that shipped real work; always when a migration or prod data was touched, or a new
standing rule was set. It carries the prod safety checks (the five financial views
must return anon **401**) and it's what keeps CARRYOVER / LEARNINGS / ROADMAP true
instead of drifting.

Four Come-With-specific rules on top of whichever variant you run:
- **`git fetch` FIRST, before writing any doc.** Other machines have shipped.
  CARRYOVER / LEARNINGS / ROADMAP are append-heavy shared files, and writing them
  against a stale base guarantees a conflict in the very file meant to tell the
  next person what's true.
- **Take the next migration number AFTER pulling.** On 2026-08-15 a migration was
  authored as `138` while the laptop was landing its own `138` and `139`. Duplicate
  numbers are a conflict in prod, not in git.
- **Re-snapshot Claude memory** into `DEV_DOCS/claude-memory/` (see the README
  there) so the other machines aren't more than one session behind.
- **`master` auto-deploys to Netlify.** Pushing `master` publishes the site and
  dashboard immediately. Work Keith hasn't green-lit goes on a branch and is
  named in CARRYOVER under "Parked / next"; never merge it to close a session out.
  Edge functions are separate — they go live only on an explicit deploy, via
  `scripts/deploy_edge_function.py` (the CLI can't do it — see CARRYOVER).

## Roles, and who owns the place

- **`master_admin` is now three people** — Keith, Martin and Henry (2026-08-15).
  `sub_admin`: Janelle, Liz. There is still **no `admin` role**.
- **`profiles.is_owner` marks the overall site owner — Keith, exactly one row.**
  Migration 140 added it plus the `protect_site_owner()` trigger: no other admin
  can change the owner's `role`, `deleted_at` or `is_owner`, or delete that row.
  Everything else on the owner's profile stays editable.
- **The demotion vector is `deleted_at`, not just `role`.** Under the 098
  deactivation contract a deactivated profile is treated as no-role by
  `is_admin()` / `is_master_admin()` / `user_can_access_module()`, so setting
  `deleted_at` locks the owner out while `role` still reads `master_admin`. Any
  new owner/role guard MUST cover both, or it isn't a guard.
- **The trigger deliberately exempts service-role callers** (`auth.uid()` null) so
  break-glass repair stays possible. That means it protects the *app*, not the
  *project*: anyone holding the service-role key, an `SBP_PAT`, the Supabase
  dashboard, GitHub or Netlify can still undo it. Edge functions running as
  service role must therefore enforce the rule themselves — `invite-user` refuses
  the owner's email for exactly this reason.

## Database / Supabase migrations

- **Project:** prod is `yaytdosxfhcqatmhctzk` (`comewith-prod`). The CLI is linked
  to staging; do not assume the link points at prod. Migrations live in
  `supabase/migrations/NNN_name.sql`, numbered in sequence (… 015–019 and up).
- **Introspect before apply.** Confirm live columns / view definitions / policies
  against prod, reconcile any `[VERIFY]` refs, and show the SQL/diff for review
  before applying anything to prod.
- **Roles:** `master_admin` / `sub_admin` / `customer`. There is **no `admin`
  role.** RLS uses the helper `public.is_admin()` (= `role in ('master_admin',
  'sub_admin')`). New admin-only tables: `for all using (public.is_admin())`.
- **NEVER use a blanket `grant ... to anon` in a migration.** Specifically, do not
  write `grant all on all tables in schema public to anon` (or to `authenticated`).
  `013_grants.sql`'s `ALTER DEFAULT PRIVILEGES` already grants the right
  privileges to new tables automatically. A broad grant silently re-grants SELECT
  on **all views too**, re-exposing financial views that were deliberately revoked
  from `anon` — this caused the **016/017 regression that 019 had to fix**. If you
  ever must re-grant, immediately re-assert every prior `revoke … from anon`, and
  verify anon access in the post-apply check (financial views must return 401).
- **Financial views are anon-revoked by design** (decision E1): `v_event_summary`,
  `v_kpi_event_financials`, `v_kpi_parties`, `v_kpi_dance_infusion`,
  `v_kpi_dashboard`. Keep them revoked. Verify with an anon REST GET → expect 401.
- **Apply discipline:** apply additively, verify on prod (objects, RLS has a real
  policy — never RLS-enabled-with-no-policy, admin can read/write, anon blocked),
  then commit the migration file so tracked history matches prod.
- **`db.py` needs `SBP_REF` passed explicitly, as a LITERAL** —
  `SBP_REF=yaytdosxfhcqatmhctzk python db.py …`. Do NOT add a bare `SBP_REF` to `.env`:
  the whole point is that the target project is visible in the command you approve, and
  `db.py` echoes it back (`[db.py] project=… source=…`).
  ⚠ `SBP_REF=$SBP_REF_PROD` **does not work from Claude Code's Bash tool** (it did from a
  normal shell, which is why it was written that way). `SBP_REF_PROD` lives only in `.env`,
  which `db.py` reads for itself and which is never sourced into the shell — so it expands
  to empty, `SBP_REF=""` lands in `os.environ`, `load_dotenv()` skips the key because it is
  already present, and you get the misleading "Set SBP_PAT and SBP_REF" error naming both
  even though only one is missing. The literal form is also what Henry's allowlist prefix
  matches, so it is the only form that runs unprompted.
- **DRY-RUN EVERY MIGRATION BEFORE APPLYING IT.** Copy the file with `commit;` swapped for
  `rollback;` and run that first. It executes the whole thing against real prod schema and
  data, then throws it away. This caught a nested-window-function error in 142 that
  re-reading the SQL had not — and with `db.py` now allowlisted for prod, the dry run is
  the only gate left between a typo and production.
- **Batch the checks — one query, one call.** Each `db.py` invocation is a separate
  prod approval, so the pre/post checks live as single UNION ALL statements in
  `supabase/checks/`: `pre_apply.sql` (edit its `targets` list, run before writing
  the migration) and `post_apply.sql` (run after; every row must read PASS). Two
  approvals per migration instead of one per check. `post_apply.sql` covers the
  anon grants behind the financial views, but still run the anon REST sweep as
  well — that exercises PostgREST end to end, which a grants query can't.
- **The anon sweep is `python scripts/check_anon_exposure.py`. Do not hand-roll it.**
  Two ways it was got wrong before, both of which reported a clean bill of health
  over a live leak (2026-08-20, LEARNINGS §37):
  - **There is no `SUPABASE_ANON_KEY` in `.env`.** The variable is
    `SUPABASE_PROD_PUBLISHABLE_KEY`. An empty apikey answers **401 for everything**,
    public or not, so a hand-rolled curl loop shows every object "blocked" and
    proves nothing. The script reads a known-public view FIRST and refuses to
    continue unless the key actually works.
  - **401 is the wrong thing to look for on a TABLE.** Views are anon-revoked and
    do answer 401. Tables carry an anon grant from 013 and rely on RLS, so they
    answer **200 with a body** — `[]` is correct, rows are a leak. Read the body,
    never the status.
- **The sweep does NOT check the five financial views. Run
  `python scripts/check_financial_views.py` as well.** `check_anon_exposure.py`
  discovers objects from the schema and is the right tool for breadth, but grep its
  output for `v_event_summary`, `v_kpi_event_financials`, `v_kpi_parties`,
  `v_kpi_dance_infusion` or `v_kpi_dashboard` and you get nothing — none of the five
  named in the rule two bullets up are in it. For months the close read its 54 confident
  lines as satisfying "all five return 401". Prod was fine; the check was decoration.
  The new script names all five, proves the publishable key works before trusting a
  single 401, and exits non-zero on any 200. **Both scripts, every close.** LEARNINGS §51.
- **`check_anon_exposure.py` DISCOVERS NOTHING — its two lists ARE the whole sweep.**
  Its own comment claimed "everything else discovered from the schema is checked too";
  there is no schema query in the file. An object named in neither `MUST_BE_EMPTY` nor
  `PUBLIC_OK` is never requested, and the closing "Nothing is exposed that should not
  be" says nothing about it. **Every migration that creates a table or view adds it to
  one of those lists in the same commit.** Grep the sweep's output for your new object
  before believing the summary line. LEARNINGS §58.
- **`revoke ... from anon` on a FUNCTION is a no-op.** Postgres grants `EXECUTE` on a
  new function to **`PUBLIC`**, and `anon` inherits it — revoking from `anon` removes a
  grant it never separately held, and the function stays wide open. Always
  `revoke all on function public.f(args) from public, anon;` then
  `grant execute on function public.f(args) to authenticated;`. 183 got this right for
  `snapshot_kpis`; 195 got it wrong anyway, because the `from anon` form *looks*
  correct and fails silently. Only the post-apply grant check caught it — which is why
  that check runs on every migration, not just the grant-ish ones. LEARNINGS §45.
- **A DESTRUCTIVE migration ships WITH its UI; an additive one may ship ahead.** The DB
  isn't branchable, so a schema change landing before its UI is normal here — but only
  when it's additive (new column/function/table), where the old UI keeps working. A drop
  (column, function signature, tightened constraint) breaks the deployed dashboard the
  instant it applies, and the outage lasts as long as the UI takes. 196 dropped
  `task_templates.event_type` and broke the live Templates page and gap scan until the
  branch merged. Either apply after the UI is merged, or split it: add the new shape →
  ship the UI → drop the old one. LEARNINGS §46.
- **`INSERT..RETURNING` enforces the SELECT policy mid-statement.** A security-definer
  helper that re-queries the table cannot see the not-yet-visible new row, so
  `.insert().select()` fails RLS even for the creator (bit us on 097 chat DMs).
  Put row-local predicates like `created_by = auth.uid()` directly in the SELECT
  policy. RLS can be smoke-tested on prod via the Management API:
  `set_config('request.jwt.claims', …)` + `set local role authenticated` inside
  BEGIN..ROLLBACK.
- **Deactivation contract (098):** `profiles.deleted_at` set = user deactivated;
  `is_admin()` / `is_master_admin()` / `user_can_access_module()` all treat that
  profile as no-role. Any new role helper MUST keep the `deleted_at is null` guard.
- **PostgREST `max_rows` is 1000 on prod, and truncation is SILENT.** A `.select()`
  with no `.range()` is a query with an undeclared cap — `ra_events` (1,327 future),
  `ra_artists` (1,594 future) and `sc_artist_cache` (1,956) are all past it today. Page
  it, and **order by a PRIMARY KEY**: ordering by a non-unique column (`event_date`,
  `next_event_date`) lets a tie straddle a page boundary and silently skip or repeat
  rows. The dashboard has `sbAll(build, pk)`; edge functions page inline. Any cap that
  can't be paged away MUST report what it dropped (`dropped_over_cap`, `scope.capped`)
  and the UI must surface it — a silent cap reads as "that's all there is". §18.
- **A delete keyed to a fetch range must be bounded at BOTH ends.** The pull functions
  clear their own source's rows before re-inserting; `gte(from)` alone was only ever
  correct because every pull ran [today, today+90]. Once a fetch can be narrower or
  start later, a bottom-bounded delete throws away everything past the window it just
  fetched. Widening a read is not just a read change if a write shares its range.

## Strategy board / KPIs

- **Categories are derived from the METRIC KEY PREFIX in `dashboard.html`, not from
  `kpi_targets.workstream`.** `radio.*` → Radio, `site.*` → Site, and so on. Do **not**
  "tidy up" by re-filing workstream values in a migration: the deployed renderer resolves
  categories client-side, so a DB re-file would silently drop those cards from the live
  board the moment it applied. A new `metric_key` lands in the right category with no
  migration; anything matching nothing renders under **Other**. LEARNINGS §21.
- **A computed metric has no history unless something writes one down.** `snapshot_kpis()`
  runs at 06:30 UTC (after the 06:00 YouTube pull) and writes `v_kpi_computed` into
  `metric_snapshots` as `source='computed'`. If you add a value to `v_kpi_computed`, it
  starts building history automatically — but it has **no prior value until tomorrow**.
- **"Prior" is defined per metric in `v_kpi_prior`, and nowhere else.** Previous completed
  event for event metrics, previous 5 uploads for recent content, nearest-30-days-ago for
  the rest. Never fall back to the LATEST reading when history is short — that compares a
  number to itself and renders a confident permanent "no change". LEARNINGS §20.
- **A value in `v_kpi_computed` is inert until a `kpi_targets` row exists.** That is the
  seam to use when a data change must not alter the deployed board: ship the view now,
  card it with the UI later.
- **Targets are edited from the dashboard ("Edit target"), never in SQL.** Duplicate active
  rows were deactivated in 142; `v_kpi_targets_current` de-dupes by `metric_key` with
  `DISTINCT ON`, so a second active row silently wins or loses on `effective_date`.
- **Charts on this board are RIGHT-ANCHORED — labels included.** The newest data is
  always at the right edge: bar tracks right-align so the most recent event sits at a
  fixed edge regardless of how many bars there are, and a sparkline's latest point is its
  last one. So chart titles and captions right-align too (`.cat-chart-cap` is
  `justify-content: flex-end`, with any secondary note pushed left by `margin-right:auto`
  so a note-less title does not fall back to the left). Bars also have a fixed preferred
  width that only shrinks on overflow — `flex: 1` makes three events fill the screen and
  stops reading as a chart at all. This took three corrective passes after merge (#9,
  #12, #15); match it rather than rediscovering it.
- **Never render a blank as zero.** "0" and "cannot be measured" are opposite claims.
  Progress bars only draw where a value AND a target exist; funnel conversions return null
  when the denominator is missing. LEARNINGS §23.
- **The funnel's ticket CTA is on the HOMEPAGE, not `event.html`.** `event.html` reads
  `v_public_recap` and is a retrospective archive page with no CTA at all, so ticket clicks
  record with `path='/'` and are matched to an event by comparing `link_url` to
  `events.ticket_url` (query string stripped both sides). **The beacon cannot backfill** —
  an event's `ticket_url` must be set BEFORE promotion starts or every click before it is
  lost. Keep `v_event_funnel` / `v_site_exposure_30d` anon-revoked like the other
  financial-adjacent views.

## Task templates (named SETS, since 196 — not per-event-type lists)

- **A template is a named SET of steps, and sets are FREE.** `task_template_sets`
  ("Party — standard", "Event Template v2") holds ordered `task_templates` rows via
  `set_id`. There is **no `event_type`** anywhere in the model — it was dropped in 196.
  Any set can be applied to any event; you pick which one when you build the checklist.
  Don't re-introduce a type filter "to tidy the picker" — keeping a v2 alongside v1 and
  trying it on one event is the whole point.
- **`events.task_template_set_id` is which set that event RUNS**, and the calendar gap
  panel measures "N of M steps missing" against it. It is written on the **first task
  actually created**, never when the set is merely picked, so an abandoned run can't
  relabel an event. An event with no set gets "no checklist picked", which is a
  different problem from "its set is empty" and is fixed on a different screen.
- **Decide and write are separate: `plan_event_tasks(event, set)` writes NOTHING**;
  `generate_day_of_tasks(event, set)` loops the plan and inserts. The dashboard walks the
  plan a step at a time so each can be edited or skipped. Any new surface must go through
  the planner — a preview that decides for itself is a third copy of the rules (the gap
  panel is already the second) and will drift silently. LEARNINGS §43.
- **Skip means skip THIS event.** Nothing is written, so the step still shows as missing
  and can be generated later. Removing a step for good is an edit on the Templates page.
- **`tasks.template_id` is the template link — NOT the title.** Steps are renamable at
  creation time, so title-matching would read a renamed step as missing forever and
  re-create the original next to it. Suppression and gap detection match on
  `template_id`, with the title kept only as a pre-195 fallback. LEARNINGS §44.
- **Gear and soundcheck steps are derived from the EVENT** (`cw_providing_gear`, the
  performing participants), not from a set, so they're offered whichever set you pick.
  They have no `template_id` and stay title-suppressed.
- **Phases order by MEANING** — planning → promo → day_of → wrap. `order by phase` is
  alphabetical and puts day_of first; harmless in a bulk insert, nonsense when a person
  is being walked through it.

## Planning / FP&A (the Planning tab, 197-202)

- **The unit of planning is an OFFERING, not an event.** `plan_offerings` is a
  repeatable thing you sell - a party, a DJ booking, a rental, a production gig -
  with a price, costs that scale with it, and a count per month. `creates_event`
  is a **flag**, not an assumption (a rental books no event), and `scale` is
  abstract and named per offering via `scale_label` ("Paid attendance", "Units").
  That indirection is deliberate: the same four tables model a SKU. Do not
  hardcode "event" into anything here.
- **Line bases are `per_unit` / `per_scale` / `pct_revenue`, and `pct_revenue` is
  EXPENSE-ONLY by constraint.** A percent-of-revenue income line is defined in
  terms of itself; forbidding it at the schema keeps every view a plain aggregate
  with no evaluation order to get wrong. `pct_revenue.amount` is a **percent**
  (6 = 6%), never a fraction.
- **`plan_offering_lines.category` must be a REAL P&L category.** It is the join
  key: `v_plan_vs_actual` matches plan to actual on `(period, ledger, section,
  category)`. Putting the unit's NAME there is exactly what broke the legacy
  `budget_lines` rows, which reported 100% variance silently for their whole
  life. LEARNINGS §47.
- **`budget_lines` rows with `version_id is null` are LEGACY and invisible to the
  planner.** The 37 hand-built rows are preserved as history; every planner view
  filters `version_id is not null`. Never "tidy up" by back-filling a version
  onto them - that double-counts against the offerings seeded from them.
- **Nothing in the planner may reach the P&L.** Plan rows live in their own
  tables and no view that computes actuals reads them. This is the separation 178
  established so a forecast can never be mistaken for money (LEARNINGS §33). A
  new planning view must not be joined into `v_pl_monthly` or anything under it.
- **A published round is frozen by TRIGGER** (`plan_frozen_guard`), not by the
  dashboard - a client-side guard is bypassed by any REST token, and actual vs
  forecast is meaningless if the forecast can be improved afterwards. Publishing
  **snapshots** the live plan rather than closing it: the `working` version stays
  editable forever, which is what a rolling forecast needs. Exactly one `working`
  version exists (partial unique index).
- **Copying into a published round is impossible by design, so publishing is a
  function.** `plan_publish_round()` creates the version in a transient state,
  fills it, then marks it published - all in one transaction. Do not try to
  reimplement it as client-side steps; the trigger will refuse them, and a
  half-copied round is worse than none.
- **`needs_review` / `provisional` are load-bearing, not decoration.** Lines
  seeded from a lump sum whose category could not be derived carry
  `needs_review`, and the board must present that offering as provisional rather
  than settled. Equally, `has_cost_model` / `has_revenue_model` exist so a
  missing side renders as "no cost" instead of a confident 100% margin - a `$0`
  line asserts "this costs nothing", which is a different claim from "not
  modelled yet". LEARNINGS §26, §48.
- **A pricing line is `quantity x rate`, and `quantity` composes with `basis` (203).**
  `per_unit` -> qty x amount ("2 DJs at $200"). `per_scale` -> qty x amount x scale,
  i.e. quantity is a multiplier ON TOP of the scale driver ("2 drinks a head"), and
  the ordinary case of qty 1 still reads as "$25 a head". `pct_revenue` -> quantity is
  **pinned to 1 by constraint**, because a percentage has no count; the UI must not
  offer a quantity box there, and switching a line onto `pct_revenue` has to send
  `quantity: 1` in the same write or the save fails a check constraint. `pct_revenue`
  is also expense-only, so **never offer it on an income line** - the schema refuses
  every save. `unit_label` ("tickets", "DJs") is display only; nothing computes from it.
- **The category box on a pricing line is a CONTROLLED LIST**, built from
  `v_pl_monthly` plus `revenue_streams` - the set a plan line can actually join to.
  Free text there is how a line reports 100% variance for its whole life without ever
  saying anything is wrong. A new line is created with an EMPTY category and
  `needs_review = true` rather than a plausible guess, and the confirm tick refuses to
  clear the flag while the category is blank.
- **`plan_offering_lines` is NOT versioned, so a published round is only half frozen.**
  `plan_publish_round()` snapshots volumes, overrides and overhead; the views join
  lines live with no version filter. Editing a price today therefore changes what a
  PUBLISHED round says it forecast - the exact thing the freeze exists to prevent.
  Known and open as of 2026-08-27; see CARRYOVER. Do not build anything new that
  relies on a published round's numbers being stable until it is fixed.
- **The forecast maths is implemented TWICE - in SQL (`v_plan_monthly`) and in JS
  (`planModelMonth`)** - because a lever that needs a round trip before it shows
  the answer is a form, not a lever. The view is the source of truth and is what
  a reload reads. **If you change one, change the other**, or the number you type
  against stops matching the number you reload into.
- **Parties have no per-head pricing yet, and that is a data fact, not an
  oversight.** There is not one paid `ticketing` row against any event of type
  `party`; priced tickets exist only for Dance Infusion. Do not seed a ticket
  price to "make the lever work" - it would feed straight into every forecast
  figure as invented evidence.

## Series contract (events.series)

`events.series` is free text. KPI views match it **exactly**. The Log Event form
MUST write `series = 'Come With Parties'` for parties and `series = 'Dance Infusion'`
for DI events, or those KPIs read empty. `'Come With Production'` is services
(we run someone else's production), not parties. `'Bookings'` (type `gig`,
added in 095) is when we're the **booked talent** at someone else's event —
performance fees go there, never under Production. The host/client who booked
us goes in `events.owner_actor_id` ("Host / booked by" in the edit-event modal).

## Mailing segments (brand delineation)

Two-level segments on `subscriber_segments`, established 2026-07-13:
- **Brand rollups** (what campaigns target): `come_with`, `dance_infusion`.
  A subscriber can hold both. Unsubscribe stays **global** (one master list).
- **Per-event segments** (cohort history): the event slug or event code,
  e.g. `come-with-7-11`, `di-02-2026-05`.

Every event import MUST add BOTH the event segment AND the matching brand
segment. Public signup widgets pass the brand segment (`come_with` on the
homepage; DI pages must pass `dance_infusion`). Never re-subscribe an
unsubscribed email during an import (e.g. `chaddercheesy@gmail.com`).

## Come With Radio (episodes live outside `events`)

- **`station_no` is the SHOW counter; `edition_seq` is the episode number.** Two
  different numbers — do not render either as "EP n" generically. `station_no`
  counts **every broadcast ever**, across series, and is displayed as **`SHOW n`**
  everywhere (dashboard, radio.html, dj.html, index.html, social-post titles,
  and the DB strings reworded in 137). `edition_seq` is a series' own numbering
  (Elements Ep1–4) and is what an audience knows — so the **rendered video keeps
  "EP n" and uses `edition_seq`** (`make_episode.py`). Elements Ep1 is SHOW 3.
  `station_no` is assigned at CREATION, not airtime, so it can drift out of
  broadcast order — `scripts/renumber_shows.py` fixes that (dry by default; never
  moves a published episode; also remaps `played_station_no` / `passed_station_no`
  / `carried_from`, which store the NUMBER rather than a key).
- **Radio episodes are numbered stations in `sc_playlists`** (`station_no`,
  lifecycle `building → testing → live → archived`), **NOT rows in `events`**.
  Do not create an event for an episode. The scheduled release date lives on
  `sc_playlists.drop_date` (radio's own tracker); the site teases the next drop
  via `get-station ?list=1 → next_drop`. Only one `building` row exists at a time
  (partial unique index) — auto-created with the next number when all are live.
- **Song memory `sc_song_log`** is the permanent played/passed/carried record —
  finalize logs played, sync/remove logs passed, finalize carries passed-not-
  played songs into the next station. Keep it in sync when touching station tracks.
- **Listener accounts** are `customer`-role auth users; `listener_*` tables are
  owner-RLS'd + anon-revoked. Never grant anon on them. `sc_playlists` /
  `sc_playlist_tracks` were also anon-revoked in **103** — they had carried
  table-level anon grants since 079 (RLS was blocking the rows, so an anon GET
  returned `200 []`, never data; now it's `401`). Public station reads are
  function-only through `get-station` (service role).
- **Venue identity lives in the DATABASE now (208), in two mechanisms with
  different authority.** `public.normalize_venue_name()` folds accents, case,
  `&`/`and`, punctuation and whitespace — that is applied **automatically**,
  because two names equal after it are the same room by construction.
  Trigram similarity **only ever suggests**, in `v_venue_name_review`: no cutoff
  separates `randall s island`/`randalls island` (0.97, same) from `green room`/
  `green room 42` (0.87, different), so a person rules and `venue_aliases` stores
  the decision. **Never auto-apply a fuzzy match** — a wrong merge rewrites
  history, which is what this exists to prevent. LEARNINGS §63.
- **`ra_events.venue_name` keeps what the feed sent; `venue_id` is the room.**
  Never normalise in place — the raw string is the only evidence of what the
  source said, and the only way to re-derive a ruling that turns out wrong.
  `link_venue_alias()` writes the alias AND re-points the history in one
  transaction; do not split it into two client calls.
- **Do not invent new folds.** A leading "The" and the feeds' "TBA - " prefix are
  plausible and are deliberately NOT folded — neither is a real collision in the
  data, and a fold ahead of evidence merges rooms that differ, silently.
- **`venueKey()` in `dashboard.html` MIRRORS `normalize_venue_name()`** and is only
  for spellings the database has not resolved. The SQL is the source of truth;
  change one, change the other. Prefer `venueIdent(e)` (the resolved `venue_id`
  when present) for grouping.
- **Venue names are FREE TEXT from three feeds — always group on `venueKey()`.**
  Prod holds `'Refuge'`, `'REFUGE'` and `'REFUGE '` (trailing space) as three
  strings; 155 of them are 149 real rooms (also Alphaville/ALPHAVILLE, Drom/DROM,
  H0l0/H0L0, public records/Public Records, `Dead Letter No. 9`/`No.9`). A `Set`
  of raw names dedupes nothing, and a dropdown then lists the same room three
  times with the artists split between them — which is how a busy room showed
  **one** artist. Fold case, punctuation and whitespace for the key; display the
  commonest spelling via `venueLabelPicker()`.
- **`next_venue` is ONE show. Never filter a venue on it.** `raWindowPool()` pins
  each artist to their soonest show in the window, so matching `next_venue` shows
  only artists whose *first* show is at that room — it hid 271 of 1,499
  artist-venue pairs. Filter on `venueKeys` (every room they play in the window).
  The tell is a singular field backing a plural question. LEARNINGS §62.
- **A filter that is on by default must say what it removed.** `Producers only`
  silently dropped a third of every room (Industry City 38 → 23). And when a room
  is isolated, show its lineup coverage — 35% of future RA events and 61% of DICE
  ones carry no lineup, so the filter gets blamed for a gap that is upstream.
- **Buzz is scored on what could be MEASURED, and the denominator moves.** Four
  inputs — top-track plays (.30), reach (.30), RSVP demand (.25), catalogue (.15) —
  each mapped 0–100 against a **fixed** anchor, never against the pool maximum: a
  pool-relative score changes when somebody else is scanned, so it cannot be
  compared week to week. An input that could not be measured is **dropped and the
  remaining weights renormalised**, never entered as 0, and the coverage % rides
  next to the number. Three absences look identical in JS and must not be conflated:
  **DICE and Ticketmaster publish no `attending` at all** (RA populates every one),
  a scan with `ok = false` **failed** (select `ok`, or a failure reads as "no
  music"), and **plays on an empty catalogue are undefined, not zero** — gating
  that wrong buries the selector DJs who upload nothing, which is who a radio show
  most wants (Ben UFO scores 56 vs 76). Catalogue size already records "uploads
  nothing"; do not charge an artist twice for it. LEARNINGS §61.
- **Rekordbox is the arrangement tool, not SoundCloud** (decided 2026-07-22).
  The set is bought and arranged in Rekordbox because SoundCloud isn't
  record-quality. The ① test push to SoundCloud + ↺ sync-back still exist for the
  first pass, but the **Rekordbox import owns final order**: dashboard
  "🎛 Import Rekordbox order" parses the playlist export (UTF-16 tab TSV, columns
  located BY HEADER NAME; also .m3u8/CSV/pasted lists), fuzzy-matches to the
  station, applies the order and pulls BPM/key. A station therefore holds songs
  that never came from SoundCloud — see `source` in migration 102 and the
  synthetic `man_…` `sc_track_id`.
- **Store metadata never overwrites Rekordbox.** `track-sources` only FILLS IN a
  missing bpm/song_key/camelot. Your own analysis of the file you own beats a
  store's tags. Matching must keep the **remix guard**: if either side names a
  remix/edit, the remixer has to match too, or the original mix matches
  "(X Remix)" and you buy the wrong track.
- **Beatport = PASTE-A-TOKEN, not a stored credential** (settled 2026-07-22 after
  testing against the live API). Hard facts, verified — don't re-litigate:
  - Access tokens live **600 seconds** (`exp - iat` on a real token). Ten minutes.
  - The **refresh token is unreachable from the browser**: not in localStorage
    (which holds only `token-refresh-result` → `{accessToken,…}`), and their site
    refreshes via a cookie JS can't read. Watching the Network tab for 30 min
    produced nothing; background tabs throttle the auto-refresh timer anyway.
  - `client_id` is in the **JWT payload** (`client_id` claim) — no hunting needed.
    `BEATPORT_CLIENT_ID` is set as a project secret.
  - So: a bookmarklet copies the current token from beatport.com, the "🛒 Where to
    buy" modal takes a paste, and `track-sources` caches it in
    `public.beatport_oauth` **only until its own JWT `exp`**. No standing
    credential at rest. Never `site_content` (anon-readable).
  - API shape confirmed: `tracks[]`, `key.camelot_number`/`camelot_letter`,
    `release.label.name`, `slug`+`id` → `beatport.com/track/<slug>/<id>`, and
    **`price.value` is in DOLLARS (1.49) not cents** — do not divide by 100.
  - Search "artist title" then **retry on title alone whenever nothing clears the
    match threshold** (not only on zero results): "Deeper Purpose Cigarettes"
    returns three confident-looking wrong hits, so a zero-results guard never
    fires and the real release is never searched for.
  `/beatport-cart` remains the way to actually fill a cart.
- **Bandcamp has no official API.** `track-sources` uses the endpoint their own
  search box calls: `POST bandcamp.com/api/bcsearch_public_api/1/autocomplete_elastic`
  with `{search_text, search_filter:"t", full_page:false, fan_id:null}`. The older
  `fuzzysearch/1/autocomplete_elastic` path is DEAD — and it answers **HTTP 200**
  with `{"error":true,"error_message":"bad function"}`, so checking `r.ok` alone
  silently reported every track as "not on Bandcamp". **Validate the payload, not
  the status**, and throw so the caller can say "couldn't reach Bandcamp" — never
  let an outage render as a definitive "not available".
- **Store matching is adversarial** — Bandcamp is full of DJ rips, bootlegs and
  flips of the track you actually want. Three guards, all regression-tested:
  (1) remix words are detected **anywhere**, not just in brackets ("Artist. Title.
  Pat Lok Flip." has none); (2) `(Radio Edit)`/`(Extended Mix)` are standard
  qualifiers, NOT remixes, and must still match their own release; (3) substring
  containment is **length-aware** — a flat score let "If U Need It" match
  "Sammy Virji: If U Need It (Callto Speed Garage Dub)". Returning "not found"
  beats sending Keith to buy the wrong file.
- **The public page never links the source playlist.** `get-station` deliberately
  does not select `sc_playlist_url`. Listeners get the FINAL MIX only
  (`mix_sc_track_url` / `mix_youtube_url`); to get the songs they come to the
  episode page and export the tracklist. Per-track links are fine. Don't
  "helpfully" re-add a playlist link.
- **Phase 1.1/2 pending:** YouTube auto-post at finalize; listener "export my
  saved playlist to my own SoundCloud" (OAuth per listener — designed, not built).

## Public artist profiles (`artist.html`) — added 2026-08-27

- **"Is this event public?" is TWO flags, and one of them is date-scoped.**
  `is_public` means *on the upcoming-events feed* — both consumers
  (`v_public_events` 030, `v_public_events_hero` 064) also filter
  `event_date >= current_date`, so setting it on a past event does nothing else
  anywhere. The past-facing flag is **`is_featured`**, which drives Recent Rooms
  via `v_public_recap`. Any "does the public see this event" gate needs
  **`is_public OR is_featured`** — `v_artist_gigs` (205) is the worked example.
  Gating on `is_public` alone took gigs from 60 to 24 and removed Dance Infusion
  #1 and #2 from every DI artist's page; only a pre-apply row count caught it.
  **Never gate a gig on `status = 'completed'`** — that was the original leak
  (065), publishing the lineup of every private booking. LEARNINGS §54.
- **An episode links on the name PRINTED, `sc_playlists.mix_by`, never on
  `assigned_actor_id`.** The FK is only whoever was given access to *build* the
  episode (130); linking on it renders "Mixed by \<guest\>" pointing at somebody
  else's profile. It is the fallback only when nothing is credited — at which
  point there is no name on screen anyway. Two public actors sharing a display
  name resolve to **neither**: no link beats the wrong link. One helper,
  `creditedArtist()` in `get-station`, decides both directions, so the episodes
  linking to an artist are exactly the episodes that artist's page lists back.
  LEARNINGS §55.
- **A profile page is only ever `public_profile = true`.** Everything public —
  the collective grid, the gigs list, the radio list, `?artist=` on `get-station`
  — gates on it, and a non-public actor must get the same empty answer as an
  unknown id so the endpoint never confirms a private profile exists.
- **The page is live-read, so a self-edit needs no deploy.** `artist-self`
  (token = `actors.edit_token`, no login) writes exactly the fields `artist.html`
  renders — bio / instagram / soundcloud / tiktok / website + photo. It cannot
  write `display_name` or `public_profile`; publishing stays Keith's toggle.

## Link-in-bio pages (`links.html`, `/links/<slug>`) — added 2026-08-31

- **The renderer is `links.html` and there is exactly one of it.** The dashboard
  editor's live preview is that same file in an iframe (`?preview=1`), fed unsaved
  form state over same-origin `postMessage`. **Never add a second renderer to
  `dashboard.html` "just for the preview"** — the planner already carries one
  deliberate duplicate (`v_plan_monthly` vs `planModelMonth`) and pays a standing
  change-both-or-drift tax for it; here there is no reason to. LEARNINGS §59.
- **The preview must show LESS than the editor holds, never more.** Inactive rows
  and rows outside their `starts_at`/`ends_at` window are dropped before posting,
  because that is what `v_public_link_items` does. A preview that flatters the page
  is worse than no preview.
- **`is_published` is a deliberate toggle and defaults false.** Publishing is never
  a side effect of editing, and an unknown slug and an unpublished slug return the
  same empty answer — the endpoint must not confirm a draft exists.
- **The schedule window is enforced in the VIEW, not the browser.** A link that
  should be gone must not ship to the page and get hidden by CSS.
- **Every rendered link carries `data-track="link:<uuid>"`.** The beacon then
  records the link's IDENTITY, not its URL, which is what `v_link_click_stats`
  matches on. Match on `link_url` instead and an edited URL silently inherits the
  old one's history, while a relative internal path never matches at all.
- **`link_pages.theme` is jsonb on purpose** — a new theme knob is a UI change, not
  a migration, which is the whole point of "customisable without Claude". Every
  key falls back to the Come With palette when absent, so a half-set theme still
  renders. A preset **fills the boxes and nothing else**: it must not wipe a
  background photo or a radius the user chose.
- **`_redirects` exists solely for these pages.** `/links` and `/links/*` rewrite
  (200) to `links.html`. Adding rules there affects the whole site — check it
  before assuming it is a links-only file.
- **Social link previews are still generic** (client-side render; crawlers do not
  run JS). `og_image_url` / `seo_description` exist on the row with nowhere to go
  until something server-renders the `<head>`. Do not add a per-page OG field to
  the editor before that exists — it would be a control that changes nothing.

## Weekly radio release (see `Radio/NOTES_WEEKLY_RELEASE.md`)

- **A private SoundCloud track will NOT embed.** `oembed` 404s on it, so the site's
  player renders nothing — and a `200` on the track PAGE proves nothing. EP 1 shipped
  with a dead embed this way. The link being saved early is fine; the track must be
  **public at publish time**.
- **The scheduled release goes through the `radio-publish-due` EDGE FUNCTION**, not
  the SQL function directly — SQL can't make HTTP calls, so it could never check the
  embed. It oembeds, flips `sharing=public` if needed, then publishes. A SoundCloud
  failure NEVER blocks the drop; it's written to `station_notes`. `cron
  radio-publish-backstop` still calls the SQL path so a broken function can't hold a
  release.
- **`/resolve?url=` does not return a PRIVATE track** from its plain permalink. Only
  `/me/tracks` lists an account's own private uploads. Env: `SC_CLIENT_ID` /
  `SC_CLIENT_SECRET`.
- **Never ask for a SoundCloud link** — retrieve it (`sc-connect` action `find_mix`,
  or ☁ Find my upload). A private upload has no shareable URL to give.
- **The Rekordbox export is the tracklist, not the dashboard.** The dashboard holds
  what was planned; the export holds what was played (EP 2: 19 planned vs 23 played,
  different order). Build cues from the export, then sync the dashboard. Export is
  UTF-16, tab separated, columns located BY HEADER NAME, and the Artist column is
  sometimes empty with the artist folded into the title.
- **Rekordbox owns bpm/key.** Store metadata only fills what's missing.
- **A missing cues column fails silently.** `render_card` needs `genres`,
  `release_date`, `show_date`, `show_venue` or that part of the card just isn't drawn.
- **Verify a render by pulling a real frame** (`ffmpeg -sseof -1.2 … -frames:v 1`),
  never by reading the code — a hardcoded stage cap silently dropped a new closing
  line from a finished 65-minute video.
- **Beatport access tokens live 600s FROM MINT**, not from copy; take one off a live
  request header, not `localStorage`. The **cart API works** — see the APIs map for
  the item shape.
- Episode material lives in **`Radio/Week N/`**, MP4 included. Show name in the video
  header is **Come With NYC Radio**.

## Media / recap links (must be publicly embeddable)

- **Validate every recap/media URL through `resolve-media` before storing.**
  SoundCloud share short-links (`on.soundcloud.com/…`) are redirects the embed
  player can't follow, and private/secret/wrong tracks oembed-404 — both fail
  **silently** on the site. The event editor already does this on save (resolves
  short links, strips `utm_*`/`si`, verifies oembed). Store only canonical,
  public URLs. `mediaKind()` matching "soundcloud.com" is NOT proof it embeds.
- **CSS: never use the `background:` shorthand on a variant/state class** (e.g.
  `.benefit`, `.audio`) layered over a base that set `background-size/position/
  repeat` — the shorthand resets them and breaks hero photos. Use `background-image:`.

## Editing `dashboard.html` (1.3 MB, one inline `<script type="module">`)

- **Never read the whole file into context.** It is ~1.3 MB and effectively all one
  inline module. Reading it once costs more than an entire ordinary session, and it
  gets read on nearly every dashboard PR. Locate the region with `grep -n`, pull only
  that window with `sed -n 'A,Bp'`, then `Edit` on an exact unique string. To review a
  change, `git diff` it — never `cat` the file.
- **Syntax-check by extraction, not by re-reading.** Extract the inline module body to
  a temp `.mjs` and run `node --check` on it. Run the **same extraction against the
  pre-edit version** (`git show HEAD:dashboard.html`) as a control — otherwise an
  artifact of the extraction itself reads as a real error introduced by the edit.
- **There is no local console check** — the Browser pane can't open `file://` URLs. The
  loop is `node --check` plus the deployed Netlify build.

- **Patch scripts must write atomically.** `open(path, "w")` truncates *before* it
  writes; on 2026-08-19 an encoding error mid-write left `dashboard.html` at zero
  bytes. Write to `path + ".tmp"` then `os.replace()`. Commit before any scripted
  sweep, because that is what recovery depends on.
- **Non-BMP characters in patch scripts:** write the literal character, never a
  surrogate-pair escape — Python turns those into lone surrogates and the write fails
  *after* it has already truncated. Use `chr(0x1F4F7)` if the literal is awkward.

## Bulk-edit surfaces (any tab with checkboxes)

- **A selection may only ever contain rows the current filter shows.** Prune on every
  render via `pruneSelection(sel, visibleRows)` and tell the user how many were
  dropped. Before this existed, narrowing a filter and applying a bulk change wrote to
  rows that were off screen — see LEARNINGS §28.
- Naming the count on the button is not enough; the count can be right while the
  membership is wrong.

## Money that leaves the business

- **Reportability (1099) is a stored decision on `actors.tax_1099_status`, never
  inferred from an expense category.** The $600 threshold is per payee per calendar
  year across every category. Payees with no actor row read `'no vendor'`, not
  `'undecided'` — they need linking before they can be ruled on. LEARNINGS §30.
- **Never seed a payment rail as an actor.** `Venmo`, `Sq *`, `Ubr `, `In *` and the
  like are statement descriptors describing how money moved, not who received it. One
  actor per rail silently merges unrelated payees. LEARNINGS §31.

## Pipeline / speculative work

- **Blue Sky = `events.stage = 'idea'`** with `expected_revenue` + `confidence`; there
  is no separate prospects table. `v_pipeline` returns the weighted number, and
  `needs_revenue_estimate` lists upcoming events with no money attached. Promote by
  moving the stage and booking real income; drop via `status = 'cancelled'`. Never let
  a speculative number reach the P&L. LEARNINGS §33.

## Photos

- **A photo needs an event *or* a subject, enforced by CHECK** — press shoots have no
  event, and inventing one puts a photo session in the events list and the P&L.
- **`is_public` defaults to FALSE.** The bucket itself is public, so never put
  documents there; publishing an image is a deliberate toggle. LEARNINGS §32.

## Secrets and third-party credentials

- **A Supabase secret is WRITE-ONLY. The Management API returns a SHA-256 DIGEST
  in the `value` field, not the value.** Every secret reads back as 64 hex
  characters whatever it holds. So: you cannot rename a secret, cannot copy one to
  another name, and **cannot judge a credential from what you read back** — length,
  charset and prefix are all properties of the digest. On 2026-08-25 that misread
  produced a confident accusation that Keith's eBay keys were "not an eBay keyset"
  (they were fine), then a test that authenticated with the *digests* and reported
  eBay's `invalid_client` as proof. **The only legitimate test of a credential is
  to send it to the system that owns it.** If an inspection contradicts what Keith
  says he entered, distrust the inspection. LEARNINGS §52.
- **`401 invalid_client` from eBay usually means the KEYSET IS DISABLED, not that
  the key is wrong.** eBay disables a production keyset until the account has a
  working Marketplace Account Deletion endpoint. `supabase/functions/ebay-account-deletion`
  is that endpoint, deployed **`--no-verify-jwt`** (eBay calls it unauthenticated; a
  gateway 401 reads to them as a dead host). It answers the challenge with
  `sha256(challengeCode + verificationToken + endpoint)`, and the endpoint in that
  hash comes from the `EBAY_DELETION_ENDPOINT` **secret, not `req.url`** — behind a
  proxy they differ and the mismatch yields a valid-looking hash eBay silently
  rejects. The shared token lives in `EBAY_VERIFICATION_TOKEN` and in the gitignored
  `ebay_verification_token.txt`; rotate BOTH sides together or verification breaks.
- **Supabase edge-function logs are empty on this plan.** `function_edge_logs`
  returns zero rows for every function, including ones invoked minutes earlier.
  Never read an empty result there as "it was never called" — run the control
  query first. §52.

## A source that costs money does not share a button with sources that don't

- **Split manual actions by COST and LATENCY, not by category.** Gear Watch had one
  "Run scan now" covering four sources. Three are API calls totalling ~6s; Facebook
  is an Apify scrape that BLOCKS until it finishes, once per target. Four of those
  never fit the edge runtime's **150s wall clock**, so the button returned 546 — and
  because the run row is written last, every press since it shipped left **no trace
  at all**. `manual` is now the free three; `facebook` is its own button behind a
  confirm naming the price. LEARNINGS §53.
- **Bound any blocking third-party call** (`AbortSignal.timeout`), leave room after
  it to finish the work, and **refuse to start one you cannot finish** — that wastes
  the credit and the wall clock together.
- **If only part of the list fits, ROTATE and say what you missed.** Starting at
  item #1 every press means the tail is never processed however often it is pressed
  — a permanent blind spot reporting itself as a clean run. Name the items skipped,
  not a count, and make `PARTIAL` its own state that the UI shows as a problem.

## Scheduled work (pg_cron → edge functions)

- **pg_cron cannot mint an admin JWT.** A scheduled job that needs an edge function calls
  a `security definer` helper that reads the function URL + **service-role key from
  `vault.decrypted_secrets` at call time** and posts via `net.http_post`. The key never
  goes in a migration. Settled 2026-08-18 (`146_gear_watch.sql`, `gear_watch_kick()`),
  closing the question `014_cron.sql` deferred in Phase 10. LEARNINGS §25.
- The edge function must accept a **service-role bearer OR an admin JWT** — the pattern
  `pull-ra-market`, `send-notice` and `send-push` already use.
- **A missing secret must be a documented no-op**, not a silent error: write the reason to
  the feature's own `last_status` and return.
- `search_path` on the helper must include `net`, or `http_post` is not found.
- **A source that is blocked or down must be reported as such — never as "nothing found".**
  Validate the payload, not the HTTP status; report per source; and mark a known-dead
  source DISABLED rather than letting it cry FAILED forever. LEARNINGS §24.
- **Never default a field that feeds a score, filter or sum.** A placeholder that flows
  into a computation stops being a placeholder and becomes invented evidence — prefer
  null and lose the signal. LEARNINGS §26.

## Scope

- This codebase is **Come With only**. Do **not** add anything Come With Fitness
  (CWF) anywhere — not in the dashboard, schema, or pages.
