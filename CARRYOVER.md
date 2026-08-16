# Carryover — 2026-08-15 (Strategy board rebuilt: real trends, six categories, a funnel · HENRY'S machine)

**FIVE closes landed on 2026-08-15, across three machines.** This is the latest, run by
**Henry**. In order: the desktop's radio-discovery audit, Henry's Notes module, Keith's
laptop radio-window/paging work, the desktop's full-admin/site-owner change, then this
Strategy rebuild. Nothing was overwritten — every "This session shipped" block is
preserved below, newest first.

**The close routine is now the MERGE ROUTINE** (`MERGE_ROUTINE.md`, was
`SESSION_CLOSE_PROMPTS.md`). Renamed because with three machines shipping, every close is
also a merge — and it now opens with a mandatory `git fetch` + migration-number check.

Pickup order: this → `DEV_DOCS/claude-memory/MEMORY.md` → `LEARNINGS.md` → `ROADMAP.md` → `CLAUDE.md`.
Ritual: `MERGE_ROUTINE.md` (was `SESSION_CLOSE_PROMPTS.md`). DI data load detail: `events/dance-infusion/DI_DATA_LOAD_LOG.md`.

## 👉 If you are the LAPTOP, start here

Nothing is lost — but one thing about this repo isn't obvious from a fresh checkout:
**Claude memory doesn't sync between machines.** The desktop's 60 memory files are
snapshotted into **`DEV_DOCS/claude-memory/`** (index: `MEMORY.md`). Read the index;
open individual files as needed. Re-snapshot at every close.

The radio window fix is merged, live, and **has now been exercised** — Keith ran a real
"↻ Pull shows" at 17:13 on 8/15 and the numbers moved (see State summary). Everything
below is current as of that pull plus this session's deploys.

## ⛔ PRIORITY CONTEXT (⚠ dated 2026-06-02 — unverified, confirm with Keith)
**Come With was set MAINTENANCE-ONLY** while the **CWF (Come With Fitness) BRD** ran as
project #1, due **June 15, 2026** — two months past. Actual work since has been steady
Come With radio/dashboard building, so this framing is stale. What still stands, and is
a hard rule either way: **nothing Come With Fitness in this repo** (dashboard / schema /
pages) — LEARNINGS §5 and CLAUDE.md "Scope".

## State summary (verified against prod 2026-08-15)
- **Prod:** Supabase `yaytdosxfhcqatmhctzk`; live at comewith.org (Netlify auto-deploy from `master`).
- **Migrations: files through `145_event_funnel.sql`; 138–145 ALL APPLIED to prod
  2026-08-15.** 141 (which had been sitting written-but-unapplied) went in this session,
  along with 142/143/144/145.
  Tracked history and prod now agree. `145_event_funnel.sql` was briefly applied to prod
  while its file sat on an unmerged branch — the unavoidable shape of a DB change that
  ships before its UI, since the database is not branchable — and PR #10 closed that gap
  the same day.
  ⚠ **The numbering collided TWICE in one day.** First: 140 was authored as `138` on the
  desktop while the laptop was landing its own 138/139. Then, hours later, on Henry's
  machine: `141_brand_favicon.sql` was authored as `140` — the number read off a local
  `master` that had not been fetched — while the desktop's `140_site_owner.sql` was already
  reaching prod. Renumbered to 141 in `19de619` because site_owner got there first.
  **Take the next number only after `git fetch`.** Note what makes this bite: neither
  collision failed loudly. Git merged two `140_*.sql` files without a murmur, because
  duplicate numbers conflict in **prod**, not in git — nothing tells you until the wrong
  thing runs, or doesn't.
  Applied via the Management API (`db.py`, `SBP_PAT` in `.env`), not the CLI — the CLI is
  linked to **staging**, so always pass the prod ref explicitly. The migration **files** are
  the tracked source of truth.
- **Financial views:** all five re-verified anon **401** on 2026-08-15, *after* both migrations
  and both deploys (`v_event_summary`, `v_kpi_event_financials`, `v_kpi_parties`,
  `v_kpi_dance_infusion`, `v_kpi_dashboard`).
  ⚠ Still **NOT revoked from `authenticated`** — the GATED BLOCKER before any customer/external login.
- **Roles:** master_admin / sub_admin / customer via `public.is_admin()`; `donor` + `staff` on `actors`.
- **`master_admin` is now THREE people** — Keith, **Martin and Henry** (promoted 2026-08-15).
  `sub_admin`: Janelle, Liz. **`profiles.is_owner` = Keith, one row**, protected by the
  `protect_site_owner()` trigger (140): no other admin can change the owner's role,
  `deleted_at` or `is_owner`, or delete that row. Verified on prod with 10 impersonation
  checks inside BEGIN..ROLLBACK — 5 blocked, 5 correctly allowed. LEARNINGS §19.
  The guard exempts service-role callers by design, so it protects the APP, not the
  PROJECT — the service key, `SBP_PAT`, Supabase dashboard, GitHub and Netlify all sit
  above it and stay Keith-only.
- **Latest LEARNINGS §:** 23.
- **Git:** `master` = `f5e4242`, pushed and live. **Everything from this session is
  merged — no open PRs.** PR #7 (categories + data layer), #8 and #11 (post-apply
  checks), #9 (bar sizing), #10 (funnel + migration 145), #12 (bar width + hover card),
  #13 (this close), #14 (calendar focus band, from another machine) and #15 (chart title
  alignment). All of it is on comewith.org.
  ⚠ **The per-event bar chart took THREE passes after its first merge** — #9 stopped it
  stretching, #12 widened it from an unreadable 26px and added a hover card, #15
  right-aligned the titles so they sat over the data instead of over empty space. Every
  one was found by Keith looking at the deployed page, because there is no local render
  check (see the risk note in this session's block).
  Older unmerged branches: `fix-lognumbers-optgroups`, `docs/roadmap-reconcile`,
  `event-hub-sprint-1`. Stale local-only branch `feat/strategy-phase1-data-truth` can be
  deleted — it was renamed to `feature/strategy-rebuild` early in the session.
- **Financial views: all five return anon 401**, re-verified at this close
  (`v_event_summary`, `v_kpi_event_financials`, `v_kpi_parties`, `v_kpi_dance_infusion`,
  `v_kpi_dashboard`), plus the six views 142/145 added.
- **Edge functions on prod:** `dj-station` **v9**, `pull-dice` **v6**, `pull-ticketmaster`
  **v8**, `pull-ra-market` **v15** — all deployed 2026-08-15 from the laptop and verified by
  reading the live bundles back.
- **Henry has prod access now** — his Supabase account was added to the **Come With** org on
  2026-08-15, so his own `SBP_PAT` reaches `comewith-prod` and `comewith-staging`. His `.env`
  carries only `SBP_PAT` + `SBP_REF_PROD`, not the `SUPABASE_PROD_URL` /
  `SUPABASE_PROD_PUBLISHABLE_KEY` the contract expects.
- **Edge-function deploys go through `scripts/deploy_edge_function.py`**, not the CLI. CLI
  2.101.0 rejects the newer `sbp_v0_…` PAT format outright ("Invalid access token format")
  *and* is linked to staging; the Management API takes the same token fine. The script
  preserves each function's existing `verify_jwt` rather than resetting it — `dj-station` is
  `verify_jwt=false` (token-gated, no login) and flipping it locks every DJ out.
- **PostgREST `max_rows` on prod is 1000** and it truncates silently. Any unbounded
  `.select()` over `ra_events` / `ra_artists` / `sc_artist_cache` is already over that line.
  The dashboard pages through `sbAll()` now; new code must too. LEARNINGS §18.
- **Radio discovery pool (prod, as of the 2026-08-15 17:13 pull — the first one AFTER the
  DICE/TM fixes went live):** 1,047 future RA events (→ 11/13), **240 DICE** (→ 9/07, was 124
  → 8/21), **40 TM across 16 venues** (was 25/10, now including Brooklyn Steel, Brooklyn
  Paramount, Brooklyn Bowl, Warsaw, Music Hall of Williamsburg, Under the 'K' Bridge Park).
  2,874 RA artists total, 1,594 with a future date. **No cron pulls any of this** — `cron.job`
  runs only publish/retention/YouTube; the pool is as fresh as the last manual "↻ Pull shows".
- **Docs freshness:** `ROADMAP.md` is reconciled to **2026-06-02** and predates the whole
  radio build; treat CARRYOVER + `DEV_DOCS/claude-memory/` as the true state.
- **Tools:** `/tools/actor-inspector.html` · `/tools/test-checklist.html` · `/tools/visualizer.html`
  deployed on comewith.org, admin-gated via the staging guard.

## Tomorrow's default
**Set a `ticket_url` on the two upcoming events — Come With #2 (14 Nov) and Dance
Infusion #3 (10 Oct).** Both are in `planning` with none, and the funnel cannot measure a
single thing without one: the homepage only renders a "Get tickets" link when a URL
exists, so there is no click to record. **The beacon cannot backfill**, so a link added
after promotion starts loses every click before it. This is five minutes in the event
editor and it is the difference between the funnel working and staying empty forever.

Then **open Strategy and actually look at it** — the rebuild is live on comewith.org but
nothing has been clicked through in a browser. Expand each of the six categories, confirm
the charts draw, and check the open/closed state survives a reload (it is per-user now).

Then, still outstanding from the desktop session: **have Martin and Henry sign in and
confirm they see everything you see** — Income, Expenses, Strategy and Users are the four
master-only modules that just opened up. Then open Users and confirm Keith carries the 👑
owner chip above the two new master chips.

Then, still outstanding from the laptop session: **reload the dashboard and pull once
more, then scan.** Everything from this session is
deployed but nothing has been clicked through a browser yet — the paging fix in particular
changes what the Radio panel loads, and it is the one change a human has to see working.

Second, cheap while you're in there: **exercise the Notes module end to end** (assign a note,
edit one, convert one to a task). Deployed and structurally verified, never clicked.

## Parked / next

**FIRST — set a `ticket_url` on Come With #2 (14 Nov) and Dance Infusion #3 (10 Oct).**
See "Tomorrow's default": the funnel is live and measuring nothing until one exists, and
the beacon cannot backfill.

Everything from this session is merged and live; there is nothing left to merge.

~~migration `141_brand_favicon.sql` is written but NOT applied~~ — **APPLIED 2026-08-15**
in this session. One row, `brand.favicon` with an empty value, so the Site Editor's picker
now renders and falls through to the static `/icons/favicon-32.png` default until
something is uploaded.

**Note on `db.py` permissions:** every call used to be blocked by the auto-mode classifier.
Henry added `Bash(SBP_REF=yaytdosxfhcqatmhctzk python db.py:*)` to
`.claude/settings.local.json` on 2026-08-15, so prod calls now run unprompted. The prod ref
is baked into the prefix on purpose — any other project still prompts. It is a standing
grant over arbitrary SQL (`db.py` has no read/write split), so the care moved from the
approval dialog into the migration: **dry-run every migration** by copying it with
`commit;` swapped for `rollback;`. That caught a nested-window-function error in 142 that
re-reading the SQL had not.

0. **Access change just landed — nobody has signed in under it yet.** Martin and Henry are
   full `master_admin` as of 2026-08-15. Worth saying out loud to both: they can now see and
   edit all company money (income, expenses, mileage, ticketing, sponsorships, donations),
   flip `financials_released` on any event, and change every other user's module access.
   They can also remove **each other** — the guard protects Keith only. If you want them
   peer-locked too, that's a further rule and it isn't built.
1. **Click through this session's work — none of it has been exercised in a browser.**
   In order: (a) hard-reload the dashboard and open Radio — the artist and cache counts
   should JUMP, because the panel was silently loading only 1,000 of 1,594 artists and
   1,000 of 1,956 cache rows; (b) set the "from" date a month or two out and confirm the
   toolbar's "📅 pool to" note and the ⚠ warning behave; (c) hit "↻ Pull shows" and read the
   toast — it now names the RA horizon, the DICE window and anything dropped; (d) open a
   real `dj.html?ep=<token>` and confirm the artist count is no longer pinned at 160.
2. **Scan the 8/18 window — 113 artists, not 677.** Post-pull truth: 982 artists in the
   8/18→9/15 window, 754 with a SoundCloud link, **641 already scanned, 113 never read**
   (517 cache rows were written 8/14 21:53, after the desktop's numbers were taken).
   "↻ Refresh music & data" on the Radio panel; it only reads artists with no cache row, so
   it's safe to re-run. NOTE: the pre-paging UI under-reported what was already scanned, so
   this count may look different in the browser than it did yesterday — the 113 is from SQL.
3. **DICE beyond ~4 weeks is still thin, by design now.** The pull aims its detail budget at
   the radio window rather than at a flat 90 days, so shows outside the window aren't
   refreshed — they're preserved from earlier pulls, not re-checked. If you move the window,
   pull again. `dropped_over_cap` in the toast is the signal that the budget bound.
4. **Scan cache never re-reads.** `raScan` skips anyone with a cache row forever; a producer
   who released tracks since their scan shows the stale catalog until "↻ re-read all" (which
   nukes and re-reads the whole window). Only 2 in-window artists are >30 days stale today —
   not yet urgent, will be.
5. **Financial views still readable by `authenticated`** — the standing GATED BLOCKER before
   any customer/external login.
6. **`feedback_log` answers anon `200 []`, not `401`.** Same latent shape 103 fixed for
   `sc_playlists` / `sc_playlist_tracks`: a table-level anon grant survives while RLS blocks
   every row, so it reads as an empty array instead of failing closed. **Verified not a leak** —
   the body is `[]` and an anon POST is refused 401. The source is visible in
   `016_feedback_log.sql` lines 27–28, which contains the exact
   `grant all on all tables in schema public to anon` that `CLAUDE.md` now forbids. One-line
   follow-up: `revoke all on public.feedback_log from anon;`. Not urgent, but it will scare
   whoever next audits `role_table_grants`.
7. **Two gate tests written but NEVER RUN:** `tests/notes_assignment_test.sql` (138) and
   `tests/notes_to_tasks_test.sql` (139). Both are `BEGIN..ROLLBACK`. They were blocked by a
   permission classifier on Henry's machine at the time; the Management API path works now, so
   they can simply be run. 138's trigger behaviour is still argued rather than observed.
8. **`.claude/settings.local.json` is tracked in git.** It's machine-local permission config —
   Henry's `Bash(git merge:*)` grant is sitting as an uncommitted modification right now and
   will follow whoever next stages everything. Probably wants `.gitignore` + `git rm --cached`.
9. **The task board's due-date window mixes UTC and local.** `calBoardTasks` builds the horizon
   with `toISOString()` but compares it to a local `today`, so in New York the window can run a
   day long after ~8pm. Pre-existing, untouched, noticed while making "next 7/30" include
   overdue. One-line fix, deliberately not bundled into an unrelated change.

## This session shipped (2026-08-15 — Strategy board rebuilt: real trends, six categories, a funnel · HENRY'S machine)
Keith opened with "the strategy page is unreadable and we're not getting actionable
insights from it". It was ~35 equal-weight cards in four workstream groups, one scroll.
Rebuilt in three phases, each its own PR.

**Migrations 141–145, all applied to prod and verified.**
- **141** — the long-parked `brand.favicon` row. One row, empty value.
- **142** — the data layer. `snapshot_kpis()` + a 06:30 UTC cron writes `v_kpi_computed`
  into `metric_snapshots`; `v_kpi_prior` / `v_kpi_event_series` / `v_kpi_content_recent` /
  `v_kpi_changed`; `user_dashboard_prefs`; duplicate active `kpi_targets` deactivated.
  Deliberately changed **zero pixels** on the deployed board.
- **143** — cards the six `*_last` metrics 142 computed but left inert.
- **144** — seeds `user_dashboard_prefs` from the old singleton. Caught late: without it
  the three hidden cards would have silently reappeared for all 5 admins.
- **145** — `v_event_funnel` + `v_site_exposure_30d`. Applied to prod ahead of its file
  for a few hours; PR #10 merged the same day and tracked history matches prod again.

**The bug under the whole complaint.** Every live-computed card had rendered
"– no prior reading" since it shipped: `v_kpi_dashboard` took `current_value` from
`coalesce(computed, snapshot)` but `prior_value` only from the second-latest **hand-logged**
reading, and nothing hand-logs net P&L or subscriber counts. The metrics that mattered
most were exactly the ones that could never show a trend. LEARNINGS §20.

**Decisions (LEARNINGS §20–23):** prior means different things per metric and now lives in
one view; board categories are derived from the **metric-key prefix, not
`kpi_targets.workstream`** (re-filing those rows would have made the already-deployed
renderer silently drop nine cards); lifetime averages moved to the drill-down and
last-event values became the headlines; **cost to raise $1 is the DI health metric**
(Keith's call), and it is `lte`, so colour follows the comparison not the sign; "0" and
"cannot be measured" are different claims, so blanks, missing targets and absent funnel
denominators all render as unknown rather than zero.

**What the numbers said once they were readable** — this is the actionable part:
last party **−$800 at 25% sell-through**; cost to raise **$0.69 vs a $0.50 target**, up
from $0.61 the event before; recent uploads averaging **103 views against 274** for the
five before them; mailing list **107 against a target of 1000**. Two genuinely good ones:
DI raised **$9,557 vs $2,943** and attendance **117 vs 42**.

**The funnel measures forward, not back.** The ticket CTA is on the **homepage**, not
`event.html` (which reads `v_public_recap` and is a retrospective archive with no CTA), so
clicks record with `path='/'` and are matched to an event by comparing `link_url` to
`events.ticket_url`. It reads empty because the beacon started 2026-07-24, the only two
events that ever had a `ticket_url` finished before that, and **neither upcoming event has
one**. LEARNINGS §23.

**Verified:** all five E1 financial views return anon **401** (plus the six new views);
138–145 confirmed applied against prod object-by-object; comewith.org confirmed serving
the Phase 2 markup; `node --check` on the extracted inline module against a `HEAD` control
at every step.

**Open risk — none of the UI has been seen rendering.** There is no local console check
for `dashboard.html`, the Browser pane cannot composite (screenshots time out, layout
boxes read `auto`), and PR #7 was merged before its preview was clicked through. The
category blocks, charts, collapse behaviour and funnel panel are all structurally verified
and visually unverified. The bar chart alone needed THREE corrective passes after it
merged — #9 (stop stretching), #12 (48px + hover card), #15 (right-align the titles) —
every one caught by Keith on the deployed page rather than by any check here. Treat that
as the honest error rate for UI shipped without a render check, and assume the parts
nobody has hovered over yet carry the same.

**Ran on Henry's machine.** All of it merged the same day (PRs #7–#13, #15) and is live
on comewith.org.

## This session shipped (2026-08-15 — site favicon, batched prod checks, toolchain · HENRY'S machine)
PRs **#3** and **#4** merged and deployed. Migration **141 written, NOT applied** (see Parked / next).

- **The public site had no favicon at all.** `dashboard.html` had linked
  `/icons/favicon-32.png` since it shipped, but no public page ever did — so comewith.org
  rendered the browser's blank default in every tab. The icon files were already in
  `/icons`; only `sw.js` and the dashboard referenced them. Static `<link>` tags now on all
  13 public pages, verified live (HTTP 200 + link present on each, and both icon assets
  serve as `image/png`).
- **`brand.favicon` is the override**, next to Logo image in Site Editor → Brand & logo.
  Uploads to `event-photos` at `site/favicon_<ts>_<name>`, same shape as `brand.logo`.
  `setFavicon()` only fires on a non-empty key, so an unset override falls through to the
  shipped icon rather than blanking the tab. **Only `index` / `watch` / `artist` honour the
  override** — the other ten don't fetch `site_content`, and adding a request per page load
  purely for a tab icon wasn't worth it until someone actually sets one.
- **Batched prod checks** — `supabase/checks/pre_apply.sql` + `post_apply.sql`, each a single
  `UNION ALL` statement. Every `db.py` call is its own prod approval, so the six-query
  checklist cost six interruptions per migration; this makes it two. `post_apply.sql` asserts
  the things this repo has actually regressed on — anon grants on the five financial views
  (016/017), anon grants on `sc_playlists`/`sc_playlist_tracks` (103, which returned
  `200 []` rather than failing loudly and hid for four migrations), RLS-enabled-with-no-policy,
  and role helpers that dropped the 098 `deleted_at` guard. **Never executed** — see above.
- **`db.py` needs `SBP_REF` passed explicitly.** `.env` holds only `SBP_REF_PROD`, so a bare
  `python db.py …` exits with "Set SBP_PAT and SBP_REF". Deliberately not fixed by adding
  `SBP_REF` to `.env`: passing it on the command line is what makes the target project
  visible in the call you're approving.
- **`db.py` has no read/write split.** Inline SQL and migration files share one code path,
  and the Management API endpoint takes multi-statement SQL, so `select …; drop …` runs
  both. No `Bash(python db.py:*)` pattern can be safely allowlisted, not even a
  `select`-prefixed one. Every call stays a manual approval, by design.
- **Toolchain on Henry's machine is complete** — `python` 3.12.10, `gh`, `node`, `git` all on
  PATH by bare name. Lesson worth keeping: **a tool that "isn't on PATH" mid-session is
  usually a stale process env, not shadowing.** Claude Code captures PATH at launch. The
  WindowsApps `python.exe` stub was never shadowing the real install — user PATH already
  ordered `Python312\` first. The fix was a restart, not clearing App Execution Aliases.
- **⚠ The `master` push guard is partial.** `Bash(git push:*)` in the local allowlist meant
  `git push origin master` — a live Netlify deploy — went through with no prompt; an `ask`
  rule did **not** override `allow`. Narrowed to per-branch-prefix allows so master prompts.
  But `gh pr merge` is still allowlisted and merging to `master` deploys just the same, so
  the PR route remains ungated. Local-settings-only, not committed.

## This session shipped (2026-08-15 — full admin for Martin & Henry, behind a site-owner guard · DESKTOP)
Migration **140 applied to prod**; `invite-user` **v9** deployed; dashboard owner chip merged.
- **Martin and Henry are `master_admin`.** Everything Keith has, with one exception.
- **The exception is enforced in the database, not the UI.** There is no role-change control
  in the dashboard at all — the hole was that `"Master admin can manage all profiles"` is
  `for all using (is_master_admin())` with no WITH CHECK, so any master_admin could PATCH
  Keith's profile row straight through PostgREST. 140 adds `profiles.is_owner` (Keith, one
  row) + a `protect_site_owner()` trigger covering `role`, `deleted_at`, `is_owner` and
  `DELETE` as one unit.
- **`deleted_at` was the vector that mattered**, not `role`: under the 098 deactivation
  contract a deactivated profile reads as no-role, so deactivating Keith would have locked
  him out with his role still saying `master_admin`.
- **Verified on prod**, impersonating Martin inside BEGIN..ROLLBACK: demote owner, deactivate
  owner, strip owner flag, grab ownership, delete owner row → **all 5 blocked**; edit owner's
  phone, change Henry's role, read company income, call `get_team_members()`, see the owner
  flag → **all 5 allowed**. Financial views re-checked anon 401 after.
- **`invite-user` hardened** — it runs as service role and so bypasses the trigger by design;
  it now refuses the owner's email itself.
- **Users tab:** 👑 owner chip + owner sorts first, so three identical "master" chips don't
  hide who runs the place (`get_team_members()` dropped/recreated to return `is_owner`).
- **Called out, not fixed:** the two new masters can remove each other and invite more
  masters; the 041–043 financial gate now applies only to Janelle and Liz; and the guard
  exempts service-role callers, so it protects the app, not the project.
- **Renamed the close routine to the MERGE ROUTINE** (`MERGE_ROUTINE.md`) with a mandatory
  Step 0: fetch before documenting, take migration numbers after pulling. This session
  authored `140` as `138` while the laptop was landing 138 and 139.

## This session shipped (2026-08-15 — Radio: future-dated windows, and the 1000-row truncation · LAPTOP)
Keith asked to plan an episode against a date months out — schedules that far ahead are
mostly confirmed, so the research is worth doing early. **No migration, no prod data writes.**
Four edge functions deployed (`dj-station` v9, `pull-dice` v6, `pull-ticketmaster` v8,
`pull-ra-market` v15), all verified by reading the live bundles back.
- **The dashboard was silently losing a third of the pool.** PostgREST `max_rows` is 1000 on
  prod and says nothing when it truncates. Every radio load was over it — 1,327 future events,
  1,594 future artists, 1,956 cache rows — and none was ordered, so *which* thousand came back
  was arbitrary. A far-future window could miss its own shows and look healthy; it also made
  scanned artists read as unscanned. All of them page through the new `sbAll()` now, ordered
  by primary key. **LEARNINGS §18.**
- **`dj-station` had the dashboard's own window bug**, on the one screen where a missing artist
  reads as "they have nothing": it filtered on `ra_artists.next_event_date` (the soonest-show
  column) *and* took a silent `.limit(160)` against a 982-artist window. Rebuilt on
  `ra_events.lineup` mirroring `raWindowPool()`, paged not capped, and it reports
  `scope.capped` + `pool_total` if the 1,500 safety stop ever binds.
- **The window can start in the future.** `dj-station` reads a start date off
  `dj_search_params` instead of anchoring on "whenever the DJ opened the link", and the
  episode form has a date field for it. All three pulls now take `from`/`to`, clamp at 180
  days instead of 90/120, and echo the window they pulled.
- **DICE gets the window, not a flat 90 days.** Its detail budget is finite and spent
  soonest-first, so aiming it at the weeks the episode is built from beats spending it on the
  near term and never arriving. Budget now scales with the window (~10 DICE shows/day in NYC;
  the flat 240 bound exactly on the 8/15 pull and cut the last two weeks off a 4-week window);
  ceiling raised 400 → 600.
- **Caught before it bit: the deletes would have eaten the pool.** All three pulls delete their
  own source's rows before re-inserting and all three bounded that delete only at the bottom —
  fine while a pull was always [today, today+90], destructive the moment a window can be
  narrower or start later. Pulling a 4-week window would have deleted every Ticketmaster show
  beyond it. Bounded at both ends now.
- **Also fixed:** the episode form REPLACED `dj_search_params` wholesale, so opening ✎ Details
  on an Elements edition and saving would have wiped its `artists`/`day_of` scope — it merges
  now. And the window's date picker is deliberately NOT capped at the pool's last date (you set
  it further out and pull to meet it); the toolbar reports how far each source reaches and
  warns when the window runs past it.
- **Verification, honestly scoped:** this laptop has neither `deno` nor `node`, so **nothing is
  compile-checked**. Changes were reviewed by reading, bracket-balanced against the pre-edit
  file as a control, and the deployed bundles grepped on prod for the new code. **Not
  verified: no human has clicked any of it in a browser** — Parked item 1.

## This session shipped (2026-08-15 — Notes module: assignment, editing, convert-to-task · HENRY'S machine)
Migrations **138 + 139 APPLIED to prod**; `master` = `0529125`, merged and **verified live on
comewith.org**. Financial views re-checked anon **401** after both. Two PRs: #1 (assignment),
#2 (edit / convert / Site bucket / due-filter). No existing row was modified by either migration.
- **Assignment (138).** `feedback_log.assigned_to` → `profiles` + `assigned_at`, partial index,
  and a BEFORE trigger that stamps the timestamp in the database rather than trusting the
  client clock. Inline picker per row, who-filter (Anyone / Mine / Unassigned / teammate),
  "logged by X" beside each note, optional assignee on quick capture. Assigning to someone
  else fires the existing `notify()` (kind `assigned`, 121) so a claim is visible without
  re-reading the tab. **Why `profiles` and not `actors`: LEARNINGS §16.**
- **Editing.** Notes were write-once. Anyone with Notes access can now edit any note — type,
  page, text, assignee, status. Closes Keith's own note from 2026-08-12, "Allow for editing
  notes that have been created."
- **Convert to task (139).** "⇢ Task" opens the Calendar's existing task modal pre-filled from
  the note; the note closes only once the task actually exists, and if the close fails the task
  stands and the note stays open *and says so*. `tasks_source_check` widened to admit `'note'`
  (the move 114 made for `'meeting'`); `tasks.feedback_note_id` mirrors `meeting_note_id`,
  `ON DELETE SET NULL` so deleting a note never deletes the work it became.
- **`site` bucket.** Added to `WS_PILLARS`, defaulted for converted notes, changeable in the
  modal. **No migration** — 116 made `pillar` free text precisely so a bucket stays a UI
  concern. The ternary chain in `wsPillarColor`/`wsPillarLabel` became `WS_PILLAR_EXTRA`.
- **Found and fixed on the way:** `calAddTask` had **no Bucket field at all** — the board could
  filter by bucket but nothing could ever set one, which is why **83 of 109 tasks have none**.
  Now on the modal for every task.
- **"Next 7 / 30 days" now includes overdue.** A horizon starting today hid exactly the work
  that most needed doing. Labels say so ("Overdue + next 7 days"); done-ness stays governed by
  the status chips rather than being silently re-decided.
- **Verification, honestly scoped:** the 1.22 MB inline module passes `node --check` after every
  merge, run with the pre-edit file as a control so an extraction artifact can't pass for a real
  error; both migrations verified on prod by introspection; the live site re-fetched and grepped
  for the shipped markup. **Not verified:** the two gate tests never ran (Parked item 7), and
  no human has clicked the feature in the real dashboard.
- **Machine note:** this session's Claude memory lives on Henry's machine
  (`~/.claude/projects/C--comewith/memory/`, 3 files) and was **deliberately NOT** copied into
  `DEV_DOCS/claude-memory/` — that folder is the desktop's set, and its README says not to merge
  two divergent memory sets by hand. The durable lessons from it went into LEARNINGS §16–17
  instead, which is where they're readable from any machine.

## This session shipped (2026-08-15 — Radio discovery audit: the window was filtering on the wrong date)
Full audit of shows → producers → tracks, run against prod. **No migration, no prod writes.**
**MERGED + DEPLOYED** (`88b2153`): dashboard live on Netlify (verified — `raWindowPool` is in
the served file), `pull-dice` **v5** and `pull-ticketmaster` **v7** live on prod (verified by
reading the deployed source back). Financial views re-checked anon **401** after the deploy.
- **The bug:** `ra_artists` collapses each artist to ONE row carrying `next_event_date` — their
  *soonest* show. The radio window filtered on that column, so an artist playing the 16th **and**
  the 25th vanished from a window starting the 18th. Keith's 8/18 + 4-week filter was hiding
  **77 artists, 70 with a SoundCloud link.** Fix indexes every date from `ra_events.lineup`;
  simulated on prod it recovers exactly those 77 (866 → 943), and re-points each artist at the
  show they play *in* the window (date, venue, that bill's genres).
- **DICE was 7 days deep, not 90.** It detail-fetched the first 160 candidates in tag order and
  saved 159 — the cap was binding exactly, so weeks 2–4 had **zero** DICE and nothing said so.
  Now date-filters before spending a fetch, takes soonest-first, caps at 240, reports
  `dropped_over_cap` + `last_date`.
- **Ticketmaster was Manhattan-only.** `city: "New York"` is a literal match at TM's end: 27
  future shows, 11 venues, nothing in Brooklyn or Queens. Now queries all five boroughs.
- **"Pull shows" swallowed TM/DICE failures**, so an outage and a real zero looked identical.
  It now names a source that didn't answer.
- Also confirmed **clean**: RA coverage itself (582 in-window events, listings to 11/10, taper
  looks like genuine listing decay, not truncation), and all five financial views at anon 401.
- **Docs:** `DEV_DOCS/claude-memory/` snapshot added so the laptop can read project state;
  `CLAUDE.md` gained a start-of-session / close-of-session section; `MERGE_ROUTINE.md`
  gained an "every close" block (re-snapshot memory, name the machine, don't merge to master
  to tidy up).

## This session shipped (2026-06-16 — Guest KPI fix: returning-match + filter-aware cards + mission/business split + ROADMAP)
Diagnose-first, then fix. Migration 040 applied to prod (additive views). Money untouched throughout.
- **Diagnosis (read-only gate):** DI#1 "28" is CORRECT — 42 RA rows = 42 tickets but 28 distinct buyers (10 multi-buys → 14 anonymous extra seats; RA captures only the buyer). The returning "1" had TWO bugs: (a) the sprint-6 ledger import skipped the 25 overlap emails → DI#1 people who also came to DI#2 got NO DI#2 attendance link (27/28 DI#1 guests lacked one); (b) returning was guest_id/email-exact, missing different-email/nickname same-person (Ethan Pollak, Liz↔Elizabeth).
- **Fix A (recovery, additive, money-free):** added the missing DI#2 `guest_event_attendance` links for existing guests in the ledger → DI#2 attendance 68→77 (+9). Backup `backups/prekpifix_2026-06-16_*.json`.
- **Fix B (migration 040):** `v_event_attendance_kpi` rebuilt with PERSON-identity matching (normalized full name + nickname canon, email fallback) — a KPI calc, not a record merge. **DI#2 returning 1 → 12** (43% of DI#1's 28; honestly below Keith's 50–75% gut because DI#1 has only 28 *identified* buyers, not 42 attendees, and some returnees used unmatchable identities). No false merges.
- **Filter-aware cards:** Guests-tab KPI cards now recompute for the active filtered set (event/returning/subscribed), immediate.
- **Mission/business split (`v_guest_spend_split`, type-driven):** DI = **mission** (MS-Society), CW Parties = **business**, else other. Surfaced in list + profile + cards, labeled "DI / Mission" vs "CW / Revenue" — never conflated.
- **Verify:** DI#1 $2,940 / DI#2 $9,557.33 financials identical before/after; guests still 97 (recovery added links, not guests); attendance 99→108.
- **ROADMAP.md reconciled** to true prod state: DONE (module series 030–040), QUEUED (audit cleanup → Artist → Vendor → Sponsor repoint → Equipment module), DECISIONS-WAITING (cold subs, door list, audit merges), GATED BLOCKER kept (financial-view lockdown before external login), pointer line kept.
- Deployed to master (Netlify). Only guest/KPI surfacing touched.

## This session shipped (2026-06-16 — Guest module: ledger import + actor graduation + repeat KPI + reconcile + audit)
Migrations 038 + 039 applied to prod (additive). DI#2 ledger people imported, money NEVER written.
- **Import:** ledger 113 rows → 77 emails → 25 overlap skipped → **52 net-new guests** + DI#2 `guest_event_attendance` (amount_spent = guest stat only, NO ticketing/income). Guests **45→97**.
- **Consent:** 0 net-new opt-outs; the **11 `ra_marketing_opt_in=False` were already-subscribed guests → UNSUBSCRIBED** (consent correction). Subscribers = 97 rows, **86 subscribed / 11 unsubscribed**.
- **Actor graduation:** `guests.actor_id` (mig 038). 13 LINK to existing actors (variants resolved: Elizabeth→Liz, Sauci→DJ Sauci Soni, Keith→Berky — no dups), **15 CREATE** clean donor/sponsor/staff, **15 FLAG** (fuzzy dups + payout artifacts — not created/merged). Actors **23→38**; 15 guests graduated. Mig 039 added `'staff'` to actor_roles enum (additive, like 029's 'donor').
- **Money untouched (proven):** DI#2 gross $9,557.33 / exp $6,557.33 / sponsor $6,225 / donations $292.44 — identical before/after. Dedupe: 0 dup guests/subs/actors. Backups in `backups/preguest_2026-06-16_*.json`.
- **Guest module (new Guests tab):** v_guest_stats list + filters (search, event, **subscribed-only = mailing list**, returning-only) + per-guest profile (history, spend, subscribed, actor link). KPI strip from `v_guest_kpis`; per-event new/returning from `v_event_attendance_kpi`.
- **Returning KPI:** DI#1 28 new/0 ret; Crossroads 2 new/1 ret (Liz); DI#2 67 new/1 ret (Claudia). 97 guests, 86 subscribed, 2 repeat, avg spend $58.87.
- **Reconciliation (read-only, NOT a double-count):** ledger $19,114.33 = gross mixed activity — sponsorship $9,950 + expense $4,750.33 + ticket $1,513 + payout $1,291 + zeffe_pkg $860 + donation $750 + comp $0. Differs from DB's reconciled net by scope (expenses/payouts/comps + in-kind sponsorship + DIY donation stream). Soft flag: ledger shows a $750 individual-donation stream + $9,950 sponsorship (incl in-kind) the DB headline consolidates differently — review if MS-facing granular totals are needed. No financial data changed.
- **Audit (`GUEST_ACTOR_AUDIT.md`, flag-only, nothing merged):** 6 same-human guest↔actor unlinked (Adam Cohen, Ethan Pollak, Francis/Theresa Berkman, Liz, Patrick — actor has no email so email-link didn't connect); 8 possible-dup/variant (Crossroads Café accent, Teri↔Theresa, etc. — 3 are likely *different* people, don't merge); Henry low-quality name; 5 payout artifacts excluded.
- Deployed to master (Netlify). Existing tabs untouched except Subscribers (prior) + new Guests tab.

## This session shipped (2026-06-16 — Attendee + mailing backfill with lifetime guest stats)
Migration 037 applied to prod (additive). Backfilled attendees → guests → subscribers, money untouched.
- **Sources:** 3 RA ticket exports (DI#1 42, Crossroads showcase 4, DI#2 20) = 66 rows → **45 unique guests** (dedupe by email). 0 no-email, 0 malformed, **0 explicit opt-outs**. The DI#2 door-list xlsx (81 names, **no emails**, fuzzy spellings overlapping RA) was **flagged, not imported** (no dedupe key → would dupe).
- **Rule applied:** subscribe everyone with email except explicit opt-outs → **all 45 subscribed** (none opted out). **Deliverability exposure flagged:** 20 cold-only guests (bought a ticket, never ticked RA marketing) — DI#1 15/28, Crossroads 1/3, DI#2 14/16. Tagged by source/segment so a cold-set unsubscribe is a one-liner if Keith wants.
- **Schema 037:** `guest_event_attendance` (additive guest↔event link carrying `amount_spent`) + `v_guest_stats` view (events_attended count+list, total_spent, first/last seen, subscribed). **Deliberately did NOT write `ticketing` rows** — that feeds `v_event_summary.ticket_revenue` and would double-count DI#1's reconciled income. Off-prompt: guest "total spent" comes from the link, event financials untouched.
- **Result:** guests 45, subscribers 45 (all subscribed), subscriber_segments 47 (per-event), attendance 47. **No dup guests/subscribers** (dedupe by email). Multi-event proven: Claudia (DI#1+DI#2 $65.70), Liz McQuillan (Crossroads+DI#1 $70). **Money untouched: ticketing=3, income=6 unchanged.** Backups in `backups/preattendee_2026-06-16_*.json`. Idempotent, tagged `[ATTENDEE BACKFILL 2026-06-16]`.
- **Surfacing:** Subscribers tab now shows real list + **Events + Total spent** columns (joined from `v_guest_stats`). Campaigns tab can send to the 45.
- **Flagged for Keith:** the 20 cold subscriptions (deliverability); the 81-name door list (no emails). See `ATTENDEE_BACKFILL_DRYRUN.md`.
- Deployed to master (Netlify). Existing tabs untouched except Subscribers enrichment.

## This session shipped (2026-06-16 — Sprint 4: chain fix + end-to-end verify + actors-only backfill)
Theme = interconnection (prove the links, not the nodes). Migrations 035 + 036 applied to prod (additive).
- **Venue-save bug = display/refetch**, not a save failure (DI#2.venue_id was correctly = Signal). The hub showed the name via a PostgREST FK-embed (`venue:venues(name)`) that returned null client-side; rest of `hubLoadEvent` already fetched owner explicitly. **Fix:** fetch venue explicitly by id (drop the embed). Now persists AND displays.
- **One gear task, not N:** migration 035 generator emits a single "Load / test / setup gear (see equipment sheet)" + new **`v_event_equipment_sheet`** view; Equipment tab has a printable **Load-in sheet**.
- **Venue-as-counterparty (A-a):** `venues.actor_id` (additive FK). Contract form preselects the venue's linked actor; "Make '<venue>' contractable" creates+links an org actor inline.
- **Delete-resurrection fixed:** generator suppression now matches a title that exists for the event INCLUDING soft-deleted → a deleted task is never re-added. Idempotency + future-only edits preserved.
- **036:** seeded a "Finalize & sign venue contract" outreach template (venue:booking) so the chain includes a contract task.
- **End-to-end verify (Phase 2 gate, rolled back) on real DI#2:** venue→contacts surface→one gear task→rider→sound, load-in→booking, contract→booking→venue contractable→deleted stays deleted. All green.
- **Backfill (Phase 4 — only persisting step):** actors + people-links ONLY, no money. Dry run found NO empty-in-DB+spreadsheet-money event → no POPULATE (DI#1 money already reconciled in the 2026-06-02 load; the `*-list.csv` files are attendee exports, not actor sources). **Created 3 actors** (32LVS, Gavin/Sara of Signal), **linked existing** (Keith→DI#1 dj solo-run, Kristen London→DI Artist Showcase), **5 people-links** (3 participants + Gavin=sound/Sara=other on Signal). Tagged `[BACKFILL 2026-06-16]`; idempotent; **DI#2 money counts identical before/after** (backup in `backups/prebackfill_2026-06-16_*.json`). Actors 20→23, no dups. Interconnection proven: backfilled Gavin auto-assigns DI#2's rider.
- **Flagged for Keith (not guessed):** Crossroads Café Artist Showcase roster; Rich Klein affiliation; Sara's exact function; the 42 DI#1 ticket-buyers (attendees). See `BACKFILL_DRYRUN.md`.
- **Off-prompt UX:** explicit-fetch over embed (matches owner pattern); load-in sheet as a print modal; venue-contractable affordance inline in the contract form; backfill deliberately conservative (people-only) because the DB already holds the money.
- Deployed to master (Netlify). Existing 16 tabs / prior hub work untouched.

## This session shipped (2026-06-16 — Sprint 3: Venue/contact matrix (3a) + conditional workflows & template editor (3b))
Resumed cleanly after a mid-3b.3 network drop (template-editor JS was the only unfinished piece — nav/panel/loadTab branch existed but `loadTemplates` et al. were missing; finished them). Both migrations were already applied to prod; re-verified objects + re-ran both test suites green before deploy.
- **Migration 033 (applied):** `venue_contacts` + `vendor_contacts` link tables (contacts are ACTORS — no parallel people table); `v_venue_contacts` / `v_vendor_contacts` views (`security_invoker`, anon-revoked) carrying a `last_event_with` recency column = the SEAM for future frequency ranking (v1 only ORDERS by it — no ML).
- **Migration 034 (applied):** `events.cw_providing_gear`; `task_templates.gear_applicability/target_function/sort_order` + unique(event_type,title); 10 seeded outreach templates (party+DI); `generate_day_of_tasks` rewritten — gear-branch + all-phase generation with due_date offsets + outreach AUTO-ASSIGN via the matrix (degrades to unassigned-with-hint).
- **3a UI:** new **Venues** tab (Venues|Vendors toggle, list→detail, contact matrix add/edit/remove, "last event here"); hub Overview venue box (set/change venue inline + ordered "contacts here" + one-click **involve**).
- **3b UI:** gear checkbox in Edit Event + hub "Gear" fact; "Generate task checklist" callout reflects gear mode; **Templates** tab editor (group by type→phase, add/edit/remove/reorder, gear applicability + target function; copy says **future-only**); **assign-task picker now grouped** (This event's people · Your team always · Venue contacts) — replaces show-everyone.
- **Off-prompt UX:** Venues + Templates as their own sidebar tabs (sanctioned surfaces); contact ordering primary→recency→last-touch; template reorder renumbers the group sequentially (clean given default sort_order=0); involve-contact one-click from the lookup.
- **Tests:** `tests/contact_matrix_test.sql` (3a gate) + `tests/conditional_workflows_test.sql` (3b) — both green, rolled back, zero persisted. JS parse-clean (170 KB module). 16 panels intact; existing 14 tabs + hub sprints 1&2 untouched (additive only). Deployed to master (Netlify).
- **Deferrals:** contract signing flow; global cross-event task board; frequency-weighted contact ranking (seam left in `last_event_with`).

## This session shipped (2026-06-16 — Event Hub Sprint 2: UX pass + Money bug + IG KPI)
Fixed the Money-section bug and closed the hub's UX gaps. **Migration 032 APPLIED to prod** (additive):
`event_participants.roles text[]` (backfilled, one-per-actor unique index), `generate_day_of_tasks`
repointed to read `roles[]` overlap, and **audit triggers on tasks/contracts/event_participants** (reuse
`audit_trigger_function`) so status/role transitions land in `audit_log` (answers the reporting Q10).
- **Money bug root cause:** the hub Money section only launched a modal and never listed line items —
  so adds showed in Overview's rollup but "vanished" in the Money tab. **Fix:** Money renders line items
  **inline** in the hub, refetching on every write; shared `loadMoneyData`/`moneySectionsHTML`/`moneyMutate`
  back both the hub-inline section and the Events-tab modal (one source of truth). Verified all 5 child types.
- **Multi-role participants:** one row per person, `roles[]` multi-select chips (custom roles allowed),
  `role` kept = `roles[1]` for back-compat. Day-of generator now catches a secondary `dj`/performer.
- **People UX:** bulk "Add people" (search + multi-select existing + stage new + batch roles/fee in one
  insert); full participant **edit** (roles/fee/bill_order/set times/contractor); fee→expense unchanged.
- **Equipment:** multi-select assign (batch purpose/dates) + edit/remove.
- **Contracts:** edit + inline document upload → `files(subject_type='contract')` → `contracts.document_id`,
  signed-URL download. (No signing flow yet — still deferred.)
- **IG KPI:** "Log IG" one form / 3 accounts (comewith/berky/danceinfusion) → upserts today's
  `metric_snapshots`; live ▲/▼ delta vs last; entry points on **Strategy toolbar** AND the **hub header**.
- **Off-prompt UX choices:** batch-default roles/fee then per-person edit (vs per-person-at-add); IG quick
  action placed in the hub **header** (always one click from landing) rather than buried in Overview;
  Overview "Open Money" now jumps to the inline section (no modal). Multi-role stored as `roles[]` (not a
  child table) — simplest additive path given how the hub + generator read roles.
- **Tests:** `tests/event_hub_sprint2_test.sql` (rolled-back, all green). **14 existing tabs untouched**
  except the Strategy toolbar's new "Log IG" button. Deployed to master (Netlify).
- **Note:** prod now has Keith's own sprint-1 live-test rows (1 contract, 1 equipment_usage, 1 task) — real, preserved.

## This session shipped (2026-06-16 — Event Hub, Sprint 1 of the module series)
Built the **Event Detail Hub** in `dashboard.html` (additive; the 14 existing tabs untouched). Reached
via an **Open** button on each Events row → a new `#panel-eventhub` with header (name + facts + **stage
stepper** + Edit core) and sections **Overview · People · Tasks · Money · Equipment · Contracts · Files**.
- **People** = `event_participants` (add existing/new actor, role, fee; **one-click Fee→expense**, manual — not auto-posted).
- **Tasks** = `tasks`+`task_assignments` (add, assign owner/doer/reviewer, inline status, soft-delete).
- **Money** = reuses the existing per-event Money panel.
- **Equipment** = `equipment_usage` — **the previously-unwritten event-attach path now writes** (feeds the day-of generator + ROI).
- **Contracts** = new `contracts` table (canonical; legacy `agreements` ignored), kinds incl. vendor/sponsor; status + mark-paid.
- **Files** = `files` table; documents in the private `agreements` bucket (signed-URL download).
- **Stage stepper** (idea→…→reported) updates `events.stage`, distinct from public `status`.
- **Day-of generator** = prominent button → `generate_day_of_tasks(p_event_id)`; idempotent.
- **Schema add:** migration **031_v_actor_full.sql** APPLIED to prod — `v_actor_full` (actors + roles[],
  `security_invoker=true`, anon-revoked). Establishes the **actor_*_details pattern** (`docs/ACTOR_DETAILS_PATTERN.md`)
  the Artist/Vendor sprints follow. No `actor_*_details` tables built yet (by design).
- **Locked decisions honored:** internal-admin only (no RLS/external-login changes), `type` enum as the
  operational axis (`series`/KPI views untouched), full cohesive design (template for the other modules).
- **Tests:** `tests/event_hub_datalayer_test.sql` — every hub write + the day-of RPC, run against prod
  **in a rolled-back transaction** (zero rows persisted; counts verified unchanged). All green.
- **Live-drive:** `EVENT_HUB_LIVE_DRIVE.md` (UI walkthrough).
- **Deferred to follow-ups:** linking an uploaded file as a contract's `document_id` (Files + Contracts
  ship separately today); a global cross-event task board; contract signing flow.

## This session shipped (2026-06-02 — data load)
Applied 023–028 + 029 to prod; populated the model with reconciled DI#1/#2 data; resolved the DI#1
duplicate (canonical "Dance Infusion #1"); proved role-overlap on real data; anon-401 held throughout.

## Open threads — needs Keith's eyes (in actor-inspector)
- **"19th & 7th Productions"** (existing contractor actor) — merge into Keith Berkman (Berky), or keep separate?
- Confirm DJ↔contractor matches + Keith = Berky.
- **Yankees-hats raffle donor** — unidentified, not loaded.
- **Held commits** (`261797d`, `5cbb51e`, close-out) — push pending Keith's go.

## Gated blocker
**Financial-view security fix** (revoke from `authenticated`, re-issue `security_invoker`) BEFORE any
customer/external login — covers existing customer-role logins too.

## How to verify
- Anon REST GET each of the 5 financial views → **401**.
- `v_kpi_dance_infusion` → DI#1 **39%**, DI#2 **31%** (% to mission = 1 − cost_to_raise).
