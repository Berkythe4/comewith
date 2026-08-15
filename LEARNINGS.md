# LEARNINGS.md — Come With

**Owner:** Keith Berkman (Berky)
**Created:** 2026-05-29
**Status:** Living document — append as understanding deepens.
**Read this before changing anything fundamental.**

---

## A note on what this document is

This records the **WHY** behind Come With's design decisions, so they outlive the chat.
How-to-work conventions live in `CLAUDE.md`; the migration history lives in
`supabase/migrations/`; current state lives in `CARRYOVER.md`. This file is the rationale.

**Append, supersede, preserve history — never delete.** If a decision changes, add a new
superseding section rather than editing the old one in place.

---

## Section 1 — Canonical event revenue / P&L model (migration 022)

One revenue basis for every event; the two workstreams differ only in framing.
- `gross_revenue = ticket_revenue + other_income (bar/vendor/merch) + donations + sponsor_cash`
- **Every event:** `net_pl = gross_revenue − total_expenses`
- **Dance Infusion also:** `total_raised = gross_revenue + sponsor_in_kind`; `cost_to_raise = total_expenses ÷ total_raised`

**Why:** Parties net P&L previously summed only the `income` table, **excluding ticket
revenue** (a 50×$20 event showed $45 net). DI counted tickets+donations+sponsors but not
other income. Now both read the same base; bar/vendor/merch counts toward DI "raised" (every
$). In-kind counts toward `total_raised` but **not** `net_pl` (it isn't cash). Implemented by
splitting `sponsor_cash` vs `sponsor_in_kind` in `v_event_summary` and computing
gross/net/raised in the `v_kpi_event_financials` rollup.

## Section 2 — "Other income" never holds ticket sales (double-count guard)

The Log Event "Other income" field is **bar / vendor / merch only**; ticket money lives
**only** in the ticketing tier section. Since `net_pl` now includes both `ticket_revenue` and
`other_income`, entering tickets as income would double-count. The field is labeled to say so.

## Section 3 — Every computed metric carries a formula hover-tooltip (standing requirement)

All computed KPIs/metrics (net P&L, sell-through, cost-to-raise, total raised, gross revenue,
card progress/trend) show a plain-language formula on hover (ⓘ), via a reusable `infoTip()` +
`FORMULAS` pattern. **Future metrics must ship with one by default.**

## Section 4 — RLS + grants rules

Roles are `master_admin` / `sub_admin` / `customer`, checked via `public.is_admin()` — there
is **no `admin` role**. New migrations must **never** blanket-`grant ... to anon` (013's
`ALTER DEFAULT PRIVILEGES` already covers new tables); a broad grant silently re-exposed the
financial views — the **016/017 regression, fixed by 019**. The five financial views stay
anon-revoked; verify **401** at every close-out. (Also enforced in `CLAUDE.md`.)

## Section 5 — CWF scope (HARD RULE)

Nothing Come With Fitness in the Come With dashboard / schema / pages until the CWF BRD is
done and there is an explicit go decision. Come With is maintenance-only until then.

## Section 6 — Soft-delete over hard-delete (events & income)

Deleting an event or income row stamps `deleted_at = now()` (reversible; preserves financial
history); all views filter `deleted_at is null`. Child tables without a `deleted_at` column
(ticketing / donations / sponsorships) hard-delete in the Money panel.

## Section 7 — Income-cleanup events use placeholder series

The 5 events created during the 2026-05-29 income reconciliation use the two existing `series`
values as **placeholders** (e.g. showcases / Production tagged "Come With Parties" so they
don't pollute the DI fundraising trend). They will be reassigned once the event model is
redesigned — see `CARRYOVER.md` "Parked / next" and `ROADMAP.md`.

## Section 8 — Public metric is "% to the mission"; expense ratio is internal-only (2026-06-02)

Same data, opposite framing. **Public** (impact report + audit) always says **"% to the mission"**
= donated ÷ total raised (positive, donor-native). DI #2 = **31%** ($3,000 of $9,557). The words
"expense ratio" / "cost-to-raise" and the dashboard's inverse (~70% / $0.70-per-$1) are **internal /
dashboard only — never on a public page.**

- The **60%-to-mission target is a FORWARD commitment from DI #3 on**, not a bar DI #2 was measured
  against. Render the figure and the target together or not at all. Exact public line:
  *"Of every dollar raised at Dance Infusion #2, $0.31 reached the mission. Our commitment from
  Dance Infusion #3 forward: 60%."*
- **DI #1 → DI #2 is an ABSOLUTE-GROWTH story, never an efficiency story:** donated **2.6×**
  ($1,140→$3,000), raised **3.3×** ($2,940→$9,557). % to mission actually *fell* 39%→31% because
  DI #2 scaled up (prime-time, full production) — that's truthful and stays on the audit as
  baseline+forward-target; **never imply % to mission rose.**
- **Why:** avoids the raw $9K-vs-$3K optics (Liz feedback) while staying honest. **Open thread:**
  update the dashboard's internal `di.cost_to_raise` target to 40% expense / 60% to mission so the
  dashboard and the public audit reconcile to one goal.

## Section 9 — DI #2 "total raised = $9,557" reconciliation (2026-06-02)

The public audit's total raised is **$9,557 = expenses $6,557 + donated $3,000** (whole-dollar).
It reconciles as: **$9,264.89 gross from others + $130 Crossroads partner donation (routed directly
to MS, never through event accounts) + ~$162 founder contribution that covered remaining event
costs** — all counted as "raised" because money is fungible. The audit carries one understated,
public-facing line: *"Total raised includes a founder contribution that covered remaining event
costs."* Note the locked **$9,557 framing figure differs from the actual gross cash inflow
($9,264.89)** — the difference is the externally-routed $130 + the founder top-up; the audit uses
$9,557 deliberately (it's the costs+donated framing, not the bank inflow).

## Section 10 — Staging is a reusable client-side admin gate (2026-06-02)

`/staging/` gates review-before-publish pages by **reusing the dashboard's Supabase session auth** —
same project + publishable key, `getSession()` → `profiles.role` ∈ {master_admin, sub_admin}
(= `public.is_admin()`). **No second password system.** One front door: sign in at
`/dashboard.html`, the session persists, every `/staging/` page sees it.

- **Reusable pattern:** a single shared `staging/guard.js` + a **2-line include** in any page's
  `<head>` (`visibility:hidden` flash-guard + the module). Adding a report needs no auth rewiring;
  list it in the `REPORTS` manifest in `staging/index.html`. To publish a page publicly, delete the
  2 lines.
- **No session → redirect to `/dashboard.html`**; signed-in non-admin → "admins only" notice;
  admin → reveal. Fail-closed (stays hidden if the guard can't load).
- **Honest caveat (load-bearing):** this is **client-side gating on a static host — NOT real
  security.** Good for keeping low-sensitivity review pages out of public view. **Genuinely
  sensitive data (financials, rosters, venues) stays in Supabase behind RLS — never as static files
  in staging.** Staging is a review surface, not a data store.

## Section 11 — Guests are a lighter layer than actors (2026-06-02)

Ticket buyers / RSVPs stay in the lighter `guests` table — **not** promoted to `actors` by default
(privacy + volume: thousands of attendees we won't otherwise track). **Promote a guest to an `actor`
only on recurrence** (a guest who becomes a sponsor / vendor / performer / repeat). Keeps the actor
graph meaningful instead of bloated.

## Section 12 — Founder out-of-pocket = a donation attributed to Keith-the-actor (2026-06-02)

Keith's out-of-pocket spend on his own events is modeled as a **donation from Keith-the-actor** (role
`donor`), counted normally in `total_raised` — **not** a special "exclude-from-aggregates" category.
Money is fungible; it made the raise possible, so it IS money raised, attributed to the donor like
any other donor (DI#1 $1,800; DI#2 $162.44). Stored in `third_party_donations` with
`donor_name='Keith Berkman'`. (Donations have no actor FK yet — §14 / backlog.)

## Section 13 — Attendance = ticket-sales proxy, basis-tagged (2026-06-02)

No scan/headcount data exists, so **RSVP ≠ attendance — never report RSVP as attendance.** Where a
count is needed, use **ticket sales as the attendance proxy and tag the basis** ("tickets issued",
"RA tickets"); leave true `attendance` null unless a real count exists. (DI#1: "42 RA tickets",
attendance null.)

## Section 14 — "% to mission" is derived, not stored (2026-06-02)

The relational model has no native "donated" / "% to mission" field. It reconciles as
**% to mission = 1 − cost_to_raise_per_dollar**, which holds because we model **net_pl = donated**
(total_raised − expenses = what reached the mission: DI#1 2940−1800=1140; DI#2 9557−6557=3000).
`v_kpi_dance_infusion.cost_to_raise_per_dollar` = expenses ÷ total_raised; public "% to mission" =
1 − that. Headline figures also kept in `events.notes`. **Backlog:** add an actor FK to
`third_party_donations` (donations currently attributed by text `donor_name`, not linked to actor rows).

## Section 15 — A cache row that answers a different question than the one you're asking (2026-08-15)

`ra_artists` holds one row per artist carrying `next_event_date` — their **soonest**
upcoming show. It reads like "when this artist plays". It is not: it's "the first
time this artist plays, counting from whenever the pull ran". The Come With Radio
window filtered on that column, so an artist playing the 16th **and** the 25th
dropped out of a window starting on the 18th. On the 2026-08-18 + 4-week window that
hid **77 artists, 70 of them with a SoundCloud link** — invisibly, because a shorter
list looks exactly like a quieter month.

**Decision:** a date window over artists is evaluated against **`ra_events.lineup`**,
which carries every date, and the artist is re-pointed at the show they actually play
*inside* the window. The artist row's own date is one more candidate, not the source
of truth — partners and manual adds have no event row at all. One shared pool
(`raWindowPool()`) now feeds the list, the counts, the venue filter and all four
scan/match passes; each had re-derived the window separately, so they could disagree.

**The general rule:** a denormalized "next/latest/current X" column is a **summary of
the pull**, not a fact about the entity. Filter on the underlying rows. If you must
filter on the summary, the window has to start at the same instant the summary did.

Two more of the same family, found in the same audit — a bound that fires silently
reads as an empty result, and an empty result reads as truth:

- **DICE** detail-fetched the first 160 candidates *in tag order* and saved 159. The
  cap was binding exactly, so a 90-day request returned 7 days: weeks 2–4 of a 4-week
  window had zero DICE shows. Now it drops out-of-window candidates before spending a
  fetch, takes the rest soonest-first, and returns `dropped_over_cap` + `last_date`.
- **Ticketmaster's `city` is a literal string match.** "New York" is Manhattan — 27
  future shows across 11 venues, nothing in Brooklyn or Queens, so Brooklyn Steel /
  Kings Theatre / Avant Gardner did not exist as far as the pull was concerned. Now
  all five boroughs are queried, and only a dead *first* call is fatal.

**Standing requirement:** any pull that truncates, caps or samples must **report what
it dropped**, and the caller must surface it. "↻ Pull shows" was also swallowing TM
and DICE failures entirely — an outage and a genuine zero rendered identically. A
source that didn't answer is now named in the toast. Silent truncation is the failure
mode here, not the truncation itself.

## Section 16 — Internal work assigns to `profiles`; work touching outside people assigns to `actors` (2026-08-15)

Notes (`feedback_log`) gained an assignee in migration 138, and the obvious move was
to copy `task_assignments`, which points at **`actors`**. That would have been wrong.

The two tables answer different questions. A **task** can land on a DJ, a vendor, or a
venue contact — people who have no login and may never have one — so `actors`, the
people-and-orgs graph, is right for it. A **note** is the internal build log; its
readers are the five `profiles` with logins, and the notification bell (121) keys on
the **auth user id**. Assigning a note to an actor would mean maintaining an
actor→user mapping that does not exist, purely to send a notification.

**The rule:** if the thing being assigned is internal team work and needs to notify,
assign to `profiles`. If it can land on someone outside the team, assign to `actors`.
Don't unify them for symmetry — they have genuinely different populations.

Two related decisions from the same build, recorded so they aren't re-litigated:
- **One assignee on a note, not a link table.** The entire point is a single owner;
  the double-work problem is not solved by a note with three owners.
- **Quick capture defaults to unassigned.** Logging something you spotted is not the
  same as committing to build it. Claiming is one click in the table.
- **A converted note keeps a back-link** (`tasks.feedback_note_id`, 139, mirroring
  `meeting_note_id`). Without it a converted note simply goes `done`, and nothing
  distinguishes "we did this" from "this became work that is still open."

## Section 17 — An empty error body means the request was rejected before the SQL ran (2026-08-15)

Applying migration 139 through the Management API returned **HTTP 400 with an empty
body**, while the identical endpoint answered read-only queries fine. Half an hour
went into the SQL. The SQL was never the problem.

PowerShell 5.1's `Get-Content -Raw` decodes using the **system ANSI codepage**, not
UTF-8. The migration's comments contain em dashes; each came back as the three
mojibake characters `â€"`, which were then re-encoded into a payload the API rejected.
Reading the same file with `[IO.File]::ReadAllText(path, [Text.Encoding]::UTF8)`
applied it first try — `201 []`.

**The diagnostic that mattered:** a *SQL* error from this endpoint comes back as JSON
with a Postgres message and code (`{"message":"Failed to run sql query: ERROR: 42703
…"}`). An **empty** body means the request was rejected before execution — so
interrogate the payload, not the statements. Bisecting confirmed it: comments stripped,
the migration applied; every SQL construct in it (multi-statement, `begin/commit`,
dollar-quoting, comments) passed individually.

**Standing rule:** read any file whose bytes are going over the wire with an explicit
encoding. The same trap is why `ROADMAP.md` prints as `Come With â€" Platform Roadmap`
in a PowerShell console while being perfectly fine on disk — use the editor/Read path
for files, not `Get-Content`, when the content matters.

## Section 18 — A row limit you did not set is still a row limit (2026-08-15)

PostgREST answers with at most `max_rows` and says **nothing** when it truncates. On
`comewith-prod` that is **1000**. Every radio load in the dashboard was already past it —
1,327 future `ra_events`, 1,594 future `ra_artists`, 1,956 `sc_artist_cache` — and none of
them carried an `.order()`, so *which* thousand came back was arbitrary and could change
between two identical calls.

The damage was invisible in exactly the way Section 15's was. A shorter artist list looks
like a quieter month. Scanned artists read as unscanned, because a third of the cache was
never delivered. And it got worse the further out you looked: a window set two months ahead
could miss its own shows while the panel looked perfectly healthy.

**The rule:** a `.select()` with no `.range()` is a query with an undeclared cap. Page it,
and order by a **primary key** — ordering by a non-unique column (`event_date`,
`next_event_date`) lets a tie straddle a page boundary and silently skip or repeat rows.
The dashboard does this through `sbAll(build, pk)`; edge functions page inline. Same rule
for any new radio surface.

**The corollary, which is the more general lesson:** *every* bound in this pipeline was
written when the data was smaller than the bound, and every one of them became a silent
filter as the data grew — the 1000-row default, `dj-station`'s `.limit(160)`, DICE's
160-then-240 detail cap, RA's page budget. A cap is only safe if exceeding it is
**reported**. Where a cap is genuinely needed, return the count it dropped and surface it
in the UI; where it is not, page.

**And the trap that nearly shipped with the fix:** all three pull functions delete their
own source's rows before re-inserting, bounded only at the bottom (`gte(from)`). That is
correct exactly as long as every pull covers "today through the end of what we keep". The
moment a pull can be *narrower* — which is the whole point of aiming it at a window — that
delete throws away everything past the window it just fetched. Bound a delete at both ends,
or don't scope the fetch. Widening a read is never just a read change if a write is keyed
to the same range.

## Section 19 — The site owner is a row, not a role (2026-08-15)

Martin and Henry were promoted to full `master_admin`. The one thing they must not be
able to do is unseat Keith. That could not be expressed in the role system, because
`master_admin` was the top of it: the policy `"Master admin can manage all profiles"`
is `for all using (is_master_admin())` **with no WITH CHECK**, so any master_admin
could `PATCH /profiles?id=eq.<keith>` straight through PostgREST. There is no
role-change control in the dashboard at all — the UI was never the boundary.

**Decision:** ownership is a flag on a row — `profiles.is_owner`, exactly one — with a
`before insert or update or delete` trigger (`protect_site_owner()`) rather than an
RLS predicate. RLS on `profiles` would have to allow master_admins to write the table
generally; the trigger can allow the write and refuse the *specific column changes*
that matter, which is the actual requirement. Ownership can be **given** by the owner
and never **taken**.

**The vector that isn't obvious: `deleted_at`.** Under the 098 deactivation contract,
`is_admin()` / `is_master_admin()` / `user_can_access_module()` all treat a profile
with `deleted_at` set as no-role. So deactivating the owner locks them out completely
while `role` still reads `master_admin` — a guard that only watched `role` would have
looked correct and protected nothing. The guard covers `role`, `deleted_at`, `is_owner`
and `DELETE` as one unit.

**The deliberate hole:** the trigger exempts callers with no JWT (`auth.uid()` is null)
— service role, Edge Functions, and Management-API sessions. Without that exemption a
bad row could only be fixed by Supabase support. The honest consequence is that this
protects the **app**, not the **project**: the service-role key, an `SBP_PAT`, the
Supabase dashboard, GitHub and Netlify all still sit above it, and they are what
ownership actually means. Keep them Keith-only.

**Corollary, and it generalises:** a service-role Edge Function bypasses every trigger
guard, so it has to re-enforce the rule itself. `invite-user` now refuses the owner's
email — Supabase happens to error on an already-registered address, but that is its
behaviour, not our guarantee. Any future function that writes `profiles` inherits this
obligation.

**What this does NOT do**, stated plainly because it's easy to assume otherwise: the
two remaining master_admins can still remove *each other*, invite further
master_admins, read and write all company money, and flip `financials_released` on any
event. The 041–043 financial gate now applies only to Janelle and Liz. That was the
accepted trade of "same full access as me, except ownership".
