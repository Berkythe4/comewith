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

---

## Section 20 — A prior value nobody writes is not a trend (2026-08-15)

The Strategy board rendered `– no prior reading` on every live-computed card since it
shipped, and nobody noticed *why*. `v_kpi_dashboard` took `current_value` from
`coalesce(computed, snapshot)` but `prior_value` **only** from `v_metric_prior` — the
second-latest **hand-logged** reading. Nothing hand-logs net P&L or subscriber counts,
so the metrics that mattered most were precisely the ones that could never show a trend.
Prod introspection settled it: `metric_snapshots` held readings for `youtube.*`,
`instagram.*` and `tiktok.*` and nothing else.

**A computed value has no history unless something writes one down.** `snapshot_kpis()`
plus a 06:30 UTC cron (after the 06:00 YouTube pull) now writes `v_kpi_computed` into
`metric_snapshots` as `source='computed'`. 27 metrics build history the same way YouTube
already did.

**"Prior" is not one thing, so it lives in one place.** `v_kpi_prior` defines it per
metric: the **previous completed event** for event metrics, the **previous 5 uploads**
for recent content, and the reading **nearest 30 days ago** for everything else. The
fallback when history is shorter than 30 days is the **earliest** reading on record,
never the latest — falling back to the latest compares a number to itself and renders a
permanent, confident "no change".

**A delta must say what it is measured against.** The old card showed `▲ 47`, which
never said *47 since when*. `prior_basis` now travels with the number, so the card reads
"▲ +47 vs about 30 days ago".

**Corollary — a nightly snapshot destroys "as of".** Once something writes daily,
"last captured" is always today and means nothing. `v_kpi_changed` walks the history and
returns when the value last actually **moved**, so a card can say "unchanged since Jul 2"
instead of "updated today" every day forever.

---

## Section 21 — The board's categories come from the metric key, not a column (2026-08-15)

The rebuild needed six categories where `kpi_targets.workstream` had four (Radio was
seven cards buried inside Audience; Site was a section of its own). The obvious move —
`update kpi_targets set workstream = 'radio' where metric_key like 'radio.%'` — is a
**silent production break**: the deployed renderer only resolves four workstreams, so the
instant that migration landed, nine cards would have vanished from the live board with no
error anywhere. The database and the front end deploy on different clocks, and a
migration cannot wait for a merge.

**So membership is derived from the metric key prefix in the client.** `radio.*` → Radio,
`site.*` → Site, and so on. Three things fall out of it: the DB change disappears
entirely; a new `metric_key` lands in the right category with no migration at all; and
anything matching nothing renders under **Other** rather than disappearing.

**The general rule:** when a schema change would alter what already-deployed code
renders, prefer deriving the same fact in the client. Ship the two halves on the same
clock, or don't couple them.

---

## Section 22 — A lifetime average cannot answer "did the last one work" (2026-08-15)

Almost every money card was an average across **every completed event ever**:
`parties.net_pl`, `parties.sell_through`, `di.cost_to_raise`. Those numbers blend a 2024
loss into a 2026 win and barely move, so the board could not answer the only question
anyone actually asks — *was the last one better than the one before?*

`v_kpi_event_series` now gives per-event values, and the `*_last` metrics read event 1
against event 2. The lifetime averages **stay**, in the drill-down, where they belong.

**Cost to raise $1 is the Dance Infusion health metric** (Keith's call, 2026-08-15): it
is what a DI event is judged by, over raised-per-event or cumulative-to-MS, which are
outcomes rather than efficiency. It is `comparison = 'lte'`, so **colour follows the
comparison, not the sign** — a falling number is the good one, and any renderer that
assumes "up is good" gets this backwards.

**Same trap, content edition:** `youtube.avg_views` is lifetime views ÷ all videos. Its
own tooltip admitted it "is NOT a read on recent performance", and it was still the
Content headline. `content.avg_views_recent` (last 5 uploads vs the 5 before) replaced
it, and immediately said something the old card structurally could not: 103 against 274.

---

## Section 23 — "0" and "cannot be measured" are different claims (2026-08-15)

A blank hand-logged card and a genuine zero rendered identically on the old board. They
are opposite statements — one is *we have no idea*, the other is *we checked and it is
nothing* — and conflating them makes the board quietly dishonest. Source badges
(`live` / `api` / `logged`, with a stale flag past 60 days) now sit on the card face
rather than inside a hover, progress bars render only where a value **and** a target
exist (0% for an untargeted metric read as failing badly rather than untracked), and
funnel conversion rates return **null, not 0**, when the denominator is missing.

**The funnel is where this matters most.** `v_event_funnel` measures site exposure →
ticket click → ticket sold → attended. It reads empty, and the empty state had to
explain itself: the beacon started 2026-07-24; the only two events that ever carried a
`ticket_url` (Knicks G5, Come With 7-11) finished before that; and **neither upcoming
event has one set**. Showing "0 clicks" would say *nobody clicked*. The truth is *nothing
was ever measurable*, so the UI names the missing field and raises a dated alert instead.

**Where the ticket CTA actually lives, because it is not where you would look.** The
ticket link is on the **homepage** — `index.html` renders one per upcoming event plus the
`#nsBtn` hero button. `event.html` reads `v_public_recap` and is a **retrospective
archive page with no CTA at all**. So a ticket click is recorded with `path='/'`, and
attributing by page path returns exactly zero forever. Clicks are matched to an event by
comparing `link_url` to `events.ticket_url`, **query string stripped on both sides** —
one stored URL is a Partiful link carrying an `fbclid` that whole-string equality misses.

**The operational consequence:** the beacon cannot backfill. A `ticket_url` added after
promotion starts loses every click that came before it, so it has to be set *first*.

---

## Section 24 — "Blocked" and "nothing found" are different answers (2026-08-18)

> ⚠ **Partly superseded the same day — see §27.** The reasoning below stands; the
> *conclusion drawn from it about Craigslist* was wrong. Craigslist RSS is dead, but
> Craigslist itself is scannable through the JSON endpoint its own search box calls, and
> Gear Watch now uses it. Do not cite this section as "Craigslist cannot be pulled".

Gear Watch was designed with Craigslist as a source. Testing it against the live endpoint
returned **HTTP 403 "Your request has been blocked"** on every search-RSS path, under a
browser user-agent and a bot one alike, while the Craigslist homepage returned 200 from
the same machine — proof of a deliberate block rather than a network problem. The HTML
search page returns 200 and contains **zero** listing markup; it is a JS shell that loads
results from an internal API.

This is §-worthy not because a source died, but because of what a dead source *renders as*
if you let it. A scan that reports "0 listings" when it was actually blocked tells the
reader **"your stolen gear is not being resold"** — the single most harmful sentence this
system could produce, and indistinguishable from good news. Same failure the Bandcamp
`fuzzysearch` endpoint produced when it answered HTTP 200 with `{"error":true}` and every
track came back "not on Bandcamp".

So the rule, now applied in both places:

- **Validate the payload, not the status.** `parseCraigslistRss` throws on a body that is
  not a feed, rather than returning `[]`.
- **A source that fails is reported as FAILED**, per source, in the digest and the panel —
  never folded into the total.
- **A source that is known-dead is reported as DISABLED, not FAILED.** Craigslist sits
  behind `CRAIGSLIST_ENABLED` (default off) with the 403 evidence in the header comment.
  A source that cries FAILED three times a day forever trains the reader to skim past
  failures — which defeats the point of reporting them at all.
- **What cannot be automated is handed back as a manual link.** OfferUp and Facebook
  Marketplace have no public API and prohibit scraping; the digest ends with saved-search
  links so they stay a 60-second human check instead of a scraper that rots.

## Section 25 — pg_cron gets a service-role bearer from vault, not a stored JWT (2026-08-18)

`014_cron.sql` deferred scheduled campaign sends in Phase 10 with an honest note: pg_cron
cannot construct an admin JWT, and the two ways out were "cron-secret header" or "vault +
service_role token in pg_net headers". Every cron job since has been inline SQL, so the
question stayed open for two months. Gear Watch needed an edge function on a schedule, so
it is now settled.

**The pattern:** `public.gear_watch_kick()` is `security definer`, reads the function URL
and the service-role key from `vault.decrypted_secrets` **at call time**, and posts via
`net.http_post`. The edge function accepts a service-role bearer OR an admin JWT — the
same door `pull-ra-market` already opened for exactly this purpose.

Three properties worth keeping in any repeat:

1. **The key never enters git.** The migration creates the *caller*; the secrets are
   stored once by hand with `vault.create_secret`.
2. **A missing secret is a documented no-op, not a failure.** With no secret the function
   writes `skipped: secrets not set` to `gear_watch_config.last_status` and returns. A
   cron job that errors silently every eight hours is worse than one that says why it did
   nothing.
3. **`search_path` must include `net`.** pg_net's functions are not in `public`, and a
   `security definer` function with a pinned search_path will not find `http_post`
   otherwise.

## Section 26 — A field you default is a signal you invented (2026-08-18)

The Craigslist parser had no per-listing location, so it filled in the obvious-looking
constant: `location: "New York (craigslist)"`. The feed is the NYC board, after all.

Then the scorer awarded **+20 for "local"** — and every listing on the board became local,
including the Philadelphia one in the fixture, which scored 55/100 and would have been
emailed as a candidate. The default did not read as missing data. It read as *evidence*.

The fix is to parse the neighbourhood out of the title's trailing parenthetical and leave
it **null** when there isn't one, because null scores nothing while a guess scores twenty.
The NYC board carries plenty of Philadelphia, Connecticut and New Jersey posts.

Generally: **a placeholder that flows into a computation stops being a placeholder.** It is
safe to default a field that is only ever displayed; it is never safe to default one that
is scored, filtered, summed or compared. Prefer null and lose the signal.

Caught by running the pipeline on a fixture before deploy — which is the argument for
having a fixture at all, given the live source was unreachable.


## Section 27 — "I couldn't reach it" is not "it can't be reached" (2026-08-18)

§24 concluded, from a real test, that Craigslist could not be scanned. Keith pushed back:
his brother had built a Craigslist car-listing puller with Claude. He was right and §24's
conclusion was wrong — and the gap between the evidence and the conclusion is the lesson.

**What the test actually proved:** the `format=rss` search paths return 403 to this
machine. **What was concluded:** Craigslist cannot be pulled. Those are different claims,
and the second does not follow from the first. One dead door is not a locked building.

**What works, verified 2026-08-18:** the internal JSON endpoint the site's own search box
calls —

```
https://sapi.craigslist.org/web/v8/postings/search/full?batch=<areaId>-0-360-0-0&cc=US&lang=en&searchPath=sss&query=<q>
```

200, with 47 live NYC results for "cdj". This is precisely the technique already written
down in CLAUDE.md for Bandcamp — *use the endpoint their own front-end calls* — and it was
not tried before declaring the source impossible. **When one access path fails, open the
site's own network tab before writing the obituary.**

Three decoding traps, all found by running it against live data rather than reasoning
about it:

1. **The payload is delta-encoded.** Posting ids accumulate from
   `decode.minPostingId`; dates are seconds after `decode.minPostedDate`; `price: -1`
   means "no price stated", and storing that as a number makes every unpriced listing the
   cheapest thing on the board.
2. **The geo string carries TWO indexes into TWO arrays** —
   `"<locationIdx>:<descriptionIdx>~lat~lon"`. Using the first for both put a Schenectady
   listing in Bushwick and handed it a local bonus.
3. **A query with no results returns `decode: 0`** — the number, not an object. Validating
   `!decode` as a bad payload reported *"nothing listed"* as *"we were blocked"*, which is
   §24's own mistake pointed the other way. Both directions are harmful; both are now
   tested.

And one that was not a Craigslist problem at all: geo terms were matched by substring, so
`"ny"` matched **albany**, and every upstate listing scored as local. Short tokens are
exactly the ones that need word boundaries. Now `isLocalTerm()`, with regression tests.

The canonical posting URL is `https://www.craigslist.org/view/d/<slug>/<token>`, built
from the `[6, …]` slug and the `[13, …]` token; the older
`/<area>/<cat>/d/<slug>/<id>.html` form 404s unless the category segment is exactly right.

**The habit worth keeping:** when a user says "I know for a fact this works", that is
evidence — usually of a path not tried. Test it before defending the earlier finding. Four
of the six bugs in this feature were found by running it against something real; none of
them were visible from the code.

---

## Section 28 — A selection the filter cannot see is a mass update with no blast radius (2026-08-19)

The Expenses tab kept selected row IDs in `expDash.sel` and never reconciled them
against the active filter. Select 30 rows, narrow the filter to 5, apply a bulk
category change — and all 30 were written. Twenty-five of them were invisible at the
moment they changed. Keith found it the way these are always found: by discovering
work he had just done had been silently overwritten.

The bulk bar was already careful in the way that is easy to be careful: every button
named the count it would affect. That is worth nothing when the count is right and
the *membership* is wrong.

**The rule: a selection may only ever contain rows the current view shows.** Prune on
every render, and say how many were dropped. A selection that quietly shrinks is only
marginally better than one that quietly grows — both are the UI knowing something the
person does not.

`pruneSelection(sel, visibleRows)` now enforces this for Expenses and Income. Anything
that grows a bulk-edit surface should call it too.

---

## Section 29 — `position: sticky` needs a scrollport that actually scrolls (2026-08-19)

Sticky table headers were added and did nothing. The CSS was right; the containing
block was not.

`.main` carries `overflow-x: auto`. Per spec that coerces `overflow-y` from `visible`
to `auto`, so `.main` became the nearest scroll container for every table inside it —
but `.main` has no height limit, so it never scrolls. The window does. A sticky
element resolves against a scrollport that never moves, so it never moves either.

Two things worth carrying:

1. **`overflow-x: auto` is never only horizontal.** It silently makes the element a
   scroll container in both axes, which changes what every `position: sticky`
   descendant is measured against.
2. **Fixing the shell has a blast radius.** Making `.main` the real scroller broke
   fourteen `window.scrollY` / `window.scrollTo` call sites that preserved scroll
   position across re-renders — they would have read 0 from a window that no longer
   moves, silently. They now go through `scrollPos()` / `scrollToPos()`, which ask
   whichever element is actually scrolling. Search for those before changing a scroll
   container again.

The first attempt at this shipped the CSS alone and was reported back as still broken.
Setting the property is not the same as verifying the behaviour.

---

## Section 30 — Reportability is a stored decision, not a category (2026-08-19)

`v_contractor_1099` was first built on `category = 'Contractors'`. It under-reported
two payees who had crossed the $600 threshold — Janelle Sochet showed $700 against
$900 actually paid, 19th & 7th showed $900 against $1,800 — because the rest of their
money sat in 'Marketing' and 'Production'. The threshold is measured per payee per
calendar year across all service payments; it does not care about our buckets.

Deeper than the query bug: **the ledger does not hold the facts the answer depends
on.** Entity type, goods-versus-services, whether a payment was a reimbursement — none
of it is in an expense row. Inferring it from category is what produced the wrong
numbers in the first place.

So the decision is stored on the actor (`tax_1099_status`), set by a person, with
'undecided' surfacing on a review list rather than being guessed. A list that says it
is incomplete is worth more than a confident wrong one.

Corollary found the same day: a payee that exists only as vendor text has nowhere to
record a decision, so it would sit on the review list forever with no action available
that could clear it. Those now read `'no vendor'` rather than `'undecided'` — a
different problem needing a different fix should not share a label.

---

## Section 31 — A payment rail is not a counterparty (2026-08-19)

Migration 158 seeded **Venmo** as an actor with a matching alias rule. Every Venmo
payment therefore resolved to a single payee called Venmo, regardless of who received
the money. It surfaced as a $650 false positive on the 1099 review list — two charges
to two different people for two different things, aggregated because they shared an
app.

The false positive was the harmless direction. The dangerous one is the same mechanism
running the other way: one person paid repeatedly through Venmo, their real total
hidden inside a bucket named after the rail.

**Vendor text from a bank statement describes how the money moved, not who received
it.** `Venmo`, `Sq *`, `Ubr `, `In *` are all descriptors. Checked for the same pattern
across PayPal, Zelle, Cash App, Square and Stripe — no other rail had an actor row.
Check again before seeding vendor aliases in bulk.

---

## Section 32 — A model that requires a parent will meet something with no parent (2026-08-19)

`event_photos.event_id` was NOT NULL, which was fine for as long as every photo came
from a show. It failed on the first press shoot: Keith paid a photographer to
photograph *him*, and there is no event those images belong to.

The tempting fix — create an event called "Press shoot" — would have put a photo
session in the events list, in the pipeline, and in the P&L next to real shows. The
schema would have stayed clean while the *meaning* of "event" quietly rotted.

A photo now hangs off a subject as well as, or instead of, an event, with a CHECK that
it has at least one so nothing can be uploaded into nowhere.

Shipped alongside it, and the more important half: **`is_public` defaulted to `true`.**
Every upload went live on the public site the moment it finished. For a library you
pick *from*, publishing is a decision, not the resting state. Default flipped to
false — and existing rows deliberately left alone, because silently un-publishing live
gallery images would have been a worse surprise than the default was.

---

## Section 33 — Blue Sky is a stage, not a new table (2026-08-19)

Keith needed somewhere to put gigs he wants but does not have, that later either
become real bookings or get dropped. The instinct is a new `prospects` table. The
right answer was already in the schema: `events.stage` permits `'idea'`.

Two numbers make it useful — `expected_revenue` and `confidence` (0-100) — and
`v_pipeline` returns the product. Ten $1,000 gigs at 30% is $3,000 of forecast, not
$10,000, which is the entire point of writing them down instead of hoping.

Because it is the same row throughout, a Blue Sky event is **promoted** (stage →
confirmed, book real income) or **dropped** (status → cancelled) without moving
anything between tables. The row survives either way, so the hit rate becomes
measurable rather than anecdotal. Nothing reaches the P&L: `v_pl_monthly` reads income
and expenses, and a Blue Sky event has neither until it is real.

Related gap closed at the same time: `v_event_money.missing_revenue` only looked
backwards — past events carrying costs with no fee. An upcoming show with no money on
it stayed invisible until it became a past show with no money on it, which is too late
to do anything about. `v_pipeline.needs_revenue_estimate` is the forward-looking
version. It currently returns all 8 upcoming events.

---

## Section 34 — Cost is incurred when it is agreed; cash moves when it is paid (2026-08-20)

161 gave income three states because most of what Come With earns is agreed long
before it is paid. Costs had no equivalent, so a DJ booked in August and paid in
October could only be recorded two ways, both wrong: leave it out (understating what
the event cost) or enter it as a normal expense (understating cash by pretending the
money had already gone). 177 gives `expenses` the mirror of income's states —
`accrued` -> `invoiced` -> `paid`.

**The split is only worth anything if each view picks a side.** Adding the column is
the easy part; deciding what counts where is the decision:

| view | basis | why |
|---|---|---|
| `v_pl_monthly` | all three | a cost is incurred when the obligation is |
| `v_cash_position` | `paid` only | otherwise a payable silently drains the float |
| `v_contractor_1099` | `paid` only, in the year **paid** | a 1099 is cash-basis |
| `v_capital` | `paid` only | nobody personally carried a bill nobody has paid |
| `v_tax_year` | `paid` only, `committed_unpaid` shown separately | filed on cash |

The 1099 and tax-year change is the subtle one: the year is now
`coalesce(settled_at::date, date)`, not `date`. A fee accrued in December and paid in
January belongs on next year's form. Every existing row backfilled to `paid` with
`settled_at` null, so the coalesce falls back to `date` and **every number this
database reported the day before was byte-identical the day after** — proved by
snapshotting 396 view keys, applying inside a transaction, diffing, and rolling back.
Nothing moves until someone records a payable on purpose.

Two smaller things fell out of it, both cases of a check that stopped making sense:

- **"No cash source" must not fire on an unpaid bill.** A bill that has not been paid
  has not left an account, so asking which account it left is nonsense — and the queue
  would have grown every time someone planned ahead properly. `v_cash_position`'s
  `unknown_src` and the Expenses tab chip both filter to `paid` now.
- **The settle dialog asks which pot paid it.** Without a `cash_source` the payment
  never draws down the float and the row lands straight back in the unknown-source
  queue, so the one moment the answer is actually known is the moment to ask.

The event Money panel gained the surface for all of it: a summary strip that separates
earned from banked, per-row status with a `settle`/`pay` action, and **"Link existing"**
— because 203 of 266 expenses were marked `event_na` and re-typing one to attach it to
an event is how you end up with the charge recorded twice.

---

## Section 35 — A forecast is not a fourth status (2026-08-20)

177 gave income and expenses three states each, but all three are **commitments**:
`accrued` already means "we have agreed to this", which is why the P&L counts it. There
was still nowhere to put the line before that — the DJ we intend to book, the bar take we
expect. Keith asked for it, and the obvious implementation was wrong.

**Adding `planned` to `expenses.status` would have been three lines of DDL and a permanent
hazard.** Every view that sums expenses or income *without* a status filter would silently
start counting speculation as fact: `v_pl_monthly`, `v_tax_year`, `v_event_summary`, the
KPI views, and the 011 / 022 / 026 / 043 / 060 views. One missed filter and a guess is in
the P&L — the exact failure §33 exists to prevent. **A separate table cannot leak, because
nothing that computes the P&L reads it.**

`budget_lines` was already the right table and nothing had ever written to it. It has
carried `(event_id, scope, category, direction, planned_amount)` since 026, its `scope`
check has permitted `'event'` from the start, it has an admin-only RLS policy — and
`v_pl_monthly_vs_budget`, the one view that turns budget into a P&L column, filters
`scope = 'period'`. Event-scoped lines are invisible to it **by construction**, not by
remembering to filter. The 37 rows in it today are all period rows and were untouched.

**Realising a forecast stamps it, it does not delete it.** When the DJ is actually booked
the line gets `realized_at` plus the id of the row it became, and stops counting as
forecast — a forecast that keeps counting after it comes true is an overstatement with a
good excuse. Keeping the row is what makes the estimate comparable to the outcome, which
is the only way forecasting ever gets better. Verified on prod: two lines in, P&L
byte-identical; commit one at $850 against a $900 plan; forecast drops to zero, the P&L
picks up exactly $850, and the −$50 variance is still readable on the budget line.

### The panel that reads like the P&L

Keith's words were "the main view format is too hard to read". The list of rows was fine
for five rows and useless for thirty. It is now the **company P&L's own table** — same
`pl-table` classes, same section bands, same carets, same Expand all / Collapse all —
scoped to one event, with four columns that read left to right as the life of a dollar:

**FORECAST** (planned, never in the P&L) · **BOOKED** (what the P&L counts, paid or not) ·
**SETTLED** (of that, money that moved) · **vs plan**.

Two things that only showed up because the render path is now under test:

- **A realised forecast was still counted**, on top of the row it became. `v_event_forecast`
  filters `realized_at`; the panel reads the *table*, and a filter that lives in only one of
  two places is a filter that does not exist.
- **Only forecast and settled rows had an editable amount**, so a committed cost — the row
  you most want to correct — was read-only. Booked is now the editable figure for anything
  real; Settled mirrors it, because settled is decided by the status, not typed.

`scripts/test_money_panel.mjs` executes the render path against fixtures and asserts the
rollup arithmetic. It exists because `node --check` proves the file parses and cannot tell
you a function the renderer calls was deleted — which is exactly what happened while
rebuilding this, twice.

---

## Section 36 — The audit, and the one pattern under most of it (2026-08-20)

A full sweep of the ecosystem: 106 tables, 177 foreign keys, 56 views, 29
security-definer functions, and a row-level census of every link the architecture
implies. Most of it was in good shape — RLS is on everywhere with a real policy on
every table, and no view is anon-readable that should not be. What it did find fell
into three groups.

### 1. "Needs a link" with no way to say "has none"

`expenses.event_na` was the only place in the schema where an intentional blank could
be *declared*. Everywhere else it looked exactly like an oversight:

- 7 expenses had no payee actor because their "vendor" is **Food & Beverage**, **Gas /
  Transportation**, **Presents** — categories, not businesses. There is no payee and
  never will be.
- 23 social posts had no event because they are Content Creation planning slots.
- 123 guests had no actor, which is correct: an attendee is not a business contact.

Every one of those would have been flagged forever. **A queue that can never reach
zero is a queue people stop reading**, which costs you the real finding sitting in it.
179 added `income.event_na`, `expenses.payee_na` and `social_posts.subject_na`; 180
added a general `data_health_waivers` so a future check can be dismissed **with a
reason** rather than needing a migration.

The same instinct fixed three checks before they shipped. "Tickets not linked to a
guest" flagged 9 rows that are deliberately **lump** rows holding reconciled Dance
Infusion totals — narrowed to single-seat tickets, 9 → 0. "Guests not linked to an
actor" became "guests whose *email already matches* an actor", 123 → 1. Getting a
check to be quiet when things are fine is most of the work of writing one.

### 2. SECURITY DEFINER + PUBLIC EXECUTE

**A definer function runs as its owner, and Postgres grants EXECUTE to PUBLIC unless
told otherwise.** On this project `authenticated` includes every radio listener who
ever signed up. `autolink_data` — which writes `actor_roles`, `guests`, `expenses` and
`income` — shipped in 181 reachable by all of them. Mine, one session old. So had
`snapshot_kpis`, which writes the KPI history the strategy board reads and which has no
other source.

Every other definer function in `public` was already fine, by one of two routes:
an internal guard, or EXECUTE granted only to `postgres` + `service_role`. 183 closed
the three that were not, allowing a **null `auth.uid()`** so pg_cron and the service
role still work — the same exemption 140's `protect_site_owner()` makes, for the same
reason.

`post_apply.sql` now checks for the whole class. Its first draft FAILed on six
functions, all false positives, and had to be narrowed three times: trigger and
event-trigger functions cannot be called directly; read-only RLS predicates are
supposed to be callable; and establishing the caller counts **however it is done** —
`actor_set_task_status` guards through `current_actor_id()`, which is the DJ
scoped-link pattern and entirely legitimate. What survives is the real thing: *it
writes, anon or authenticated can call it, and it never asks who is calling.*

### 3. Automation that leaves a receipt

`autolink_data()` is **dry by default** and only makes links already implied by data
somebody entered on purpose — an actor who is the payee on two expenses **is** a
vendor; a donor name that matches an actor's display name **exactly** is that actor.
No fuzzy matching, ever: §31 exists because one bad merge collapsed unrelated payees
into a single actor.

Every run of either function writes a row to `data_health_runs` with a per-action
summary, and the panel renders that history. **A nightly process that mutates the actor
graph and leaves no receipt is not automation, it is drift with a schedule.** The
preview and the apply are the same function with one flag, because a preview produced
by different code than the apply is not a preview.

What it actually did on first run: 11 actor roles inferred from relationships that
already existed, 1 guest linked by exact email, 8 donations linked by exact name (in
179), 16 blank event stages filled from the status each event already carried. Zero
risky matches — the 7 category-shaped "vendors" were correctly left alone.

---

## Section 37 — The ledger was public, and every check said it was fine (2026-08-20)

Found while answering a question about private SoundCloud links. Anonymous callers
could read **29 expense rows with payee names and amounts, 59 ticketing rows, 16
donations with donor names, 12 sponsorships and 9 income rows**. Closed in 185.

### The bug

043 put this on all five money tables:

```sql
for select using (can_see_event_financials(event_id))
```

and defined the helper as:

```sql
select public.is_master_admin()
    or (p_event_id is not null
        and exists (select 1 from events e
                     where e.id = p_event_id and e.financials_released));
```

**The second branch asks what has been released and never who is asking.** RLS
policies apply to role `public`, which includes `anon`, so the moment an event had
`financials_released = true` its money became world-readable. The intent — *staff
see an event's money only once it is released* — was half-expressed: the release
condition was written down, the staff condition was assumed. 185 adds
`public.is_admin()` to that branch and nothing else changes.

### The part worth remembering: two checks that couldn't see it

**1. A grant check cannot see an RLS leak.** `post_apply.sql` verifies the five
financial VIEWS are anon-revoked, and they are — that check has been passing
honestly for months. But the leak was on the underlying TABLES, which carry an anon
grant from 013's default privileges and rely on RLS alone. Different mechanism,
invisible to a grants query.

**2. Every REST spot-check in this repo was run with an empty API key.** `.env` has
no `SUPABASE_ANON_KEY` — the variable is `SUPABASE_PROD_PUBLISHABLE_KEY`. An empty
apikey returns **401 for everything**, public or not. So the check printed a column
of `401`s, which is exactly what a passing run looks like. Every anon verification I
reported this session — after 177, 178, 180 — was vacuous. The invariant happened to
hold for the views; nobody had established that, we had just been told it in a
convincing font.

Both are the same failure: **checking a proxy for the invariant instead of the
invariant.** The same shape as the two checks in `post_apply.sql` that cried wolf on
their first run, and the same shape as the Bandcamp endpoint that answered HTTP 200
with `{"error":true}` and made every track read as "not available".

`scripts/check_anon_exposure.py` is the fix. It reads a known-public view first and
**refuses to run** unless the key actually works, then reads the **body** of every
sensitive table — because on a table `200 []` is correct and `200` with rows is a
breach, and the status alone cannot tell you which one you have.

### And a smaller one, same sweep

`v_equipment_roi` (purchase prices, revenue per item), `v_mailing_list_health`
(list size) and `v_metric_prior` (the KPI scoreboard) were anon-readable with
nothing public reading them. Revoked in 186. `v_kpi_targets_current` was left alone
and flagged instead: `tools/visualizer.html` reads it with **no sign-in at all**, so
revoking would have broken a working tool silently — which is the thing this whole
section is about.

## Section 38 — An inline editor must not narrow the field it edits (2026-08-21)

The Social Calendar's timeline view became a list view built to the events-list
spec: the same `data-table`, the same banding, the same chip filters, and the four
fields that actually move a post along editable in place.

Two of those four fields could not take the obvious editor, and both would have
failed the same way — **silently, on save, destroying data the user could see a
second earlier**.

- **`social_posts.channels` is an ARRAY.** A single `<select>` bound to it turns a
  post that goes out on Instagram *and* TikTok into a post that goes out on one of
  them, the moment anybody touches that cell for any reason. The cell renders one
  removable chip per channel and a separate `＋` select that only offers channels
  the post does not already have. Adding is additive; removing is explicit.
- **`social_posts.content_pillar` is FREE TEXT.** The modal has always been a text
  box, so the column holds whatever anyone typed. A select over a hardcoded list of
  pillars silently rewrites any value not on that list to whichever option the
  browser picked instead. The options are **derived** — the five we always use, plus
  every distinct value on file — so an off-list `takeover` renders selected and
  survives an edit to the row's stage.

The general rule: **an inline editor is a promise that the widget can express every
value the column can hold.** Where it cannot, either widen the widget or leave the
field to the modal. A `<select>` over an array or over free text is not an editor,
it is a delete button with a friendly label. Same family as §26 (never default a
field that feeds a computation) — a lossy control and an invented placeholder both
replace real data with something that merely looks reasonable.

Regression-tested in `scripts/test_content_center.mjs`: a two-channel post keeps both
chips and is not re-offered a channel it already has; an off-list pillar comes back
`selected`. Both checks were mutation-tested — breaking them fails the suite.

While there: the timeline view is **gone**, not kept alongside. Keeping it would have
left two views called "List", and the list strictly supersedes it — everything the
timeline showed, plus editing, plus filters. Views: Calendar, List, Board.

## Section 39 — A deferral has to carry the check that would kill it (2026-08-21)

186 revoked three anon-readable internal views and deliberately left a fourth,
`v_kpi_targets_current`, with this written into the migration:

> tools/visualizer.html reads it ANONYMOUSLY - it has no sign-in at all - so
> revoking would break a working internal tool without warning.

**Every clause of that was wrong.** `tools/visualizer.html` line 7 loads
`/staging/guard.js` and line 60 imports its Supabase client from it — the same admin
gate every `/staging/` page uses; with no session it redirects to `/dashboard.html`.
And the tool could never have worked signed-out anyway. Asked as the public with a
key that actually works:

| what the visualizer reads | as anon |
|---|---|
| `metric_definitions` | `200 []` — RLS returns nothing, so the metric picker is empty |
| `v_data_points` | `401` — already revoked, nothing to chart |
| `v_kpi_targets_current` | **`200`, real rows** |

Two of its three sources were already closed. The grant was not holding a working
tool up; it was publishing every KPI target we have set. Revoked in **187**.
`authenticated` keeps SELECT, verified before and after, so nothing changed for a
signed-in admin.

### The part worth remembering

This is §37's failure one level up. There, two checks tested a **proxy** for the
invariant. Here, the decision not to check rested on a **premise stated in prose**
that nobody had ever executed — and prose in a migration header reads exactly as
authoritative as a verified fact, because it sits in the same file.

So: **when you defer a security fix, write down the command that would prove the
premise wrong, and run it.** "The tool reads it anonymously" is a claim about a live
system; it takes one `curl` with a real key. Three lines of investigation would have
closed this on 2026-08-20 instead of parking it as a decision for Keith.

Corollary, and the reason this was cheap to find: **the leak stayed on a list.** The
carryover named it as item 2 and `check_anon_exposure.py` carried it in `PUBLIC_OK`
with a comment saying why. A known exception that is written down where the sweep
will show it is survivable; the same exception held only in someone's head is not.
It has now moved to `MUST_BE_EMPTY`, and **nothing in `public` is anon-readable
except the public site feed.**

## Section 40 — An invoice is not revenue (2026-08-21)

Invoicing shipped across 188–194. Every design decision in it follows from one
rule, written into 188's header because it is the thing that will be forgotten
first:

**The income row is the revenue. An invoice is the document that asks for it.**

Nothing in the invoicing schema sums into the P&L. If an invoice also booked
revenue, every invoiced job would count twice — once as the accrual and once as
the bill — and the error would look like growth.

```
income row (accrued)   the money we are owed          <- the P&L counts this
invoice + lines        the document asking for it     <- counts nothing
invoice_payments       cash arriving against the doc  <- still not revenue
income row (received)  settled, when the invoice is    <- the cash date
                       paid IN FULL
```

161 gave income three states — `accrued -> invoiced -> received` — and the
middle one had been unreachable for months because nothing could produce an
invoice. Every income row on prod was `accrued` or `received`. Sending an
invoice is now what moves rows to `invoiced`, and paying it in full is what
moves them to `received`, carrying the payment's date and method as the cash
date and source.

**Partial payments live on the INVOICE, not the income row.** An income row has
one `amount` and one `settled_at` and cannot be half settled. A deposit is
recorded against the invoice; the rows behind it flip only when the balance
reaches zero. That is why `invoice_payments` exists at all rather than a
`paid_amount` column.

### Totals are computed, never stored

`subtotal / discount / tax / total / paid / balance` live in views over the
lines and payments. A stored total is a number that can disagree with the rows
underneath it, and on an invoice that disagreement is the difference between
what you charged and what you can prove you charged.

The consequence is that the arithmetic exists in **two** places — `v_invoice_totals`
for the dashboard, `computeTotals()` in `template.ts` for the PDF and the
client's web page — and they must agree forever. So:

- both were checked against prod inside `BEGIN..ROLLBACK` on the case designed
  to pull them apart (a per-line discount **and** an invoice-level discount, a
  non-taxable line, tax on top of a discount): subtotal 2270, discount 227,
  taxable base 1935, tax 171.73, total 2214.73 — identical to the cent;
- `scripts/check_invoice_sql.sql` and `scripts/test_invoice.mjs` hold the two
  halves, so changing one without the other fails a test rather than a client's
  arithmetic.

The invoice-level discount is **apportioned across the taxable base in
proportion**, so turning tax on can never charge tax on money the client is not
paying. Rounding happens once per line and once on the tax, so the printed lines
add up to the printed total.

### Two smaller rules from the same build

**Tax off and tax at 0% are different claims.** The row only prints when
somebody has asserted it. Same family as §23 — never render a blank as zero.

**Snapshot what the document says; look up what the document needs.** The
bill-to name and email are copied onto the invoice at creation, because a
document that silently rewrites itself when a contact is renamed is not a record
of anything. The payment block is read fresh each render, because if you change
banks you want the current details, not last year's.

## Section 41 — Grants are checked before RLS (2026-08-21)

188 tried to make `invoice_settings` master-admin-only and did it twice:

```sql
create policy invoice_settings_master on public.invoice_settings
  for all using (public.is_master_admin());          -- correct
revoke all on public.invoice_settings from anon, authenticated;   -- fatal
```

The revoke made the Payment details screen unopenable **by everyone, including
the owner**. The screen is ordinary dashboard code reading the table as the
signed-in user; with no grant, PostgREST refuses before RLS is ever consulted.

Two things worth keeping:

1. **Revoking a grant does not make a table "admin only". It makes it
   nobody-only.** RLS is what decides *who among the grantees*; the grant is what
   decides *whether the role can reach the table at all*. They are not two
   strengths of the same dial.
2. **The failure presents as a problem with the CALLER.** The error is
   `permission denied`, so the dashboard reported "Payment details are
   master-admin only" and Keith went and checked a role that was correct all
   along (`berky@comewith.org`, master_admin, `is_owner`). A permission error
   should name what was refused, not guess why.

Fixed in 191: `grant select, update to authenticated`, policy unchanged. Proved
on prod inside `BEGIN..ROLLBACK` by impersonating real accounts — berky 1 row,
henry 1 row, janelle (sub_admin) 0 rows — and the anon sweep still reports the
table blocked.

The shape generalises past this table: **if a screen cannot read something, check
the grant before you check the policy**, because the grant failure is the one
that lies about whose fault it is.

## Section 42 — A screen that changes shared chrome must have it reset by the opener (2026-08-21)

`openKpi()` is one modal reused by every screen in the dashboard. The invoice
editor is the only caller that changes its *chrome* rather than just its body: it
widens the modal, and it hides the submit button because it has no single save
action.

Both leaked into whatever opened next, and both were reported as separate bugs
by Keith days apart:

- **width** — Record a payment and Send opened at 1040px, laid out for a form of
  two fields;
- **submit** — `$('kpiModalSubmit').style.display = 'none'` survived, so after
  opening *any* invoice, Send, Record a payment and Payment details all rendered
  **with no submit button at all**. "There is no send button" was this.

The fix is not "remember to put it back". It is that **`openKpi` resets every
property any screen is allowed to set** — width, submit visibility, and the new
optional third button — and `closeKpi` does the same. A caller may change the
chrome; a caller may not be trusted to restore it, because the caller does not
know what opens next.

The test asserts all three resets and was mutation-checked: re-introducing
either leak fails the suite. This is a guard rather than a fix precisely because
it bit twice.

Worth noting what did *not* catch it: `node --check` passed happily both times.
The file parses perfectly with a leaked inline style. Only running the screens in
order, or asserting on the reset, shows it — the same lesson as §38's note that
the money panel's own test caught a helper defined outside its region while the
syntax check saw nothing wrong.

## Section 43 — A generator that decides and writes in the same breath can't be reviewed (2026-08-21)

`generate_day_of_tasks` read the event, chose the steps, and inserted them in one
transaction. The only interaction it offered was a `confirm()` describing what it
was *about* to do in prose. You found out what had actually been created by
looking at the task list afterwards.

Making the steps editable first was not a UI problem. It was that **the decision
and the write were the same function**, so there was no moment at which the
proposal existed and could be shown to anyone. The fix is the split:

- `plan_event_tasks(event, set)` — decides, returns rows, writes **nothing**
- `generate_day_of_tasks` — loops the plan and inserts

The generator is now a thin consumer of the planner, which matters more than it
looks: the calendar gap panel **already** re-implemented the generator's filter
client-side, to promise a count that matched what the button would create. That
was two copies of the rules. A third — a preview that decided separately from the
thing that wrote — would have drifted the first time either changed, and drifted
*silently*, because a preview that over-promises looks exactly like a preview
that's right.

The corollary is worth stating: **a review step replaces a confirm dialog, it
does not sit next to one.** Both entry points dropped their `confirm()`. Asking
"are you sure?" before showing someone the eleven things they're agreeing to was
never really a question.

## Section 44 — If people can rename a thing, its identity cannot be its name (2026-08-21)

Template-generated tasks were matched back to their template by
`lower(trim(title))`, in two places: the generator's "already exists" suppression,
and the gap panel's "N of M workflow steps missing". That was survivable for as
long as the titles were machine-written and nobody could edit them.

The whole point of the review flow is that you *can* edit them — rename
"Confirm vendor arrival window" to "Check with Sal re: 6pm drop" as you create it.
Under title-matching that rename means the step reads as missing **forever**, and
the next run re-creates the original alongside your renamed one. The feature would
have quietly corrupted the thing that measures it.

So `tasks.template_id` carries the link, with the title kept only as a fallback
for rows created before the change (34 backfilled). The general rule:

> The moment you let a user edit a field, that field stops being available as a
> join key. Identity has to move to something they can't type.

Same shape as the `station_no` / `edition_seq` split (§ radio): a number a human
reads is not the number the system keys on.

## Section 45 — `revoke ... from anon` on a new function is a no-op (2026-08-21)

Migration 195 shipped with:

```sql
revoke all on function public.plan_event_tasks(uuid) from anon;
```

The post-apply check reported **FAIL** — anon could still execute it. Postgres
grants `EXECUTE` on a newly created function to **`PUBLIC`**, and `anon` inherits
that. Revoking from `anon` removes a grant `anon` never separately held.

The correct form revokes PUBLIC and grants the role you actually want back:

```sql
revoke all on function public.f(args) from public, anon;
grant execute on function public.f(args) to authenticated;
```

This is the same fact 183 acted on for `snapshot_kpis` (`from public, anon,
authenticated`) — the knowledge was in the repo and still didn't survive contact
with a new migration, because the `from anon` form *looks* correct and fails
silently. Both functions guard themselves with `is_admin()` anyway, so nothing
was ever exposed; what was exposed was the gap between the check and the review.

**Reading the SQL did not catch this. The post-apply grant check did.** Which is
the argument for running it every time rather than when a migration "looks
grant-ish" — this one wasn't about grants at all, it was about a new function.

## Section 46 — A migration that DROPS something must ship with its UI, not before it (2026-08-21)

196 dropped `task_templates.event_type`. The deployed dashboard still selected
that column, so between applying the migration and merging the branch, two live
surfaces were broken on comewith.org: the Templates page rendered an error panel,
and the calendar's gap scan degraded to "scan unavailable".

The database is not branchable — a DB change that ships before its UI is the
normal shape here, and 145 did exactly that. But 145 was **additive**. The
asymmetry is the lesson:

- **Additive** (new column, new function, new table): the old UI keeps working,
  and the window is harmless. Apply whenever.
- **Destructive** (drop a column, drop a function signature, tighten a
  constraint): the old UI breaks the instant it applies. The window is a live
  outage, and its length is however long the UI takes.

For destructive changes, either apply *after* the UI is merged and ready to
deploy, or make the change in two migrations — add the new shape, ship the UI,
drop the old shape — which is the only version with no window at all.

What made this survivable rather than serious was luck about *what* broke: both
casualties were read-only surfaces that already degrade gracefully. The apply
button kept working only because adding a defaulted second parameter left the
one-argument RPC call resolvable — verified, not assumed. Had the new signature
not been call-compatible, dropping the old function would have broken the
feature the migration existed to improve.

Also worth recording, since it was a deliberate choice and not an oversight:
`event_type` was **dropped rather than left nullable**. A stale column that
nothing reads is worse than no column — the next person filters on it and gets a
silently empty list, which is the failure mode this file keeps returning to.

---

## Section 47 — A forecast keyed on the unit's NAME can never be compared to actuals (2026-08-21)

`budget_lines` held a hand-built forecast that looked completely reasonable:
`Come With Party #1 (7/11)`, `DJ Gig #1`, `Equipment Rental #1`..`#6`, each with
an income row and an expense row, plus standing `Marketing` $500 and `Software`
$230. Thirty-seven rows, $33,469. Someone had done real work.

None of it could ever have produced a variance number. `v_pl_monthly_vs_budget`
joins plan to actual **on `(period, category)`** — and those rows put the *unit's
name* in `category`. No P&L category is called "DJ Gig #1", so the join matched
nothing, every line reported 100% variance, and it had been doing so silently
since the day it was written.

The dashboard made it worse in the quietest possible way: `renderPLGrid()`
fetched that view on **every** P&L open and built a `planned` lookup from it —
then never read the variable. A variance feature wired to nothing, sitting on
top of a join that could not succeed. Neither half was visible as a failure,
because the output of both is *no output*.

Two rules out of it.

**A field that is a JOIN KEY is not a label.** `category` is how plan meets
actual. The moment it carries "which one is this" instead of "what kind is
this", the two sides stop being comparable and nothing announces it. The unit's
identity needed its own home — which is what `plan_offerings` is.

**Dead code that computes is worse than dead code that sits there.** An unused
`const` is a tidiness problem. An unused `const` built from a network fetch on
every render is a cost being paid for an answer nobody reads, and it reads to
the next person as "variance is handled here" — which is precisely why it
survived. When the feature moved to its own tab, the fetch and the lookup were
deleted rather than left "in case".

## Section 48 — Seed from what you can derive; flag what you had to guess (2026-08-21)

Migration 199 seeded six offerings out of those 37 legacy rows. Most of it was
honest derivation: the amounts are Keith's own budget figures, and the ticket
economics came from real `ticketing` history (Dance Infusion: 61 paid heads at
$28.06, both computed, neither invented).

One thing could not be derived. The $1,200 sitting against "Come With Party #1"
is a single number covering venue, talent and marketing, and the P&L category it
belongs to is simply not recoverable from it. Splitting it three ways would have
looked like data and behaved like data — feeding contribution, margin, breakeven
and every variance figure downstream — while being something we made up. That is
§26 exactly: a placeholder that reaches a computation stops being a placeholder.

So the split was not invented. Each seeded line carries `needs_review`, the
offering reports `provisional`, and the board says so in the header rather than
presenting the forecast as settled. The number is usable *and* labelled.

The same logic forbade a tempting shortcut in the other direction. Equipment
Rental and Event Production have no expense in the legacy budget, and Artist
Showcase has no income. Writing a `$0.00` line would have been easy and would
have asserted "this costs nothing" — a claim nobody made. No line is written,
and `v_plan_offering_unit` exposes `has_cost_model` / `has_revenue_model` so a
missing side renders as "no cost" rather than as a confident 100% margin. "Zero"
and "not modelled" are opposite claims; the schema now has room to say which.

The forward-looking half of this is that the guess is *correctable and tracked*:
`v_event_contribution` puts each event's real contribution next to what the
model predicted. First reading, every completed event came in under model —
Come With 7-11 contributed **−$900** against a modelled **+$1,150**. A model
that is wrong and says so is worth having; a model that is wrong and confident
is not.

## Section 49 — Verify the verifier, especially when it agrees with you (2026-08-21)

Three checks lied during this session, each in a way that would have passed
unnoticed.

**The grant check.** A hand-written query counted anon grants of *any* privilege
type and reported the new planning objects as FAIL — and, in the same run, the
five canonical financial views as FAIL too. Those are known-good and verified
401 in two prior closes. The disagreement with established truth is what exposed
the check: the blanket-grant era left non-SELECT anon grants behind that were
never revoked, so only `privilege_type = 'SELECT'` means anything. `post_apply.sql`
already documents this in a comment. The lesson is not "read the comment" — it
is that **a check whose result contradicts something you already know to be true
is reporting on itself, not on the system.** Re-run with
`has_table_privilege()`, which is authoritative, before believing either answer.

**The syntax check.** This machine has no JS runtime at all — `node`, `deno`,
`bun`, `npx` all absent — so `node --check`, the loop CLAUDE.md documents, could
not run. The substitute (esprima via Python) failed on the *pre-edit* file at
`||=`, then `?.[`, then `catch {`, then top-level `await`: four ES2019–2022
constructs an ES2017 parser cannot see. Each failure looked exactly like a
syntax error in the new code. **The control run is the entire method.** Only
once the unmodified file parsed did a PASS on the patched file mean anything —
and a deliberately broken brace was then injected to confirm the checker still
caught real errors, because a checker that passes everything passes your bug too.

**The function grants.** §45 landed in `CLAUDE.md` mid-session, from another
machine, saying `revoke ... from anon` on a function is a silent no-op. 201 had
used the correct `from public, anon` form — but that was checked against prod
with `has_function_privilege()` rather than assumed from reading the SQL, and
the check found something else: 197's two trigger guards had been created with
no grants at all and so carried `PUBLIC EXECUTE`. Harmless in fact (invoker
trigger functions refuse a direct call), closed anyway in 202. Reading your own
migration proves what you *wrote*, never what the database *did*.

## Section 50 — A filtered view sent onward must say what it filtered out (2026-08-25)

"Email task list" on Calendar & Tasks sends the board's current view — the active
filters, in the order on screen. That was the ask, and it is the right default:
the list you are looking at is the list you mean.

But it creates a hazard the unfiltered version never had. A recipient cannot see
your filter chips. A mail containing four tasks, sent while "high priority ·
overdue only" was set, reads as *"there are four things outstanding"* — and the
reader has no way to tell it apart from a mail that genuinely contains
everything. The sender knows the difference for about ten seconds; the mail
outlives that, gets forwarded, and gets acted on.

**So the filter travels with the mail.** `taskFilterWords()` turns the active
filters into a sentence, which rides in three places: the subject line, the grey
scope line under the heading ("Filtered view — high priority · overdue only.
Sorted by due date (ascending)."), and an italic note giving the count as **N of
M**. An unfiltered send carries none of it, so the note's presence is itself
information.

This is the same shape as the silent-cap rule in §18 — a truncated list that does
not announce its truncation reads as a complete one — and the same shape as §28,
where a bulk edit applied to rows a filter had scrolled out of view. **Any time a
subset leaves the screen it was defined on, it has to carry its own definition
with it.** The board, the email and the export are three views of one list, and
only the board shows the chips.

Two smaller rules fell out of building it:

- **Do not offer a second control that can contradict the first.** The hub's mail
  had an "Include completed tasks" checkbox. Once the board grew a `done` status
  chip, that checkbox could disagree with the board — tick one, untick the other,
  and you send a list that does not match what you were looking at when you
  pressed send. It was deleted; the chip decides, and grouped mode derives from it.
- **Two boards describing the same filters must not describe them twice.** The
  hub and the calendar were one copy-paste away from two slightly different
  sentences for the same state. `taskFilterWords` / `taskScopeLine` /
  `taskFilterNote` are shared, and `calFilterWords` is a wrapper that passes
  `withEvent` — the hub has no event dropdown and must never claim one.

## Section 51 — The close was verifying five views that nothing actually checked (2026-08-25)

`CLAUDE.md` and `MERGE_ROUTINE.md` both open the close with the same instruction:
all five financial views must return anon **401** — `v_event_summary`,
`v_kpi_event_financials`, `v_kpi_parties`, `v_kpi_dance_infusion`,
`v_kpi_dashboard`. That check is the direct descendant of the 016/017 regression
and is the single most-repeated rule in the repo.

`scripts/check_anon_exposure.py` is the tool the same file says to use, and says
not to hand-roll. **It does not check any of the five.** Its output was read for
months as satisfying the instruction because it is long, it names dozens of
objects, it ends with "Nothing is exposed that should not be", and four of the
five have names that *look* like things in the list (`v_kpi_targets_current` is
in there; `v_kpi_dashboard` is not). Grepping its 54 lines for the five names
returns nothing.

Prod was fine — verified independently, all five 401. The invariant held. **The
check on it did not exist.** For as long as that was true, a real regression
would have produced exactly the same clean-looking close.

The failure is one step further out than §37's. There the check read a proxy for
the invariant (grants instead of the response body, an empty key instead of a
working one). Here the check simply did not cover the thing, and the *ritual*
was the proxy — running a script named `check_anon_exposure` felt like checking
anon exposure. **A checklist item is only as good as the assertion behind it; if
you cannot point at the line that would go red, the item is decoration.**

`scripts/check_financial_views.py` now names all five explicitly, reads a known-
public view first and refuses to continue unless the key actually works (§37),
and exits non-zero on any 200. Both scripts run at every close — the sweep for
breadth, this one for the five that are named in the rules.

## Section 52 — A secret cannot be read back, and the thing that looks like its value is a digest (2026-08-25)

Gear Watch's eBay source had never run. Keith added the credentials; the scanner
still said `NOT CONFIGURED`, because they had been saved under eBay's own field
labels — `App ID` and `Cert ID` — rather than the names the code reads. Fine so
far. What happened next cost two round trips and produced a confident, wrong
accusation.

The Supabase **Management API returns a SHA-256 digest in the `value` field of a
secret, not the secret.** Every secret reads back as 64 hex characters whatever
it holds. Read that way, `App ID` and `Cert ID` looked like "64 hex characters,
no hyphens, no `PRD` marker" — which is emphatically not the shape of an eBay
keyset. From that I concluded the values were wrong, told Keith so, copied those
digests into `EBAY_CLIENT_ID` / `EBAY_CLIENT_SECRET`, and tested **the digests**
against eBay's OAuth endpoint. eBay answered `401 invalid_client`, which read as
confirmation. It was nothing of the kind: I had authenticated with a hash of his
password and reported that his password was wrong.

Proving it took one line — set a secret to a URL known exactly, read it back,
and compare against `sha256(url)`. It matched.

**Rules that follow:**

- **A secret is write-only.** It can be written and then *exercised*; it can never
  be inspected. There is no rename, no copy, no "move this value to that name" —
  any of those require plaintext nobody has. Re-entering it is not a workaround,
  it is the only mechanism.
- **Never characterise a credential from a read.** Length, charset and prefix are
  all properties of the digest. The only legitimate test of a credential is to
  send it to the system that owns it and see what that system says.
- **When an inspection contradicts the user, distrust the inspection first.**
  Keith said twice that the values were correct. Both times the reply was a
  sharper description of the digest. He was right, the tool was lying, and the
  tie-breaker — "what does eBay say when we use the real value?" — was the one
  test not being run.

Same family as §49 and §37: the check measured a proxy for the thing. Here the
proxy was so faithful in shape (fixed-width hex, stable per value, different per
secret) that nothing about it looked like a placeholder.

*Applied correctly the same hour, for once:* the edge-function log query returned
zero rows for `ebay-account-deletion`, which would have proved eBay never called
it. Running the same query with no filter also returned zero — for functions
invoked minutes earlier. The log source was simply empty on this plan, so the
finding was reported as "cannot tell" rather than "eBay never called".

## Section 53 — A paid, blocking source must not share a button with free, instant ones (2026-08-25)

"Run scan now" on Gear Watch had **never once succeeded.** `gear_watch_runs` held
19 rows, all `cron`, zero `manual`. Reproduced before changing anything: HTTP
**546 at 151.3 seconds** — the edge runtime's 150s wall clock.

The cause was one button standing for four sources that are not comparable.
Reverb, Craigslist and eBay are API calls that finish in about six seconds
between them. Facebook is an Apify scrape via `run-sync-get-dataset-items`, which
**holds the connection open until the scrape completes** — tens of seconds each,
once per target, sequentially. Four of those cannot fit in 150s and never could.
Because the run row is written last, every press died before writing anything:
no error row, no log line, nothing but a toast that faded. The feature had a
failure mode with no evidence, which is why it survived from the day it shipped.

Three things it now does:

- **Split by cost and latency, not by category.** `manual` is the free three;
  `facebook` is its own button behind a confirm that names the price. Cron is
  unchanged. A source excluded by mode says which mode excluded it — it never
  reports zero.
- **Bound the blocking call.** Facebook runs against a 110s deadline with
  `AbortSignal.timeout` per scrape, leaving room to score and store, and refuses
  to start a scrape it cannot finish rather than burning the credit and the wall
  clock together.
- **Rotate, or the tail of the list is never searched.** Only two or three
  scrapes fit. Starting at target #1 every press meant the last targets would
  *never* be searched, however many times the button was pressed — a permanent
  blind spot that reports itself as a completed scan. Each run now starts where
  the last one stopped, and the status names the gear rather than a count:
  "searched AlphaTheta Wave 8, Pioneer XDJ-AZ, Pioneer CDJ-3000; **NOT searched
  KRK Rokit 5** — press again to continue where this left off." `PARTIAL` is its
  own state and the UI toasts it as a problem, not a success. §18 again: a
  subset that does not announce itself reads as the whole.

**And the reason eBay was dark had nothing to do with any of this.** eBay
disables a production keyset until the account has a working Marketplace Account
Deletion endpoint, and a disabled keyset answers `401 invalid_client` — byte-for-
byte what a wrong password returns. The keys were right the whole time. A new
`ebay-account-deletion` function (deployed `--no-verify-jwt`, since eBay calls it
unauthenticated) answers the challenge with
`sha256(challengeCode + verificationToken + endpoint)`; the endpoint in that hash
comes from a secret rather than `req.url`, because behind a proxy those differ
and the mismatch yields a valid-looking hash eBay silently rejects. With it
verified, eBay went from `NOT CONFIGURED` to `ok — 173 listing(s)`, and the
scan's reach went from 173 listings to 346.


---

## Section 54 — "Is it public?" is two flags in this schema, and one of them is date-scoped (2026-08-27)

`v_artist_gigs` (065) listed a gig on a public artist profile when the event was
`is_public = true` **or** `status = 'completed'`. The second half was a leak:
every completed event published its participant list by name, whether or not it
had ever been announced. Private bookings and anything deliberately left
unpublished went public the moment somebody marked it complete. Keith spotted it
from the page, not from the SQL.

The obvious fix was to drop the `completed` half and gate on `is_public` alone.
That is what migration **204** did — and the pre-apply check is the only reason
it did not stand. Counting the rows before and after showed gigs falling from
**60 to 24**, and the list of what vanished included **Dance Infusion #1 and
#2** — the flagship public shows, gone from every DI artist's page.

**`is_public` is not "this event faces the public". It is "this event is on the
upcoming-events feed".** 030's own column comment says so, and both consumers
(`v_public_events` 030, `v_public_events_hero` 064) also filter
`event_date >= current_date`. Nobody had ever set it on a past event because
doing so had no effect anywhere. The past-facing flag is a different one:
**`is_featured`**, which is what puts an event in Recent Rooms on the homepage
via `v_public_recap` (061/063/184).

**205 gates on `is_public OR is_featured`** — announced upcoming, or publicly
recapped. In Keith's words when asked which past events should count: *"everything
that is showing in recent rooms."*

Two things worth keeping:

- **Encode the rule, don't sync the flags.** The alternative was hand-flipping
  `is_public` on today's featured events. It would have produced the identical
  result this afternoon and then drifted silently: the next event featured in
  Recent Rooms would have been missing from its artists' profiles, with nothing
  on either screen to explain why.
- **A count before and after is a cheap way to be wrong out loud.** The migration
  read correctly, dry-ran clean, and would have been a quiet regression on nine
  public profiles. What caught it was asking how many rows this changes and which
  ones — before applying, not after. `supabase/checks/pre_apply.sql` exists for
  exactly this; a view whose whole job is a `where` clause deserves it most.

A side effect worth its own line: the four **Growth & Networking** events
(Elements, We Belong Here, Hulaween, JunXion) are festivals the team attended to
network, and they had been listed as **gigs** on artist profiles all along.
Carrying neither flag, they now correctly do not appear.

---

## Section 55 — Link the name that is printed, not the record it came from (2026-08-27)

Radio episodes now link their "Mixed by <name>" credit to that artist's public
profile, and each artist's profile lists the episodes back.

An episode has two things that look like the answer: `sc_playlists.mix_by`, free
text, the name actually rendered on the page; and `assigned_actor_id` (130), a
real foreign key to `actors`. The FK is the tempting one — it is typed, it
cannot go stale, it needs no matching.

**It is also the wrong one.** `assigned_actor_id` is whoever was given access to
*build* the episode. When a guest mixes an episode Keith set up, linking on the
FK renders "Mixed by \<guest\>" pointing at Keith's profile — a link that is
confidently, invisibly wrong. Matching on `mix_by` can only ever fail by not
finding anybody, which renders as plain text and tells no lies. The FK is the
fallback only when nothing is credited at all, at which point there is no name on
screen to link anyway.

The same reasoning settles the duplicate case: two public actors sharing a
display name resolve to **neither**. A 50% chance of pointing a fan at the wrong
person is not better than no link.

**One rule, used in both directions.** The same `creditedArtist()` decides which
profile an episode links to and which episodes a profile lists, so the two can
never disagree — an episode that links to an artist is exactly an episode that
artist's page lists back.

And the read stayed inside `get-station`. The reverse lookup wanted a list of
published episodes for an actor, which a small anon view over `sc_playlists`
would have served — but `sc_playlists` was anon-revoked in 103 precisely so that
public station reads go through the function, and a new view would have undone
that quietly. It is a `?artist=<id>` mode instead: published episodes only, and a
non-public actor gets the same empty answer as an unknown id, so the endpoint
never confirms that a private profile exists.


---

## Section 56 — A derived number that is right by coincidence is still not a measurement (2026-08-27)

Every surface showing an episode's length — homepage card, radio hub, episode
page, and the artist profile card added the same day — summed
`sc_playlist_tracks.duration_ms`. That is the total length of the SOURCE TRACKS,
and it had never been the runtime of the mix. Keith caught it from the profile
card: *"it shows the incorrect minutes."*

The fix was straightforward once looked for. `sc-connect`'s `mix_stats` action
already fetched the published mix's own track object from api.soundcloud.com to
read its play count, and was discarding `duration`. It now stores
`sc_playlists.mix_duration_ms` (206) and every surface reads that.

**The part worth writing down is what the numbers showed.**

| SHOW | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
|---|---|---|---|---|---|---|---|
| real | 61 | 65 | 64 | 60 | 58 | 43 | 56 |
| summed | 86 | 98 | 65 | 60 | 59 | 43 | 109 |

Four of seven were within a minute. **That is why it survived this long.**
Anybody sanity-checking the hub would have looked at SHOW 3, 4, 5 or 6, seen a
plausible hour, and moved on. The wrongness only shows on the episodes whose
tracks came from SoundCloud at full length, where a DJ set's cutting and
overlapping makes the sum over-report by nearly double.

And the near-misses are not the calculation partly working. Those four are
Beatport/Rekordbox episodes whose stored track lengths are **preview clips** —
averaging 1.8 minutes, as short as 40 seconds. They land near an hour because the
clips are short and there are roughly as many of them as an hour needs. Two
unrelated errors cancelling is not a measurement; it is a coincidence that will
stop holding the moment an episode mixes both sources.

I got this wrong in the first commit, in the confident direction: I wrote that
SHOW 6's "43 min" was an hour-long mix under-reported by a third, reasoning from
the preview-clip durations without having the real one to check against. The
backfill produced 43 minutes. **The claim was written into a migration comment, a
column comment on prod and two edge functions before the number existed to test
it against** — and had the backfill not run in the same session, it would have
sat there as the durable explanation of a bug it described backwards.

Two rules out of it:

- **When replacing a computed value with a measured one, print both, per row,
  before writing down why the old one was wrong.** The diff is the evidence, and
  it costs one query.
- **"Close on most rows" is not partial correctness for a derived quantity.**
  Either the derivation models the thing or it does not. Ask which rows it is
  wrong on and why *those*; if the answer is "the ones where the inputs happen to
  be shaped differently", the agreement everywhere else is luck. Same family as
  §26 (a placeholder that reaches a computation becomes invented evidence) and
  §23 (never render a blank as zero): a number on screen is a claim, and it
  carries no marking to say it was a guess.

---

## Section 57 — Prove the column is read, not just that nothing broke (2026-08-27)

Migration 203 added `quantity` to `plan_offering_lines` so a pricing line could
say "100 tickets at $25" instead of "$2,500". `quantity` defaults to 1, so the
whole change is meant to be **inert**: every existing line must compute exactly
what it computed before.

The obvious check is a before/after comparison, and it was done properly — a
single fingerprint over every number the two planning views can produce
(`v_plan_offering_unit`, `v_plan_monthly`, `v_event_contribution`, and the lines
themselves), captured against prod, then recomputed inside the dry-run
transaction. Identical, all four.

**And it would have been identical if the views had ignored `quantity`
completely.** That is the trap. "Nothing changed" is exactly what a column
nobody reads looks like, and it is also exactly what a correct migration looks
like — the same evidence supports both, so on its own it distinguishes nothing.
An inertness check can only ever tell you the change is *harmless*; it cannot
tell you the change is *there*.

So the dry run also doubled the quantity on one `per_unit` cost line and asserted
the offering's cost per unit rose by exactly that line's amount. It did
($6,557 → $13,114). Only then did the fingerprint mean what it appeared to mean:
not "nothing happened", but "the column is wired in and reads as 1 today".

A third run inserted a `pct_revenue` line with `quantity = 2` and required it to
FAIL, because a percentage has no count and the constraint pinning it to 1 is
load-bearing rather than decorative. It failed with `23514`, as intended.

**The shape of the lesson.** §49 said verify the verifier when it agrees with
you. This is its sibling: when an expected result is "no change", a passing check
is indistinguishable from a check pointed at nothing. Pair every inertness proof
with a deliberate perturbation that MUST move the number, and a deliberate
violation that MUST be refused. Three runs, not one.

**Two other things fell out of the same session, both of them checks lying:**

- **The dry run caught a schema fact reading the repo could not.** 203 rebuilt
  `v_plan_offering_unit` from the definition in 198 — but 199 had quietly
  replaced that view with two extra columns (`has_revenue_model`,
  `has_cost_model`). `create or replace view` refuses to drop columns, so the
  apply would have failed. Reading your own migration history tells you what was
  *written*; only prod tells you what the database *is* (§45 again, from the
  other direction).

- **An RLS negative control erased its own evidence.** The probe checked that
  `anon` could not insert a pricing line by attempting the insert inside a
  plpgsql `BEGIN … EXCEPTION` block and raising `'PROBLEM: anon inserted'` if it
  succeeded. That block is a **subtransaction** — the raise rolls back the very
  insert it is reporting, so the follow-up count finds nothing and the probe
  passes whatever the database does. Record outcomes in a variable and return
  them as rows; never `raise` to report a failure you are trying to observe.
  (The first run also leaked the admin's JWT claims into the anon block, because
  `set_config(..., true)` is transaction-local, not block-local: `auth.uid()`
  still returned Keith, `is_admin()` still passed, and the "anon" test was the
  admin test wearing a different role name. Clear the claims before switching
  role, or you are testing nothing.)

---

## Section 58 — A sweep that walks a hand-written list is a list, not a sweep (2026-08-31)

`scripts/check_anon_exposure.py` was written to answer one question: is anything
readable by the public that should not be? It ends with the line **"Nothing is
exposed that should not be."** and a zero exit code, and it has been run at the
close of most sessions since it was written.

Line 38 of that file said:

> The ones worth naming explicitly, so a failure reads as a sentence rather than
> a table name. **Everything else discovered from the schema is checked too.**

The second sentence was not true. There is no schema query anywhere in the file.
`main()` iterates `MUST_BE_EMPTY`, then iterates `PUBLIC_OK`, and stops. An object
in neither list is **never requested at all** — no row is printed for it, and the
closing "Nothing is exposed" says nothing whatsoever about it.

This surfaced when migration 207 added two tables and three views and a full run
came back clean without ever mentioning them. It is the identical failure to the
five financial views in §51, in a different file: a check whose *output* was read
as covering something its *code* never touched. §51 caught it in what the script
omitted; this is the same trap one level up, in a comment that stated the
opposite of the code directly above it.

Two things were wrong and both are now fixed. The comment says, in capitals, that
the two lists **are** the whole sweep and that every new table and view must be
added to one of them in the same migration that creates it. And the 207 objects
are named: `link_pages`, `link_items` and `v_link_click_stats` under
`MUST_BE_EMPTY`, `v_public_link_pages` and `v_public_link_items` under
`PUBLIC_OK`.

The general rule, now stated three times in this document under three different
disguises: **a green check proves something about the objects it names, and
nothing at all about the objects it does not.** When a check reports confidently,
grep its output for the thing you actually care about before believing it. If the
name is not there, the check did not run — however clean the summary line looks.

An honest comment would have made this visible years earlier than an audit did.
A comment that describes behaviour the code does not have is worse than no
comment: it is a claim, read as evidence, that nobody re-derives.

---

## Section 59 — A preview that is not the page is not evidence (2026-08-31)

The links-page editor needed a live preview: change a colour, drag a row, see it.
The obvious build is a small renderer inside `dashboard.html` that draws
approximately what `links.html` draws. That was rejected, and the reason is
already written down elsewhere in this repo at some cost.

The planner implements its forecast maths **twice** — in SQL (`v_plan_monthly`)
and in JS (`planModelMonth`) — because a lever that needs a round trip before it
shows a number is a form, not a lever. That was the right call there, and it is
paid for with a standing rule in CLAUDE.md: *if you change one, change the other,
or the number you type against stops matching the number you reload into.* A rule
like that is a permanent tax, and it is only worth paying when there is no way to
have one implementation.

Here there was a way. `links.html` accepts `?preview=1`; in that mode it does not
touch the database and does not load the beacon, it waits for a same-origin
`postMessage` carrying `{page, items}` and renders that through the exact code
path a visitor gets. The dashboard embeds it in an iframe and posts the unsaved
form state on every keystroke. One renderer. The preview cannot drift from the
page, because it **is** the page — a layout change is visible in the editor
without anybody remembering to mirror it.

Two details that make it honest rather than merely convenient:

- **The preview filters like the public view does.** Inactive rows and rows
  outside their `starts_at` / `ends_at` window are dropped before posting, because
  those are exactly what `v_public_link_items` removes. A preview that shows more
  than the visitor will see flatters the page, which is the one thing a preview
  must never do.
- **The icon list is read out of the iframe** (`CW_LINK_ICONS`, same origin)
  rather than copied into the editor. If the renderer learns a new glyph the
  editor offers it, with nothing to keep in step. Where the read fails, the icon
  field stays free text and still works — the suggestions were only ever a
  convenience, so their absence degrades rather than breaks.

The general shape: **when a second surface has to show what a first surface
produces, embedding the first one costs less than reimplementing it — and unlike
a reimplementation, it cannot be wrong.** Reach for the duplicate only when the
two surfaces genuinely cannot share a runtime, as SQL and the browser cannot.

---

## Section 60 — Prefer the foreign key, but not when almost nothing has one (2026-09-01)

Auto-assigning the DJ to the radio week's production tasks looked like a
one-liner: `sc_playlists.assigned_actor_id` is a real FK to `actors`, and
`task_assignments` already takes an `actor_id`. Join them and ship.

Counting first is what stopped that. On prod, `assigned_actor_id` is set on
**one** of ten stations. `mix_by` — free text, the DJ's name — is set on **nine**.
The column with referential integrity is the one nobody fills in; the column that
is actually maintained is the one holding a string. A feature built only on the FK
would have done nothing on nine episodes out of ten, including the one currently
being built, and would have looked like a bug rather than like missing data.

So the resolution order is: **`assigned_actor_id` when set** — it is explicit and
unambiguous — then `mix_by` matched by name. All four names in use today
(`Berky`, `KRNeY`, `Henry`, `32LVS`) resolve to exactly one live actor, so the
fallback is what makes the feature work at all.

Note this is the mirror image of §55, not a contradiction of it. There, an episode
credit must link on the **printed name** (`mix_by`) and never on
`assigned_actor_id`, because the FK records who was given access to *build* the
episode and rendering it as the credit points a visitor at the wrong artist. Here
the question is who should *do the work*, and the FK — when it is set — is exactly
the right answer. Same two columns, opposite precedence, because the question is
different. Anyone "tidying" one of these to match the other will break the other.

The guard is the same in both places, and it is the part that must not be dropped:
**zero matches and two matches both assign nobody, and say why.** A task on the
wrong person's board is worse than a task on nobody's — and worse here than in
§55, because `tasks` carries an "Actors can read assigned tasks" policy, so
assigning does not merely label a row, it *shows that row to that person*. A bad
name match is an unintended disclosure, not just a wrong label.

Two smaller things the same change is careful about. Which steps count as "the
DJ's" is a **default, not a decision** — every row is a ticked-or-unticked
checkbox in the confirm modal, showing the name it will land on, so the split is
visible before anything is written rather than encoded invisibly in a constant.
And the returned rows are mapped back to steps **by title, not by position**:
assuming an insert returns in payload order is exactly the sort of assumption that
fails silently and puts the wrong task on someone's board.

---

## Section 61 — A missing input is not a zero, and a zero is not always a measurement (2026-09-01)

The Buzz score ranks NYC artists for radio. It blended three inputs — RSVP demand
(0.40), follower reach (0.35), catalogue size (0.15) — each scaled against the
current pool's maximum, plus flat bonuses. Asked to add top-track plays, the
honest answer was that the plays were the smallest of four problems, and every one
of them was invisible in the output.

**1. The heaviest input was scoring a ticketing platform, not an artist.** RA
publishes an attending count on all 941 of its future events. DICE publishes it on
none of its 263, Ticketmaster on none of its 48. `Number(e.attending) || 0` turned
that absence into a zero, so an artist whose shows happen to be sold through DICE
lost 40% of their score for a fact about the vendor. Galantis scored **47**. Icona
Pop, 45. Both are now in the top four. The score was not measuring buzz; it was
measuring who sells through Resident Advisor.

**2. Scaling against the pool maximum makes the number incomparable with itself.**
Every component was `100 * log1p(x) / max(log1p(x))` across the current list, so an
artist's score moved when somebody *else* was scanned or a bigger name appeared on
a lineup. "72 last week, 68 now" said nothing about the artist. It also compressed
everything: the old scores ran p25 28 to p90 59 and **nobody ever exceeded 81**,
because the weights summed to 0.90 and the last ten points needed an RA editorial
pick. Fixed anchors — lo scores 0, hi scores 100, log-spaced, chosen off the real
prod spread — make the number mean the same thing every week, for the same reason
`v_kpi_prior` refuses to compare a metric against its own latest reading (§20).

**3. A failed scan looked exactly like an artist with no music.** 191 cache rows
have `ok = false`. The dashboard was not even selecting `ok`, so a scan that failed
contributed `0 songs, 0 followers` and scored as fact.

The fix for all three is the same shape, and it is the general lesson: **an input
that could not be measured is dropped from the weighted average and shrinks the
DENOMINATOR, instead of entering the numerator as a zero.** What is left is a score
over what was actually knowable, with the coverage shown next to it — 100% for
710 artists, 70% for 381, 55% for 316. 27 artists have nothing measurable at all
and are now shown as "—", because ranking them 0 would place them below artists we
know to be small, which is a claim the data does not support (§26 again).

**Then the part that nearly shipped wrong.** Having decided a successful scan makes
plays "known", zero uploads made plays a known zero. That is defensible — until you
look at who it demotes. **Ben UFO**: 109,333 followers, 2,010 RSVPs, no original
uploads, scored **56**. Craig Richards, Joseph Capriati, Jyoty all the same. 393
scanned artists have no tracks. These are selector DJs who play other people's
records, and they are precisely who a radio show wants to book.

The error was treating one fact as two independent measurements. **The plays of a
catalogue that does not exist are not zero, they are undefined.** Catalogue size
already records "uploads nothing" — counting it again as "nothing gets played"
punishes the same fact twice. Gating plays on having at least one track puts Ben UFO
at **76**, which is obviously right. So: `catalog = 0` is a real measurement,
`plays` on an empty catalogue is not, and the two live one line apart.

The distinction is not "is the number zero" but **"did anybody look, and was there
anything there to look at"**. Three different absences — the platform never
publishes it, the scan failed, there is nothing to measure — all arrive as `0` in
JavaScript and must not all be scored as one.

---

## Section 62 — A filter is a claim about the world, and this one was three claims short (2026-09-02)

"When I isolate Refuge and Industry City I only see one artist." Refuge had two
dozen artists on its bills. Three separate defects stacked to produce that one,
and none of them announced itself.

**1. The venue dimension is free text from three feeds, and it is not consistent.**
The artist pool held `'Refuge'`, `'REFUGE'` and `'REFUGE '` — three distinct
strings, the last differing only by a trailing space (hex `…474520`). The dropdown
listed all three; two rendered identically. They held 23, 1 and 1 artists. Picking
either look-alike is exactly how a busy room shows one name. The same fragmentation
hit Alphaville/ALPHAVILLE, Drom/DROM, H0l0/H0L0, public records/Public Records and
`'Dead Letter No. 9'`/`'Dead Letter No.9'` — **155 stored strings for 149 real
rooms**. Group on a canonical key (case, punctuation and whitespace folded); show
the most common spelling. A `Set` of raw strings deduplicates nothing when the
strings disagree about capitalisation.

**2. Each artist was pinned to ONE show, so a venue filter could only ever see the
artists whose FIRST show in the window happened to be there.** `raWindowPool()`
re-points every artist at their soonest show and stores a single `next_venue`;
`raRadioList()` then compared `a.next_venue === venue`. An artist playing Nowadays
on the 5th and Refuge on the 20th was a Nowadays artist, and Refuge never showed
them. Measured across the window: Bossa Nova had 82 artists on its bills and the
filter offered 62; Dead Letter No.9, 29 and 4; Signal, 59 and 41. In total
**1,228 of 1,499 artist-venue pairs were reachable** — the filter was quietly
answering a different question than the one being asked.

A one-row-per-entity projection is fine for "when do they next play". It is wrong
the moment it backs a filter over a dimension the entity has *many* of. The clue is
that the field is singular (`next_venue`) while the question is plural.

**3. `Producers only` is on by default and was removing a third of the room in
silence** — Industry City 38 → 23, Refuge 24 → 19. The filter is wanted; its
silence was not. It now reports what it hid, the same rule as any other cap (§18):
a number with no note reads as "that is all there is".

**And the ceiling nobody could see.** 35% of future RA events and 61% of DICE ones
carry **no lineup at all**; Industry City had 9 shows and 3 lineups. No filter can
show artists that were never pulled, so isolating a room now states its own
coverage — "9 shows in this window, 3 with a lineup" — because otherwise the
filter is blamed for a gap that lives upstream in the pull.

The general lesson is the one that keeps recurring in this repo under new costumes:
**an empty result is a claim, and it is usually the least likely of several.**
Before trusting "there is only one artist here", check whether the key matched,
whether the projection could represent the answer, whether a default filter ate it,
and whether the source data was ever collected. Here all four were wrong at once,
and each alone would have looked like a quiet, plausible truth.

---

## Section 63 — Two mechanisms, because no single threshold separates them (2026-09-02)

Asked to auto-match misspelled venue names to one canonical room, the tempting
build is one similarity score and a cutoff. The data refuses it. On prod:

    randall s island   <->  randalls island        0.97   same room
    hotel 50 bowery    <->  hotel 50 bowery ny     0.91   same room
    green room         <->  green room 42          0.87   DIFFERENT venues
    314 scholes        <->  314 scholes st         0.88   same room
    brooklyn army terminal <-> … pier 4            0.86   different spaces

There is no cutoff that keeps the first four and rejects the third. A score is a
measure of *string* similarity and the question is about *places*, so similarity
can raise the question but must never answer it.

So the feature is **two mechanisms with different authority**:

**Deterministic folds apply automatically, with no review.** `normalize_venue_name()`
folds accents, case, `&`/`and`, punctuation and whitespace. Two names equal after
that are the same room by construction — there is no judgement in it, so asking a
human to confirm it would be theatre. This alone merged `'Refuge'`/`'REFUGE'`/
`'REFUGE '`, `'Crossroads Cafe'`/`'Crossroads Café'`, `'telos.haus'`/`'Telos Haus'`,
`'Dead Letter No. 9'`/`'No.9'`, five duplicate rows inside the `venues` table
itself (including `'Acoustik Garden Lounge'` entered twice), and linked **2,745 of
3,217 historical events** on the spot.

**Fuzzy similarity only ever suggests.** It fills a review queue ordered by how
many events each spelling holds, shows the closest room with its score, and waits.
A wrong merge here silently rewrites history — which is the exact thing the
feature exists to protect — so the cost of a bad automatic decision is far higher
than the cost of a click.

**What is deliberately NOT folded is as important as what is.** A leading "The"
and the feeds' "TBA - " prefix are both plausible rules, and neither appears as a
real collision in the data today. A fold invented ahead of evidence is
indistinguishable from a bug: it merges rooms that differ and leaves nothing
behind to show it happened. They stay suggestions.

**Keep the raw string; add a resolved id.** `ra_events.venue_name` still holds
exactly what the feed sent, and `venue_id` is the room it resolves to. Normalising
in place would have destroyed the only record of what the source actually said,
and left no way to re-derive a ruling that turns out wrong. The alias table is the
same idea: every judgement is a row, so it can be read, changed, or disagreed with
later. "Not a venue" (`status = 'ignored'`) is a stored decision too — otherwise
`TBA` and `listen` climb back to the top of the queue every week and the queue
teaches you to ignore it.

**The alias and the back-link must happen together.** `link_venue_alias()` writes
the alias *and* re-points every historical event with that spelling, in one
transaction. As two client calls it eventually fails halfway and leaves a venue
with a correct alias and 34 events still pointing at nothing — a state nobody
would think to look for.
