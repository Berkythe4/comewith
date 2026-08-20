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

