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
