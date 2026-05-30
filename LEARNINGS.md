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
