# Session 2026-08-21 — Planning (FP&A) — Keith's machine

Third session on 2026-08-21. Started as a CSS bug report, ended as a new module.

## The arc

**Set out to do:** fix some white boxes on the P&L tab.

**Found:** the P&L stylesheet was authored against a light theme and dropped into the
dark dashboard. It painted five backgrounds with `var(--panel, #fff)` and `--panel` is
defined nowhere in the file — so they fell back to literal white under `--ink` cream.
Only *some* cards broke, which is what made it look arbitrary: `.accent` and `.warn`
tint with `rgba()` and composite correctly over the plum-black page. Keith had hard
refreshed and still seen white, for the simple reason that the fix had never been
pushed.

**Then asked for:** build the tool out into an FP&A function — easily update forward
numbers, always keep actuals lined up against them, pull levers and make decisions. He
ran exactly this at Maersk, weekly at times, usually twice monthly.

**Found, again:** a forecast already existed and had never worked. `budget_lines` held
37 hand-built rows in nearly the right shape, but they stored the *unit's name* in
`category` — and the view that compares plan to actual joins **on category**. "DJ Gig
#1" matches no P&L category, so every row had reported 100% variance silently since it
was written. Worse, `renderPLGrid()` fetched that view on every P&L open and built a
`planned` lookup it never read. Two dead things stacked on each other, neither of which
could announce itself, because the output of both is no output.

## Key decisions

- **The unit of planning is an OFFERING, not an event.** Keith wants to rebuild this for
  a fashion company deciding SKU order quantities, so `creates_event` is a flag and
  `scale` is abstract with a per-offering label. The same four tables serve both. This
  shaped every other decision.
- **A published round is frozen by trigger, not by the dashboard** — a client guard is
  bypassed by any REST token. Publishing snapshots the live plan rather than closing it,
  so the working plan stays editable forever, which is what a rolling forecast needs.
- **The legacy 37 rows were preserved, not migrated.** `version_id` marks planner rows;
  the old forecast stays readable and cannot double-count.
- **Seeded models are labelled, not laundered.** Amounts are Keith's; the category
  behind each lump is not recoverable, so lines carry `needs_review` and offerings read
  `provisional`. No `$0` line is written where the truth is "not modelled" —
  `has_cost_model` says so instead, rather than showing a confident 100% margin.
- **The forecast maths is implemented twice**, SQL for truth and JS for the levers,
  because a lever that needs a round trip is a form. Recorded in CLAUDE.md as a
  change-both-or-neither rule.

## What the data said, immediately

- A DJ booking contributes **$75 on $500** (15%); a party contributes **$1,150** (46%).
- Every completed event came in **under model** — Come With 7-11 at −$900 against a
  modelled +$1,150.
- Parties have **no paid ticket rows at all**, so per-head pricing could not be derived
  and was not invented.

## One honest note

Three separate checks lied during this session, and each would have passed unnoticed if
taken at face value: a hand-written grant query that called the five known-good
financial views failures; a JS parser that rejected the *unmodified* dashboard at four
different ES2019–2022 constructs; and a migration whose grants looked right in the SQL
but had to be checked against prod anyway — where it turned out two *other* functions
had been created with `PUBLIC EXECUTE`. The pattern is the same each time: the check
disagreed with something already known to be true, and that disagreement was the only
signal. LEARNINGS §49.

The corollary is the session's real open risk: **the anon REST sweep never ran.** This
machine has no publishable key, so grants were verified in SQL — authoritative for
grants, blind to PostgREST. Flagged at the top of CARRYOVER *and* CLAUDE.md so the
desktop runs it first.

## Parked / next

1. Run the anon sweep on the desktop. Then delete the block from CLAUDE.md.
2. Click through the Planning tab — nothing has been exercised in a browser.
3. Confirm the six provisional pricing models.
4. Decide the two open model questions: party ticket pricing, and whether an equipment
   rental should book an event.
5. Not built: saved scenarios, unit→event conversion, cash runway, forecast drift.
