# Session — 2026-08-21 · Task templates: named sets, reviewed step by step

Ran on **Henry's machine**. Second close of the day; the desktop's invoicing
session is unrelated and untouched.

## The arc

**Set out to do:** put a pop-up in front of template application so each task can
be edited before it's created, one at a time, with a skip.

**What that turned out to require:** the pop-up was never the hard part. The
generator *decided and wrote in the same function*, so there was no moment at
which a proposal existed to show anybody. Splitting `plan_event_tasks` (decides,
writes nothing) out of `generate_day_of_tasks` (loops the plan, inserts) was the
actual feature; the modal queue is a consumer of it.

**Then it required something else:** letting people rename a step at creation
time quietly breaks the only link between a template and its task, which was
`lower(trim(title))` — used both by the generator's suppression and by the gap
panel's "N of M steps missing". A rename would have read as permanently missing
*and* re-created the original next to it. `tasks.template_id` exists because of
that, not because anyone asked for it.

**Then the scope grew, on request:** templates should be named sets you pick
from, not one list per event type. That's a second migration and a rewritten
Templates page, and it's the one that bit.

## Key decisions

- **Sets are fully free — no `event_type` at all.** Henry's call, from two
  options offered. Any set applies to any event. `event_type` was *dropped*, not
  left nullable: a stale column nothing reads is worse than no column.
- **The event remembers its set** (`events.task_template_set_id`), written on the
  first task actually created rather than when the set is picked — so abandoning
  a run never relabels an event. The gap panel measures against it.
- **Skip is per-event.** Nothing is written, so the step stays available. Removing
  a step for good is an edit on the Templates page — different intent, different
  screen.
- **`calAddTask` was split, not duplicated** (`taskFormOptions` /
  `taskFormMarkup` / `taskFormCollect`), honouring the comment already on it
  about not growing a second task form.
- **Duplicate** on a set is how you get a v2 without risking v1. It's the feature
  the request was really about.

## The honest note

**196 broke production for roughly half an hour.** It drops
`task_templates.event_type`, which the deployed dashboard still selected, so the
Templates page rendered an error panel and the calendar gap scan degraded to
"scan unavailable" between applying the migration and merging PR #17. I applied
it before the UI was ready because 195 had been safely additive and I carried
that habit forward without re-checking whether it still held. It didn't.

What kept it from being worse was luck about *what* broke — both casualties were
read-only surfaces that already degrade gracefully, and the apply button survived
only because a defaulted second parameter left the one-argument RPC call
resolvable (verified, not assumed). Had the new signature not been
call-compatible, dropping the old function would have broken the very feature the
migration existed to improve.

Second, smaller: `revoke ... from anon` on a new function is a no-op — EXECUTE
goes to PUBLIC and anon inherits it. 195 shipped with the wrong form. The
post-apply grant check caught it; reading the SQL had not, twice.

Both are now LEARNINGS (§45, §46) and CLAUDE.md rules.

## Parked / next

- **Nobody has clicked through any of it.** Verified at the data layer against
  prod — planner output, RLS as a genuine authenticated admin, the 3b suite,
  `node --check`, 16 headless form assertions — but the modals themselves have
  never run in a browser, because this machine can't sign in. First job next time.
- `.env` here lacks `SUPABASE_PROD_PUBLISHABLE_KEY`, so
  `scripts/check_anon_exposure.py` can't run on this machine; the anon checks were
  done by hand with curl, proving the key live before trusting any 401.
- If the set picker gets long, sort by usage — a sort, not a filter.
