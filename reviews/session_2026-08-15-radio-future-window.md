# Session review — 2026-08-15 · Radio: future-dated windows (laptop)

Third close of the day, after the desktop's discovery audit and Henry's Notes module.

## What happened

Started as "pick up where we left off". The first real finding was that CARRYOVER was
already wrong: it described `radio/window-by-lineup` as unmerged and undeployed, but Keith
had green-lit it at 16:52 — `88b2153` merged it and both edge functions were deployed the
same minute. Verified by reading the live bundles off prod rather than trusting the file.

Keith then ran a real "↻ Pull shows" (the first through the fixed functions). Ticketmaster
was genuinely fixed — 25 → 40 events, 10 → 16 venues, Brooklyn present. DICE improved but
landed on **exactly** its new 240 cap, so the last two weeks of a 4-week window were still
empty. That reframed the session: he asked for the ability to select a start date as far
out as the sources offer, and to have every limitation in that path fixed.

## The arc

Expected one fix (`dj-station`'s `.limit(160)`). Found four, three of them silent:

- `dj-station` didn't just cap — it filtered on `ra_artists.next_event_date`, the *same*
  soonest-show bug the dashboard had just been rebuilt to avoid. On the one screen where a
  missing artist reads as "they have nothing".
- **The dashboard was silently losing a third of the pool.** PostgREST `max_rows` is 1000;
  every radio load is past it and none was ordered. That one is worse than the original
  bug — it was corrupting the artist list, the cache, and the "unscanned" count all at once,
  and it gets worse the further out you look.
- The pulls could only look forward from today, capped at 90/120 days.
- And the fix for that would have been destructive: all three pulls delete their own
  source's rows bounded only at the bottom. Aiming a pull at a 4-week window would have
  deleted every Ticketmaster show beyond it. Caught before deploy, not after.

Two more found in passing: the episode form replaced `dj_search_params` wholesale (opening
✎ Details on an Elements edition and saving would have wiped its fixed lineup scope), and
my own first draft capped the date picker at the pool's last date — which locks the door
from the inside, since the way you extend the pool is to set a date past it and pull.

## Decisions

- Windows are evaluated against `ra_events.lineup` **everywhere**, not just the dashboard.
  Two implementations of the same window is one too many; `dj-station` mirrors
  `raWindowPool()` and the comment in each points at the other.
- **A cap is only acceptable if exceeding it is reported.** Where it can be paged away,
  page it. LEARNINGS §18, now also in CLAUDE.md.
- DICE's finite detail budget is aimed at the **radio window**, not a flat 90 days. Shows
  outside the window are preserved from earlier pulls rather than refreshed — a deliberate
  trade, and the reason the delete had to be bounded at both ends.

## Parked / next

- **Nothing here has been clicked in a browser.** All four functions are deployed and
  verified by bundle read-back, prod invariants re-checked — but no human has seen the
  paging fix load a full pool. That's Parked item 1.
- DICE beyond the window stays stale until you pull with the window moved.
- Unchanged: financial views still readable by `authenticated` (the gated blocker); no cron
  pulls the market at all.

## One honest note

This laptop has neither `deno` nor `node`, so **none of this is compile-checked** — a real
gap versus Henry's close, which ran `node --check` against a control file. What I had was
reading, a bracket-balance pass baselined against the pre-edit file, and the fact that all
four deploys came back ACTIVE with the new code present in the bundle. That is evidence the
platform bundled them, not evidence they behave correctly. The first pull and the first
`dj.html` load are the real test, and I'd rather that be stated than implied.
