# Session review — 2026-08-15 · Strategy board rebuild (Henry's machine)

Fifth close of the day. Ran on Henry's machine. All of it merged the same day (PRs
#7–#13 and #15) and is live on comewith.org. Nothing left open.

## The arc

**Set out to do:** make the Strategy page readable. Keith's framing was "it's unreadable
and we're not getting the actionable insights we want from it" — ~35 equal-weight cards
in four workstream groups, one scroll, no hierarchy.

**What we discovered instead:** the page wasn't only badly laid out, it was structurally
incapable of showing a trend. `v_kpi_dashboard` took `current_value` from
`coalesce(computed, snapshot)` but `prior_value` **only** from the second-latest
hand-logged reading — and nothing hand-logs net P&L, sell-through or subscriber counts.
Prod introspection settled it: `metric_snapshots` held `youtube.*`, `instagram.*`,
`tiktok.*` and nothing else. Every card anyone would actually make a decision on had read
"– no prior reading" since the day it shipped, and `as_of` was hardcoded to
`CURRENT_DATE` so all of them claimed to be "updated today" forever.

That reordered the work: fix the data layer first (Phase 1), then the page (Phase 2),
then the funnel (Phase 3). No amount of layout would have helped a number with no history.

**Second discovery, at the funnel.** The plan was to attribute ticket clicks by page path.
That would have returned zero forever: `event.html` reads `v_public_recap` and is a
retrospective archive page with no ticket CTA at all. The ticket links are on the
**homepage**. Clicks record with `path='/'`, so attribution had to be by matching
`link_url` to `events.ticket_url` instead.

## Key decisions (LEARNINGS §20–23)

- **"Prior" is per-metric and lives in one view.** Previous completed event for event
  metrics, previous 5 uploads for content, nearest-30-days-ago for everything else —
  falling back to the *earliest* reading, never the latest, which would compare a number
  to itself and render a confident permanent "no change".
- **Categories derive from the metric-key prefix, not `kpi_targets.workstream`.** Re-filing
  those rows would have made the already-deployed renderer silently drop nine cards the
  instant the migration landed. The DB and the front end deploy on different clocks.
- **Last-event values are the headlines; lifetime averages moved to the drill-down.**
  Cost to raise $1 is the DI health metric (Keith's call). It is `lte`, so colour follows
  the comparison, not the sign.
- **"0" and "cannot be measured" are different claims.** Blanks, missing targets and
  absent funnel denominators all render as unknown rather than zero.

## What it surfaced the moment it worked

Last party **−$800 at 25% sell-through**. Cost to raise **$0.69 against a $0.50 target**,
worse than the $0.61 before it. Recent uploads averaging **103 views vs 274** for the five
prior. Mailing list **107 against 1000**. And two good ones: DI raised **$9,557 vs $2,943**,
attendance **117 vs 42**.

## Parked / next

- **Set `ticket_url` on Come With #2 and Dance Infusion #3.** The funnel is live and
  measures nothing until one exists, and the beacon cannot backfill.
- Nothing left to merge. The bar chart needed three corrective passes after its first
  merge (#9 stretching, #12 width + hover card, #15 title alignment) — all three found by
  Keith on the live page.

## One honest note

None of this UI has been seen rendering. There is no local console check for
`dashboard.html`, the Browser pane cannot composite so screenshots time out and layout
boxes read `auto`, and PR #7 was merged before anyone clicked its deploy preview. Six
category blocks, five charts, the collapse behaviour and the funnel panel are all
structurally verified — `node --check` against a HEAD control at every step — and visually
unverified. The bar chart alone needed THREE corrective passes after merging (#9, #12,
#15), every one caught by Keith on the deployed page rather than by anything here. That
is the honest error rate for UI shipped without a render check. The first person to open Strategy is doing
QA, whether they mean to or not.
