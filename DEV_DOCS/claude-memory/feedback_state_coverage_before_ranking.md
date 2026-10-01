---
name: feedback-state-coverage-before-ranking
description: "Before answering \"top N / biggest / most\" from our data, check and state how much of the pool is actually measured; unmeasured rows are not small rows"
metadata:
  node_type: memory
  type: feedback
  originSessionId: a9f22648-7b83-432d-b7e4-f51edd7827bc
  modified: 2026-09-30T22:01:38.489Z
---

When Keith asks for a ranking ("top 5 biggest DICE artists", "sort by followers"), first measure coverage — how many of the pool have the metric at all — and say it in the answer. On 2026-09-30 I gave a "top 5 by followers" that was really the top 5 of the 53 of 338 artists we had measured; Keith caught it only because he knew Purple Disco Machine and Vintage Culture have SoundCloud. The real cause was three bugs (pullers wiping SoundCloud links, a sort on RA's count, unscanned links) — see LEARNINGS §74/§75.

**Why:** a missing value sorts like a zero, so a ranking over partial data looks complete and confidently omits the biggest names. Keith checks results against what he knows about the scene.

**How to apply:** for any "top/most/biggest" answer, report "N of M measured" alongside it; if coverage is low, investigate why before ranking (see [[project-radio-discovery-window]]). Pair with the repo rule "never render a blank as zero".
