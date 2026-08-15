---
name: feedback-time-estimates
description: "User wants an estimated time at the start of EVERY prompt + a \"spent / remaining\" ticker as I go. Reaffirmed 2026-06-01 to apply to all prompts, not just non-trivial ones."
metadata: 
  node_type: memory
  type: feedback
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

**⚠ CALIBRATION 2026-06-02 — the "spent" ticker was running WAY too high.** Keith: "the time
you say you spend at the top is always way higher than the actual elapsed time." Correct. I
have **no real wall-clock**; the "spent ~X min" figures were guesses scaled to how much *work*
a turn felt like — and work-volume massively inflates them. A turn with lots of output still
elapses in **seconds to ~1–2 minutes** of Keith's real time. Fix:
- Keep the **up-front "est X min"** (forward-looking; it's what lets Keith decide wait vs.
  step away — the useful part).
- **Do NOT fabricate an inflated "spent" total.** Either omit the retrospective "spent"
  number, or give an honestly small one (most turns: well under a minute to a few minutes —
  NOT tens of minutes, regardless of how much code/output was produced).
- When unsure of elapsed time, say so qualitatively ("quick" / "a few min") rather than
  inventing a precise minute count.

**Reaffirmed & broadened 2026-06-01: keep time estimates up for ALL future prompts**, not
just non-trivial chunks. Every response should lead with an "est X min" up front. Even
quick/trivial tasks get a short estimate (e.g. "est <1 min"). (See the 2026-06-02 calibration
above for how to handle the "spent" part — keep it honest/small or omit, never inflated.)

When starting any chunk of work (single edit, phase, sprint, multi-step task), state an
estimated time up front. As work progresses, give a "spent X min / remaining Y min" update
at natural check-in points — at each commit, sprint completion, or major transition.

**Why:** User wants to decide whether to step away vs. wait, and wants visibility into
whether my estimates are accurate so they can calibrate trust. Established 2026-05-28 at
the start of Phase 6 after watching Phases 3-5 go faster than expected without time
visibility.

**How to apply:**
- Phase-level estimate at the start of a phase ("Phase 6 estimated at ~15 min wall-clock")
- Sprint-level estimate as part of the kickoff for each sprint
- After each commit or sprint completion: "Spent X min / remaining ~Y min on this phase"
- Be honest. Don't pad. Update the remaining estimate if I'm clearly running ahead or
  behind, and explain why.
- This is a process preference, not a metric to game — if a task genuinely needs more
  time than I estimated, say so rather than rushing to hit the number.
- **ESTIMATES ARE INFORMATIONAL, NEVER A CONSTRAINT (clarified 2026-06-01).** Never cut
  scope, skip a fix, or rush to fit within an estimate. If extra work surfaces mid-task
  (a bug, broken path, cleanup, a discrepancy worth surfacing), fix it properly even if it
  means running over. Just flag that we're over and why. Correctness and completeness
  always win over hitting the number.

**Estimate units: WALL-CLOCK FROM USER'S PERSPECTIVE.** NOT "how long would a human
take to type this code." Tool calls fire in parallel, files write instantly, multi-
sprint phases compress dramatically.

Calibration from Phase 6:
- Original estimate: 75 minutes (mistakenly using "if a human typed this" units)
- Actual wall-clock: ~14 minutes from "start sprint 1" to "Phase 6 close" commit
- Ratio: ~5x overestimate. Going forward, divide my "human" gut estimate by 5
  to get a reasonable wall-clock number for the user. Or estimate directly in
  wall-clock minutes from the start.

Better defaults for future estimates:
- Single-file frontend edit + commit: 1-2 min
- New Edge Function scaffold + write + deploy + commit: 2-3 min
- Phase scoping conversation: 2-3 min
- Multi-sprint phase (4-6 sprints, no debug): 10-20 min wall-clock
- Multi-sprint phase WITH RLS/auth debugging: add 10-15 min
- End-to-end verification by the user (their hands): not counted in my estimate;
  acknowledge separately.
