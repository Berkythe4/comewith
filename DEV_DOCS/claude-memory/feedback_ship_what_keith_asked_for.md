---
name: feedback-ship-what-keith-asked-for
description: "Work Keith asked for IS green-lit — commit and push it, and always state deploy status; he checks the live dashboard, not the working tree"
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 8ebdc8d3-10da-4c31-9305-ba69f006c2f9
  modified: 2026-08-25T16:06:03.686Z
---

When Keith asks for a change, that request **is** the green light. Build it,
commit it, push it. `master` auto-deploys to Netlify, so pushing is what makes it
real for him.

**Why:** on 2026-08-25 the event-hub tasks board was built, tested and verified,
then deliberately left uncommitted because "master auto-deploys and Keith hasn't
green-lit it" (the CLAUDE.md rule). He went to Dance Infusion → Tasks, saw the
**old** UI, and said "whatever you did I don't really understand and it is not
right." Nothing was wrong with the build. It had never left the working tree. The
rule about parking un-green-lit work on a branch exists for work he did *not* ask
for — speculative refactors, half-finished migrations — not for the thing he just
requested.

**How to apply:**
- Feature he asked for → commit and push. Say "pushed, live in a minute or two
  after Netlify builds, hard-refresh."
- Something built on your own initiative, or a migration he hasn't seen → branch,
  and name it in CARRYOVER under "Parked / next".
- Either way, **end the message with the deploy state in plain words.** "Done"
  and "visible to Keith" are different states and only he can tell them apart.
  Never let him discover it by clicking.
- He verifies in the browser on the live site, so "the tests pass" is not the
  same claim as "you can see it now".

Related: [[project-two-machine-handoff]] (master auto-deploys; held work goes on a
branch), [[feedback-time-estimates]].
