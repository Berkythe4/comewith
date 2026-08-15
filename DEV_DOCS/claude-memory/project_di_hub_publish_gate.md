---
name: project-di-hub-publish-gate
description: Do NOT publish the Dance Infusion
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

The DI2 public hub page at `events/dance-infusion-2/index-v2.html` is
built and works against the get-event-hub Edge Function. It's currently
visible only on localhost.

**Hold:** Do NOT push it to comewith.org / production in Phase 11
cutover, do NOT share the URL externally, do NOT include it in any
"launched" announcement, until the user explicitly approves.

**Why:** User wants to complete the DI2 impact report (financial
reconciliation, MS Society impact, lessons learned for DI3) and
incorporate it into the hub before going live. Posting the hub with
incomplete content would undercut the story.

**How to apply:**
- When planning Phase 11 cutover, exclude the DI hub from the deploy
  manifest unless the user has lifted this hold
- If asked to "share the hub" or "preview the hub", confirm the user
  has approved before sharing any URL beyond localhost
- The dance-infusion-2/index-v2.html file stays in the repo; this is
  a *publish* gate, not a *build* gate

This memory can be deleted (or its contents inverted to "cleared to
publish") once the user gives the go-ahead.
