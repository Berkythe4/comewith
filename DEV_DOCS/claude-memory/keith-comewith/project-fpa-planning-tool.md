---
name: project-fpa-planning-tool
description: "The Planning board's unit model is deliberately generic because Keith intends to rebuild it for a goods business (SKUs) next"
metadata: 
  node_type: memory
  type: project
  originSessionId: 9285a14c-2927-4aa5-9b36-c63f3a5610ad
  modified: 2026-08-22T01:29:15.158Z
---

The Planning board shipped on **2026-08-21** (migrations 197-202). Its schema is
generic on purpose, and the reason is not in the repo's own docs as an intention:

**Keith plans to build a version of this tool for a fashion company**, to decide
how many units of each SKU to order based on profitability. He said the structure
"could replicate into any service or goods" and that SKUs will need more built on
top, but the foundation should carry over.

That is why `plan_offerings` models a *unit of business* rather than an event:
`creates_event` is a flag, `scale` is abstract with a per-offering
`scale_label`, and the three line bases (`per_unit` / `per_scale` /
`pct_revenue`) were chosen to cover a SKU's economics as well as an event's.
When extending this, **do not collapse the abstraction back into event-specific
columns** — that forecloses the second product.

Also outstanding from that conversation, and not yet built: he wants a
forecastable unit to be **convertible into a real event** ("turned into / merged
into an event"). `plan_offerings.event_type` + `series` exist so a unit knows
what it would become; there is no button yet.

Related: [[user-fpa-background]]
