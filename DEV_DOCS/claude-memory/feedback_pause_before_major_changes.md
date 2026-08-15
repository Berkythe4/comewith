---
name: feedback-pause-before-major-changes
description: "Martin (2nd operator, less experienced than Keith) — pause and review severity before any major structural change; minor changes proceed normally"
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 279f1814-6e4e-4f39-a354-6447da747ad9
  modified: 2026-07-30T01:44:05.619Z
---

**Martin Kearney** also works in this repo with me, not just Keith — `martin@comewith.org`,
role `sub_admin`, alias-tagged to the KRNeY actor record. He told me on 2026-07-30 he is
**less experienced at building and making changes** than Keith.

His instruction: **before any really major structural change, pause and review the severity
of the potential next steps together.** Minor, contained changes (a read-time filter, a
rescan, a UI tweak) should just proceed — he explicitly said those are fine.

**Why:** he can't always tell from the outside whether a proposed step is routine or
risky, so an unrequested big change costs him the chance to weigh it. A one-line severity
call up front gives him that.

**How to apply:** lead with a plain severity read before acting — what it touches, whether
it's reversible, and what happens if it's wrong. Treat as MAJOR (pause first): schema
migrations that drop/rename/backfill, anything touching RLS or the financial gate, edits
to live/published episodes or public pages, deleting rows, deploying something that changes
customer-facing behaviour, force-pushes. Treat as MINOR (just do it, mention it): additive
columns, read-time filters, UI/CSS, rescans and other re-runnable data refreshes, new
migrations that only add.

Do not assume authorisation carries over from Keith — check who is actually talking. Keith
is `master_admin`; Martin is `sub_admin` and does not see company financials unless
released. Related: [[project_staff_access_model]], [[project_user_management]].
