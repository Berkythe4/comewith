---
name: project-social-calendar-v2-connector
description: "Social Calendar v2 (migr 217-220) + social-mcp v5 claude.ai connector (6 tools incl. log_results) shipped 2026-10-02; two captions, hidden-not-dropped editor, secret-in-path MCP; connector added, Mon/Fri tasks live claude.ai-side"
metadata:
  node_type: memory
  type: project
  originSessionId: 70f036fe-3254-4ab0-9c2f-729329ffaf68
  modified: 2026-10-02T16:42:19.314Z
---

Shipped + deployed 2026-10-02 (commit 871f727; now social-mcp v5, 6 tools incl. log_results). Migration 217 added
account/format/phase/claude_caption/claude_drafted_at/results to social_posts, widened
stage with ready+approved, notes.author_name, connector_log. `caption` IS the final
caption. A trigger refuses claude_caption from any signed-in user (connector only).
218 restored 9 soft-deleted 'planned' rows 217's stage map wrongly moved.

Connector = edge fn `social-mcp`, URL `.../functions/v1/social-mcp/<SOCIAL_MCP_SECRET>`
(secret in desktop `.env` + Supabase secret; never commit). 6 tools (log_results = posted-only, results-only, merged); no deletes;
never `caption`; only edits stage idea/drafted. claude.ai request-header auth is beta,
so the secret rides in the path; `x-api-key` also accepted.

**Why:** Keith wants Claude to plan/draft from claude.ai + scheduled tasks while
Janelle owns the final copy.
**How to apply:** Keith added the connector 2026-10-02; the Mon/Fri tasks are claude.ai
scheduled tasks (prompts in reviews/session_2026-10-02.md). Editor result boxes must match
the connector's metrics or a save wipes one (LEARNINGS 86). Secret was rotated once. Re-run
`python scripts/e2e_social_mcp.py` after any connector change. Related:
[[project-social-calendar]] [[feedback-ship-what-keith-asked-for]]
