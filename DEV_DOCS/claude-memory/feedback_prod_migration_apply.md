---
name: feedback_prod_migration_apply
description: "Apply prod migrations via the Supabase Management API using SBP_PAT in .env — not manual SQL-Editor paste, not the CLI (it's linked to staging)"
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 90f54be3-1f54-4bc6-ba07-0bdfee57a183
---

Prod DDL is applied via the **Supabase Management API**, the same way 023–030 were:
`POST https://api.supabase.com/v1/projects/$SBP_REF_PROD/database/query` with
`Authorization: Bearer $SBP_PAT` and body `{"query": "<sql>"}`. Token + prod ref +
prod URL/publishable key live in `.env` (`SBP_PAT`, `SBP_REF_PROD`,
`SUPABASE_PROD_URL`, `SUPABASE_PROD_PUBLISHABLE_KEY`). DDL returns `[]` / HTTP 201
on success. `jq` is NOT installed — build the JSON payload with Python
(`/c/Python314/python`, `json.dumps`).

**Why:** the `supabase` CLI is linked to **staging** (`qjosjafbizxdtkisyrqm`), so
`db push` / `migration up` would hit the wrong project (and staging's remote
migration table is empty, so push would try to re-apply everything). The Management
API is the only working prod path and has run 7+ times. Don't fall back to "paste
into SQL Editor" — Keith considers that making him do it manually.

**How to apply:** read the migration file, POST it via the Management API with the
SBP_PAT, then verify with anon REST (publishable key) — see [[feedback_no_broad_anon_grants]]
and [[project_kpi_layer]] for the anon-401 financial-view checks.
