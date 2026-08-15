---
name: feedback_prod_migration_apply
description: "Apply prod migrations via the Supabase Management API using SBP_PAT in .env — not manual SQL-Editor paste, not the CLI (it's linked to staging)"
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 90f54be3-1f54-4bc6-ba07-0bdfee57a183
  modified: 2026-08-15T20:54:19.948Z
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
and [[project_kpi_layer]] for the anon-401 financial-view checks. `db.py` at the repo
root does this; pass the prod ref explicitly (`SBP_REF=yaytdosxfhcqatmhctzk python db.py "…"`)
because `.env`'s `SBP_REF` points at **staging**.

**Edge functions are the same story** (confirmed 2026-08-15). `supabase functions deploy`
is unusable here for a second reason on top of the staging link: CLI 2.101.0 rejects the
newer `sbp_v0_…` PAT format outright — *"Invalid access token format. Must be like
`sbp_0102...1920`"* — even though the Management API accepts that exact token. Deploy with
**`python scripts/deploy_edge_function.py <slug> [<slug>…]`** (multipart POST to
`/v1/projects/<ref>/functions/deploy?slug=…`). It reads the live function first and
preserves its `verify_jwt` — a deploy ships new code, it does not silently change who may
call the function. Verify by GETting `/functions/<slug>/body` and grepping for something
new in the source; the returned `version` also increments.
