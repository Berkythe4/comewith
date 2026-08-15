---
name: feedback-no-broad-anon-grants
description: "In Supabase migrations, never use broad `grant all on all tables to anon` — it silently re-exposes views revoked earlier"
metadata: 
  node_type: memory
  type: feedback
  originSessionId: d97501fd-d79e-460c-80ae-ea0889c23091
---

In a Supabase migration, do NOT include `grant all on all tables in schema public to anon` (or to authenticated). It re-grants SELECT on every table AND view, silently undoing earlier targeted `revoke select ... from anon` statements.

**Why:** On comewith prod 2026-05-29, migrations 016 and 017 each included that broad grant as "idempotent, mirrors 013." It re-exposed the 4 KPI financial views that 015 had revoked (decision E1) — anon could read `v_kpi_parties` etc. again. Caught in post-apply verification (anon GET returned 200 not 401); fixed with corrective migration 019 re-revoking. See [[project-kpi-layer]].

**How to apply:** New tables created in a migration already inherit anon/authenticated/service_role grants from 013's `ALTER DEFAULT PRIVILEGES` — no explicit grant block is needed. If you ever must re-grant, immediately re-assert every prior `revoke ... from anon` after it. Always verify anon access (expect 401 on revoked views) in the post-apply check.
