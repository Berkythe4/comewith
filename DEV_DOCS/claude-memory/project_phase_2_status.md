---
name: project-phase-2-status
description: "Phase 2 (auth bootstrap) closed 2026-05-28 on comewith-staging. Captures staging state, what was validated, and the open items Phase 3 needs to address."
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

Phase 2 closed 2026-05-28 on comewith-staging only (prod untouched).

## Staging state (live on Supabase, not in repo)
- `berky@comewith.org` exists in `auth.users`, `profiles.role = master_admin`
- `test+customer@comewith.org` exists as a permanent fixture for RLS testing
  (role=customer, no associated client/agreement after Phase 2 cleanup)
- Auth → Providers: magic-link email enabled, signups OFF, email confirmation ON
- Auth → URL Configuration: Site URL set to `http://localhost:3000` placeholder
- Schema-level grants applied (matching 013_grants.sql)

## Validated
- `handle_new_user` auth trigger creates profile rows correctly
- `auth.uid()`-based admin/customer RLS isolation works:
  Berky reads 72 audit_log rows; throwaway customer reads 0
- `is_admin()` / `is_master_admin()` helper functions evaluate correctly
  under impersonated JWT claims via `set_config('request.jwt.claim.sub', ...)`

## Open for Phase 3
- Anon-INSERT into `inquiries` fails RLS via SQL Editor `set local role anon`
  even with a valid PERMISSIVE policy. See [[project-anon-rls-sql-editor]] for
  full notes. First real inquiry submission from the supabase-js client is
  the validation gate.
- Real Site URL + production redirect URLs (currently localhost placeholder)
- Phase 3 itself is undefined; PHASE0_README hinted "frontend rewrites" but
  the scope hasn't been speced. Spar with user before building.

**Why:** Phase 2 produced dashboard + auth state that lives on Supabase's
servers, not in the repo. Without this memory, a future session would have
to ask the user "did Berky get provisioned?" / "is magic-link configured?"

**How to apply:** Before starting Phase 3 or any session touching auth, read
this to know what's already in place. The companion migration is
`supabase/migrations/013_grants.sql`.

Related: [[project-anon-rls-sql-editor]]
