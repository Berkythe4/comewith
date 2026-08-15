---
name: project_user_management
description: Revamped Users tool — staff profiles, alias tagging (staff↔actor), gig KPIs, activity log; master-only
metadata:
  type: project
  originSessionId: 23f44bb5-a672-44a4-9c2e-b8eac9975d80
---

The **Team** module was rebuilt into a tabbed **Users** management tool (master-only; module key still `team`, label renamed to "Users"), 2026-06-26, migrations 058+059, commit 97c62f2.

**Three tabs (`TEAM.tab`):**
- **People** — `renderTeamPeople`: user list + KPI cards (users / signed-in / linked aliases / alias gigs). Each row → `openUserDetail`.
- **Activity log** — `renderTeamActivity`/`loadActivity`: reads `audit_log`, filterable by **date range + user + table** (server-side re-query, limit 500; defaults to last 30 days on first open).
- **Modules** — `renderTeamModules`: the original sign-off/release control (unchanged).

**User detail** (`renderUserDetail`): last sign-in / member-since / password / invite status; editable name+phone (`saveUserProfile`); **gig KPIs** (gigs / fees earned / last gig) + gig history; recent activity; access (scope select + per-module override chips); **Aliases** — link a staff login to their other actor records (e.g. KRNeY the DJ → Martin) via `actors.user_id`, searchable picker (`openLinkAlias`) + `unlinkAlias`. Gig KPIs come from `event_participants` via the linked aliases.

**Backend:**
- **058 → superseded by 094 (2026-07-15):** `v_team_members` (profiles ⨝ auth.users, owner-run, `WHERE is_master_admin()`) tripped Supabase's `auth_users_exposed` security advisory — any API-exposed view reading auth.users granted to `authenticated` gets flagged regardless of row gating. **094 dropped the view and replaced it with RPC `public.get_team_members()`** (SECURITY DEFINER, same columns + master-only gate, EXECUTE revoked from public/anon, granted to authenticated/service_role); dashboard uses `sb.rpc('get_team_members')`. Applied to prod + verified 2026-07-15 (anon rpc → 401, old view endpoint → 404), commit 7b6d43e. Rule: never expose auth.users through a public-schema view — use a definer RPC. After creating a view/function you may need `notify pgrst, 'reload schema'` or PostgREST 404s it (PGRST205). Indexes on actors.user_id + audit_log(occurred_at/actor_id/table_name) still from 058.
- **059** — extended `audit_trigger_function()` (shared, SECURITY DEFINER, logs auth.uid()+email; needs an `id` column) to 10 more tables so the activity log is comprehensive: actors, events, venues, sponsorships, ticketing, third_party_donations, inquiries, social_posts, conversations, equipment_inventory (now 17 audited total). `audit_log` is master-only read (no new exposure). Skipped high-volume/low-signal tables (conversation_messages, social_post_notes, metric_snapshots, guest_event_attendance). Keep dashboard `AUDIT_TABLES_LIST` in sync with the audited set.

**Verified e2e:** master-only view (staff 0 / anon 401); alias link → KRNeY's 6 gigs rolled up under Martin (then unlinked clean — real KRNeY↔Martin link is Keith's call); new audit trigger logged martin's actor update as martin. `actors.user_id` link/unlink goes through can_see_people() RLS (042; master passes). See [[project_staff_access_model]], [[project_actor_model_and_equipment]].
