-- =============================================================================
-- 028_test_checklist_state.sql  —  persistence for the reusable test checklist
-- (Supabase-backed, NOT localStorage — same admin-only pattern as dashboard_prefs.)
-- NOT APPLIED — review before apply. Push held.
-- =============================================================================
create table public.test_checklist_state (
  test_key    text primary key,
  checked     boolean not null default false,
  notes       text,
  updated_by  uuid default auth.uid(),
  updated_at  timestamptz not null default now()
);
alter table public.test_checklist_state enable row level security;
create policy "Admins can manage test checklist state" on public.test_checklist_state for all using (public.is_admin());
-- Grants inherited from 013 default privileges. NO blanket anon grant.
-- DOWN: drop table public.test_checklist_state;
