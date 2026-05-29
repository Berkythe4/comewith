-- ============================================================
-- COME WITH — 016 FEEDBACK LOG  (additive)
-- Capture-while-you-work scratchpad for Keith + Liz. Not a tracker.
-- Admin-only, matching the project convention (public.is_admin()).
-- ============================================================
begin;

create table if not exists public.feedback_log (
  id          uuid primary key default gen_random_uuid(),
  created_at  timestamptz not null default now(),
  created_by  uuid default auth.uid(),
  type        text not null check (type in ('bug','enhancement','idea')),
  note        text not null,
  status      text not null default 'open' check (status in ('open','done')),
  page        text
);

create index if not exists feedback_log_status_created
  on public.feedback_log (status, created_at desc);

alter table public.feedback_log enable row level security;

create policy "Admins can manage feedback_log"
  on public.feedback_log for all using (public.is_admin());

-- idempotent grants (mirrors 013); RLS still gates row access
grant usage on schema public to anon, authenticated, service_role;
grant all on all tables in schema public to anon, authenticated, service_role;

commit;
