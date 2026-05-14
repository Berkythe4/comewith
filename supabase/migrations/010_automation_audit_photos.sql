-- =============================================================================
-- 010_automation_audit_photos.sql
-- Automation job registry, audit log, and photo metadata table.
-- =============================================================================

create table public.automation_jobs (
  id              uuid primary key default gen_random_uuid(),
  name            text not null unique,
  description     text,
  cron_expression text not null,
  edge_function   text not null,
  enabled         boolean not null default true,
  last_run_at     timestamptz,
  last_status     text,
  next_run_at     timestamptz,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now()
);

create trigger set_updated_at
  before update on public.automation_jobs
  for each row execute function public.handle_updated_at();

alter table public.automation_jobs enable row level security;

create policy "Admins can manage automation jobs"
  on public.automation_jobs for all
  using (public.is_admin());

-- =============================================================================
-- Automation runs — one row per execution
-- =============================================================================
create table public.automation_runs (
  id              uuid primary key default gen_random_uuid(),
  job_id          uuid references public.automation_jobs(id) on delete cascade,
  started_at      timestamptz not null default now(),
  completed_at    timestamptz,
  duration_ms     integer,
  status          text not null check (status in ('running', 'success', 'failed')),
  error_message   text,
  output          jsonb
);

create index idx_automation_runs_job_id on public.automation_runs(job_id);
create index idx_automation_runs_started_at on public.automation_runs(started_at desc);

alter table public.automation_runs enable row level security;

create policy "Admins can read automation runs"
  on public.automation_runs for select
  using (public.is_admin());

-- =============================================================================
-- Audit log — INSERT/UPDATE/DELETE on sensitive tables
-- =============================================================================
create table public.audit_log (
  id              bigserial primary key,
  table_name      text not null,
  row_id          text not null,
  action          text not null check (action in ('INSERT', 'UPDATE', 'DELETE')),
  actor_id        uuid references public.profiles(id),
  actor_email     text,
  old_data        jsonb,
  new_data        jsonb,
  occurred_at     timestamptz not null default now()
);

create index idx_audit_log_table_row on public.audit_log(table_name, row_id);
create index idx_audit_log_actor on public.audit_log(actor_id);
create index idx_audit_log_occurred_at on public.audit_log(occurred_at desc);

alter table public.audit_log enable row level security;

create policy "Master admin can read audit log"
  on public.audit_log for select
  using (public.is_master_admin());

-- Audit trigger function — attach to any table you want logged.
create or replace function public.audit_trigger_function()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  actor_email_val text;
begin
  -- Look up the actor's email from profiles (may be null for system actions)
  select email into actor_email_val from public.profiles where id = auth.uid();

  insert into public.audit_log (table_name, row_id, action, actor_id, actor_email, old_data, new_data)
  values (
    tg_table_name,
    coalesce((new.id)::text, (old.id)::text),
    tg_op,
    auth.uid(),
    actor_email_val,
    case when tg_op = 'INSERT' then null else to_jsonb(old) end,
    case when tg_op = 'DELETE' then null else to_jsonb(new) end
  );

  return coalesce(new, old);
end;
$$;

-- Attach audit triggers to sensitive tables.
create trigger audit_agreements
  after insert or update or delete on public.agreements
  for each row execute function public.audit_trigger_function();

create trigger audit_profiles
  after insert or update or delete on public.profiles
  for each row execute function public.audit_trigger_function();

create trigger audit_income
  after insert or update or delete on public.income
  for each row execute function public.audit_trigger_function();

create trigger audit_expenses
  after insert or update or delete on public.expenses
  for each row execute function public.audit_trigger_function();

-- =============================================================================
-- Photos — metadata for Supabase Storage objects
-- =============================================================================
create table public.photos (
  id              uuid primary key default gen_random_uuid(),
  event_id        uuid references public.events(id) on delete cascade,
  artist_id       uuid references public.artists(id) on delete cascade,
  equipment_id    uuid references public.equipment_inventory(id) on delete cascade,
  bucket          text not null,
  storage_path    text not null,
  caption         text,
  tags            text[] not null default '{}',
  taken_at        timestamptz,
  uploaded_by     uuid references public.profiles(id),
  width_px        integer,
  height_px       integer,
  size_bytes      integer,
  created_at      timestamptz not null default now(),
  deleted_at      timestamptz
);

create index idx_photos_event_id on public.photos(event_id);
create index idx_photos_artist_id on public.photos(artist_id);
create index idx_photos_tags on public.photos using gin(tags);

alter table public.photos enable row level security;

create policy "Admins can manage photos"
  on public.photos for all
  using (public.is_admin());

-- Event photos are public-readable for the event hub page.
create policy "Public can read event photos"
  on public.photos for select
  using (event_id is not null and deleted_at is null);
