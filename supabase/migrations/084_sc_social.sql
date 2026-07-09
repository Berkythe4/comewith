-- =============================================================================
-- 084_sc_social.sql
-- Log of follow / repost actions taken on the connected SoundCloud account, so
-- the UI can show what's already done (✓), avoid duplicates, and give a history.
-- The actions themselves run through the sc-social edge fn using the stored
-- OAuth token. Admin RLS.
-- =============================================================================
begin;

create table if not exists public.sc_social_log (
  id uuid primary key default gen_random_uuid(),
  action text not null check (action in ('follow','unfollow','repost','unrepost')),
  target_type text not null check (target_type in ('user','track')),
  target_id text not null,          -- SoundCloud numeric id
  target_label text,                -- artist name / track title for display
  ok boolean not null default true,
  detail text,
  created_at timestamptz not null default now()
);
create index if not exists idx_sc_social_target on public.sc_social_log(target_type, target_id);

alter table public.sc_social_log enable row level security;
drop policy if exists "Admins manage sc_social_log" on public.sc_social_log;
create policy "Admins manage sc_social_log" on public.sc_social_log for all using (public.is_admin()) with check (public.is_admin());

commit;
-- POST: follow/repost history table (anon-blocked).
