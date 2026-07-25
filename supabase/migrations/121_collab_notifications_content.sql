-- =============================================================================
-- 121_collab_notifications_content.sql
-- The collaboration layer: work in-system, get notified, keep everything tied
-- together. Three pillars:
--   1. notifications — "you've been asked to action X" → a bell, not a task.
--   2. content_assets — one store for full-length content AND social-ready
--      clips, attachable to an event and/or a post, surfaced in every layer.
--   3. social_posts brief/draft/review workflow — Keith briefs, Janelle drafts,
--      either can request feedback; the whole exchange lives on the post.
-- Admin-only, anon-revoked, audited.
-- =============================================================================
begin;

-- ---- 1. Notifications -------------------------------------------------------
create table if not exists public.notifications (
  id           uuid primary key default gen_random_uuid(),
  user_id      uuid not null references auth.users(id) on delete cascade,   -- recipient
  from_user_id uuid references auth.users(id) on delete set null,           -- who triggered it
  kind         text not null default 'fyi',   -- feedback_request / draft_request / assigned / mention / approved / fyi
  title        text not null,
  body         text,
  subject_type text,                            -- post / station / event / task / meeting
  subject_id   uuid,
  status       text not null default 'unread' check (status in ('unread', 'read', 'done')),
  created_at   timestamptz not null default now(),
  read_at      timestamptz,
  done_at      timestamptz
);
create index if not exists notifications_user_idx on public.notifications (user_id, status, created_at desc);

alter table public.notifications enable row level security;
-- You see and manage YOUR notifications; any admin can create one for a teammate.
drop policy if exists "See own notifications" on public.notifications;
create policy "See own notifications" on public.notifications for select
  using (user_id = auth.uid());
drop policy if exists "Create notifications" on public.notifications;
create policy "Create notifications" on public.notifications for insert
  with check (public.is_admin() and from_user_id = auth.uid());
drop policy if exists "Update own notifications" on public.notifications;
create policy "Update own notifications" on public.notifications for update
  using (user_id = auth.uid()) with check (user_id = auth.uid());
drop policy if exists "Delete own notifications" on public.notifications;
create policy "Delete own notifications" on public.notifications for delete
  using (user_id = auth.uid());
revoke all on public.notifications from anon;

-- ---- 2. Content assets (full-length + social-ready clips) -------------------
create table if not exists public.content_assets (
  id           uuid primary key default gen_random_uuid(),
  event_id     uuid references public.events(id) on delete cascade,
  post_id      uuid references public.social_posts(id) on delete set null,
  station_id   uuid references public.sc_playlists(id) on delete set null,   -- radio episode
  kind         text not null default 'full' check (kind in ('full', 'clip')),
  media        text not null default 'video' check (media in ('video', 'audio', 'image', 'other')),
  url          text,
  storage_path text,
  label        text,
  artist_id    uuid references public.actors(id) on delete set null,
  is_public    boolean not null default false,   -- also lives on the public site
  duration_note text,                            -- e.g. "0:30", "full set"
  created_by   uuid references auth.users(id) on delete set null,
  created_at   timestamptz not null default now(),
  updated_at   timestamptz not null default now()
);
create index if not exists content_assets_event_idx on public.content_assets (event_id);
create index if not exists content_assets_post_idx on public.content_assets (post_id) where post_id is not null;
create index if not exists content_assets_station_idx on public.content_assets (station_id) where station_id is not null;

alter table public.content_assets enable row level security;
drop policy if exists "Admins manage content_assets" on public.content_assets;
create policy "Admins manage content_assets" on public.content_assets for all
  using (public.is_admin()) with check (public.is_admin());
revoke all on public.content_assets from anon;
drop trigger if exists audit_content_assets on public.content_assets;
create trigger audit_content_assets after insert or update or delete on public.content_assets
  for each row execute function public.audit_trigger_function();

-- ---- 3. Post brief / draft / review workflow -------------------------------
alter table public.social_posts add column if not exists brief       text;   -- Keith's notes for the drafter
alter table public.social_posts add column if not exists draft_by    uuid references auth.users(id) on delete set null;
alter table public.social_posts add column if not exists review_by    uuid references auth.users(id) on delete set null;
alter table public.social_posts add column if not exists copy_status  text default 'brief'
  check (copy_status in ('brief', 'drafting', 'feedback', 'approved'));

commit;
