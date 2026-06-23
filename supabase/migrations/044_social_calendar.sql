-- =============================================================================
-- 044_social_calendar.sql
-- Social content calendar: posts with a workflow pipeline + threaded,
-- timestamped notes for Keith <-> Janelle collaboration.
--
-- ADDITIVE, NON-FINANCIAL — safe to apply to prod (same class as 041). Two new
-- leaf tables (no cross-module coupling), gated by the existing module system:
-- access == public.user_can_access_module('social-calendar'). master_admin
-- always; marketing/full once the module is signed_off (flipped on below).
--
-- Field set follows 2026 content-calendar best practices (status pipeline,
-- multi-channel, content pillar/tags, owner, asset link+status, target date,
-- optional event link, CTA).
-- =============================================================================
begin;

-- 1. Posts --------------------------------------------------------------------
create table if not exists public.social_posts (
  id            uuid primary key default gen_random_uuid(),
  title         text not null,
  caption       text,                                   -- the post copy
  channels      text[] not null default '{}',           -- instagram|tiktok|facebook|x|youtube|email|blog|other
  series        text,                                   -- 'Come With Parties' | 'Dance Infusion' | null (general)
  content_pillar text,                                  -- free-text tag (announcement, recap, promo, BTS…)
  stage         text not null default 'idea'
                  check (stage in ('idea','drafted','review','planned','scheduled','posted','archived')),
  scheduled_for timestamptz,                            -- target publish date/time
  posted_at     timestamptz,                            -- set when actually posted
  owner_id      uuid references public.profiles(id) on delete set null,  -- assignee
  event_id      uuid references public.events(id) on delete set null,    -- optional tie to an event
  link_url      text,                                   -- destination / live post URL
  asset_url     text,                                   -- creative file/link
  asset_status  text default 'none'
                  check (asset_status in ('none','requested','in_progress','ready')),
  cta           text,
  created_by    uuid references public.profiles(id) default auth.uid(),
  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now(),
  deleted_at    timestamptz
);

create index idx_social_posts_stage on public.social_posts(stage) where deleted_at is null;
create index idx_social_posts_scheduled on public.social_posts(scheduled_for);
create index idx_social_posts_event on public.social_posts(event_id);

create trigger set_updated_at
  before update on public.social_posts
  for each row execute function public.handle_updated_at();

alter table public.social_posts enable row level security;
create policy "Social calendar module access" on public.social_posts for all
  using (public.user_can_access_module('social-calendar'))
  with check (public.user_can_access_module('social-calendar'));

-- 2. Threaded notes (conversation history) ------------------------------------
create table if not exists public.social_post_notes (
  id          uuid primary key default gen_random_uuid(),
  post_id     uuid not null references public.social_posts(id) on delete cascade,
  author_id   uuid references public.profiles(id) default auth.uid(),
  body        text not null,
  created_at  timestamptz not null default now()
);

create index idx_social_post_notes_post on public.social_post_notes(post_id, created_at);

alter table public.social_post_notes enable row level security;
-- Notes inherit the module gate. Editing/deleting a note is allowed only to its
-- author or a master_admin; everyone with module access can read + add notes.
create policy "Notes read+add via module" on public.social_post_notes for select
  using (public.user_can_access_module('social-calendar'));
create policy "Notes insert via module" on public.social_post_notes for insert
  with check (public.user_can_access_module('social-calendar'));
create policy "Notes edit own or master" on public.social_post_notes for update
  using (author_id = auth.uid() or public.is_master_admin());
create policy "Notes delete own or master" on public.social_post_notes for delete
  using (author_id = auth.uid() or public.is_master_admin());

-- 3. Release the module: it is now built, and signed off so marketing (Janelle)
--    sees it. (Keith can un-sign from the Team tab anytime.)
update public.module_registry
   set built = true, signed_off = true, signed_off_at = now()
 where key = 'social-calendar';

commit;

-- =============================================================================
-- POST-APPLY VERIFICATION:
--   * select built, signed_off from module_registry where key='social-calendar'; -- t,t
--   * anon REST GET social_posts -> [] (RLS, module gate; not anon-readable).
--   * master can insert a post + a note; note.author_id defaults to the caller.
-- ROLLBACK:
--   drop table if exists public.social_post_notes;
--   drop table if exists public.social_posts;
--   update public.module_registry set built=false, signed_off=false,
--     signed_off_at=null where key='social-calendar';
-- =============================================================================
