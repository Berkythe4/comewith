-- =============================================================================
-- 025_content_tags.sql  —  Phase C: content items + signature tags (additive)
-- Spec §1.6, §4.5.8, §6 Phase C. NOT APPLIED — review before apply. Push held.
-- =============================================================================

-- ----------------------------------------------------------------------------
-- content_items — individual pieces (video/reel), graded on views via snapshots
-- ----------------------------------------------------------------------------
create table public.content_items (
  id             uuid primary key default gen_random_uuid(),
  series_id      uuid references public.content_series(id) on delete set null,
  event_id       uuid references public.events(id) on delete set null,  -- nullable: standalone OR from an event
  title          text not null,
  platform       text,            -- youtube | instagram | tiktok | ...
  url            text,
  published_at   date,
  publish_status text not null default 'draft' check (publish_status in ('draft','review','published')),
  embed_on       jsonb not null default '[]'::jsonb,  -- e.g. ["homepage","event:<id>","di_report"]
  featured       boolean not null default false,
  notes          text,
  created_at     timestamptz not null default now(),
  updated_at     timestamptz not null default now(),
  deleted_at     timestamptz
);
create index idx_content_items_series on public.content_items(series_id);
create index idx_content_items_event on public.content_items(event_id);
create index idx_content_items_status on public.content_items(publish_status) where deleted_at is null;

create trigger set_updated_at before update on public.content_items
  for each row execute function public.handle_updated_at();

alter table public.content_items enable row level security;
create policy "Admins can manage content items" on public.content_items for all using (public.is_admin());
-- Published content is public (website surfacing). NOT financial — safe for anon.
create policy "Public can read published content" on public.content_items
  for select using (publish_status = 'published' and deleted_at is null);

-- ----------------------------------------------------------------------------
-- tags + taggables — generic, polymorphic; slice KPIs by signature
-- ----------------------------------------------------------------------------
create table public.tags (
  id          uuid primary key default gen_random_uuid(),
  name        text not null unique,
  kind        text,            -- 'signature' | 'theme' | ...
  created_at  timestamptz not null default now()
);
alter table public.tags enable row level security;
create policy "Admins can manage tags" on public.tags for all using (public.is_admin());

create table public.taggables (
  id            uuid primary key default gen_random_uuid(),
  tag_id        uuid not null references public.tags(id) on delete cascade,
  subject_type  text not null check (subject_type in ('event','content_item','actor')),
  subject_id    uuid not null,
  created_at    timestamptz not null default now()
);
create unique index idx_taggables_unique on public.taggables(tag_id, subject_type, subject_id);
create index idx_taggables_subject on public.taggables(subject_type, subject_id);

alter table public.taggables enable row level security;
create policy "Admins can manage taggables" on public.taggables for all using (public.is_admin());

-- Seed the signature tag mentioned in the spec.
insert into public.tags (name, kind) values ('booth-to-wall', 'signature')
on conflict (name) do nothing;

-- Grants: 013 default privileges; admin-only via RLS (+ public read of PUBLISHED
-- content_items only). No anon grants on tags/taggables.

-- DOWN: drop table public.taggables; drop table public.tags; drop table public.content_items;
