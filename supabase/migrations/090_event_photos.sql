-- =============================================================================
-- 090_event_photos.sql
-- Event photo galleries: admins attach photos to events in the hub; the public
-- site shows them on event.html (linked from the homepage "Recent rooms").
--
-- Storage: existing PUBLIC `event-photos` bucket (already holds artist photos +
-- event heroes; "Public can read event photos" / "Admins can manage event
-- photos" storage policies already exist). Two sizes are uploaded per photo,
-- generated client-side: storage_path (~1600px full) + thumb_path (~480px grid
-- thumbnail) — the grid serves thumbs to keep egress inside the Free plan.
--
-- Security follows the 030 pattern: the TABLE is admin-only (RLS + anon grant
-- stripped, since 013 default privileges auto-grant ALL to anon on new tables);
-- anon reads ONLY the dedicated view, which exposes photos of publicly surfaced
-- events (featured recaps or public upcoming) with is_public=true.
-- =============================================================================
begin;

create table if not exists public.event_photos (
  id           uuid primary key default gen_random_uuid(),
  event_id     uuid not null references public.events(id) on delete cascade,
  storage_path text not null,
  thumb_path   text,
  caption      text,
  sort_order   int  not null default 100,
  is_public    boolean not null default true,
  created_at   timestamptz not null default now(),
  created_by   uuid default auth.uid()
);
comment on table public.event_photos is
  'Photo gallery entries per event. Files live in the public event-photos bucket (paths are URL-reachable regardless of is_public — is_public only gates what the site LISTS). storage_path = ~1600px full, thumb_path = ~480px grid thumb.';
create index if not exists event_photos_event_idx
  on public.event_photos (event_id, sort_order, created_at);

alter table public.event_photos enable row level security;
drop policy if exists "Admins manage event photos" on public.event_photos;
create policy "Admins manage event photos" on public.event_photos
  for all using (public.is_admin()) with check (public.is_admin());

-- Least privilege: strip 013's auto-grant from the table — anon reads the view only.
revoke all on public.event_photos from anon;

create or replace view public.v_public_event_photos as
  select p.event_id, p.storage_path, p.thumb_path, p.caption, p.sort_order
  from public.event_photos p
  join public.events e on e.id = p.event_id
  where p.is_public = true
    and e.deleted_at is null
    and (e.is_featured = true or e.is_public = true);

comment on view public.v_public_event_photos is
  'Anon-readable photo feed for event.html. Photos flagged is_public on events that are publicly surfaced (is_featured recaps or is_public upcoming). No internal fields.';

revoke all    on public.v_public_event_photos from anon, authenticated;
grant  select on public.v_public_event_photos to   anon, authenticated;

commit;
