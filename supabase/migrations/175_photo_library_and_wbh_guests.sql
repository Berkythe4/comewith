-- ============================================================
-- COME WITH — 175 right guests, and photos that do not need an event
--
-- 1. WRONG PEOPLE. 172 attached 'Zachary Storey' and a bare 'Amanda' to We
--    Belong Here. Both were wrong. It was Zach Mackey and Amanda Brundige —
--    VIP fans and friends, not crew. Zachary Storey is a different person who
--    came in through the 2026-06-16 guest ledger and should not be on this trip
--    at all.
--
-- 2. PHOTOS WITHOUT AN EVENT. event_photos requires an event_id, which was fine
--    while every photo came from a show. It does not survive first contact with
--    a press shoot: Keith paid Tori Mumtaz $250 to photograph HIM, and there is
--    no event those images belong to. Inventing one would put a photo session in
--    the events list and, worse, in the P&L alongside real shows.
--
--    So a photo now hangs off a SUBJECT as well as, or instead of, an event:
--
--        event_id          nullable  — the show, if there was one
--        subject_actor_id  new       — who or what the photos are OF
--
--    A photo needs at least one of the two, enforced below, so nothing can be
--    uploaded into nowhere and quietly become unfindable.
--
-- 3. PUBLIC BY DEFAULT WAS BACKWARDS. is_public defaulted to true, so every
--    upload was live on the site the moment it finished. For a press shoot that
--    is exactly wrong — these are a library to pick FROM, and publishing is a
--    decision, not the resting state. Default flips to false.
--
--    EXISTING PHOTOS ARE NOT TOUCHED. Anything already public stays public;
--    silently un-publishing live gallery images would be a worse surprise than
--    the default was.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. The actual guests
-- ---------------------------------------------------------------
insert into public.actors (display_name, kind, status, notes)
select v.name, 'person', 'active', v.note
  from (values
    ('Zach Mackey',     'VIP fan and friend. Guest at We Belong Here, June 2026.'),
    ('Amanda Brundige', 'VIP fan and friend. Guest at We Belong Here, June 2026.')
  ) as v(name, note)
 where not exists (
   select 1 from public.actors a
    where a.deleted_at is null and a.display_name = v.name);

-- The placeholder 'Amanda' created by 172 becomes the real record if it is still
-- unused elsewhere; otherwise it is retired rather than left as a decoy.
update public.actors
   set deleted_at = now(),
       notes = coalesce(notes, '') || ' [superseded by Amanda Brundige, 175]'
 where deleted_at is null
   and display_name = 'Amanda'
   and exists (select 1 from public.actors b
                where b.deleted_at is null and b.display_name = 'Amanda Brundige');

-- Off the trip: the placeholder and the wrong Zach.
delete from public.event_participants p
 using public.events e, public.actors a
 where e.id = p.event_id and a.id = p.actor_id
   and e.slug = 'we-belong-here-2026'
   and a.display_name in ('Amanda', 'Zachary Storey');

insert into public.event_participants (event_id, actor_id, role, notes)
select e.id, a.id, 'guest', 'VIP fan / friend — Keith''s guest'
  from public.events e
  join public.actors a on a.deleted_at is null
 where e.slug = 'we-belong-here-2026'
   and a.display_name in ('Angela Tabone', 'Zach Mackey', 'Amanda Brundige')
   and not exists (
     select 1 from public.event_participants p
      where p.event_id = e.id and p.actor_id = a.id and p.role = 'guest');

update public.events
   set notes = 'Keith won tickets. Brought Angela Tabone (Dance Infusion), Zach Mackey '
            || 'and Amanda Brundige — VIP fans and friends. No revenue: nothing was sold '
            || 'and the tickets cost nothing. Costs are travel and on-site spend only.'
 where slug = 'we-belong-here-2026';

-- ---------------------------------------------------------------
-- 2 + 3. Photos: optional event, required subject, private by default
-- ---------------------------------------------------------------
alter table public.event_photos
  alter column event_id drop not null,
  add column if not exists subject_actor_id uuid references public.actors(id),
  add column if not exists shoot_label text,
  alter column is_public set default false;

comment on column public.event_photos.subject_actor_id is
  'Who or what the photos are OF. Lets a press shoot exist without a fake event.';
comment on column public.event_photos.shoot_label is
  'Names a session that is not an event, e.g. "Press shoot — July 2026".';

alter table public.event_photos
  drop constraint if exists event_photos_needs_a_home;
alter table public.event_photos
  add constraint event_photos_needs_a_home
  check (event_id is not null or subject_actor_id is not null);

create index if not exists idx_event_photos_subject
  on public.event_photos (subject_actor_id)
  where subject_actor_id is not null;

-- The credits view has to cope with both shapes now. Dropped rather than
-- replaced: the new columns sit in the middle of the list, and CREATE OR REPLACE
-- can only append. Nothing depends on this view, so dropping is safe.
drop view if exists public.v_photo_credits;

create view public.v_photo_credits as
select
  p.id                                    as photo_id,
  p.event_id,
  coalesce(e.name, p.shoot_label, 'Unfiled') as event_name,
  e.event_date,
  p.subject_actor_id,
  s.display_name                          as subject,
  p.shoot_label,
  p.storage_path,
  p.thumb_path,
  p.caption,
  p.is_public,
  coalesce(p.shot_on, e.event_date)       as shot_on,
  p.photographer_actor_id,
  a.display_name                          as photographer,
  a.instagram                             as photographer_instagram,
  coalesce(p.credit_line, a.display_name) as credit,
  (p.photographer_actor_id is null)       as needs_credit
from public.event_photos p
left join public.events e on e.id = p.event_id and e.deleted_at is null
left join public.actors a on a.id = p.photographer_actor_id
left join public.actors s on s.id = p.subject_actor_id;

revoke select on public.v_photo_credits from anon;

commit;

-- DOWN: restore event_id NOT NULL (only possible once every subject-only photo is
--   removed), drop the constraint/index/columns, set is_public default back to
--   true, and re-point the We Belong Here participants.
