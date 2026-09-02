-- =============================================================================
-- 208_venue_normalization.sql
-- One static venue per real room, and a historical record that points at it.
--
-- THE PROBLEM, measured on prod before writing this:
--   * ra_events carries venue_name as FREE TEXT from three feeds and NO venue_id
--     at all, so 3,217 events have never been linked to a venue record.
--   * 406 distinct venue spellings in ra_events; 16 of them are the same room
--     written differently - 'Refuge' / 'REFUGE' / 'REFUGE ' (trailing space),
--     'Crossroads Cafe' / 'Crossroads Café', 'telos.haus' / 'Telos Haus',
--     'Dead Letter No. 9' / 'Dead Letter No.9', 'TBA - Brooklyn' / 'TBA Brooklyn'.
--   * The venues TABLE has the same disease: 5 collisions, including
--     'Acoustik Garden Lounge' entered twice.
--
-- THE SHAPE, and why it is two mechanisms rather than one:
--
--   1. DETERMINISTIC folding is applied AUTOMATICALLY. normalize_venue_name()
--      folds accents, case, '&'/'and', punctuation and whitespace. Two strings
--      that agree after that ARE the same room - there is no judgement in it, so
--      no human needs to confirm it.
--
--   2. FUZZY similarity only ever SUGGESTS. It is not safe to auto-merge, and the
--      data says so plainly: 'green room' vs 'green room 42' scores 0.87 and they
--      are different venues, while 'randall s island' vs 'randalls island' scores
--      0.97 and they are one. No threshold separates those two cases, so the
--      review queue exists and a person decides. A wrong merge silently rewrites
--      history, which is exactly what this migration is meant to protect.
--
-- Deliberately NOT folded: a leading "The", and the "TBA - " prefix the feeds put
-- on to-be-announced locations. Both are plausible folds and NEITHER appears as a
-- real collision in the data today, so they stay suggestions rather than rules -
-- a fold invented ahead of evidence is how two different rooms become one.
--
-- Unmatched names are NOT auto-created as venues. 203 of the 406 spellings have
-- no venue row, and many are not rooms at all ('TBA', 'listen', 'summer club');
-- pouring them into a table Keith curates by hand would be a worse mess than the
-- one being fixed. They surface in a review queue instead, ordered by how many
-- events they hold.
-- =============================================================================
begin;

-- Trigram similarity, for the SUGGESTION queue only. Supabase keeps extensions
-- out of public on purpose.
create schema if not exists extensions;
create extension if not exists pg_trgm with schema extensions;

-- ── 1. The deterministic key ─────────────────────────────────────────────────
-- IMMUTABLE because a generated column and an index both depend on it. Every
-- fold here is one this repo has actually seen in prod data; see the header.
create or replace function public.normalize_venue_name(p_name text)
returns text
language sql
immutable
parallel safe
as $$
  select nullif(
           btrim(
             regexp_replace(
               regexp_replace(
                 replace(
                   lower(
                     translate(
                       coalesce(p_name, ''),
                       'ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÑÒÓÔÕÖØÙÚÛÜÝàáâãäåçèéêëìíîïñòóôõöøùúûüýÿ',
                       'AAAAAACEEEEIIIINOOOOOOUUUUYaaaaaaceeeeiiiinoooooouuuuyy')),
                   '&', ' and '),
                 '[^a-z0-9]+', ' ', 'g'),
               '\s+', ' ', 'g')),
           '');
$$;

comment on function public.normalize_venue_name(text) is
  'Deterministic venue key: accents folded, lower-cased, & -> and, punctuation and whitespace collapsed. Two names equal after this ARE the same room. Deliberately does NOT strip a leading "The" or a "TBA - " prefix - neither is a real collision in the data, and an invented fold merges rooms that differ.';

revoke all on function public.normalize_venue_name(text) from public, anon;
grant execute on function public.normalize_venue_name(text) to authenticated, service_role;

-- ── 2. Merge the duplicates already inside `venues` ──────────────────────────
-- Survivor = most linked events, then the one carrying curated detail (capacity /
-- area / partner / links), then the oldest row. Chosen from the data rather than
-- hardcoded, so this reads the same way if it is ever re-run.
with ranked as (
  select v.id, public.normalize_venue_name(v.name) as nkey,
         row_number() over (
           partition by public.normalize_venue_name(v.name)
           order by (select count(*) from public.events e where e.venue_id = v.id) desc,
                    (case when v.capacity is not null then 1 else 0 end
                     + case when coalesce(v.area,'') <> '' then 1 else 0 end
                     + case when v.is_partner then 1 else 0 end
                     + case when coalesce(v.ra_url,'') <> '' then 1 else 0 end
                     + case when coalesce(v.ticket_url,'') <> '' then 1 else 0 end) desc,
                    v.created_at asc nulls last,
                    v.id asc) as rn
    from public.venues v
   where v.deleted_at is null
),
survivors as (select nkey, id as keep_id from ranked where rn = 1),
losers    as (select r.id as drop_id, s.keep_id
                from ranked r join survivors s on s.nkey = r.nkey
               where r.rn > 1),
-- Repoint everything that references a losing row BEFORE it disappears.
moved_events as (
  update public.events e set venue_id = l.keep_id
    from losers l where e.venue_id = l.drop_id returning 1),
moved_contacts as (
  update public.venue_contacts c set venue_id = l.keep_id
    from losers l where c.venue_id = l.drop_id returning 1)
-- Soft-delete rather than hard-delete: a venue row is history, and something not
-- caught by the two FKs above may still name it.
update public.venues v
   set deleted_at = now(),
       name = v.name || ' (merged duplicate)'
  from losers l
 where v.id = l.drop_id;

-- ── 3. The canonical key on venues ───────────────────────────────────────────
alter table public.venues
  add column if not exists name_norm text
  generated always as (public.normalize_venue_name(name)) stored;

-- Unique only across LIVE rows: the merged duplicates above keep their old key
-- and must not collide with the survivor.
create unique index if not exists venues_name_norm_live_uidx
  on public.venues (name_norm) where deleted_at is null and name_norm is not null;

-- ── 4. Aliases: every spelling that has been ruled on ────────────────────────
create table if not exists public.venue_aliases (
  alias_norm  text primary key,
  alias_raw   text not null,
  venue_id    uuid references public.venues(id) on delete cascade,
  -- 'linked'  = this spelling IS that venue
  -- 'ignored' = reviewed and deliberately not a venue ('TBA', 'listen', …), so it
  --             stops coming back to the top of the queue every week
  status      text not null default 'linked' check (status in ('linked', 'ignored')),
  source      text not null default 'manual' check (source in ('auto', 'manual')),
  note        text,
  created_at  timestamptz not null default now(),
  created_by  uuid references public.profiles(id),
  constraint venue_aliases_linked_needs_venue
    check (status <> 'linked' or venue_id is not null)
);

create index if not exists venue_aliases_venue_idx on public.venue_aliases (venue_id);

comment on table public.venue_aliases is
  'One row per venue spelling that has been ruled on. Deterministic folds never need a row here - normalize_venue_name() already makes them equal. This table is for the judgement calls: a misspelling, an abbreviation, a room name, or a string ruled NOT a venue at all (status = ignored).';

alter table public.venue_aliases enable row level security;
drop policy if exists "Admins manage venue aliases" on public.venue_aliases;
create policy "Admins manage venue aliases" on public.venue_aliases for all
  using (public.is_admin()) with check (public.is_admin());
revoke all on public.venue_aliases from anon;

-- ── 5. The resolver ──────────────────────────────────────────────────────────
-- Order matters: an explicit alias always beats the deterministic key, because an
-- alias is a decision somebody made and the key is only a guess that two strings
-- look alike.
create or replace function public.resolve_venue_name(p_name text)
returns uuid
language sql
stable
as $$
  with k as (select public.normalize_venue_name(p_name) as nkey)
  select coalesce(
    (select a.venue_id from public.venue_aliases a, k
      where a.alias_norm = k.nkey and a.status = 'linked'),
    (select v.id from public.venues v, k
      where v.name_norm = k.nkey and v.deleted_at is null limit 1)
  );
$$;

revoke all on function public.resolve_venue_name(text) from public, anon;
grant execute on function public.resolve_venue_name(text) to authenticated, service_role;

-- ── 6. The historical link ───────────────────────────────────────────────────
-- venue_name STAYS as the feed sent it. venue_id is the resolved room. Keeping
-- both is the point: the raw string is evidence, and rewriting it would destroy
-- the only record of what the source actually said.
alter table public.ra_events add column if not exists venue_id uuid references public.venues(id);
create index if not exists ra_events_venue_idx on public.ra_events (venue_id) where venue_id is not null;

comment on column public.ra_events.venue_id is
  'Resolved canonical venue. Set by the ra_events_resolve_venue trigger on write; null means the spelling has no venue yet and is sitting in v_venue_name_review.';

create or replace function public.ra_events_resolve_venue()
returns trigger
language plpgsql
as $$
begin
  if new.venue_name is distinct from coalesce(old.venue_name, '') or new.venue_id is null then
    new.venue_id := public.resolve_venue_name(new.venue_name);
  end if;
  return new;
end;
$$;

drop trigger if exists ra_events_resolve_venue on public.ra_events;
create trigger ra_events_resolve_venue before insert or update on public.ra_events
  for each row execute function public.ra_events_resolve_venue();

-- Backfill the history.
update public.ra_events e
   set venue_id = public.resolve_venue_name(e.venue_name)
 where e.venue_id is null and coalesce(btrim(e.venue_name), '') <> '';

-- ── 7. The review queue ──────────────────────────────────────────────────────
-- Every spelling with no venue and no ruling, heaviest first, with the closest
-- existing venues as SUGGESTIONS. similarity() is why pg_trgm is here; it never
-- decides anything on its own.
create or replace view public.v_venue_name_review as
with unresolved as (
  select public.normalize_venue_name(e.venue_name) as alias_norm,
         min(btrim(e.venue_name))                  as example_raw,
         count(*)                                  as events,
         min(e.event_date)                         as first_seen,
         max(e.event_date)                         as last_seen,
         string_agg(distinct e.source, '/')        as sources
    from public.ra_events e
   where coalesce(btrim(e.venue_name), '') <> ''
     and e.venue_id is null
   group by 1
)
select u.alias_norm, u.example_raw, u.events, u.first_seen, u.last_seen, u.sources,
       s.venue_id       as suggest_venue_id,
       s.venue_name     as suggest_venue_name,
       s.score          as suggest_score
  from unresolved u
  left join lateral (
    select v.id as venue_id, v.name as venue_name,
           extensions.similarity(v.name_norm, u.alias_norm) as score
      from public.venues v
     where v.deleted_at is null and v.name_norm is not null
       and extensions.similarity(v.name_norm, u.alias_norm) >= 0.45
     order by extensions.similarity(v.name_norm, u.alias_norm) desc, v.name
     limit 1) s on true
 where not exists (select 1 from public.venue_aliases a where a.alias_norm = u.alias_norm);

comment on view public.v_venue_name_review is
  'Venue spellings with no canonical room and no ruling yet, heaviest first, each with its closest existing venue as a SUGGESTION only. Nothing here is applied automatically - "green room" and "green room 42" score 0.87 and are different venues.';

revoke all on public.v_venue_name_review from anon;
grant select on public.v_venue_name_review to authenticated;

-- What the feeds and the venue book disagree about, for the same queue's header.
create or replace view public.v_venue_link_health as
select (select count(*) from public.ra_events where coalesce(btrim(venue_name),'') <> '')            as events_with_a_name,
       (select count(*) from public.ra_events where venue_id is not null)                            as events_linked,
       (select count(*) from public.v_venue_name_review)                                             as spellings_unresolved,
       (select count(*) from public.venue_aliases where status = 'linked')                           as aliases_linked,
       (select count(*) from public.venue_aliases where status = 'ignored')                          as aliases_ignored,
       (select count(*) from public.venues where deleted_at is null)                                 as venues_live;

revoke all on public.v_venue_link_health from anon;
grant select on public.v_venue_link_health to authenticated;

-- ── 8. Applying a ruling ─────────────────────────────────────────────────────
-- One function so the alias and the back-link to history always happen together.
-- Doing it as two client calls is how a venue gets an alias while its 34 events
-- keep pointing at nothing.
create or replace function public.link_venue_alias(
  p_alias_raw text,
  p_venue_id  uuid,
  p_note      text default null)
returns integer
language plpgsql
security definer
set search_path = public
as $$
declare
  v_norm text := public.normalize_venue_name(p_alias_raw);
  v_rows integer;
begin
  if not public.is_admin() then
    raise exception 'not authorised';
  end if;
  if v_norm is null then
    raise exception 'that name normalises to nothing';
  end if;
  if p_venue_id is not null and not exists (
       select 1 from public.venues where id = p_venue_id and deleted_at is null) then
    raise exception 'no live venue with that id';
  end if;

  insert into public.venue_aliases (alias_norm, alias_raw, venue_id, status, source, note, created_by)
  values (v_norm, btrim(p_alias_raw), p_venue_id,
          case when p_venue_id is null then 'ignored' else 'linked' end,
          'manual', p_note, auth.uid())
  on conflict (alias_norm) do update
    set venue_id = excluded.venue_id, status = excluded.status,
        alias_raw = excluded.alias_raw, note = excluded.note,
        created_by = excluded.created_by, created_at = now();

  -- Re-point every historical event that used this spelling, in the same call.
  update public.ra_events e
     set venue_id = p_venue_id
   where public.normalize_venue_name(e.venue_name) = v_norm
     and e.venue_id is distinct from p_venue_id;
  get diagnostics v_rows = row_count;
  return v_rows;
end;
$$;

revoke all on function public.link_venue_alias(text, uuid, text) from public, anon;
grant execute on function public.link_venue_alias(text, uuid, text) to authenticated;

-- Create a venue from an unmatched spelling and link it, in one step.
create or replace function public.create_venue_from_name(p_name text)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  v_id uuid;
  v_norm text := public.normalize_venue_name(p_name);
begin
  if not public.is_admin() then
    raise exception 'not authorised';
  end if;
  if v_norm is null then
    raise exception 'that name normalises to nothing';
  end if;
  select id into v_id from public.venues where name_norm = v_norm and deleted_at is null limit 1;
  if v_id is null then
    insert into public.venues (name) values (btrim(p_name)) returning id into v_id;
  end if;
  perform public.link_venue_alias(p_name, v_id, 'created from the review queue');
  return v_id;
end;
$$;

revoke all on function public.create_venue_from_name(text) from public, anon;
grant execute on function public.create_venue_from_name(text) to authenticated;

notify pgrst, 'reload schema';

commit;
