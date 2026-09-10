-- 211: when a feed sends no lineup, recover the artists we ALREADY TRACK from
-- the event title.
--
-- THE PROBLEM
-- 62% of future DICE events (344 of 555) and 35% of RA ones carry an empty
-- `lineup`. The show is stored, but nothing connects it to an artist, so it is
-- invisible to the artist pool, the watchlist coverage strip and buzz. Claptone
-- at Marquee Skydeck on 2026-10-03 is the worked example: pull-dice v22 captured
-- the event correctly, and Claptone -- who is ON THE WATCHLIST -- still showed as
-- having no upcoming NYC show, because his name exists only in the title.
--
-- WHY NOT JUST PARSE THE TITLE
-- Because most of these titles are not artist names. Measured over the 586
-- lineup-less future events: "CTRL ALT DLT", "FULLSOME", "Julie's Top 5 UK Music
-- Party", "WOW NYFW x SAFE - City of Stars". Deriving artists from those would
-- fill ra_artists with invented evidence -- the thing LEARNINGS SS26 forbids.
-- Matching every title against the whole 3,349-name artist pool was measured too
-- and is just as bad: it links "Stone Street Oktoberfest" to an artist called
-- Stone (5 times), "Alec Monopoly at the NYSC Summer Club" to Alec, and drags in
-- "Dreams", "Cosmo" and "Special Guest".
--
-- SO THE RULE IS INVERTED: never DERIVE a name from a title, only CONFIRM one we
-- already track. Same principle as the Bandsintown genre rule -- the answer is
-- guaranteed by WHO WE ASK ABOUT, not by parsing the response. Measured over the
-- same 586 events this fires 4 times, all four correct, no false positives:
-- Claptone (watchlist), Chris Lake, BLOND:ISH and Ragie Ban.
-- Low recall on purpose. We are not reconstructing 586 lineups; we are making
-- sure a show by someone we care about cannot go unlinked.
--
-- WHY A TRIGGER AND NOT A PULLER CHANGE
-- Two of the four hits are RA, not DICE. Fixing it inside pull-dice would solve
-- a third of it and then need copying into pull-ra-market and pull-ticketmaster
-- -- three copies of one rule, drifting. This sits where ra_events_resolve_venue
-- already sits: one implementation, every source, and it survives the delete-and-
-- reinsert that every puller does.
--
-- Additive. New function + view + trigger, and a backfill of existing rows.
-- The raw `title` is never modified; a lineup the feed actually sent is never
-- touched. Entries added here carry "via": "title" so they stay tellable apart
-- from what a feed supplied.

begin;

-- 1) One normaliser, not two. The folds a credit needs (accents, case, &/and,
--    punctuation, whitespace) are exactly the folds 208 already defined for
--    venues, and CLAUDE.md is explicit that a second copy is a thing to keep in
--    sync forever. This delegates rather than duplicating.
--    COUPLING, stated out loud: changing normalize_venue_name() changes credit
--    matching too. That is the price of having one implementation.
create or replace function public.normalize_credit(p_name text)
returns text
language sql
immutable
parallel safe
set search_path = public
as $$
  select public.normalize_venue_name(p_name);
$$;

comment on function public.normalize_credit(text) is
  'Normalise an artist credit for matching. Delegates to normalize_venue_name so there is exactly one set of folds. See migration 211.';

-- 2) The names we actually care about: the watchlist, our partners, and anyone
--    whose track has been on a station. Credits are split on the usual joiners
--    so "Lane 8, Kasablanca" contributes both names separately.
--    Names shorter than 4 characters after normalising are excluded -- they are
--    the ones that collide with ordinary title words.
create or replace view public.v_tracked_artists as
with src as (
  select label      as name, 'watchlist' as origin from public.watchlist
   where kind = 'artist' and not archived and label is not null
  union
  select name, 'partner' from public.ra_artists where is_partner and name is not null
  union
  select artist_name, 'played' from public.sc_playlist_tracks where artist_name is not null
), parts as (
  select btrim(p) as name, src.origin
    from src,
         lateral regexp_split_to_table(
           src.name,
           '\s*(?:,|/|\+|\mx\M|\mvs\.?\M|\mb2b\M|\mfeat\.?\M|\mft\.?\M|\mand\M|\mwith\M)\s*'
         ) as p
), keyed as (
  select public.normalize_credit(name) as key, name, origin
    from parts
   where btrim(coalesce(name, '')) <> ''
)
select distinct on (key) key, name, origin
  from keyed
 where key is not null
   and length(replace(key, ' ', '')) >= 4
 order by key, origin;

comment on view public.v_tracked_artists is
  'Artists we track by name (watchlist / partner / played on a station), normalised and split, for confirming a credit inside an event title. See migration 211.';

-- Internal strategy data. Views carry a table-level anon grant from 013 unless
-- revoked, and this one lists the watchlist.
revoke all on public.v_tracked_artists from anon;
grant select on public.v_tracked_artists to authenticated;

-- 3) Fill an EMPTY lineup with the tracked artists the title confirms.
create or replace function public.ra_events_link_known_artists()
returns trigger
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_title text;
  v_names jsonb;
begin
  -- A lineup the feed actually sent is evidence. Only ever fill a blank.
  if new.lineup is not null and jsonb_array_length(new.lineup) > 0 then
    return new;
  end if;

  -- Normalise SEGMENT BY SEGMENT, keeping the separators as a "|" marker.
  -- Punctuation is the only thing distinguishing a credit from part of a longer
  -- name, and plain normalising destroys it:
  --     "Scratch presents: Leo Vagnati, Eric Eric, Gianni Blanco"
  --     "Teksupport: Westend"
  -- both flatten to "<word> <name>", but the first is Gianni Blanco -- a
  -- different artist from the Blanco we track -- and the second really is
  -- Westend. Keeping the boundary tells them apart.
  select ' ' || string_agg(public.normalize_credit(seg), ' | ' order by ord) || ' '
    into v_title
    -- A colon only separates when a SPACE follows it: "presents: Westend" is a
    -- boundary, "BLOND:ISH" is a name. Splitting on every colon cut that artist
    -- in half and lost them.
    -- "&" and "+" are deliberately NOT separators -- "Above & Beyond" and
    -- "Sultan + Shepard" are single acts. The joining-word list below already
    -- covers a genuine "A & B", because normalise turns "&" into " and ".
    from regexp_split_to_table(coalesce(new.title, ''), '[,;/()\[\]|]|:\s|\s[-–—]\s')
         with ordinality as t(seg, ord)
   where btrim(coalesce(public.normalize_credit(seg), '')) <> '';

  if v_title is null or btrim(replace(v_title, '|', '')) = '' then
    return new;
  end if;
  v_title := regexp_replace(v_title, '\s+', ' ', 'g');

  -- Whole-word containment, plus: a SINGLE-WORD name must sit at the start of a
  -- segment or right after a joining word. That is what rejects "Blanco" inside
  -- "Gianni Blanco" while still accepting "Teksupport: Westend". Multi-word
  -- names carry their own evidence and need no such guard.
  select jsonb_agg(distinct jsonb_build_object('name', t.name, 'via', 'title'))
    into v_names
    from public.v_tracked_artists t
   where position(' ' || t.key || ' ' in v_title) > 0
     and (
       position(' ' in btrim(t.key)) > 0
       or v_title ~ ('(^ | \| | (?:presents?|pres|with|w|feat|ft|featuring|at|x|b2b|vs|versus|and|guest|only) )'
                     || t.key || ' ')
     );

  if v_names is not null then
    new.lineup := v_names;
  end if;
  return new;
end;
$$;

comment on function public.ra_events_link_known_artists() is
  'Fills an empty ra_events.lineup with tracked artists named in the title. Never overwrites a lineup a feed supplied. See migration 211.';

-- Postgres grants EXECUTE on a new function to PUBLIC, which anon inherits, so
-- "revoke ... from anon" alone is a silent no-op (LEARNINGS SS45). Revoke from
-- PUBLIC first. The trigger function needs no grant at all -- a trigger runs as
-- the table owner -- and normalize_credit is only ever called server-side.
revoke all on function public.normalize_credit(text) from public, anon;
grant execute on function public.normalize_credit(text) to authenticated;
revoke all on function public.ra_events_link_known_artists() from public, anon, authenticated;

drop trigger if exists ra_events_link_known_artists on public.ra_events;
create trigger ra_events_link_known_artists
  before insert or update on public.ra_events
  for each row execute function public.ra_events_link_known_artists();

-- 4) Backfill. Touching the row fires the trigger above; the venue trigger also
--    re-runs and is idempotent.
update public.ra_events
   set lineup = lineup
 where lineup is null or jsonb_array_length(lineup) = 0;

commit;
