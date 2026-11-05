-- 231: "Copy to my SoundCloud" for a list of songs that is not a station.
--
-- The Scene Map now shows, per club, the radio songs by artists who played there,
-- and offers that list as a SoundCloud playlist. The export plumbing (215/216:
-- sc-connect export_mine_start -> SoundCloud approval -> sc-oauth -> export_as)
-- already does exactly this for a station; the only thing tying it to a station is
-- sc_dj_exports.playlist_id NOT NULL.
--
-- So an export row now carries EITHER a station (playlist_id) OR its own track list
-- (track_ids + title), never neither. Additive: relaxing NOT NULL and adding two
-- nullable columns leaves the deployed dashboard, dj.html and both functions working
-- unchanged - every existing writer still sends playlist_id.

begin;

alter table public.sc_dj_exports
  add column if not exists track_ids text[],
  add column if not exists title     text;

alter table public.sc_dj_exports alter column playlist_id drop not null;

alter table public.sc_dj_exports drop constraint if exists sc_dj_exports_source_chk;
alter table public.sc_dj_exports add constraint sc_dj_exports_source_chk
  check (playlist_id is not null or (track_ids is not null and cardinality(track_ids) > 0 and title is not null));

commit;
