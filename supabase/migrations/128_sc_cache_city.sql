-- 128: store the SoundCloud profile city on the scan cache so the bulk of
-- artists (who are read via sc-enrich, not sc-match) get an NYC-local signal.
-- Additive; sc-enrich fills it, the dashboard reads it for the 📍 NYC tag/filter.
alter table public.sc_artist_cache add column if not exists city text;
