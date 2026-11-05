-- 239: (Keith, 2026-11-05)
-- 1. "Pacha NYC Basement" is part of Pacha New York - fold the venue row in. Its one
--    listing is "Pacha NYC presents: BLOND:ISH special guest Luciano" (RA 2467966).
-- 2. Three SHOW 9 songs (drop 2026-09-24) were saved with no show. Keith: check the
--    remixers. Each now points at the show that put it in the crate, from ra_events:
--      Supermode - Tell Me Why (MEDUZA Remix)  -> MEDUZA, We Belong Here: Central Park,
--                                                 Wollman Rink, 2026-10-02 (RA 2476530)
--      Afrojack - Monopoly Man                 -> AFROJACK Presents: CONTROL, Brooklyn
--                                                 Storehouse, 2026-10-02 (DICE)
--      Phantogram - When I'm Small (Jkyl & Hyde Redo)
--                                              -> Phantogram (DJ Set), Elsewhere,
--                                                 2026-09-26 (RA 2518635). The credited
--                                                 artist's own show, a week inside the
--                                                 window; the remixers (Jkyl & Hyde) play
--                                                 Brooklyn Storehouse on 2026-10-24.
--    Each update must hit exactly one row whose show is still blank.

begin;

do $$
declare d public.venues; k public.venues;
begin
  if (select count(*) from public.venues where name = 'Pacha NYC Basement' and deleted_at is null) <> 1
     or (select count(*) from public.venues where name = 'Pacha New York' and deleted_at is null) <> 1 then
    raise exception 'expected one live Pacha NYC Basement and one Pacha New York';
  end if;
  select * into d from public.venues where name = 'Pacha NYC Basement' and deleted_at is null;
  select * into k from public.venues where name = 'Pacha New York' and deleted_at is null;
  insert into public.venue_aliases (alias_norm, alias_raw, venue_id, status, source, note)
  values (d.name_norm, btrim(d.name), k.id, 'linked', 'manual', '239: room of Pacha New York')
  on conflict (alias_norm) do update set venue_id = excluded.venue_id, status = 'linked', note = excluded.note;
  update public.venue_aliases set venue_id = k.id where venue_id = d.id;
  update public.ra_events     set venue_id = k.id where venue_id = d.id;
  update public.events        set venue_id = k.id where venue_id = d.id;
  update public.venue_contacts set venue_id = k.id where venue_id = d.id;
  update public.venues set deleted_at = now(), updated_at = now(),
         notes = concat_ws(E'\n', notes, '239: folded into Pacha New York (' || k.id || ')') where id = d.id;
end $$;

do $$
declare n int;
begin
  update public.sc_playlist_tracks e set show_date = x.d::date, show_venue = x.v, show_url = x.u
    from (values
      ('114a5b91-f5f6-4f07-9b6f-f15b17d19e99', '2026-10-02', 'Wollman Rink',        'https://ra.co/events/2476530'),
      ('0d2c5f7d-80ff-48e6-8eab-1f345bf0b6e1', '2026-10-02', 'Brooklyn Storehouse', 'https://dice.fm/event/3o756x-afrojack-presents-control-2nd-oct-brooklyn-storehouse-new-york-city-tickets'),
      ('57897315-9f86-4953-9934-de7ad7f7ee71', '2026-09-26', 'Elsewhere',           'https://ra.co/events/2518635')
    ) x(id, d, v, u)
   where e.id = x.id::uuid and nullif(btrim(e.show_venue), '') is null and e.show_date is null;
  get diagnostics n = row_count;
  if n <> 3 then raise exception 'expected 3 SHOW 9 songs with a blank show, updated %', n; end if;
end $$;

commit;
