-- 249: merge the two DI #3 event segments into one.
--
-- DI #3 attendees were tagged under two names: 'di-03-2026-11' (12 people, the
-- 2026-11-05 import, event-code style) and 'dance-infusion-3' (34 people,
-- 2026-11-11..16). 3 people hold both, so 43 distinct people.
--
-- Kept: 'dance-infusion-3' = events.slug, which is what both live importers write
-- (control-center.js eventSeg = event.slug; import-engine.js "event slug"). Merging
-- the other way would see the slug tag re-created by the next import.
-- No mailing_campaigns.segment_filter references either name (checked 2026-11-18).
-- Every one of the 43 already holds the dance_infusion brand segment; untouched.

begin;

do $$
declare n_people int; n_before int;
begin
  select count(distinct subscriber_id) into n_people
    from subscriber_segments where segment in ('di-03-2026-11','dance-infusion-3');
  if n_people <> 43 then raise exception 'expected 43 DI#3 people, found %', n_people; end if;
end $$;

-- people tagged only under the old name get the slug tag (keep their original added_at)
insert into subscriber_segments (subscriber_id, segment, added_at)
select o.subscriber_id, 'dance-infusion-3', min(o.added_at)
  from subscriber_segments o
 where o.segment = 'di-03-2026-11'
   and not exists (select 1 from subscriber_segments k
                    where k.subscriber_id = o.subscriber_id and k.segment = 'dance-infusion-3')
 group by o.subscriber_id;

delete from subscriber_segments where segment = 'di-03-2026-11';

do $$
declare n_rows int; n_people int; n_old int;
begin
  select count(*), count(distinct subscriber_id) into n_rows, n_people
    from subscriber_segments where segment = 'dance-infusion-3';
  select count(*) into n_old from subscriber_segments where segment = 'di-03-2026-11';
  if n_rows <> 43 or n_people <> 43 or n_old <> 0 then
    raise exception 'merge check failed: rows % people % old %', n_rows, n_people, n_old;
  end if;
end $$;

commit;
