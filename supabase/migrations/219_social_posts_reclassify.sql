-- =============================================================================
-- 219_social_posts_reclassify.sql
-- Data only. Re-classifies LIVE social posts (Keith's sprint, 2026-10-02):
--   radio  : title has CWR/radio or series 'Come With Radio' -> phase radio,
--            format from the title (recap -> carousel, release/Ep<n> -> reel)
--   DI     : upcoming DI posts -> collab + awareness (all four are also named
--            individually, and those specific values win)
--   4 named posts retitled + classified.
-- Built BY ID from scripts/plan_social_reclassify.py run over
-- backups/social_posts_pre_reclassify_2026-10-02.json - the dry run printed the
-- full before/after table from the same plan. 29 rows, all live.
--
-- Guards (asserted below, the transaction aborts on any failure):
--   * every UPDATE carries `deleted_at is null` and must hit exactly 1 row
--   * total rows changed = 29
--   * no deleted row changes at all (full-row checksum)
--   * caption, stage, scheduled_for and owner_id are identical on every row
--   * only title / account / format / phase differ, and only on the 29
-- =============================================================================
begin;

create temp table _pre on commit drop as
  select id, deleted_at is not null as deleted, scheduled_for,
         title as b_title, account as b_account, format as b_format, phase as b_phase,
         md5(row_to_json(p)::text) as full_row,
         md5(concat_ws('|', caption, stage, scheduled_for::text, owner_id::text)) as protected,
         md5(concat_ws('|', title, account, format, phase)) as classified
    from public.social_posts p;

do $$
declare n int; total int := 0;
begin
  update public.social_posts set phase = 'radio'
   where id = 'a14e526b-fff9-4499-ac22-3d0b92f1a959' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'a14e526b-fff9-4499-ac22-3d0b92f1a959'; end if; total := total + n;
  update public.social_posts set phase = 'radio', format = 'reel'
   where id = '796658c5-8e04-4012-b5e6-ae210774e245' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '796658c5-8e04-4012-b5e6-ae210774e245'; end if; total := total + n;
  update public.social_posts set phase = 'radio', format = 'carousel'
   where id = 'db4ff634-669f-4dc3-8000-83eac22eaa75' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'db4ff634-669f-4dc3-8000-83eac22eaa75'; end if; total := total + n;
  update public.social_posts set phase = 'radio', format = 'reel'
   where id = 'f23cf591-ab5b-4022-a7c3-051c73f4067f' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'f23cf591-ab5b-4022-a7c3-051c73f4067f'; end if; total := total + n;
  update public.social_posts set phase = 'radio'
   where id = '65eda410-ede8-4c2e-9d77-eaedacf4c72a' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '65eda410-ede8-4c2e-9d77-eaedacf4c72a'; end if; total := total + n;
  update public.social_posts set phase = 'radio'
   where id = '9f30ea4f-bc94-4980-b9d4-89beba08581b' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '9f30ea4f-bc94-4980-b9d4-89beba08581b'; end if; total := total + n;
  update public.social_posts set format = 'reel'
   where id = 'c1606982-9d65-4524-a206-e1e0e34c27c9' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'c1606982-9d65-4524-a206-e1e0e34c27c9'; end if; total := total + n;
  update public.social_posts set phase = 'radio', format = 'carousel'
   where id = 'bf643590-0a0b-4b6e-91fd-ca60a122a20b' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'bf643590-0a0b-4b6e-91fd-ca60a122a20b'; end if; total := total + n;
  update public.social_posts set phase = 'radio', format = 'carousel'
   where id = '71285b39-fb20-4d35-961a-74d43fef8720' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '71285b39-fb20-4d35-961a-74d43fef8720'; end if; total := total + n;
  update public.social_posts set phase = 'radio'
   where id = '878d5bac-c9f2-49e7-871b-ba1d06c7ff05' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '878d5bac-c9f2-49e7-871b-ba1d06c7ff05'; end if; total := total + n;
  update public.social_posts set phase = 'radio'
   where id = '8cc8d330-0a5e-41f3-b10f-7fe1e812fbc7' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '8cc8d330-0a5e-41f3-b10f-7fe1e812fbc7'; end if; total := total + n;
  update public.social_posts set phase = 'radio', format = 'reel'
   where id = 'b995491c-6de4-49e5-98d3-48c8bf6cfaeb' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'b995491c-6de4-49e5-98d3-48c8bf6cfaeb'; end if; total := total + n;
  update public.social_posts set phase = 'radio', format = 'carousel'
   where id = '3f02f2dc-edaf-4f34-9d97-854c9c08c7ed' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '3f02f2dc-edaf-4f34-9d97-854c9c08c7ed'; end if; total := total + n;
  update public.social_posts set format = 'reel'
   where id = '15b35ac2-8cad-4649-9621-883d76aac7bf' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '15b35ac2-8cad-4649-9621-883d76aac7bf'; end if; total := total + n;
  update public.social_posts set format = 'reel'
   where id = '19a538f4-baff-4878-a2be-4702bd1ede9b' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '19a538f4-baff-4878-a2be-4702bd1ede9b'; end if; total := total + n;
  update public.social_posts set format = 'reel'
   where id = '20171702-a1b5-40f8-b422-95816d95461e' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '20171702-a1b5-40f8-b422-95816d95461e'; end if; total := total + n;
  update public.social_posts set phase = 'radio', format = 'carousel'
   where id = 'bc7cd825-4d5d-4085-b3cb-af1fcba291a7' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'bc7cd825-4d5d-4085-b3cb-af1fcba291a7'; end if; total := total + n;
  update public.social_posts set format = 'reel'
   where id = '2e4e5cad-0f97-40c6-a631-55dba518de71' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '2e4e5cad-0f97-40c6-a631-55dba518de71'; end if; total := total + n;
  update public.social_posts set phase = 'radio', format = 'carousel'
   where id = '22aa1dab-bbf7-45b2-9274-727981e6ac8d' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '22aa1dab-bbf7-45b2-9274-727981e6ac8d'; end if; total := total + n;
  update public.social_posts set phase = 'radio'
   where id = 'a8ca1d35-953f-450e-bb6a-057bea1e1d9b' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'a8ca1d35-953f-450e-bb6a-057bea1e1d9b'; end if; total := total + n;
  update public.social_posts set phase = 'radio', format = 'carousel'
   where id = 'c6cd4250-c48a-48f4-a5d3-d1d415d4602b' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'c6cd4250-c48a-48f4-a5d3-d1d415d4602b'; end if; total := total + n;
  update public.social_posts set phase = 'radio', format = 'reel'
   where id = '40aac9c3-c1e5-4076-80de-3cb403e82c01' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '40aac9c3-c1e5-4076-80de-3cb403e82c01'; end if; total := total + n;
  update public.social_posts set format = 'reel'
   where id = '41b17258-c926-4063-8ba8-6f86e01accf7' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '41b17258-c926-4063-8ba8-6f86e01accf7'; end if; total := total + n;
  update public.social_posts set format = 'reel'
   where id = '537b83d5-cbd0-4871-bceb-dc25df88e483' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '537b83d5-cbd0-4871-bceb-dc25df88e483'; end if; total := total + n;
  update public.social_posts set format = 'reel'
   where id = 'cf9bbb17-9eb1-4e04-8757-4846ff75400e' and deleted_at is null;  -- radio
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'cf9bbb17-9eb1-4e04-8757-4846ff75400e'; end if; total := total + n;
  update public.social_posts set title = 'DI3 Official Announcement', format = 'carousel', phase = 'awareness', account = 'collab'
   where id = 'c3a36e7b-0896-42ee-a36c-f0fe7d702086' and deleted_at is null;  -- specific
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'c3a36e7b-0896-42ee-a36c-f0fe7d702086'; end if; total := total + n;
  update public.social_posts set title = 'DI3 b2b: Miss Vee × Rainbow Tutu', format = 'reel', phase = 'awareness', account = 'collab'
   where id = '75ec2bbc-bb97-4da9-96c5-e8a03723001b' and deleted_at is null;  -- specific
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '75ec2bbc-bb97-4da9-96c5-e8a03723001b'; end if; total := total + n;
  update public.social_posts set title = 'DI3 Headliner: Emmy Adelle', format = 'reel', phase = 'awareness', account = 'collab'
   where id = 'b02c898b-55b6-42ad-9f9a-2719a5f5347f' and deleted_at is null;  -- specific
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', 'b02c898b-55b6-42ad-9f9a-2719a5f5347f'; end if; total := total + n;
  update public.social_posts set title = 'DI3 Sponsor spotlight + spots still open', format = 'carousel', phase = 'sponsors', account = 'di'
   where id = '3cbfda67-7766-4abe-a653-b288c1c8ed47' and deleted_at is null;  -- specific
  get diagnostics n = row_count; if n <> 1 then raise exception 'expected 1 live row for %', '3cbfda67-7766-4abe-a653-b288c1c8ed47'; end if; total := total + n;
  if total <> 29 then raise exception 'changed % rows, expected 29', total; end if;
end $$;

do $$
declare bad int;
begin
  -- zero deleted posts changed
  select count(*) into bad from _pre x join public.social_posts p using (id)
   where x.deleted and md5(row_to_json(p)::text) <> x.full_row;
  if bad <> 0 then raise exception '% deleted posts changed', bad; end if;
  -- captions, stages, dates and owners identical everywhere
  select count(*) into bad from _pre x join public.social_posts p using (id)
   where md5(concat_ws('|', p.caption, p.stage, p.scheduled_for::text, p.owner_id::text)) <> x.protected;
  if bad <> 0 then raise exception '% posts had caption/stage/date/owner changed', bad; end if;
  -- classification changed on exactly the planned ids and nowhere else
  select count(*) into bad from _pre x join public.social_posts p using (id)
   where (md5(concat_ws('|', p.title, p.account, p.format, p.phase)) <> x.classified)
         <> (p.id in ('a14e526b-fff9-4499-ac22-3d0b92f1a959', '796658c5-8e04-4012-b5e6-ae210774e245', 'db4ff634-669f-4dc3-8000-83eac22eaa75', 'f23cf591-ab5b-4022-a7c3-051c73f4067f', '65eda410-ede8-4c2e-9d77-eaedacf4c72a', '9f30ea4f-bc94-4980-b9d4-89beba08581b', 'c1606982-9d65-4524-a206-e1e0e34c27c9', 'bf643590-0a0b-4b6e-91fd-ca60a122a20b', '71285b39-fb20-4d35-961a-74d43fef8720', '878d5bac-c9f2-49e7-871b-ba1d06c7ff05', '8cc8d330-0a5e-41f3-b10f-7fe1e812fbc7', 'b995491c-6de4-49e5-98d3-48c8bf6cfaeb', '3f02f2dc-edaf-4f34-9d97-854c9c08c7ed', '15b35ac2-8cad-4649-9621-883d76aac7bf', '19a538f4-baff-4878-a2be-4702bd1ede9b', '20171702-a1b5-40f8-b422-95816d95461e', 'bc7cd825-4d5d-4085-b3cb-af1fcba291a7', '2e4e5cad-0f97-40c6-a631-55dba518de71', '22aa1dab-bbf7-45b2-9274-727981e6ac8d', 'a8ca1d35-953f-450e-bb6a-057bea1e1d9b', 'c6cd4250-c48a-48f4-a5d3-d1d415d4602b', '40aac9c3-c1e5-4076-80de-3cb403e82c01', '41b17258-c926-4063-8ba8-6f86e01accf7', '537b83d5-cbd0-4871-bceb-dc25df88e483', 'cf9bbb17-9eb1-4e04-8757-4846ff75400e', 'c3a36e7b-0896-42ee-a36c-f0fe7d702086', '75ec2bbc-bb97-4da9-96c5-e8a03723001b', 'b02c898b-55b6-42ad-9f9a-2719a5f5347f', '3cbfda67-7766-4abe-a653-b288c1c8ed47'));
  if bad <> 0 then raise exception '% posts changed outside (or missing from) the plan', bad; end if;
end $$;

commit;
