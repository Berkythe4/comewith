-- ============================================================
-- COME WITH — 159 merge duplicate vendor actors
--
-- 158's consolidation surfaced two actors that are the same party twice:
--
--   'Crossroads Café'  and  'Crossroads CafÃ©'   <- mojibake: é encoded twice
--   '19th & 7th Productions (Michael McManus)'  and  '19th & 7th Productions'
--
-- Both had expenses attached, so the P&L was splitting one vendor across two
-- lines. The mojibake one is the giveaway that a UTF-8 name went through a
-- Latin-1 step at some point.
--
-- Fourteen tables carry a foreign key to actors, so this repoints EVERY one of
-- them generically rather than assuming only expenses is involved — guessing
-- wrong would leave a dangling reference or silently orphan a row. The loser is
-- then soft-deleted, never hard-deleted, so the merge is reversible.
--
-- NOT merged: 'Jennifer Alderman' and 'Jennifer Taveras' share a first name and
-- are different people.
-- ============================================================
begin;

do $$
declare
  r record;
  pair record;
  merged int := 0;
begin
  for pair in
    select keep.id as keep_id, lose.id as lose_id,
           keep.display_name as keep_name, lose.display_name as lose_name
      from (values
        -- (surviving name, duplicate name)
        ('Crossroads Caf' || chr(233),            'Crossroads Caf' || chr(195) || chr(169)),
        ('19th & 7th Productions (Michael McManus)', '19th & 7th Productions')
      ) as v(keep_name, lose_name)
      join public.actors keep on keep.display_name = v.keep_name and keep.deleted_at is null
      join public.actors lose on lose.display_name = v.lose_name and lose.deleted_at is null
     where keep.id <> lose.id
  loop
    -- Repoint every foreign key that references actors, whatever table it lives in.
    for r in
      select tc.table_name as t, kcu.column_name as c
        from information_schema.table_constraints tc
        join information_schema.key_column_usage kcu
          on kcu.constraint_name = tc.constraint_name and kcu.table_schema = tc.table_schema
        join information_schema.constraint_column_usage ccu
          on ccu.constraint_name = tc.constraint_name
       where tc.constraint_type = 'FOREIGN KEY'
         and ccu.table_name = 'actors'
         and tc.table_schema = 'public'
    loop
      -- Junction tables carry unique keys like (actor_id, role). When the keeper
      -- already holds the same pair, repointing collides — and the right answer
      -- there is to DROP the duplicate's row, not to keep both. Anything without
      -- such a constraint simply repoints.
      begin
        execute format('update public.%I set %I = $1 where %I = $2', r.t, r.c, r.c)
          using pair.keep_id, pair.lose_id;
      exception when unique_violation then
        execute format('delete from public.%I where %I = $1', r.t, r.c)
          using pair.lose_id;
        raise notice 'collision on %.% - dropped the duplicate rows instead', r.t, r.c;
      end;
    end loop;

    update public.actors set deleted_at = now() where id = pair.lose_id;
    merged := merged + 1;
  end loop;

  raise notice 'merged % duplicate actor pair(s)', merged;
end $$;

-- Point the alias at the survivor so a future import can never re-create the split.
insert into public.vendor_aliases (pattern, actor_id, note)
select 'crossroads', a.id, 'merged by 159'
  from public.actors a
 where a.deleted_at is null and a.display_name = 'Crossroads Caf' || chr(233)
on conflict (pattern) do update set actor_id = excluded.actor_id;

insert into public.vendor_aliases (pattern, actor_id, note)
select 'cross x roads', a.id, 'merged by 159'
  from public.actors a
 where a.deleted_at is null and a.display_name = 'Crossroads Caf' || chr(233)
on conflict (pattern) do update set actor_id = excluded.actor_id;

commit;

-- DOWN: restore the soft-deleted actors (set deleted_at = null) and repoint by hand;
-- the original split is not reconstructible from this migration alone.
