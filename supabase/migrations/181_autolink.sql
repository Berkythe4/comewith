-- ============================================================
-- COME WITH — 181 auto-link, nightly, with a receipt
--
-- 180 finds what is not linked. This closes the ones that can be closed without
-- a judgement call, and writes down exactly what it did every single time.
--
-- THE RULE FOR WHAT autolink_data() IS ALLOWED TO DO: only a link that is
-- already implied by data somebody entered on purpose.
--   * an actor who is the payee on 2 expenses IS a vendor - the role is implied
--     by the relationship, not guessed from their name
--   * a donor name that matches an actor's display name EXACTLY is that actor
--   * a guest whose email matches an actor's email EXACTLY is that actor
--   * a vendor string that matches an actor name exactly, or a vendor_alias
--     pattern somebody wrote by hand, is that actor
--   * a row on an event inherits that event's ledger
-- FUZZY MATCHING IS NOT IN HERE AND SHOULD NOT BE ADDED. LEARNINGS §31 exists
-- because one bad merge collapsed unrelated payees into a single actor, and
-- store matching (§ the radio notes) had to grow three guards for the same
-- reason. Returning "not linked" beats linking the wrong person.
--
-- DRY BY DEFAULT. autolink_data() with no argument changes nothing and returns
-- what it WOULD do, the same convention scripts/renumber_shows.py uses. The
-- nightly job runs it for real; the dashboard button offers both.
--
-- EVERY RUN WRITES A ROW to data_health_runs with a per-action summary. That is
-- the point: a nightly process that silently mutates the actor graph and leaves
-- no receipt is not automation, it is drift.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. Audit sweep -> a run row
-- ---------------------------------------------------------------
create or replace function public.snapshot_data_health(p_source text default 'manual')
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_total int;
  v_sev   jsonb;
  v_by    jsonb;
  v_id    uuid;
begin
  select count(*) into v_total from public.v_data_health;

  select coalesce(jsonb_object_agg(severity, n), '{}'::jsonb) into v_sev
    from (select severity, count(*) as n from public.v_data_health group by severity) s;

  select coalesce(jsonb_object_agg(check_key, n), '{}'::jsonb) into v_by
    from (select check_key, count(*) as n from public.v_data_health group by check_key) c;

  insert into public.data_health_runs (kind, source, total, by_severity, summary)
  values ('audit', p_source, v_total, v_sev, v_by)
  returning id into v_id;

  return jsonb_build_object('run_id', v_id, 'total', v_total,
                            'by_severity', v_sev, 'by_check', v_by);
end $$;

revoke all on function public.snapshot_data_health(text) from anon;

-- ---------------------------------------------------------------
-- 2. The safe links
-- ---------------------------------------------------------------
create or replace function public.autolink_data(p_apply boolean default false,
                                                p_source text default 'manual')
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  r_roles int := 0; r_donor int := 0; r_guest int := 0;
  r_vend  int := 0; r_alias int := 0; r_ledger_x int := 0; r_ledger_i int := 0;
  v_sum jsonb; v_id uuid;
begin
  -- `on commit drop` only fires at COMMIT, so calling this twice inside one
  -- transaction - dry run then apply, which is exactly how the dashboard button
  -- works - collided on the second call. Drop first.
  drop table if exists _roles;
  drop table if exists _donor;
  drop table if exists _guest;
  drop table if exists _vend;
  drop table if exists _alias;

  -- ---- a. roles implied by relationships that already exist ----------------
  create temp table _roles on commit drop as
  select distinct a.id as actor_id, v.role
    from public.actors a
    cross join lateral (values
      ('vendor',        exists (select 1 from public.expenses x where x.vendor_actor_id = a.id and x.deleted_at is null)),
      ('customer',      exists (select 1 from public.income i where i.actor_id = a.id and i.deleted_at is null)),
      ('artist',        exists (select 1 from public.event_participants p where p.actor_id = a.id)),
      ('sponsor',       exists (select 1 from public.sponsorships s where s.actor_id = a.id)),
      ('venue_contact', exists (select 1 from public.venues n where n.actor_id = a.id and n.deleted_at is null)),
      ('donor',         exists (select 1 from public.third_party_donations d where d.actor_id = a.id))
    ) as v(role, hit)
   where a.deleted_at is null and v.hit
     and not exists (select 1 from public.actor_roles r
                      where r.actor_id = a.id and r.active);
  select count(*) into r_roles from _roles;

  -- ---- b. donations -> actor, exact display name ---------------------------
  create temp table _donor on commit drop as
  select d.id, a.id as actor_id
    from public.third_party_donations d
    join public.actors a on a.deleted_at is null
     and lower(btrim(a.display_name)) = lower(btrim(d.donor_name))
   where d.actor_id is null and coalesce(d.donor_name, '') <> '';
  select count(*) into r_donor from _donor;

  -- ---- c. guests -> actor, exact email -------------------------------------
  create temp table _guest on commit drop as
  select g.id, (select a.id from public.actors a
                 where a.deleted_at is null and lower(a.email) = lower(g.email)
                 order by a.created_at limit 1) as actor_id
    from public.guests g
   where g.deleted_at is null and g.actor_id is null and coalesce(g.email, '') <> ''
     and exists (select 1 from public.actors a
                  where a.deleted_at is null and lower(a.email) = lower(g.email));
  select count(*) into r_guest from _guest;

  -- ---- d. expense vendor string -> actor, exact name then hand-written alias
  create temp table _vend on commit drop as
  select x.id, a.id as actor_id, 'exact'::text as how
    from public.expenses x
    join public.actors a on a.deleted_at is null
     and lower(btrim(a.display_name)) = lower(btrim(x.vendor))
   where x.deleted_at is null and x.vendor_actor_id is null and not x.payee_na
     and coalesce(x.vendor, '') <> '';
  select count(*) into r_vend from _vend;

  create temp table _alias on commit drop as
  select x.id, va.actor_id, 'alias'::text as how
    from public.expenses x
    join public.vendor_aliases va
      on lower(btrim(x.vendor)) like lower(va.pattern) || '%'
   where x.deleted_at is null and x.vendor_actor_id is null and not x.payee_na
     and coalesce(x.vendor, '') <> ''
     and not exists (select 1 from _vend v where v.id = x.id);
  select count(*) into r_alias from _alias;

  -- ---- e. ledger follows the event it is filed against ---------------------
  select count(*) into r_ledger_x
    from public.expenses x join public.events e on e.id = x.event_id
   where x.deleted_at is null and e.deleted_at is null
     and x.ledger is distinct from (case when e.series ilike '%dance infusion%'
                                         then 'dance_infusion' else 'come_with' end);
  select count(*) into r_ledger_i
    from public.income i join public.events e on e.id = i.event_id
   where i.deleted_at is null and e.deleted_at is null
     and i.ledger is distinct from (case when e.series ilike '%dance infusion%'
                                         then 'dance_infusion' else 'come_with' end);

  if p_apply then
    insert into public.actor_roles (actor_id, role, active)
    select actor_id, role, true from _roles
    on conflict do nothing;

    update public.third_party_donations d set actor_id = s.actor_id
      from _donor s where d.id = s.id;

    update public.guests g set actor_id = s.actor_id
      from _guest s where g.id = s.id and s.actor_id is not null;

    update public.expenses x set vendor_actor_id = s.actor_id
      from _vend s where x.id = s.id;
    update public.expenses x set vendor_actor_id = s.actor_id
      from _alias s where x.id = s.id;

    update public.expenses x
       set ledger = case when e.series ilike '%dance infusion%' then 'dance_infusion' else 'come_with' end
      from public.events e
     where e.id = x.event_id and x.deleted_at is null and e.deleted_at is null
       and x.ledger is distinct from (case when e.series ilike '%dance infusion%'
                                           then 'dance_infusion' else 'come_with' end);
    update public.income i
       set ledger = case when e.series ilike '%dance infusion%' then 'dance_infusion' else 'come_with' end
      from public.events e
     where e.id = i.event_id and i.deleted_at is null and e.deleted_at is null
       and i.ledger is distinct from (case when e.series ilike '%dance infusion%'
                                           then 'dance_infusion' else 'come_with' end);
  end if;

  v_sum := jsonb_build_object(
    'actor_roles_inferred', r_roles,
    'donations_linked',     r_donor,
    'guests_linked',        r_guest,
    'payees_linked_exact',  r_vend,
    'payees_linked_alias',  r_alias,
    'expense_ledger_fixed', r_ledger_x,
    'income_ledger_fixed',  r_ledger_i,
    'total',                r_roles + r_donor + r_guest + r_vend + r_alias + r_ledger_x + r_ledger_i);

  insert into public.data_health_runs (kind, source, total, summary, note)
  values (case when p_apply then 'autolink' else 'autolink_dry' end, p_source,
          (v_sum->>'total')::int, v_sum,
          case when p_apply then 'applied' else 'dry run - nothing was changed' end)
  returning id into v_id;

  return v_sum || jsonb_build_object('run_id', v_id, 'applied', p_apply);
end $$;

revoke all on function public.autolink_data(boolean, text) from anon;

-- ---------------------------------------------------------------
-- 3. Nightly, after the KPI snapshot at 06:30
-- ---------------------------------------------------------------
-- Auto-link first, then audit, so the morning number reflects a system that has
-- already tidied what it safely can. Both leave a row in data_health_runs.
select cron.unschedule('data-health-nightly')
 where exists (select 1 from cron.job where jobname = 'data-health-nightly');

select cron.schedule('data-health-nightly', '0 7 * * *', $cron$
  select public.autolink_data(true, 'cron');
  select public.snapshot_data_health('cron');
$cron$);

-- Seed the history with today, so the dashboard has something to show and the
-- first nightly run has a baseline to be compared against.
select public.autolink_data(false, 'migration-181');
select public.snapshot_data_health('migration-181');

commit;

-- DOWN: select cron.unschedule('data-health-nightly');
--   drop function public.autolink_data(boolean, text);
--   drop function public.snapshot_data_health(text);
