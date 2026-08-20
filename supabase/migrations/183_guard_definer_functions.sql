-- ============================================================
-- COME WITH — 183 guard the security-definer functions
--
-- Found by auditing every SECURITY DEFINER function in public against its ACL.
-- A definer function runs as its OWNER, and Postgres grants EXECUTE to PUBLIC
-- unless told not to. On this project `authenticated` includes every radio
-- listener account (customer role, created by anyone who signs up), so three
-- functions were reachable by people who should never have been near them:
--
--   autolink_data          WRITES actor_roles, guests, expenses, income.
--                          A listener could have rewritten the actor graph.
--                          Added in 181, one session old. Mine.
--   snapshot_data_health   writes data_health_runs.
--   snapshot_kpis          writes metric_snapshots - KPI history, which the
--                          strategy board reads and which has no other source.
--                          Pre-existing.
--
-- Every other definer function in public already guards itself internally
-- (actor_set_task_status, generate_day_of_tasks, get_team_members,
-- radio_schedule_go_live) or is granted only to postgres + service_role
-- (gear_watch_kick, the radio_publish_* family). These three were the gap.
--
-- THE GUARD ALLOWS A NULL auth.uid(), deliberately. pg_cron and the service role
-- have no JWT, and the nightly data-health job runs through exactly that path.
-- It is the same exemption 140's protect_site_owner() makes, and for the same
-- reason: this protects the APP, not the project.
--
-- The bodies below are lifted verbatim from 181 with only the guard inserted, so
-- the two migrations cannot drift.
-- ============================================================
begin;

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
  -- 183: SECURITY DEFINER functions are executable by PUBLIC unless somebody
  -- says otherwise, and `authenticated` on this project includes every radio
  -- listener account. Without this, any signed-in listener could rewrite the
  -- actor graph. auth.uid() is null for pg_cron and the service role, which is
  -- how the nightly job still runs - the same exemption 140 makes for
  -- break-glass repair.
  if auth.uid() is not null and not public.is_admin() then
    raise exception 'snapshot_data_health is admin only' using errcode = '42501';
  end if;

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
  -- 183: SECURITY DEFINER functions are executable by PUBLIC unless somebody
  -- says otherwise, and `authenticated` on this project includes every radio
  -- listener account. Without this, any signed-in listener could rewrite the
  -- actor graph. auth.uid() is null for pg_cron and the service role, which is
  -- how the nightly job still runs - the same exemption 140 makes for
  -- break-glass repair.
  if auth.uid() is not null and not public.is_master_admin() then
    raise exception 'autolink_data is master_admin only' using errcode = '42501';
  end if;

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
-- snapshot_kpis: same class, pre-existing
-- ---------------------------------------------------------------
-- Fixed by REVOKE rather than by a guard, because nothing calls it but the 06:30
-- cron job - not the dashboard, not an edge function - so nobody legitimate loses
-- anything. Rewriting a working function to add a guard it does not need would be
-- the larger change, and an attempt to splice one in programmatically refused
-- rather than half-apply, which is the behaviour you want from that kind of edit.
-- postgres and service_role keep EXECUTE, which is the path pg_cron uses.
revoke all on function public.snapshot_kpis() from public, anon, authenticated;

commit;

-- DOWN: restore 181's unguarded definitions (not recommended) and the original
--   snapshot_kpis from its own migration.
