-- ============================================================
-- COME WITH — 201 plan_publish_round(): freeze a reading of the plan
--
-- The working plan is the LIVE plan — it is always the current best view and it
-- stays editable forever. Publishing does not lock it; publishing takes a
-- SNAPSHOT, so "how did we do against the forecast" can name which forecast.
-- Run it at whatever cadence suits — twice monthly, weekly during a push.
--
-- WHY THIS IS A FUNCTION AND NOT THREE CLIENT CALLS. The 197 freeze trigger
-- refuses any write to a row belonging to a published version — which is the
-- whole point of a snapshot, and which also means a client cannot insert the
-- copied rows into it. Doing it in steps from the dashboard would mean either
-- weakening the trigger or leaving a half-copied round behind when a call fails
-- midway. Here the version is created in a transient state, filled, and only
-- then marked published, all inside one transaction. A failure rolls the whole
-- thing back and leaves no orphan.
--
-- security definer, because the trigger must be satisfied by ordering rather
-- than by privilege — but the caller is still checked: an admin JWT or the
-- service role, nothing else. Follows the 183 guard pattern.
-- ============================================================
begin;

create or replace function public.plan_publish_round(p_label text default null)
returns uuid
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_src   uuid;
  v_new   uuid;
  v_label text;
begin
  -- An anonymous or non-admin caller gets nothing. auth.uid() is null for the
  -- service role, which is allowed on purpose (break-glass / scheduled use).
  if auth.uid() is not null and not public.is_admin() then
    raise exception 'admin only' using errcode = '42501';
  end if;

  select id into v_src from public.plan_versions where status = 'working' limit 1;
  if v_src is null then
    raise exception 'there is no working plan to publish';
  end if;

  v_label := coalesce(nullif(btrim(p_label), ''),
                      'FC ' || to_char(now() at time zone 'utc', 'YYYY-MM-DD'));

  -- Transient status so the 197 freeze trigger lets the copy through. It is
  -- flipped to 'published' at the end of this same transaction, so no row is
  -- ever visible to another session in this state.
  insert into public.plan_versions (label, status, horizon_months, basis_period, notes, created_by)
  select v_label, 'archived', s.horizon_months, to_char(current_date, 'YYYY-MM'),
         'Snapshot of the working plan taken ' || to_char(now() at time zone 'utc', 'YYYY-MM-DD HH24:MI') || ' UTC.',
         auth.uid()
    from public.plan_versions s where s.id = v_src
  returning id into v_new;

  insert into public.plan_volumes (version_id, offering_id, period, units, scale, notes)
  select v_new, offering_id, period, units, scale, notes
    from public.plan_volumes where version_id = v_src;

  insert into public.plan_overrides (version_id, period, ledger, section, category, amount, reason, created_by)
  select v_new, period, ledger, section, category, amount, reason, created_by
    from public.plan_overrides where version_id = v_src;

  insert into public.budget_lines
    (scope, period, category, direction, planned_amount, ledger, version_id, label, notes, event_id, event_type)
  select scope, period, category, direction, planned_amount, ledger, v_new, label, notes, event_id, event_type
    from public.budget_lines where version_id = v_src;

  update public.plan_versions
     set status = 'published', published_at = now()
   where id = v_new;

  return v_new;
end;
$$;

comment on function public.plan_publish_round(text) is
  'Freezes a copy of the working plan as a published round and returns its id. '
  'The working plan itself is untouched and stays editable — publishing takes a '
  'reading, it does not close the books.';

revoke all on function public.plan_publish_round(text) from public, anon;
grant execute on function public.plan_publish_round(text) to authenticated, service_role;

commit;

-- DOWN: drop function if exists public.plan_publish_round(text);
