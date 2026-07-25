-- =============================================================================
-- 119_client_error_log.sql
-- Surface CLIENT-side errors in the Activity log. DB triggers only record
-- successful writes; an operation that errors (constraint, RLS, network) rolls
-- back and leaves no trace, so the only record was a toast the user saw once.
-- A security-definer RPC writes those into audit_log as action='ERROR', so they
-- appear in the same Activity log with no new surface to check.
-- =============================================================================
begin;

-- audit_log.action only allowed INSERT/UPDATE/DELETE.
alter table public.audit_log drop constraint if exists audit_log_action_check;
alter table public.audit_log add constraint audit_log_action_check
  check (action = any (array['INSERT', 'UPDATE', 'DELETE', 'ERROR']));

create or replace function public.log_client_event(p_message text, p_context text default null)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.audit_log (table_name, row_id, action, actor_id, actor_email, new_data)
  values (
    '(client)', '-', 'ERROR', auth.uid(),
    (select email from public.profiles where id = auth.uid()),
    jsonb_build_object('message', left(coalesce(p_message, ''), 2000),
                       'context', left(coalesce(p_context, ''), 500))
  );
end $$;
grant execute on function public.log_client_event(text, text) to authenticated;

commit;
