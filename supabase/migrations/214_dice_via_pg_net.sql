-- 214: route DICE calls through the database (pg_net).
--
-- From 2026-09-30 DICE (Cloudflare in front of api.dice.fm) answers 403 to every
-- request from the Supabase EDGE runtime, even with the exact headers dice.fm's own
-- web client sends. The same request made from the DATABASE via pg_net gets 200.
-- So pull-dice keeps all its logic but performs its HTTP through these two helpers:
--   dice_fetch_enqueue(reqs)  queue GET/POST calls, return pg_net request ids
--   dice_fetch_collect(ids)   read back whichever of those responses have landed
-- pg_net is asynchronous and its worker only sees committed rows, so enqueue and
-- collect are separate RPCs; the edge function polls collect.
--
-- These make outbound HTTP as the database, so they are fenced hard:
--   * host is fixed to https://api.dice.fm; only /unified_search (POST) and
--     /events/<id> (GET) are accepted; anything else raises.
--   * EXECUTE is service_role only. Revoked from PUBLIC as well as anon and
--     authenticated, because a new function's EXECUTE goes to PUBLIC (LEARNINGS §45).

begin;

create or replace function public.dice_fetch_enqueue(reqs jsonb)
returns bigint[]
language plpgsql
security definer
set search_path = public, net
as $$
declare
  r jsonb;
  ids bigint[] := '{}';
  p text;
  h jsonb := jsonb_build_object(
    'User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36',
    'Accept', 'application/json', 'Content-Type', 'application/json',
    'Origin', 'https://dice.fm', 'Referer', 'https://dice.fm/',
    'X-Api-Timestamp', '2025-04-16', 'X-Client-Platform', 'web',
    'X-Client-Timezone', 'America/New_York',
    'X-Device-Id', gen_random_uuid()::text);
begin
  if jsonb_typeof(reqs) <> 'array' or jsonb_array_length(reqs) > 200 then
    raise exception 'dice_fetch_enqueue: expected an array of at most 200 requests';
  end if;
  for r in select * from jsonb_array_elements(reqs) loop
    p := r->>'path';
    if p = '/unified_search' then
      ids := ids || net.http_post(
        url := 'https://api.dice.fm/unified_search',
        body := coalesce(r->'body', '{}'::jsonb),
        headers := h, timeout_milliseconds := 15000);
    elsif p ~ '^/events/[A-Za-z0-9]{1,40}$' then
      ids := ids || net.http_get(
        url := 'https://api.dice.fm' || p,
        headers := h, timeout_milliseconds := 15000);
    else
      raise exception 'dice_fetch_enqueue: path not allowed: %', p;
    end if;
  end loop;
  return ids;
end $$;

create or replace function public.dice_fetch_collect(ids bigint[])
returns table (id bigint, status_code int, content text, timed_out boolean, error_msg text)
language sql
security definer
set search_path = public, net
as $$
  select r.id, r.status_code, r.content, r.timed_out, r.error_msg
  from net._http_response r
  where r.id = any(ids);
$$;

revoke all on function public.dice_fetch_enqueue(jsonb) from public, anon, authenticated;
revoke all on function public.dice_fetch_collect(bigint[]) from public, anon, authenticated;
grant execute on function public.dice_fetch_enqueue(jsonb) to service_role;
grant execute on function public.dice_fetch_collect(bigint[]) to service_role;

commit;
