-- 097: Team chat — channels / members / messages + realtime
--
-- In-app chat between staff: one team-wide channel, DMs per pair (dm_key =
-- sorted "uid:uid"), optional per-event threads. Emails sent to a teammate
-- from the Users tab are mirrored into the DM as kind='email' rows so one
-- thread tracks both. Delivery is Supabase Realtime (publication below) with
-- client-side polling fallback. All access is staff-only (public.is_admin());
-- DMs are member-only — master does NOT get implicit read of others' DMs.
-- New tables inherit grants from 013 default privileges; anon explicitly
-- revoked below, and RLS has real policies (never enabled-with-no-policy).

create table if not exists public.chat_channels (
  id          uuid primary key default gen_random_uuid(),
  kind        text not null check (kind in ('team','dm','event')),
  name        text,
  event_id    uuid references public.events(id) on delete set null,
  dm_key      text unique,
  created_by  uuid references public.profiles(id) on delete set null,
  created_at  timestamptz not null default now(),
  deleted_at  timestamptz
);

create table if not exists public.chat_members (
  channel_id   uuid not null references public.chat_channels(id) on delete cascade,
  user_id      uuid not null references public.profiles(id) on delete cascade,
  last_read_at timestamptz not null default now(),
  primary key (channel_id, user_id)
);
create index if not exists idx_chat_members_user on public.chat_members (user_id);

create table if not exists public.chat_messages (
  id              uuid primary key default gen_random_uuid(),
  channel_id      uuid not null references public.chat_channels(id) on delete cascade,
  author_id       uuid references public.profiles(id) on delete set null,
  kind            text not null default 'chat' check (kind in ('chat','email','system')),
  body            text not null,
  email_subject   text,
  conversation_id uuid references public.conversations(id) on delete set null,
  created_at      timestamptz not null default now(),
  edited_at       timestamptz,
  deleted_at      timestamptz
);
create index if not exists idx_chat_messages_chan on public.chat_messages (channel_id, created_at);

-- ---- visibility helper -----------------------------------------------------
create or replace function public.can_see_channel(ch uuid)
returns boolean language sql stable security definer set search_path = public as $$
  select exists (
    select 1 from public.chat_channels c
    where c.id = ch and c.deleted_at is null
      and ( (c.kind in ('team','event') and public.is_admin())
            -- creator visibility is needed for INSERT..RETURNING on a fresh DM,
            -- before its member rows exist (creator is always a member anyway)
            or c.created_by = auth.uid()
            or exists (select 1 from public.chat_members m
                       where m.channel_id = ch and m.user_id = auth.uid()) )
  );
$$;

-- ---- RLS -------------------------------------------------------------------
alter table public.chat_channels enable row level security;
alter table public.chat_members  enable row level security;
alter table public.chat_messages enable row level security;

drop policy if exists chat_channels_select on public.chat_channels;
-- created_by is checked directly on the row (not via the helper): INSERT..RETURNING
-- enforces the SELECT policy mid-statement, where can_see_channel()'s re-query
-- cannot see the just-inserted row yet.
create policy chat_channels_select on public.chat_channels
  for select using (created_by = auth.uid() or public.can_see_channel(id));
drop policy if exists chat_channels_insert on public.chat_channels;
create policy chat_channels_insert on public.chat_channels
  for insert with check (public.is_admin() and created_by = auth.uid());
drop policy if exists chat_channels_update on public.chat_channels;
create policy chat_channels_update on public.chat_channels
  for update using (public.is_master_admin() or created_by = auth.uid());
drop policy if exists chat_channels_delete on public.chat_channels;
create policy chat_channels_delete on public.chat_channels
  for delete using (public.is_master_admin());

-- Members: the channel creator seeds the member rows (both sides of a DM);
-- anyone can update their own last_read_at.
drop policy if exists chat_members_select on public.chat_members;
create policy chat_members_select on public.chat_members
  for select using (public.can_see_channel(channel_id));
drop policy if exists chat_members_insert on public.chat_members;
create policy chat_members_insert on public.chat_members
  for insert with check (
    public.is_admin() and exists (
      select 1 from public.chat_channels c
      where c.id = channel_id
        and (c.created_by = auth.uid() or c.kind in ('team','event') or user_id = auth.uid())
    )
  );
drop policy if exists chat_members_update on public.chat_members;
create policy chat_members_update on public.chat_members
  for update using (user_id = auth.uid());
drop policy if exists chat_members_delete on public.chat_members;
create policy chat_members_delete on public.chat_members
  for delete using (public.is_master_admin() or user_id = auth.uid());

drop policy if exists chat_messages_select on public.chat_messages;
create policy chat_messages_select on public.chat_messages
  for select using (public.can_see_channel(channel_id));
drop policy if exists chat_messages_insert on public.chat_messages;
create policy chat_messages_insert on public.chat_messages
  for insert with check (
    public.is_admin() and author_id = auth.uid() and public.can_see_channel(channel_id)
  );
drop policy if exists chat_messages_update on public.chat_messages;
create policy chat_messages_update on public.chat_messages
  for update using (author_id = auth.uid());
drop policy if exists chat_messages_delete on public.chat_messages;
create policy chat_messages_delete on public.chat_messages
  for delete using (public.is_master_admin() or author_id = auth.uid());

-- ---- anon lockdown ---------------------------------------------------------
revoke all on public.chat_channels, public.chat_members, public.chat_messages from anon;

-- ---- seed the team-wide channel -------------------------------------------
insert into public.chat_channels (kind, name)
select 'team', 'Team'
where not exists (select 1 from public.chat_channels where kind = 'team' and deleted_at is null);

-- ---- realtime --------------------------------------------------------------
-- The supabase_realtime publication exists on Supabase projects but currently
-- has no tables; add chat_messages so inserts stream to subscribed clients
-- (postgres_changes enforces RLS per subscriber).
do $$
begin
  if not exists (select 1 from pg_publication where pubname = 'supabase_realtime') then
    create publication supabase_realtime;
  end if;
  if not exists (select 1 from pg_publication_tables
                 where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'chat_messages') then
    alter publication supabase_realtime add table public.chat_messages;
  end if;
end $$;
