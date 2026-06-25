-- =============================================================================
-- 056_conversations.sql
-- Email conversations: every email sent to an actor/venue opens a thread; all
-- messages (outbound, status events like bounces, inbound replies, manual notes)
-- are logged for the whole team — unless the thread is marked 'restricted', in
-- which case only master + creator + an explicit ACL can see it.
-- New tables inherit grants from 013 ALTER DEFAULT PRIVILEGES (no manual grants).
-- =============================================================================
begin;

create table if not exists public.conversations (
  id uuid primary key default gen_random_uuid(),
  subject text not null,
  actor_id uuid references public.actors(id) on delete set null,   -- the counterparty
  recipient_email text not null,
  source text,                       -- human label of where it was sent from
  source_kind text,                  -- 'actor' | 'venue' | 'event_people'
  source_id uuid,                    -- actor/venue/event id (for the deep link)
  event_id uuid references public.events(id) on delete set null,
  created_by uuid references public.profiles(id),
  visibility text not null default 'team' check (visibility in ('team','restricted')),
  last_message_at timestamptz not null default now(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  deleted_at timestamptz
);
create index if not exists idx_conversations_actor on public.conversations(actor_id);
create index if not exists idx_conversations_last on public.conversations(last_message_at desc);

create table if not exists public.conversation_messages (
  id uuid primary key default gen_random_uuid(),
  conversation_id uuid not null references public.conversations(id) on delete cascade,
  direction text not null check (direction in ('outbound','inbound','note','event')),
  from_email text,
  to_email text,
  subject_line text,
  body text,
  resend_id text,                    -- Resend message id, for webhook correlation
  status text,                       -- queued|sent|delivered|bounced|complained|opened
  created_by uuid references public.profiles(id),
  created_at timestamptz not null default now(),
  meta jsonb not null default '{}'
);
create index if not exists idx_convmsg_conv on public.conversation_messages(conversation_id, created_at);
create index if not exists idx_convmsg_resend on public.conversation_messages(resend_id) where resend_id is not null;

create table if not exists public.conversation_acl (
  conversation_id uuid not null references public.conversations(id) on delete cascade,
  user_id uuid not null references public.profiles(id) on delete cascade,
  primary key (conversation_id, user_id)
);

alter table public.conversations        enable row level security;
alter table public.conversation_messages enable row level security;
alter table public.conversation_acl      enable row level security;

-- Can the current user see this thread? master, creator, team-visible (with the
-- conversations module), or explicitly ACL'd.
create or replace function public.can_see_conversation(p_conv uuid)
returns boolean language sql stable security definer set search_path = public as $$
  select exists (
    select 1 from public.conversations c
    where c.id = p_conv and c.deleted_at is null and (
         public.is_master_admin()
      or c.created_by = auth.uid()
      or (c.visibility = 'team' and public.user_can_access_module('conversations'))
      or exists (select 1 from public.conversation_acl a where a.conversation_id = c.id and a.user_id = auth.uid())
    )
  );
$$;
grant execute on function public.can_see_conversation(uuid) to authenticated;

create policy "Conversations visible" on public.conversations for select
  using (public.can_see_conversation(id));
create policy "Conversations insert" on public.conversations for insert
  with check (public.user_can_access_module('conversations'));
create policy "Conversations update" on public.conversations for update
  using (public.is_master_admin() or created_by = auth.uid())
  with check (public.is_master_admin() or created_by = auth.uid());
create policy "Conversations delete" on public.conversations for delete
  using (public.is_master_admin() or created_by = auth.uid());

create policy "Messages visible" on public.conversation_messages for select
  using (public.can_see_conversation(conversation_id));
create policy "Messages insert" on public.conversation_messages for insert
  with check (public.can_see_conversation(conversation_id) and public.user_can_access_module('conversations'));

create policy "ACL manage" on public.conversation_acl for all
  using (public.is_master_admin() or exists (select 1 from public.conversations c where c.id = conversation_id and c.created_by = auth.uid()))
  with check (public.is_master_admin() or exists (select 1 from public.conversations c where c.id = conversation_id and c.created_by = auth.uid()));

-- Register the Conversations module (signed off so all staff roles see it).
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('conversations', 'Conversations', 'Audience',
        (select coalesce(max(sort_order),0)+1 from public.module_registry),
        true, true, false, array['operations','marketing','full'])
on conflict (key) do update set built = true, signed_off = true;

commit;
