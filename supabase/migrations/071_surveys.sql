-- =============================================================================
-- 071_surveys.sql
-- Post-event survey system. Responses tag to event / actor / customer(guest) /
-- subscriber via per-recipient tokenized invites; a public token allows anonymous
-- responses tagged to the event.
--
-- Security: all tables are ADMIN-ONLY via RLS (is_admin()). The public survey form
-- never touches these tables directly — it goes through the survey-get / survey-submit
-- edge functions (service role), exactly like agreement signing / artist-self.
-- =============================================================================
begin;

create table if not exists public.surveys (
  id              uuid primary key default gen_random_uuid(),
  title           text not null,
  intro           text,
  event_id        uuid references public.events(id) on delete set null,
  status          text not null default 'draft' check (status in ('draft','open','closed')),
  allow_anonymous boolean not null default true,
  public_token    text unique not null default encode(gen_random_bytes(18),'hex'),
  created_at      timestamptz not null default now(),
  created_by      uuid,
  updated_at      timestamptz not null default now()
);

create table if not exists public.survey_questions (
  id          uuid primary key default gen_random_uuid(),
  survey_id   uuid not null references public.surveys(id) on delete cascade,
  sort_order  int not null default 0,
  prompt      text not null,
  qtype       text not null check (qtype in ('rating','nps','choice','yesno','short_text','long_text')),
  options     jsonb not null default '[]'::jsonb,
  required    boolean not null default false
);
create index if not exists survey_questions_survey on public.survey_questions(survey_id, sort_order);

create table if not exists public.survey_invites (
  id            uuid primary key default gen_random_uuid(),
  survey_id     uuid not null references public.surveys(id) on delete cascade,
  token         text unique not null default encode(gen_random_bytes(18),'hex'),
  event_id      uuid references public.events(id) on delete set null,
  actor_id      uuid references public.actors(id) on delete set null,
  guest_id      uuid references public.guests(id) on delete set null,
  subscriber_id uuid references public.subscribers(id) on delete set null,
  email         text,
  label         text,
  sent_at       timestamptz,
  responded_at  timestamptz,
  created_at    timestamptz not null default now()
);
create index if not exists survey_invites_survey on public.survey_invites(survey_id);

create table if not exists public.survey_responses (
  id            uuid primary key default gen_random_uuid(),
  survey_id     uuid not null references public.surveys(id) on delete cascade,
  invite_id     uuid references public.survey_invites(id) on delete set null,
  event_id      uuid references public.events(id) on delete set null,
  actor_id      uuid references public.actors(id) on delete set null,
  guest_id      uuid references public.guests(id) on delete set null,
  subscriber_id uuid references public.subscribers(id) on delete set null,
  anonymous     boolean not null default false,
  source        text,
  submitted_at  timestamptz not null default now()
);
create index if not exists survey_responses_survey on public.survey_responses(survey_id);
create index if not exists survey_responses_event  on public.survey_responses(event_id);
create index if not exists survey_responses_actor  on public.survey_responses(actor_id);

create table if not exists public.survey_answers (
  id          uuid primary key default gen_random_uuid(),
  response_id uuid not null references public.survey_responses(id) on delete cascade,
  question_id uuid not null references public.survey_questions(id) on delete cascade,
  value       jsonb
);
create index if not exists survey_answers_response on public.survey_answers(response_id);

-- Admin-only RLS on every table (public form uses edge functions / service role).
alter table public.surveys           enable row level security;
alter table public.survey_questions  enable row level security;
alter table public.survey_invites    enable row level security;
alter table public.survey_responses  enable row level security;
alter table public.survey_answers    enable row level security;

do $$
declare t text;
begin
  foreach t in array array['surveys','survey_questions','survey_invites','survey_responses','survey_answers']
  loop
    execute format('drop policy if exists %I on public.%I', 'admin_'||t, t);
    execute format('create policy %I on public.%I for all using (public.is_admin()) with check (public.is_admin())', 'admin_'||t, t);
  end loop;
end $$;

-- Nav: Surveys in the Audience group (after Campaigns, before Social Calendar).
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
select 'surveys', 'Surveys', 'Audience', 155, true, false, false,
       coalesce((select default_roles from public.module_registry where key = 'campaigns'), '{}')
on conflict (key) do update set
  label = excluded.label, nav_group = excluded.nav_group,
  sort_order = excluded.sort_order, built = excluded.built;

notify pgrst, 'reload schema';
commit;
