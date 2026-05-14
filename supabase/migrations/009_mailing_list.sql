-- =============================================================================
-- 009_mailing_list.sql
-- Self-hosted mailing list (per decision #8). Resend Audiences is the delivery
-- mechanism; the source of truth lives here. CAN-SPAM compliant.
-- =============================================================================

create table public.subscribers (
  id              uuid primary key default gen_random_uuid(),
  email           text not null,
  full_name       text,
  status          text not null default 'pending'
                    check (status in ('pending', 'subscribed', 'unsubscribed', 'bounced', 'complained')),
  source          text,  -- 'comewith.org/subscribe', 'event:dance-infusion-2', etc.
  confirmed_at    timestamptz,
  unsubscribed_at timestamptz,
  unsubscribe_token text not null default encode(gen_random_bytes(24), 'hex'),
  resend_contact_id text,  -- ID in Resend Audiences after sync
  guest_id        uuid references public.guests(id) on delete set null,
  ip_at_signup    inet,
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now()
);

create unique index idx_subscribers_email_active on public.subscribers(lower(email))
  where status in ('pending', 'subscribed');
create index idx_subscribers_status on public.subscribers(status);
create index idx_subscribers_unsubscribe_token on public.subscribers(unsubscribe_token);

create trigger set_updated_at
  before update on public.subscribers
  for each row execute function public.handle_updated_at();

alter table public.subscribers enable row level security;

-- Public can INSERT (sign up). Edge function handles the actual confirm flow.
create policy "Anyone can subscribe"
  on public.subscribers for insert
  with check (true);

-- Public can SELECT/UPDATE only via their unsubscribe token (handled by Edge Function).
-- The Edge Function uses the service_role key, bypassing RLS.

create policy "Admins can manage subscribers"
  on public.subscribers for all
  using (public.is_admin());

-- =============================================================================
-- Segments — many-to-many tagging
-- =============================================================================
create table public.subscriber_segments (
  id              uuid primary key default gen_random_uuid(),
  subscriber_id   uuid not null references public.subscribers(id) on delete cascade,
  segment         text not null,
  added_at        timestamptz not null default now()
);

create unique index idx_subscriber_segments_unique on public.subscriber_segments(subscriber_id, segment);
create index idx_subscriber_segments_segment on public.subscriber_segments(segment);

alter table public.subscriber_segments enable row level security;

create policy "Admins can manage subscriber segments"
  on public.subscriber_segments for all
  using (public.is_admin());

-- =============================================================================
-- Mailing campaigns — sends sent via Resend
-- =============================================================================
create table public.mailing_campaigns (
  id              uuid primary key default gen_random_uuid(),
  name            text not null,
  subject         text not null,
  preview_text    text,
  body_html       text,
  body_text       text,
  segment_filter  text,
  sent_at         timestamptz,
  scheduled_for   timestamptz,
  resend_broadcast_id text,
  status          text not null default 'draft'
                    check (status in ('draft', 'scheduled', 'sending', 'sent', 'failed')),
  recipient_count integer not null default 0,
  created_by      uuid references public.profiles(id),
  created_at      timestamptz not null default now(),
  updated_at      timestamptz not null default now()
);

create index idx_mailing_campaigns_status on public.mailing_campaigns(status);
create index idx_mailing_campaigns_sent_at on public.mailing_campaigns(sent_at desc);

create trigger set_updated_at
  before update on public.mailing_campaigns
  for each row execute function public.handle_updated_at();

alter table public.mailing_campaigns enable row level security;

create policy "Admins can manage mailing campaigns"
  on public.mailing_campaigns for all
  using (public.is_admin());

-- =============================================================================
-- Mailing events — delivery, open, click, bounce, complaint from Resend webhooks
-- =============================================================================
create table public.mailing_events (
  id              uuid primary key default gen_random_uuid(),
  campaign_id     uuid references public.mailing_campaigns(id) on delete cascade,
  subscriber_id   uuid references public.subscribers(id) on delete cascade,
  event_type      text not null,
  resend_event_id text,
  metadata        jsonb not null default '{}'::jsonb,
  occurred_at     timestamptz not null default now()
);

create index idx_mailing_events_campaign_id on public.mailing_events(campaign_id);
create index idx_mailing_events_subscriber_id on public.mailing_events(subscriber_id);
create index idx_mailing_events_type on public.mailing_events(event_type);
create index idx_mailing_events_occurred_at on public.mailing_events(occurred_at desc);

alter table public.mailing_events enable row level security;

create policy "Admins can read mailing events"
  on public.mailing_events for select
  using (public.is_admin());
