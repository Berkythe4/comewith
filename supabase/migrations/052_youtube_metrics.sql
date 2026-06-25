-- =============================================================================
-- 052_youtube_metrics.sql  (additive)
-- Richer YouTube metrics from the Data API: per-video table + new KPI cards.
-- Values are written by the pull-youtube-stats Edge Function (metric_snapshots /
-- youtube_videos). Admin-only RLS; public YouTube data (not financial).
-- =============================================================================
begin;

create table if not exists public.youtube_videos (
  video_id      text primary key,
  title         text,
  published_at  timestamptz,
  views         bigint not null default 0,
  likes         bigint not null default 0,
  comments      bigint not null default 0,
  thumbnail_url text,
  fetched_at    timestamptz not null default now()
);
alter table public.youtube_videos enable row level security;
create policy "Admins manage youtube videos" on public.youtube_videos for all using (public.is_admin());

-- New KPI cards (content workstream). Values flow from metric_snapshots via the function.
insert into public.kpi_targets (metric_key, workstream, label, target_value, comparison, unit, effective_date, active) values
  ('youtube.total_views',       'content', 'Total views',       50000, 'gte', '',     current_date, true),
  ('youtube.videos',            'content', 'Videos published',     50, 'gte', '',     current_date, true),
  ('youtube.total_likes',       'content', 'Total likes',        1000, 'gte', '',     current_date, true),
  ('youtube.engagement_rate',   'content', 'Engagement rate',       5, 'gte', '%',    current_date, true),
  ('youtube.days_since_upload', 'content', 'Days since upload',     14, 'lte', 'days', current_date, true)
on conflict do nothing;

commit;

-- DOWN: drop table youtube_videos; delete the 5 new kpi_targets rows.
