-- =============================================================================
-- 054_agreement_event_link.sql  (additive)
-- agreements.event_id — optional link from a client agreement to an event, so a
-- signed agreement can auto-file into that event's Files tab (kind='contract').
-- Keeps the free-text event_date/venue_name for un-linked / ad-hoc agreements.
-- Admin-only RLS already covers agreements; no policy change needed.
-- =============================================================================
begin;
alter table public.agreements
  add column if not exists event_id uuid references public.events(id) on delete set null;
create index if not exists idx_agreements_event on public.agreements(event_id) where event_id is not null;

-- Allow the auto-filed HTML agreement snapshot into the private 'agreements'
-- bucket (previously PDF-only). The file-agreement Edge Function writes a
-- text/html snapshot that surfaces in the linked event's Files → Contract bucket.
update storage.buckets
  set allowed_mime_types = array['application/pdf','text/html']
  where id = 'agreements'
    and not ('text/html' = any(coalesce(allowed_mime_types, array[]::text[])));
commit;
-- DOWN: alter table public.agreements drop column event_id;
--       update storage.buckets set allowed_mime_types = array['application/pdf'] where id='agreements';
