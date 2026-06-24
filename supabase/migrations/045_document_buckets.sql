-- =============================================================================
-- 045_document_buckets.sql  (additive, reversible)
-- Files-tab redesign: per-document-type "buckets" + a vendor link on files.
--
--  A) files.vendor_actor_id — the counterparty/vendor a file relates to (distinct
--     from uploaded_by = who uploaded it). Lets the Files tab show "vendor".
--  B) document_types — admin-extensible registry of doc-type buckets (Contract,
--     Technical Rider, …). files.kind stores the slug; the UI renders one bucket
--     per registry row (in `sort` order) plus any ad-hoc kinds present on files.
--     "Add document type" inserts a row here, so new buckets persist for everyone.
--
-- ADDITIVE ONLY: 1 nullable column + 1 table + seed. No DROP / destructive change.
-- RLS: admin-only (is_admin()). New objects inherit 013 default privileges; NO anon
-- grant (CLAUDE.md 013/016/017/019 discipline). document_types is not financial.
-- =============================================================================
begin;

-- A) vendor/counterparty on files
alter table public.files
  add column if not exists vendor_actor_id uuid references public.actors(id) on delete set null;
create index if not exists idx_files_vendor on public.files(vendor_actor_id) where vendor_actor_id is not null;

-- B) doc-type bucket registry
create table if not exists public.document_types (
  id          uuid primary key default gen_random_uuid(),
  slug        text unique not null,
  label       text not null,
  sort        integer not null default 100,
  is_standard boolean not null default false,
  created_at  timestamptz not null default now()
);
alter table public.document_types enable row level security;
create policy "Admins manage document types"
  on public.document_types for all using (public.is_admin());

insert into public.document_types (slug, label, sort, is_standard) values
  ('contract',          'Contract',          10,  true),
  ('technical_rider',   'Technical Rider',   20,  true),
  ('hospitality_rider', 'Hospitality Rider', 30,  true),
  ('stage_plot',        'Stage Plot',        40,  true),
  ('insurance_coi',     'Insurance / COI',   50,  true),
  ('permit',            'Permit / License',  60,  true),
  ('invoice',           'Invoice',           70,  true),
  ('receipt',           'Receipt',           80,  true),
  ('floor_plan',        'Floor Plan',        90,  true),
  ('w9',                'W-9 / Tax',         100, true),
  ('photo',             'Photo',             110, true),
  ('other',             'Other',             999, true)
on conflict (slug) do nothing;

commit;

-- DOWN:
--   drop table public.document_types;
--   alter table public.files drop column vendor_actor_id;
