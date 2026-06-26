-- 066_e2e_review.sql
-- Shared state for the end-to-end testing checklist. One row per (test, reviewer);
-- each reviewer (martin / henry / keith) has their own checked + notes so the
-- columns are independent. Admin-only (Martin + Henry are sub_admin).

create table public.e2e_review (
  test_key   text not null,
  reviewer   text not null,
  checked    boolean not null default false,
  notes      text,
  updated_at timestamptz not null default now(),
  primary key (test_key, reviewer)
);
alter table public.e2e_review enable row level security;
create policy "Admins manage e2e review" on public.e2e_review
  for all using (public.is_admin()) with check (public.is_admin());
-- Table privileges inherited from 013 default privileges (no blanket grants).

notify pgrst, 'reload schema';
