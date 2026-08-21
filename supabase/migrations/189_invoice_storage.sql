-- ============================================================
-- COME WITH — 189 the invoices bucket
--
-- PRIVATE, and it has to stay that way. An invoice PDF carries the client's
-- name, what they were charged, and — depending on `invoice_settings` — the
-- Bluevine routing and account numbers printed in the "how to pay" block. The
-- event-photos and artist-photos buckets are public because publishing is their
-- job; this one is the opposite, and LEARNINGS §32 already made the general
-- point that a public bucket is never a place for documents.
--
-- The client still gets their PDF without any anon grant on the bucket: the
-- public invoice page asks the `invoice-public` edge function, which holds the
-- service role, matches on `public_token`, and streams the bytes back. Same
-- shape as `get-station`.
--
-- MIME is restricted to PDF and HTML. Nothing else has any business here, and a
-- narrow allow-list means a bug that tries to upload something odd fails loudly
-- at the storage layer rather than quietly succeeding.
-- ============================================================
begin;

insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values ('invoices', 'invoices', false, 10485760, array['application/pdf', 'text/html'])
on conflict (id) do update
  set public = false,
      file_size_limit = excluded.file_size_limit,
      allowed_mime_types = excluded.allowed_mime_types;

drop policy if exists "Admins can upload invoices" on storage.objects;
create policy "Admins can upload invoices"
  on storage.objects for insert to authenticated
  with check (bucket_id = 'invoices' and public.is_admin());

drop policy if exists "Admins can read invoices" on storage.objects;
create policy "Admins can read invoices"
  on storage.objects for select to authenticated
  using (bucket_id = 'invoices' and public.is_admin());

drop policy if exists "Admins can update invoices" on storage.objects;
create policy "Admins can update invoices"
  on storage.objects for update to authenticated
  using (bucket_id = 'invoices' and public.is_admin())
  with check (bucket_id = 'invoices' and public.is_admin());

drop policy if exists "Admins can delete invoices" on storage.objects;
create policy "Admins can delete invoices"
  on storage.objects for delete to authenticated
  using (bucket_id = 'invoices' and public.is_admin());

-- 'invoice' joins the standard document buckets so a filed invoice lands in its
-- event's Files tab next to the contract, rather than in an "other" pile.
insert into public.document_types (slug, label, sort, is_standard)
values ('invoice', 'Invoice', 45, true)
on conflict (slug) do nothing;

commit;

-- DOWN:
--   delete from storage.buckets where id = 'invoices';   -- objects must go first
--   drop policy if exists "Admins can upload invoices" on storage.objects;  (etc.)
--   delete from public.document_types where slug = 'invoice';
