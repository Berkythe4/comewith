-- =============================================================================
-- 012_storage.sql
-- Storage bucket creation + RLS policies for Supabase Storage.
-- Run AFTER all table migrations.
-- =============================================================================

-- Create buckets. The id and name must match exactly.
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values
  ('agreements',       'agreements',       false, 10485760, array['application/pdf']),
  ('event-photos',     'event-photos',     true,  5242880,  array['image/jpeg', 'image/png', 'image/webp']),
  ('artist-photos',    'artist-photos',    true,  5242880,  array['image/jpeg', 'image/png', 'image/webp']),
  ('equipment-photos', 'equipment-photos', true,  5242880,  array['image/jpeg', 'image/png', 'image/webp']),
  ('receipts',         'receipts',         false, 5242880,  array['image/jpeg', 'image/png', 'image/webp', 'application/pdf']),
  ('sponsor-logos',    'sponsor-logos',    true,  2097152,  array['image/png', 'image/svg+xml', 'image/jpeg'])
on conflict (id) do nothing;

-- =============================================================================
-- Storage RLS policies
-- =============================================================================

-- AGREEMENTS bucket — private. Admins can upload/read, customers can read their own.
create policy "Admins can upload agreements"
  on storage.objects for insert
  with check (bucket_id = 'agreements' and public.is_admin());

create policy "Admins can read agreements"
  on storage.objects for select
  using (bucket_id = 'agreements' and public.is_admin());

create policy "Admins can update agreements"
  on storage.objects for update
  using (bucket_id = 'agreements' and public.is_admin());

create policy "Admins can delete agreements"
  on storage.objects for delete
  using (bucket_id = 'agreements' and public.is_master_admin());

-- EVENT PHOTOS bucket — public read, admin write.
create policy "Public can read event photos"
  on storage.objects for select
  using (bucket_id = 'event-photos');

create policy "Admins can upload event photos"
  on storage.objects for insert
  with check (bucket_id = 'event-photos' and public.is_admin());

create policy "Admins can manage event photos"
  on storage.objects for update
  using (bucket_id = 'event-photos' and public.is_admin());

create policy "Admins can delete event photos"
  on storage.objects for delete
  using (bucket_id = 'event-photos' and public.is_admin());

-- ARTIST PHOTOS bucket — public read, admin write.
create policy "Public can read artist photos"
  on storage.objects for select
  using (bucket_id = 'artist-photos');

create policy "Admins can manage artist photos"
  on storage.objects for all
  using (bucket_id = 'artist-photos' and public.is_admin());

-- EQUIPMENT PHOTOS bucket — public read, admin write.
create policy "Public can read equipment photos"
  on storage.objects for select
  using (bucket_id = 'equipment-photos');

create policy "Admins can manage equipment photos"
  on storage.objects for all
  using (bucket_id = 'equipment-photos' and public.is_admin());

-- RECEIPTS bucket — private, admin only.
create policy "Admins can manage receipts"
  on storage.objects for all
  using (bucket_id = 'receipts' and public.is_admin());

-- SPONSOR LOGOS bucket — public read, admin write.
create policy "Public can read sponsor logos"
  on storage.objects for select
  using (bucket_id = 'sponsor-logos');

create policy "Admins can manage sponsor logos"
  on storage.objects for all
  using (bucket_id = 'sponsor-logos' and public.is_admin());
