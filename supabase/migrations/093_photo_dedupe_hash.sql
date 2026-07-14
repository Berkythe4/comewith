-- =============================================================================
-- 093_photo_dedupe_hash.sql
-- Duplicate-photo guard. Storage can't dedupe content (and our paths are
-- timestamped), and filenames lie (IMG_8024 (1).jpg / two cameras both naming
-- IMG_0001). The dashboard hashes the ORIGINAL file bytes (SHA-256 via Web
-- Crypto) before upload and skips any hash already on the event. Nullable:
-- photos uploaded before this migration have no hash and are never matched.
-- =============================================================================
begin;

alter table public.event_photos
  add column if not exists file_hash text;

comment on column public.event_photos.file_hash is
  'SHA-256 hex of the ORIGINAL uploaded file (pre-downscale), computed in the browser. Upload UI skips files whose hash already exists on the event. Null = uploaded before dedupe existed.';

create index if not exists event_photos_hash_idx
  on public.event_photos (event_id, file_hash);

commit;
