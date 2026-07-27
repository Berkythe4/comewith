-- 124: raise event-photos bucket limit 15MB -> 50MB so phone video CLIPS
-- (content_assets uploads under clips/…) fit. Photos are resized client-side to
-- well under this; only short social videos need the headroom.
update storage.buckets set file_size_limit = 52428800 where id = 'event-photos';
