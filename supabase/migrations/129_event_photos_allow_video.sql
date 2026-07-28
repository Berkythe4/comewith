-- 129: the event-photos bucket also holds social CLIPS (video) + audio, uploaded
-- via openAssetModal — but allowed_mime_types was images-only, so every video
-- upload was rejected by the bucket (and, via a UI bug, hung on "Saving…").
-- Broaden the allowlist to image + video + audio, and raise the size ceiling to
-- 200MB so short social clips fit. Still restricted (no arbitrary file types).
update storage.buckets
  set file_size_limit = 209715200,   -- 200 MB
      allowed_mime_types = array[
        'image/jpeg','image/png','image/webp','image/svg+xml','image/gif','image/heic','image/heif',
        'video/mp4','video/quicktime','video/webm','video/x-msvideo','video/x-matroska','video/mpeg','video/3gpp',
        'audio/mpeg','audio/mp4','audio/aac','audio/wav','audio/x-wav','audio/ogg','audio/flac','audio/x-m4a'
      ]
  where id = 'event-photos';
