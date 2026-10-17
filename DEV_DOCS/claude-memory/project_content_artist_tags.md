---
name: project_content_artist_tags
description: Migration 221 (2026-10-17) — many artists per piece of event content via artist_ids arrays; event hub Content tab + Content Center + post/asset modals
metadata:
  node_type: memory
  type: project
  originSessionId: 53d8dce8-b2f8-4fe1-9584-dcd4fd25f289
  modified: 2026-10-17T18:49:13.166Z
---

Shipped + deployed 2026-10-17 (eee3141, dbbb305). `artist_ids` (uuid[]) on content_assets / social_posts / event_photos, and a json `artist_ids` array on each `events.recap_videos` item. Legacy single `artist_id` is kept, mirrored to `artist_ids[0]`. `v_artist_content` expands both and now skips staged recap items. Shared picker helpers: `ensureArtistPick()` / `artistTagsHTML()` / `artistPickDD()` in dashboard.html.

**Why:** Keith wanted multiple artists tagged on all event content (b2b sets, crowd shots, lineup recaps).

**How to apply:** any form that writes `artist_ids` must have read them first. The post composer hides the Artists box if it can't load them (LEARNINGS §94). Tags on photos/posts/assets are internal-only. Parked: a Content Center artist filter, and tagged photos/clips on artist.html. Related: [[project_artist_profiles]], [[project_social_calendar_v2_connector]].
