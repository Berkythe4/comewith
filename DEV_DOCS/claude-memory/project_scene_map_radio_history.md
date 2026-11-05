---
name: project-scene-map-radio-history
description: "Scene Map rebuilt 2026-11-05 as radio history - tracklists primary (v_radio_track_shows), other-shows layer, Explore table drives map, master tracklist + SC export per view; migrations 229-239"
metadata:
  node_type: memory
  type: project
  originSessionId: 97bac3d7-052f-4724-9c05-ea483eb276b7
  modified: 2026-11-05T21:28:38.975Z
---

Scene Map (dashboard, internal) was rebuilt on 2026-11-05 into a radio-history tool; default mode "Radio history", old map kept as "Upcoming shows".

- Primary source = OUR tracklists: `v_radio_track_shows` (233-237) resolves each aired song's recorded show to a room + party. Listings lineup matching is a secondary "other shows" layer (dimmed, "picked for X", ignores `via:'title'` lineup entries).
- Any open club or Explore filter (radio show / artist / song / search) renders ONE master tracklist under the map (`sceneCtx()` → `renderSceneList()`), with "Copy to my SoundCloud" (231: `sc_dj_exports.track_ids/title`).
- Venues: rooms = one club (Keith), BASEMENT separate, Pacha variants → Pacha New York, Elements Festival = Pocono Raceway; 7 pins fixed from RA addresses.
- Rekordbox songs: `sc_link_id/url` via `scripts/link_radio_tracks_soundcloud.py` (116/152); never used by push/sync.
- Tiles: Esri (CARTO returns API-KEY images at 200).

**Why:** Keith wants to see which clubs our radio songs come from and export a club's (or artist's / show's) songs as a playlist; he found earlier versions hard to read.
**How to apply:** keep tracklists authoritative, attribute every listings-derived song, keep the state bar/exits obvious; not yet clicked through live or export-tested end to end (as of 2026-11-05). See [[project-ra-market-radio]] and [[feedback-ship-what-keith-asked-for]].
