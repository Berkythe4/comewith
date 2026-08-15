---
name: project_show_vs_episode_numbering
description: "Radio numbering — station_no is the global SHOW counter (SHOW n), edition_seq is the per-series episode number (Elements Ep1-4); renumbered to broadcast order 2026-08-04"
metadata: 
  node_type: memory
  type: project
  originSessionId: 9caef454-ca9a-4416-aa06-a6721798c2fc
  modified: 2026-08-04T16:38:36.668Z
---

Two numbers, two names, settled 2026-08-04. Never render either as a generic
"EP n" again — that collision is what prompted this.

- **`sc_playlists.station_no` = the SHOW counter.** Every broadcast ever, across
  all series. Displayed **`SHOW n`** in the dashboard, radio.html, dj.html,
  index.html, the auto social-post title, and the DB strings (migration 137).
- **`edition_seq` = the episode number inside a series** (Elements Ep1–4). This
  is what an audience knows, so the **rendered video keeps "EP n" and draws
  `edition_seq`** — an Elements video says EP 1, not SHOW 3.

**Renumbered to broadcast order** (Keith signed off): `station_no` is assigned at
CREATION, not airtime, and the Elements editions were planned after NYC Ep3 but
drop two weeks earlier. Now 1–2 = NYC Ep1–2, 3–6 = Elements Ep1–4, 7 = NYC Ep3.
Elements Ep4 is **SHOW 6** — 2 NYC + 4 Elements, the count Keith works from.

`scripts/renumber_shows.py` is the tool (dry by default, `--apply` to write). It
refuses to move a published episode — that number is in the slug, the public page
and the played-song history — parks numbers high first because `station_no` is
uniquely indexed, and remaps `sc_song_log.played_station_no` / `passed_station_no`
and `sc_playlist_tracks.carried_from`, which store the NUMBER rather than a
foreign key and would otherwise point at the wrong show. It prints the inverse
mapping to undo.

**When replacing a security-definer function, re-check its ACL after.** 137
replaced `radio_publish_station`; `create or replace` preserved the revokes, but
verify — a reset ACL there would expose the publish path over REST. See
[[feedback_no_broad_anon_grants]] and [[project_radio_release_pipeline]].
