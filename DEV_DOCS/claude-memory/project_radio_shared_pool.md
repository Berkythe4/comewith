---
name: project-radio-shared-pool
description: Migration 136 — track attribution (added_by) + per-person claimed/maybe/veto marks so several people can build one station from a shared song pool
metadata: 
  node_type: memory
  type: project
  originSessionId: 279f1814-6e4e-4f39-a354-6447da747ad9
  modified: 2026-07-29T23:51:16.506Z
---

Several people now build a Come With Radio station from the same pool of songs
(decided 2026-07-30). Migration **136** applied to prod:

- `sc_playlist_tracks.added_by` → `profiles(id)`, **default `auth.uid()`**. A default,
  not a trigger, because the service-role callers (sc-connect sync + carry-over,
  dj-station) have no user and NULL is the honest answer. Rows predating 136 are NULL
  and render as "—".
- `sc_track_marks (track_id, user_id, mark)`, `mark in ('claimed','maybe','veto')`,
  unique per (track, person). **A child table, not a column** — the whole point is to
  surface disagreement, so two people must hold different positions on the same track
  at once; one column lets the last writer silently overwrite the other's call.
- RLS: any admin **reads all** marks (seeing the disagreement is the feature), writes
  **only their own** row. Verified on prod as Martin: own insert OK, writing as Keith
  refused with 42501.
- **Episode-scoped on purpose.** A song's role in a set is per-set, so marks cascade
  with the track and `sc_song_log` stays purely played/passed/carried. Keith explicitly
  chose per-episode over "follows the song forever".

UI lives in the ⛶ Arrange view (see [[project_beatport_cart_api]] era work): columns
"Added by" (initials, stable hue per person) and "Calls" (one chip per teammate);
clicking cycles your own mark none → claimed → maybe → veto → clear.

NOT yet surfaced: the main station tracklist and `dj.html` (the DJ scoped link) don't
show marks — only the Arrange table does. Related: [[project_radio_episode_planning]],
[[project_radio_release_pipeline]].
