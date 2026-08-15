---
name: project_radio_video_bookends
description: Radio YouTube render — intro/closing story slides + per-track genre & release date (awaiting Keith sign-off)
metadata: 
  node_type: memory
  type: project
  originSessionId: db51d261-6033-4aa4-a2da-cd68f16c298d
  modified: 2026-07-27T18:32:39.466Z
---

Radio episode YouTube render (`Radio/render/render_episode.py`, driven by
`make_episode.py`) got bookends + song facts, built 2026-07-27:

- **Intro & closing "story slides"** — one accumulating slide each, revealing a
  beat at a time (~1.7–2.4s) then holding ~5s. Intro: cover+ON AIR+wordmark →
  credits → purpose ("every track is an artist playing NYC soon") → the move →
  comewith.org. Closing: thanks → tracklist w/ cost downloadable at site →
  follow @comewith → "we plug back in next Thursday · <date>". Both are
  **OVERLAID on the opening/closing of the mix** so total runtime is unchanged
  (no dead air). Flags: `--mixed-by/--drop-date/--next-date/--no-bookends`.
  `make_episode` auto-pulls mix_by + this/next `drop_date` from prod.
- **Song cards** now show a genre + release-year line under the title (show
  date/venue chips stay). `genres` already stored; **release_date is new** —
  migration **125** added `sc_playlist_tracks.release_date` (nullable, applied
  to prod), and `track-sources` fills release_date+label from Beatport
  (`publish_date`/`new_release_date`), FILL-only. So release dates only appear
  after the "🛒 Where to buy" run; until then that half of the line omits itself.
- **Sign-off tool:** `Radio/render/preview_bookends.py` renders just intro + a
  sample card + closing → stills + silent mp4 in `Radio/Video/_preview/`,
  reusing the real render fns. Storyboard artifact was published for review.

STATUS: **awaiting Keith's sign-off** on intro/closing copy/timing/@handle
before it's considered final. Handle "@comewith" and the "next Thursday" cadence
are assumptions. See [[project_radio_release_pipeline]].
