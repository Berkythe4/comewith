---
name: project-radio-weekly-release-runbook
description: COMEWITHRADIO/NOTES_WEEKLY_RELEASE.md (was Radio/, renamed 2026-10-19) is the weekly episode runbook — read it before touching a release; key traps summarised here
metadata: 
  node_type: memory
  type: reference
  originSessionId: 279f1814-6e4e-4f39-a354-6447da747ad9
  modified: 2026-07-30T06:44:48.194Z
---

**`Radio/NOTES_WEEKLY_RELEASE.md`** (in the repo) is the runbook for producing a weekly
episode. Written 2026-07-30 after EP 2. Read it first; the same conventions are
summarised in CLAUDE.md under "Weekly radio release".

The traps that cost the most time on EP 2:

- **The Rekordbox export is the tracklist, not the dashboard.** Dashboard = planned,
  export = played. EP 2: 19 planned vs 23 played, completely different order. Build
  cues from the export (`Radio/render/cues_from_rekordbox.py`), then sync the DB.
  Export is UTF-16 tab TSV, columns BY HEADER NAME, and the Artist column is sometimes
  empty with the artist folded into the title.
- **A missing cues column fails SILENTLY.** `render_card` needs `genres`,
  `release_date`, `show_date`, `show_venue` — omit one and that part of the card simply
  isn't drawn. Build cues FROM the DB so video and site can't disagree.
- **`match_mix.py` audio fingerprinting is NOT trustworthy** — on EP 2 it placed tracks
  nearly in reverse at 0.10–0.75 confidence with everything tempo-stretched. Ask for a
  deck **history** export; failing that generate the fill-in times sheet
  (`EPN_times.txt`) and let them type while watching a draft render. Martin typed the
  times over the `guess` value rather than after the `=` — parse from wherever the
  numbers actually are.
- **`INTRO_BEATS`/`OUTRO_BEATS` stage cap** was hardcoded `min(i,4)`, so a new reveal
  beat rendered nothing and a new closing line was missing from a finished 65-min
  video. `preview_bookends.py` holds its own copy of that loop. **Verify a bookend by
  pulling a real frame**, never by reading the code.
- `weekly_prep.py` rewrites the cues every run; it now carries times forward by
  artist+title then by POSITION (export titles differ from the DB's cleaned ones).
- MP4 and all episode material live in **`COMEWITHRADIO/Episode N/`** (the folder was `Radio/` until
  2026-10-19; Keith renamed it on purpose; tools use `_paths.radio_dir()`). Header title is **Come With
  NYC Radio** (`--title`). Artwork must be **square** — portrait letterboxes.

Added 2026-10-19 (Ep 7 / SHOW 11):
- **Dry-run the publish before every drop**: `radio-publish-due?dry=1&station=N` (secret is in
  the cron command). SHOW 11's SoundCloud field held the YouTube URL and go-live was unset.
- **Cards name every upcoming show** (`all_shows.py` -> `cues_with_shows.py`); a Dance
  Infusion show makes an orange spotlight card. Keith liked both (and the 12/31 &ME catch).
- **Read the text of any art before putting it on screen**: the DI3 flier misspells Emmy
  Adelle as "Adele"; Keith said don't use it.
- Compute weekdays, never write them from memory (a draft said "Wed 10/22"; it's a Thursday).

Related: [[project_radio_scheduled_release]], [[project_radio_release_pipeline]],
[[project_beatport_cart_api]].
