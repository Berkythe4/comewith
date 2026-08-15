---
name: project_radio_episode_planning
description: "Radio episode planning — edit any episode, future skeletons, assignment, DJ scoped link (dj.html), recap notes, social-post link"
metadata: 
  node_type: memory
  type: project
  originSessionId: db51d261-6033-4aa4-a2da-cd68f16c298d
  modified: 2026-07-28T19:37:53.415Z
---

Built 2026-07-28 (migration 130, commits 60fac76/9d3596b). sc_playlists gained
recap_notes, assigned_to (profiles), assigned_actor_id (actors), dj_token,
dj_search_params (jsonb {weeks,genres}); social_posts gained station_id; new
'planned' status (future skeletons — the one-building partial unique index only
constrains status='building', so many 'planned' coexist).

- **Edit any episode** (incl. completed/archived): Radio Control Center (ccMode) →
  open the episode → ✎ Details. `ccWithStation(id, raOpenEpisodeDetails)` loads
  that station into raState.playlist. raOpenEpisodeDetails is now async (loads
  raLoadEpLookups: profiles + person-actors + recent social_posts).
- **Future skeletons**: CC "➕ Plan a future episode" (data-ra-ccplan →
  raCcPlanEpisode) inserts a numbered status='planned' row, drop = latest+7.
- **Assignment**: episode details 👤 Assigned to — a user AND/OR a DJ (person
  actor). Shown as a chip on CC cards (raAssignChip).
- **DJ scoped link**: details → 🎧 DJ access → generate/copy/revoke dj_token
  (raDjGenLink/raDjRevoke). Link = comewith.org/dj.html?ep=<token>. **dj.html** is
  a login-free token-gated workspace; **dj-station** edge fn (service-role, token
  is the only credential; revoke=clear dj_token) returns ONLY NYC artists playing
  within dj_search_params.weeks (+ genre overlap) with their sc_artist_cache
  songs, plus the episode's current sc_playlist_tracks. **DJ CAN ADD SONGS**
  (2026-07-28, migration 132 allows source='dj'): dj-station action='add'/'remove'
  (token-gated; add dedups on the unique key + source='dj'; remove only pulls the
  DJ's own source='dj' picks, never curated). dj.html has ＋add/✓added·remove per
  song. Builder shows a green "🎧 DJ pick" badge on source='dj' tracks.
- **Special editions** (migration 131: sc_playlists.edition_name + edition_seq):
  Control Center "✨ Special edition" (data-ra-ccedition → raCcSpecialEdition) —
  name + first drop + # of drops; a row per day (title + DJ assign: user or
  actor); inserts N Planned episodes (numbered on, consecutive drop_dates), shown
  with an ✨ edition chip. Built for the Come With Elements Radio 4-drop festival
  run (Thu 8/6–Sun 8/10; Martin+Henry as support DJs, 32LVS/Taylor guest, Keith).
- **Recap notes**: details → 📝 Recap notes (INTERNAL free-form, never public;
  not in RA_PL_COLS-exposed public paths). For Janelle's recap content.
- **Social-post link**: details → 🔗 Linked social post (1:1 via
  social_posts.station_id). The post modal shows the linked episode + its recap
  notes (raLoadPostEpisode) so Janelle drafts recap with the material in front.

- **Elements Radio tool** (2026-07-28): the Come With Elements Radio edition
  (station_no 4–7, Ep1–4 = Thu/Fri/Sat/Sun of the Aug 6–9 festival) is scoped to
  the FESTIVAL lineup, not the weekly NYC pool. `Radio/Elements-26/elements_tool.py`
  matched all 110 main+Thursday artists to SoundCloud, pulled songs into
  sc_artist_cache, created **source='elements'** ra_artists (next_event_date NULL
  so they stay OUT of the general Artist Radio list + Scene tools, which load
  ra_artists gte today), and set each episode's `dj_search_params =
  {pool:'elements', day, artists:[…]}`. dj-station gained a **fixed-lineup mode**:
  when `params.artists` is set it returns exactly those artists in order (no date
  window) + `scope.pool/day`; dj.html header reads "Your crate = the Elements
  <day> lineup". All 4 episodes have DJ links. NOT included: Disco Den (~55 house
  DJs) + Fun Factory (games). To re-run/extend, edit LINEUP in elements_tool.py.

See [[project_radio_release_pipeline]], [[project_sc_match]].
