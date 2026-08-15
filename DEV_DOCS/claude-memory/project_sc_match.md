---
name: project_sc_match
description: "sc-match edge fn resolves a SoundCloud profile from an artist name; radio panel \"Match SoundCloud\" + Beatport-for-all"
metadata: 
  node_type: memory
  type: project
  originSessionId: db51d261-6033-4aa4-a2da-cd68f16c298d
  modified: 2026-07-27T20:57:27.716Z
---

Built 2026-07-27 so DICE/Ticketmaster artists (and ~339 RA artists with no link)
stop being dead ends in the radio Artist-Radio panel.

- **`sc-match` edge fn** (admin-only, deployed to prod): name → SoundCloud
  profile via api-v2 `/search/users`. Reuses sc-enrich's scraped `client_id`
  (`site_content.ops.sc_client_id`). CONSERVATIVE: accepts only a candidate whose
  permalink/username/full_name NORMALIZES to the query (NFKD strip diacritics +
  drop non-alphanumeric; official/music/dj affixes allowed via stripAffix),
  ranked verified → followers → track_count; no exact match ⇒ matched:false.
  `write:true` fills EMPTY `ra_artists.soundcloud` (ilike name, is null) — never
  overwrites. Verified on real DICE names (ARTBAT→artbatmusic,
  Curbi→curbiofficial, Interplanetary Criminal, Massane, Main Phase, bad tuner…).
- **Panel**: "☁ Match SoundCloud (N)" button (shows when noScCount>0) →
  `raMatchSc()` collects no-link artists in window (all sources), batches of 25 to
  sc-match with write, mirrors matches into `raState.artists`, then calls
  `raScan()` to pull their songs. So DICE artists fold into the normal producer
  workflow (scan reads SC songs via sc-enrich).
- **Beatport catalog for EVERY artist**: the 🛒 Beatport button (data-ra-bpcat)
  now renders on all artist cards, not just no-SoundCloud ones. track-sources
  `artist_catalog` already returns per release: bpm, song_key, camelot,
  release_date, label, price, sample_url (preview) — so "all Beatport songs with
  BPM/key/Camelot" was already there; it just needed exposing.

So per artist now: SoundCloud songs (match→scan) + Beatport catalogue (🛒), both
with previews; BPM/key/Camelot come from Beatport. Camelot shows as harmonic
codes (chips), NOT yet a visual color wheel — offered as a follow-up.
See [[project_ra_market_radio]] and [[project_radio_release_pipeline]].
