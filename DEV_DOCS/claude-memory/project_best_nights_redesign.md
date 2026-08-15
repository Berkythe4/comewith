---
name: project_best_nights_redesign
description: "Best Nights rework — signed-off plan (Buzz, NYC-local, public heat-map, avatars) + Phase-1 build status"
metadata: 
  node_type: memory
  type: project
  originSessionId: db51d261-6033-4aa4-a2da-cd68f16c298d
  modified: 2026-07-28T13:39:09.048Z
---

Best Nights tab is being reworked from competition-only ("when to throw") into
the scene control room. Proposal signed off by Keith 2026-07-27 (artifact:
"Best Nights — Redesign Proposal"). **Decisions:**
- **Buzz score** weighting: show-demand 40 / reach 35 / catalog 15 / editorial 10,
  **plus a small bump for producers with a release in the last 6 months**. Demand
  = RA attending/interested on their NYC dates (lives on ra_events → needs an
  artist→event join; that's why Buzz is built in the Best Nights tab, which loads
  events, not the artist panel which doesn't).
- **NYC-local**: pull city from ANY easily available source. DONE via SoundCloud
  profile city (sc-match captures `city`/`country`, fills `ra_artists.city`).
- **Default view**: Map.
- **Heat-map is PUBLIC-facing** — keep internal triggers (who-to-book, Buzz) OUT
  of it. Internal triggers live in a SEPARATE internal control center that also
  shows the map. Phase order confirmed: P1 now → P2 public club heat-map (clubs on
  map → every radio track from artists tied to that venue, all weeks, as a
  playlist) → P3 listener/artist avatars (saved playlist + shows attended/going).

**Phase 1 SHIPPED 2026-07-27** (commit 3b15bb6):
- NYC-local "📍 NYC" tag + "📍 NYC locals" filter (city from SC).
- "➕ Add" artist/venue tool (auto-captures artist SC/city/followers via sc-match;
  saves artist to ra_artists flagged `is_partner` — partners ALWAYS shown,
  bypassing date-window + producers-only). Venues save to the EXISTING CRM
  `venues` table, extended by migration 127 with area/lat/lng/links/genres/
  is_partner/source (one venue entity for CRM + heat-map). Migration 126 added
  ra_artists.city + is_partner. NOTE: `venues` was NOT new — it's the CRM's
  actor-linked table; 127 fixed 126's bad unique(lower(name)) index.
- Also that day: unified "↻ Refresh music & data" button replaced the scan/match
  clutter (permanent toolbar = discovery; collapsible station builder = episode).

**Best Nights TAB — Buzz + Who-to-book SHIPPED 2026-07-27** (commit e8f9b5b):
- `mkBuzzList()` computes per-artist Buzz 0–100 (demand40/reach35/catalog15/
  editorial10 + recent-release bump; log-normalized) from ra_events lineups
  (demand/pick) + shared raState.artists/cache (reach/catalog/city).
- "🎧 Who to book" section in renderMarket (after the KPI cards) ranks lineup
  artists by Buzz with next-show/RSVP/reach/catalog/pick/recent + 📍 NYC tag,
  SoundCloud link, watch btn; "📍 NYC locals" toggle (marketState.bookNyc).

Also fixed same day (commit fb61b6e): Add-modal "Look up" was dead (main script
is type="module" → inline onclick can't reach module fns; rewired via delegated
listeners on #kpiModalOverlay). NYC filter showed 0 because 0/1889 had a city —
sc-enrich now captures SC profile city (migration 128 = sc_artist_cache.city) and
a one-time backfill resolved the pool (691 cities, 430 NYC-local).

**NYC SCENE CONTROL CENTER shipped 2026-07-28** (commits bc6898c/a984cb1) — Best
Nights reworked into a tabbed internal control center (renderMarket → Overview /
Artists / Venues). Data layer `scArtistData`/`scVenueData`/`scSnapshot`/`raBuzzMap`
(all memoized on a data signature). Always-on SNAPSHOT (both the Scene tab AND the
radio panel): events & artists per source (RA/DICE/TM), SC linked/scanned/
producers, songs+avg/artist, NYC count, venues. ARTISTS = sortable table
(buzz/rsvp/reach/songs/shows/next) + filters (source/genre/NYC/producers/has-SC/
search) + row drill-in with a BUZZ BREAKDOWN (demystifies the score) & all shows
w/ platform+ticket links. VENUES = aggregated from shows (shows/avg+peak RSVP/
artists/genres/platform tags/capacity from venues table) + drill-in. Ticketing
PLATFORM derived from source, tagged across shows/artists/venues. Buzz carried to
Artist Radio (🔥 Buzz sort + per-card chip). User feedback that drove it: "who to
book" was clunky → became the sortable Artists explorer; wanted a data web +
snapshot of what's captured. Tested via extracted-fn harness (ALL PASS).

**Cross-nav + internal Scene Map SHIPPED 2026-07-28** (commit 6cbb3f1):
- Control-center cross-nav: artist drill-in → click a venue (jumps to Venues tab
  expanded); venue drill-in → buzz-ranked artist chips jump to the artist.
- New left-nav item **🗺️ Scene Map** (key `scene-map`, nav_group Radio sort 30,
  pushed client-side in renderNav; panel `#panel-scene-map`; loadTab case;
  loadSceneMap/renderSceneMap/sceneSelect). Leaflet 1.9.4 + CARTO dark tiles
  loaded on-demand from jsdelivr (no CSP on the dashboard). Plots venues by
  geocoded coords, dot size=shows/demand, color=platform; click → shows/artists
  panel + green lines to venues that SHARE artists (the scene web); chips
  deep-link into the control center.
- Venue coords: geocoded via OSM Nominatim + a Photon fallback (NYC-bbox
  validated, messy names cleaned) into `venues.lat/lng`, source='scene'.
  **206 of 235 located** (~29 unnamed/private spaces remain). loadMarket +
  scVenueData carry lat/lng. Scripts: scratchpad/geocode_venues.py + geocode_retry.py.
- Scene Map FILTERS + radio color (commit cf22106): filter by min shows / genre /
  "🎙 On the radio" (+ platform, size-by). Venues hosting artists who've appeared
  on any radio station (`raState.radioArtists` = split set of sc_playlist_tracks
  credits, 36 distinct) get a DISTINCT lime color + legend; `onRadio` flagged on
  venues AND artists in the data layer (memo sig includes radioArtists.size).

STILL TODO: better venue geocoding coverage (via addresses); the PUBLIC club
heat-map (hold until ≥4 weeks of data, keep internal triggers OFF it — the Scene
Map is the internal precursor); P3 avatars. See [[project_sc_match]] and
[[project_ra_market_radio]].
