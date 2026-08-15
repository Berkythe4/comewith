---
name: project_ra_market_radio
description: "Resident Advisor public-data integration (078) — RA Market scheduling intel + Radio station from upcoming artists' SoundCloud; private RA data stays CSV"
metadata: 
  node_type: memory
  type: project
  originSessionId: 6ae7a3fb-5f8f-4348-8f22-cc711985cf02
---

Built 2026-07-08 (migration 078, commits 82f4de3 + earlier). Uses RA's undocumented
GraphQL at **https://ra.co/graphql** — introspection is ENABLED, no auth needed for
public data. Headers required: Content-Type, Referer https://ra.co/events, a browser
User-Agent.

**Key API finding (probed live):** PUBLIC = `eventListings(filters:{areas:{eq:INT},
listingDate:{gte,lte}}, pageSize, page)` → event id/title/date/startTime/attending
(RSVP)/interestedCount/isTicketed/pick{id}/genres{name}/flyerFront/contentUrl/
venue{name}/artists{id,name,soundcloud,instagram,followerCount,image,contentUrl}.
Artist type also has bandcamp/website/soundcloudUser. **Area 8 = New York.** GATED
(returns AUTH_NOT_AUTHORIZED / 401, needs to BE the logged-in promoter): Event
`totalRevenue`/`totalTicketsSold`/`totalTicketAllocation`/`guestlist`, and the whole
`Ticketing_*` namespace + `multiEventDashboard`. **So YOUR money/tickets/guestlist
stays on the CSV flow (import + RA-export) — decided with Keith.** RA GraphQL repeats
the same event across pages → dedupe by id (upsert-batch dup key errored).

**Schema (078, admin RLS, anon-blocked):** `ra_events` (ra_id pk, event_date,
venue_name, attending, interested_count, is_ticketed, is_pick, genres[], lineup jsonb,
content_url, fetched_at) + `ra_artists` (ra_id pk, name, soundcloud, follower_count,
image, next_event_date/title/venue). Module `ra-market` (Insights, sort 196,
signed_off=false → master-only). `ops.ra_area_id` (=8) editable in Site Editor.

**Edge fn `pull-ra-market`** (verify_jwt ON; accepts admin JWT OR service-role — auth
checks the JWT **role claim** === service_role, NOT string-match, because the project's
SUPABASE_SERVICE_ROLE_KEY env is the new sb_secret_ format, so string compare failed).
Pages eventListings (area/days/maxPages), replaces the forward window, dedupes events +
artists (keeps each artist's soonest show). Live run: **569 NYC events, 908 artists,
701 with SoundCloud.** Callable by pg_cron later.

**Dashboard RA Market screen** (`loadRaMarket`/`renderRaMarket`/`renderRaRadio`/
`raPlay` in dashboard.html): best-nights table (opportunity = median RSVP / sqrt(event
count)) with YOUR avg attendance by weekday alongside; top venues + hot genres; RA-pick
+ ticketed %. **Radio**: `raRadioList()` dedupes artists by soundcloud within 2/3/4
weeks; ▶ loads their SoundCloud profile into `w.soundcloud.com/player` iframe (a profile
URL plays their tracks); "Copy lineup" export. Radio is INTERNAL for now — offered a
public radio.html as a follow-up (more ToS-sensitive to republish RA data publicly).

**Radio v2 (2026-07-08, migration 079, commit 89e016d) — SoundCloud integration:**
RA only gives an artist's SoundCloud PROFILE url, not tracks. So `sc-tracks` edge fn
(admin/service, same role-claim auth) uses SoundCloud's INTERNAL read API
(api-v2.soundcloud.com): extracts a `client_id` from soundcloud.com JS bundles (regex
`client_id[:=]"..."`), caches it in `site_content ops.sc_client_id`, refreshes on 401;
`/resolve?url=` → user, `/users/{id}/tracks`. **"Songs not sets" = kind==='track' AND
duration <= maxMinutes (default 15min)** — SC "sets" are kind=playlist (excluded); DJ
MIXES are kind=track so duration is the real filter. Returns plays/created_at/permalink/
artwork. Verified: Four Tet 2 songs; a set-only artist 0 songs / 5 sets excluded.
`079`: `sc_playlists` + `sc_playlist_tracks` (admin RLS, anon-secure, unique
(playlist_id,sc_track_id)). Radio UI: sort (soonest/followers/name) + search; per-artist
"🎵 songs" expander (Recent/Popular); persistent station (add/remove/reorder/clear),
**▶ Play station uses the SC Widget API** (w.soundcloud.com/player/api.js — load track,
bind FINISH → load next), Copy links.

**SoundCloud playlist EXPORT — BUILT + LIVE (2026-07-08, migration 080, commit 4947310).**
Keith has Artist Pro; registered app "Come With Radio" (SC_CLIENT_ID/SECRET secrets set on
prod). OAuth 2.1 Authorization-Code + PKCE. Endpoints: authorize
`https://secure.soundcloud.com/authorize`, token `https://secure.soundcloud.com/oauth/token`,
API `https://api.soundcloud.com`; header `Authorization: OAuth <token>` (NOT Bearer); create =
POST /playlists `{playlist:{title,description,sharing:'private',tracks:[{id:Number}]}}`, update =
PUT /playlists/{id}; refresh tokens single-use; tokens ~1h. **Registered Redirect URI MUST be
`https://yaytdosxfhcqatmhctzk.supabase.co/functions/v1/sc-oauth`** (redirect_uri_mismatch if not).
`080`: `sc_oauth` singleton token store (admin RLS) + `sc_playlists.sc_playlist_id/url`.
Functions: **sc-oauth** (public/--no-verify-jwt, callback: code→token, stores, redirects
dashboard.html?sc=connected) + **sc-connect** (admin: status/start/export; start makes PKCE
state+verifier→authorize URL; export refreshes-if-expired then POST|PUT /playlists, stores
sc_playlist_id+url). UI: state-aware export modal (Connect my SoundCloud full-redirect → else
Create/Update playlist → opens SC link); ?sc= return toast. Verified backend: status
configured:true, start builds valid PKCE URL. NOTE: SC deprecating numeric track id → urn
(soundcloud:tracks:N) — if exports start failing on track ids, switch to urn. Untested end-to-end
(needs Keith's browser login).

**Radio v3 — producer detection + real filtering (2026-07-09, migration 081, commits
c6c6ef7/e64b1d3):** BIG rebuild after "not usable" feedback. (1) **Connect button was
dead** — it lives in the kpiModal which is OUTSIDE #panel-ra-market, so the panel-scoped
click handler never fired. Moved ALL radio clicks to a DOCUMENT-level delegated handler
(guarded to only act on data-ra-* targets). (2) **Producer detection**: `081 sc_artist_cache`
keyed by normalized SoundCloud URL (SURVIVES RA re-pulls, which delete/recreate ra_artists);
`sc-enrich` edge fn batch-resolves + classifies producer = has ≥1 original SONG (kind=track,
45s–maxMin dur; sets/mixes = long tracks excluded). Caches songs. Frontend: "Producers only"
toggle (default ON) hides pure DJs; "Scan for producers" batch-scans window (20/call, live
progress, cancelable); expanding one artist scans just them. Genre filter (ra_artists.genres
now populated by pull), window 2/3/4/6w, sort soonest/most-songs/followers/name. Producer
badge + song count; songs from cache instant, recent/popular. **CRITICAL BUG FOUND: RA stores
soundcloud URLs with `www.` but SC resolve 404s on www.** (and 404≠401 so no client_id
refresh) — normalize by stripping www everywhere (sc-enrich norm + frontend raKey). Pre-scanned
the whole NYC window: **~55% are producers** (206/377 at midpoint). Note sc-tracks fn now
orphaned (raToggleSongs uses sc-enrich).

**Radio v4 (2026-07-09, migrations 082-084):** (a) **Publishable tracklist** — RA event.cost
is PUBLIC free-text ("$20+","30"); pull stores next_cost/next_event_url on ra_artists; each
station track captures show_date/venue/cost/url; "📋 Publish tracklist" = shareable gig-guide,
also becomes the SoundCloud playlist DESCRIPTION on export (sc-connect accepts body.description).
(b) **Unlisted radio page** — radio.html (noindex, token-gated) + get-station public fn; stations
have public_token + published(default FALSE) → "Preview · unlisted" until Keith flips it; dashboard
"🌐 Preview page" button. (c) **SoundCloud CONNECT WORKS** — sc_oauth has a live token (Keith
connected successfully). (d) **Follow/repost tool (084)** — sc-social (admin): follow PUT
/me/followings/{id}, repost POST /reposts/tracks/{id} (+undo), OAuth token w/ refresh, logs to
sc_social_log; per-artist ＋follow / per-song ↻repost buttons, manual+selective, ✓ done-state,
repost confirms, optimistic w/ revert. Header `Authorization: OAuth <token>`.
(e) **BPM/Camelot — mostly a DEAD END for this use case:** SoundCloud tracks are unreleased edits
NOT in any BPM/key DB. Tested MusicBrainz→AcousticBrainz (keyless): only **1/15** matched.
GetSongBPM registration was broken for Keith (accept-all email + dead link). Spotify audio-features
deprecated for new apps. Only audio-analysis (infeasible here) would work. Kept sc-bpm as a
"released tracks only" best-effort button + keyToCamelot() util (verified 10/10). Set expectations
LOW. **Still OPEN/requested: Ticketmaster** (official, free key at developer.ticketmaster.com,
filter Dance/Electronic at source — widens MARKET intel, no SoundCloud so not radio) **+ Dice**
(NO official API — reverse-engineer like RA or skip). Awaiting Keith's TM key + Dice call.

**Radio v5 (2026-07-09, migration 086):** SPLIT into two Insights modules — **market**
("Best Nights", panel-market, loadMarket/renderMarket) + **ra-market** ("Artist Radio",
radio-only now). Best Nights: night score = demand(RSVP)+star power(headliner follower
reach, ra_events.lineup[].follower_count /1000) vs competition; weekday table + **📅 month
CALENDAR** (heat-colored, click night→lineup, marketState.view/month/selDate); venues
**paged+sortable** (marketState.venuePage/venueSort); **hot genres CLICKABLE** (marketState.genre)
→ filters venues + "Biggest artists coming" (top by follower_count, IG+SC links). Radio: window
now has a **start date** (raState.radioStart, default today) via shared **raWindow()** used by
raRadioList/renderRaRadio/raScan; per-artist **📷 IG** button. **Ticketmaster LIVE** (085
pull-ticketmaster, ra_events.source='tm', STRICT genre=Dance/Electronic; no RSVP→demand stays
RA-only; TM_API_KEY set). Player fix: moved to a persistent sticky **#raDock** OUTSIDE the
re-rendered body (was 'working then stopped' — re-renders wiped the iframe); playback driven by
Widget API READY→play() (auto_play URL param unreliable). SoundCloud SYNC-BACK: sc-connect
action=sync GETs the exported playlist, reconciles final order/adds/removes (replaced BPM idea).
**Instagram auto-follow is IMPOSSIBLE** (no API) — IG buttons open the profile for manual follow;
automated follow/repost is SoundCloud-only.

**Radio v6 (2026-07-10, migrations 088-089, commits d1affd1→c93e81c):** DATA-COMPLETENESS +
watchlist + station hardening. (a) **"Few August shows" root cause:** pull-ra-market only fetched
~12 pages (600 events) = all July; Aug/Sep + their artists never pulled. Cap → 40 pages (~2000):
now ~1011 events Jul–Oct (Aug 298). **088 `ra_artists.source`** ('ra'|'tm'): both pulls delete only
their OWN source (pull-ra-market's window-delete had been wiping TM too). (b) **TM performers now
upserted into ra_artists** (source=tm, no socials/RSVP) so they show in artist views. (c) **Radio
shows no-SoundCloud artists** (removed the `.not('soundcloud',is,null)` filter; MISS VEE now appears)
— play/songs hidden, badge "no SoundCloud on RA". **VeeDay simply isn't in RA's lineup data** (can't
pull what RA doesn't publish → add via watchlist "create as new artist"). (d) **Best-Nights DOW
weighting** `DOW_WEIGHT`=[.65,.15,.30,.55,.80,1,1] (Sun..Sat) so Monday is never "best"; best
flipped Mon→Fri; calendar dropped the confusing median number. (e) **Watchlist (089
`watchlist.actor_id`):** "Collaborator" reason; **multi-select** reason filter (toggle chips,
marketState.watchReasons[]); each artist's upcoming nights as **little boxes** (artistShows() helper
matches event lineups); note inline next to reason; collaborators **link to a roster actor or create
as new artist** (writes actor_id + actor_roles role='artist', "🔗 roster" tag). (f) **SoundCloud
station hardening (sc-connect):** export **pre-validates each track** via public GET /tracks/{id} +
skips uploader-blocked/deleted (embeddable_by==='none' or non-200) → reports `skipped`, doesn't fail
whole playlist; **body is now form-encoded Rails params** (`playlist[title]`, `playlist[tracks][][id]`)
because SC rejects JSON ("Could not parse JSON request body"); **sync is non-destructive on an
incomplete snapshot** — a reorder is remove-then-re-add so the API briefly omits in-flight tracks;
old sync deleted them (DATA LOSS). Now trusts `track_count`, deletes NOTHING if fewer come back,
returns `incomplete` (GET uses ?access=playable,preview,blocked). Recovered 2 lost tracks + removed a
stray duplicate "Weekly station" row. (g) **Player STILL "opens but doesn't play"** — env, not code:
SC's OWN ▶ fails → Chrome ad-blocker or third-party-cookie/Tracking-Protection blocking the embed.
Added `allow="encrypted-media"` + guaranteed "open ↗" fallback; **open Q: Keith's incognito test**
(fails→cookies, works→extension). raLoadPlaylist can still race duplicate station rows (unique guard
would fix).

To get service_role for local testing: Management API `GET /v1/projects/{ref}/api-keys?
reveal=true` (delete the temp file after). See [[project_api_integrations]],
[[project_site_review_audit]], [[project_public_events]].
