---
name: project-radio-scheduled-release
description: "Scheduled episode publish now runs through the radio-publish-due edge function, which makes the SoundCloud mix public before the page goes live (EP 1 dead-embed fix)"
metadata: 
  node_type: memory
  type: project
  originSessionId: 279f1814-6e4e-4f39-a354-6447da747ad9
  modified: 2026-07-30T06:44:11.564Z
---

**A private SoundCloud track will not embed** — `oembed` 404s on it, so the site's
player renders nothing. A `200` on the track PAGE proves nothing. EP 1 (2026-07-23)
shipped with a dead embed for exactly this: the link was saved while the track was
private, the page auto-published on schedule, and nothing checked.

**Fixed 2026-07-30.** pg_cron used to call `public.radio_publish_due()` directly —
pure SQL, so it *cannot* make HTTP calls and could never verify the embed. It now
calls the **`radio-publish-due` edge function** (secret-gated via
`RADIO_PUBLISH_SECRET`, `verify_jwt:false`), which per due station:

1. oembeds `mix_sc_track_url` — if it answers, changes nothing
2. if not, finds the track and `PUT track[sharing]=public`, then re-checks
3. if no URL at all, finds the upload by title/runtime and stores it
4. calls `radio_publish_station()` **regardless** — a SoundCloud problem must never
   hold a drop — and writes any failure to `station_notes` so it is visible

`cron radio-publish-backstop` (`3-58/15`) still calls the SQL path, so a broken edge
function delays the SoundCloud fix but never blocks the release.

Supports `?dry=1&station=N` for testing without publishing — use it.

**API facts, both cost time:**
- `/resolve?url=` does **NOT** return a PRIVATE track from its plain permalink. Only
  `/me/tracks` lists an account's own private uploads; match on permalink there.
- Env names are **`SC_CLIENT_ID` / `SC_CLIENT_SECRET`** (not `SOUNDCLOUD_*`). Getting
  this wrong means the token refresh silently never runs and every call goes out with
  an expired token.

Also: `sc-connect` v15 action **`find_mix`** retrieves a link by hand (dashboard: 🚀 Go
live → ☁ Find my upload). **Never ask for a SoundCloud link** — a private upload has
no shareable URL to give. Related: [[project_radio_release_pipeline]],
[[project_radio_weekly_release_runbook]].
