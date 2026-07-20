---
name: beatport-cart
description: Build Keith's Beatport cart from a radio station's tracklist — match every song on Beatport, add matches to his cart (or emit buy links), report unmatched. Invoke when Keith says he's ready to buy a station's tracks or types /beatport-cart [ep-number].
---

# Beatport cart builder

Goal: Keith says "make the cart" → he ends up one click from checkout at
https://www.beatport.com/cart with the station's songs in it.

Give an "est X min" up front and a spent/remaining ticker (standing rule).

## 0 · Which station

Argument `[ep-number]` → that station; otherwise the current WORKING station
(newest `status in ('building','testing')`). Query prod via the Management API
(SBP_PAT + SBP_REF_PROD in `.env`, POST /v1/projects/{ref}/database/query,
send a browser User-Agent or Cloudflare 403s):

```sql
select t.title, t.artist_name, t.sc_track_id from sc_playlist_tracks t
join sc_playlists p on p.id = t.playlist_id
where p.station_no = <N>  -- or the working-station filter
order by t.sort;
```

## 1 · Beatport token (file: `.beatport_token.json`, repo root, gitignored)

- If the file exists, use `access_token` from it.
- On 401, refresh: `POST https://api.beatport.com/v4/auth/o/token/` (form:
  `grant_type=refresh_token`, `refresh_token=<stored>`, `client_id=<see below>`)
  and REWRITE the file with the full new JSON (refresh tokens rotate).
  The public client_id is the one Beatport's own docs frontend uses — find it
  at https://api.beatport.com/v4/docs/ (network tab / embedded JS); the
  beets-beatport4 project README documents the same flow.
- If missing/dead beyond refresh, STOP and tell Keith (browser flow, ~1 min):
  log in at beatport.com → DevTools Network → filter `token` → copy the JSON
  response of `/v4/auth/o/token/` → paste; save it verbatim to
  `.beatport_token.json`. Never commit it; never store it in `site_content`
  (that table is anon-readable).

## 2 · Match each track

`GET https://api.beatport.com/v4/catalog/search/?q=<artist> <title>&type=tracks&per_page=10`
with `Authorization: Bearer <token>`.

- Strip noise before searching: "(Original Mix)", "feat./ft.", bracketed
  qualifiers; try artist+title, then title alone if zero hits.
- Score candidates: artist-name and title similarity (case/punct-insensitive).
  A remix ONLY matches if the remixer string matches too — "Song (X Remix)" must
  not match the original mix. Confidence: high / low / none.
- Collect per match: beatport track `id`, `slug`, price if present, URL
  `https://www.beatport.com/track/<slug>/<id>`.

## 3 · Cart (attempt, with graceful fallback)

The cart API is internal/undocumented — introspect before writing:
`GET https://api.beatport.com/v4/my/cart/` (Bearer) to find the default cart id
and the item shape, then try adding each matched track (likely
`POST /v4/my/cart/<cart_id>/items/` with `{"item_id": <track_id>, "item_type": "track"}`
— verify the shape against what GET returned; adapt). Re-GET the cart to
CONFIRM the count actually grew — never claim success from a 2xx alone.

If adding fails after a couple of shape attempts, don't burn time: fall back to
a "Buy links" list (one Beatport URL per matched track) — that's still fast for
Keith. Say plainly which mode happened.

## 4 · Report (all of it in the final message)

- ✅ Added to cart: n/m — link https://www.beatport.com/cart
- 🔗 Buy-link fallback list (if used)
- ⚠️ Low-confidence matches: song → what was matched (Keith should eyeball)
- ❌ Unmatched: artist — title (likely not sold on Beatport; suggest the
  SoundCloud link)
- Optional: offer to write matches back to `sc_song_log` (a `beatport_url`
  column does not exist yet — offer, don't assume).

## Cautions

- Read-only + Keith's-own-cart writes ONLY. Never touch checkout/payment.
- This rides an unofficial token flow (same tier as our RA/SoundCloud usage);
  if Beatport blocks it, report and fall back to links — don't fight it.
