---
name: comewith-render-tool-credentials
description: SBP_PAT is a Supabase management token (never share it); RLS blocks the public key from radio tables; get-station?t=TOKEN is the credential-free read route
metadata: 
  node_type: memory
  type: reference
  originSessionId: b01e18b6-2c8a-425a-83b2-d8f30e373b0a
  modified: 2026-08-27T00:44:55.340Z
---

Established 2026-08-26 while making the radio tool shareable.

- **`SBP_PAT` in `Comewith/.env` is a Supabase MANAGEMENT token**, not a database
  login. It runs arbitrary SQL on production and deploys/deletes edge functions.
  Never send it to a collaborator. `.env` is gitignored; `.env.example` in the
  repo has every non-secret value pre-filled. Mint a **separate** PAT per machine
  at supabase.com/dashboard/account/tokens so a lost laptop is one revocation.
- **The publishable key cannot read `sc_playlists` / `sc_playlist_tracks`** — RLS
  denies it (verified, HTTP 401 42501). There is no read-only DB route.
- **The credential-free route is `get-station?t=<public_token>`**, the endpoint
  radio.html already uses. It returns a full episode — *including an unpublished
  one* — with every field the render cues need. `tracklist_from_txt.py --token`
  uses it, and falls back to the token stored in the episode folder's
  `episode.json`. A token reads one episode and can never write: `--write-order`
  refuses to run from one.
- **Dashboard admin ≠ Supabase project access.** martin@, henry@ and berky@ are
  all `master_admin` in `profiles`, but only berky@ owns the Supabase org.
- **The Supabase CLI rejects the current PAT format** ("Invalid access token
  format" — CLI 2.101.0 expects 40 chars, the token is 43). Deploy edge functions
  via the Management API instead: multipart POST to
  `/v1/projects/{ref}/functions/deploy?slug=NAME`, and **send a browser
  User-Agent** or Cloudflare answers 403 "error code: 1010".

Used by [[comewith-radio-video-tool]].
