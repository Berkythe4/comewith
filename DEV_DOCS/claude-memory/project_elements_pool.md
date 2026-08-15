---
name: project_elements_pool
description: "Elements Festival pool — Thursday (Ep1) carries ALL festival producers; every SoundCloud song cap removed; 18 zero-song acts are mix-only DJs and that's correct"
metadata: 
  node_type: memory
  type: project
  originSessionId: 9caef454-ca9a-4416-aa06-a6721798c2fc
  modified: 2026-08-03T18:33:28.136Z
---

Elements Festival 2026 lives inside the Come With Radio stack, not as a separate
tool: offline scripts in `Radio/Elements-26/` seed `ra_artists (source='elements')`
+ `sc_artist_cache`, and the 4 episodes are `sc_playlists` rows scoped via
`dj_search_params`. See [[project_radio_episode_planning]].

**Thursday = the whole festival (2026-08-03, Keith's own edition).** Ep1 is the
early slot with a 10-act bill, so `elements_thursday.py` scopes it to every
PRODUCER across all four days + Disco Den (139 artists), ordered Thu first, each
tagged with the day it plays. Fri/Sat/Sun (Martin, Henry, unassigned) stay on
their own night — that's intentional, not an oversight. The day map is derived
from what Ep1–4 already hold; do not re-declare the lineup a third time.

**No song caps anywhere.** Removed `fetch_songs(want=15)`, `dj-station`'s
`.slice(0,12)`, and `elements_disco.py`'s private 15-item copy of the rule. Above
& Beyond went 15 → 404 songs. A cap here is invisible: a short crate reads as the
artist's whole catalogue. `sc-enrich`'s `.slice(0,200)` is a storage guard, fine.

**Zero songs is usually the CORRECT answer.** 18 Elements acts post only DJ sets
(Lightcode = 20-min guided meditations, Sirens = a 60-min podcast, Koopmusik =
live sets). They are DJs, `is_producer=false`, and they're excluded from producer
scopes on purpose. Only 2 of the 19 were actually wrong profiles. Still unresolved
and left alone deliberately: **MLE** (verified @mlemusicc has 0 uploads, @mle8 has
the music), **Sirens** (@sirens_la may be a different act), **DJ Dad** (no
confident match). Per [[feedback_flag_suspicious_artist_matches]] these stay
flagged, not guessed — nobody re-checks a profile that looks filled in.

**An artist name can exist under several sources.** `ra_artists` holds Brainrack
and Flash Gea twice — once from Elements with a SoundCloud, once from RA with
`soundcloud = null`. Any name→row lookup must prefer the row that HAS a profile,
or the artist silently gets an empty crate. Fixed in `dj-station` v8; check the
same pattern anywhere else names are used as keys.

**Deploying edge functions:** the Supabase CLI rejects the newer `sbp_v0_…` PAT
(`LegacyInvalidAccessTokenError`). Use the Management API multipart endpoint
`POST /v1/projects/<ref>/functions/deploy?slug=<slug>` with `metadata` + `file`
parts, and send a browser User-Agent or Cloudflare returns 403 code 1010.
Related: [[feedback_prod_migration_apply]].
