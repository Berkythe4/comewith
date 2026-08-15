---
name: feedback-flag-suspicious-artist-matches
description: "A SoundCloud match with a low follower count or zero songs is probably the wrong profile — flag it and advise, never accept it silently"
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 279f1814-6e4e-4f39-a354-6447da747ad9
  modified: 2026-07-30T03:45:56.566Z
---

Martin, 2026-07-30: **"you looked up the wrong kettama artist, i saw low follower
count and no songs were pulled. if you see something like this flag it and advise."**

An automated name→SoundCloud match that lands on a profile with a **low follower
count**, **zero tracks**, or **zero songs pulled** is a red flag for a wrong profile,
not a quiet fact to store. Surface it with a recommendation.

Real case: the Elements lineup matched "Kettama" to a profile with 0 tracks. The actual
artist is `KETTAMA (G-TOWN FOREVER)` — he appended a tagline to his SoundCloud display
name, so an exact-name matcher missed him. Also seen with 0 tracks: MLE, Cloud
Conductor, Diis, DJ Dad, Elkind, Funky Pickles.

**Why:** a booked festival artist with no songs is almost always a matching failure
rather than an artist with no music, and it silently costs a lineup act their whole
crate. Keith/Martin can resolve it in seconds if told; they cannot see it if it just
looks like an empty result.

**How to apply:** after any name→profile match run, report profiles with 0 tracks or
implausibly low followers for a booked act, and propose the likely correct handle.
Do not overwrite a match without saying so. The matcher in
`Radio/Elements-26/elements_tool.py` requires an exact normalised name against
permalink/username/full_name, so any artist whose display name carries an extra
tagline, label suffix or emoji will be missed.

**Naming:** reference such an artist WITHOUT the parenthetical tagline — store and
display `KETTAMA`, not `KETTAMA (G-TOWN FOREVER)`. This does not contradict
[[feedback_preserve_artist_symbols]]: keep the artist's own spelling and casing, but a
tagline appended to a SoundCloud display name is not part of the name. Note that
`sc_playlists.dj_search_params.artists` matches `ra_artists.name` by exact string, so
renaming an artist means updating those episode params too or the lineup stops
resolving.
