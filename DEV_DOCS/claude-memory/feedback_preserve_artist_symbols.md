---
name: feedback-preserve-artist-symbols
description: "Artist/track names must be reproduced byte-for-byte — never transliterate or strip symbols, diacritics or unusual casing"
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 279f1814-6e4e-4f39-a354-6447da747ad9
  modified: 2026-07-30T03:12:32.479Z
---

Martin, 2026-07-30: **"symbols in artist names is quite common, you will need to read
and regurgitate them, don't translate."**

Reproduce artist and track names EXACTLY as the source has them. Never normalise for
display or storage:

- diacritics stay: `Theø`, `Zoë Johnston`, `Giolibrí`, `Ammo Amor`
- punctuation stays: `MIND | MATTER` (pipe, not `I`), `X Club.` (trailing dot),
  `Fretless'`, `J.Gill`, `Adastra/Viligir`
- deliberate casing stays: `Let me d&be`, `LEVEL UP`, `nickcurly`, `it's murph`
- ampersands stay: `Above & Beyond`, `Walker & Royce`

**Why:** these are real names as the artists spell them. "Translating" a symbol
misnames a real person or act, and it's how `MIND | MATTER` vs `MIND I MATTER` and
`Above & Beyond` became matching bugs in the first place.

**How to apply:** normalising is fine INSIDE a comparison (`_norm()` in
`Radio/Elements-26/elements_sc.py` strips punctuation to match names), but anything
WRITTEN to the database, a tracklist, an export, a caption or the public site must be
the verbatim string from SoundCloud / Beatport / the artist. When cleaning a title,
only remove what is genuinely redundant — never re-spell what remains. Watch that
console encoding (`errors='replace'`) never leaks a `?` or `�` into stored data.

Related: [[project_radio_shared_pool]], [[feedback_pause_before_major_changes]].
