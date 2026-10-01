---
name: comewith-radio-video-tool
description: "Radio episode MP4 tool lives at Comewith/Radio — read RUN_SHEET.html first; folders are \"Episode N\"; episode number != station_no"
metadata: 
  node_type: memory
  type: project
  originSessionId: b01e18b6-2c8a-425a-83b2-d8f30e373b0a
  modified: 2026-08-27T00:44:37.532Z
---

Built 2026-08-26. The Come With NYC Radio episode video is made by
`Comewith/Radio/Make Radio MP4.bat` — double-click, type the episode number.
**Open `Comewith/Radio/RUN_SHEET.html` before touching any of it**; it is the
whole process as a flow chart. (Also published at
https://claude.ai/code/artifact/677fbf74-c3b6-4258-abff-683b462a2e69)

Four things that are not obvious from the code:

- **Episode number is not `station_no`.** The Elements run took shows 3–6, so
  NYC Radio Ep 3 is SHOW 7. `make_episode.resolve_episode()` converts; you only
  ever type the episode number.
- **Episode folders are `Radio/Episode N/`** (renamed from `Week N` on
  2026-08-26; `_paths.py` still accepts the old name).
- **Tracklists are typed by hand, not exported from Rekordbox.** The `.txt` in
  the episode folder is the only place the played ORDER and the START TIMES
  exist; everything else comes from the database.
- **All video copy is in `Radio/render/templates.json`** — slides, card band,
  static tags. Editing it needs no code change.

Credential model matters here — see [[comewith-render-tool-credentials]].
Repo/push details in [[comewith-repo-and-push-auth]].
