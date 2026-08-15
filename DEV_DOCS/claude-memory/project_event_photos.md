---
name: event-photos
description: "Event photo galleries — migration 090, hub Photos tab, public event.html linked from Recent Rooms (deployed 2026-07-14)"
metadata: 
  node_type: memory
  type: project
  originSessionId: fd236bc8-6af6-467b-bb1d-c8f4a3db2b2f
---

Deployed 2026-07-14 (commit 57250b6; migration 090 applied to prod + verified). **Hub Photos tab**: drag-drop upload → browser makes TWO sizes via downscaleImage (~1600px full at 0.85 + ~480px thumb at 0.8) → public `event-photos` bucket at `event/<eventId>/<ts>_<name>_full|_thumb.jpg`; per-photo caption, "on site" toggle, reorder (full reindex ×10), delete (removes storage files too). **Public `event.html?id=`**: reads `v_public_recap` (featured events; also used by homepage Recent Rooms, exposes id) + `v_public_event_photos` (anon view: is_public photos on featured/public events). Recap videos embedded up top (YouTube/SoundCloud — zero Supabase egress), lazy thumb grid + full-size lightbox below. Homepage recap cards click through via `data-event`; `[data-media]` play chips still lightbox.

Security follows the 030 pattern (013 default privs auto-grant ALL to anon on new tables — table grant stripped, SELECT on the view only; financial views re-verified 401). Bucket is PUBLIC: paths are URL-reachable regardless of is_public — the flag only gates listing; never upload sensitive files there.

**Cost basis (Free plan, verified 2026-07-14):** org on Free (1GB storage / 5GB egress); storage then = 5.7MB. Two-size scheme ≈ 18MB per 50-photo event → years of headroom; egress supports ~1.5–2k gallery views/mo; upgrade trigger = Pro $25/mo (250GB egress + on-the-fly image transforms, which would replace the two-size scheme). Watch the egress graph in the Supabase dashboard. Related: [[event-import-tool]].
