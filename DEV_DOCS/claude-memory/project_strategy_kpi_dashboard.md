---
name: project_strategy_kpi_dashboard
description: How the Strategy-tab KPI cards get their values (v_kpi_dashboard) and which are live-computed vs manual
metadata: 
  node_type: memory
  type: project
  originSessionId: 23f44bb5-a672-44a4-9c2e-b8eac9975d80
---

Strategy tab KPI cards all read **`v_kpi_dashboard`** (`loadStrategy`). Structure: `v_kpi_targets_current` (one row per active metric_key, with target/comparison/unit/label/workstream) LEFT JOIN value sources.

**Fix 2026-06-24 (migration 051, commit b72c681):** event-derived financial KPIs (`di.raised_per_event`, `di.cost_to_raise`, `di.attendance`, `parties.net_pl`, `parties.sell_through`) were always **null** — `v_kpi_dashboard` only joined `v_metric_latest` (= latest `metric_snapshots`) and nobody hand-logs those. Added **`v_kpi_computed`** (derives them live from `v_kpi_dance_infusion`/`v_kpi_parties`, **completed events only**) and repointed `v_kpi_dashboard` to `coalesce(computed, manual-snapshot)`. Also added 4 cards: `di.to_ms_total`, `parties.net_pl_total`, `audience.subscribers` (mailing list), `guest.repeat_pct`. Both views kept **anon-revoked** (financial; verify 401).

**YouTube auto-pull LIVE (2026-06-24, commit 133c479):** Edge fn `pull-youtube-stats` (deployed `--no-verify-jwt`; auth = `YT_CRON_SECRET` query param for cron OR a valid admin JWT) hits YouTube Data API v3 `channels?part=statistics` and upserts `youtube.subscribers` + `youtube.avg_views` (=views/videos) into metric_snapshots daily. Channel = **"Come With!" `UCH5W3mPz3YBTCVI5LOnz3BA`** (handle comewithnyc). Secrets on prod: YOUTUBE_API_KEY, YOUTUBE_CHANNEL_ID, YT_CRON_SECRET. Daily `pull-youtube-stats` pg_cron job (06:00 UTC via pg_net). Manual "↻ YouTube" button on Strategy toolbar (`refreshYoutube` → invoke). `youtube.watch_hours` still manual (needs YouTube Analytics API + OAuth, not the Data API). Instagram auto-pull NOT built (needs Meta app + IG Business token). **Richer YT metrics (migration 052, commit 25eed0b):** function now walks uploads playlist → per-video stats; writes youtube.total_views/videos/total_likes/engagement_rate/days_since_upload + upserts `youtube_videos` (per-video). Strategy page shows those cards + a "Top YouTube videos" list. (Note: kpi_targets has some duplicate active rows e.g. youtube.avg_views ×4 — harmless, v_kpi_targets_current dedups by metric_key.)

**Can't connect YouTube/IG subscribers to the guest list — hard API limit:** YouTube/Instagram do NOT expose subscriber/follower identities or emails (privacy), so there's no way to match them to guests. The EMAIL mailing list IS connected: `v_guest_stats.subscribed` flags each guest as a mailing subscriber by email (86 of 97 guests subscribed); the Guests tab shows it with a "subscribed" filter. That's the only audience↔guest link that's actually obtainable.

**So the value sources are:** (a) live-computed event KPIs via `v_kpi_computed` (di.*/parties.* aggregates, completed-only) — always current; (b) manual `metric_snapshots` via the "Log numbers"/"Log IG followers" flows for the rest (IG followers, YouTube subs, saves/shares, presale velocity, watch hours, avg views). A manual card reads null until someone logs it — that's expected, NOT a bug. To add a new live card: add the value to `v_kpi_computed` + a `kpi_targets` row (workstream/label/target/comparison gte|lte/unit). NOTE: send labels as plain ASCII (em-dash mojibakes through the Management API).
