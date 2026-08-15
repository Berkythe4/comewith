---
name: project_api_integrations
description: Status of external API/data integrations + the in-dashboard API map; Instagram scaffold; RA/Partiful have no public API
metadata:
  type: project
  originSessionId: 23f44bb5-a672-44a4-9c2e-b8eac9975d80
---

In-dashboard **API map** (2026-06-26, commit a728b6a): a "🔌 APIs" button in the main-header opens `#apiOverlay` (reuses the workflow-map wf-* shell) showing `API_CONNECTIONS` grouped by `API_GROUPS` = Established / Pending / No-official-API / Future. Edit the `API_CONNECTIONS` JS array to update it. Sits next to the "🗺️ Workflow" button.

**Status:**
- **Established (live):** YouTube Data API (pull-youtube-stats, daily cron — see [[project_strategy_kpi_dashboard]]); Resend email (transactional/campaigns/conversations + webhook — see [[project_email_conversations]]); Supabase (platform).
- **Pending — needs Keith's setup:** Instagram Graph API. `pull-instagram-stats` Edge Function is **scaffolded but NOT deployed** (mirrors pull-youtube-stats; writes instagram.followers/media/reach/profile_views into metric_snapshots). Needs: Meta/Facebook dev app + IG switched to Business/Creator linked to a FB Page + long-lived token (instagram_basic + instagram_manage_insights + pages_read_engagement) + the IG Business account id. Then set secrets **IG_USER_ID + IG_ACCESS_TOKEN** (+ optional IG_CRON_SECRET) and deploy `--no-verify-jwt`. Note IG long-lived tokens expire ~60 days (need refresh).
- **No official/public API (workaround only):** **Resident Advisor** — no public API; ticketing comes via RA CSV export → in-app importer; unofficial ra.co/graphql exists but ToS-restricted/unstable (don't rely on it). **Partiful** — no public API; manual for now.
- **Future opportunities (real APIs, not yet wired):** GA4 Data API (website funnel — highest-value next, free); TikTok Display/Business API; Eventbrite/DICE (if ticketing moves there — they have proper APIs unlike RA/Partiful); Spotify/Bandsintown (talent vetting, not own-performance).

Pattern for adding a pull-* integration: Edge fn auth = cron secret OR admin JWT; upsert into metric_snapshots (metric_key, value, captured_on, series_id null, source); add a kpi_targets row + the metric shows on Strategy.
