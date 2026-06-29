# Come With — Session Summary for Claude Desktop (2026-06-29)

Covers work since the Dance Infusion full-handoff (session started 2026-06-28). All of
this is **applied to prod (Supabase `yaytdosxfhcqatmhctzk`) and pushed to `master`**
(Netlify live). Migrations **067–074**; edge functions **survey-get / survey-submit /
survey-send** deployed and **send-campaign** redeployed.

## 1. Dance Infusion impact report — now public + dashboard-editable
- Moved from static local JSON to **Supabase** (migration 067): `events.impact_report` (jsonb)
  + `events.impact_report_public` (the publish **toggle**) + anon view `v_public_impact_report`.
- **Edit it in the dashboard:** open the DI#2 event → hub header → **Impact report** (text,
  hero + inline photos, the publish toggle).
- Public pages `events/dance-infusion/di-02-2026-05/reports/impact-report.html` + `public-audit.html`
  read Supabase (fall back to local JSON only on localhost). The old `/staging` gate is gone — the
  **toggle is the gate**. Homepage `#di` shows a **"Read the #2 Impact Report"** button when published.
- Content locked per Keith: attendance **117**, DI#1 sponsors **0**, **reach removed**, audit goal
  **50%**, Yankees donor = **New York Yankees**, human-moment quote + DI#3 "what's next" render from
  saved content. **The DI#2 report is published.**

## 2. Pricing tool (new Sales module, between Inquiries & Agreements)
- Migrations 068 (`pricing_config`, admin single-row + nav row) + 070 (`events.quote` jsonb).
- Pure engine `assets/pricing-engine.js` (+ `scripts/test_pricing.mjs`, 14 passing tests).
- Quote builder: **DJ / equipment rental** (live from `equipment_inventory.daily_rate`) **/ labor /
  lighting / travel (mileage + drive time) / weekend-peak-rush surcharges**; every default editable;
  **per-DJ custom rates**; copy + print/PDF.
- **Save a quote to an event**, or with no event it **creates a planning event** from the quote.
- Still **master-only** (built, not signed off).

## 3. Survey system (new Audience module)
- Migrations 071 (5 survey tables, admin RLS) + 072 (anon `v_public_survey`) + 073
  (`mailing_campaigns.survey_id`). Edge fns survey-get / survey-submit (public) + survey-send (admin).
- Public **`survey.html`** (rating / NPS / multiple-choice / yes-no / text). Dashboard **Surveys**
  module: builder, open/close, share public link OR email **personal tokenized links** (event guests /
  segment / picked people), results filterable by **event / person**.
- **Tagging:** every response ties to event + actor/guest/customer/subscriber (anonymous public link
  tags to the event only).
- **Wired together:** attach a survey to a **campaign** → each recipient gets a personal link in the
  email; place it with `{{survey}}` / `{{survey_link}}` (else appended). Impact report shows a
  "Tell us about your night" button. **First DI#2 feedback survey is live + open.**
- Still **master-only** (built, not signed off).

## 4. Campaigns
- **Edit** any not-yet-sent campaign; **CC** field (069); **attach a survey**; campaign rows show a
  📋 survey badge; plain-text line breaks now render as `<br>`; survey appears in test send + preview.

## 5. Event hub + operations
- **Guest list (pre-assign expected customers)** on the Customers tab — searchable picker of existing
  guests/contacts, add-new, remove; the Customers list is searchable.
- **Equipment load-in checkoff** (074 `equipment_usage.loaded_at`) — persistent checkboxes, "X/Y
  loaded", "Mark all loaded".
- Fixes: event hub now reloads the `audited`/`financials_released` flags (were saved but not re-read);
  content/showcase events show **net P&L** in the Events "Result" column (the Crossroads showcase's
  −$1,400 was hidden behind IG reach).

## 6. Other
- **Social calendar email** (✉️) — sends an inline-HTML snapshot via `send-notice`, with a recipient
  picker (team + contacts).
- **Users** access chips colour-coded: **blue = role default**, **green/red = grant/revoke override**.
  (Reminder: marketing scope sees Guests, Subscribers, Campaigns, Social Calendar, Conversations,
  Notes, Site Editor by default; Surveys once signed off.)
- **Homepage collective** ("Residents & regulars") = portrait **photo cards** using each artist's
  profile photo (managed in the Artists tab). Falls back to an initial tile if no photo.
- **Pop-out fix:** removed click-outside-to-close (a drag-select ending off the modal was dismissing
  forms) — pop-outs close via Cancel / Esc only.
- **In-dashboard Workflow map** gained **Quote/Pricing**, **Email campaign**, and **Survey/feedback**
  steps (numbers auto-computed by position).

## Open / next
- **Release to staff** (Team → Modules, flip `signed_off`): **Pricing, Surveys, Templates** are
  master-only today.
- Add artist **photos** for the new collective cards; flip more artists onto the collective
  (only Berky + KRNeY are on).
- External: verify **comewith.org** as a sending domain in Resend before the first real email blast.

## Conventions reminder (unchanged)
- Prod is `yaytdosxfhcqatmhctzk`; apply DDL via the Supabase Management API with `SBP_PAT` (CLI is
  linked to staging). Roles: master_admin / sub_admin / customer; RLS uses `is_admin()`. Never a
  blanket `grant ... to anon`. Financial views stay anon-revoked. Series contract: 'Come With Parties'
  / 'Dance Infusion' exact-match.
