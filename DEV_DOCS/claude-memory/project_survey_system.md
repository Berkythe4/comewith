---
name: project_survey_system
description: "Survey system (DEPLOYED) — build surveys, tokenized/anonymous public form, results tagged to event/actor/customer, wired into the impact report"
metadata: 
  node_type: memory
  type: project
  originSessionId: 59fd8c20-e65f-4382-bad4-d875b13df3be
---

Post-event survey system, built + DEPLOYED 2026-06-29 (migrations 071+072 on prod,
3 edge functions deployed, frontend pushed).

- **Schema (071):** `surveys` (public_token, event_id, status draft/open/closed,
  allow_anonymous), `survey_questions` (qtype: rating/nps/choice/yesno/short_text/
  long_text, options jsonb, sort_order, required), `survey_invites` (per-recipient
  token + event/actor/guest/subscriber tags), `survey_responses` (tags + anonymous),
  `survey_answers` (value jsonb). All tables **admin-only RLS**; public form never
  touches them directly.
- **072:** `v_public_survey` anon view (open + anonymous surveys only → id, event_id,
  public_token, title) so public pages can find an event's survey link.
- **Edge fns (verify_jwt off):** `survey-get` + `survey-submit` (public; token = an
  invite token OR a survey public_token; submit tags the response to the invite's
  event/actor/guest/subscriber, or just the event for anonymous). `survey-send`
  (admin; creates a tokenized invite per recipient + emails each via Resend).
- **Public page:** `survey.html?t=<token>` renders question types, submits via
  survey-submit. survey.html is at site ROOT (use `/survey.html` from subfolder pages).
- **Dashboard module:** Surveys (Audience group, sort 155, signed_off=false → master
  only until released). Builder (questions, reorder, open/close), Share/send (copy
  public link + email tokenized links to typed emails and/or an event's guests),
  Results = per-question aggregates + response list, filterable by event/actor. Code:
  `loadSurveys`/`renderSurveyDetail`/`openSurveyQuestion`/`sendSurvey` in dashboard.html.
- **Impact report wiring:** the public report shows a "Tell us about your night" button
  when an OPEN anonymous survey is linked to that event (reads v_public_survey). See
  [[project_impact_report_supabase]].
- **Campaign attachment (073):** `mailing_campaigns.survey_id`. Attach an OPEN survey
  to a campaign (create/edit form picker; campaign row shows "📋 Survey: <title>"). On
  send, `send-campaign` creates a per-recipient (subscriber + CC) tokenized invite and
  appends a "Share your feedback →" button to each email — so the impact-report email +
  its survey go out together, each response tagged to the recipient. See [[project_email_campaigns]].
- First live survey: "Dance Infusion #2 — Your Feedback" (6 Qs, open, linked to DI#2).
- Backend verified end-to-end (get → submit → response tagged to DI#2 event → cleaned up).
- Pop-out backspace-close fixed app-wide via a capture-phase window keydown guard.
