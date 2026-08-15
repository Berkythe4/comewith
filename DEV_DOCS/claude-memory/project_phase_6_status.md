---
name: project-phase-6-status
description: "Phase 6 (customer-facing flows) closed 2026-05-28. index-v2.html inquiry form (anon writes), customer_portal-v2.html (signed-in agreements view), inquiry-notify Edge Function. Anon-RLS mystery resolved as a PostgREST RETURNING quirk."
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

Phase 6 closed 2026-05-28. Five sprint commits, all individually
mergeable. Final commit chain: `216234b` sprint 1 → `11244fb` sprint 4.

## What ships

### Frontend
- `index-v2.html` — public inquiry form. Vanilla JS + supabase-js
  CDN. Anon `insert(payload)` with NO `.select()` chain (sends
  `Prefer: return=minimal`)
- `customer_portal-v2.html` — magic-link login, then post-login
  shell showing customer's agreements with status badge, total,
  and "Review & sign" deep link to `sign.html?token=X` for any
  unsigned-but-sent agreement that still has a valid token
- Both files at repo root, parallel to existing `index.html` and
  `customer_portal.html` (legacy Apps Script versions untouched)

### Edge Function
- `inquiry-notify` — public, --no-verify-jwt. POST {email}, looks
  up most recent inquiry from this email in the last 5 minutes
  via service-role, emails all master_admins via Resend with the
  full inquiry details, Reply-To set to the inquirer

## The Phase 2 anon-RLS resolution (important)

While testing sprint 1, the same `42501: new row violates RLS`
error from Phase 2 reproduced. After deeper investigation, the
root cause turned out to be **PostgREST's default
`Prefer: return=representation`**, which makes the INSERT also
SELECT the inserted row back. anon has no SELECT policy on
inquiries, so the SELECT-after-INSERT fails, and Postgres
surfaces it as a misleading "WITH CHECK violation" error.

Fix: don't ask for the row back when anon writes. supabase-js
chains `.insert(x)` → no select needed → uses `return=minimal`
under the hood. Updated `project_anon_rls_sql_editor.md` to mark
the original hypothesis superseded and document the actual cause.

## Schema-side cleanup also done
- Inquiries policies restored to migration-spec form
  (Anyone-can-insert, Admins read/update, Master-admin delete).
  Phase 2 debugging had left the INSERT policy with
  `TO authenticated, anon` instead of the default `TO public`.
- The "NUCLEAR allow all" debug policy was dropped.
- Test rows from this phase's debugging were deleted from
  inquiries.

## Out of scope (deferred)
- services_selection.html and equipment list rewrite — not
  blocking; original scope had them in Phase 6 but the user
  picked the 2-page version. Worth doing in Phase 6.5 if needed.
- Rate-limiting on inquiry-notify to prevent spam.
- HCaptcha or similar bot prevention on the public form.
- Customer portal showing inquiries history.

## Hard-coded values to revisit in Phase 11
- Same as Phase 5: signing URL base in send-agreement function
  is still `http://localhost:8765/sign.html`, needs
  `https://comewith.org/sign.html` at cutover.

## Open for Phase 7
- Dance Infusion event hub (public event pages, ticketing CSV
  import from Zeffy + Resident Advisor, sponsor admin UI)
- Different domain than inquiries/agreements — separate Phase 6
  work doesn't carry over much, but the patterns (anon writes
  via PostgREST, customer auth via magic-link) do
- The Dance Infusion event already has its own `events/` folder
  in the repo with historical data + an AUDIT_REPORT.md

## Time tracking note
Estimated 75 min, spent ~70 min. The Phase 2 anon-RLS debug
that landed in sprint 1 (~10 min extra) was offset by sprints
2-4 running faster than estimated. Process-feedback memory at
[[feedback-time-estimates]] applies.

Related: [[project-phase-5-status]], [[project-anon-rls-sql-editor]],
[[feedback-time-estimates]]
