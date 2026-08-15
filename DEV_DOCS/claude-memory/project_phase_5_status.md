---
name: project-phase-5-status
description: "Phase 5 (transactional Edge Functions for agreement flow) closed 2026-05-28. Three functions deployed, sign.html ships, end-to-end agreement signing works once RESEND_API_KEY is set."
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

Phase 5 closed 2026-05-28 (same day as Phases 2/3/4). Five sprint
commits, all individually mergeable. Final commit: `daa30ae`.

## What ships

### Edge Functions (deployed to staging)
- `send-agreement` — admin-auth required. Takes {agreement_id},
  creates agreement_links token, emails customer via Resend with
  signing link, sets agreement status='sent'
- `get-agreement-by-token` — public, --no-verify-jwt. Takes
  {token}, returns agreement payload (no admin-only fields) +
  link metadata. Returns 410 for expired tokens
- `mark-signed` — public, --no-verify-jwt. Takes {token,
  signature_name}, records signature, sets status='signed',
  marks link used, emails ALL master_admins via Resend with
  "Agreement signed by X" notification

### Frontend
- `sign.html` at repo root. Reads `?token=X`, calls
  get-agreement-by-token, renders agreement, accepts typed-name
  signature, posts to mark-signed
- `dashboard-v2.html` Agreements tab grew an "Actions" column
  with Send/Resend buttons that invoke send-agreement

### Schema
- No new migrations. Uses existing `agreement_links` table
  (created in 004_inquiries_agreements.sql) and
  `agreements.client_signature_url` repurposed to hold the
  typed-name signature (column name is a hand-me-down)

## Required out-of-band setup
- **RESEND_API_KEY** must be set as a Supabase Edge Function
  secret before any emails actually send:
    `supabase secrets set RESEND_API_KEY=re_xxx`
  Without it, sends fail with a clear error in the toast/UI
  ("RESEND_API_KEY not set...")

## Hard-coded values to revisit in Phase 11
- Signing URL base in `send-agreement/index.ts`:
  `SIGN_BASE_URL = "http://localhost:8765/sign.html"`
  Will need to switch to `https://comewith.org/sign.html`
  at production cutover

## Email identities
- From: `Berky <berky@comewith.org>` for customer-facing,
  `Come With <berky@comewith.org>` for admin notifications
- Reply-To: `berky@comewith.org` for everything
- `hello@comewith.org` is NOT configured on the Resend domain yet;
  if added later, can switch admin-notify From to it

## What did NOT ship (deferred)
- PDF generation. Web signing is the e-signature record. If a
  PDF compliance copy is ever needed, would slot in after
  mark-signed (use pdf-lib in Deno or Resend Attachments)
- inquiry-notify Edge Function — would fire on new public inquiry.
  Belongs with Phase 6 when the public form goes live
- Real Site URL (still `http://localhost:3000` placeholder in
  Supabase Auth config)

## Open for Phase 6
- Public inquiry form on index.html → INSERT inquiries via anon
- Customer portal (customer_portal.html) → magic-link login,
  shows signed-in customer's agreements + invoices
- services_selection.html → public-facing service catalog
- inquiry-notify Edge Function so Berky gets emails when public
  form submits
- Will be the **first real test of anon-RLS** (the Phase 2
  SQL Editor anomaly hypothesis verifies here)

## How to apply
- New Edge Functions: `supabase functions new <name>` to scaffold,
  edit `index.ts` to use raw `Deno.serve()` (cleaner than the
  scaffold's `withSupabase` wrapper), keep `deno.json` minimal,
  deploy with `supabase functions deploy <name>` (add
  `--no-verify-jwt` for public endpoints)
- Always use `npm:@supabase/supabase-js@2` and
  `npm:resend@4` import specifiers, not the deno.json mappings

Related: [[project-phase-4-status]], [[project-phase-3-status]]
