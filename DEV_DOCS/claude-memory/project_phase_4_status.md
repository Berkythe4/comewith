---
name: project-phase-4-status
description: Phase 4 (admin dashboard writes) closed 2026-05-28. dashboard-v2.html now does inline status edits + Add Income + Add Expense (with receipt upload). Ledger replaceable from Sheets.
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

Phase 4 closed 2026-05-28 (same day as Phases 2 + 3). Four sprints,
commits `2ada308`, `f0fdd78`, `ba05471`, plus the sprint-4 polish
commit. Each is a self-contained increment.

## What ships
- Inline status `<select>` editors on inquiries and agreements rows
- "+ Add income" button → modal → INSERT into public.income
- "+ Add expense" button → modal with file upload → upload to
  Storage 'receipts' bucket → INSERT into public.expenses with
  receipt_path set
- Generic modal infrastructure (`FORM_DEFS`, `openModal`, file
  upload, ESC/click-outside/cancel close, auto-focus first field)
- Toast notification system reused across status writes and
  modal inserts; flash-green on newly-inserted rows after reload
- JWT-expired detection in BOTH inline edits and modal submits
  triggers an explicit `sb.auth.signOut()`

## Out of scope (deferred)
- CRUD on clients, equipment, events — Phase 4.5 or rolls into
  the relevant domain phase (Phase 7 for events, Phase 6 for
  clients via the public flow)
- Editing / deleting existing income/expense rows — only inserts
  for now; if Berky mis-keys, fix via db.py or Supabase dashboard
- Viewing uploaded receipts (would need signed URLs from the
  private bucket) — Phase 4.5 or Phase 6
- Agreement creation form — agreement workflow is heavyweight
  enough that it should land with Phase 5's send-agreement
  Edge Function

## Out-of-band staging state (still applies)
- `uri_allow_list`: includes `http://localhost:8765/**`
- Site URL: `http://localhost:3000` placeholder
- A few test rows still in inquiries / agreements / clients from
  Phase 2 debugging — visible in v2 dashboard for verification

## Open for Phase 5
- Edge Functions need the `supabase` CLI installed locally
  (`supabase functions serve` for dev, `supabase functions deploy`
  for staging)
- First function: `send-agreement` — generates PDF from a
  template (existing ComeWith_Events_Services_Agreement.html
  could be the source), uploads to 'agreements' bucket, emails
  the signing link via Resend
- Second function: `inquiry-notify` — triggers on insert to
  public.inquiries (or called from the future Phase 6 public
  form), sends Berky a "new inquiry" email
- Resend integration needs RESEND_API_KEY from existing setup
- Magic-link redirect target should be configured when Phase 6
  brings the customer portal online

## How to apply
- Adding a new write form: add a key to `FORM_DEFS` in
  `dashboard-v2.html`. The modal plumbing handles the rest.
  For file uploads, also add an entry to `RECEIPT_FIELD_TO_COLUMN`.
- Adding inline status edit on a new table: add the table +
  valid statuses to `STATUS_VALUES`, render `statusSelect(...)`
  in the renderer.

Related: [[project-phase-3-status]], [[project-phase-2-status]]
