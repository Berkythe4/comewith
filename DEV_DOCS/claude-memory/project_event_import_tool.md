---
name: event-import-tool
description: "Drag-and-drop event-export importer (RA tickets/scans/guestlist + Partiful) on the event hub Customers tab — engine, tests, standing rules"
metadata: 
  node_type: memory
  type: project
  originSessionId: fd236bc8-6af6-467b-bb1d-c8f4a3db2b2f
---

Built 2026-07-13: **"⬆ Import event exports"** on the event hub Customers tab (dashboard.html). Drop multiple CSVs at once — auto-classified by header signature (RA ticket list / RA scan data / RA guest list / Partiful) — hit Run, and it creates guests+ticketing+gea, sets `total_attendance` from scan data (unique barcodes scanned) + marks the event completed, and updates the mailing list (subscribers + per-event & brand segments), then renders a confirmations/dedupes/flags/errors report. **Idempotent — re-running the same files writes nothing new.**

Architecture: pure logic in `assets/import-engine.js` (parseCsv, classifyExport, matchName, buildImportPlan — no DOM/Supabase), executed by `hubRunImport()` in dashboard.html via supabase-js as the logged-in admin. **Tests: `node scripts/test_import_engine.mjs` (40 assertions against the real 7-11 exports in `events/come-with/7-11/` — those CSVs are test fixtures, keep them).** All standing rules encoded: barcode-keyed ticketing, Partiful Going = ticket+customer (host excluded, Maybes ignored), name-dedupe incl. "Name #2" collapse + email-handle matching (Knostalgia↔knostalgiamusic@), door walk-ins become no-email customers, opt-outs never re-subscribed, brand segment from events.series.

DEPLOYED 2026-07-13 (commits d26d489+7adf7ac pushed to master → Netlify). **2026-07-14 (091 + cd1551f):** `guest_event_attendance.attended` = person-level scanned-in (true/false/null-unknown; ticketing.attended only covers RA barcodes); importer sets it from scan data (48 tests); 7-11 backfilled 21✓/22✗/13—; hub Customers tab is now a sortable table w/ role/source/scanned filters + search. Same session: Artists tab got multi-select ✉ Email + ⧉ Copy emails (shared `copyEmailList()`, comma-separated for Google Drive sharing) and event People got ⧉ Copy emails; send-actor-email edge fn redeployed to SKIP the "Open in Come With dashboard" footer link for `source_kind='event_people'` (performers can't access the dashboard — Keith's ask; other sources keep the link). Related: [[cw711-import]], [[project-mailing-list-architecture]], [[project-email-conversations]].
