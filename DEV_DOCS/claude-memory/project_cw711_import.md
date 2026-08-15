---
name: cw711-import
description: Come With 7-11 (2026-07-11) attendance + mailing import applied to prod 2026-07-13 — patterns for future post-event imports
metadata: 
  node_type: memory
  type: project
  originSessionId: fd236bc8-6af6-467b-bb1d-c8f4a3db2b2f
---

Come With 7-11 post-event import applied to prod 2026-07-13. `total_attendance=27` from RA **scan data** (unique barcodes scanned; scan export = `ScanCount` per barcode, covers tickets AND guestlist). 37 RA tickets → ticketing rows keyed on **barcode** not order number (orders repeat across multi-ticket orders; `import_ticketing.py` keys on order number and would collapse them — don't use it for multi-ticket RA lists). +22 guests/+22 subscribers (segment `come-with-7-11`; segments are per-event slugs, there is no master `come_with` segment in data). Chad HG = chaddercheesy@gmail.com, an explicit DI#2 unsubscribe — never re-subscribe. Full log + flags (duplicate Victoriarose guest, 4 opted_in=false guestlist performers not subscribed) in `events/come-with/7-11/IMPORT_LOG.md`. RA export tip: the basic "list" export lacks Email/Opt-In columns — need the fuller export (like DI#2's) plus the separate scan-data export. Partiful exports have no emails, RSVPs only — reference, never imported. Backups `backups/pre711import_2026-07-13_*.json`. Related: [[project-historical-events-backfill]], [[project-mailing-list-architecture]].

**Rule from Keith (2026-07-13): Partiful "Going" RSVPs count as tickets sold AND as customers (guest records) even without email**, deduped by NAME against RA tickets/guestlist/scans AND against existing guests, host excluded. For 7-11: +13 ticketing rows `source='partiful'` (`external_id='partiful:<name-slug>'`, $0, attended null) + 12 new no-email guests (opted_in_mailing=false) + gea links; Alexander Moody matched his existing di2-ledger guest and got cross-brand segments. No-email guests can't email-dedupe — future imports must name-check against them. Maybes don't count; attendance stays scan-based. Door walk-ins with scans also become no-email guests (gea source 'ra_door'; barcode+scan time in guest notes for later matching) — 7-11's Garth/Kyle/Martin/Erika Scott; this 'Martin' is a DIFFERENT person from Just Martin/KRNeY per Keith. Event totals: 50 tickets sold (37 RA + 13 Partiful), 56 gea links (24 RA + 15 guestlist + 13 Partiful + 4 door), 27 attended.
