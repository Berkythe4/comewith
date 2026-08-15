---
name: project-phase-7-status
description: Phase 7 (Dance Infusion event hub) closed 2026-05-28. DI2 seeded from dance_infusion.json. Hub page at events/dance-infusion-2/index-v2.html. CSV importer for RA tickets (Zeffy adapter is a stub).
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

Phase 7 closed 2026-05-28. Five sprint commits.

## What ships

### Seed
- `seed_dance_infusion_2.py` parses `events/dance-infusion-2/dance_infusion.json`
  (the canonical event data file) and loads venue + event + 9 sponsors +
  9 sponsorships + 5 artists + 5 artist_bookings + 5 raffle_prizes +
  4 expenses into the DB. Idempotent.

### Edge Function
- `get-event-hub` (public). POST {slug} → returns event + venue +
  sponsorships (with sponsor name/website/logo) + artist_bookings
  (with artist details) + raffle_prizes. Bypasses admin-only RLS
  on the related tables using service-role.

### Public page
- `events/dance-infusion-2/index-v2.html` — parallel to existing
  `index.html`. Hero with status chip, lineup grid (signature_color
  accents on artist cards), sponsors grouped by tier (Benefactor
  → Vendor), raffle prizes grid.

### Dashboard
- dashboard-v2.html grew three new read-only admin tabs:
  Sponsors, Sponsorships (with sponsor+event embed), Artists.

### Importer
- `import_ticketing.py` takes (event-slug, source, csv-path). Has
  a `resident_advisor` adapter matching the RA export. Adding a
  Zeffy or other source is a 10-line adapter function + register
  in ADAPTERS dict.
- Initial import: 16 RA tickets for DI2 ($350.40 revenue; matches
  AUDIT_REPORT $438 minus 4 comp tickets that have no Order
  number).

## What did NOT ship (deferred)
- Zeffy adapter (the file in the events folder turned out to be
  a Bluevine bank statement, not the Zeffy export). The structure
  for adding it is in place.
- Sponsor / artist / sponsorship WRITE flows in the dashboard
  (Phase 7 was read-only; Phase 4 patterns apply when ready)
- Public events index page (just /events) listing all events —
  only the specific DI2 page exists
- Image uploads for sponsor logos / artist photos (Storage
  buckets exist from Phase 0 but no upload UI)

## Hard-coded values to revisit in Phase 11
- Same as previous phases: any localhost references in functions
  swap to comewith.org

## Open for Phase 8
- Mailing list (self-hosted per decision #8). Public subscribe
  form, double-opt-in confirmation, unsubscribe via tokenized URL,
  segments.
- Resend integration is wired (Phase 5/6) so transactional
  confirmation emails are easy.

## Time tracking
Estimated ~15-20 min wall-clock, came in around 12 min. Five
sprints, each well under estimate. Pattern is repeating itself
across phases — frontend + Edge Function + admin tab work is
faster than the original conservative numbers suggested. See
[[feedback-time-estimates]] for the calibration note.

Related: [[project-phase-6-status]], [[feedback-time-estimates]]
