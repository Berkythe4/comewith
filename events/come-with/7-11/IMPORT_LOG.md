# Come With 7-11 — Attendance + Mailing Import (2026-07-13)

**Event:** `come-with-7-11` (Come With Parties) · **Prod:** `yaytdosxfhcqatmhctzk`
**Backups:** `backups/pre711import_2026-07-13_*.json` (guests, subscribers, subscriber_segments, guest_event_attendance, ticketing)

## Sources
| File | What | Used for |
|---|---|---|
| `20260711-ComeWith-list (3).csv` | 37 RA tickets, 24 unique buyers, emails + Marketing Opt In | guests, ticketing, attendance links, subscribers |
| `20260711-ComeWith-scandata.csv` | 74 barcodes (tickets + guestlist), ScanCount | `total_attendance` + `ticketing.attended` |
| `ComeWith_7-13_guests_partiful.csv` | 38 Partiful RSVPs (22 Going / 16 Maybe), **no emails** | 13 Going → ticketing rows (see below) |
| RA guestlist (imported pre-event via dashboard) | 15 comp links already existed | untouched |

## What was written
- **events**: `total_attendance = 27` (unique barcodes scanned; 28 total scans — KRNeY's barcode scanned twice = re-entry), `status = 'completed'`. 8 of the 27 scans were RA tickets, 19 were guestlist barcodes.
- **guests**: +22 (24 unique buyer emails; Liz McQuillan + chaddercheesy already existed). Source tag `event-import:come-with-7-11`, `opted_in_mailing=true` per locked backfill rule (7 explicit Opt-in, 17 blank, 0 opt-outs).
- **ticketing**: 37 rows, `external_id = barcode` (order numbers repeat across multi-ticket orders), `amount_paid = 0` (Free before 7 — no revenue impact), `attended` from scan data (8 true).
- **guest_event_attendance**: +24 RA links (quantity = tickets per buyer) → 39 total with the 15 pre-event guestlist comps.
- **subscribers**: +22 subscribed (107 subscribed / 119 total). **Chad HG (`chaddercheesy@gmail.com`) NOT re-subscribed** — explicit DI#2 unsubscribe honored.
- **subscriber_segments**: segment `come-with-7-11` on 30 subscribers (22 new buyers + Liz + 7 already-subscribed guestlist guests).

## Flagged — not acted on
- **Victoriarose Vargas is now two guest rows**: `victoriarose.business@outlook.com` (existing) + `tv0192837465@gmail.com` (RA ticket). Email-dedupe can't merge; decide manually.
- **4 guestlist guests with `opted_in_mailing=false` NOT subscribed** (recorded flag honored): SheDay, KRNeY, Lunaera, Janelle Sochet. Flip in dashboard if they should be on the list.
- **Door check-ins with no email** counted in attendance but have no guest row: Garth, Erika Scott, Kyle, Martin, Sergio Pena (+ no-email guestlist names Timmy D, Chloe Pizza Hands, Campbell Gordon, Just Martin).
- **Apple private relay buyer** (6 tickets, 1 scan) subscribed as `w6vwp4z8qs@privaterelay.appleid.com` — relay addresses do deliver, but it's unnameable.
- **Partiful Going ≠ attendance** — its 22 "Going" were not added to the count (no check-in data, no emails).

SQL applied: single idempotent transaction (generated; re-runnable — every insert is guarded by not-exists).

## Follow-up 2026-07-13: Partiful Going → tickets sold

Per Keith: Partiful "yes" RSVPs count as tickets sold, deduped against RA.
- **+13 ticketing rows**, `source='partiful'`, `external_id='partiful:<name-slug>'`, $0, `attended=null` (Partiful exports no check-in data). Event now totals **50 tickets** (37 RA + 13 Partiful). Backup: `backups/prepartiful_2026-07-13_ticketing.json`.
- **Deduped (already on RA side, name-matched):** Liz McQuillan, Victoriarose Vargas, Knostalgia (=Knyckolas Sutherland), Lila (=Lila Bey), Marc (=Marc Getzoff), Steve (=Steve Schuffenhauer), emma stroble (guestlist), Kyle (door-scanned). **Keith Berkman excluded as host.**
- **Inserted:** Alexander Moody, Brendan, Christopher Workman JR, Eddie Mota, Keila Hernandez, Pedro Santiago (+1), Kevin McConville, Knostalgia's +1, Marlo, Michael Borchardt, Rhythm, Shayan Habibi, Shideh Almasi.
- Maybes (16) not counted. `total_attendance` stays 27 (scan-based door count — tickets sold ≠ attendance).

## Follow-up 2 (2026-07-13): Partiful people ARE customers (Keith's call)

No-email guests are fine (guestlist import already created them — Just Martin, Timmy D, etc.). Added **12 new guest records** (email null, `opted_in_mailing=false`, dedupe by name) + attendance links, and wired all 13 Partiful ticketing rows to their guests. **Alexander Moody deduped to his existing DI#2-ledger guest** (`alex.imoody@gmail.com`) — he's now tagged cross-brand (`dance_infusion` + `come_with` + both event segments). Event totals: 52 attendance links (24 RA + 15 guestlist + 13 Partiful). Backup: `backups/prepartifulguests_2026-07-13_guests.json`.
**Dedupe caveat for future imports:** these 12 have no email, so if one later buys an RA ticket, email-dedupe won't match them — check names against no-email guests before inserting.

## Follow-up 3 (2026-07-13): door walk-ins → customer records

Added the 4 scanned door walk-ins as no-email guests + `gea` links (`source='ra_door'`, ticket_type 'Guest'): **Garth, Kyle, Martin, Erika Scott** — barcodes/scan times in each guest's notes for later matching. "Martin" confirmed by Keith as a **different** Martin (not Just Martin / KRNeY). Kyle = same Kyle as the Partiful Going RSVP. Sergio Pena (0 scans, no-show) not added. Event now: **56 gea links** (24 RA + 15 guestlist + 13 Partiful + 4 door). Backup: `backups/predoorguests_2026-07-13_gea.json`.
