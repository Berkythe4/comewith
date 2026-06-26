# Come With — content brief for the customer front-end prototypes

Real brand facts pulled from the repo + live DB (2026-06-26). Use these so the
prototypes are populated with accurate content, not lorem ipsum.

## Brand
- **Name:** Come With  ·  **Base:** Brooklyn, NY  ·  **Site:** comewith.org
- **Instagram:** @comewithnyc  ·  **Email:** berky@comewith.org
- **Founder / resident:** Keith "Berky" Berkman
- **Palette (current):** `#2A1B2E` plum · `#2b1c12`/`#3D2B1F` espresso brown · `#7A6F5F` taupe · `#8FB339` olive-lime (accent) · `#EDE4D3` cream
- **Fonts (current):** DM Sans (body), DM Mono (mono/labels)
- **Voice:** warm, underground, community-first — not corporate clubland.

## What Come With does (4 pillars)
1. **Parties** — house / disco / melodic DJ parties in NYC & Brooklyn.
2. **Production** — DJ, sound, lighting & hosting for other people's events (brings + sets up the full rig).
3. **Dance Infusion** — a charity dance-event series benefiting the **National MS Society**.
4. **Content creation** — recorded sets & artist showcases.

## Dance Infusion — impact (the mission story)
- **Dance Infusion #1** (Sep 8, 2025): 42 tickets → **$1,140 donated to the National MS Society** (39% to mission; proof-of-concept, solo-run).
- **Dance Infusion #2** (May 9, 2026): 117 guests; raised ~$9.5k gross → **~$3,000 net to the National MS Society**.
- **~$4,140+ raised for the National MS Society** across two benefits so far. "% to mission" is the public framing.

## Events (real — for the events section)
| Date | Event | Type | Note |
|---|---|---|---|
| **Jul 11, 2026** | **Come With 7-11** | Party | **UPCOMING** — tickets on Resident Advisor |
| Jun 13, 2026 | Knicks G5 Watch Party | Party | Crossroads Café · ~60 |
| May 9, 2026 | Dance Infusion #2 | Benefit | Signal · 117 · for MS Society |
| Apr 18, 2026 | DI Artist Showcase — Kristen London & 32LVS | Content | recorded showcase |
| Apr 16, 2026 | Maxwell House 4/20 | Production | gear + hosting for a host |
| Jan 17, 2026 | Crossroads Café Artist Showcase | Content | |
| Sep 8, 2025 | Dance Infusion #1 | Benefit | 42 tix · $1,140 to MS |

## The collective / lineup (DJs & performers)
Berky (founder/resident), KRNeY, SPF 50, Kristen London, 32LVS, Just Martin, Henry, Kloud9, DJ Sauci Soni. Venue partners incl. Signal, Crossroads Café, Acoustik Garden.

## Conversion paths (what the site should drive)
- **Tickets / RSVP:** Resident Advisor (tickets), Partiful (RSVPs) — external links per event.
- **Bookings & inquiries:** the inquiry form is the key conversion — it should POST to the `inquiries` table (anon insert is allowed) which feeds the dashboard CRM. Fields: name, email, phone, event type, event date, services, message.
- **Email list:** newsletter signup (the `subscribers` table / subscribe Edge Function exists).
- **Follow:** @comewithnyc, YouTube "Come With!".

## Live data hooks (optional wiring for prototypes)
- Public events: `GET {SUPABASE_URL}/rest/v1/v_public_events` (anon) — already exists.
- Inquiry submit: `POST {SUPABASE_URL}/rest/v1/inquiries` (anon insert, `Prefer: return=minimal`).
- Newsletter: `subscribe` Edge Function.
