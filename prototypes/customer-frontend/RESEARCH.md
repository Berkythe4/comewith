# Come With customer front-end — research dossier (2026-06-26)

Compiled to inform the redesign. Three strands: (A) evidence-based web-design
principles, (B) comparable-company website analysis, (C) event/ticketing UX.
Every claim is tied to a verified source URL.

---

# A. Evidence-based design principles (NN/g · Baymard · web.dev · W3C)

## 1. First impressions & aesthetics-usability
Users form an aesthetic judgment in ~**50ms** (Lindgaard et al. 2006) and it rarely
changes; a polished design also triggers the **aesthetic-usability effect** (people
rate attractive sites as more usable and forgive flaws). → The hero must look
polished + on-brand instantly.
Source: https://www.nngroup.com/articles/first-impressions-human-automaticity/

## 2. Visual hierarchy & scanning (F-pattern / layer-cake)
Users scan, they don't read. Without clear headings they fall into the inefficient
**F-pattern**; strong headings/subheadings enable the efficient **layer-cake** scan.
Front-load key words.
Sources: https://www.nngroup.com/articles/f-shaped-pattern-reading-web-content/ ·
https://www.nngroup.com/articles/layer-cake-pattern-scanning/

## 3. Above the fold / hero
Above-the-fold gets **57%** of viewing time; first 3 screenfuls = **81%**; **65%+**
of above-fold time is in the top half. Put the core message + primary CTA up top.
Source: https://www.nngroup.com/articles/scrolling-and-attention/

## 4. Navigation — visible & short
Hiding nav behind a hamburger **cuts discoverability ~½**; visible nav was **>20%**
better. Keep desktop nav visible, few items, consider sticky.
Source: https://www.nngroup.com/articles/hamburger-menus/

## 5. CTA — wording, contrast, placement
Specific labels ("RSVP for [Event]") beat vague "Get Started". What matters is
**high contrast/isolation**, not the specific color (HubSpot +21%, Slack +34% CTR
from contrast changes). Place CTAs next to what they act on (up to ~29% lift).
Sources: https://www.nngroup.com/articles/get-started/ ·
https://www.nngroup.com/articles/closeness-of-actions-and-objects-gui/ ·
https://cxl.com/blog/which-color-converts-the-best/

## 6. Speed / Core Web Vitals
Targets: **LCP < 2.5s, CLS < 0.1, INP < 200ms**. Measured results: Economic Times
**−43% bounce**, Rakuten **+61% conversion**, Vodafone **+8% sales**, AliExpress
**−15% bounce**.
Sources: https://web.dev/case-studies/vitals-business-impact ·
https://web.dev/articles/defining-core-web-vitals-thresholds

## 7. Mobile-first
Mobile ≈ **52–64%** of global traffic. Primary actions in the center **thumb zone**;
touch targets ≥ **1cm (~0.4in)**.
Sources: https://www.nngroup.com/articles/touch-target-size/ ·
https://gs.statcounter.com/platform-market-share/desktop-mobile/worldwide/

## 8. Typography & readability
Line length **50–75 chars**; generous body size + line height; font choice alone
changed reading speed up to **35%**.
Sources: https://www.nngroup.com/articles/legibility-readability-comprehension/ ·
https://www.nngroup.com/articles/glanceable-fonts/

## 9. Color & contrast (WCAG AA)
Text contrast **≥ 4.5:1** (normal), **≥ 3:1** (large text ≥18pt / UI). SC 1.4.3.
Source: https://www.w3.org/WAI/WCAG22/Understanding/contrast-minimum

## 10. Social proof & trust
Named testimonials (name + affiliation), real numbers, recognizable logos; users
trust external reviews more than on-site quotes. Four credibility factors: design
quality, disclosure, current content, connection to the web.
Sources: https://www.nngroup.com/articles/social-proof-ux/ ·
https://www.nngroup.com/articles/trustworthy-design/

## 11. Form design (inquiry / RSVP)
**Single column**, **minimize fields** (~8; long forms lose **17–18%**), **labels
above fields** (never inline-only placeholders).
Sources: https://baymard.com/learn/form-design ·
https://baymard.com/blog/avoid-multi-column-forms ·
https://baymard.com/blog/mobile-forms-avoid-inline-labels

## 12. Whitespace
Generous negative space groups content, guides the eye to the CTA, and improves
scannability.
Source: https://www.nngroup.com/articles/characteristics-minimalism/

### Top 10 rules to design by
1. Win the first 50ms (polished, on-brand hero).
2. Lead above the fold (core message + primary CTA up top).
3. Design for scanning (headings, front-loaded words).
4. Nav visible & short.
5. Primary CTA = highest contrast + specific wording.
6. Fast: LCP<2.5s, CLS<0.1, INP<200ms.
7. Mobile-first, center thumb-zone actions, ≥1cm targets.
8. WCAG AA contrast (4.5:1 / 3:1).
9. Forms: single column, ~8 fields, labels above.
10. Credible social proof with room to breathe.

---

# B. Event / nightlife / ticketing UX (with live teardowns)

**Source tiers:** [Primary] = original research (NN/g, Baymard, Stanford GSB, WCAG, web.dev, peer-reviewed, first-party platform docs); [Vendor] = platform/blog (directional).

## Events listing / calendar
- Each card = a mini-page; **date top-left** (users scan top/left first). Keep card layout identical so date/venue/lineup/price compare column-to-column. [Primary] https://www.nngroup.com/articles/list-entries/
- Surface **price or "Free / RSVP"** on the card; limit badges to **2–3** (Sold Out / Few Left / Benefit). [Primary] same
- Filtering is conversion-critical (weak list UX → 67–90% abandonment vs 17–33%); show live match counts. [Primary] https://baymard.com/research/ecommerce-product-lists
- Location → list + link to Google/Apple Maps, don't embed. [Primary] https://www.nngroup.com/articles/store-finders-and-locators/
- Scarcity: high-demand → raw counts feel scarcer; low-demand → percentages. Keep truthful. [Primary] https://pmc.ncbi.nlm.nih.gov/articles/PMC10135727/

## Event detail
- Lead with **who-what-when-where** + primary CTA above the fold (consider sticky); then description → lineup → run-of-show → FAQ. [Vendor] Eventbrite; [Primary] https://www.nngroup.com/articles/scrolling-and-attention/
- Vibe hero image from a **past** event (82% of attendees prefer feel-conveying images). FAQ pre-answers 21+, refunds, transit, accessibility. [Vendor] Eventbrite

## Ticket / RSVP CTAs + external ticketing
- Specific labels ("Get Tickets" / "RSVP"), one dominant high-contrast primary CTA, ≥4.5:1 text / ≥3:1 UI contrast, slim sticky mobile CTA. [Primary] https://www.nngroup.com/articles/get-started/ · https://webaim.org/articles/contrast/
- **Resident Advisor** = embed listing widget and/or "Tickets via Resident Advisor ↗"; RA hosts checkout (no on-domain checkout). [Primary] https://support.ra.co/article/7-ticket-widget
- **Partiful** = "RSVP on Partiful ↗" link (no embed/paid). [Primary] https://partiful.com/
- Hand-offs: open new tab, label it, show the partner brand (jarring unbranded jumps trigger distrust). [Primary] https://www.nngroup.com/articles/new-browser-windows-and-tabs/ · https://baymard.com/blog/perceived-security-of-payment-form

## Email signup
- Email-only / minimal fields (3-field ~25% vs 6+ ~15%); real incentive (presale codes, early lineup drops); two-step click-trigger or exit-intent beat passive bars; double opt-in for deliverability. [Vendor] OptinMonster/BDOW; [Primary] https://www.nngroup.com/reports/email-newsletter-design/

## Past-event media / recaps (social proof)
- Make past events a **permanent browsable recap archive** (photos/aftermovies) — the content flywheel. Best model: **Boiler Room** (boilerroom.tv) tagged recap archive. Show numbers only when impressive. [Primary] https://www.nngroup.com/articles/social-proof-ux/ ; [Vendor] Ticket Fairy

## Charity / benefit framing
- Lead with **one identifiable beneficiary story, not stats** (adding stats lowers giving). Pair with "$X = outcome". Highlight **one suggested amount**. Frame as "support the cause," not gift-in-exchange. Build trust before the ask; show National MS Society credibility (Charity Navigator). [Primary] GoFundMe Pro (Small/Slovic) https://pro.gofundme.com/c/blog/identifiable-victim-effect/ · Stanford GSB https://www.gsb.stanford.edu/insights/how-nonprofits-make-ask-framing-donation-requests · https://www.nngroup.com/articles/commitment-levels/
- Best model: **Hyde Park Jazz Benefit** — merged "tickets or donate" CTA, 501(c)(3) + one-line mission, honoree recognition.

## Mobile
- Mobile-first: >68% of ticketing on phones; buying is last-minute (57% ≤1 week before). Tap targets ≥48×48px; LCP ≤2.5s, INP ≤200ms, CLS ≤0.1 (1s→10s load = +123% bounce). Add-to-calendar, click-to-map. [Primary] https://web.dev/articles/accessible-tap-targets · https://web.dev/articles/vitals ; [Vendor] Eventbrite/BRI

## Live teardowns
| Site | Listing | Ticket CTA | Recaps |
|---|---|---|---|
| Public Records (publicrecords.nyc) | upcoming-only; date/type/time/room/artists | "Get tickets" → Dice | none |
| Elsewhere (elsewhere.club/calendar) | **best filters** (type+timeframe+room+genre) | "Buy Tickets" → Eventbrite | none |
| Nowadays (nowadays.nyc) | no own calendar → RA page | off-site RA | strong "Safer Space" + "Residents" values pages |
| Boiler Room (boilerroom.tv) | upcoming carousel | "Tickets" link-out | **best-in-class tagged recap video archive** |
| Hyde Park Jazz Benefit | single benefit | **merged tickets-or-donate** (PayPal+Zelle+check) | 501(c)(3) + mission + honorees |
| Partiful (partiful.com) | location pre-RSVP; mutuals surface | one-tap "Going" | "see who's going" feed |

## Blueprint distilled
Hero = past-event vibe photo + dual identity (parties + Dance Infusion benefit) + one primary CTA + two-step email capture. Events = consistent cards (date/venue/lineup/type tag) + filters + **Upcoming vs permanent Past/Recaps archive** (Boiler Room model). RA → themed widget or "Tickets via RA ↗"; community → "RSVP on Partiful ↗". Dance Infusion = one beneficiary story + "$X=outcome" + suggested amount + MS Society credibility + merged tickets-or-donate. Mobile-first, ≥48px targets, fast.

# C. Comparable companies (13 verified, June 2026)

**A. Production / creative agencies** — *23 Layers* (twentythreelayers.com): personality-led "We are…" hero, sub-brands as nav lanes, downloadable capabilities deck as a lead magnet. *Eventique* (eventique.com): outcome CTA copy ("Request Proposal / Tell Us About Your Event") repeated down-page, result-framed portfolio tiles, "New" badge for the next event. *Empire Entertainment* (empireentertainment.com): cinematic full-bleed photo hero + "Watch our reel", dedicated **Talent** nav lane, tap-through case-study pages.

**B. DJ collectives / party brands** — *Nowadays / Mister Saturday Night* (nowadays.nyc): photo-carousel atmosphere, ticketing delegated to RA behind one CTA, standout **Safer Space** page with affordability line ("if cover is a barrier, reach out"). *Public Records* (publicrecords.nyc): values-led mission hero before nav, **Live/Club/Etc. calendar filter**, CTA paths separated by intent. *The Lot Radio* (thelotradio.com): a **live audio player IS the hero**, genre tags + resident "R" badges, playful sticker texture over a clean base.

**C. Nightlife / music brands** — *Elsewhere* (elsewhere.club): full-bleed hero + **newsletter-first "Join the list"** above the fold, events tagged by type + physical zone, Memberships surfaced as revenue. *Teksupport/TCE* (tcepresents.com): minimal wordmark hero + live location/time stamp, **Day/Night mode toggle**, portfolio-of-brands, ticketing via Tixr/DICE. *Cityfox* (cityfox.us): sells **the experience/production values** (sound specs, light/art) over lineup, named membership tier, flagship recurring "experiences".

**D. Benefit / charity orgs** — *Dancers Against Cancer* (imadanceragainstcancer.org): rotating **"Hope Stories"** (real people) + donate, **real-time donation progress bars** (goal + %), past events kept live showing final raised, **tiered giving → outcomes** ("$50 covers co-pays"), "Donate Now" ×6. *Sweat with Pride* (sweatwithpride.com): **three live impact counters in the hero** (participants / $ raised / minutes), real-time donation feed + leaderboards, team ("bring a crew") fundraising. *(National MS Society / Bike MS pages 403 to fetch — their P2P playbook = DonorDrive team pages, fundraising minimums, finish-line progress, prize tiers.)*

**E. Boutique experiential** — *Daybreaker* (daybreaker.com): rotating tagline hero + "View reel", **scale stat band** ("150 events across the globe"), city-segmented SMS+email presale capture inline with listings, **mission woven into the main flow** (not siloed). *Mirrored Media* (mirroredmedia.com): video-background hero + one CTA, **value-pillar block** ("Magic is in our DNA"), case-study carousel, award badge.

## Cross-cutting patterns
1. **Experience-first hero** (full-bleed photo / looping reel / live player) — vibe sold before a word.
2. **One calendar, tagged/filtered by program type** (Public Records' Live/Club/Etc., Elsewhere's type+zone) → maps to Parties / Dance Infusion / Production.
3. **Newsletter/SMS capture above the fold + presale gating** (Elsewhere, Daybreaker, Teksupport).
4. **Brand site vs. ticketing-rail split** — owned site stays atmospheric; transactions on RA/Dice/Tixr.
5. **Portfolio-of-brands + tap-through case studies** (TCE, 23 Layers, Empire) — past work = clickable proof.
6. **Membership/community as a named tier** (Elsewhere, Cityfox, Daybreaker) beats a flat list.
7. **Transparent impact + progress mechanics** for charity (DAC bars/totals/tiers; Sweat with Pride counters/feed/leaderboards/teams).
8. **Values-led, low-friction copy** (Nowadays affordability line, Public Records mission-first).

## What this means for Come With
A **single parent brand with three legible lanes** (Parties · Dance Infusion · Production), an **experience-first hero**, **one filtered calendar**, ticketing delegated to RA/Partiful, and marquee nights packaged as **tap-through recap/case-study pages** that double as a production portfolio and social proof. For **Dance Infusion**: adopt the charity toolkit — progress bar (goal + %), persistent past totals, tiered outcome giving framed around MS, a donation feed/leaderboard, and team fundraising — with values-led copy so it reads as mission, not solicitation.

---

## How the 3 prototypes map to the research
- **V1 "Pulse"** (after-dark / nightlife) — experience-first dark hero, one tagged calendar, RA/Partiful CTAs, newsletter capture, marquee strip. For the **party-goer**. (Nowadays / Elsewhere / Lot Radio / Daybreaker DNA.)
- **V2 "Marquee"** (editorial / premium production studio) — light, whitespace, value-pillars, services-forward, result-framed work, "Request a proposal" CTA, capabilities lane. For **bookers / clients / sponsors**. (23 Layers / Eventique / Empire / Mirrored / TCE DNA.)
- **V3 "Infusion"** (community / mission) — warm, Dance-Infusion-forward, live impact counters + progress + tiered outcome giving + "Hope Story", merged tickets-or-donate, values-led copy. For the **community & donors**. (DAC / Sweat with Pride / Daybreaker / Nowadays Safer Space DNA.)


