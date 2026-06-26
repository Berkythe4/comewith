# Come With customer front-end — 3 prototypes, compared

Three complete, research-informed directions for the public site, built locally
(not on the live homepage). All three are single self-contained HTML files that
share one data layer (`_shared.js`): they pull **live upcoming events** from the
public Supabase view and **submit real inquiries** to the CRM. Open any file in a
browser to test.

| | **V1 · Pulse** | **V2 · Marquee** | **V3 · Infusion** |
|---|---|---|---|
| **File** | `v1-pulse.html` | `v2-marquee.html` | `v3-infusion.html` |
| **Direction** | After-dark / nightlife | Editorial / premium studio | Community / mission-first |
| **Primary audience** | Party-goers, fans | Bookers, clients, sponsors | Community, donors, MS supporters |
| **Mood** | Dark, electric, energetic | Light, calm, sophisticated | Warm, friendly, heartfelt |
| **Palette** | Plum/espresso bg + lime + hot-pink | Cream paper + espresso + plum/olive | Cream + plum + rose + gold |
| **Type system** | Archivo (heavy grotesque) + DM Sans/Mono | Fraunces (editorial serif) + DM Sans/Mono | Poppins (rounded) + DM Sans/Mono |
| **Hero** | Huge wordmark + next-show ticket card + marquee strip | "Events, produced / Parties, curated / Causes, served" + reel placeholder | Split: heart-led headline + **live impact card** (counters + progress bar) |
| **Primary CTA** | "See upcoming events" / "Get tickets ↗" | "Request a proposal" | "Get tickets" / "Support the cause" |
| **Events shown as** | Tagged cards, Upcoming + Past archive | Case-study cards (work) + clean upcoming list | Warm cards, Upcoming + Past, benefit-tagged |
| **Dance Infusion** | Bold section w/ stats + founder story | Editorial split w/ "% to mission" transparency | **Front and center** — story, tiers ($X=outcome), progress to goal, Safer-Space line |
| **Production / booking** | Secondary ("Book us") | **Lead offer** — 3 lanes, "how we work", proposal form | Secondary ("we also produce") |
| **Research DNA** | Nowadays · Elsewhere · Lot Radio · Daybreaker | 23 Layers · Eventique · Empire · Mirrored · TCE | Dancers Against Cancer · Sweat with Pride · Daybreaker |

## Strengths & trade-offs

**V1 · Pulse** — *Strengths:* instantly reads "cool party brand", strongest first-50ms vibe, the next-show ticket card converts ticket intent fast, fun motion (marquee). *Trade-offs:* dark/energetic can feel less "premium" to a corporate booker; the charity angle is present but not dominant.

**V2 · Marquee** — *Strengths:* most credible for paid production/booking & sponsors, "Request a proposal" repeated, work-as-portfolio builds trust, calm whitespace reads premium. *Trade-offs:* lower party energy; could feel agency-formal to the underground crowd; needs real photography to shine.

**V3 · Infusion** — *Strengths:* best for the mission — live impact counters + progress bar + tiered outcome giving + one human story are exactly the charity playbook; warmest, most inclusive tone (Safer-Space copy). *Trade-offs:* leads with the benefit, so the for-profit party/production side is demoted; risks reading as "a charity" rather than "a party brand that gives back" if not balanced.

## Scorecard (1–5, subjective, research-anchored)

| Criterion | V1 Pulse | V2 Marquee | V3 Infusion |
|---|---|---|---|
| Brand energy / "cool" | **5** | 3 | 4 |
| Premium / booker credibility | 3 | **5** | 3 |
| Mission / donor clarity | 3 | 3 | **5** |
| Drives **ticket** intent | **5** | 3 | 4 |
| Drives **booking/proposal** intent | 3 | **5** | 3 |
| First-impression impact (50ms) | **5** | 4 | 4 |
| Accessibility (contrast, labels, targets) | 4 | **5** | 4 |
| Mobile-first feel | 4 | 4 | **5** |
| Dev effort to ship | 4 | 4 | 4 |

## Recommendation
Come With's biggest public audience is **party-goers**, and the research is unanimous that nightlife sites win on **experience-first** energy — so **V1 "Pulse" is the strongest single default**. But the ideal production site is a **hybrid**:

- **Base = V1 Pulse** (energy, events-first, ticket CTA).
- **Fold in V3's Dance Infusion block** verbatim (live impact counters + progress-to-goal + tiered giving + the human story + Safer-Space line) — it's the proven charity toolkit and Come With's most distinctive asset.
- **Fold in V2's "Request a proposal" production lane + work-as-case-studies** as a clearly separate path for bookers/sponsors.

That gives the research's winning structure: **one parent brand, three legible lanes (Parties · Dance Infusion · Production), experience-first hero, one filtered calendar, ticketing delegated to RA/Partiful.** See `RESEARCH.md` for full sourcing.

## What's wired vs. mocked (for honest testing)
- **Live:** upcoming events (Supabase `v_public_events`), inquiry-form submit (→ CRM `inquiries`), newsletter capture (routed through inquiries as a tagged signup).
- **Mocked/placeholder:** hero/recap **imagery** (gradient placeholders — needs real event photos), past-event "case study" stats (curated in `_shared.js`), the Dance Infusion **$ goal/progress** (illustrative target), donation checkout (would link to a real donation rail), RA ticket **widget** (we link out per research; embedding the RA widget is a future option).

## Audit result
Structural + reference-integrity + syntax audit run on all three: **all pass** — valid `lang`/viewport/meta, sticky nav, every JS `#id` reference resolves, inputs labeled, tags balanced, no placeholder text, inline JS syntax-clean. One real bug (a malformed `\'` escape that broke the form/newsletter scripts) was caught and fixed during the audit.
