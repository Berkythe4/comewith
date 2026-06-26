# Come With — customer front-end prototypes (LOCAL — not deployed)

Three research-informed redesign directions for the public site, built locally so
nothing touches the live homepage (`/index.html`) until you pick a direction.

## How to view
Open any of these in a browser (double-click works — they're self-contained):
- **`v1-pulse.html`** — after-dark / nightlife direction (party-goers)
- **`v2-marquee.html`** — editorial / premium studio direction (bookers, sponsors)
- **`v3-infusion.html`** — community / mission-first direction (donors, MS supporters)

They pull **live upcoming events** and **submit real inquiries** to the Supabase
backend (read-only public endpoints + anon inquiry insert), so the events section
and the booking form actually work. If you prefer, run a tiny local server in this
folder (`python -m http.server`) and visit `http://localhost:8000/v1-pulse.html`.

## The docs
- **`RESEARCH.md`** — the full sourced dossier: (A) evidence-based design principles
  (NN/g, Baymard, web.dev, W3C), (B) event/ticketing UX with live teardowns,
  (C) 13 comparable companies analyzed, all with verified source URLs.
- **`COMPARISON.md`** — the three prototypes compared, scored, with a recommendation.
- **`CONTENT.md`** — real Come With brand facts/content used to populate the builds.
- **`_shared.js`** — the shared live-data layer (events fetch + inquiry submit).

## Status
Local prototypes only. Imagery is placeholder (gradients) pending real event photos;
ticketing links out to Resident Advisor per the research. Pick a direction (or the
recommended hybrid in COMPARISON.md) and it becomes the new `index.html`.
