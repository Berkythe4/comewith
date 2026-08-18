# Gear Watch — stolen-rig resale scan

Built 2026-08-18 after the vehicle break-in (NYPD grand larceny). Scans the
resale market three times a day for the stolen DJ rig, scores every candidate,
and alerts on the ones worth a human look.

**The rule that matters most:** a high-scoring hit goes to the **detective**,
with its score breakdown. Never contact the seller, never arrange a meet.

## Pieces

| Piece | Where |
|---|---|
| Schema, cron, module row | `supabase/migrations/146_gear_watch.sql` |
| Scanner | `supabase/functions/scan-gear-market/index.ts` |
| Confidence model | `supabase/functions/scan-gear-market/scoring.ts` |
| Tests (20, no network) | `node --test supabase/functions/scan-gear-market/scoring.test.ts` |
| Dashboard panel | `dashboard.html` → `loadGearWatch()`, `panel-gearwatch` |
| Loss schedule (the claim) | `Financial/ComeWith_Stolen_Gear_Loss.xlsx` |

## Sources, and why only three

| Site | How | Note |
|---|---|---|
| Reverb | Official API, `REVERB_TOKEN` | **Live-verified 2026-08-18.** Search + a per-candidate detail lookup |
| eBay | Browse API, `EBAY_CLIENT_ID` + `EBAY_CLIENT_SECRET` | Client-credentials OAuth, minted per run. **Never run live yet.** |
| Craigslist | `sapi.craigslist.org` — the internal JSON endpoint its own search box calls | **Live-verified 2026-08-18.** No credentials needed |
| OfferUp | — | No public API, bot-blocked |
| Facebook Marketplace | — | No public API, scraping prohibited |

The last two are emitted as **saved-search links** in the digest and in the
panel, so they stay a 60-second manual check instead of a scraper that breaks
silently.

**Craigslist's search RSS is dead** — every `format=rss` path answers HTTP 403, and the
HTML search page is a JS shell with no listings in it. The working route is the endpoint
the front-end itself calls, which is the same technique `track-sources` uses for Bandcamp.
Its payload is delta-encoded; the decode and its traps are documented in `parsers.ts` and
pinned by tests. See LEARNINGS §27.

**You can exercise the whole pipeline right now, with no deploy and no keys:**

```
node scripts/gearwatch_live_test.mjs
THEFT_DATE=2026-08-16 node scripts/gearwatch_live_test.mjs "pioneer cdj 3000"
```

It hits Craigslist live, runs the real decoder and the real scorer, and prints each
listing's score breakdown plus every drop and why. This is the fastest way to check the
scan still works after any change to the model.

It also writes **`gearwatch_results.html`** in the repo root — the same results as a
clickable page, so links can actually be opened (and it's gitignored, since it holds
listing data and is regenerated every run). **Until migration 146 is applied and the
function deployed, that file is the only place results exist** — nothing is being stored,
and no scan is running on a schedule.

**A source that fails is reported as FAILED, never as "nothing found."** An
outage that renders as "your gear isn't listed anywhere" is the most dangerous
bug this system could have — same lesson as the Bandcamp `bcsearch` endpoint
that answered HTTP 200 with an error body.

## The score

Gates first — a gate means "this is not our gear", so the listing is **dropped,
not stored with a low score**:

1. the model has to be named in the listing
2. accessory listings are out (`case`, `skin`, `decal`, `stand`, `bag` in the
   **title** — checked on the title only, because half of the legitimate
   listings say "comes with case" in the description)
3. anything posted **before the theft date** is out

Then weighted signals, capped at 100:

| Signal | Points |
|---|---|
| Model match (base) | 35 |
| **Serial number in the listing** | **+60** |
| Two different stolen models in one listing (bundle) | +30 |
| Local (NYC metro) | +25 |
| Listed after the theft — **graduated** | +25 (≤3 days), +20 (≤7), +15 (≤14), +10 (≤30), +5 (≤60) |
| A pair, of a model we lost a pair of | +18 |
| Price far below market (<60%) | +15 (or +8 under 80%) |
| Seller feedback under 10 | +10 |
| "no box" / "cash only" / "must go" language | +5 each, capped at 10 |

**Calibrated 2026-08-18 against real listings**, after the first live run scored a
CDJ-3000 in Brooklyn posted the day after the theft at only 55. The bands now mean:

| Score | Reading | What happens |
|---|---|---|
| 100 | Serial match, or the full rig in one listing | push + email |
| 85–99 | Model + local + fresh — the shape of a real hit | push + email |
| 65–84 | Model + one strong signal | email |
| 35–64 | Model only, or distant, or stale | logged in the panel, no alert |

Recency is graduated on purpose: stolen gear is flipped within days, so "listed
tomorrow" and "listed in three weeks" are not the same claim. A flat bonus threw that
away. Note the recency signal only works once the **theft date** is set.

Every award is stored in `score_breakdown`, because "87/100" is not something
you can hand a detective and "serial match +60, local +20" is.

**Recovering serials is worth more than any other input.** The serial rule is
dormant until `gear_watch_targets.serial` is filled in — paste them into the
panel as they arrive.

## Install (in order)

1. **Apply the migration** to prod (`yaytdosxfhcqatmhctzk`) — introspect first,
   per the repo rule:
   ```
   SBP_REF=$SBP_REF_PROD python db.py supabase/checks/pre_apply.sql   # targets: gear_watch_*
   SBP_REF=$SBP_REF_PROD python db.py supabase/migrations/146_gear_watch.sql
   ```
2. **Get the API credentials** (~15 min, both free):
   - **Reverb** → sign in → user menu → **My Profile** → **API & Integrations** tab →
     **Generate New Token**. Name it (e.g. "Gear Watch"), and tick the **`public`** and
     **`read_listings`** scopes — that is all a search needs; do not grant write scopes.
     **Reverb personal tokens do not expire**, so this is set-once.
     Test it before it goes anywhere near prod:
     ```
     REVERB_TOKEN=<token> node scripts/gearwatch_live_test.mjs "pioneer cdj 3000"
     ```
     A working token prints `reverb: N listing(s)`; a bad one prints
     `reverb FAILED: HTTP 401 — token rejected`.
   - **eBay** → developer.ebay.com → create an app → production App ID (client id) +
     Cert ID (client secret)
3. **Set the function secrets** (Supabase dashboard → Edge Functions → Secrets):
   `REVERB_TOKEN`, `EBAY_CLIENT_ID`, `EBAY_CLIENT_SECRET`
   (`RESEND_API_KEY` and the VAPID keys are already set — the digest goes
   through `send-notice`, the push through `send-push`.)
4. **Deploy the function** — the CLI can't do this, use the script:
   ```
   python scripts/deploy_edge_function.py scan-gear-market
   ```
5. **Store the two vault secrets once** so pg_cron can call the function.
   pg_cron cannot mint an admin JWT (the reason scheduled sends were deferred in
   `014_cron.sql`), so the job uses a service-role bearer read from vault at call
   time — it is never in git:
   ```sql
   select vault.create_secret('<service-role key>', 'gear_watch_srk');
   select vault.create_secret('https://yaytdosxfhcqatmhctzk.supabase.co/functions/v1/scan-gear-market', 'gear_watch_url');
   ```
   Until both exist the cron job is a **documented no-op** — it writes
   `skipped: secrets not set` to `gear_watch_config.last_status` rather than
   failing silently every eight hours.
6. **Open Gear Watch in the dashboard** (Operations → Gear Watch, master only)
   and set, under ⚙ Settings: **theft date** (required — it's the hard gate),
   digest email, and thresholds. Set `gear_watch_config.push_user_id` to Keith's
   auth user id for the phone alert on high hits.
7. **Run it by hand once** ("🔍 Run scan now") and confirm all three sources
   report `ok — N listing(s) fetched` rather than FAILED.

## Schedule

`0 12 * * *`, `0 18 * * *`, `0 0 * * *` UTC = **8am / 2pm / 8pm ET**. Spread
across the day on purpose: listings get posted and pulled within hours, so a
single nightly scan would miss a same-day flip entirely.

## Reverb needs two calls, not one

**Reverb's search payload has no seller location and no seller feedback.** Verified
against the live API: `location` is absent and `shop` holds only `{slug,
preferred_seller}`. Scoring off search alone left every Reverb hit location-unknown —
the largest source in the scan, structurally blind to the strongest non-serial signal
there is.

The **listing detail** endpoint (`/api/listings/<id>`) carries both:

```
location: {region, locality, country_code, display_location}
shop:     {feedback_count, rating_percentage, preferred_seller}
```

So the scan does search → model gate → **detail lookup for the survivors only** (a
handful, not the ~50 results per query) → score. The lookup is capped at 40 per run and
the cap is reported in the source line when it bites, never silently.

Two signals cut the other way, both added after seeing real Reverb inventory:
**priced above 1.5× market −10** (a fence prices to move; 3× market is dealer retail or a
bundle) and **established dealer −10** (Reverb's preferred-seller badge — a shop with a
storefront and a return policy is the least likely route for stolen goods). Both lower a
score; neither gates a listing out.

## Known soft spots

- **The Craigslist endpoint is undocumented** and can change shape without notice. The
  decoder validates the payload and throws — so a shape change reads as FAILED, never as
  "nothing listed". Re-check with `node scripts/gearwatch_live_test.mjs`.
- **`CRAIGSLIST_NYC_AREA = 17`** is the area id in the `batch` parameter. If results stop
  looking like NYC, that is the number to check.
- **Craigslist returns nearby-area results too** (upstate, New Jersey, Connecticut). That
  is deliberate — stolen gear travels — and the geo score, not the query, decides local.
- **eBay `itemCreationDate` is not always present** in Browse search summaries.
  When a listing has no date, the pre-theft gate can't apply — the listing is
  kept rather than guessed at.
- **The cases and stands are not scan targets.** They carry no serial, resell
  for little, and "CDJ case" matches thousands of listings. They stay on the
  loss schedule; they are not worth hunting.
- Reverb is queried once per target (4 calls/run). Well inside any sane rate
  limit, but if targets grow, batch them.
