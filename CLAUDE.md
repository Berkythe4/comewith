# Come With — Project Conventions

Operational conventions for this repo. These are binding — follow them exactly.
Broader migration history and architecture live in `ROADMAP.md`.

## Start of session (read these, in this order)

Keith works from **two machines** (desktop + laptop). Claude Code's memory is
per-machine and does **not** sync, so everything a session needs to resume lives
**in the repo**:

1. **`CARRYOVER.md`** — where the last session left off, and what's next. Start here.
2. **`DEV_DOCS/claude-memory/MEMORY.md`** — index of the desktop's Claude memory,
   snapshotted into the repo so either machine can read it. Open individual files
   from there as needed; they're background, and reflect what was true when written.
3. **`LEARNINGS.md`** — numbered, append-only design decisions with rationale.
4. **`ROADMAP.md`** — architecture + phase history (older; CARRYOVER wins on state).
5. **`reviews/session_YYYY-MM-DD.md`** — the narrative of a given session.

Then `git log --oneline -15` and `git branch -a` — work is sometimes parked on a
branch specifically so it does **not** auto-deploy (see below).

## End of session (the close routine)

**`SESSION_CLOSE_PROMPTS.md` is the ritual** — Quick / Standard / Full by session
size. Run it at the end of any session that shipped real work; always when a
migration or prod data was touched, or a new standing rule was set. It carries the
prod safety checks (the five financial views must return anon **401**) and it's
what keeps CARRYOVER / LEARNINGS / ROADMAP true instead of drifting.

Two Come-With-specific additions to whichever variant you run:
- **Re-snapshot Claude memory** into `DEV_DOCS/claude-memory/` (see the README
  there) so the other machine isn't more than one session behind.
- **`master` auto-deploys to Netlify.** Pushing `master` publishes the site and
  dashboard immediately. Work Keith hasn't green-lit goes on a branch and is
  named in CARRYOVER under "Parked / next"; never merge it to close a session out.
  Edge functions are separate — they go live only on an explicit deploy.

## Database / Supabase migrations

- **Project:** prod is `yaytdosxfhcqatmhctzk` (`comewith-prod`). The CLI is linked
  to staging; do not assume the link points at prod. Migrations live in
  `supabase/migrations/NNN_name.sql`, numbered in sequence (… 015–019 and up).
- **Introspect before apply.** Confirm live columns / view definitions / policies
  against prod, reconcile any `[VERIFY]` refs, and show the SQL/diff for review
  before applying anything to prod.
- **Roles:** `master_admin` / `sub_admin` / `customer`. There is **no `admin`
  role.** RLS uses the helper `public.is_admin()` (= `role in ('master_admin',
  'sub_admin')`). New admin-only tables: `for all using (public.is_admin())`.
- **NEVER use a blanket `grant ... to anon` in a migration.** Specifically, do not
  write `grant all on all tables in schema public to anon` (or to `authenticated`).
  `013_grants.sql`'s `ALTER DEFAULT PRIVILEGES` already grants the right
  privileges to new tables automatically. A broad grant silently re-grants SELECT
  on **all views too**, re-exposing financial views that were deliberately revoked
  from `anon` — this caused the **016/017 regression that 019 had to fix**. If you
  ever must re-grant, immediately re-assert every prior `revoke … from anon`, and
  verify anon access in the post-apply check (financial views must return 401).
- **Financial views are anon-revoked by design** (decision E1): `v_event_summary`,
  `v_kpi_event_financials`, `v_kpi_parties`, `v_kpi_dance_infusion`,
  `v_kpi_dashboard`. Keep them revoked. Verify with an anon REST GET → expect 401.
- **Apply discipline:** apply additively, verify on prod (objects, RLS has a real
  policy — never RLS-enabled-with-no-policy, admin can read/write, anon blocked),
  then commit the migration file so tracked history matches prod.
- **`INSERT..RETURNING` enforces the SELECT policy mid-statement.** A security-definer
  helper that re-queries the table cannot see the not-yet-visible new row, so
  `.insert().select()` fails RLS even for the creator (bit us on 097 chat DMs).
  Put row-local predicates like `created_by = auth.uid()` directly in the SELECT
  policy. RLS can be smoke-tested on prod via the Management API:
  `set_config('request.jwt.claims', …)` + `set local role authenticated` inside
  BEGIN..ROLLBACK.
- **Deactivation contract (098):** `profiles.deleted_at` set = user deactivated;
  `is_admin()` / `is_master_admin()` / `user_can_access_module()` all treat that
  profile as no-role. Any new role helper MUST keep the `deleted_at is null` guard.

## Series contract (events.series)

`events.series` is free text. KPI views match it **exactly**. The Log Event form
MUST write `series = 'Come With Parties'` for parties and `series = 'Dance Infusion'`
for DI events, or those KPIs read empty. `'Come With Production'` is services
(we run someone else's production), not parties. `'Bookings'` (type `gig`,
added in 095) is when we're the **booked talent** at someone else's event —
performance fees go there, never under Production. The host/client who booked
us goes in `events.owner_actor_id` ("Host / booked by" in the edit-event modal).

## Mailing segments (brand delineation)

Two-level segments on `subscriber_segments`, established 2026-07-13:
- **Brand rollups** (what campaigns target): `come_with`, `dance_infusion`.
  A subscriber can hold both. Unsubscribe stays **global** (one master list).
- **Per-event segments** (cohort history): the event slug or event code,
  e.g. `come-with-7-11`, `di-02-2026-05`.

Every event import MUST add BOTH the event segment AND the matching brand
segment. Public signup widgets pass the brand segment (`come_with` on the
homepage; DI pages must pass `dance_infusion`). Never re-subscribe an
unsubscribed email during an import (e.g. `chaddercheesy@gmail.com`).

## Come With Radio (episodes live outside `events`)

- **`station_no` is the SHOW counter; `edition_seq` is the episode number.** Two
  different numbers — do not render either as "EP n" generically. `station_no`
  counts **every broadcast ever**, across series, and is displayed as **`SHOW n`**
  everywhere (dashboard, radio.html, dj.html, index.html, social-post titles,
  and the DB strings reworded in 137). `edition_seq` is a series' own numbering
  (Elements Ep1–4) and is what an audience knows — so the **rendered video keeps
  "EP n" and uses `edition_seq`** (`make_episode.py`). Elements Ep1 is SHOW 3.
  `station_no` is assigned at CREATION, not airtime, so it can drift out of
  broadcast order — `scripts/renumber_shows.py` fixes that (dry by default; never
  moves a published episode; also remaps `played_station_no` / `passed_station_no`
  / `carried_from`, which store the NUMBER rather than a key).
- **Radio episodes are numbered stations in `sc_playlists`** (`station_no`,
  lifecycle `building → testing → live → archived`), **NOT rows in `events`**.
  Do not create an event for an episode. The scheduled release date lives on
  `sc_playlists.drop_date` (radio's own tracker); the site teases the next drop
  via `get-station ?list=1 → next_drop`. Only one `building` row exists at a time
  (partial unique index) — auto-created with the next number when all are live.
- **Song memory `sc_song_log`** is the permanent played/passed/carried record —
  finalize logs played, sync/remove logs passed, finalize carries passed-not-
  played songs into the next station. Keep it in sync when touching station tracks.
- **Listener accounts** are `customer`-role auth users; `listener_*` tables are
  owner-RLS'd + anon-revoked. Never grant anon on them. `sc_playlists` /
  `sc_playlist_tracks` were also anon-revoked in **103** — they had carried
  table-level anon grants since 079 (RLS was blocking the rows, so an anon GET
  returned `200 []`, never data; now it's `401`). Public station reads are
  function-only through `get-station` (service role).
- **Rekordbox is the arrangement tool, not SoundCloud** (decided 2026-07-22).
  The set is bought and arranged in Rekordbox because SoundCloud isn't
  record-quality. The ① test push to SoundCloud + ↺ sync-back still exist for the
  first pass, but the **Rekordbox import owns final order**: dashboard
  "🎛 Import Rekordbox order" parses the playlist export (UTF-16 tab TSV, columns
  located BY HEADER NAME; also .m3u8/CSV/pasted lists), fuzzy-matches to the
  station, applies the order and pulls BPM/key. A station therefore holds songs
  that never came from SoundCloud — see `source` in migration 102 and the
  synthetic `man_…` `sc_track_id`.
- **Store metadata never overwrites Rekordbox.** `track-sources` only FILLS IN a
  missing bpm/song_key/camelot. Your own analysis of the file you own beats a
  store's tags. Matching must keep the **remix guard**: if either side names a
  remix/edit, the remixer has to match too, or the original mix matches
  "(X Remix)" and you buy the wrong track.
- **Beatport = PASTE-A-TOKEN, not a stored credential** (settled 2026-07-22 after
  testing against the live API). Hard facts, verified — don't re-litigate:
  - Access tokens live **600 seconds** (`exp - iat` on a real token). Ten minutes.
  - The **refresh token is unreachable from the browser**: not in localStorage
    (which holds only `token-refresh-result` → `{accessToken,…}`), and their site
    refreshes via a cookie JS can't read. Watching the Network tab for 30 min
    produced nothing; background tabs throttle the auto-refresh timer anyway.
  - `client_id` is in the **JWT payload** (`client_id` claim) — no hunting needed.
    `BEATPORT_CLIENT_ID` is set as a project secret.
  - So: a bookmarklet copies the current token from beatport.com, the "🛒 Where to
    buy" modal takes a paste, and `track-sources` caches it in
    `public.beatport_oauth` **only until its own JWT `exp`**. No standing
    credential at rest. Never `site_content` (anon-readable).
  - API shape confirmed: `tracks[]`, `key.camelot_number`/`camelot_letter`,
    `release.label.name`, `slug`+`id` → `beatport.com/track/<slug>/<id>`, and
    **`price.value` is in DOLLARS (1.49) not cents** — do not divide by 100.
  - Search "artist title" then **retry on title alone whenever nothing clears the
    match threshold** (not only on zero results): "Deeper Purpose Cigarettes"
    returns three confident-looking wrong hits, so a zero-results guard never
    fires and the real release is never searched for.
  `/beatport-cart` remains the way to actually fill a cart.
- **Bandcamp has no official API.** `track-sources` uses the endpoint their own
  search box calls: `POST bandcamp.com/api/bcsearch_public_api/1/autocomplete_elastic`
  with `{search_text, search_filter:"t", full_page:false, fan_id:null}`. The older
  `fuzzysearch/1/autocomplete_elastic` path is DEAD — and it answers **HTTP 200**
  with `{"error":true,"error_message":"bad function"}`, so checking `r.ok` alone
  silently reported every track as "not on Bandcamp". **Validate the payload, not
  the status**, and throw so the caller can say "couldn't reach Bandcamp" — never
  let an outage render as a definitive "not available".
- **Store matching is adversarial** — Bandcamp is full of DJ rips, bootlegs and
  flips of the track you actually want. Three guards, all regression-tested:
  (1) remix words are detected **anywhere**, not just in brackets ("Artist. Title.
  Pat Lok Flip." has none); (2) `(Radio Edit)`/`(Extended Mix)` are standard
  qualifiers, NOT remixes, and must still match their own release; (3) substring
  containment is **length-aware** — a flat score let "If U Need It" match
  "Sammy Virji: If U Need It (Callto Speed Garage Dub)". Returning "not found"
  beats sending Keith to buy the wrong file.
- **The public page never links the source playlist.** `get-station` deliberately
  does not select `sc_playlist_url`. Listeners get the FINAL MIX only
  (`mix_sc_track_url` / `mix_youtube_url`); to get the songs they come to the
  episode page and export the tracklist. Per-track links are fine. Don't
  "helpfully" re-add a playlist link.
- **Phase 1.1/2 pending:** YouTube auto-post at finalize; listener "export my
  saved playlist to my own SoundCloud" (OAuth per listener — designed, not built).

## Weekly radio release (see `Radio/NOTES_WEEKLY_RELEASE.md`)

- **A private SoundCloud track will NOT embed.** `oembed` 404s on it, so the site's
  player renders nothing — and a `200` on the track PAGE proves nothing. EP 1 shipped
  with a dead embed this way. The link being saved early is fine; the track must be
  **public at publish time**.
- **The scheduled release goes through the `radio-publish-due` EDGE FUNCTION**, not
  the SQL function directly — SQL can't make HTTP calls, so it could never check the
  embed. It oembeds, flips `sharing=public` if needed, then publishes. A SoundCloud
  failure NEVER blocks the drop; it's written to `station_notes`. `cron
  radio-publish-backstop` still calls the SQL path so a broken function can't hold a
  release.
- **`/resolve?url=` does not return a PRIVATE track** from its plain permalink. Only
  `/me/tracks` lists an account's own private uploads. Env: `SC_CLIENT_ID` /
  `SC_CLIENT_SECRET`.
- **Never ask for a SoundCloud link** — retrieve it (`sc-connect` action `find_mix`,
  or ☁ Find my upload). A private upload has no shareable URL to give.
- **The Rekordbox export is the tracklist, not the dashboard.** The dashboard holds
  what was planned; the export holds what was played (EP 2: 19 planned vs 23 played,
  different order). Build cues from the export, then sync the dashboard. Export is
  UTF-16, tab separated, columns located BY HEADER NAME, and the Artist column is
  sometimes empty with the artist folded into the title.
- **Rekordbox owns bpm/key.** Store metadata only fills what's missing.
- **A missing cues column fails silently.** `render_card` needs `genres`,
  `release_date`, `show_date`, `show_venue` or that part of the card just isn't drawn.
- **Verify a render by pulling a real frame** (`ffmpeg -sseof -1.2 … -frames:v 1`),
  never by reading the code — a hardcoded stage cap silently dropped a new closing
  line from a finished 65-minute video.
- **Beatport access tokens live 600s FROM MINT**, not from copy; take one off a live
  request header, not `localStorage`. The **cart API works** — see the APIs map for
  the item shape.
- Episode material lives in **`Radio/Week N/`**, MP4 included. Show name in the video
  header is **Come With NYC Radio**.

## Media / recap links (must be publicly embeddable)

- **Validate every recap/media URL through `resolve-media` before storing.**
  SoundCloud share short-links (`on.soundcloud.com/…`) are redirects the embed
  player can't follow, and private/secret/wrong tracks oembed-404 — both fail
  **silently** on the site. The event editor already does this on save (resolves
  short links, strips `utm_*`/`si`, verifies oembed). Store only canonical,
  public URLs. `mediaKind()` matching "soundcloud.com" is NOT proof it embeds.
- **CSS: never use the `background:` shorthand on a variant/state class** (e.g.
  `.benefit`, `.audio`) layered over a base that set `background-size/position/
  repeat` — the shorthand resets them and breaks hero photos. Use `background-image:`.

## Editing `dashboard.html` (1.3 MB, one inline `<script type="module">`)

- **Never read the whole file into context.** It is ~1.3 MB and effectively all one
  inline module. Reading it once costs more than an entire ordinary session, and it
  gets read on nearly every dashboard PR. Locate the region with `grep -n`, pull only
  that window with `sed -n 'A,Bp'`, then `Edit` on an exact unique string. To review a
  change, `git diff` it — never `cat` the file.
- **Syntax-check by extraction, not by re-reading.** Extract the inline module body to
  a temp `.mjs` and run `node --check` on it. Run the **same extraction against the
  pre-edit version** (`git show HEAD:dashboard.html`) as a control — otherwise an
  artifact of the extraction itself reads as a real error introduced by the edit.
- **There is no local console check** — the Browser pane can't open `file://` URLs. The
  loop is `node --check` plus the deployed Netlify build.

## Scope

- This codebase is **Come With only**. Do **not** add anything Come With Fitness
  (CWF) anywhere — not in the dashboard, schema, or pages.
