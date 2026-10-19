# Come With NYC Radio — weekly release notes

**Read this before touching a weekly episode.** Written 2026-07-30 after EP 2, where
most of these were learned the hard way. Ordered as the week actually runs.

---

## 0 · Where things live

Everything for a week goes in **`Radio/Week N/`** — the mix, artwork, tracklist,
cues, paperwork and the finished MP4. Not a shared `Radio/Video/` dump.

```
Radio/Week 2/
  CWR_Ep2.073026.WAV      the recorded mix
  CWR_Ep2.txt             Rekordbox export  <- THE TRACKLIST (see §2)
  CWR_EP2_artwork.png     this week's cover
  EP2_cues.csv            what the renderer reads
  EP2_times.txt           the fill-in sheet Keith/Martin types into
  EP2_youtube.txt         title + description + chapters
  EP2_buylist.txt         store links + prices
  EP2_checklist.md        release checklist
  CWR_Ep2_YouTube.mp4     the finished video
```

`make_episode.py --week 2` finds the mix and cues in that folder by itself.

---

## 1 · The header and the show name

The video header reads **`COME WITH NYC RADIO · EP N`**. Pass it explicitly:

```
--title "Come With NYC Radio"
```

The default is "Come With Radio", which is the *brand*; the show is **Come With NYC
Radio**. The Elements editions are a different show name again.

---

## 2 · The Rekordbox export is the tracklist, NOT the dashboard

The dashboard tracklist is what was *planned*. The deck export is what was *played*.
On EP 2 the dashboard had 19 tracks in one order; the export had 23 in a completely
different order, including four library tracks that were never logged.

**Always build the cues from the export**, then sync the dashboard to match:

```
python Radio/render/cues_from_rekordbox.py \
    --txt "Radio/Week 2/CWR_Ep2.txt" \
    --mix "Radio/Week 2/CWR_Ep2.073026.WAV" \
    --out "Radio/Week 2/EP2_cues.csv" \
    --times "Radio/Week 2/EP2_times.txt" --ep 2
```

Export format traps, all real:
- **UTF-16 with a BOM, tab separated.** Not UTF-8.
- **Locate columns BY HEADER NAME.** Rekordbox reorders them per view.
- **The Artist column is sometimes EMPTY**, with the artist folded into the title
  ("AC Slater - For The Funk"). Split on `" - "`.
- A **playlist/crate** export is not a **history** export. If the source lengths sum
  to roughly double the mix, some of those tracks were probably never played — ask
  before trusting it. EP 2: 130:40 of source against a 65:00 mix.

---

## 3 · Track start times

There is no reliable automatic route. In order of preference:

1. **Deck history export** (Rekordbox/Engine) — exact, one minute of effort. Week 1
   had `HISTORY 2026-07-22.m3u8`. Ask for it first, every week.
2. **The fill-in sheet** — `EP2_times.txt`, one line per track. Generate it, hand it
   over, let them type while watching a draft render.
3. **Tap-along** — `Radio/render/tap_times.html`. Works, costs the length of the mix.

**`match_mix.py` (audio fingerprinting) was tried on EP 2 and is not trustworthy.**
It returned confidences of 0.10–0.75, placed the tracks in nearly reverse order,
put two tracks 14s apart, and failed to find one at all. Every match reported a
−8%..+4% tempo stretch, i.e. it was straining to fit. Don't ship its output.

**Watch how the sheet comes back.** On EP 2 the times were typed over the `guess`
value, not after the `=`. Parse from wherever the numbers actually are; don't insist
on the template. Sanity checks that caught real problems:
- strictly increasing
- last start + last duration ≈ the mix length (EP 2 landed on 65:00 exactly)

---

## 4 · Every field the card draws

`render_card` reads these from the cues CSV. Miss one and it **silently** disappears
— that is how EP 2 first rendered with no genre line and no show chips:

| column | shows as |
|---|---|
| `artist`, `title` | the big lines |
| `genres` | `HOUSE · TECH HOUSE` |
| `release_date` | `— RELEASED 2021` |
| `show_date` | the lime date chip |
| `show_venue` | the venue chip |

Full header: `idx,start,artist,title,bpm,song_key,camelot,duration_ms,genres,release_date,show_date,show_venue,show_cost,show_url`

Build the cues **from the database** once it's synced, so the video and the website
can't disagree. Fill gaps at the source, not just in the CSV:
- **genre** — from the Rekordbox `Genre` column
- **show date/venue** — from `ra_artists.next_event_date/next_venue`. For a
  collaboration where one partner has no NYC date, use the other's booking.
- **release_date** — Beatport, via a pasted token (§7)

---

## 5 · Reveal beats — the stage cap

`INTRO_BEATS` / `OUTRO_BEATS` list one entry per reveal, and **the last entry is the
hold**. The stage number is capped at `len(BEATS) - 2`. It used to be hardcoded to
`min(i, 4)`, so adding a sixth beat rendered nothing and the new closing line was
silently missing from a finished 65-minute video.

`preview_bookends.py` keeps its **own** copy of that loop — fix both or the preview
disagrees with the render.

**Always verify a bookend by pulling a real frame**, never by reading the code:

```
ffmpeg -sseof -1.2 -i out.mp4 -frames:v 1 -y frame.png
```

Current closing tease: `TEASE_LINE` in `render_episode.py`, one line to edit.

---

## 6 · The SoundCloud link, and why the embed died on EP 1

**The link is never the problem — the track's PRIVACY is.** A permalink saved while
the track is private stays correct forever; what fails is `oembed`, which 404s on a
private track, so the site's player renders nothing. EP 1 shipped with a dead embed
for exactly this reason: link saved early, page auto-published on schedule, nobody
ever checked the embed.

`oembed` is the ground truth. A `200` on the track PAGE means nothing —
`soundcloud.com/comewithnyc/cwr_ep2-073026` answered 200 while its oembed 404'd.

**The scheduled path now handles it.** pg_cron calls the edge function
**`radio-publish-due`** (not the SQL function directly, which cannot make HTTP calls):

1. oembed the stored URL — if it answers, change nothing
2. if not, find the track on the account and flip `sharing=public`, then re-check
3. if there's no URL at all, find the upload by title/runtime and store it
4. publish either way — a SoundCloud problem must never hold the drop; the failure is
   written to `station_notes` so it is visible instead of silent

`cron.radio-publish-backstop` calls the SQL function on a slower cadence, so a broken
edge function delays the SoundCloud fix but never blocks the release.

**API trap:** `/resolve?url=` does **NOT** return a private track from its plain
permalink. Only `/me/tracks` lists an account's private uploads — match on permalink
there as the fallback. Env names are `SC_CLIENT_ID` / `SC_CLIENT_SECRET`.

## 6b · Retrieving a link by hand

**Standard as of 2026-07-30.** A private upload has no shareable URL, so asking for
one was always the wrong ask. `sc-connect` action **`find_mix`** reads
`/me/tracks`, which returns the authenticated account's own uploads *including
private ones with their secret_token*, and matches on the mix's exact runtime.

In the dashboard: 🚀 Go live → **☁ Find my upload**. Falls back to a candidate
picker when nothing matches on length. Pasting a link is still there, second.

- A **private** track is retrievable but **will not embed** — oembed 404s. Go live
  flips it public and refreshes the permalink.
- Store only canonical public URLs; validate through `resolve-media`. Short
  `on.soundcloud.com` links do not embed.

---

## 7 · Beatport tokens

Access tokens live **600 seconds from mint**, not from when they're copied. A token
read from `localStorage['token-refresh-result']` on a tab that's been sitting is
usually already dead — one arrived 173s past expiry on 2026-07-29. Take it off a
live request header instead (🌐 Find a live token, in the 🛒 Where to buy panel).

Fill `release_date` only where it is NULL; never overwrite. Two EP 2 tracks are
genuinely absent from Beatport (a DJ edit, a Bandcamp-only remix) — the card degrades
to genre only, which is fine. Beware near-misses: searching "Sweater (Swami Sound &
Age Remix)" returns `DaDa Sound Project — Tama Tama`, which would have stamped a
wrong year on the card. The remix/version guards exist for this.

---

## 8 · Artwork

**Square.** The renderer center-crops to a square, so a portrait 3:4 cover ends up
letterboxed with white bars inside its own frame (EP 2: 1151×1542). Ask for a square
export.

---

## 9 · Before handing anything over

- `ffprobe` the MP4: duration must equal the mix (bookends are *overlaid*, so the
  runtime never grows), 1920×1080, h264 + aac.
- Pull a **track-card frame** and a **final frame** and actually look at them.
- Name a draft so it cannot be uploaded by mistake (`EP2_DRAFT_do-not-upload.mp4`),
  and delete it once the real one exists.
- Flag anything with **uncleared samples** — EP 2 opens on an Afroman edit, which is
  a likely YouTube Content ID claim. Fine on SoundCloud.

---

## 10 · `weekly_prep.py` gotcha

It rewrites `EPN_cues.csv` every run. It used to write blanks first and read them back
two steps later, so re-running **wiped the times** and the chapters stayed `(time)`
forever. It now carries times forward, matching on artist+title and falling back to
**position** — needed because export titles ("Bam Bam (Original Mix)") differ from the
dashboard's cleaned ones ("Bam Bam"). Re-run it after the times land so the chapters
bake in, then check they aren't `(time)`.
