# Come With Radio

Working files for the Come With Radio show. This is the local, in-repo home for
art and documents — it is **not** what the website serves. See the note below.

A visual of the whole weekly process lives at **`/radio-workflow.html`**
(comewith.org/radio-workflow.html) and is linked from the dashboard's Artist
Radio panel.

## Folders

- **Artwork/** — the show's artwork (`Radio_Thumbnail.jpg`). Station-level, not
  per-episode.
- **render/** — the tools (Python scripts + the tap tool). Never edited
  per-episode. See `render/README.md`.
- **Episode N/** — **one folder per episode.** Everything for that episode lives
  here: the recorded `mix.wav`, the Rekordbox History export, `tracklist.json`
  (the locked timestamps), `youtube_chapters.txt`, the working `EP*` files, and
  the finished `CWR_EpN_YouTube.mp4`. Heavy audio/video is git-ignored; the small
  records (tracklist, chapters, cues, docs) are kept.

## Making the video — just double-click it

**`Make Radio MP4.bat`**, in the Comewith folder. It asks which episode, then
does the whole thing: pulls the tracklist, refuses to go on if any start time is
blank, renders, ffprobes the result against the mix, and grabs three frames for
you to look at. No terminal, no flags, nothing to remember.

Put these in `Radio/Episode N/` first:

| | |
|---|---|
| the mix | `CWR_EpN.<date>.wav` — the recorded set |
| the Rekordbox **history** export | a `.txt` — this is the tracklist, *not* the dashboard |
| the artwork | square PNG/JPG with `artwork` in the name (optional — falls back to the station art) |

Then double-click, and type the episode number.

> **Type the number the AUDIENCE knows.** Ep 3, not show 7. They are two
> different numbers — see below.

If the tracklist or the start times aren't there yet, it stops and prints
exactly what to run or go and get. It never renders a half-timed video.

### Episode number vs show number

`station_no` is the global show counter; the episode number is that show's place
in **its own** series. The Elements run took shows 3–6, so **NYC Radio Ep 3 is
SHOW 7**. `make_episode.py` resolves one to the other against the database — you
only ever type the episode number. Burning "EP 7" onto a video the world knows
as Ep 3 would be wrong, and so would pulling Elements Ep 1's credits into Ep 3's
intro slide; both were possible before the resolver existed.

### If you'd rather run the pieces

```
python Radio/render/weekly_prep.py --episode 3      # chapters + buy list + checklist
python Radio/render/make_episode.py --episode 3     # cues -> times -> render
python Radio/render/verify_episode.py --episode 3   # ffprobe + frames
```

`make_episode.py --episode N` finds the mix, the history and the cues in
`Radio/Episode N/` by itself, and writes the video there. `--dry` renders the cards
plus a one-second preview if you just want to eyeball the design.

## Important: this folder does NOT feed the live site

The website and dashboard read images and mixes from **Supabase storage**, not
from this folder:

- **Station artwork** — upload it in the dashboard: Radio → open the station →
  **✎ Details** → "📻 Come With Radio — station artwork". That stores it in the
  `radio-mixes` bucket under `brand/` and records the URL in
  `site_content.ops.radio_artwork`. Keeping the source here is just a backup /
  working copy.
- **Episode cover** — same modal, "This episode's cover" (goes to `radio-mixes`
  under `covers/`, saved on `sc_playlists.cover_url`).
- **The mix audio** — uploaded during the 🚀 Go live step (`radio-mixes` bucket).

So: keep your **originals** here for safekeeping and editing; **upload** the ones
you want public through the dashboard.

## Big files

Large `.mp4` mixes and hi-res masters can bloat the git repo. If a file is big
(roughly >25 MB), consider keeping it out of git — tell Claude and it'll add a
gitignore rule so the folder stays organized without committing the heavy file.
