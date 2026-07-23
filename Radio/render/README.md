# Come With Radio — video render + weekly prep

Self-contained tools that turn a recorded mix into the **Option 4 "Now Playing"**
YouTube video, and assemble the paperwork for a release. All local, all
read-only against prod except where noted. Nothing here changes the live site.

Requires: Python 3 + Pillow (installed), ffmpeg (installed), and `.env` with
`SBP_PAT` + `SBP_REF_PROD` for the DB pulls.

---

## The weekly one-liner

```
python Radio/render/weekly_prep.py            # current working station
python Radio/render/weekly_prep.py --station 1
```

READ-ONLY. Writes four files and prints a status line:

| File | What |
|---|---|
| `Radio/render/EP{N}_cues.csv` | tracklist — fill the `start` column, then render |
| `Radio/Documents/EP{N}_youtube.txt` | YouTube title + description + chapters |
| `Radio/Documents/EP{N}_buylist.txt` | every track + Beatport/Bandcamp link + price |
| `Radio/Documents/EP{N}_checklist.md` | release checklist, pre-ticked for what's done |

Run it anytime — start of the week to see what's left, again after buying/timing
to refresh. It never writes to the DB or the site.

---

## Making the video — the one-liner

```
python Radio/render/make_episode.py --audio "Radio/Video/EP1_mix.wav"
# with deck history:  --history "Radio/Video/EP1_history.txt"
```

`make_episode.py` does the whole thing for the working station: pulls the cues,
takes the times (from `--history`, or the cues if already filled, or stops and
tells you to tap them in), and renders `Radio/Video/EP{N}.mp4`. If you'd rather
run the pieces yourself, they're below.

## Making the video — the pieces

**1 · Get the cues (tracklist).** `weekly_prep.py` already wrote
`Radio/render/EP{N}_cues.csv`. (Or run `make_cues.py` on its own.)

**2 · Add the track start times.** Three ways, easiest first:

- **Rekordbox / Engine history** — if your gear logged the session (standalone
  writes it to the USB; laptop keeps it in the app), export it and send it over —
  the times get merged into the cues for you.
- **Tap-along** — open `Radio/render/tap_times.html` in a browser (double-click,
  no server). Load the mix audio + the cues CSV, play it, tap SPACE as each track
  drops, hit **Download cues CSV**. That file has the times filled in.
- **By hand** — type `mm:ss` into the `start` column.

The renderer refuses to run if any `start` is blank, so you can't ship a
half-timed video by accident.

**3 · Render.**

```
python Radio/render/render_episode.py \
    --cues Radio/render/EP1_cues.csv \
    --audio "Radio/Video/EP1_mix.wav" \
    --cover Radio/Artwork/Radio_Thumbnail.jpg \
    --out Radio/Video/EP1.mp4 \
    --ep "EP 1"
```

Out comes a 1920×1080 H.264 MP4 in `Radio/Video/`. Upload it to YouTube, paste
the link into the dashboard (✎ Details → YouTube link).

Flags: `--dry` renders the cards + a 1-second preview (fast sanity check);
`--title` overrides the show name.

### What it looks like
Per track: your cover, artist (big), title, BPM + Camelot/key chips, "up next",
and a lime progress bar that fills across the whole mix. No camera, no editor, no
ticket pitch — just the music on screen. Brand pulled from `radio.html`; type is
Arial Black + Consolas (stand-ins for Archivo / DM Mono).

---

## How the render works (so future-you can fix it)

`render_episode.py` → Pillow draws one PNG card per track → an ffmpeg `concat`
demuxer holds each card for its track's duration → a lime bar source is
alpha-gated by time with `geq` (`x < BAR_W·t/total`) and `overlay`-ed on top →
audio muxed. The progress-bar geometry lives in one place (`BAR_X/Y/W/H`) shared
between the PNG rail and the ffmpeg fill — edit there only.

> Gotcha we already hit: ffmpeg `drawbox`'s `t` is **thickness, not time**, so it
> can't animate a bar. The `geq`-alpha + `overlay` route is deliberate. And `geq`
> needs r/g/b **and** alpha or it errors.

---

## Files
- `make_episode.py` — one command: cues → times → render (the one-liner above)
- `weekly_prep.py` — the read-only paperwork assembler
- `make_cues.py` — just the cues CSV
- `import_history.py` — read start times from a Rekordbox/Engine history export
  (preview-first; never writes without `--write`)
- `render_episode.py` — the video renderer
- `tap_times.html` — offline tap-along timestamp tool
- `EP{N}_cues.csv` — generated; your working file (git-ignored)

## If something goes wrong — revert

Everything here is **new, additive files** — no live app code, no DB, no
functions were touched by the render tools. To undo the whole render toolkit:
`git revert <commit>` (or delete `Radio/render/`). Nothing on the site or in the
dashboard depends on it. The generated `EP*` files are regenerable anytime with
`weekly_prep.py`, so deleting them is harmless.
