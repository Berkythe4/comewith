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
- **Week N/** — **one folder per episode.** Everything for that episode lives
  here: the recorded `mix.wav`, the Rekordbox History export, `tracklist.json`
  (the locked timestamps), `youtube_chapters.txt`, the working `EP*` files, and
  the finished `CWR_EpN_YouTube.mp4`. Heavy audio/video is git-ignored; the small
  records (tracklist, chapters, cues, docs) are kept.

## The weekly one-liners

```
python Radio/render/weekly_prep.py --week 1     # cues + chapters + buy list + checklist → Week 1/
python Radio/render/make_episode.py --week 1    # mix + times → Radio/Week 1/CWR_Ep1_YouTube.mp4
```

`make_episode --week N` auto-finds the mix, the times, and the tracklist in
`Radio/Week N/`, and renders the video there. No camera, no editor.

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
