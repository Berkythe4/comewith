# Come With Radio

Working files for the Come With Radio show. This is the local, in-repo home for
art and documents — it is **not** what the website serves. See the note below.

## Folders

- **Artwork/** — station artwork and per-episode covers (the source images).
  `Radio_Thumbnail.jpg` is the current station thumbnail. Name episode covers
  clearly, e.g. `EP1_cover.jpg`.
- **Video/** — finished YouTube mix videos (`EP1.mp4`) and any raw ingredients.
- **Documents/** — anything else radio-related: notes, guest-mix agreements,
  release checklists, drop schedule, etc.

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
