# How to make the episode MP4

> **The visual version of this page — the whole run as one flow chart — is at**
> https://claude.ai/code/artifact/677fbf74-c3b6-4258-abff-683b462a2e69
> Start there. This page is the same thing in writing, and works offline.

Everything you need is one double-click. This sheet is what to put in the folder
first, what the tool asks you, and what to do when it complains.

> The deep background — why the renderer does what it does, the SoundCloud and
> Beatport traps, the release checklist — is in `NOTES_WEEKLY_RELEASE.md`. You
> don't need it to make a video.

---

## 1 · Put three things in `Radio\Episode N\`

| | | |
|---|---|---|
| **the mix** | `CWR_Show7_ep.3.wav` | any `.wav` / `.mp3` / `.m4a`. Name it however you like. |
| **the tracklist** | `Track List Show 7 radio ep.3.txt` | typed by hand. Must start with `Track List`. |
| **the artwork** | `CWR_EP3_artwork.png` | optional. Needs `artwork` in the name, and must be **square**. |

Nothing else is required. There is no Rekordbox step any more.

### The tracklist file

A plain text file, one line per track, **in the order you played them**:

```
1 ALL THE TIME - John Summit - 0:00
2 Pop Pop - Channel Tres - 1:39
3 Evergreen Kings - Adriatic - 3:47
```

Number, then the track, then the start time. That's it.

It's forgiving about the rest — these all read fine:

```
11 Twisted - Budino 22:23          no dash before the time
24 On Again - Mau P -47:02         no space either
9 I Got 5 On It - LNZ (Franky Rizardo Remix) - 18:23
```

It doesn't care which side of the dash the artist is on, or whether you wrote
the remixer the same way the database did.

**This file is the only place two things exist**: the order the set was actually
played in, and where each track starts. Everything else — SoundCloud links,
artwork, track lengths, each artist's next NYC show — comes out of the database
automatically. Don't retype any of that.

### If the artwork is portrait

Get a square export. The renderer center-crops to a square, so a 3:4 cover ends
up with white bars inside its own frame on every card.

---

## 2 · Double-click `Make Radio MP4.bat`

It lives in the `Radio` folder. (It also works from the `Comewith` folder — it finds the repo root either way.)

It asks you three things:

**"Episode number:"** → the number the **audience** knows. Ep 3, not show 7.

> Those really are two different numbers. The Elements run took shows 3–6, so
> NYC Radio **Ep 3 is SHOW 7**. The tool converts one to the other itself — you
> only ever type the episode number. Typing 7 would get you Elements Ep 4.

**"Next drop date:"** → when the *next* episode goes out, as `2026-09-10`. This
is the date on the closing slide, beside "WE PLUG BACK IN THURSDAY". Press Enter
to accept whatever the dashboard has scheduled, but it's a plan until you
confirm it — and it gets burned into the video.

**"Update the WEBSITE's tracklist order too? (y/N)"** → say yes once you're happy
with the running order. The dashboard shows what was *planned*; nothing corrects
it to what was *played* unless you do this. It only ever renumbers the existing
tracks — it never deletes anything, so the SoundCloud links and saved playlists
are safe.

Then it shows you the matched tracklist and waits for a **Y** before rendering
anything. An hour-long mix takes about 15 minutes.

---

## 3 · Look at the frames

When it finishes it opens `Radio\Episode N\_preview\` with three stills: the intro,
a track card, and the last frame.

**Actually look at them.** Every automatic check can pass on a video that drew
the wrong words — that has happened here twice, once with a closing line that
silently didn't render into a finished 65-minute video. `ffprobe` cannot see
text.

Then upload to YouTube and paste the link into the dashboard: **✎ Details →
YouTube link**.

---

## When it stops and tells you something

| It says | What to do |
|---|---|
| `There is no folder "Radio\Episode N"` | Make the folder, put the mix and tracklist in. |
| `No mix audio` | Drop the recorded set into the episode folder. |
| `No typed tracklist` | Write the `.txt` described above. |
| `typed track(s) matched NOTHING` | That line doesn't exist in the dashboard. Add the track there, or fix the spelling in the txt. |
| `dashboard track(s) were NOT played` | The dashboard has a song you didn't play. Remove it in the dashboard. |
| `time goes backwards` | A start time is earlier than the one before it — a typo in the txt. |
| `Verification FAILED` | Don't upload. The line above says which check failed. |

It refuses to write the website's order while anything is unmatched, on purpose.

---

## Doing it by hand

The `.bat` is only a wrapper. The same three steps:

```
python Radio/render/tracklist_from_txt.py --episode 3                 # check the match
python Radio/render/tracklist_from_txt.py --episode 3 --out-cues      # write the cues
python Radio/render/tracklist_from_txt.py --episode 3 --write-order   # fix the site order

python Radio/render/make_episode.py --episode 3 --next-date 2026-09-10
python Radio/render/verify_episode.py --episode 3
```

`make_episode.py --dry` renders the cards plus a one-second preview, if you only
want to check the design.

The paperwork — YouTube title, description, chapters, buy list, release
checklist — is separate and read-only:

```
python Radio/render/weekly_prep.py --episode 3
```

---

## Changing how the card looks

All in `Radio/render/render_episode.py`:

| What | Where |
|---|---|
| the comewith.org/radio band | `BAND_URL`, `BAND_LEAD`, `BAND_ITEMS` |
| the intro / closing copy | `EDITIONS["weekly"]` |
| the closing tease line | `TEASE_LINE` |
| brand colours | `LIME`, `CREAM`, `DIM`, `FAINT` |
| progress bar geometry | `BAR_X`, `BAR_Y`, `BAR_W`, `BAR_H` — **shared with ffmpeg, edit here only** |

Two traps worth knowing before you edit:

- **Symbols.** The card draws from one font with no emoji fallback. `▸ ♥ ↓ · —`
  work. `▶ ♡ ⬇ 🎟 ★` render as empty boxes. Always render a card and look before
  trusting a new symbol.
- **Reveal beats.** `INTRO_BEATS` / `OUTRO_BEATS` — the last entry is the hold,
  and the stage number is capped at `len(BEATS) - 2`. `preview_bookends.py` keeps
  its own copy of that loop, so fix both or the preview lies to you.

Preview a card change without a full render:

```
python Radio/render/make_episode.py --episode 3 --dry
python Radio/render/preview_bookends.py
```
