# Come With NYC Radio Ep 8 (SHOW 12): release checklist

**Guest DJ Kristen London.** Drop date: **Thu 2026-11-05** (dashboard). Mix length **1:09:14**. 17 tracks.
Runbook: `NOTES_WEEKLY_RELEASE.md`. Tracklist source: her brief, `Come With Radio.docx` (no Rekordbox export).

Status key: [x] done  [ ] to do  [?] waiting on Keith

## A. Tracklist onto the website (DONE 2026-10-30)
- [x] Brief -> `Track List Ep8 (Kristen London).txt` + `EP8_tracklist.txt` (Rekordbox-shaped, UTF-16) + `EP8_times.txt`
- [x] SHOW 12 synced to the 17 played tracks, in order (`apply_station_from_rekordbox.py`, pairs for #1/#2 so their SoundCloud links survive)
- [x] 50 unplayed picks logged passed@12 and carried to the next show. Backup: `backups/presync_2026-10-30_show12_tracks.json`
- [x] Show chips from `ra_events.lineup`, counted from the drop date. Aliases set by hand: #1 Alex Kassian (remixer, BASEMENT 11/6), #2 Locussolus = DJ Harvey (Knockdown 11/13), #17 Jennifer = Jen Cardini (Public Records 11/29)
- [x] #11 Best Mate (Spank Remix): Best Mate is a duo that includes **Will Automagic** (2026-11-02). Alias set by hand like #1/#2/#17: chip = Brooklyn Pacha 11/7 (Louie Vega, ra.co/events/2557249); card also names BASEMENT 11/21 (ra.co/events/2553762, same night as #13 Cormac). Added to the "Catch them live" line in `desc_public` + `EP8_youtube.txt`; cues/shows regenerated, MP4 re-rendered
- [x] Her brief saved as the public description (`desc_public`, verbatim from the docx) + `desc_sc` + `recap_notes` + `mix_duration_ms`
- [x] Read back through `get-station` (token): 17 tracks, brief, "Mixed by" links to her profile
- [ ] Keith: eyeball SHOW 12 in the dashboard
- [?] Brief says "the **Come Here** prompt": typo for "Come With"? Left verbatim
- [ ] Genres / release dates: none on any of the 17 (no Rekordbox tags, no Beatport run). Cards render without the genre line. Optional: Beatport run to fill release_date

## B. Video
- [x] Cues `EP8_cues.csv` (all shows per card); preview cards checked (1, 5, 11, 17)
- [x] Render `CWR_Ep8_YouTube.mp4` with the SHOW 12 cover (`CWR_Ep8_cover.jpg`); next drop on the closing slide = 2026-11-19 (episodes are 2 weeks apart); intro reads GUEST MIX BY KRISTEN LONDON; closing slide adds "Follow Kristen London @kristenmlondon_dj"
- [x] `CWR_Ep8_mix.mp3` (320 kbps) for SoundCloud
- [x] Cover reads **EP.110626** (Fri 11/6); drop stays **Thu 11/5**. Keith 2026-11-02: 11/5 is correct, cover art stays as is

## C. Upload + schedule (DONE 2026-11-03)
- [x] YouTube: Keith uploaded iaPdU8cj3x8 = scheduled premiere Thu 11/5 3:00 PM ET (20:00 UTC; read off the watch page: isUpcoming, startTimestamp)
- [x] SoundCloud: Keith uploaded track 2429271267 PRIVATE + scheduled; found via `find_mix` (/me/tracks), length matches to the ms (4,154,514)
- [x] SHOW 12: SC url (Keith, = the found track's permalink) + track id + YouTube url (cleaned of `?si=`) + `scheduled_go_live` 2026-11-05 20:00 UTC
- [x] Dry run `radio-publish-due?dry=1&station=12`: due 1, no warning; "sharing=private, left alone (SoundCloud owns the release)". The page waits for SoundCloud's own release (up to 120 min, then publishes anyway with a `station_notes` warning). SC's scheduled time should be the same 3 PM ET
- [ ] Social: release post at/after go-live; tag @kristenmlondon_dj

## D. After it goes live
- [ ] Press play on both players on the episode page; read `station_notes` on SHOW 12
