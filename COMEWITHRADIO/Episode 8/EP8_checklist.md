# Come With NYC Radio Ep 8 (SHOW 12): release checklist

**Guest DJ Kristen London.** Drop date: **Thu 2026-11-05** (dashboard). Mix length **1:09:14**. 17 tracks.
Runbook: `NOTES_WEEKLY_RELEASE.md`. Tracklist source: her brief, `Come With Radio.docx` (no Rekordbox export).

Status key: [x] done  [ ] to do  [?] waiting on Keith

## A. Tracklist onto the website (DONE 2026-10-30)
- [x] Brief -> `Track List Ep8 (Kristen London).txt` + `EP8_tracklist.txt` (Rekordbox-shaped, UTF-16) + `EP8_times.txt`
- [x] SHOW 12 synced to the 17 played tracks, in order (`apply_station_from_rekordbox.py`, pairs for #1/#2 so their SoundCloud links survive)
- [x] 50 unplayed picks logged passed@12 and carried to the next show. Backup: `backups/presync_2026-10-30_show12_tracks.json`
- [x] Show chips from `ra_events.lineup`, counted from the drop date. Aliases set by hand: #1 Alex Kassian (remixer, BASEMENT 11/6), #2 Locussolus = DJ Harvey (Knockdown 11/13), #17 Jennifer = Jen Cardini (Public Records 11/29)
- [x] #11 Best Mate (Spank Remix): NO NYC show in any feed under either name; card and page draw no show line
- [x] Her brief saved as the public description (`desc_public`, verbatim from the docx) + `desc_sc` + `recap_notes` + `mix_duration_ms`
- [x] Read back through `get-station` (token): 17 tracks, brief, "Mixed by" links to her profile
- [ ] Keith: eyeball SHOW 12 in the dashboard
- [?] Brief says "the **Come Here** prompt": typo for "Come With"? Left verbatim
- [ ] Genres / release dates: none on any of the 17 (no Rekordbox tags, no Beatport run). Cards render without the genre line. Optional: Beatport run to fill release_date

## B. Video
- [x] Cues `EP8_cues.csv` (all shows per card); preview cards checked (1, 5, 11, 17)
- [x] Render `CWR_Ep8_YouTube.mp4` with the SHOW 12 cover (`CWR_Ep8_cover.jpg`); next drop on the closing slide = 2026-11-19 (episodes are 2 weeks apart); intro reads GUEST MIX BY KRISTEN LONDON; closing slide adds "Follow Kristen London @kristenmlondon_dj"
- [x] `CWR_Ep8_mix.mp3` (320 kbps) for SoundCloud
- [?] Cover reads **EP.110626** (Fri 11/6) but the dashboard drop is **Thu 11/5**. Pick one

## C. Upload + schedule (Keith)
- [ ] YouTube: upload the MP4, scheduled premiere at the drop (text in `EP8_youtube.txt`)
- [ ] SoundCloud: upload the MP3 PRIVATE / scheduled, same time; cover art; description from `EP8_youtube.txt`
- [ ] Then: SC url + track id (`find_mix`, never paste), YouTube url, `scheduled_go_live` on SHOW 12
- [ ] Dry run `radio-publish-due?dry=1&station=12` before the drop (§97)
- [ ] Social: release post at/after go-live; tag @kristenmlondon_dj

## D. After it goes live
- [ ] Press play on both players on the episode page; read `station_notes` on SHOW 12
