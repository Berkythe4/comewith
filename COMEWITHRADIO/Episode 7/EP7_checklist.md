# Come With NYC Radio Ep 7 (SHOW 11): release checklist

**Emmy Adelle spotlight**: she plays Dance Infusion #3 (fundraiser), Sat 2026-11-14 @ Signal.
Drop date: **Thu 2026-10-22**. Mixed by **Berky**. Mix length **1:10:23**. 19 tracks.
Runbook: `NOTES_WEEKLY_RELEASE.md`. Tracklist source: `CWR_Ep7_V2.txt` (confirmed as played).
Note: the `Radio` folder is now `COMEWITHRADIO` (renamed 2026-10-19).

Status key: [x] done  [ ] to do  [?] waiting on Keith

## A. Tracklist onto the website (DONE 2026-10-18)
- [x] V2 confirmed as exactly what was played
- [x] Emmy's DI #3 show added by hand (`ra_events man:2026-11-14-emmy-adelle-signal` + `ra_artists man_emmy-adelle`)
- [x] SHOW 11 synced: 19 tracks in played order; 36 unplayed picks logged passed@11, carried to Ep 8
- [x] Emmy's SoundCloud + Instagram saved on her artist row (2026-10-19)
- [ ] Keith: eyeball SHOW 11 in the dashboard

## B. Times + cues (DONE 2026-10-19)
- [x] Mix audio in (`CWR_Ep2_V1.mp4`, 70:23) -> `CWR_Ep7_mix.mp3` (320 kbps, audio only)
- [x] Keith filled in `EP7_times.txt`
- [x] `EP7_cues.csv` (render/cues_with_shows.py): times + all shows + DI spotlight
- [x] Release dates: 17 from Keith's Beatport run, +1 from iTunes (Don't Kill My Vibe); Money (Know Good Flip) undated

## C. Cards + spotlight
- [x] Multiple shows per card; orange DI cards for Emmy's 4 tracks; real cover; DI thumbnail sticker
- [x] Closing slide = DI logo (flier dropped: misspells Emmy)  + "Catch Emmy Adelle live at our fundraiser" (`card_previews/closing_slide.png`)
- [x] Intro credit: "MIXED BY BERKY FT. EMMY ADELLE"
- [x] Flier NOT used anywhere (misspells Emmy as "Adele")
- [?] Website: show the extra shows + a DI tag on the episode page too (needs a site change; not built)

## D. Video
- [x] Render `CWR_Ep7_YouTube_v2.mp4` (1080p h264, AAC 320k, 70:23)
- [x] ffprobe + real frames checked: intro, card 5 (two shows), card 13 (DI), closing DI-logo slide
- [x] `EP7_youtube.txt`: title, YouTube + SoundCloud descriptions, chapters, tags (DRAFT)

## E. Upload + schedule
- [x] Keith uploaded: YouTube csjUEHmjyq8 = scheduled premiere Thu 10/22 3:00 PM ET (title, desc, chapters, thumbnail verified)
- [x] SoundCloud track 2420632227 found via /me/tracks (length matches to the ms; desc + tags + art verified)
- [x] SoundCloud set private + scheduled 10/22 3 PM by Keith (oembed 404 confirmed)
- [x] Fixed: SC field held the YouTube link; now SC url + track id + YouTube url + duration stored
- [x] SC + public descriptions saved on SHOW 11 (desc_sc incl. tracklist + hashtags)
- [x] Scheduled go-live 2026-10-22 19:00 UTC (3 PM ET, = YouTube premiere); dry run clean ("already embeddable")
- [x] Social: Emmy spotlight caption Wed->Thu 10/22; release brief says post at/after 3 PM ET; recap drafted; flier-spelling note on Oct 25 post
- [?] Keith: move the release post from 11:04 AM to 3 PM ET+ in the dashboard (connector can't change times)
- [x] Wording -> "100% of your $25 ticket goes to the National MS Society" (site + SoundCloud done; YouTube = Keith)

## F. After it goes live (10/22)
- [ ] Press play on both players on the episode page; check `station_notes` on SHOW 11
- [ ] Beatport cart for anything unpurchased (/beatport-cart 11), if wanted
