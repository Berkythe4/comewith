# Shared SoundCloud song-fetch rule for the Elements tooling.
#
# ONE definition of "a song, not a DJ set", imported by both elements_tool.py
# (first build) and elements_rescan.py (re-pull). It used to be inline in
# elements_tool.py only, which is exactly how it drifted from sc-enrich's
# contract and put 251 multi-hour DJ sets into sc_artist_cache as "songs" on
# 2026-07-28. Keep in step with supabase/functions/sc-enrich/index.ts.
import json, time, urllib.request

SONG_MIN_MS = 45_000            # below this it's a clip / ID snippet
SONG_MAX_MS = 15 * 60 * 1000    # above this it's a DJ set / mix / livestream
UA = {"User-Agent": "Mozilla/5.0 Chrome/126"}


def fetch_songs(api, cid, uid, want=15, max_pages=6):
    """Return (songs, set_count) for a SoundCloud user id.

    Pages until `want` real SONGS are collected rather than `want` uploads: an
    artist who mostly posts sets previously came back mix-only or empty, because
    the item cap was filled by sets before any song was reached.
    """
    out, sets, pages = [], 0, 0
    url = f"{api}/users/{uid}/tracks?limit=50&client_id={cid}"
    while url and len(out) < want and pages < max_pages:
        pages += 1
        try:
            js = json.load(urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=15))
        except Exception:
            break
        for t in js.get("collection", []):
            if t.get("kind") != "track":
                continue
            d = t.get("duration") or 0
            if d > SONG_MAX_MS:
                sets += 1
                continue
            if d < SONG_MIN_MS:
                continue
            if t.get("streamable") is False:
                continue
            out.append({
                "sc_track_id": str(t["id"]), "title": t.get("title"),
                "permalink_url": t.get("permalink_url"), "duration_ms": d,
                "playback_count": t.get("playback_count") or 0,
                "created_at": t.get("created_at"), "artwork_url": t.get("artwork_url"),
            })
            if len(out) >= want:
                break
        nxt = js.get("next_href")
        url = (nxt + f"&client_id={cid}") if nxt and len(out) < want else None
        if url:
            time.sleep(0.15)
    return out, sets
