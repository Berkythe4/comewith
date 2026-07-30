# Shared SoundCloud song-fetch rule for the Elements tooling.
#
# ONE definition of "a real song by THIS artist", imported by elements_tool.py
# (first build) and elements_rescan.py (re-pull). It lived inline in
# elements_tool.py only, which is how it drifted from sc-enrich's contract and
# put 251 multi-hour DJ sets into sc_artist_cache as "songs" on 2026-07-28.
# Keep in step with supabase/functions/sc-enrich/index.ts.
#
# Three rules, each earned from a real miss:
#  1. LENGTH   45s <= duration <= 15min. Longer = DJ set / mix / livestream.
#  2. OWNERSHIP a track must belong to this profile. Albums and playlists on a
#     profile can hold anyone's music, so the merge below checks t.user.id.
#  3. CREDIT   even a track uploaded BY this profile can be someone else's
#     release. "LEVEL UP - The Other Side" sits on soundcloud.com/zingaraa, so
#     it was served as a Zingara song (found 2026-07-30). SoundCloud carries the
#     rights credit in publisher_metadata.artist, which said "LEVEL UP" — that is
#     authoritative, so a track whose credit omits this artist is dropped.
#     Collaborations keep working: "Zingara, LEVEL UP" still names Zingara.
import json, re, time, urllib.request

SONG_MIN_MS = 45_000            # below this it's a clip / ID snippet
SONG_MAX_MS = 15 * 60 * 1000    # above this it's a DJ set / mix / livestream
UA = {"User-Agent": "Mozilla/5.0 Chrome/126"}
# Containers are merged because /users/{id}/tracks does NOT return everything:
# LEVEL UP shows track_count=46 but serves 0 from /tracks — their catalogue is
# all inside albums/playlists. sc-enrich hit the same wall (commit f81873a).
CONTAINERS = ["albums", "playlists", "playlists_without_albums", "spotlight"]


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def clout(track):
    """Engagement, for choosing between duplicate uploads of the same song.

    Likes + reposts + comments, because those are deliberate acts; plays are the
    tiebreak only. Keith asked for "the song with the most clout on SoundCloud".
    """
    g = lambda k: int(track.get(k) or 0)
    return (g("likes_count") + g("reposts_count") + g("comment_count"),
            g("playback_count"))


def dupe_key(title, artist_names=()):
    """Identity of a SONG, so two uploads of it collapse but two VERSIONS do not.

    Artists re-post their own material: Big Gigantic have "Hot Mic" twice (54,409
    plays and 2,158), and Boys Noize have 'Boys Noize "Shut It Down" ft. TAICHU,
    Taube' alongside a plain "Shut It Down". So the key ignores the artist's own
    name, surrounding quotes, a "(feat. X)" group and the no-op "(Original Mix)".

    It deliberately KEEPS version qualifiers — "(Extended VIP Mix)" is a different
    record, not a duplicate, and must survive as its own entry.
    """
    t = (title or "").lower()
    for n in artist_names:
        n = (n or "").strip().lower()
        if len(n) >= 3:
            t = t.replace(n, " ")
    t = re.sub(r"[\"“”'’]", " ", t)
    # bracketed feature group first, so a trailing "[Extended Mix]" is not eaten
    t = re.sub(r"[\(\[][^)\]]*\b(?:feat|ft|featuring)\b[^)\]]*[\)\]]", " ", t)
    # then a bare trailing "ft. A, B" with no bracket left after it
    t = re.sub(r"\b(?:feat|ft)\.?\s+[^\[\(]*$", " ", t)
    t = re.sub(r"[\(\[]\s*original mix\s*[\)\]]", " ", t)
    return _norm(t)


def credited_elsewhere(track, artist_names):
    """Return the crediting artist string if this track is someone ELSE's release.

    None means "keep it" — either the artist is credited, or there is no evidence
    either way. We only ever drop on POSITIVE contrary evidence, because a missing
    or sloppy credit must not cost an artist their own song.
    """
    pm = ((track.get("publisher_metadata") or {}).get("artist") or "").strip()
    if not pm:
        return None                                   # no rights data -> keep
    names = [n for n in (_norm(x) for x in artist_names) if n]
    if not names:
        return None
    # Compare against the WHOLE normalised credit and by TOKEN overlap. Never split
    # the credit on separators: splitting on "&" tore "Above & Beyond" into
    # ["above","beyond"], matched neither, and dropped that artist's entire official
    # catalogue (2026-07-30) — only their untagged uploads survived. Every check
    # below fails toward KEEPING, because wrongly dropping costs an artist their own
    # song while wrongly keeping merely leaves a track to eyeball.
    cred = _norm(pm)
    #  a) whole name inside the credit — "abovebeyond" in "abovebeyond, zojohnston"
    if any(n in cred for n in names):
        return None
    #  b) credit inside the name — artist "Flash Gea" registered simply as "Flash"
    if any(cred and cred in n for n in names):
        return None
    #  c) a distinctive word in common — "MIND | MATTER" vs the credit "MIND I MATTER"
    #     normalise to mindmatter vs mindimatter, which neither (a) nor (b) catches.
    tokens = {t for raw in artist_names for t in re.findall(r"[a-z0-9]{4,}", (raw or "").lower())}
    if any(t in cred for t in tokens):
        return None
    #  d) no distinctive token to reason about (e.g. "LA sad", whose releases are
    #     credited to its members) — that is not evidence of a foreign track.
    if not tokens:
        return None
    #  e) the title itself naming them (a remix or flip of theirs).
    if any(n in _norm(track.get("title")) for n in names):
        return None
    return pm


def fetch_songs(api, cid, uid, artist_names=(), want=15, max_pages=6):
    """Return (songs, set_count, dropped) for a SoundCloud user id.

    `artist_names` = every name this profile is known by (SoundCloud username and
    the festival lineup name); used for the CREDIT rule.
    `dropped` = [(title, crediting_artist)] so callers can report what was skipped.

    Pages until `want` real SONGS are collected rather than `want` uploads: an
    artist who mostly posts sets used to come back mix-only or empty, because the
    item cap filled with sets before any song was reached.
    """
    def get(url):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=20) as r:
                return json.load(r)
        except Exception:
            return None

    # `best` keys on the SONG, not the upload, so a re-post loses to whichever copy
    # has the most engagement. Insertion order (SoundCloud's own, newest first) is
    # preserved via `order` so a dedupe never silently reshuffles the crate.
    out, sets, dropped, seen = [], 0, [], set()
    best, order, dupes = {}, [], []

    def consider(t):
        nonlocal sets
        if t.get("kind") != "track" or not t.get("id"):
            return
        tid = str(t["id"])
        if tid in seen:
            return
        if str((t.get("user") or {}).get("id") or uid) != str(uid):
            return                                    # OWNERSHIP: not this profile's
        d = t.get("duration") or 0
        if d > SONG_MAX_MS:
            seen.add(tid); sets += 1; return           # LENGTH: a DJ set
        if d < SONG_MIN_MS or t.get("streamable") is False:
            return
        who = credited_elsewhere(t, artist_names)
        if who:
            seen.add(tid); dropped.append((t.get("title"), who)); return   # CREDIT
        seen.add(tid)
        row = {
            "sc_track_id": tid, "title": t.get("title"),
            "permalink_url": t.get("permalink_url"), "duration_ms": d,
            "playback_count": t.get("playback_count") or 0,
            "likes_count": t.get("likes_count") or 0,
            "reposts_count": t.get("reposts_count") or 0,
            "comment_count": t.get("comment_count") or 0,
            "created_at": t.get("created_at"), "artwork_url": t.get("artwork_url"),
            "credited_artist": ((t.get("publisher_metadata") or {}).get("artist") or None),
        }
        k = dupe_key(t.get("title"), artist_names)
        prev = best.get(k)
        if prev is None:
            best[k] = row; order.append(k)
        elif clout(t) > clout(prev):
            dupes.append((prev["title"], row["title"]))      # loser, winner
            best[k] = row
        else:
            dupes.append((row["title"], prev["title"]))

    url, pages = f"{api}/users/{uid}/tracks?limit=50&client_id={cid}", 0
    while url and len(best) < want and pages < max_pages:
        pages += 1
        js = get(url)
        if not js:
            break
        for t in js.get("collection", []):
            consider(t)          # no early break: a later upload may out-clout an earlier one
        nxt = js.get("next_href")
        url = (nxt + f"&client_id={cid}") if nxt and len(best) < want else None
        if url:
            time.sleep(0.15)

    # Still short? Their catalogue may live in albums/playlists instead.
    if len(best) < want:
        for path in CONTAINERS:
            if len(best) >= want:
                break
            js = get(f"{api}/users/{uid}/{path}?limit=50&client_id={cid}")
            for c in (js.get("collection", []) if isinstance(js, dict) else []) or []:
                for t in (c.get("tracks") or []):
                    consider(t)
                if len(best) >= want:
                    break
            time.sleep(0.15)

    out = [best[k] for k in order][:want]
    return out, sets, dropped, dupes
