#!/usr/bin/env python3
"""
release_dates.py — fill the `release_date` column of a cues CSV.

    python Radio/render/release_dates.py --cues <cues.csv>            # dry, prints
    python Radio/render/release_dates.py --cues <cues.csv> --write

WHY NOT track-sources: that edge function is the right tool when the songs are
rows on a station, and its best source (Beatport) needs a pasted token that
lives 600 seconds. This works off a cues CSV and only uses APIs that need no
auth at all, so it can run unattended:

    iTunes  → releaseDate, excellent dance/electronic coverage
    Deezer  → album release_date, good on club/extended mixes iTunes misses
    MusicBrainz → first-release-date, the long tail

MATCHING IS ADVERSARIAL, same lesson as the store matching in track-sources:
a search for "Chris Lorenzo Appetite" will happily return someone else's
"Appetite", and an original mix will match a remix of itself. So:

  * artist must overlap on a distinctive token;
  * VERSION WORDS MUST AGREE — if either side says remix/flip/edit/bootleg, the
    other has to as well, and a named remixer has to appear on both. Otherwise
    you date the original when you played the remix, or vice versa;
  * "(Extended Mix)" / "(Original Mix)" / "(Radio Edit)" are standard qualifiers,
    NOT remixes — they must not block a match to the same release;
  * the EARLIEST credible date wins, because a track gets re-released on
    compilations for years and the compilation date is not the release date.

Returning nothing beats returning a confident wrong year on a card.
"""
import argparse, csv, io, json, os, re, sys, time, urllib.parse, urllib.request

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

UA = {"User-Agent": "ComeWithRadio/1.0 (berky@comewith.org)"}

# Words that mean "this is a different record from the original".
REMIX_W = r"(remix|rmx|bootleg|flip|edit|mashup|refix|rework|vip|dub|remaster)"
# Words that are just a version of the SAME record — never a mismatch signal.
QUALIFIER = re.compile(r"\((extended|original|radio|club|instrumental|clean|dirty|"
                       r"extended mix|original mix|radio edit|club mix)[^)]*\)", re.I)


def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def tokens(s):
    return {t for t in norm(s).split() if len(t) >= 4}


def who_tokens(s):
    """Remixer names, at a LOWER length threshold than artist tokens.

    A remixer is often three characters — YDG, GRiZ, MPH — and the 4-char artist
    threshold silently dropped them, leaving an empty remixer set that made the
    "same remixer?" guard skip itself entirely."""
    return {t for t in norm(s).split() if len(t) >= 2}


def version_of(title):
    """(is_remix, remixer_tokens) for a title, ignoring standard qualifiers."""
    t = QUALIFIER.sub(" ", title or "")
    is_rm = bool(re.search(REMIX_W, t, re.I))
    who = set()
    if is_rm:
        for m in re.finditer(r"[\(\[]([^)\]]*)[\)\]]", t):
            inner = m.group(1)
            if re.search(REMIX_W, inner, re.I):
                who |= who_tokens(re.sub(REMIX_W, " ", inner, flags=re.I))
        # "Artist. Title. Pat Lok Flip." — the remix credit needn't be bracketed
        m = re.search(r"([^()\[\]]+)\s+" + REMIX_W + r"\b", t, re.I)
        if m and not who:
            who |= who_tokens(m.group(1))
    return is_rm, who


def base_title(title):
    """Title without any bracketed group — what the record is actually called."""
    return norm(re.sub(r"[\(\[][^)\]]*[\)\]]", " ", title or ""))


def acceptable(want_artist, want_title, got_artist, got_title):
    """Would a careful person say these are the same recording?"""
    wt, gt = base_title(want_title), base_title(got_title)
    if not wt or not gt:
        return False
    # length-aware containment: "If U Need It" must not swallow a much longer title
    if wt != gt:
        short, long_ = (wt, gt) if len(wt) <= len(gt) else (gt, wt)
        if short not in long_ or len(short) < 0.6 * len(long_):
            return False
    # Artist must overlap on a distinctive token — but a short name has none.
    # tokens() drops anything under 4 characters, so YDG, MPH and Hol! produced an
    # EMPTY set and could never match themselves, no matter how exact the title.
    # Fall back to comparing the whole normalised name in that case.
    if not (tokens(want_artist) & tokens(got_artist)):
        nw, ng = norm(want_artist).replace(" ", ""), norm(got_artist).replace(" ", "")
        if not (nw and ng and (nw in ng or ng in nw)):
            return False
    w_rm, w_who = version_of(want_title)
    g_rm, g_who = version_of(got_title)
    if w_rm != g_rm:
        return False                       # original vs remix — different records
    # Both are remixes. If EITHER names a remixer they must share one. Requiring
    # both to name someone let a bare "(remix)" match any named remix: MusicBrainz
    # has "Speaker Knockerz, Young Scooter - Lonely (remix)" from 2014, which was
    # accepted as the 9B49 remix and would have printed RELEASED 2014 on that card.
    if w_rm and (w_who or g_who) and not (w_who & g_who):
        return False
    return True


def get(url, tries=2):
    for _ in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=20) as r:
                return json.load(r)
        except Exception:
            time.sleep(0.6)
    return None


# Rekordbox file-naming leftovers. They are part of the FILENAME, not the record,
# and they poison a search ("Girl$ (YDG Remix)FINAL" finds nothing anywhere).
LEFTOVER = re.compile(r"[\s_-]*(final|finalv?\d*|v\d+|master|wav|mp3|clean|dirty)\s*$", re.I)


def clean_title(title):
    t = (title or "").strip()
    for _ in range(3):                     # "(...)FINAL_v1" needs more than one pass
        n = LEFTOVER.sub("", t).strip()
        if n == t:
            break
        t = n
    return re.sub(r"\s{2,}", " ", t)


def split_lead_artist(title):
    """'Dom Dolla - Girl$ (YDG Remix)' -> ('Dom Dolla', 'Girl$ (YDG Remix)').

    Rekordbox users routinely put the ORIGINAL artist in front of the title when
    the file is a remix or edit, while the Artist column holds the remixer. The
    record is filed under the original artist everywhere, so searching for the
    remixer's name alone finds nothing.
    """
    m = re.match(r"^\s*([^()\[\]]{2,40}?)\s+-\s+(.+)$", title or "")
    return (m.group(1).strip(), m.group(2).strip()) if m else (None, None)


def queries(artist, title):
    """Search strings to try, best first. The title-alone retry is not a
    zero-results fallback — it runs whenever nothing has CLEARED the bar, because
    a search can return three confident-looking wrong hits and never trip a
    zero-results guard (the 'Deeper Purpose Cigarettes' lesson)."""
    ct = clean_title(title)
    out = ["%s %s" % (artist, ct), ct]
    lead, tail = split_lead_artist(ct)
    if lead:
        out += ["%s %s" % (lead, tail), tail]
    seen, uniq = set(), []
    for q in out:
        k = norm(q)
        if k and k not in seen:
            seen.add(k); uniq.append(q)
    return uniq


def _match_all(artist, title, got_pairs):
    """got_pairs = [(got_artist, got_title, date_or_payload)] -> accepted payloads."""
    ok = []
    lead, tail = split_lead_artist(clean_title(title))
    for ga, gt, d in got_pairs:
        if not d or (isinstance(d, tuple) and not any(d)):
            continue
        if acceptable(artist, clean_title(title), ga, gt):
            ok.append(d)
        elif lead and acceptable(lead, tail, ga, gt):
            ok.append(d)                   # matched as the ORIGINAL artist's record
    return ok


def from_itunes(artist, title):
    out = []
    for q in queries(artist, title):
        js = get("https://itunes.apple.com/search?term=%s&entity=song&limit=15"
                 % urllib.parse.quote(q))
        pairs = [(r.get("artistName", ""), r.get("trackName", ""),
                  ((r.get("releaseDate") or "")[:10], r.get("primaryGenreName") or ""))
                 for r in (js or {}).get("results", [])]
        out = [(d, g, "iTunes") for d, g in _match_all(artist, title, pairs) if d]
        if out:
            break
        time.sleep(0.2)
    return out


def from_deezer(artist, title):
    for q in queries(artist, title):
        js = get("https://api.deezer.com/search?q=%s&limit=15" % urllib.parse.quote(q))
        cand = []
        for r in (js or {}).get("data", []):
            cand.append(((r.get("artist") or {}).get("name", ""), r.get("title", ""),
                         (r.get("album") or {}).get("id")))
        keep = [c for c in cand if _match_all(artist, title, [(c[0], c[1], "x")])]
        out = []
        for ga, gt, alb in keep[:4]:
            if not alb:
                continue
            a = get("https://api.deezer.com/album/%s" % alb)
            d = (a or {}).get("release_date", "")[:10]
            g = ((a or {}).get("genres") or {}).get("data") or []
            gname = g[0].get("name", "") if g else ""
            if d:
                out.append((d, gname, "Deezer"))
            time.sleep(0.15)
        if out:
            return out
        time.sleep(0.2)
    return []


def from_musicbrainz(artist, title):
    for q in queries(artist, title):
        js = get("https://musicbrainz.org/ws/2/recording?query=%s&fmt=json&limit=12"
                 % urllib.parse.quote(q))
        pairs = []
        for r in (js or {}).get("recordings", []):
            ga = ", ".join(c.get("name", "") for c in r.get("artist-credit", []) if isinstance(c, dict))
            pairs.append((ga, r.get("title", ""), ((r.get("first-release-date") or "")[:10], "")))
        out = [(d, g, "MusicBrainz") for d, g in _match_all(artist, title, pairs) if d]
        time.sleep(1.1)                    # MusicBrainz asks for 1 req/sec
        if out:
            return out
    return []


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cues", required=True)
    ap.add_argument("--write", action="store_true", help="write the dates back into the CSV")
    a = ap.parse_args()

    with open(a.cues, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
        cols = list(rows[0].keys())
    if "release_date" not in cols:
        raise SystemExit("This cues file has no release_date column.")

    found = got_g = 0
    for i, t in enumerate(rows, 1):
        if (t.get("release_date") or "").strip() and (t.get("genres") or "").strip():
            found += 1
            continue
        artist, title = t["artist"], t["title"]
        hits = from_itunes(artist, title) or from_deezer(artist, title) or from_musicbrainz(artist, title)
        if hits:
            # earliest wins — a later date is a re-release / compilation
            date, genre, src = sorted(hits)[0]
            t["release_date"] = date
            # Genre is only ever FILLED IN, never overwritten: the DJ's own tag
            # beats a store's guess, same rule track-sources follows for bpm/key.
            if genre and not (t.get("genres") or "").strip():
                t["genres"] = genre
                got_g += 1
            found += 1
            print("  %2d  %-34s %s   (%s, %d hit%s)%s"
                  % (i, artist[:34], date, src, len(hits), "" if len(hits) == 1 else "s",
                     ("  +genre " + genre) if genre and t["genres"] == genre else ""))
        else:
            print("  %2d  %-34s —        no confident match: %s" % (i, artist[:34], title[:44]))
        time.sleep(0.25)

    print("\n%d/%d dated" % (found, len(rows)))
    if a.write:
        with io.open(a.cues, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols)
            w.writeheader()
            w.writerows(rows)
        print("wrote", a.cues)
    else:
        print("(dry run — pass --write to save)")


if __name__ == "__main__":
    main()
