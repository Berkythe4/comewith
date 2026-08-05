"""krney_times_sheet.py - write the KRNeY start-times confirmation sheet
plus the genre/hashtag block for every set.

The tracklist only exists as a photo of a handwritten page, so every line in
TRACKS is a READING, not data. Uncertain ones are marked (?) and listed again
at the top so Keith checks nine lines rather than re-reading twenty-seven.

To apply corrections: edit TRACKS below and re-run. The unsure list, the gap
check and the tag block all rebuild themselves from it.
"""
import collections, csv, io, json, os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

RENDER = r"C:\Users\Admin\Documents\Comewith\Radio\Elements-26\render"
OUT = os.path.join(RENDER, "Elements_Ep2_KRNeY_times.txt")

# (start, song as written, sure?, note)
# sure=False => I could not read it confidently off the photo.
TRACKS = [
    ("0:00",  "Fast lane",          True,  "no time written - assumed the opener"),
    ("2:04",  "Curves",             True,  ""),
    ("4:16",  "Mask off",           True,  ""),
    ("5:51",  "Purple Rhythm",      False, "second word unclear"),
    ("7:35",  "Water Falls",        True,  ""),
    ("9:21",  "Evergreen",          False, "reads Evkvround / Everground - unclear"),
    ("10:16", "Orange",             True,  ""),
    ("12:06", "God speed",          True,  "first time crossed out; 12:06 confirmed"),
    ("13:56", "Type Shit",          True,  ""),
    ("16:41", "Wet me",      True,  "name and last digit both overwritten"),
    ("18:36", "Telekness",          False, "time overwritten; name may be 'Telekinesis'"),
    ("21:04", "Gimme Gimme",        True,  ""),
    ("23:00", "Carry me home",      True,  ""),
    ("25:11", "Feel the vibe",      True,  "last digit unclear"),
    ("28:11", "Forever young",      True,  "first time crossed out; 28:11 confirmed"),
    ("32:22", "The Fade",           False, "could be 'The Jade'"),
    ("34:18", "Feel it in my",      True,  "title unfinished on the page; time given as 3481 - read as 34:18"),
    ("37:16", "Make you Right",     True,  ""),
    ("39:59", "Psycho",             False, "seconds unclear - 39:59 or 39:39"),
    ("42:05", "Long walks",         True,  ""),
    ("44:08", "amigud",             False, "name unclear"),
    ("47:42", "gossip",             True,  ""),
    ("50:24", "Hype up",            False, "reads 'Hytyup' - unclear"),
    ("51:36", "Dream On",           True,  ""),
    ("54:03", "Torch",              True,  "'Gorillat' written beside it - the artist"),
    ("55:58", "Surround sounds",    True,  ""),
    ("57:19", "Let 'em talk",       True,  "'Mikayli' written beside it - the artist"),
]

MIX = os.path.join(RENDER, "Audio_Video_Final", "CWR_ElementsEp2_KRNeY.wav")

# What krney_identify.py could work out by matching each title against the 157
# booked Elements artists. INFERENCES, not confirmations — the only one KRNeY
# corroborated himself is Mikayli, which he wrote on the page.
# The aligned tracklist: KRNeY's handwritten TIMES married to the real artists
# and titles he entered in the dashboard (krney_align.py), then dated. Keyed by
# start time, which is the one field both halves agree on exactly.
IDENT = {}
_cp = os.path.join(RENDER, "Elements_Ep2_KRNeY_cues.csv")
if os.path.exists(_cp):
    for r in csv.DictReader(open(_cp, encoding="utf-8-sig")):
        IDENT[r["start"]] = r


def mmss(s):
    s = int(s); return "%d:%02d" % (s // 60, s % 60)


def secs(t):
    p = [int(x) for x in t.split(":")]
    return p[0] * 60 + p[1] if len(p) == 2 else p[0] * 3600 + p[1] * 60 + p[2]


def genres_for(path):
    rows = list(csv.DictReader(open(os.path.join(RENDER, path), encoding="utf-8-sig")))
    c = collections.Counter()
    for r in rows:
        for g in [x.strip() for x in r["genres"].replace("/", ",").split(",") if x.strip()]:
            c[g.title()] += 1
    return len(rows), c


# Genre strings come from Rekordbox tags and SoundCloud fields, and some of them
# are PLATFORM CATEGORIES or label names rather than genres. Naively slugging them
# puts #dancedj and #matrodasound on a public upload, which helps nobody and looks
# careless. DROP = not a genre; the rest map to what people actually search.
DROP = {"Dance & Dj", "Matrodasound"}
TAGS = {
    "Dance & Edm": ["#edm"],
    "Drum & Bass": ["#dnb", "#drumandbass"],
    "Ukg":         ["#ukgarage", "#ukg"],
    "Deep Tech":   ["#deeptech"],
    "Future Bass": ["#futurebass"],
    "Electro Pop": ["#electropop"],
    "Tech House":  ["#techhouse"],
}


def tags_for(g):
    if g in DROP:
        return []
    return TAGS.get(g, ["#" + "".join(ch for ch in g.lower() if ch.isalnum())])


def build():
    L = []
    A = L.append
    A("EP 2 - COME WITH ELEMENTS RADIO : KRNeY - TRACK START TIMES")
    A("=" * 86)
    A("MIX FILE   : CWR_ElementsEp2_KRNeY.wav")
    mixlen = ""
    try:
        import subprocess
        mixlen = mmss(float(subprocess.check_output(
            ["ffprobe", "-v", "error", "-show_entries", "format=duration",
             "-of", "default=nk=1:nw=1", MIX], text=True).strip()))
    except Exception:
        mixlen = "?"
    A("MIX LENGTH : %s" % mixlen)
    A("SOURCE     : CWR_ElementsEp2_KRNeY_TrackList.jpg (handwritten, read off the photo)")
    A("TRACKS     : %d" % len(TRACKS))
    A("")
    A("READ THIS FIRST")
    A("  These are my READINGS of your handwriting, not data from a file. The times are")
    A("  legible almost everywhere; several song NAMES are not. Correct anything wrong,")
    A("  fill in the artists, and I will build the cues and render from it.")
    A("")
    unsure = [(i, t) for i, t in enumerate(TRACKS, 1) if not t[2]]
    A("  %d line(s) I could not read with confidence - please check these first:" % len(unsure))
    for i, t in unsure:
        A("     %2d. %-18s %-7s  %s" % (i, t[1], t[0], t[3]))
    A("")
    A("  The last track runs to the end of the mix (%s)." % mixlen)
    A("")
    A("-" * 86)
    A("  #   START    SONG                      ARTIST               YEAR GENRE")
    A("-" * 86)
    for i, (start, name, sure, note) in enumerate(TRACKS, 1):
        mark = "  " if sure else " ?"
        d = IDENT.get(start)
        if d:
            A(" %2d%s  %-7s  %-24s  %-20s %-4s %s"
              % (i, mark, start, d["title"][:24], d["artist"][:20],
                 (d["release_date"] or "-")[:4], d["genres"] or ""))
        else:
            A(" %2d%s  %-7s  %-24s  %s" % (i, mark, start, name[:24], "** not matched - as written **"))
    A("-" * 86)
    A("")
    A("NOTES FROM THE PAGE")
    for i, (start, name, sure, note) in enumerate(TRACKS, 1):
        if note:
            A("   %2d  %-24s %s" % (i, name[:24], note))
    A("")
    A("GAPS BETWEEN TRACKS (a very short or very long one usually means a misread time)")
    A("")
    for i in range(len(TRACKS) - 1):
        d = secs(TRACKS[i + 1][0]) - secs(TRACKS[i][0])
        flag = ""
        if d < 60:
            flag = "  <-- under a minute, check"
        elif d > 240:
            flag = "  <-- over four minutes, check"
        A("   %2d -> %2d   %-7s to %-7s   %s%s"
          % (i + 1, i + 2, TRACKS[i][0], TRACKS[i + 1][0], mmss(d), flag))
    A("")
    A("")
    A("=" * 86)
    A("GENRES + HASHTAGS FOR EACH SET")
    A("=" * 86)
    A("")
    A("Put the ALWAYS block on every upload, then that set's own genres under it.")
    A("SoundCloud reads the first few tags hardest, so lead with the specific ones and")
    A("let the brand tags follow. YouTube only surfaces the first three above the title,")
    A("so the order below is the order to paste.")
    A("")
    A("ALWAYS (every episode, both platforms)")
    A("   #comewithradio #comewithnyc #elementsfestival #elements2026 #electronicmusic")
    A("   #djmix #livemix #brooklyn #nyc")
    A("")
    
    SETS = [
        ("EP 1 - WATER - Berky - Thu Aug 6", "Elements_Ep1_cues.csv"),
        ("EP 3 - EARTH - Henry - Sat Aug 8", "Elements_Ep3_Henry_cues.csv"),
    ]
    for label, f in SETS:
        n, c = genres_for(f)
        A(label)
        A("   %d tracks. Genres actually in the set, most common first:" % n)
        items = ["%s (%d)" % (g, k) for g, k in c.most_common() if g not in DROP]
        line = "      "
        for it in items:
            if len(line) + len(it) + 2 > 84:
                A(line.rstrip().rstrip(",")); line = "      "
            line += it + ", "
        A(line.rstrip().rstrip(","))
        dropped = [g for g, _ in c.most_common() if g in DROP]
        if dropped:
            A("      (ignored, not genres: %s)" % ", ".join(dropped))
        A("   Tags:")
        tags, seen = [], set()
        for g, _ in c.most_common():
            for t in tags_for(g):
                if t not in seen:
                    seen.add(t); tags.append(t)
        line = "      "
        for t in tags:
            if len(line) + len(t) + 1 > 84:
                A(line); line = "      "
            line += t + " "
        A(line.rstrip())
        A("")
    
    A("EP 2 - FIRE - KRNeY - Fri Aug 7")
    if IDENT:
        c2 = collections.Counter()
        for d in IDENT.values():
            for g in [x.strip() for x in (d.get("genres") or "").replace("/", ",").split(",") if x.strip()]:
                c2[g.title()] += 1
        A("   %d of %d written lines matched to the tracklist KRNeY built in the dashboard" % (len(IDENT), len(TRACKS)))
        A("   through his DJ link. Those artists and titles are HIS, not a guess - the")
        A("   handwriting only supplied the times. Genres and years were then looked up")
        A("   from those real artist + title pairs.")
        A("   Genres across the identified tracks:")
        items = ["%s (%d)" % (g, k) for g, k in c2.most_common() if g not in DROP]
        line = "      "
        for it in items:
            if len(line) + len(it) + 2 > 84:
                A(line.rstrip().rstrip(",")); line = "      "
            line += it + ", "
        A(line.rstrip().rstrip(","))
        A("   Tags (only %d of the tracks carry a genre so far):" % sum(1 for d in IDENT.values() if d.get("genres")))
        tags, seen2 = [], set()
        for g, _ in c2.most_common():
            for t in tags_for(g):
                if t not in seen2:
                    seen2.add(t); tags.append(t)
        line = "      "
        for t in tags:
            if len(line) + len(t) + 1 > 84:
                A(line); line = "      "
            line += t + " "
        A(line.rstrip())
        A("   %d written line(s) matched nothing in his list, and 4 of his tracks have no" % (len(TRACKS) - len(IDENT)))
        A("   written line: Candy Shop Remix, Purple Plum Trees, Hot Mic, FVKVRVND.")
        A("   Those are the open questions - see the table above.")
    else:
        A("   Cannot be listed yet - run krney_identify.py first.")
    A("")
    A("EP 4 - AIR - 32LVS - Sun Aug 9")
    A("   No tracklist received yet.")
    A("")
    A("=" * 86)
    
    txt = "\n".join(L) + "\n"
    io.open(OUT, "w", encoding="utf-8", newline="\r\n").write(txt)
    bad = sum(1 for ch in txt if ord(ch) > 126 and ch not in "\r\n")
    print("wrote", OUT)
    print("non-ASCII chars:", bad, "| lines:", len(L), "| unsure:", len(unsure))
    

if __name__ == "__main__":
    build()
