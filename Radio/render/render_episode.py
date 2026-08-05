#!/usr/bin/env python3
"""
render_episode.py — turn a recorded mix + a cues CSV into the Option 4
"Now Playing" video (1920x1080, YouTube-ready MP4).

Pipeline: Pillow renders one card PNG per track (cover + artist/title/BPM/key +
up-next), then ffmpeg holds each card for its track's duration, animates a lime
progress bar across the bottom (drawbox width = t/total), and muxes the audio.
No camera, no editor.

Usage:
    python Radio/render/render_episode.py \
        --cues Radio/render/EP1_cues.csv \
        --audio "Radio/Video/EP1_mix.wav" \
        --cover Radio/Artwork/Radio_Thumbnail.jpg \
        --out Radio/Video/EP1.mp4 --ep "EP 1"

The video opens with an intro slide and ends with a closing slide (staged reveal,
overlaid on the mix so total runtime is unchanged) — see draw_intro/draw_outro.
`make_episode.py` fills the bookend meta (mixed-by, drop dates) from prod for you.

`start` in the cues CSV is mm:ss (e.g. 4:12) or plain seconds. Any blank starts
are reported and abort, so you never render a half-timed video by accident.

  --dry            render cards + a 1s preview only (fast sanity check)
  --title TEXT     show name (default "Come With Radio")
  --mixed-by TEXT  intro credit;  --drop-date / --next-date  YYYY-MM-DD (bookends)
  --no-bookends    skip the intro + closing;  --intro-secs / --outro-secs  retime
"""
import argparse, csv, os, random, re, subprocess, sys, tempfile
import sys as _sys
try:
    _sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
from PIL import Image, ImageDraw, ImageFont, ImageFilter

W, H = 1920, 1080
FONTS = os.environ.get("WINDIR", "C:/Windows") + "/Fonts"

# ---- brand tokens (from radio.html) -----------------------------------------
BG_TOP, BG_BOT = (30, 17, 40), (15, 8, 20)
CREAM, DIM, FAINT = (237, 228, 211), (179, 167, 184), (124, 113, 131)
LIME, LIME_DK, HOT, LINE = (167, 224, 46), (143, 179, 57), (255, 92, 168), (74, 50, 82)

# Progress-bar geometry — SHARED with ffmpeg's drawbox so the moving fill lands
# exactly on the track drawn in the PNG. Edit here only.
BAR_X, BAR_H = 96, 16
BAR_W = W - 2 * BAR_X
BAR_Y = H - 118

def font(name, size):
    return ImageFont.truetype(os.path.join(FONTS, name), size)

# Arial Black = the Archivo-ish heavy display; Consolas = the DM-Mono role.
def F_disp(s): return font("ariblk.ttf", s)
def F_bold(s): return font("arialbd.ttf", s)
def F_body(s): return font("arial.ttf", s)
def F_mono(s): return font("consola.ttf", s)
def F_monob(s): return font("consolab.ttf", s)

def parse_start(v):
    v = (v or "").strip()
    if not v:
        return None
    if ":" in v:
        parts = [float(p) for p in v.split(":")]
        if len(parts) == 2:
            return parts[0] * 60 + parts[1]
        if len(parts) == 3:
            return parts[0] * 3600 + parts[1] * 60 + parts[2]
    return float(v)

def fmt_clock(sec):
    sec = int(round(sec)); return "%d:%02d" % (sec // 60, sec % 60)

def fmt_show_date(s):
    s = (s or "").strip()
    if not s:
        return ""
    try:
        from datetime import datetime
        dt = datetime.strptime(s[:10], "%Y-%m-%d")
        return "%s  %s %d" % (dt.strftime("%a"), dt.strftime("%b"), dt.day)   # Fri  Jul 31
    except Exception:
        return s

def ffprobe_duration(path):
    out = subprocess.check_output(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration",
         "-of", "default=nk=1:nw=1", path], text=True)
    return float(out.strip())

def rounded_cover(path, size, radius):
    im = Image.open(path).convert("RGB")
    # center-crop to square, then resize
    s = min(im.size); l = (im.width - s) // 2; t = (im.height - s) // 2
    im = im.crop((l, t, l + s, t + s)).resize((size, size), Image.LANCZOS)
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, size, size], radius, fill=255)
    out = Image.new("RGBA", (size, size), (0, 0, 0, 0)); out.paste(im, (0, 0), mask)
    return out

def make_background():
    # vertical gradient + a soft lime glow top-right (blurred ellipse)
    bg = Image.new("RGB", (W, H), BG_BOT)
    top = Image.new("RGB", (1, H), BG_TOP)
    grad = Image.new("L", (1, H))
    for y in range(H):
        grad.putpixel((0, y), int(255 * (1 - y / H) ** 1.3))
    bg = Image.composite(top.resize((W, H)), bg, grad.resize((W, H)))
    glow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow)
    gd.ellipse([W * 0.55, -H * 0.35, W * 1.15, H * 0.55], fill=(167, 224, 46, 34))
    glow = glow.filter(ImageFilter.GaussianBlur(150))
    bg = Image.alpha_composite(bg.convert("RGBA"), glow).convert("RGB")
    return bg

def _layer():
    return Image.new("RGBA", (W, H), (0, 0, 0, 0))


def _veil(im, blur):
    """Blur a field until its EDGE is gone. Anything large enough to tint the
    frame has to lose its outline, or it reads as a drawn shape rather than
    light — the first Elements pass left a visible arc where the pool ended."""
    return im.filter(ImageFilter.GaussianBlur(blur))


def _settle(bg, washes, marks, wash_alpha=0.62):
    """Colour fields first, pulled back so the purple stays the ground; then the
    thin work at full weight. Marks carry almost no colour area, so they can be
    bright without shifting the ground — and structure is what reads as an
    element, not opacity."""
    lay = _layer()
    for p in washes:
        lay = Image.alpha_composite(lay, p)
    a = lay.split()[3]
    lay.putalpha(a.point(lambda v: int(v * wash_alpha)))
    for p in marks:
        lay = Image.alpha_composite(lay, p)
    return Image.alpha_composite(bg.convert("RGBA"), lay).convert("RGB")


def make_background_water():
    """WATER — the single-element backdrop for Berky's episode.

    Each episode in the run is assigned an element and its cover art says which;
    the backdrop should agree. This is water alone rather than all four, and it
    still has to stay out of the way: the purple remains the ground, and every
    piece sits where the card is empty — the pool low, the surface light in the
    dead band above the progress rail, bubbles up the outer edges clear of the
    cover on the left and the track text on the right.

    What makes it read as water is the horizontal REPETITION, not the blue. So
    the wash is weak and the ripples do the work, with amplitude and spacing
    growing toward the bottom so the surface has some depth to it.
    """
    import math
    bg = make_background()
    rnd = random.Random(20260806)            # fixed: the backdrop must not shimmer
    washes, marks = [], []

    # The pool. Two fields — a deep cool body and a lighter shelf above it — so
    # the colour shifts with depth instead of sitting flat.
    p = _layer(); pd = ImageDraw.Draw(p)
    pd.ellipse([-W * 0.30, H * 0.66, W * 1.30, H * 1.70], fill=(26, 118, 176, 58))
    pd.ellipse([-W * 0.10, H * 0.86, W * 1.10, H * 1.50], fill=(18, 74, 132, 46))
    washes.append(_veil(p, 210))

    # Surface light: a soft pale band where the water meets the air, in the empty
    # strip under the track text and above the progress rail.
    s = _layer(); sd = ImageDraw.Draw(s)
    sd.ellipse([-W * 0.20, H * 0.60, W * 1.20, H * 0.80], fill=(150, 214, 234, 30))
    washes.append(_veil(s, 150))

    # Ripples. Broken into segments with gaps rather than drawn edge to edge:
    # full-width lines plus the crossing caustics I tried first read as a
    # WIREFRAME GRID, not water — a mesh of straight lines is the one thing real
    # water never makes. Light on a surface is interrupted, so these are too, and
    # each segment carries its own weight so the band shimmers instead of ruling.
    r = _layer(); rd = ImageDraw.Draw(r)
    for k in range(15):
        t = k / 14.0
        y0 = H * (0.615 + 0.40 * t ** 1.25)
        amp = 4 + 15 * t
        base = 26 + 62 * t                      # nearer the bottom = closer = brighter
        wid = 3 if t < 0.6 else 4
        x = rnd.uniform(-40, 120)
        while x < W + 20:
            seg = rnd.uniform(160, 520) * (0.6 + 0.8 * t)
            a_seg = int(base * rnd.uniform(0.45, 1.0))
            pts = [(px, y0 + amp * math.sin(px / (120.0 + 90.0 * t) + k * 0.7))
                   for px in range(int(x), int(min(x + seg, W + 20)), 8)]
            if len(pts) > 1 and a_seg > 5:
                rd.line(pts, fill=(168, 230, 246, a_seg), width=wid, joint="curve")
            x += seg + rnd.uniform(70, 260)     # the gap is what stops it ruling
    # Blur barely at all. A 2px line under a 5px blur spreads over ~15px and its
    # peak alpha collapses — the first water pass drew all of this and none of it
    # was visible. Softness has to come from the alpha, not from the blur.
    marks.append(_veil(r, 2))

    # Bubbles rising up the outer edges — the left column is clear of the cover,
    # the right of the track text, so nothing drifts behind a word.
    b = _layer(); bd = ImageDraw.Draw(b)
    for _ in range(70):
        side = rnd.random() < 0.5
        x = W * (rnd.uniform(0.005, 0.055) if side else rnd.uniform(0.945, 0.995))
        y = H * (1.02 - 0.72 * rnd.random() ** 0.7)
        rad = rnd.uniform(1.6, 5.2)
        a = int(190 * max(0.0, 1.0 - (H - y) / (H * 0.72)) ** 1.2 * rnd.uniform(0.55, 1.0))
        if a > 6:
            bd.ellipse([x - rad, y - rad, x + rad, y + rad], outline=(202, 242, 252, a), width=2)
    marks.append(_veil(b, 1))

    return _settle(bg, washes, marks, wash_alpha=0.58)


def make_background_fire():
    """FIRE — Martin. Cover is a red/orange flame swirl.

    Heat comes from the EMBERS, not the wash: orange spread wide over purple just
    makes brown. So the wash stays low and tight to the floor, and the sparks and
    the updraught do the reading.
    """
    import math
    bg = make_background()
    rnd = random.Random(20260807)
    washes, marks = [], []

    f = _layer(); fd = ImageDraw.Draw(f)
    fd.ellipse([-W * 0.25, H * 0.74, W * 1.25, H * 1.60], fill=(196, 34, 8, 62))
    fd.ellipse([-W * 0.05, H * 0.88, W * 1.05, H * 1.44], fill=(242, 88, 4, 46))
    washes.append(_veil(f, 205))

    # Updraught: tapering vertical wisps off the floor. Curved and uneven —
    # straight ones would read as a fence.
    u = _layer(); ud = ImageDraw.Draw(u)
    for k in range(22):
        x0 = W * rnd.uniform(0.0, 1.0)
        h = H * rnd.uniform(0.14, 0.34)
        a0 = 40 + 46 * rnd.random()
        sway = rnd.uniform(18, 62) * (1 if rnd.random() < 0.5 else -1)
        pts = []
        n = 16
        for i in range(n + 1):
            t = i / n
            pts.append((x0 + sway * math.sin(t * 2.4 + k), H * 1.01 - h * t))
        for i in range(n):
            a = int(a0 * (1 - i / n) ** 1.5)
            if a > 4:
                ud.line([pts[i], pts[i + 1]], fill=(255, 158, 62, a), width=3)
    marks.append(_veil(u, 4))

    # Sparks, densest low, thinning as they rise.
    em = _layer(); ed = ImageDraw.Draw(em)
    for _ in range(150):
        x = W * rnd.random()
        y = H * (1.03 - 0.60 * rnd.random() ** 0.6)
        r = rnd.uniform(1.3, 4.4)
        a = int(225 * max(0.0, 1.0 - (H - y) / (H * 0.60)) ** 1.4 * rnd.uniform(0.5, 1.0))
        if a > 6:
            ed.ellipse([x - r, y - r, x + r, y + r], fill=(255, 198, 116, min(235, a)))
    marks.append(_veil(em, 2))
    return _settle(bg, washes, marks, wash_alpha=0.60)


def make_background_earth():
    """EARTH — Henry. Cover is aerial terrain: olive fields and river channels.

    So this is CONTOURS, not strata. Nested meandering lines read as land seen
    from above and, importantly, don't collide with water's parallel ripples —
    two elements in one run must not look like each other.
    """
    import math
    bg = make_background()
    rnd = random.Random(20260808)
    washes, marks = [], []

    g = _layer(); gd = ImageDraw.Draw(g)
    gd.ellipse([-W * 0.25, H * 0.70, W * 1.25, H * 1.60], fill=(96, 104, 24, 66))
    gd.ellipse([-W * 0.05, H * 0.86, W * 1.05, H * 1.40], fill=(132, 108, 32, 44))
    washes.append(_veil(g, 205))

    # Contour lines: each one a wandering curve, each nested slightly inside the
    # last, so they group into landforms instead of ruling parallel lines.
    c = _layer(); cd = ImageDraw.Draw(c)
    for band in range(4):
        seed_y = H * (0.70 + 0.085 * band)
        f1, f2 = rnd.uniform(180, 320), rnd.uniform(420, 900)
        p1, p2 = rnd.uniform(0, 6.3), rnd.uniform(0, 6.3)
        for k in range(5):
            off = k * rnd.uniform(9, 17)
            a = int((104 + 74 * band / 3.0) * (1 - k / 7.0))
            pts = [(x, seed_y + off + 26 * math.sin(x / f1 + p1) + 44 * math.sin(x / f2 + p2))
                   for x in range(-20, W + 40, 12)]
            cd.line(pts, fill=(214, 202, 126, max(26, a)), width=3, joint="curve")
    marks.append(_veil(c, 2))

    # Grain — dry dust rather than rising sparks, so it sits low and still.
    d_ = _layer(); dd = ImageDraw.Draw(d_)
    for _ in range(190):
        x = W * rnd.random()
        y = H * (1.02 - 0.34 * rnd.random() ** 0.5)
        r = rnd.uniform(0.9, 2.6)
        a = int(175 * rnd.uniform(0.4, 1.0))
        dd.ellipse([x - r, y - r, x + r, y + r], fill=(214, 190, 128, a))
    marks.append(_veil(d_, 1))
    return _settle(bg, washes, marks, wash_alpha=0.58)


def make_background_air():
    """AIR — 32LVS. Cover is cyan smoke against black.

    Sinuous tapering wisps, not the straight lime streaks the four-element
    backdrop uses for wind: those two would otherwise be the same gesture twice.
    Held to the upper band and the outer thirds, clear of the track text.
    """
    import math
    bg = make_background()
    rnd = random.Random(20260809)
    washes, marks = [], []

    a_ = _layer(); ad = ImageDraw.Draw(a_)
    ad.ellipse([W * 0.30, -H * 0.42, W * 1.24, H * 0.42], fill=(52, 152, 190, 44))
    ad.ellipse([-W * 0.24, H * 0.06, W * 0.42, H * 0.72], fill=(40, 128, 168, 30))
    washes.append(_veil(a_, 195))

    # Wisps: long, slow, thinning at both ends — smoke, not wind.
    w = _layer(); wd = ImageDraw.Draw(w)
    for k in range(14):
        y0 = H * rnd.uniform(0.05, 0.34) if k % 2 == 0 else H * rnd.uniform(0.62, 0.94)
        x0 = W * rnd.uniform(-0.15, 0.55)
        span = W * rnd.uniform(0.35, 0.85)
        amp = rnd.uniform(22, 70)
        f = rnd.uniform(260, 520)
        a0 = 104 + 62 * rnd.random()
        n = 40
        pts = [(x0 + span * (i / n), y0 + amp * math.sin((x0 + span * (i / n)) / f + k))
               for i in range(n + 1)]
        for i in range(n):
            t = i / n
            a = int(a0 * math.sin(math.pi * t) ** 0.8)      # fade in and out at the ends
            if a > 4:
                wd.line([pts[i], pts[i + 1]], fill=(176, 232, 248, a), width=4)
    marks.append(_veil(w, 3))

    m = _layer(); md = ImageDraw.Draw(m)
    for _ in range(70):
        x = W * rnd.random()
        y = H * rnd.random()
        r = rnd.uniform(1.0, 3.2)
        md.ellipse([x - r, y - r, x + r, y + r], fill=(220, 242, 252, int(120 * rnd.uniform(0.3, 1.0))))
    marks.append(_veil(m, 2))
    return _settle(bg, washes, marks, wash_alpha=0.56)


def make_background_ether():
    """ETHER — Janelle. The one with no natural vocabulary, so it takes it from
    the cover: a pale radiant bloom ringed by faint geometry.

    Kept to the top-right, where the lime glow already lives and the card is
    empty, plus a sparse starfield. This is the one I'd most want a second
    opinion on — the other four describe themselves; ether doesn't.
    """
    import math
    bg = make_background()
    rnd = random.Random(20260810)
    washes, marks = [], []

    # Pushed up and right so the rings clear the artist name — display type is
    # the one thing on the card nothing may cross.
    cx, cy = W * 0.87, H * 0.15
    b = _layer(); bd = ImageDraw.Draw(b)
    bd.ellipse([cx - W * 0.30, cy - H * 0.42, cx + W * 0.30, cy + H * 0.42], fill=(150, 178, 226, 46))
    washes.append(_veil(b, 190))

    # Concentric polygons, rotating a little each ring — the cover's geometry,
    # drawn faintly enough to read as structure rather than shapes.
    p = _layer(); pd = ImageDraw.Draw(p)
    for k in range(7):
        rad = 105 + k * 62
        rot = k * 0.28
        sides = 6
        pts = [(cx + rad * math.cos(rot + i * 2 * math.pi / sides),
                cy + rad * 0.86 * math.sin(rot + i * 2 * math.pi / sides)) for i in range(sides)]
        pd.line(pts + [pts[0]], fill=(206, 220, 250, max(30, int(120 - k * 12))), width=3, joint="curve")
    marks.append(_veil(p, 2))

    # Rays out of the bloom, short and uneven.
    r_ = _layer(); rd = ImageDraw.Draw(r_)
    for k in range(20):
        ang = rnd.uniform(0, 6.283)
        r0 = rnd.uniform(90, 150); r1 = r0 + rnd.uniform(60, 240)
        a = int(72 + 62 * rnd.random())
        rd.line([(cx + r0 * math.cos(ang), cy + r0 * 0.86 * math.sin(ang)),
                 (cx + r1 * math.cos(ang), cy + r1 * 0.86 * math.sin(ang))],
                fill=(220, 232, 252, a), width=3)
    marks.append(_veil(r_, 3))

    s = _layer(); sd = ImageDraw.Draw(s)
    for _ in range(120):
        x = W * rnd.random(); y = H * rnd.random()
        r = rnd.uniform(0.8, 2.4)
        sd.ellipse([x - r, y - r, x + r, y + r], fill=(230, 236, 254, int(140 * rnd.uniform(0.25, 1.0))))
    marks.append(_veil(s, 1))
    return _settle(bg, washes, marks, wash_alpha=0.56)


def make_background_elements():
    """The weekly backdrop with the four elements worked into it.

    The brief was "keep the purple subtle" — so this does NOT replace the ground,
    it breathes four presences into the same gradient at very low alpha. Each one
    is placed where the card layout is EMPTY, so nothing ever competes with type:
    the cover sits left-of-centre, the track text right of it, the progress rail
    across the bottom.

        fire   low left, under the cover — embers, warm, heaviest blur
        water  low right — a cool pool with slow horizontal ripples
        air    top centre — a pale lifted haze, the lightest touch of the four
        wind   upper right — long thin streaks, carrying the brand lime

    Everything is drawn oversized and Gaussian-blurred, so there is no visible
    edge anywhere — at this alpha the eye reads atmosphere, not shapes.
    """
    import math
    bg = make_background()                       # same purple, same lime glow

    def layer():
        return Image.new("RGBA", (W, H), (0, 0, 0, 0))

    def veil(im, blur):
        """Any large colour field gets blurred until its EDGE is gone. A pool this
        size at blur 26 leaves a visible arc across the frame — the giveaway that
        it is a drawn ellipse rather than light."""
        return im.filter(ImageFilter.GaussianBlur(blur))

    # Two kinds of layer, treated very differently.
    #   WASHES are broad colour fields. They are what tint the ground, so they are
    #     kept weak and pulled back hard — a wide orange wash over purple just
    #     makes brown, which is what went wrong first time.
    #   MARKS are thin: embers, ripples, streaks, motes. They carry almost no
    #     colour area, so they can be much brighter without moving the ground at
    #     all — and STRUCTURE is what actually reads as an element. Depicting the
    #     four is a drawing problem, not an opacity problem.
    washes, marks = [], []

    # FIRE — low left, tight into the corner, heat rising out of it.
    f = layer(); fd = ImageDraw.Draw(f)
    fd.ellipse([-W * 0.20, H * 0.78, W * 0.32, H * 1.42], fill=(206, 48, 18, 52))
    fd.ellipse([-W * 0.12, H * 0.90, W * 0.22, H * 1.30], fill=(255, 122, 30, 40))
    washes.append(veil(f, 200))

    # Scatter from a FIXED SEED: identical on every render (a backdrop that
    # shimmered between frames would be unusable) but with no structure.
    # The first version stepped both x and y by golden-ratio multiples of the
    # same index, which puts every particle on a lattice — the embers came out
    # as three straight diagonal lines across the frame, reading as scratches
    # rather than sparks. Two irrationals do not make a random 2-D scatter.
    rnd = random.Random(20260806)

    em = layer(); ed = ImageDraw.Draw(em)
    for _ in range(130):
        x = W * rnd.uniform(0.0, 0.34)
        y = H * (1.03 - 0.62 * rnd.random() ** 0.62)
        r = rnd.uniform(1.4, 5.0)
        a = int(235 * max(0.0, 1.0 - (H - y) / (H * 0.62)) ** 1.35 * rnd.uniform(0.55, 1.0))
        if a > 5:
            ed.ellipse([x - r, y - r, x + r, y + r], fill=(255, 196, 108, min(235, a)))
    marks.append(veil(em, 3))

    # WATER — low right. The pool is veiled to nothing; the ripples do the work.
    w = layer(); wd = ImageDraw.Draw(w)
    wd.ellipse([W * 0.62, H * 0.72, W * 1.30, H * 1.42], fill=(30, 132, 190, 58))
    washes.append(veil(w, 200))

    rip = layer(); rd = ImageDraw.Draw(rip)
    for k in range(9):
        y0 = H * (0.775 + k * 0.032)
        amp = 9 + k * 2.2
        a = int(120 - k * 10)
        pts = [(x, y0 + amp * math.sin(x / 150.0 + k * 0.85))
               for x in range(int(W * 0.58), W + 40, 10)]
        rd.line(pts, fill=(158, 232, 246, max(24, a)), width=3, joint="curve")
    marks.append(veil(rip, 5))

    # AIR — NOT a haze. A pale wash across the top is what greyed the whole upper
    # frame first time and fought the lime glow. Air is drift instead: motes
    # lifting up the left side, clear of the header and the track text.
    ai = layer(); ad = ImageDraw.Draw(ai)
    for _ in range(72):
        t = rnd.random()
        x = W * rnd.uniform(0.01, 0.33)
        y = H * (0.06 + 0.62 * t)
        r = rnd.uniform(1.1, 3.6)
        ad.ellipse([x - r, y - r, x + r, y + r],
                   fill=(230, 226, 250, int((60 + 80 * (1 - t)) * rnd.uniform(0.5, 1.0))))
    marks.append(veil(ai, 4))

    # WIND — long streaks off the upper right in the brand lime, so the fourth
    # element and the accent are one gesture. Held below the header line and
    # above the artist name, where the card is empty.
    # Brighter and paler than the pure brand lime: these sit inside the existing
    # lime glow in the top-right corner, and at the accent's own value they
    # simply vanished into it.
    wi = layer(); wid = ImageDraw.Draw(wi)
    for k in range(12):
        y0 = H * (0.105 + k * 0.025)
        amp = 14 + k * 3.2
        a = int(150 - k * 9)
        pts = [(x, y0 + amp * math.sin(x / 330.0 + k * 0.55))
               for x in range(int(W * 0.44), W + 60, 12)]
        wid.line(pts, fill=(206, 240, 140, max(34, a)), width=2, joint="curve")
    marks.append(veil(wi, 3))

    lay = layer()
    for p in washes:                             # colour first, pulled back
        lay = Image.alpha_composite(lay, p)
    al = lay.split()[3]
    lay.putalpha(al.point(lambda v: int(v * 0.62)))
    for p in marks:                              # then the drawing, at full weight
        lay = Image.alpha_composite(lay, p)
    return Image.alpha_composite(bg.convert("RGBA"), lay).convert("RGB")


# One backdrop per element, so a card agrees with its cover art. "elements"
# carries all four at once and is the run-overview look.
BACKDROPS = {"weekly": make_background,
             "elements": make_background_elements,
             "water": make_background_water,
             "fire": make_background_fire,
             "earth": make_background_earth,
             "air": make_background_air,
             "ether": make_background_ether}


def chip(draw, x, y, text, fnt, accent=False):
    tw = draw.textlength(text, font=fnt)
    ascent, descent = fnt.getmetrics(); th = ascent + descent
    padx, pady = 30, 18
    col = LIME if accent else CREAM
    border = LIME_DK if accent else LINE
    draw.rounded_rectangle([x, y, x + tw + 2 * padx, y + th + 2 * pady],
                           radius=(th + 2 * pady) // 2, outline=border, width=2)
    draw.text((x + padx, y + pady), text, font=fnt, fill=col)
    return x + tw + 2 * padx

def truncate(draw, text, fnt, maxw):
    if draw.textlength(text, font=fnt) <= maxw:
        return text
    while text and draw.textlength(text + "…", font=fnt) > maxw:
        text = text[:-1]
    return text + "…"

def render_card(bg, cover, track, idx, ntracks, ep_label, title_text, nxt, out_png, progress=0.0):
    im = bg.copy()
    d = ImageDraw.Draw(im)
    # header
    d.text((BAR_X, 66), ("%s  ·  %s" % (title_text.upper(), ep_label)).upper(),
           font=F_mono(30), fill=FAINT)
    live = "● NOW PLAYING"
    lw = d.textlength(live, font=F_monob(30))
    d.text((W - BAR_X - lw, 66), live, font=F_monob(30), fill=LIME)

    # cover
    cs = 470; cx, cy = BAR_X, (H - cs) // 2 - 30
    im.paste(cover, (cx, cy), cover)

    # meta column
    mx = cx + cs + 96
    mw = W - BAR_X - mx
    d.text((mx, cy + 8), "%02d / %02d" % (idx, ntracks), font=F_mono(32), fill=FAINT)
    # Shrink the artist to fit before resorting to an ellipsis. "Louis The Child,
    # Joey Purp" came out as "Louis The Chil…", which loses the second artist
    # entirely — on a card whose whole job is to credit them. Collaborations are
    # common enough that the name has to flex. Truncation stays as the last
    # resort for a single name genuinely too long at the floor size.
    aname = track["artist"] or "—"
    asize = 132
    while asize > 60 and d.textlength(aname, font=F_disp(asize)) > mw:
        asize -= 3
    artist = truncate(d, aname, F_disp(asize), mw)
    # Keep the baseline where it was at full size, so shorter names don't shift.
    d.text((mx, cy + 58 + (132 - asize) * 0.62), artist, font=F_disp(asize), fill=CREAM)
    # Same for the title — these run long once a track is credited as
    # "Dom Dolla & Tiga - Don't Worry Baby (Sam Sidewayz Remix)", and the remixer
    # is exactly the part an ellipsis eats.
    tname = track["title"] or ""
    tsize = 58
    while tsize > 38 and d.textlength(tname, font=F_body(tsize)) > mw:
        tsize -= 2
    ttitle = truncate(d, tname, F_body(tsize), mw)
    d.text((mx, cy + 214 + (58 - tsize) * 0.5), ttitle, font=F_body(tsize), fill=DIM)

    # song facts — genre + release year (about the SONG). Shown above the show
    # chips (which are about the artist's upcoming gig).
    gline = fmt_genre_release(track)
    if gline:
        d.text((mx, cy + 286), truncate(d, gline, F_mono(30), mw), font=F_mono(30), fill=DIM)
        chy = cy + 348
    else:
        chy = cy + 300
    # chips: the artist's upcoming show — date + venue (what viewers care about)
    chx = mx
    dt = fmt_show_date(track.get("show_date"))
    if dt:
        chx = chip(d, chx, chy, dt, F_mono(34), accent=True) + 22
    ven = (track.get("show_venue") or "").strip()
    if ven:
        ven = truncate(d, ven, F_mono(34), (mx + mw) - chx - 90)
        chip(d, chx, chy, ven, F_mono(34))

    # progress rail + lime fill baked in at this track's position (stepped per
    # track — fast + robust; a per-frame animated bar via ffmpeg geq is far too
    # slow for an hour-long render).
    d.rounded_rectangle([BAR_X, BAR_Y, BAR_X + BAR_W, BAR_Y + BAR_H],
                        radius=BAR_H // 2, fill=(48, 38, 58))
    fw = int(BAR_W * max(0.0, min(1.0, progress)))
    if fw > BAR_H:
        d.rounded_rectangle([BAR_X, BAR_Y, BAR_X + fw, BAR_Y + BAR_H],
                            radius=BAR_H // 2, fill=LIME)
    # No "UP NEXT". Naming the next track spoils the set for anyone who hasn't
    # read the tracklist — the whole point of listening through is not knowing
    # what's coming. "LAST TRACK" stays: it gives the ending its shape without
    # revealing a single thing.
    if not nxt:
        d.text((BAR_X, BAR_Y + 40), "LAST TRACK", font=F_mono(30), fill=FAINT)
    site = "comewith.org"
    sw = d.textlength(site, font=F_mono(30))
    d.text((W - BAR_X - sw, BAR_Y + 40), site, font=F_mono(30), fill=(179, 167, 184))

    im.convert("RGB").save(out_png)

def fmt_genre_release(t):
    """'HOUSE · GARAGE   —   RELEASED 2024' from a track's genres + release_date.
    genres may be a list, a pipe/comma string, or empty; release_date any date-ish
    string. Either piece is optional; returns '' when we know nothing."""
    parts = []
    g = t.get("genres")
    if isinstance(g, list):
        gl = [str(x).strip() for x in g if str(x).strip()]
    else:
        gl = [x.strip() for x in re.split(r"[|,/]", str(g or "")) if x.strip()]
    if gl:
        parts.append(" · ".join(gl[:2]).upper())
    rd = str(t.get("release_date") or "").strip()
    if rd:
        m = re.match(r"(\d{4})", rd)
        parts.append("RELEASED " + (m.group(1) if m else rd))
    return "   —   ".join(parts)

def fmt_md(s):
    s = (s or "").strip()
    if not s:
        return ""
    try:
        from datetime import datetime
        dt = datetime.strptime(s[:10], "%Y-%m-%d")
        return "%s %d" % (dt.strftime("%b"), dt.day)      # Jul 30
    except Exception:
        return s

def next_thursday(drop, given):
    """The closing 'plug back in' date. Prefer an explicit --next-date; else this
    episode's drop + 7 days; else nothing."""
    if given:
        return given
    if not drop:
        return ""
    try:
        from datetime import datetime, timedelta
        return (datetime.strptime(drop[:10], "%Y-%m-%d") + timedelta(days=7)).strftime("%Y-%m-%d")
    except Exception:
        return ""

def _ctext(d, cx, y, text, fnt, fill):
    """Draw horizontally-centered text; return the y just below it."""
    w = d.textlength(text, font=fnt)
    d.text((cx - w / 2, y), text, font=fnt, fill=fill)
    a, de = fnt.getmetrics(); return y + a + de

def _cdivider(d, cx, y, half=210):
    d.rounded_rectangle([cx - half, y, cx + half, y + 4], radius=2, fill=LIME_DK)

# ---- EDITION COPY -----------------------------------------------------------
# The bookends say what the episode IS, and that differs by edition. The weekly
# NYC show is "every artist is playing New York soon"; the Elements run is four
# consecutive nights of producers playing that festival. Keeping the copy in one
# table means a new special edition is a dict entry, not a fork of the renderer.
#
# `brand`      headline on the intro slide
# `intro_a/b`  the two body lines (a = what this is, b = what to do about it)
# `outro_a/b`  the two closing body lines
# `next_label` prefix on the closing date pill
# `tease`      final closing line; "" hides it
EDITIONS = {
    "weekly": {
        "brand": "COME WITH RADIO",
        "intro_a": "Every track is an artist playing New York — soon.",
        "intro_b": "Hear who's playing, grab your tickets, go live.",
        "intro_cta": "TICKETS · WHO'S PLAYING & WHERE · WHERE TO LISTEN",
        "outro_a": "Tickets to every artist you just heard — plus when &",
        "outro_b": "where they play next, and the mix to replay — at",
        "next_label": "WE PLUG BACK IN NEXT THURSDAY",
        "tease": "SOMETHING ELEMENTAL IS COMING",
    },
    "elements": {
        "brand": "COME WITH ELEMENTS RADIO",
        "intro_a": "Every track is a producer playing Elements this weekend.",
        "intro_b": "Four nights, four episodes. This one is Berky's run.",
        # No "TICKETS" here — Elements tickets aren't ours to sell, and pointing
        # at them from our own slide reads like we're the box office. What the
        # site actually gives you during the weekend is the schedule and the run.
        "intro_cta": "SET TIMES · WHO'S PLAYING & WHEN · THE FULL RUN",
        "outro_a": "Every artist you just heard plays Elements this weekend —",
        "outro_b": "set times, the rest of the run and the mix to replay — at",
        "next_label": "THE RUN CONTINUES",
        "tease": "FOUR NIGHTS · FOUR EPISODES · ONE WEEKEND",
        # The bookends belong to the RUN, not the episode: every night opens and
        # closes on ether while the track cards stay in that DJ's own element.
        # Defaulted here rather than left to a flag, so it cannot be forgotten on
        # one episode and quietly break the set.
        "bookend_backdrop": "ether",
    },
}
ED = EDITIONS["weekly"]          # replaced in main() by --edition

# ---- INTRO: one accumulating "slide" that tells the station's story ----------
# Fixed layout (lines never shift); `stage` reveals more of it each beat, so the
# story assembles in place, then the full slide holds. Overlays the opening of
# the mix so total runtime is unchanged.
def draw_intro(bg, cover_sm, ep_label, mixed_by, drop_date, stage):
    im = bg.copy(); d = ImageDraw.Draw(im); cx = W // 2
    cs = cover_sm.size[0]
    im.paste(cover_sm, (cx - cs // 2, 92), cover_sm)          # cover, top-center
    tag = "● ON AIR"
    tw = d.textlength(tag, font=F_monob(28))
    d.text((cx - tw / 2, 92 + cs + 22), tag, font=F_monob(28), fill=LIME)
    # The brand headline is drawn at whatever size FITS: "COME WITH ELEMENTS
    # RADIO" is 9 characters longer than "COME WITH RADIO" and would have run
    # off both edges at a fixed 86pt.
    bsize = 86
    while bsize > 40 and d.textlength(ED["brand"], font=F_disp(bsize)) > W - 2 * 110:
        bsize -= 2
    _ctext(d, cx, 92 + cs + 66, ED["brand"], F_disp(bsize), CREAM)
    if stage >= 1:                                            # credits
        cred = "  ·  ".join([p for p in [
            ep_label.upper(),
            ("MIXED BY " + mixed_by.upper()) if mixed_by else "",
            fmt_show_date(drop_date).upper()] if p])
        _ctext(d, cx, 92 + cs + 176, cred, F_mono(30), DIM)
    if stage >= 2:
        _cdivider(d, cx, 92 + cs + 232)
        _ctext(d, cx, 92 + cs + 262, ED["intro_a"], F_body(44), CREAM)
    if stage >= 3:
        _ctext(d, cx, 92 + cs + 330, ED["intro_b"], F_body(44), DIM)
    if stage >= 4:
        _ctext(d, cx, 92 + cs + 416, ED["intro_cta"], F_mono(28), FAINT)
        _ctext(d, cx, 92 + cs + 452, "comewith.org", F_disp(44), LIME)
    return im.convert("RGB")

# ---- CLOSING: sincere thanks, tracklist download, follow, next drop ----------
def draw_outro(bg, cover_sm, next_date, stage):
    im = bg.copy(); d = ImageDraw.Draw(im); cx = W // 2
    tag = "● THAT'S THE SHOW"
    tw = d.textlength(tag, font=F_monob(28))
    d.text((cx - tw / 2, 150), tag, font=F_monob(28), fill=LIME)
    _ctext(d, cx, 200, "THANK YOU FOR", F_disp(78), CREAM)
    _ctext(d, cx, 292, "PLUGGING IN.", F_disp(78), CREAM)
    if stage >= 1:
        _cdivider(d, cx, 430)
        _ctext(d, cx, 462, ED["outro_a"], F_body(40), CREAM)
        _ctext(d, cx, 512, ED["outro_b"], F_body(40), CREAM)
    if stage >= 2:
        _ctext(d, cx, 576, "comewith.org", F_disp(60), LIME)
    if stage >= 3:
        _ctext(d, cx, 690, "Follow @comewithnyc so you never miss a drop.", F_body(40), DIM)
    if stage >= 4:
        nd = fmt_md(next_date)
        txt = ED["next_label"] + ((" · " + nd.upper()) if nd else "")
        tw2 = d.textlength(txt, font=F_mono(38))
        pad = 30
        d.rounded_rectangle([cx - tw2 / 2 - pad, 772, cx + tw2 / 2 + pad, 772 + 78],
                            radius=40, outline=LIME_DK, width=2)
        d.text((cx - tw2 / 2, 772 + 20), txt, font=F_mono(38), fill=LIME)
    # Closing line under the pill — dimmer and smaller so the chip stays a clean
    # single line, and revealed last, after the date lands. On the weekly show
    # this is the cryptic tease for whatever is coming; inside a special edition
    # it names the run you are already in. "" hides it.
    if stage >= 5 and ED["tease"]:
        _ctext(d, cx, 892, ED["tease"], F_mono(30), DIM)
    return im.convert("RGB")

# reveal cadence (seconds per beat); last beat is the full-slide hold.
INTRO_BEATS = [1.7, 1.7, 2.4, 2.4, 2.4, 5.0]     # stages 0,1,2,3,4, then hold@4
OUTRO_BEATS = [1.8, 2.4, 2.2, 2.2, 2.6, 2.2, 5.0]  # stages 0..5, then hold@5

def build_intro(work, bg, cover_sm, ep_label, mixed_by, drop_date, total_secs, hold_extra=0.0):
    """PNG frames for the intro, scaled to fit total_secs. Returns [(png,dur)].

    `hold_extra` is added to the FINAL beat only — the one where the whole slide
    is assembled. Stretching total_secs instead would slow every reveal down with
    it; the ask is for the finished slide to stay up longer, not for the story to
    tell itself more slowly.
    """
    scale = total_secs / sum(INTRO_BEATS)
    out = []
    for i, base in enumerate(INTRO_BEATS):
        stage = min(i, len(INTRO_BEATS) - 2)      # last beat = hold on the final stage
        png = os.path.join(work, "intro_%02d.png" % i)
        draw_intro(bg, cover_sm, ep_label, mixed_by, drop_date, stage).save(png)
        dur = base * scale + (hold_extra if i == len(INTRO_BEATS) - 1 else 0.0)
        out.append((png, max(0.1, dur)))
    return out

def build_outro(work, bg, cover_sm, next_date, total_secs):
    scale = total_secs / sum(OUTRO_BEATS)
    out = []
    for i, base in enumerate(OUTRO_BEATS):
        stage = min(i, len(OUTRO_BEATS) - 2)      # last beat = hold on the final stage
        png = os.path.join(work, "outro_%02d.png" % i)
        draw_outro(bg, cover_sm, next_date, stage).save(png)
        out.append((png, max(0.1, base * scale)))
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cues", help="cues CSV (idx,start,artist,title,bpm,...)")
    ap.add_argument("--json", help="tracklist.json from match_mix.py — timing driven ONLY by this")
    ap.add_argument("--meta", help="optional cues CSV to pull bpm/key from, joined by order")
    ap.add_argument("--abitrate", default="192k")
    ap.add_argument("--audio", help="the mix; omit only with --no-audio")
    ap.add_argument("--no-audio", action="store_true",
                    help="render the SLIDES ONLY, silent — for reviewing the deck before the mix exists")
    ap.add_argument("--duration", help="total runtime for --no-audio (mm:ss, h:mm:ss or seconds). "
                                       "Defaults to the last cue's start + its own length.")
    ap.add_argument("--edition", default="weekly", choices=sorted(EDITIONS),
                    help="which bookend copy to use")
    ap.add_argument("--backdrop", default="weekly", choices=sorted(BACKDROPS),
                    help="'elements' works fire / air / wind / water into the same purple ground")
    # The bookends are the RUN's identity, not the episode's: across a series they
    # should open and close the same way whoever mixed it, while the track cards
    # stay in that DJ's own element. So they can take their own ground and art.
    ap.add_argument("--bookend-backdrop", choices=sorted(BACKDROPS),
                    help="backdrop for the intro + closing only (default: same as --backdrop)")
    ap.add_argument("--bookend-cover", help="cover art for the intro only (default: same as --cover)")
    ap.add_argument("--cover", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--ep", default="EP 1")
    ap.add_argument("--title", default="Come With Radio")
    ap.add_argument("--mixed-by", default="", help="intro credit — who mixed it")
    ap.add_argument("--drop-date", default="", help="this episode's drop date YYYY-MM-DD (intro)")
    ap.add_argument("--next-date", default="", help="next drop YYYY-MM-DD (closing 'plug back in'); default = drop-date + 7")
    ap.add_argument("--no-bookends", action="store_true", help="skip the intro + closing slides")
    ap.add_argument("--intro-secs", type=float, default=15.5)
    ap.add_argument("--intro-hold", type=float, default=5.0,
                    help="extra seconds the FULLY assembled intro slide stays up, on top of "
                         "--intro-secs. Reveal pacing is unaffected. 0 restores the old timing.")
    ap.add_argument("--outro-secs", type=float, default=16.0)
    ap.add_argument("--dry", action="store_true", help="cards + 1s preview only")
    a = ap.parse_args()

    global ED
    ED = EDITIONS[a.edition]

    if not a.audio and not a.no_audio:
        raise SystemExit("Give --audio, or --no-audio to render the slides on their own.")
    need = [(a.json or a.cues), a.cover] + ([a.audio] if a.audio else [])
    for p in need:
        if not p or not os.path.exists(p):
            raise SystemExit("Missing file: " + str(p))

    if a.json:
        # Timing + tracklist come ONLY from the JSON; bpm/key optionally joined
        # from --meta by track order. Edit the JSON and re-render — no re-matching.
        import json as _json
        data = _json.load(open(a.json, encoding="utf-8"))
        meta = {}
        if a.meta and os.path.exists(a.meta):
            with open(a.meta, encoding="utf-8-sig") as f:
                for i, r in enumerate(csv.DictReader(f), 1):
                    meta[i] = r
        tracks = []
        for r in data:
            m = meta.get(r.get("order"), {})
            tracks.append({"start": str(r.get("start_sec", "")), "artist": r.get("artist", ""),
                           "title": r.get("title", ""), "bpm": m.get("bpm", ""),
                           "song_key": m.get("song_key", ""), "camelot": m.get("camelot", ""),
                           "genres": r.get("genres") or m.get("genres", ""),
                           "release_date": r.get("release_date") or m.get("release_date", ""),
                           "show_date": r.get("show_date") or m.get("show_date", ""),
                           "show_venue": r.get("show_venue") or m.get("show_venue", "")})
    else:
        if not a.cues:
            raise SystemExit("Give --cues or --json.")
        with open(a.cues, encoding="utf-8-sig") as f:
            tracks = list(csv.DictReader(f))
    if not tracks:
        raise SystemExit("No tracks to render.")

    # parse + validate start times
    blanks = []
    for i, t in enumerate(tracks):
        t["_start"] = parse_start(t.get("start"))
        if t["_start"] is None:
            blanks.append(i + 1)
    if blanks:
        raise SystemExit("These track rows have a blank `start` — fill them first: %s"
                         % ", ".join(map(str, blanks)))
    starts = [t["_start"] for t in tracks]
    if starts != sorted(starts):
        raise SystemExit("Start times aren't in increasing order — check the cues.")

    if a.audio:
        total = ffprobe_duration(a.audio)
    elif a.duration:
        total = parse_start(a.duration)
        if total is None:
            raise SystemExit("Couldn't read --duration %r (use mm:ss, h:mm:ss or seconds)." % a.duration)
    else:
        # No mix yet: the last card needs an end, so fall back to that track's own
        # length from the cues. Better than an arbitrary tail — the deck then runs
        # for as long as the set actually should.
        last_ms = 0
        try: last_ms = int(float(tracks[-1].get("duration_ms") or 0))
        except Exception: pass
        total = starts[-1] + (last_ms / 1000.0 if last_ms else 180.0)
        print("No audio — deck runs to %s (last cue + its own length)." % fmt_clock(total))
    if starts[-1] >= total:
        raise SystemExit("Last track start (%s) is past the %s length (%s)."
                         % (fmt_clock(starts[-1]), "audio" if a.audio else "requested", fmt_clock(total)))
    ends = starts[1:] + [total]

    work = tempfile.mkdtemp(prefix="cwr_render_")
    bg = BACKDROPS[a.backdrop]()
    cover = rounded_cover(a.cover, 470, 34)
    print("Rendering %d track cards…" % len(tracks))
    cards = []
    for i, t in enumerate(tracks):
        nxt = None
        if i + 1 < len(tracks):
            n = tracks[i + 1]
            nxt = ("%s — %s" % (n.get("artist", ""), n.get("title", ""))).strip(" —")
        png = os.path.join(work, "card_%03d.png" % i)
        render_card(bg, cover, t, i + 1, len(tracks), a.ep, a.title, nxt, png,
                    progress=starts[i] / total if total else 0)
        dur = max(0.1, ends[i] - starts[i])
        cards.append([png, dur])

    # Bookends: the intro overlays the OPENING of the mix and the closing overlays
    # the ENDING, so the audio (the mix) is untouched and total runtime is exactly
    # the mix length. We just steal the first/last few seconds of card time.
    intro = outro = []
    if not a.no_bookends:
        bk_name = a.bookend_backdrop or ED.get("bookend_backdrop")
        bk_bg = BACKDROPS[bk_name]() if bk_name else bg
        cover_sm = rounded_cover(a.bookend_cover or a.cover, 300, 24)
        intro_secs = min(a.intro_secs, cards[0][1] * 0.7)          # don't eat a whole track
        outro_secs = min(a.outro_secs, cards[-1][1] * 0.7)
        drop = a.drop_date
        nxt_date = next_thursday(drop, a.next_date)
        # The extra hold is real screen time, so it comes out of the first card
        # too — otherwise the intro would run past its own slot and every cue
        # after it would sit late against the mix.
        hold = max(0.0, min(a.intro_hold, cards[0][1] * 0.7 - intro_secs))
        intro = build_intro(work, bk_bg, cover_sm, a.ep, a.mixed_by, drop, intro_secs, hold)
        outro = build_outro(work, bk_bg, cover_sm, nxt_date, outro_secs)
        cards[0][1] = max(0.1, cards[0][1] - intro_secs - hold)    # first card starts after intro
        cards[-1][1] = max(0.1, cards[-1][1] - outro_secs)         # last card ends before outro
        print("Intro %.1fs (+%.1fs hold) + closing %.1fs (overlaid on the mix — total unchanged)."
              % (intro_secs, hold, outro_secs))
    concat = list(intro) + [tuple(c) for c in cards] + list(outro)

    listfile = os.path.join(work, "list.txt")
    with open(listfile, "w", encoding="utf-8") as f:
        for png, dur in concat:
            f.write("file '%s'\n" % png.replace("\\", "/"))
            f.write("duration %.3f\n" % dur)
        f.write("file '%s'\n" % concat[-1][0].replace("\\", "/"))  # concat needs last repeated

    # Slides concat + audio. The progress bar is baked into each PNG (stepped per
    # track), so no per-frame video filter is needed — this renders an hour-long
    # mix in seconds instead of choking a per-pixel geq bar.
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    cmd = ["ffmpeg", "-y",
           "-f", "concat", "-safe", "0", "-i", listfile]
    if a.audio:
        cmd += ["-i", a.audio]
    cmd += ["-vf", "fps=30,format=yuv420p",
            "-c:v", "libx264", "-preset", "veryfast", "-crf", "20"]
    if a.audio:
        cmd += ["-c:a", "aac", "-b:a", a.abitrate]
    else:
        cmd += ["-an"]                     # slides only — no silent track to strip later
    cmd += ["-pix_fmt", "yuv420p", "-movflags", "+faststart"]
    if a.dry:
        cmd += ["-t", "1"]
    if a.audio:
        cmd += ["-shortest"]               # meaningless (and a warning) with no second input
    cmd += [a.out]
    print("Encoding video (total %s)…" % fmt_clock(total))
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        sys.stderr.write(r.stderr[-1600:])
        raise SystemExit("ffmpeg failed.")
    print("DONE ->", a.out)
    print("Upload that MP4 to YouTube, then paste the link into ✎ Details.")

if __name__ == "__main__":
    main()
