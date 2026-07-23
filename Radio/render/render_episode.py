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

`start` in the cues CSV is mm:ss (e.g. 4:12) or plain seconds. Any blank starts
are reported and abort, so you never render a half-timed video by accident.

  --dry           render cards + a 1s preview only (fast sanity check)
  --title TEXT    show name (default "Come With Radio")
"""
import argparse, csv, os, re, subprocess, sys, tempfile
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

def render_card(bg, cover, track, idx, ntracks, ep_label, title_text, nxt, out_png):
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
    artist = truncate(d, track["artist"] or "—", F_disp(132), mw)
    d.text((mx, cy + 58), artist, font=F_disp(132), fill=CREAM)
    ttitle = truncate(d, track["title"] or "", F_body(58), mw)
    d.text((mx, cy + 214), ttitle, font=F_body(58), fill=DIM)

    # chips: BPM, key
    chx = mx; chy = cy + 300
    if track.get("bpm"):
        chx = chip(d, chx, chy, "%s BPM" % track["bpm"], F_mono(34)) + 22
    keytxt = " · ".join([x for x in [track.get("camelot"), track.get("song_key")] if x])
    if keytxt:
        chip(d, chx, chy, keytxt, F_mono(34), accent=True)

    # progress track (empty) — a dim rail; ffmpeg overlays the moving lime fill.
    d.rounded_rectangle([BAR_X, BAR_Y, BAR_X + BAR_W, BAR_Y + BAR_H],
                        radius=BAR_H // 2, fill=(48, 38, 58))
    # up next
    up = ("UP NEXT — %s" % nxt) if nxt else "LAST TRACK"
    d.text((BAR_X, BAR_Y + 40), up, font=F_mono(30), fill=FAINT)
    site = "comewith.org"
    sw = d.textlength(site, font=F_mono(30))
    d.text((W - BAR_X - sw, BAR_Y + 40), site, font=F_mono(30), fill=(179, 167, 184))

    im.convert("RGB").save(out_png)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cues", required=True)
    ap.add_argument("--audio", required=True)
    ap.add_argument("--cover", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--ep", default="EP 1")
    ap.add_argument("--title", default="Come With Radio")
    ap.add_argument("--dry", action="store_true", help="cards + 1s preview only")
    a = ap.parse_args()

    for p in (a.cues, a.audio, a.cover):
        if not os.path.exists(p):
            raise SystemExit("Missing file: " + p)

    with open(a.cues, encoding="utf-8-sig") as f:
        tracks = list(csv.DictReader(f))
    if not tracks:
        raise SystemExit("Cues file has no rows.")

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

    total = ffprobe_duration(a.audio)
    if starts[-1] >= total:
        raise SystemExit("Last track start (%s) is past the audio length (%s)."
                         % (fmt_clock(starts[-1]), fmt_clock(total)))
    ends = starts[1:] + [total]

    work = tempfile.mkdtemp(prefix="cwr_render_")
    bg = make_background()
    cover = rounded_cover(a.cover, 470, 34)
    print("Rendering %d track cards…" % len(tracks))
    concat = []
    for i, t in enumerate(tracks):
        nxt = None
        if i + 1 < len(tracks):
            n = tracks[i + 1]
            nxt = ("%s — %s" % (n.get("artist", ""), n.get("title", ""))).strip(" —")
        png = os.path.join(work, "card_%03d.png" % i)
        render_card(bg, cover, t, i + 1, len(tracks), a.ep, a.title, nxt, png)
        dur = max(0.1, ends[i] - starts[i])
        concat.append((png, dur))

    listfile = os.path.join(work, "list.txt")
    with open(listfile, "w", encoding="utf-8") as f:
        for png, dur in concat:
            f.write("file '%s'\n" % png.replace("\\", "/"))
            f.write("duration %.3f\n" % dur)
        f.write("file '%s'\n" % concat[-1][0].replace("\\", "/"))  # concat needs last repeated

    # Moving progress bar: a lime source (BAR_W x BAR_H) whose ALPHA is gated by
    # time via geq — opaque for x < BAR_W * (t/total), transparent after — then
    # overlaid on the cards. drawbox can't do this (its `t` is thickness, not
    # time), so this geq-alpha overlay is the reliable route.
    bar_src = "color=c=0x%02x%02x%02x:s=%dx%d:d=%.3f:r=30" % (
        LIME[0], LIME[1], LIME[2], BAR_W, BAR_H, total)
    # geq needs r/g/b (lime) alongside the time-gated alpha, or it errors.
    geq_a = "if(lt(X\\,%d*min(T/%.3f\\,1))\\,255\\,0)" % (BAR_W, total)
    fc = ("[2:v]format=rgba,geq=r='%d':g='%d':b='%d':a='%s'[bar];"
          "[0:v]fps=30,format=yuv420p[base];"
          "[base][bar]overlay=%d:%d:format=auto,format=yuv420p[v]"
          % (LIME[0], LIME[1], LIME[2], geq_a, BAR_X, BAR_Y))

    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    cmd = ["ffmpeg", "-y",
           "-f", "concat", "-safe", "0", "-i", listfile,
           "-i", a.audio,
           "-f", "lavfi", "-i", bar_src,
           "-filter_complex", fc,
           "-map", "[v]", "-map", "1:a",
           "-c:v", "libx264", "-preset", "veryfast", "-crf", "20",
           "-c:a", "aac", "-b:a", "192k",
           "-pix_fmt", "yuv420p", "-movflags", "+faststart"]
    if a.dry:
        cmd += ["-t", "1"]
    cmd += ["-shortest", a.out]
    print("Encoding video (total %s)…" % fmt_clock(total))
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        sys.stderr.write(r.stderr[-1600:])
        raise SystemExit("ffmpeg failed.")
    print("DONE ->", a.out)
    print("Upload that MP4 to YouTube, then paste the link into ✎ Details.")

if __name__ == "__main__":
    main()
