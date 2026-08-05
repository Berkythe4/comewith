#!/usr/bin/env python3
"""
make_thumbs.py — turn the full-size Elements cover art into upload-ready files.

    python Radio/render/make_thumbs.py --src "Radio/Elements-26/render/Cover Art"

The masters are 8334x8334 and 20-60 MB each, which nothing will accept:

    YouTube thumbnail   1280x720, 16:9, under 2 MB
    SoundCloud artwork  square, at least 800x800, under 2 MB

So each cover produces both. The YouTube one puts the WHOLE square over a
blurred, darkened zoom of itself: a 16:9 centre crop was tried first and it
beheaded the Come With mark, which is a tall triangle — the wordmark came out cut
in half and unreadable at thumbnail size. Hard black bars would have kept it but
wasted a third of an already small image.

Quality is stepped down until the file is under the limit rather than guessed at,
so a busier cover doesn't silently come out over budget.
"""
import argparse, os, sys
from PIL import Image, ImageEnhance, ImageFilter

try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

YT_LIMIT = 2 * 1024 * 1024          # YouTube's own thumbnail ceiling
SC_LIMIT = 2 * 1024 * 1024


def save_under(im, path, limit, q0=92):
    """Save JPEG, stepping quality down until it fits. Returns (bytes, quality)."""
    for q in range(q0, 40, -6):
        im.save(path, "JPEG", quality=q, optimize=True, progressive=True)
        n = os.path.getsize(path)
        if n <= limit:
            return n, q
    return os.path.getsize(path), q


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="folder of cover art")
    ap.add_argument("--out", help="where to write (default: <src>/upload)")
    a = ap.parse_args()
    out = a.out or os.path.join(a.src, "upload")
    os.makedirs(out, exist_ok=True)

    names = sorted(f for f in os.listdir(a.src)
                   if f.lower().endswith((".jpg", ".jpeg", ".png")))
    if not names:
        raise SystemExit("No images in %s" % a.src)

    for n in names:
        src = os.path.join(a.src, n)
        im = Image.open(src)
        mode, size, mb = im.mode, im.size, os.path.getsize(src) / 1024 / 1024
        im = im.convert("RGB")                      # the masters are CMYK
        stem = os.path.splitext(n)[0]

        # YouTube — the WHOLE square, centred over a blurred fill of itself.
        #
        # A 16:9 centre crop was the first attempt and it beheaded the mark: the
        # Come With triangle is tall, so cropping the top and bottom off a square
        # cut the wordmark in half and left it unreadable at thumbnail size. Hard
        # black bars would keep it but waste a third of an already small image.
        # Blurring a zoomed copy behind fills the frame and keeps the art intact.
        w, h = im.size
        bg = im.resize((1280, 1280), Image.LANCZOS).crop((0, 280, 1280, 1000))
        bg = bg.filter(ImageFilter.GaussianBlur(28))
        bg = ImageEnhance.Brightness(bg).enhance(0.72)   # sit it back behind the art
        fg = im.resize((720, 720), Image.LANCZOS)
        yt = bg.copy()
        yt.paste(fg, ((1280 - 720) // 2, 0))
        yp = os.path.join(out, "%s_yt_1280x720.jpg" % stem)
        yn, yq = save_under(yt, yp, YT_LIMIT)

        # SoundCloud — square, 1600 is comfortably over their 800 minimum.
        s = min(im.size)
        sq = im.crop(((w - s) // 2, (h - s) // 2, (w - s) // 2 + s, (h - s) // 2 + s))
        sq = sq.resize((1600, 1600), Image.LANCZOS)
        sp = os.path.join(out, "%s_sc_1600.jpg" % stem)
        sn, sq_ = save_under(sq, sp, SC_LIMIT)

        print("%-22s %sx%s %s %5.1f MB  ->  yt %4d KB (q%d) · sc %4d KB (q%d)"
              % (n, size[0], size[1], mode, mb, yn // 1024, yq, sn // 1024, sq_))

    print("\nwrote %d file(s) to %s" % (len(names) * 2, out))


if __name__ == "__main__":
    main()
