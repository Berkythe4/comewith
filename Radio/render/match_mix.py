#!/usr/bin/env python3
"""
match_mix.py — recover a DJ mix's tracklist timestamps by AUDIO-MATCHING each
owned source track against the recorded mix. Pure local processing; nothing is
uploaded. (PHASE 1 of the video build.)

How it works
------------
1. ffmpeg decodes the mix and every source file to mono 11025 Hz (handles
   WAV/AIFF/FLAC/MP3 — no librosa/soundfile needed).
2. Each signal becomes a log-mel spectrogram (robust to EQ/level, unlike raw
   samples), z-scored per band.
3. For each source track we take its highest-energy QUERY_SECONDS window (the
   part actually played, not the intro that's buried under the previous track)
   and slide it across the mix via FFT cross-correlation, summed over mel bands
   and length-normalised → a cosine-similarity score at every offset.
4. DJ tracks are tempo-shifted and time-stretched. If the straight match is weak
   we re-try the query time-stretched from -8% to +8% (1% steps) and keep the
   best. The winning stretch is reported as tempo_shift_pct.
5. The track's START in the mix = matched offset − (query start in the track),
   tempo-adjusted.

Outputs (into Radio/render/)
   tracklist.json          — [{order, artist, title, start_sec, start_hms,
                               confidence, tempo_shift_pct, source_file}]
   youtube_chapters.txt     — "HH:MM:SS Artist – Title", first line 00:00:00
   match_report.md          — matched / unmatched / low-confidence / overlaps

Usage
   python Radio/render/match_mix.py \
     --mix "Radio/Video/01 CWR_Ep1_Final.wav" \
     --tracks "G:/Radio Station/Week 1/Tracks" \
     [--mix-seconds 600]   # analyse only the first N s of the mix (test mode)
     [--quick]             # skip the tempo search (fast first pass)

Timestamps are ESTIMATES for your review — hand-correct start_sec in
tracklist.json and the video re-renders without re-matching.
"""
import argparse, glob, hashlib, json, os, subprocess, sys, tempfile
import numpy as np
from scipy.io import wavfile
from scipy.signal import stft, fftconvolve
try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass

SR = 11025
N_FFT = 2048
HOP = 512               # ~46 ms/frame
N_MELS = 40
FPS = SR / HOP          # feature frames per second
QUERY_SECONDS = 60
AUDIO_EXT = (".wav", ".aiff", ".aif", ".flac", ".mp3", ".m4a")

# ---- decode + features ------------------------------------------------------
def decode_mono(path, cache, seconds=None):
    """ffmpeg → mono 11025 16-bit wav → numpy float32 [-1,1]. Cached by path+mtime."""
    key = hashlib.md5((path + str(os.path.getmtime(path)) + str(seconds)).encode()).hexdigest()[:16]
    wav = os.path.join(cache, key + ".wav")
    if not os.path.exists(wav):
        cmd = ["ffmpeg", "-y", "-v", "error"]
        if seconds:
            cmd += ["-t", str(seconds)]
        cmd += ["-i", path, "-ac", "1", "-ar", str(SR), "-sample_fmt", "s16", wav]
        subprocess.run(cmd, check=True)
    sr, x = wavfile.read(wav)
    if x.dtype == np.int16:
        x = x.astype(np.float32) / 32768.0
    else:
        x = x.astype(np.float32)
    return x

_MEL = None
def mel_filter():
    global _MEL
    if _MEL is not None:
        return _MEL
    # simple mel filterbank over the STFT bins
    n_bins = N_FFT // 2 + 1
    fmax = SR / 2
    def hz2mel(f): return 2595 * np.log10(1 + f / 700)
    def mel2hz(m): return 700 * (10 ** (m / 2595) - 1)
    mpts = np.linspace(hz2mel(40), hz2mel(fmax), N_MELS + 2)
    hz = mel2hz(mpts)
    bins = np.floor((N_FFT + 1) * hz / SR).astype(int)
    fb = np.zeros((N_MELS, n_bins), np.float32)
    for m in range(1, N_MELS + 1):
        l, c, r = bins[m - 1], bins[m], bins[m + 1]
        if c == l: c += 1
        if r == c: r += 1
        for k in range(l, c):
            if 0 <= k < n_bins: fb[m - 1, k] = (k - l) / max(1, c - l)
        for k in range(c, r):
            if 0 <= k < n_bins: fb[m - 1, k] = (r - k) / max(1, r - c)
    _MEL = fb
    return fb

def logmel(x):
    f, t, Z = stft(x, fs=SR, nperseg=N_FFT, noverlap=N_FFT - HOP, boundary=None)
    P = np.abs(Z) ** 2
    M = mel_filter() @ P
    return np.log1p(M).astype(np.float32)          # [N_MELS, frames]

def zscore(F):
    mu = F.mean(axis=1, keepdims=True)
    sd = F.std(axis=1, keepdims=True) + 1e-6
    return (F - mu) / sd

# ---- matching ---------------------------------------------------------------
def best_query_window(Ft, win_frames):
    """Pick the highest-energy contiguous window of the track feature = the part
    most likely actually played (not the intro)."""
    if Ft.shape[1] <= win_frames:
        return 0, Ft
    energy = (Ft ** 2).sum(axis=0)
    csum = np.concatenate([[0], np.cumsum(energy)])
    win_e = csum[win_frames:] - csum[:-win_frames]
    s = int(np.argmax(win_e))
    return s, Ft[:, s:s + win_frames]

def time_stretch(F, factor):
    """Resample the feature along time by `factor` (>1 = longer/slower)."""
    n = F.shape[1]
    m = max(1, int(round(n * factor)))
    idx = np.linspace(0, n - 1, m)
    lo = np.floor(idx).astype(int); hi = np.minimum(lo + 1, n - 1); frac = idx - lo
    return F[:, lo] * (1 - frac) + F[:, hi] * frac

def ncc(Mz, Qz):
    """Normalised cross-correlation of query Qz over mix Mz (both z-scored per
    band). Returns (best_lag_frames, score in ~[-1,1])."""
    B, Lm = Mz.shape
    Lq = Qz.shape[1]
    if Lq >= Lm:
        return 0, 0.0
    corr = np.zeros(Lm - Lq + 1, np.float32)
    for b in range(B):
        corr += fftconvolve(Mz[b], Qz[b][::-1], mode="valid")
    # local energy of the mix window per lag (summed over bands), for normalisation
    e = (Mz ** 2).sum(axis=0)
    csum = np.concatenate([[0], np.cumsum(e)])
    win_e = csum[Lq:] - csum[:len(corr)]
    qn = float((Qz ** 2).sum())
    denom = np.sqrt(np.maximum(win_e, 1e-6) * max(qn, 1e-6))
    score = corr / denom
    lag = int(np.argmax(score))
    return lag, float(score[lag])

def score_curve(Mz, Qz):
    """Normalised cross-correlation curve of query Qz over mix Mz (both z-scored)."""
    B, Lm = Mz.shape
    Lq = Qz.shape[1]
    if Lq >= Lm:
        return None
    corr = np.zeros(Lm - Lq + 1, np.float32)
    for b in range(B):
        corr += fftconvolve(Mz[b], Qz[b][::-1], mode="valid")
    e = (Mz ** 2).sum(axis=0)
    csum = np.concatenate([[0], np.cumsum(e)])
    win_e = csum[Lq:] - csum[:len(corr)]
    qn = float((Qz ** 2).sum())
    return corr / np.sqrt(np.maximum(win_e, 1e-6) * max(qn, 1e-6))

def match_track(Mz, Ft, tempos, min_start=0.0, max_start=None):
    """Best alignment of track feature Ft in mix Mz across tempos, restricted to
    starts in [min_start, max_start] — the ordering constraint that kills the
    false positives blind matching produces."""
    win = int(QUERY_SECONDS * FPS)
    q0_frames, Q = best_query_window(Ft, win)
    q0_sec = q0_frames / FPS
    best = {"score": -1, "start": min_start, "tempo": 0}
    for s in tempos:
        factor = 1.0 / (1 + s / 100.0)     # played +s% faster → source time compressed
        Qz = zscore(time_stretch(Q, factor)) if s != 0 else zscore(Q)
        sc = score_curve(Mz, Qz)
        if sc is None:
            continue
        starts = np.arange(len(sc)) / FPS - q0_sec * factor
        mask = starts >= max(0.0, min_start)
        if max_start is not None:
            mask &= starts <= max_start
        if not mask.any():
            continue
        idx = np.flatnonzero(mask)
        j = idx[int(np.argmax(sc[idx]))]
        if sc[j] > best["score"]:
            best = {"score": float(sc[j]), "start": float(starts[j]), "tempo": s}
    return max(0.0, best["start"]), best["score"], best["tempo"]

# ---- map a known tracklist (artist/title, in order) to source files ----------
def _norm(s):
    s = str(s or "").lower()
    for junk in ("(original mix)", "(extended mix)", "(extended)", "(feat", "feat.", "original mix"):
        s = s.replace(junk, " ")
    out = []
    for ch in s:
        out.append(ch if ch.isalnum() else " ")
    return " ".join("".join(out).split())

def map_files(tracklist, files):
    """For each {artist,title} pick the best-matching source file by token overlap."""
    fnorm = [(f, set(_norm(os.path.basename(f)).split())) for f in files]
    mapping = []
    for t in tracklist:
        want = set(_norm((t["artist"] + " " + t["title"])).split())
        best, bestscore = None, 0.0
        for f, toks in fnorm:
            if not toks:
                continue
            inter = len(want & toks)
            sc = inter / max(1, len(want))
            if sc > bestscore:
                best, bestscore = f, sc
        mapping.append(best if bestscore >= 0.5 else None)
    return mapping

# ---- io helpers -------------------------------------------------------------
def parse_name(fn):
    base = os.path.splitext(os.path.basename(fn))[0]
    if " - " in base:
        a, t = base.split(" - ", 1)
        return a.strip(), t.strip()
    return "", base.strip()

def hms(sec):
    sec = int(round(sec)); return "%02d:%02d:%02d" % (sec // 3600, (sec % 3600) // 60, sec % 60)

def read_tracklist(path):
    import csv
    with open(path, encoding="utf-8-sig") as f:
        return [{"artist": r.get("artist", "").strip(), "title": r.get("title", "").strip()}
                for r in csv.DictReader(f)]

def build_score_grid(Mz, Ft, tempos):
    """For one track: best correlation at each possible START frame in the mix,
    taking the max over the tempo search. Returns (grid[T], tempo_at[T])."""
    T = Mz.shape[1]
    grid = np.full(T, -2.0, np.float32)
    tg = np.zeros(T, np.int16)
    win = int(QUERY_SECONDS * FPS)
    q0f, Q = best_query_window(Ft, win)
    for s in tempos:
        factor = 1.0 / (1 + s / 100.0)
        Qz = zscore(time_stretch(Q, factor)) if s != 0 else zscore(Q)
        sc = score_curve(Mz, Qz)
        if sc is None:
            continue
        lags = np.arange(len(sc))
        starts = np.round(lags - q0f * factor).astype(int)   # lag → track-start frame
        ok = (starts >= 0) & (starts < T)
        tmp = np.full(T, -2.0, np.float32)
        np.maximum.at(tmp, starts[ok], sc[ok])
        imp = tmp > grid
        grid[imp] = tmp[imp]; tg[imp] = s
    return grid, tg

def dp_align(grids, gap_bins):
    """Global monotonic alignment: choose one start per track, strictly increasing
    and >= gap_bins apart, maximising the summed correlation. O(N·T) via prefix
    max — a single bad peak can't cascade because the WHOLE path is scored."""
    N = len(grids); T = len(grids[0])
    best = grids[0].astype(np.float64).copy()
    back = [np.zeros(T, np.int32) for _ in range(N)]
    for i in range(1, N):
        prev = best
        pm_val = np.maximum.accumulate(prev)
        pm_arg = np.zeros(T, np.int32); cur = 0
        for t in range(T):
            if prev[t] > prev[cur]:
                cur = t
            pm_arg[t] = cur
        sv = np.full(T, -1e18); sa = np.zeros(T, np.int32)
        if gap_bins < T:
            sv[gap_bins:] = pm_val[:T - gap_bins]
            sa[gap_bins:] = pm_arg[:T - gap_bins]
        best = grids[i] + sv
        back[i] = sa
    end = int(np.argmax(best)); path = [end]
    for i in range(N - 1, 0, -1):
        end = int(back[i][end]); path.append(end)
    path.reverse()
    return path

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mix", required=True)
    ap.add_argument("--tracks", required=True)
    ap.add_argument("--tracklist", help="ordered cues CSV (artist,title) — enables ordered/constrained matching (recommended)")
    ap.add_argument("--mix-seconds", type=int, default=None)
    ap.add_argument("--min-gap", type=float, default=25.0, help="min seconds between consecutive track starts")
    ap.add_argument("--quick", action="store_true", help="skip tempo search")
    ap.add_argument("--outdir", default=os.path.dirname(os.path.abspath(__file__)))
    a = ap.parse_args()

    files = sorted([f for f in glob.glob(os.path.join(a.tracks, "*"))
                    if f.lower().endswith(AUDIO_EXT)])
    if not files:
        raise SystemExit("No audio files in " + a.tracks)
    tempos = [0] if a.quick else list(range(-8, 9))

    cache = os.path.join(tempfile.gettempdir(), "cwr_match_cache")
    os.makedirs(cache, exist_ok=True)

    print("Decoding + analysing the mix%s…" % (" (first %ds)" % a.mix_seconds if a.mix_seconds else ""))
    mix = decode_mono(a.mix, cache, seconds=a.mix_seconds)
    mix_len = len(mix) / SR
    Mz = zscore(logmel(mix))
    HIGH, LOW = 0.30, 0.18

    if a.tracklist:
        # ---- ORDERED mode: global monotonic alignment (Viterbi/DP) ------------
        tl = read_tracklist(a.tracklist)
        mp = map_files(tl, files)
        have = [(i, t, f) for i, (t, f) in enumerate(zip(tl, mp)) if f]
        missing = [(i, t) for i, (t, f) in enumerate(zip(tl, mp)) if not f]
        print("Matching %d tracks (%d with source, %d missing) via global alignment…\n"
              % (len(tl), len(have), len(missing)))
        grids, tgs = [], []
        for i, t, f in have:
            g, tg = build_score_grid(Mz, logmel(decode_mono(f, cache)), tempos)
            grids.append(g); tgs.append(tg)
            print("  built curve: %s - %s" % (t["artist"], t["title"][:40]))
        gap_bins = int(a.min_gap * FPS)
        path = dp_align(grids, gap_bins)

        by_idx = {}
        for k, (i, t, f) in enumerate(have):
            bin_ = path[k]
            by_idx[i] = {"artist": t["artist"], "title": t["title"],
                         "start_sec": round(bin_ / FPS, 2), "confidence": round(float(grids[k][bin_]), 3),
                         "tempo_shift_pct": int(tgs[k][bin_]), "source_file": os.path.basename(f)}
        for i, t in missing:                        # interpolate a missing source between neighbours
            lo = max([by_idx[j]["start_sec"] for j in by_idx if j < i], default=0)
            his = [by_idx[j]["start_sec"] for j in by_idx if j > i]
            hi = min(his) if his else lo + 120
            by_idx[i] = {"artist": t["artist"], "title": t["title"], "start_sec": round((lo + hi) / 2, 2),
                         "confidence": 0.0, "tempo_shift_pct": 0, "source_file": None, "interpolated": True}
        ordered = [by_idx[i] for i in sorted(by_idx)]
        for n, r in enumerate(ordered, 1):
            r["order"] = n; r["start_hms"] = hms(r["start_sec"])
        matched = ordered; unmatched = []
        for r in ordered:
            tag = "INTERP" if r.get("interpolated") else ("%.2f" % r["confidence"])
            print("  %2d  %-9s conf %-7s t%+d%%  %s - %s"
                  % (r["order"], r["start_hms"], tag, r["tempo_shift_pct"], r["artist"], r["title"][:34]))
        rows = ordered
    else:
        # ---- BLIND mode (fallback) --------------------------------------------
        print("Matching %d source files blind (tempo steps: %d)…\n" % (len(files), len(tempos)))
        rows = []
        for i, f in enumerate(files, 1):
            artist, title = parse_name(f)
            Ft = logmel(decode_mono(f, cache))
            start, score, tempo = match_track(Mz, Ft, tempos)
            rows.append({"artist": artist, "title": title, "start_sec": round(start, 2),
                         "confidence": round(score, 3), "tempo_shift_pct": tempo,
                         "source_file": os.path.basename(f)})
            print("  %2d/%2d  %-6.1f  conf %.2f  t%+d%%  %s"
                  % (i, len(files), start, score, tempo, os.path.basename(f)[:52]))
        matched = [r for r in rows if r["confidence"] >= LOW and r["start_sec"] < mix_len - 5]
        matched.sort(key=lambda r: r["start_sec"])
        for n, r in enumerate(matched, 1):
            r["order"] = n; r["start_hms"] = hms(r["start_sec"])

    os.makedirs(a.outdir, exist_ok=True)
    with open(os.path.join(a.outdir, "tracklist.json"), "w", encoding="utf-8") as fp:
        json.dump(matched, fp, indent=2, ensure_ascii=False)

    # chapters (first forced to 0)
    with open(os.path.join(a.outdir, "youtube_chapters.txt"), "w", encoding="utf-8") as fp:
        for n, r in enumerate(matched):
            stamp = "00:00:00" if n == 0 else r["start_hms"]
            fp.write("%s %s – %s\n" % (stamp, r["artist"], r["title"]))

    # overlaps: consecutive starts < 45s apart = suspicious for a mix
    overlaps = [(matched[i], matched[i + 1]) for i in range(len(matched) - 1)
                if matched[i + 1]["start_sec"] - matched[i]["start_sec"] < 45]
    low = [r for r in matched if r["confidence"] < HIGH]
    unmatched = [r for r in rows if r not in matched]

    with open(os.path.join(a.outdir, "match_report.md"), "w", encoding="utf-8") as fp:
        fp.write("# Mix match report\n\n")
        fp.write("Mix: `%s` — analysed %s of %.1f min.\n\n"
                 % (os.path.basename(a.mix), ("first %.1f min" % (a.mix_seconds/60)) if a.mix_seconds else "all", mix_len/60))
        fp.write("**%d of %d source files placed.** Tempo search: %s.\n\n"
                 % (len(matched), len(files), "off (quick)" if a.quick else "-8%..+8%"))
        fp.write("## Placed tracks\n\n| # | Start | Conf | Tempo | Track |\n|--|--|--|--|--|\n")
        for r in matched:
            flag = " ⚠️" if r["confidence"] < HIGH else ""
            fp.write("| %d | %s | %.2f%s | %+d%% | %s – %s |\n"
                     % (r["order"], r["start_hms"], r["confidence"], flag, r["tempo_shift_pct"], r["artist"], r["title"]))
        if low:
            fp.write("\n## ⚠️ Low confidence (<%.2f) — check these\n\n" % HIGH)
            for r in low:
                fp.write("- **%s** at %s (conf %.2f) — verify the start by ear.\n" % (r["title"], r["start_hms"], r["confidence"]))
        if overlaps:
            fp.write("\n## ⚠️ Suspiciously close (<45s apart)\n\n")
            for a1, b1 in overlaps:
                fp.write("- %s (%s) then %s (%s) — %.0fs apart. One may be a mis-match.\n"
                         % (a1["title"], a1["start_hms"], b1["title"], b1["start_hms"], b1["start_sec"]-a1["start_sec"]))
        if unmatched:
            fp.write("\n## Not placed\n\n")
            for r in sorted(unmatched, key=lambda r: -r["confidence"]):
                if r.get("source_file"):
                    fp.write("- %s — best conf %.2f. Not in this window / not in the mix.\n" % (r["source_file"], r["confidence"]))
                else:
                    fp.write("- **%s – %s** — SOURCE FILE MISSING (was track %s in the list). Interpolate its start from neighbours or add the source.\n"
                             % (r["artist"], r["title"], r.get("order_hint", "?")))
        fp.write("\n---\nEstimates for review. Hand-correct `start_sec` in tracklist.json; the video re-renders from that without re-matching.\n")

    print("\nWrote tracklist.json, youtube_chapters.txt, match_report.md to %s" % a.outdir)
    print("Placed %d/%d. %d low-confidence, %d suspiciously-close, %d unplaced."
          % (len(matched), len(files), len(low), len(overlaps), len(unmatched)))

if __name__ == "__main__":
    main()
