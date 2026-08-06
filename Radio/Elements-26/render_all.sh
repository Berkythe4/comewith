#!/usr/bin/env bash
# Re-render all four Elements episodes.
#
# Three things changed on 6 Aug:
#
#   --bookend-cover  the ETHER mark on the first AND last slide. The bookends
#                    belong to the RUN, not the night — every episode opens and
#                    closes on the same art while the track cards stay in that
#                    DJ's own element. (The closing slide never drew a cover at
#                    all before; draw_outro took one and ignored it.)
#   --next-date      the date of the NEXT episode. The default was drop+7, a
#                    weekly-show assumption that put "AUG 13" on Ep 1 when Ep 2
#                    goes out the following afternoon.
#   --next-label     Ep 4 only. The run does not continue after Sunday; it hands
#                    back to the NYC show on the 20th, and a pill reading "THE
#                    RUN CONTINUES" there sends people back for an episode that
#                    does not exist.
#
# Written out four times rather than looped over a table. The filenames do not
# follow one pattern — Ep 1's audio is CWR_Elements_Day1_Berky.wav while its
# video is CWR_ElementsEp1_Berky.mp4 — and a clever loop that derives one from
# the other writes a correct-looking render to the wrong path.
#
# Serial on purpose. Two ffmpeg jobs writing at once once produced a file with a
# plausible size and duration whose NAL units were shredded and whose audio was
# -91 dB. Renders are ~15 minutes each; the set takes about an hour.
set -euo pipefail
cd "$(dirname "$0")/../.."

R="Radio/Elements-26/render"
AV="$R/Audio_Video_Final"
ART="$R/Cover Art"
ETHER="$ART/ETHER_JANELLE.JPG"
RE="python Radio/render/render_episode.py"
COMMON=(--edition elements --bookend-cover "$ETHER" --title "Come With Elements Radio")

echo "=== EP 1 · WATER · Berky · Thu 6 Aug ==="
$RE "${COMMON[@]}" \
  --cues "$R/Elements_Ep1_cues.csv" \
  --audio "$AV/CWR_Elements_Day1_Berky.wav" \
  --backdrop water --cover "$ART/WATER_KEITH.JPG" \
  --out "$AV/CWR_ElementsEp1_Berky.mp4" \
  --ep "EP 1" --mixed-by "Berky" \
  --drop-date 2026-08-06 --next-date 2026-08-07

echo "=== EP 2 · FIRE · KRNeY · Fri 7 Aug ==="
$RE "${COMMON[@]}" \
  --cues "$R/Elements_Ep2_KRNeY_cues.csv" \
  --audio "$AV/CWR_ElementsEp2_KRNeY.wav" \
  --backdrop fire --cover "$ART/FIRE_MARTIN.JPG" \
  --out "$AV/CWR_ElementsEp2_KRNeY.mp4" \
  --ep "EP 2" --mixed-by "KRNeY" \
  --drop-date 2026-08-07 --next-date 2026-08-08

echo "=== EP 3 · EARTH · Henry · Sat 8 Aug ==="
$RE "${COMMON[@]}" \
  --cues "$R/Elements_Ep3_Henry_cues.csv" \
  --audio "$AV/CWR_ElementsEp3_Henry.wav" \
  --backdrop earth --cover "$ART/EARTH_HENRY.JPG" \
  --out "$AV/CWR_ElementsEp3_Henry.mp4" \
  --ep "EP 3" --mixed-by "Henry" \
  --drop-date 2026-08-08 --next-date 2026-08-09

echo "=== EP 4 · AIR · 32LVS · Sun 9 Aug (finale) ==="
$RE "${COMMON[@]}" \
  --cues "$R/Elements_Ep4_32LVS_cues.csv" \
  --audio "$AV/CWR_ElementsEp4_32LVS.mp3" \
  --backdrop air --cover "$ART/AIR_32LVS.JPG" \
  --out "$AV/CWR_ElementsEp4_32LVS.mp4" \
  --ep "EP 4" --mixed-by "32LVS" \
  --drop-date 2026-08-09 --next-date 2026-08-20 \
  --next-label "BACK TO NYC RADIO"

echo "All four rendered."
