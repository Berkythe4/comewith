# -*- coding: utf-8 -*-
"""Field repairs shared by every reader of a Rekordbox export.

There are two parsers for these files -- cues_from_rekordbox.py (the render
path) and apply_station_from_rekordbox.py (the database path) -- plus a third
in dashboard.html. They must agree about what a title and an artist ARE, or the
video, the episode page and the fill-in sheet print three different tracklists
for the same set. The parsing stays where it is; only these repairs are shared,
because they are the part that was silently diverging.

Every rule here comes from a real Ep 4 row, and each is narrow on purpose:

  Ep4 #2   Take Care (Extended Revisit) (Extended Revisit)
  Ep4 #13  I Want My Freedom (feat. Hero Baldwin) (feat. Hero Baldwin)
           Rekordbox prints Mix Name into Track Title while the tag already
           carries it, so the qualifier lands twice.

  Ep4 #16  Lane 8, Kasablanca >
           A dangling ">" from whatever wrote the tag. Not part of a name.

  Ep4 #15  MMXX ? XII (Kolsch Remix [Extended])
           A dash that lost its encoding upstream -- the Album field on the same
           row carries the identical corruption, which is what rules out the
           character being real. Only a FREE-STANDING "?" between two words is
           touched: a genuine "?" ends a question and sits against the word
           before it, so it can never match.
"""
import re

_DUPE_QUALIFIER = re.compile(r"\s*(\([^()]+\))\s*\1(?=\s|$)")
_DANGLING = re.compile(r"\s*[>»]\s*$")
_LOST_DASH = re.compile(r"(?<=\w) \? (?=\w)")


def clean_title(s):
    s = _DUPE_QUALIFIER.sub(r" \1", str(s or "")).strip()
    return _LOST_DASH.sub(" - ", s)


def clean_artist(s):
    return _DANGLING.sub("", str(s or "")).strip()
