#!/usr/bin/env python3
"""Where an episode's folder lives.

One place, because four tools need the same answer and they must never disagree
about it. The folders were called `Week N` until 2026-08-26 and are `Episode N`
now; the old name is still accepted so a folder that never got renamed keeps
working instead of failing with "there is no folder".
"""
import os

NAMES = ("Episode %s", "Week %s")


def episode_dir(root, n, must_exist=True):
    """Radio/Episode N (preferred) or Radio/Week N (legacy). None if neither."""
    for pat in NAMES:
        p = os.path.join(root, "Radio", pat % n)
        if os.path.isdir(p):
            return p
    return None if must_exist else os.path.join(root, "Radio", NAMES[0] % n)


def episode_dir_or_die(root, n):
    p = episode_dir(root, n)
    if not p:
        raise SystemExit("No folder for episode %s — looked for %s" % (
            n, " and ".join(os.path.join("Radio", pat % n) for pat in NAMES)))
    return p
