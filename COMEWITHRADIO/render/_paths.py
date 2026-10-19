#!/usr/bin/env python3
"""Where the radio folder and an episode's folder live.

One place, because several tools need the same answer and they must never
disagree about it.

The radio folder was `Radio/` until 2026-10-19 and is `COMEWITHRADIO/` in the
repo now (Keith renamed it on purpose). The tool packet handed to guest DJs
(package_tool.py) still unpacks as `Radio/`, so both names are accepted, the
new one first.

Episode folders were called `Week N` until 2026-08-26 and are `Episode N` now;
the old name is still accepted so a folder that never got renamed keeps
working instead of failing with "there is no folder".
"""
import os

RADIO_NAMES = ("COMEWITHRADIO", "Radio")
NAMES = ("Episode %s", "Week %s")


def radio_dir(root):
    """root/COMEWITHRADIO (preferred) or root/Radio. The first that exists,
    else the preferred name so a caller creating it gets the current one."""
    for n in RADIO_NAMES:
        p = os.path.join(root, n)
        if os.path.isdir(p):
            return p
    return os.path.join(root, RADIO_NAMES[0])


def episode_dir(root, n, must_exist=True):
    """<radio>/Episode N (preferred) or <radio>/Week N (legacy). None if neither."""
    base = radio_dir(root)
    for pat in NAMES:
        p = os.path.join(base, pat % n)
        if os.path.isdir(p):
            return p
    return None if must_exist else os.path.join(base, NAMES[0] % n)


def episode_dir_or_die(root, n):
    p = episode_dir(root, n)
    if not p:
        base = os.path.basename(radio_dir(root))
        raise SystemExit("No folder for episode %s — looked for %s" % (
            n, " and ".join(os.path.join(base, pat % n) for pat in NAMES)))
    return p
