"""List slides that have a title but no body (no table, chart, content card or
picture, and no real text) - what a leftover page of a stripped table looks like.

    python scan_empty_bodies.py "deck.pptx"
"""
import re
import sys
from pptx import Presentation

PX = 9525


def _type(sh):
    try:
        return sh.shape_type
    except NotImplementedError:
        return None


def walk(shs):
    for sh in shs:
        if _type(sh) == 6:
            yield from walk(sh.shapes)
        else:
            yield sh


def empty_body(slide) -> bool:
    chars = 0
    for sh in walk(slide.shapes):
        if getattr(sh, "has_table", False) and sh.has_table:
            return False
        if getattr(sh, "has_chart", False) and sh.has_chart:
            return False
        t = _type(sh)
        if t == 1 and sh.name.startswith("Rounded Rectangle") \
                and sh.width / PX > 300:
            return False                      # a content card
        if t == 13 and sh.width / PX > 200 and sh.top / PX > 100:
            return False                      # a real picture, not the logo
        if sh.has_text_frame and sh.top is not None and sh.top / PX > 100 \
                and sh.top / PX < 690:
            x = sh.text_frame.text.strip()
            if "All Rights Reserved" not in x and "Source" not in x[:7]:
                chars += len(x)
    return chars < 60


def title(slide):
    ts = [sh.text_frame.text.strip() for sh in walk(slide.shapes)
          if sh.has_text_frame and sh.text_frame.text.strip()
          and sh.top is not None and sh.top / PX < 70]
    return max(ts, key=len) if ts else ""


if __name__ == "__main__":
    prs = Presentation(sys.argv[1])
    n = 0
    for i, s in enumerate(prs.slides, 1):
        if empty_body(s):
            n += 1
            print(i, title(s)[:100].encode("ascii", "replace").decode())
    print(n, "slides with no body of", len(prs.slides))
