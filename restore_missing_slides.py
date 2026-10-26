"""Put back EVERY slide a finished sample deck lacks (output step only).

    python restore_missing_slides.py "deck.pptx" "source.pptx" ["out.pptx"]

``source`` is the same deck built with FMI_DECK_STYLE=off, so no page was
dropped by the house style (volume pages of a country, YoY pages, extra
company profiles).  Slides are matched by title (first, second ... copy of a
repeated title).  A source slide the deck lacks is copied in - shapes, tables,
charts, padlocks - straight after the slide that precedes it in the source.
Slides the deck already has are not touched.
"""
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx import Presentation                          # noqa: E402
from fmi_report import deck_style as D                 # noqa: E402
from restore_country_volume_pages import _copy_shape   # noqa: E402


def main(deck, source, out):
    prs, src = Presentation(str(deck)), Presentation(str(source))
    lst = prs.slides._sldIdLst
    slides = list(prs.slides)
    have = Counter(D.title_of(s) for s in slides)
    seen = Counter()
    cursor = None                    # the deck slide the next insert follows
    where = {}                       # title -> list of deck slides in order
    for s in slides:
        where.setdefault(D.title_of(s), []).append(s)
    added = []
    for s in src.slides:
        t = D.title_of(s)
        if not t:                    # untitled (divider / breaker): leave
            continue
        k = seen[t]
        seen[t] += 1
        if k < have[t]:
            cursor = where[t][k]
            continue
        new = prs.slides.add_slide((cursor or slides[0]).slide_layout)
        for sh in list(new.shapes):
            sh._element.getparent().remove(sh._element)
        for sh in s.shapes:
            _copy_shape(s, new, sh)
        el = lst[-1]
        lst.remove(el)
        ids = [e.rId for e in lst]
        pos = 0
        if cursor is not None:
            pos = [sl.slide_id for sl in prs.slides].index(cursor.slide_id) + 1
        lst.insert(pos, el)
        cursor = new
        added.append(t)
    prs.save(str(out))
    print(f"{len(added)} slides restored; deck now {len(prs.slides)} slides")
    for t in added:
        print("  +", t[:100].encode("ascii", "replace").decode())


if __name__ == "__main__":
    a = sys.argv
    main(a[1], a[2], a[3] if len(a) > 3 else a[1])
