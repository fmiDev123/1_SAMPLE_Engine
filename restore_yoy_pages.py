"""Put a country's YoY pages back into a finished sample deck.

    python restore_yoy_pages.py "deck.pptx" "source.pptx" "Germany" ["out.pptx"]

``source`` is the same deck built with FMI_DECK_STYLE=off, so every YoY page
exists.  Each YoY page of ``country`` that the deck lacks is added (shapes,
tables, charts, padlocks copied) straight after the page it follows in the
source, found by title.  Nothing else changes.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx import Presentation                          # noqa: E402
from fmi_report import deck_style as D                 # noqa: E402
from restore_country_volume_pages import _copy_shape   # noqa: E402


def main(deck, source, country, out):
    prs, src = Presentation(str(deck)), Presentation(str(source))
    have = {D.title_of(s) for s in prs.slides}
    todo, prev = [], None
    for s in src.slides:
        t = D.title_of(s)
        if D.is_yoy(t):
            if t.startswith(country) and t not in have:
                todo.append((prev, s, t))
        else:
            prev = t
    added, lost = 0, []
    lst = prs.slides._sldIdLst
    for anchor, s, t in todo:
        idx = next((i for i, x in enumerate(prs.slides)
                    if D.title_of(x) == anchor), None)
        if idx is None:
            lost.append(t)
            continue
        new = prs.slides.add_slide(prs.slides[idx].slide_layout)
        for sh in list(new.shapes):
            sh._element.getparent().remove(sh._element)
        for sh in s.shapes:
            _copy_shape(s, new, sh)
        el = lst[-1]
        lst.remove(el)
        lst.insert(idx + 1, el)
        added += 1
        print("  added after", idx + 1, "-", t[:70])
    prs.save(str(out))
    print(f"{added} YoY pages added, {len(lost)} without an anchor; "
          f"deck now {len(prs.slides)} slides")
    for t in lost:
        print("  no anchor:", t.encode("ascii", "replace").decode())


if __name__ == "__main__":
    a = sys.argv
    main(a[1], a[2], a[3], a[4] if len(a) > 4 else a[1])
