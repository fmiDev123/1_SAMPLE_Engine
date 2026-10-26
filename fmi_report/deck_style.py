"""Deck-level house style - applied to every deck the export builds.

This is the structure the approved sample carries, now part of the engine so
a locked SAMPLE and an unlocked REPORT are written the same way:

  1. A Value Analysis table split over (1/2) and (2/2) is merged onto one slide.
  2. Company share, company profile and the company cross-sections are moved
     to the start of the Global / regional / country blocks.
  3. Section numbers (dividers + agenda) follow the new order, and the agenda
     cards are re-laid group by group.
  4. Column A of every table is widened to fit the longest segment name.
  5. SAMPLES only (a locked deck): YoY slides are removed when the deck is
     still over SLIDE_LIMIT slides, and one company profile is kept.

Volume is KEPT on every segment page (global, regional, country) alongside
Value - both are shown for every segment in every report (client, 2026-10-22;
reverses the previous "Volume off every segment page" step). ``volume_off``
and ``country_volume_off`` below are kept for reference / reuse but are no
longer called from ``apply``.

Switch off for a build with FMI_DECK_STYLE=off.  It works on the finished
.pptx, so it is shape-based and does not touch what the page builders draw.
"""
from __future__ import annotations

import copy
import os
import re
from pathlib import Path

from PIL import ImageFont
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

PX = 9525                                   # EMU per slide px
SLIDE_LIMIT = 200

_DEFER = 0                                  # >0: a caller applies the style


class defer:
    """`with defer():` - the export inside is NOT styled (a locked build
    styles the deck after the padlock, so the unlocked pass is skipped)."""
    def __enter__(self):
        global _DEFER
        _DEFER += 1

    def __exit__(self, *a):
        global _DEFER
        _DEFER -= 1


def enabled() -> bool:
    return os.environ.get("FMI_DECK_STYLE", "on").strip().lower() != "off"


def deferred() -> bool:
    return _DEFER > 0


def _pl():
    from .locking import _pptlock
    return _pptlock()


# ───────────────────────── helpers ─────────────────────────
def px(v):
    return int(round(v * PX))


def texts(shapes):
    for s in shapes:
        if s.shape_type == 6:
            yield from texts(s.shapes)
        if s.has_text_frame:
            yield s.text_frame.text


def title_of(slide):
    try:
        return _pl()._slide_heading(slide).strip()
    except Exception:
        return ""


def bar_text(slide):
    """'Section 7.1' of a section divider, else ''."""
    for sh in slide.shapes:
        if sh.name == "FMI Section Bar" and sh.has_text_frame:
            return sh.text_frame.text.strip()
    return ""


def divider_title(slide):
    for sh in slide.shapes:
        if sh.name in ("FMI Section Title", "FMI Breaker Title") \
                and sh.has_text_frame:
            return sh.text_frame.text.strip()
    return ""


def is_breaker(slide):
    return any(sh.name == "FMI Section Breaker" for sh in slide.shapes)


def box(sh):
    return (sh.left / PX, sh.top / PX, (sh.left + sh.width) / PX,
            (sh.top + sh.height) / PX)


def inside(inner, outer, tol=3):
    return (inner[0] >= outer[0] - tol and inner[1] >= outer[1] - tol
            and inner[2] <= outer[2] + tol and inner[3] <= outer[3] + tol)


def remove_shape(sh):
    sh._element.getparent().remove(sh._element)


def set_runs_text(shape, fn):
    """Apply ``fn`` to the shape's text, keeping the first run's formatting:
    the whole text of each paragraph lands in its first run."""
    for p in shape.text_frame.paragraphs:
        runs = list(p.runs)
        if not runs:
            continue
        full = "".join(r.text for r in runs)
        runs[0].text = fn(full)
        for r in runs[1:]:
            r.text = ""


# ───────────────────── slide list management ───────────────────
class Deck:
    def __init__(self, prs):
        self.prs = prs
        self.lst = prs.slides._sldIdLst

    def slides(self):
        return list(self.prs.slides)

    def ids(self):
        return list(self.lst)

    def delete(self, slide):
        for sid in list(self.lst):
            if self.prs.part.related_part(sid.rId) is slide.part:
                self.prs.part.drop_rel(sid.rId)
                self.lst.remove(sid)
                return

    def reorder(self, id_order):
        for sid in list(self.lst):
            self.lst.remove(sid)
        for sid in id_order:
            self.lst.append(sid)


# ─────────────── 1. volume off the segment pages ───────────────
VOL_SEG_TITLE = re.compile(
    r"(Market Volume Analysis, By (?!Region|Country)|"
    r"YoY Analysis of .* Market Volume, By )")


def drop_panel(slide, title_shape):
    """Remove the card a panel-title text box sits in, with everything in it."""
    tb = box(title_shape)
    cards = [s for s in slide.shapes
             if s.shape_type == 1 and s.name.startswith("Rounded Rectangle")
             and s.width / PX > 300
             and box(s)[0] <= tb[0] + 2 and box(s)[1] <= tb[1] + 2
             and box(s)[2] >= tb[2] - 2 and box(s)[3] >= tb[3] - 2]
    if not cards:
        return None
    card = min(cards, key=lambda s: s.width * s.height)
    cb = box(card)
    for s in list(slide.shapes):
        if s is card or inside(box(s), cb):
            remove_shape(s)
    return cb


def volume_panel_title(slide):
    for s in slide.shapes:
        if s.has_text_frame and s.shape_type == 17:
            t = s.text_frame.text
            if re.search(r"\bVolume\b", t) and (
                    "Shares & Variance" in t or "Volume YoY" in t
                    or re.match(r"^.* Volume, By ", t)):
                return s
    return None


def tidy_volume_words(slide):
    for s in slide.shapes:
        if not s.has_text_frame or s.shape_type != 17:
            continue
        if re.search(r"volume and value|volume & value", s.text_frame.text,
                     re.I):
            set_runs_text(s, lambda t: re.sub(
                r"(?i)volume (and|&) value", "value", t))


def strip_part_suffix(slide):
    for s in slide.shapes:
        if s.has_text_frame and s.shape_type == 17:
            if re.search(r"\(\d/\d\)\s*$", s.text_frame.text) and \
                    s.top / PX < 20:
                set_runs_text(s, lambda t: re.sub(r"\s*\(\d/\d\)\s*$", "", t))
                return


def shift_up(slide, below_top, delta):
    for s in slide.shapes:
        if s.top is not None and s.top / PX >= below_top - 1 \
                and 100 < s.top / PX < 690:
            s.top = s.top - px(delta)


def has_volume_panel(slide):
    return volume_panel_title(slide) is not None


def country_volume_off(deck):
    gone, dropped_base = 0, set()
    for sl in deck.slides():
        t = title_of(sl)
        if not re.match(r"^.*(Market Tables|YoY Analysis|Segmental "
                        r"Analysis)", t):
            continue
        if re.match(r"^YoY Analysis of ", t):           # global / regional
            continue
        base = re.sub(r"\s*\(\d/2\)\s*$", "", t)
        if re.search(r"\(1/2\)\s*$", t):
            if has_volume_panel(sl):
                deck.delete(sl)
                dropped_base.add(base)
                gone += 1
            continue
        if re.search(r"\(2/2\)\s*$", t):
            if base in dropped_base:
                strip_part_suffix(sl)
                tidy_volume_words(sl)
            continue
        vt = volume_panel_title(sl)
        if vt is None:
            continue
        cb = drop_panel(sl, vt)
        if cb is None:
            continue
        others = [s for s in sl.shapes if s.shape_type == 1
                  and s.name.startswith("Rounded Rectangle")
                  and s.width / PX > 300 and s.top / PX > cb[3] - 5]
        if others:
            top2 = min(s.top / PX for s in others)
            shift_up(sl, top2, top2 - cb[1])
        else:
            widen_row(sl, cb[1], 1250)
        tidy_volume_words(sl)
    return gone


def widen_row(slide, row_top, full_w):
    """The remaining card on a row takes the whole width."""
    scale = None
    for s in slide.shapes:
        if s.top is None:
            continue
        if abs(s.top / PX - row_top) < 4 and s.shape_type == 1 \
                and s.left / PX < 40:
            scale = full_w / (s.width / PX)
            break
    if not scale:
        return
    for s in slide.shapes:
        if s.top is None or s.name.startswith("Rectangle"):
            continue
        if s.top / PX >= row_top - 2 and s.left / PX < 640 \
                and s.height / PX < 330 and s.top / PX < 700:
            if s.shape_type == 13:                       # padlock picture
                s.left = px(25 + (1230 - s.width / PX) / 2) if \
                    s.left / PX > 200 else s.left
                continue
            s.width = int(s.width * scale)


def volume_off(deck):
    removed = 0
    for sl in deck.slides():
        if VOL_SEG_TITLE.search(title_of(sl)):
            deck.delete(sl)
            removed += 1
    removed += country_volume_off(deck)
    # divider bullet lists: 'Market Volume Analysis' goes
    for sl in deck.slides():
        if bar_text(sl).startswith("Section 7.") or \
                re.match(r"^Section \d+\.\d+$", bar_text(sl)):
            for sh in sl.shapes:
                if sh.has_text_frame and "Market Volume Analysis" in \
                        sh.text_frame.text:
                    for p in list(sh.text_frame.paragraphs):
                        if "Market Volume Analysis" in "".join(
                                r.text for r in p.runs):
                            p._p.getparent().remove(p._p)
    return removed


# ───────────────────── 2. merge (1/2)+(2/2) value tables ─────────────────
def tables_of(slide):
    return [s for s in slide.shapes if s.has_table]


def _cards(shapes, width_min=300):
    """Auto-shape 'cards' at least width_min px wide that frame a table or
    commentary block. Most templates name these 'Rounded Rectangle'; the
    Fact.MR template draws the same framing boxes as plain 'Rectangle'
    (2026-10-22, Water & Wastewater Infrastructure Market) - try the usual
    name first so FMI/TMR behaviour is unchanged, then fall back.
    """
    found = [s for s in shapes if s.shape_type == 1
             and s.name.startswith("Rounded Rectangle")
             and s.width / PX > width_min]
    return found or [s for s in shapes if s.shape_type == 1
                     and s.name.startswith("Rectangle")
                     and s.width / PX > width_min]


def merge_value_pairs(deck):
    merged = 0
    sl = deck.slides()
    i = 0
    while i < len(sl) - 1:
        a, b = sl[i], sl[i + 1]
        ta, tb_ = title_of(a), title_of(b)
        if re.search(r"\(1/2\)\s*$", ta) and re.search(r"\(2/2\)\s*$", tb_) \
                and re.sub(r"\(\d/2\)", "", ta) == re.sub(r"\(\d/2\)", "", tb_) \
                and "Value Analysis" in ta \
                and len(tables_of(a)) == 1 and len(tables_of(b)) == 1:
            if merge_pair(deck, a, b):
                merged += 1
                sl = deck.slides()
                continue
        i += 1
    return merged


def merge_pair(deck, a, b):
    ta, tb_ = tables_of(a)[0], tables_of(b)[0]
    rows_a = list(ta.table._tbl.tr_lst)
    rows_b = list(tb_.table._tbl.tr_lst)[1:]            # drop B's header
    n = len(rows_a) + len(rows_b)
    row_h = 23.0
    t_top = ta.top / PX
    new_h = n * row_h
    # extra shapes that sit below B's table (the commentary card)
    b_card = [s for s in _cards(b.shapes)
              if s.top / PX > tb_.top / PX + tb_.height / PX]
    card_h = (b_card[0].height / PX) if b_card else 0
    table_bottom = t_top + new_h + 24                    # + source line
    avail = 696 - table_bottom - 10
    if b_card and avail < min(card_h, 150):
        return False
    if table_bottom > 690:
        return False
    for tr in rows_b:
        ta.table._tbl.append(copy.deepcopy(tr))
    for tr in ta.table._tbl.tr_lst:
        tr.h = px(row_h)
    ta.height = px(new_h)
    # A's card, its source line and padlock
    card_a = [s for s in _cards(a.shapes) if 100 < s.top / PX < t_top][0]
    card_a.height = px(table_bottom - card_a.top / PX - 6)
    for s in a.shapes:
        if s.has_text_frame and s.text_frame.text.startswith("Source:") \
                and s.top / PX > t_top:
            s.top = px(t_top + new_h + 2)
    for s in a.shapes:
        if s.shape_type == 13 and s.top / PX > t_top + 20 \
                and s.top / PX < 690 and s.width / PX < 60:
            s.top = px(t_top + new_h / 2 - s.height / PX / 2)
    # B's commentary card, moved onto A
    if b_card:
        dy = (table_bottom + 4) - b_card[0].top / PX
        cb = box(b_card[0])
        for s in list(b.shapes):
            if s.top is None:
                continue
            if inside(box(s), cb) or s is b_card[0]:
                if s.shape_type == 13:                  # a picture: own rel
                    import io
                    a.shapes.add_picture(io.BytesIO(s.image.blob), s.left,
                                         s.top + px(dy), s.width, s.height)
                    continue
                el = copy.deepcopy(s._element)
                a.shapes._spTree.append(el)
                new = a.shapes[-1]
                new.top = new.top + px(dy)
        # fit the card to the page bottom
        cards = [s for s in _cards(a.shapes)
                 if s.top / PX > table_bottom - 2]
        if cards:
            c = cards[0]
            c.height = px(min(c.height / PX, 696 - c.top / PX))
    strip_part_suffix(a)
    deck.delete(b)
    return True


# ───────────────────── 3/4. YoY and company profiles ─────────────────────
def is_yoy(title):
    return bool(re.match(r"^YoY Analysis of ", title)
                or re.search(r"\bYoY Analysis\b", title))


PROFILE_TAIL = re.compile(
    r" (Executive Market Snapshot|Regional Penetration & Win Position|"
    r"SWOT Analysis|Key Financials|Strategy Analysis) \(\d/5\)\s*$")


def keep_one_profile(deck):
    first, gone = None, 0
    for sl in deck.slides():
        m = PROFILE_TAIL.search(title_of(sl))
        if not m:
            continue
        name = title_of(sl)[:m.start()]
        first = first or name
        if name != first:
            deck.delete(sl)
            gone += 1
    # the divider that lists the companies keeps only the first
    if first:
        for sl in deck.slides():
            if divider_title(sl).startswith("Company Profiles"):
                for sh in sl.shapes:
                    if sh.has_text_frame and "Snapshot" in sh.text_frame.text:
                        for p in list(sh.text_frame.paragraphs):
                            t = "".join(r.text for r in p.runs)
                            if t and first not in t:
                                p._p.getparent().remove(p._p)
    return gone, first


# ───────────────────── 5. reorder blocks ─────────────────────
def segments_of(deck):
    """[(start_index, end_index, key)] over the deck: every divider/breaker
    opens a segment that runs to the next divider."""
    sl = deck.slides()
    marks = [i for i, s in enumerate(sl)
             if is_breaker(s) or bar_text(s) or
             any(x.name == "FMI Section Title" for x in s.shapes)]
    out = []
    for k, i in enumerate(marks):
        j = marks[k + 1] if k + 1 < len(marks) else len(sl)
        out.append((i, j))
    return out


def reorder_blocks(deck):
    sl = deck.slides()
    ids = deck.ids()
    segs = segments_of(deck)

    def info(i):
        s = sl[i]
        return (is_breaker(s), bar_text(s), divider_title(s))

    # The deck follows the research team's table of contents (2026-10-14):
    # cross-sectional blocks stay inside their own geography's outlook where
    # the builder put them, and the company share + profile blocks belong to
    # the Competition Assessment, the last section - they are moved there
    # (right behind its breaker), never to the front of the Global outlook.
    share_k = next((k for k, (i, j) in enumerate(segs)
                    if info(i)[2] == "Market Share by Leading Players"), None)
    prof_k = next((k for k, (i, j) in enumerate(segs)
                   if info(i)[2].startswith("Company Profiles")), None)
    comp_k = next((k for k, (i, j) in enumerate(segs)
                   if info(i)[0] and info(i)[2] == "Competition Assessment"),
                  None)
    moved = set()
    inserts = {}                       # breaker segment k -> [segment k...]
    comp_blocks = [[k] for k in (share_k, prof_k) if k is not None]
    if comp_blocks and comp_k is not None:
        inserts[comp_k] = comp_blocks
    for blks in inserts.values():
        for blk in blks:
            moved.update(blk)

    new = []
    for k, (i, j) in enumerate(segs):
        pre = [x for x in range(0, segs[0][0])] if k == 0 else []
        new.extend(pre)
        if k in moved:
            continue
        new.extend(range(i, j))
        for blk in inserts.get(k, []):
            for kk in blk:
                new.extend(range(segs[kk][0], segs[kk][1]))
    tail_start = segs[-1][1]
    new.extend(range(tail_start, len(sl)))
    assert sorted(new) == list(range(len(sl))), "reorder lost slides"
    deck.reorder([ids[i] for i in new])
    return len(moved)


# ───────────────────── 6. section numbers + agenda ─────────────────
def renumber_sections(deck):
    sl = deck.slides()
    order, mapping = [], {}
    for s in sl:
        m = re.match(r"^Section (\d+)$", bar_text(s))
        if m:
            old = int(m.group(1))
            if old not in mapping:
                order.append(old)
                mapping[old] = len(order)
    for s in sl:
        for sh in s.shapes:
            if sh.name == "FMI Section Bar" and sh.has_text_frame:
                t = sh.text_frame.text.strip()
                m = re.match(r"^Section (\d+)(\.\d+)?$", t)
                if m and int(m.group(1)) in mapping:
                    new = f"Section {mapping[int(m.group(1))]}{m.group(2) or ''}"
                    set_runs_text(sh, lambda _t, n=new: n)
    return mapping


def section_groups(deck):
    """({section number: the geography breaker it sits under, or None},
    [breaker titles in deck order]) - read off the deck as it now stands."""
    cur, out, order = None, {}, []
    for s in deck.slides():
        if is_breaker(s):
            cur = divider_title(s) or cur
            if cur and cur not in order:
                order.append(cur)
            continue
        m = re.match(r"^Section (\d+)$", bar_text(s))
        if m:
            out.setdefault(int(m.group(1)), cur)
    return out, order


def fix_agenda(deck, mapping):
    """Agenda cards: numbers follow the new section numbers, and every card
    sits under the geography it now belongs to, in the columns and at the
    pitch the page already has.  A deck that did not change is left as built."""
    ag = next((s for s in deck.slides() if title_of(s) == "Agenda"), None)
    if ag is None:
        return
    cards = [sh for sh in ag.shapes if sh.shape_type == 1
             and sh.name.startswith("Rounded Rectangle")
             and sh.has_text_frame and sh.text_frame.text.strip()]
    if not cards:
        return
    num_re = re.compile(r"^(\d+)\.\s+(.*)$")
    col_of = lambda c: 0 if c.left / PX < 640 else 1
    ordered = sorted(cards, key=lambda c: (col_of(c), c.top))
    # original grouping: a header and the numbered cards that follow it
    heads, item_group, cur = {}, {}, None
    for c in ordered:
        txt = c.text_frame.text.strip()
        m = num_re.match(txt)
        if m:
            item_group[id(c)] = cur
        else:
            cur = txt
            heads[txt] = c
    changed = any(k != v for k, v in mapping.items())
    for c in cards:
        m = num_re.match(c.text_frame.text.strip())
        if m and int(m.group(1)) in mapping:
            set_runs_text(c, lambda _t, n=mapping[int(m.group(1))],
                          l=m.group(2): f"{n}. {l}")
    gmap, border = section_groups(deck)
    new_group, moved = {}, False
    for c in cards:
        m = num_re.match(c.text_frame.text.strip())
        if not m:
            continue
        no = int(m.group(1))                 # already the NEW number
        g = gmap.get(no)
        if g in heads or (g is None and not heads):
            tgt = g
        elif g is None:
            tgt = item_group[id(c)] if no not in gmap else None
        else:
            tgt = item_group[id(c)]
        if tgt != item_group[id(c)]:
            moved = True
        new_group[id(c)] = tgt
    overflow = max((c.top + c.height) / PX for c in cards) > 700
    if not changed and not moved and not overflow:
        return
    items = [c for c in cards if num_re.match(c.text_frame.text.strip())]
    num = lambda c: int(num_re.match(c.text_frame.text.strip()).group(1))
    order = [None] + [h for h in border if h in heads]
    order += [h for h in heads if h not in order]
    groups = []
    for h in order:
        its = sorted([c for c in items if new_group[id(c)] == h], key=num)
        if h is None and not its:
            continue
        groups.append((heads.get(h) if h is not None else None, its))
    is_item = lambda c: bool(num_re.match(c.text_frame.text.strip()))
    it = sorted([c for c in cards if is_item(c) and col_of(c) == 0],
                key=lambda c: c.top)
    gaps = [b.top - a.top for a, b in zip(it, it[1:]) if b.top - a.top > 0]
    pitch = min(gaps) if gaps else px(34)
    hd_gap = max([c2.top - h.top for h in cards if not is_item(h)
                  for c2 in it if c2.top > h.top][:1] or [pitch])
    hd_gap = max(hd_gap, pitch)
    y0 = min(c.top for c in cards)
    limit = px(700)
    base_x = {k: min(c.left for c in cards if col_of(c) == k)
              for k in (0, 1) if any(col_of(c) == k for c in cards)}
    dx0 = {id(c): c.left - base_x[col_of(c)] for c in cards}
    hdr_w = max((c.width for c in cards if not is_item(c)), default=px(612))
    itm_w = max((c.width for c in cards if is_item(c)), default=px(594))

    def place(col_x, wh, wi):
        """Fill the columns in order; None when the last one overruns."""
        pos, col, y = {}, 0, y0
        for hd, its in groups:
            seq = ([hd] if hd is not None else []) + its
            first_two = hd_gap + (pitch if its else 0)
            if y + first_two > limit and col < len(col_x) - 1:
                col, y = col + 1, y0
            for c in seq:
                if y + c.height > limit and col < len(col_x) - 1:
                    col, y = col + 1, y0
                pos[id(c)] = (col, y)
                y += hd_gap if c is hd else pitch
        end = max(yy + c.height for c in cards for (cc, yy) in [pos[id(c)]])
        return pos if end <= limit else None

    two_x = [base_x[0], base_x.get(1, base_x[0] + px(638))]
    plan = place(two_x, hdr_w, itm_w)
    xs, wh, wi = two_x, hdr_w, itm_w
    if plan is None:                              # 3 columns, 414 px each
        xs, wh, wi = [px(15), px(435), px(855)], px(414), px(396)
        plan = place(xs, wh, wi)
    comp = 1.0
    if plan is None:                              # still too tall: tighten
        pitch0, hd0 = pitch, hd_gap
        need = max(yy + c.height - y0 for c in cards
                   for (cc, yy) in [(0, y0)]) or 1
        for f in (0.9, 0.8, 0.7, 0.6):
            pitch, hd_gap = int(pitch0 * f), int(hd0 * f)
            plan = place(xs, wh, wi)
            if plan is not None:
                comp = f
                break
        if plan is None:
            return
    for c in cards:
        col, y = plan[id(c)]
        head_card = not is_item(c)
        c.left = xs[col] + (0 if head_card else px(18))
        c.top = y
        if xs is not two_x:
            c.width = wh if head_card else wi
        if comp < 1.0:
            c.height = int(c.height * comp)


# ───────────────────── 7. column A ─────────────────────
_FONT_CACHE = {}


def _font(bold):
    key = bool(bold)
    if key not in _FONT_CACHE:
        for f in (("calibrib.ttf" if bold else "calibri.ttf"),
                  "segoeuib.ttf", "arialbd.ttf"):
            try:
                _FONT_CACHE[key] = ImageFont.truetype(f, 14)   # 14px = 10.5pt
                break
            except Exception:
                continue
    return _FONT_CACHE[key]


def text_px(text, size_pt, bold):
    f = _font(bold)
    return f.getlength(text) * (size_pt * 4 / 3) / 14.0


_DASH_CHARS = "-–—"
_DASH_SPLIT = re.compile(r"(?<=[%s])" % _DASH_CHARS)
# export_native._break_lines's own mark for a line it cut mid-piece (a
# single word, or a dash-piece of one, too wide for one line on its own) -
# invisible in PowerPoint, so rejoining strips it rather than treating the
# join as a word boundary and adding a space ("stainless-" cut again,
# mid-piece, must not come back as "stainles s-steel").
_CONT_MARK = "​"


def _pieces(word):
    """A word's pieces at the points a greedy wrap may break it without a
    space - mirrors export_native._pieces, so a label widened here rewraps
    the way it was first wrapped."""
    return [p for p in _DASH_SPLIT.split(str(word)) if p] or [word]


def _rejoin_lines(lines):
    """The original one-line text behind lines a greedy wrap cut: a line
    marked with ``_CONT_MARK`` continues the next with no space - set by
    ``export_native._break_lines`` at every join that takes none, a dash
    break and a mid-piece cut alike.  A line ending in a dash but NOT
    marked is a dash standing on its own (a leading bullet, "– Item") and
    does take a space before what follows - the mark, not the trailing
    character, says whether a join is a word boundary."""
    if not lines:
        return ""
    prev_cont = lines[0].endswith(_CONT_MARK)
    out = lines[0][:-1] if prev_cont else lines[0]
    for ln in lines[1:]:
        cont = ln.endswith(_CONT_MARK)
        body = ln[:-1] if cont else ln
        out = out + body if prev_cont else out + " " + body
        prev_cont = cont
    return out


def _rewrap(text, avail, size, bold):
    """Greedy word-wrap of ``text`` into lines that each fit ``avail`` px -
    mirrors export_native._break_lines, so a label cell widened here
    rewraps at its NEW width instead of keeping line breaks cut for its
    old, narrower one (2026-10-22: a label pre-wrapped by the page builder
    for a cramped column read as broken mid-word once this pass widened
    the column, because the column grew but the already-written line
    breaks never got recomputed)."""
    def fits(s):
        return text_px(s, size, bold) <= avail
    lines, cur = [], ""
    for word in str(text).split():
        for j, piece in enumerate(_pieces(word)):
            trial = (cur + (" " if j == 0 else "") + piece) if cur else piece
            if fits(trial):
                cur = trial
                continue
            if cur:
                lines.append(cur)
            cur = piece
            while len(cur) > 1 and not fits(cur):
                k = len(cur) - 1
                while k > 1 and not fits(cur[:k]):
                    k -= 1
                lines.append(cur[:k])
                cur = cur[k:]
    lines.append(cur)
    return [ln for ln in lines if ln] or [str(text)]


def _reflow_cell(cell, avail):
    """Re-cut a label cell's lines to fit ``avail`` px, now that its column
    has been widened - a no-op for a cell that was never cut into more
    than one line (its own wrap, if any, is PowerPoint's at paint time, not
    a baked-in line break)."""
    paras = list(cell.text_frame.paragraphs)
    lines = [p.text for p in paras if p.text]
    if len(paras) < 2 or not lines:
        return False
    first_run = next((r for p in paras for r in p.runs), None)
    if first_run is None:
        return False
    size = first_run.font.size.pt if first_run.font.size else 11.0
    bold = bool(first_run.font.bold)
    color = (first_run.font.color.rgb
             if first_run.font.color and first_run.font.color.type is not None
             else None)
    name = first_run.font.name
    new_lines = _rewrap(_rejoin_lines(lines), avail, size, bold)
    if new_lines == lines:
        return False
    tf = cell.text_frame
    for p in paras[1:]:
        p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    for i, ln in enumerate(new_lines):
        p = p0 if i == 0 else tf.add_paragraph()
        if i:
            p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        if color is not None:
            r.font.color.rgb = color
        if name:
            r.font.name = name
    return True


def widen_first_column(prs):
    changed = 0
    for sl in prs.slides:
        for sh in sl.shapes:
            if not sh.has_table or sh.name == "FMI Contents Table":
                continue
            tbl = sh.table
            ncols = len(tbl.columns)
            if ncols < 3:
                continue
            try:                  # the Market Signals table keeps its widths:
                head = tbl.cell(0, 0).text_frame.text.strip().lower()
            except Exception:     # the lock lays its text over fixed columns
                head = ""
            if head.startswith("questions we"):
                continue
            need = 0.0
            for r in tbl.rows:
                c = r.cells[0]
                lines = [p.text for p in c.text_frame.paragraphs if p.text]
                txt = _rejoin_lines(lines) if len(lines) > 1 \
                    else c.text_frame.text.strip()
                if not txt:
                    continue
                size = 11.0
                bold = True
                for p in c.text_frame.paragraphs:
                    for run in p.runs:
                        if run.font.size:
                            size = run.font.size.pt
                        bold = bool(run.font.bold) or bold
                need = max(need, text_px(txt, size, bold))
            need += 18                                     # cell margins
            cur = tbl.columns[0].width / PX
            total = sum(c.width for c in tbl.columns) / PX
            cap = total * 0.42
            want = min(max(cur, need), cap)
            if want <= cur + 1:
                continue
            others = total - cur
            scale = (total - want) / others
            floor = 38.0
            if min(c.width / PX for c in list(tbl.columns)[1:]) * scale < floor:
                want = total - floor * (ncols - 1)
                if want <= cur + 1:
                    continue
                scale = (total - want) / others
            widths = [want] + [c.width / PX * scale
                               for c in list(tbl.columns)[1:]]
            acc = 0
            for k, w in enumerate(widths):
                if k < len(widths) - 1:
                    ew = px(w)
                    acc += ew
                else:
                    ew = sum(c.width for c in tbl.columns) - acc
                tbl.columns[k].width = Emu(ew)
            for r in tbl.rows:
                c = r.cells[0]
                n_before = len(c.text_frame.paragraphs)
                if _reflow_cell(c, want - 18):
                    n_after = len(c.text_frame.paragraphs)
                    if n_after < n_before:
                        r.height = Emu(int(r.height * n_after / n_before))
            changed += 1
    return changed


# ───────────────────────── main ─────────────────────────
def apply(src, dst=None, sample=False) -> dict:
    """Style the deck at ``src`` and save it to ``dst`` (default: in place).
    ``sample`` is True for a locked deck (also trims it - see the docstring).
    Returns a log of what changed."""
    prs = Presentation(str(src))
    deck = Deck(prs)
    log = {"start": len(prs.slides)}
    log["volume_removed"] = 0       # Volume kept on every segment page
    log["merged_pairs"] = merge_value_pairs(deck)
    log["profiles_dropped"], log["yoy_removed"] = 0, 0
    if sample:
        gone, _first = keep_one_profile(deck)
        log["profiles_dropped"] = gone
        if len(prs.slides) > SLIDE_LIMIT:
            y = [s for s in deck.slides() if is_yoy(title_of(s))]
            for s in y:
                deck.delete(s)
            log["yoy_removed"] = len(y)
    log["blocks_moved"] = reorder_blocks(deck)
    mapping = renumber_sections(deck)
    fix_agenda(deck, mapping)
    log["tables_widened"] = widen_first_column(prs)
    log["end"] = len(prs.slides)
    prs.save(str(dst or src))
    return log


def style_file(path, sample=False):
    """In-place `apply` that never loses the deck: any failure leaves the
    unstyled file as it was and says so."""
    path = Path(path)
    if not enabled():
        return None
    tmp = path.with_suffix(".styled.pptx")
    try:
        log = apply(path, tmp, sample=sample)
        os.replace(tmp, path)
        return log
    except Exception as exc:                       # pragma: no cover
        try:
            tmp.unlink()
        except OSError:
            pass
        print(f"WARNING: deck style skipped for {path.name}: {exc!r}")
        return None
