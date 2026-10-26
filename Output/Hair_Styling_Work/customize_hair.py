"""Customise a built Hair Styling report deck - same deck-level post-process as the AIT sample
(copied from Output/AIT_Sample_work/customize_ait.py; sample-only steps switched off).

No engine source is touched: this works on the finished .pptx only.

  1. Volume is dropped from every SEGMENT page (global, regional, country);
     it stays on the regional / country level pages (summary, by region /
     country, country overviews).
  2. A Value Analysis table split over (1/2) and (2/2) is merged onto one slide.
  3. YoY slides are removed when the deck is still over 200 slides.
  4. One company profile is kept.
  5. Company share, company profile and the company cross-sections are moved
     to the start of the Global / regional / country blocks.
  6. Section numbers (dividers + agenda) follow the new order.
  7. Column A of every table is widened to fit the longest segment name.

usage: python customize_ait.py IN.pptx OUT.pptx
"""
import copy
import re
import sys

from PIL import ImageFont
from pptx import Presentation
from pptx.util import Emu

sys.path.insert(0, r"D:\PY\FMI_PPT")
import pptlock as pl                      # noqa: E402  (read-only helper)

PX = 9525                                   # EMU per slide px
SLIDE_LIMIT = 200


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
        return pl._slide_heading(slide).strip()
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
    b_card = [s for s in b.shapes if s.shape_type == 1
              and s.name.startswith("Rounded Rectangle")
              and s.width / PX > 300
              and s.top / PX > tb_.top / PX + tb_.height / PX]
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
    card_a = [s for s in a.shapes if s.shape_type == 1
              and s.name.startswith("Rounded Rectangle")
              and s.width / PX > 300 and 100 < s.top / PX < t_top][0]
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
        cards = [s for s in a.shapes if s.shape_type == 1
                 and s.name.startswith("Rounded Rectangle")
                 and s.width / PX > 300 and s.top / PX > table_bottom - 2]
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

    # indices of whole blocks to move
    def collect_cross(start_k):
        """a 'Cross-sectional Assessment' divider and its 8.x children."""
        out = [start_k]
        bar = info(segs[start_k][0])[1]
        parent = bar.replace("Section ", "")
        k = start_k + 1
        while k < len(segs):
            b = info(segs[k][0])[1]
            if re.match(rf"^Section {re.escape(parent)}\.\d+$", b):
                out.append(k)
                k += 1
            else:
                break
        return out

    cross_blocks = []
    for k, (i, j) in enumerate(segs):
        br, bar, ttl = info(i)
        if not br and ttl == "Cross-sectional Assessment":
            cross_blocks.append(collect_cross(k))
    share_k = next((k for k, (i, j) in enumerate(segs)
                    if info(i)[2] == "Market Share by Leading Players"), None)
    prof_k = next((k for k, (i, j) in enumerate(segs)
                   if info(i)[2].startswith("Company Profiles")), None)
    breakers = [k for k, (i, j) in enumerate(segs) if info(i)[0]]
    # destination per cross block: the breaker of the geography it follows
    def breaker_before(k):
        c = [b for b in breakers if b < k]
        return c[-1] if c else None
    # a country-level deck has no breaker of its own before its cross block;
    # the last breaker before it is the right anchor too

    moved = set()
    inserts = {}                       # breaker segment k -> [segment k...]
    glob_breaker = None
    for blk in cross_blocks:
        br = breaker_before(blk[0])
        inserts.setdefault(br, []).append(blk)
        if glob_breaker is None:
            glob_breaker = br
    # global start gets share + profile ahead of its cross block
    front = []
    if share_k is not None:
        front.append([share_k])
    if prof_k is not None:
        front.append([prof_k])
    if glob_breaker is not None:
        inserts[glob_breaker] = front + inserts.get(glob_breaker, [])
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


# ───────────────────── 6. section numbers + agenda ─────────────────────
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


def fix_agenda(deck, mapping, new_order_names):
    """Agenda cards: numbers follow the mapping and the cards are re-laid in
    the new order, group by group."""
    ag = None
    for s in deck.slides():
        if title_of(s) == "Agenda":
            ag = s
            break
    if ag is None:
        return
    cards = [sh for sh in ag.shapes if sh.shape_type == 1
             and sh.name.startswith("Rounded Rectangle")]
    heads, items = [], {}
    for sh in cards:
        t = sh.text_frame.text.strip()
        m = re.match(r"^(\d+)\.\s+(.*)$", t)
        if m:
            items[int(m.group(1))] = (sh, m.group(2))
        else:
            heads.append(sh)
    # restate numbers
    for old, (sh, label) in items.items():
        new = mapping.get(old)
        if new is not None:
            set_runs_text(sh, lambda _t, n=new, l=label: f"{n}. {l}")
    # layout: left column, right column - (header, [old numbers])
    groups = []
    hmap = {h.text_frame.text.strip(): h for h in heads}
    for name, olds in new_order_names:
        groups.append((hmap.get(name), olds))
    left_x, right_x, step = 15, 653, 34
    cols = [[], []]
    left_names = ("Global Market Outlook",)
    col = 0
    # market-signals card sits first on the left column
    y = {0: 119, 1: 119}
    first = items.get(1)
    if first:
        sh = first[0]
        sh.left, sh.top = px(33), px(y[0])
        y[0] += step
    for h, olds in groups:
        if h is None:
            continue
        if h.text_frame.text.strip() in ("Market Overview",):
            c = 0
        elif h.text_frame.text.strip() in ("Global Market Outlook",):
            c = 0
        else:
            c = 1
        x0 = left_x if c == 0 else right_x
        h.left, h.top = px(x0), px(y[c])
        y[c] += step + 4
        for o in olds:
            if o in items:
                sh = items[o][0]
                sh.left, sh.top = px(x0 + 18), px(y[c])
                y[c] += step
    return


def fix_agenda_numbers(deck, mapping):
    for s in deck.slides():
        if title_of(s) != "Agenda":
            continue
        for sh in s.shapes:
            if sh.shape_type == 1 and sh.has_text_frame:
                m = re.match(r"^(\d+)\.\s+(.*)$", sh.text_frame.text.strip())
                if m and int(m.group(1)) in mapping:
                    set_runs_text(sh, lambda _t, n=mapping[int(m.group(1))],
                                  l=m.group(2): f"{n}. {l}")


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
            need = 0.0
            for r in tbl.rows:
                c = r.cells[0]
                txt = c.text_frame.text.strip()
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
            changed += 1
    return changed


# ───────────────────────── main ─────────────────────────
def main(src, dst):
    prs = Presentation(src)
    deck = Deck(prs)
    log = {"start": len(prs.slides)}
    log["volume_removed"] = volume_off(deck)
    log["merged_pairs"] = merge_value_pairs(deck)
    log["profiles_dropped"] = 0              # full report keeps every profile
    n = len(prs.slides)
    if False:                                 # slide-limit trim is sample-only
        y = [s for s in deck.slides() if is_yoy(title_of(s))]
        for s in y:
            deck.delete(s)
        log["yoy_removed"] = len(y)
    else:
        log["yoy_removed"] = 0
    log["blocks_moved"] = reorder_blocks(deck)
    mapping = renumber_sections(deck)
    log["section_map"] = mapping
    fix_agenda_numbers(deck, mapping)
    log["tables_widened"] = widen_first_column(prs)
    log["end"] = len(prs.slides)
    prs.save(dst)
    return log


if __name__ == "__main__":
    out = main(sys.argv[1], sys.argv[2])
    for k, v in out.items():
        print(f"{k}: {v}")
