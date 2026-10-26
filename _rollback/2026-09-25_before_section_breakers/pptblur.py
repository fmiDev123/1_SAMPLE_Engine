"""pptblur - the padlock, over a BLUR of the figure it withholds.

A companion to `pptlock.py`, and like it a SINGLE self-contained file that
imports nothing from any report engine.  `pptlock` decides WHAT to withhold;
this decides WHAT THE READER SEES IN ITS PLACE:

    from pptblur import blur_pptx
    blur_pptx("report.pptx", "report_locked.pptx",
              policy=LockPolicy(open_years={2025, 2026, 2036},
                                totals_only=True))

or from the command line:

    python pptblur.py report.pptx out.pptx --open-years 2025,2026,2036 --totals-only


THE RULE: BLUR THE VALUE, LOCK THE PANEL
----------------------------------------
A padlock alone says "something was here".  A blur says "a figure of THIS
SHAPE was here" - the reader sees the number's width and weight, enough to
believe the deck is real and complete, and not enough to read it.  So:

    every withheld figure    is cleared and replaced by a BLURRED RASTER
                             of that very text, in place, at its own size
    every CONTAINER that     takes ONE padlock, centred - a table, a chart,
    lost a figure            or a panel of loose figures clustered together

The padlock is deliberately NOT on the value.  One per value buries the blur
it is meant to sit on and turns a size table into a wall of icons; one per
panel says the same thing once and leaves the blurs legible AS blurs.  The
rasters are kept faint (peak alpha ~0.27, near-black ink) because that is
what the reference deck measures at - darker reads as a redaction bar,
lighter as a smudge on the slide.

WHAT IS TREATED HOW
-------------------
`pptlock.scan` finds four kinds of value; each is blurred in the way its
container allows:

    table cell    blurred in place, the cell's text cleared; the TABLE takes
                  the padlock
    shape value   blurred in place - a KPI card, or a bare figure on a
                  bespoke page; the panel it clusters into takes the padlock
    chart label   a native chart draws its labels from a number format, so
                  anything else can only arrive as the label's PICTURE FILL:
                  the blur is fed through `pptlock`'s own label plumbing, and
                  the CHART takes the padlock
    prose figure  "xx" - DrawingML has no inline picture, so a figure inside
                  a sentence can only become a character (`pptlock`'s own
                  conclusion, and its behaviour)

WHAT IS NEVER WITHHELD: STRUCTURE
---------------------------------
A chart stripped of its axis is not locked, it is broken - and a reader who
cannot see 2021 / 2026 / 2036 under the bars cannot tell what the blurs are
even about.  Two things therefore survive whatever the policy says:

    a lone YEAR or span of years    a category label, not a figure - the same
                                    rule `pptlock.prose_figures` already
                                    applies inside a sentence
    an AXIS TICK                    three or more short lone figures sharing
                                    a left edge and a width, EVENLY spaced
                                    down the slide (or sharing a top edge
                                    across it).  Data labels fail the test:
                                    they share a column but sit where their
                                    series lands, so their gaps are ragged.

Measured against the reference deck, this keeps 143 labels readable across 24
slides - its 0% / 50% / 100% scales, its -451 / +0 / +451 BPS axis and its
year categories - while still withholding the BPS markers beside them.

Nothing about a market, a template or a page is known here: every rect comes
from the deck, every string from the deck, and the policy from `pptlock`.

WHAT IT CANNOT DO
-----------------
* A chart or table pasted as an IMAGE is opaque, and is left alone.
* Blurring needs Pillow.  Without it the module still runs and still locks -
  it simply places the padlock with no blur under it, and says so in the
  report.
* A font is looked up by name in the OS font directories; when the deck's own
  font is missing the nearest available one is rasterised instead.  A blur is
  a shape, not a reading, so a substituted font costs nothing.
"""
from __future__ import annotations

import argparse
import io
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu, Pt

import pptlock
from pptlock import LockPolicy, LockArt, EMU_PER_PX

__all__ = ["BlurStyle", "BlurReport", "blur_pptx", "render_blur",
           "LockPolicy"]


# ═══════════════════════════ how it looks ════════════════════════════
@dataclass
class BlurStyle:
    """The blur's measurements, and where the padlock goes.

    **The value is blurred; the padlock is not on the value.** One padlock is
    placed per CONTAINER - a table, a chart, or a cluster of loose figures -
    centred on it, which is how the reference deck marks a panel as withheld.
    A padlock per value buries the blur it is supposed to sit on and turns a
    size table into a wall of icons; one padlock says the same thing once.

    The blur's numbers are the reference's own rasters, measured: a 3x
    supersampled render, a Gaussian at ~0.19 of the raster height, an alpha
    ceiling of 0.27 over near-black ink.
    """

    blur_values: bool = True           # a withheld figure becomes a blur of itself
    lock_values: bool = False          # ... and NOT a padlock on top of it
    lock_containers: bool = True       # one padlock per table / chart / cluster
    # what a figure inside a sentence becomes: "inline-lock" redraws the
    # paragraph with a padlock where the figure stood (the reference's own
    # answer); "xx" keeps `pptlock`'s placeholder, which never mis-wraps
    prose: str = "inline-lock"
    # a chart without its axis is broken, not locked: leave scale ticks and
    # the years that name a column readable (see `axis_paths` / `is_structure`)
    keep_structure: bool = True
    # a table cell that is a SENTENCE carrying figures ("Scale base (58.3%
    # share)") cannot take an inline padlock - a cell has no room to redraw -
    # so it is withheld whole, blur and all, which is what the reference does.
    # "prose" instead keeps the words and swaps the figures for "xx".
    narrative_cells: str = "blank"
    # `totals_only` narrows the open columns to the total row - but only where
    # there IS one, or a table without a total loses even its open columns
    totals_only_needs_total: bool = True

    supersample: int = 3               # render at Nx, then blur, then place
    blur_frac: float = 0.19            # Gaussian radius, as a fraction of height
    alpha: float = 0.27                # peak opacity of the blurred ink
    ink: str = "#1F1E1E"               # what the reference's rasters measure
    pad_frac: float = 0.55             # slack around the text, x its height
    # a blur is only worth drawing when the box is big enough to show one
    min_px: float = 7.0

    # the container padlock: a fraction of the container's height, held
    # between two sizes so it reads on a KPI strip and does not swamp a chart
    lock_frac: float = 0.12
    lock_min: float = 22.0
    lock_max: float = 64.0
    # loose figures within this many px of one another are ONE panel, and so
    # take one padlock between them
    cluster_gap: float = 90.0


@dataclass
class BlurReport:
    blurred: int = 0
    locked: int = 0
    prose: int = 0
    kept_structure: int = 0
    skipped_images: int = 0
    notes: list = field(default_factory=list)

    def __str__(self):
        out = [f"blurred {self.blurred} value(s)", f"locked {self.locked}",
               f"{self.prose} figure(s) in prose", 
               f"{self.kept_structure} axis/year label(s) kept"]
        if self.skipped_images:
            out.append(f"{self.skipped_images} large picture(s) left alone")
        return "; ".join(out + self.notes)


# ═════════════════════════ rasterising a figure ══════════════════════
_FONT_DIRS = [Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts",
              Path("/usr/share/fonts"), Path("/Library/Fonts"),
              Path.home() / ".fonts"]
_FONT_FILE = {"segoe ui": "segoeui.ttf", "calibri": "calibri.ttf",
              "arial": "arial.ttf", "helvetica": "arial.ttf",
              "times new roman": "times.ttf", "bell mt": "BELL.TTF",
              "verdana": "verdana.ttf", "tahoma": "tahoma.ttf"}
_FONT_CACHE: dict = {}


def _font(name: str, px: int):
    """A PIL font for ``name`` at ``px``, or the default bitmap font."""
    from PIL import ImageFont
    key = (str(name).lower(), px)
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]
    stem = _FONT_FILE.get(str(name or "").strip().lower())
    cands = []
    if stem:
        cands += [d / stem for d in _FONT_DIRS]
    cands += [d / f for d in _FONT_DIRS
              for f in ("segoeui.ttf", "calibri.ttf", "arial.ttf",
                        "DejaVuSans.ttf")]
    for c in cands:
        try:
            if c.exists():
                f = ImageFont.truetype(str(c), px)
                _FONT_CACHE[key] = f
                return f
        except Exception:
            continue
    try:
        f = ImageFont.truetype("DejaVuSans.ttf", px)
    except Exception:
        f = ImageFont.load_default()
    _FONT_CACHE[key] = f
    return f


def render_blur(text: str, height_px: float, style: BlurStyle = None,
                font_name: str = "Segoe UI"):
    """A blurred picture of ``text``, and the size to place it at.

    Returns ``(png_bytes, w_px, h_px)`` - the width follows the text, so a
    blurred "1,234,567" is visibly longer than a blurred "4.1%", which is the
    whole point: the reader sees the SHAPE of what was withheld.
    Returns None when Pillow is missing or the box is too small to bother.
    """
    style = style or BlurStyle()
    text = str(text or "").strip()
    if not text or height_px < style.min_px:
        return None
    try:
        from PIL import Image, ImageDraw, ImageFilter
    except Exception:
        return None

    ss = max(1, int(style.supersample))
    fs = max(6, int(round(height_px * ss)))
    font = _font(font_name, fs)
    dummy = Image.new("L", (1, 1))
    box = ImageDraw.Draw(dummy).textbbox((0, 0), text, font=font)
    tw, th = max(1, box[2] - box[0]), max(1, box[3] - box[1])
    pad = int(round(fs * style.pad_frac))
    W, H = tw + 2 * pad, th + 2 * pad

    ink = style.ink.lstrip("#")
    rgb = tuple(int(ink[i:i + 2], 16) for i in (0, 2, 4))
    img = Image.new("RGBA", (W, H), rgb + (0,))
    ImageDraw.Draw(img).text((pad - box[0], pad - box[1]), text,
                             font=font, fill=rgb + (255,))
    img = img.filter(ImageFilter.GaussianBlur(max(1.0, H * style.blur_frac)))
    # Hold the ink near-black and NORMALISE the opacity to the ceiling: a
    # Gaussian spreads a short string much thinner than a long one, so a flat
    # multiplier would make "4.1%" nearly invisible next to "1,234,567".
    # Scaling each raster to the same PEAK is what keeps every blurred figure
    # reading with the weight the reference deck's rasters have.
    a = img.getchannel("A")
    peak = a.getextrema()[1] or 255
    k = style.alpha * 255.0 / peak
    img.putalpha(a.point(lambda v: min(255, int(v * k))))

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue(), W / ss, H / ss


# ═══════════════════════════ placing it ══════════════════════════════
def _add_blur(slide, text, cx, cy, height_px, style, font_name="Segoe UI",
              max_w=None):
    """Put a blurred picture of ``text`` centred on (cx, cy).  True if drawn.

    ``max_w`` holds the smudge inside its container: a narrative cell's
    sentence is far wider than the cell, and a blur that runs across the
    neighbouring columns reads as damage rather than as a withheld value.
    """
    made = render_blur(text, height_px, style, font_name)
    if not made:
        return False
    png, w, h = made
    if max_w and w > max_w:
        h *= max_w / w
        w = max_w
    slide.shapes.add_picture(io.BytesIO(png),
                             Emu(int((cx - w / 2) * EMU_PER_PX)),
                             Emu(int((cy - h / 2) * EMU_PER_PX)),
                             Emu(int(w * EMU_PER_PX)),
                             Emu(int(h * EMU_PER_PX)))
    return True


def _shape_font(shp, para_i: int, default: float = 14.0):
    """(point size, font name) of a shape's paragraph, as far as it says."""
    size, name = default, "Segoe UI"
    try:
        para = shp.text_frame.paragraphs[para_i]
        for r in para.runs:
            if r.font.size:
                size = r.font.size.pt
            if r.font.name:
                name = r.font.name
            if r.font.size and r.font.name:
                break
    except Exception:
        pass
    return size, name


def _cell_font(cell, default: float = 9.0):
    size, name = pptlock._cell_font_pt(cell, default), "Segoe UI"
    try:
        for para in cell.text_frame.paragraphs:
            for r in para.runs:
                if r.font.name:
                    return size, r.font.name
    except Exception:
        pass
    return size, name


class _BlurAsArt:
    """A stand-in for `pptlock.LockArt` that hands back a BLUR of one value.

    A native chart draws its labels from a number format, so the only way to
    put anything else in a label is as its picture FILL - plumbing `pptlock`
    already has (`_lock_chart_label`). Rather than copy that DrawingML, this
    presents the blurred raster through the same two methods it calls, so a
    chart label is blurred by exactly the code path that would have padlocked
    it.
    """

    def __init__(self, path):
        self._p = path

    def plain(self):
        return self._p

    def labelled(self, unit="", white=False):
        return self._p

    @property
    def available(self):
        return self._p is not None


def _blur_file(text, height_px, style, tmpdir):
    """A blurred raster of ``text`` written to disk; None if it cannot be."""
    made = render_blur(text, height_px, style)
    if not made:
        return None
    png, _w, _h = made
    tmpdir.mkdir(parents=True, exist_ok=True)
    p = tmpdir / f"blur{abs(hash((text, round(height_px, 1)))):x}.png"
    if not p.exists():
        p.write_bytes(png)
    return p


# ═══════════════ what is STRUCTURE and must stay readable ════════════
# A chart without its axis is not a locked chart, it is a broken one.  The
# reference deck leaves 143 such labels readable across 24 slides - its scale
# ticks (0% / 50% / 100%, -451 / +0 / +451), and every year that names a
# column.  `pptlock` cannot see them: to its scan a tick is a text box holding
# one figure, exactly like a KPI value.  The difference is not in the string,
# it is in the ARRANGEMENT - ticks come in threes or more, same width, evenly
# spaced down one edge - so that is what is measured here.
_YEAR_ONLY = re.compile(r"^\s*(?:19|20)\d{2}\s*$")
_YEAR_SPAN = re.compile(
    r"^\s*(?:19|20)\d{2}\s*[-–—/]\s*(?:(?:19|20)?\d{2})\s*$")


def is_structure(text) -> bool:
    """True when a lone figure is STRUCTURE - a year or a span of years.

    `pptlock.prose_figures` already refuses to withhold a bare year inside a
    sentence ("a bare year is structure rather than a figure"); a year on its
    own, as a chart's category label, is the same thing and must survive for
    the same reason.
    """
    s = str(text or "").strip()
    return bool(_YEAR_ONLY.match(s) or _YEAR_SPAN.match(s))


def _even(vals, tol=0.3) -> bool:
    """True when values are evenly spaced - an axis, not a scatter of data."""
    if len(vals) < 3:
        return False
    vals = sorted(vals)
    gaps = [b - a for a, b in zip(vals, vals[1:])]
    mean = sum(gaps) / len(gaps)
    if mean <= 0:
        return False
    return max(abs(g - mean) for g in gaps) <= tol * mean * 1.6


def axis_paths(slide, max_w=70.0, max_h=26.0) -> set:
    """Shape paths that read as axis ticks, and so are left alone.

    A tick is one of three or more short lone figures that share a left edge
    and a width and are evenly spaced down the slide (a value axis), or share
    a top edge and a height and are evenly spaced across it (a category axis).
    Data labels fail the test: the BPS markers on a share-shift page share a
    column but sit wherever their series lands, so their gaps are ragged.
    """
    cols, rows = {}, {}
    for path, shp in pptlock._iter_shapes(slide.shapes):
        if not getattr(shp, "has_text_frame", False):
            continue
        t = shp.text_frame.text.strip()
        if not t or len(t) > 10 or not pptlock._is_lone_figure(t):
            continue
        r = pptlock._abs_rect(shp)
        if r is None:
            continue
        x, y, w, h = r
        if w <= max_w:
            cols.setdefault((round(x / 2), round(w / 2)), []).append((y, path))
        if h <= max_h:
            rows.setdefault((round(y / 2), round(h / 2)), []).append((x, path))
    out = set()
    for group in list(cols.values()) + list(rows.values()):
        if len(group) >= 3 and _even([g[0] for g in group]):
            out.update(p for _v, p in group)
    return out


def totals_only_applies(table, policy) -> bool:
    """True when narrowing the open columns to the total row means anything.

    `LockPolicy.totals_only` says "keep the open years on the TOTAL row only".
    On a size table that reads perfectly: the horizon and the estimate survive
    once, at the bottom, and every breakdown row locks.  On a table with NO
    total row it is a trapdoor - there is no row to keep them on, so the whole
    grid goes dark, including the columns the policy calls open.

    The reference deck draws exactly this line: its share table (Share 2026 /
    Share 2036 / Variance / F-CAGR, no total row) keeps all three open columns
    on every row, while its year-grid table one slide later, which does have a
    Total, keeps them only there.
    """
    if not policy.totals_only:
        return False
    rows = list(table.rows)
    return any(policy.is_total_label(r.cells[0].text) for r in rows[1:])


def _prose_cell(cell, policy) -> int:
    """Withhold the figures INSIDE a phrase cell, leaving the phrase.

    An editorial cell ("143,010 -> 210,400", "Lead: 1.1% Sodium Fluoride") is
    a sentence that happens to start with a digit, not a value.  Blanking it
    loses the sentence AND the protected name inside it; the reference keeps
    both and withholds only the figures, which is `pptlock`'s prose rule
    applied one level down.
    """
    n = 0
    for para in cell.text_frame.paragraphs:
        marks = pptlock._prose_marks(pptlock._para_text(para), policy)
        if not marks:
            continue

        def swap(s, _m=None):
            for a, b, _whole, unit in reversed(
                    pptlock._prose_marks(s, policy)):
                s = s[:a] + policy.placeholder + unit + s[b:]
            return s

        if pptlock._rewrite_paragraph(para, swap):
            n += len(marks)
    return n


# ═══════════ a figure INSIDE a sentence: the inline padlock ══════════
# DrawingML has no inline picture, so a padlock cannot be dropped into a run
# of text.  The reference deck solves it the only way the format allows: it
# takes the paragraph away from PowerPoint and redraws it as its own
# fragments - a text box per stretch of words, a picture where each figure
# stood.  That means OWNING the wrap, which is what `_layout_paragraphs`
# does: measure with the deck's own font, break on the shape's width, and
# emit boxes at the positions measured.  Half-measures do not work here -
# leaving the paragraph in place and covering the figure needs PowerPoint's
# wrap, which cannot be read back.
_PT_PX = 96.0 / 72.0                       # 1 pt at 96 dpi
_LINE_K = 1.22                             # line box, as a multiple of size


def _para_style(para, default_pt=11.0):
    """(size_pt, font name, bold, rgb) as far as the paragraph declares it."""
    size, name, bold, rgb = default_pt, "Segoe UI", None, None
    for r in para.runs:
        f = r.font
        if f.size and size == default_pt:
            size = f.size.pt
        if f.name and name == "Segoe UI":
            name = f.name
        if bold is None:
            bold = f.bold
        if rgb is None:
            try:
                rgb = f.color.rgb
            except Exception:
                rgb = None
        if f.size and f.name:
            break
    return size, name, bool(bold), rgb


def _measure(font, s):
    if not s:
        return 0.0
    return font.getlength(s)


def _layout_paragraphs(shp, policy, style):
    """Lay a text frame out as fragments, with a slot where each figure was.

    Returns ``[(x, y, line_h, size_pt, name, bold, rgb, [(kind, text, w)])]``
    per line, in slide px - ``kind`` being "text" or "lock".  None when the
    shape holds nothing to withhold, or cannot be measured.
    """
    try:
        from PIL import ImageFont            # noqa: F401
    except Exception:
        return None
    rect = pptlock._abs_rect(shp)
    if rect is None:
        return None
    try:
        tf = shp.text_frame
        paras = list(tf.paragraphs)
    except Exception:
        return None
    if not any(pptlock._prose_marks(pptlock._para_text(p), policy)
               for p in paras):
        return None

    x0, y0, w0, h0 = rect
    pad_l = (tf.margin_left or 0) / EMU_PER_PX or 7.5
    pad_r = (tf.margin_right or 0) / EMU_PER_PX or 7.5
    pad_t = (tf.margin_top or 0) / EMU_PER_PX or 3.75
    avail = max(20.0, w0 - pad_l - pad_r)

    lines, y = [], y0 + pad_t
    for para in paras:
        text = pptlock._para_text(para)
        size_pt, name, bold, rgb = _para_style(para)
        px = max(6, int(round(size_pt * _PT_PX)))
        font = _font(name, px)
        line_h = size_pt * _PT_PX * _LINE_K
        if not text.strip():
            y += line_h
            continue

        marks = pptlock._prose_marks(text, policy)
        # tokens: the words between the figures, and a slot for each figure
        tokens, at = [], 0
        for a, b, _whole, unit in marks:
            if a > at:
                tokens.append(("text", text[at:a]))
            tokens.append(("lock", unit))
            at = b
        if at < len(text):
            tokens.append(("text", text[at:]))
        if not marks:
            tokens = [("text", text)]

        lock_w = size_pt * _PT_PX * pptlock.LOCK_ART_ASPECT * 1.05
        cur, cur_w = [], 0.0

        def flush():
            nonlocal cur, cur_w, y
            if cur:
                lines.append((x0 + pad_l, y, line_h, size_pt, name, bold, rgb,
                              list(cur)))
                y += line_h
            cur, cur_w = [], 0.0

        for kind, payload in tokens:
            if kind == "lock":
                if cur_w + lock_w > avail and cur:
                    flush()
                cur.append(("lock", "", lock_w))
                cur_w += lock_w
                if payload:                       # the unit the padlock keeps
                    uw = _measure(font, payload)
                    cur.append(("text", payload, uw))
                    cur_w += uw
                continue
            # words, kept with the space that follows them
            for word in re.findall(r"\S+\s*", payload):
                ww = _measure(font, word)
                if cur_w + ww > avail and cur:
                    flush()
                    word_s = word.lstrip()
                    ww = _measure(font, word_s)
                    word = word_s
                if cur and cur[-1][0] == "text":
                    k, t, wd = cur[-1]
                    cur[-1] = (k, t + word, wd + ww)
                else:
                    cur.append(("text", word, ww))
                cur_w += ww
        flush()
        if y > y0 + h0 + line_h:                 # ran past the shape: give up
            return None
    return lines


def _draw_inline(slide, shp, lines, art, style):
    """Blank the shape's text and redraw the measured lines beside padlocks."""
    from pptx.util import Pt as _Pt
    try:
        for para in shp.text_frame.paragraphs:
            for r in list(para.runs):
                r.text = ""
    except Exception:
        return 0
    placed = 0
    for x, y, line_h, size_pt, name, bold, rgb, parts in lines:
        cx = x
        for kind, text, w in parts:
            if kind == "lock":
                if art and art.available:
                    pptlock._place_lock(slide, art, cx + w / 2,
                                        y + line_h * 0.52, size_pt)
                    placed += 1
            elif text.strip():
                box = slide.shapes.add_textbox(
                    Emu(int(cx * EMU_PER_PX)), Emu(int(y * EMU_PER_PX)),
                    Emu(int(max(w + 4, 6) * EMU_PER_PX)),
                    Emu(int(line_h * 1.25 * EMU_PER_PX)))
                tf = box.text_frame
                tf.word_wrap = False
                tf.margin_left = tf.margin_right = 0
                tf.margin_top = tf.margin_bottom = 0
                run = tf.paragraphs[0].add_run()
                run.text = text
                run.font.size = _Pt(size_pt)
                run.font.name = name
                run.font.bold = bold
                if rgb is not None:
                    run.font.color.rgb = rgb
            cx += w
    return placed


def _clusters(rects, gap):
    """Group rects into panels: two rects joined when their boxes, grown by
    ``gap``, touch. Single-linkage, so a row of KPI cards is one panel and a
    chart on the other side of the slide is another."""
    parent = list(range(len(rects)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def grown(r):
        return (r[0] - gap / 2, r[1] - gap / 2, r[2] + gap, r[3] + gap)

    for i in range(len(rects)):
        ax, ay, aw, ah = grown(rects[i])
        for j in range(i + 1, len(rects)):
            bx, by, bw, bh = grown(rects[j])
            if ax < bx + bw and bx < ax + aw and ay < by + bh and by < ay + ah:
                parent[find(i)] = find(j)
    out = {}
    for i, r in enumerate(rects):
        out.setdefault(find(i), []).append(r)
    return list(out.values())


def _bbox(rects):
    x0 = min(r[0] for r in rects)
    y0 = min(r[1] for r in rects)
    x1 = max(r[0] + r[2] for r in rects)
    y1 = max(r[1] + r[3] for r in rects)
    return x0, y0, x1 - x0, y1 - y0


def _place_container_lock(slide, art, rect, style, rep):
    """ONE padlock, centred on a container that has something withheld."""
    if not (style.lock_containers and art and art.available):
        return
    x, y, w, h = rect
    size = min(style.lock_max, max(style.lock_min, h * style.lock_frac))
    # `_place_lock` sizes the artwork from a TEXT point size, so hand it the
    # size that yields the height we want
    pptlock._place_lock(slide, art, x + w / 2, y + h / 2,
                        size / pptlock.LOCK_IMG_K)
    rep.locked += 1


# ═════════════════════════════ the pass ══════════════════════════════
def blur_pptx(src, dst, policy: LockPolicy = None, style: BlurStyle = None,
              assets_dir=None) -> BlurReport:
    """Withhold every value the policy names, showing a blur under the lock.

    The deck is read once by `pptlock.scan`, so what is withheld here is
    exactly what `pptlock` would withhold - this module only changes what
    stands in the value's place.
    """
    policy = policy or LockPolicy()
    style = style or BlurStyle()
    prs = Presentation(str(src)) if not isinstance(src, Presentation().__class__) \
        else src
    art = LockArt(assets_dir)
    rep = BlurReport()
    if render_blur("1,234", 12, style) is None:
        rep.notes.append("Pillow not available - locked without the blur")

    cands = pptlock.scan(prs, policy)
    slides = list(prs.slides)
    fallback = set()               # prose the inline layout could not measure
    import tempfile
    tmpdir = Path(tempfile.mkdtemp(prefix="pptblur-"))

    # Every container that loses a value earns ONE padlock, collected here as
    # {slide index: [rect, ...]}: tables and charts by their own rect, loose
    # figures by the rect of the panel they cluster into.
    boxed: dict = {}
    loose: dict = {}

    # 0. what must stay readable whatever the policy says: the axis ticks and
    #    category years a chart is unreadable without (see `axis_paths`)
    axes = {}
    if style.keep_structure:
        for n, slide in enumerate(slides, 1):
            paths = axis_paths(slide)
            if paths:
                axes[n] = paths

    def _structural(c) -> bool:
        return (style.keep_structure and c.kind == "shape-value"
                and (is_structure(c.text)
                     or c.shape_path in axes.get(c.slide, ())))

    # A table with no total row must not be narrowed to one (see
    # `totals_only_applies`): its open columns stay open on every row, which
    # `pptlock`'s scan cannot know because it decides column by column.
    reopened = set()
    if policy.totals_only and style.totals_only_needs_total:
        seen = {}
        for c in cands:
            if c.kind != "table-cell":
                continue
            key = (c.slide, c.shape_path)
            if key not in seen:
                try:
                    tbl = pptlock._shape_at(slides[c.slide - 1],
                                            c.shape_path).table
                    hdr = [cell.text.strip() for cell in tbl.rows[0].cells]
                    seen[key] = (totals_only_applies(tbl, policy),
                                 [policy.is_open_label(h) for h in hdr])
                except (IndexError, AttributeError):
                    seen[key] = (True, [])
            narrows, open_col = seen[key]
            col = c.ref[1]
            if not narrows and col < len(open_col) and open_col[col]:
                reopened.add(c.key)

    # 1. the kinds this module draws itself, in ONE pass per shape so a table
    #    is emptied once and its blurs placed from one set of row heights
    mine = []
    for c in cands:
        if c.kind not in ("table-cell", "shape-value"):
            continue
        if c.key in reopened:
            rep.kept_structure += 1
            continue
        if _structural(c):
            rep.kept_structure += 1
            continue
        # a cell that is a PHRASE is a sentence, not a value.  With
        # `narrative_cells="prose"` its figures become "xx" and the words
        # stay; the default withholds the cell whole (see the pass below).
        if c.kind == "table-cell" and style.narrative_cells == "prose" \
                and not pptlock._is_lone_figure(c.text):
            slide = slides[c.slide - 1]
            try:
                cell = pptlock._shape_at(slide, c.shape_path).table.cell(*c.ref)
            except (IndexError, AttributeError):
                continue
            rep.prose += _prose_cell(cell, policy)
            continue
        mine.append(c)

    by_shape = {}
    for c in mine:
        by_shape.setdefault((c.slide, c.shape_path, c.kind), []).append(c)

    for (slide_no, shape_path, kind), group in by_shape.items():
        slide = slides[slide_no - 1]
        try:
            shp = pptlock._shape_at(slide, shape_path)
        except (IndexError, AttributeError):
            continue
        rect = pptlock._abs_rect(shp)
        if rect is None:
            continue

        if kind == "table-cell":
            tbl = shp.table
            x0, y0 = rect[0], rect[1]
            for c in group:
                r, col = c.ref
                cell = tbl.cell(r, col)
                size, fname = _cell_font(cell)
                pptlock._blank_cell(cell)
                cx = x0 + pptlock._col_x(tbl, col) + \
                    tbl.columns[col].width / EMU_PER_PX / 2
                cy = y0 + pptlock._row_y(tbl, r) + \
                    tbl.rows[r].height / EMU_PER_PX / 2
                # a cell's height is fixed, so the blur is sized off the FONT
                if style.blur_values and \
                        _add_blur(slide, c.text, cx, cy, size, style, fname):
                    rep.blurred += 1
                if style.lock_values:
                    pptlock._place_lock(slide, art, cx, cy, size, c.unit)
                    rep.locked += 1
            # the whole table is what the padlock speaks for
            boxed.setdefault(slide_no, []).append(rect)
        else:                                            # shape-value
            c = group[0]
            vi = c.ref[0]
            size, fname = _shape_font(shp, vi)
            try:
                para = shp.text_frame.paragraphs[vi]
                npara = len(shp.text_frame.paragraphs)
                for r in list(para.runs):
                    r.text = ""
            except Exception:
                continue
            cy = rect[1] + rect[3] * (vi + 0.5) / max(1, npara)
            cx = rect[0] + rect[2] / 2
            if style.blur_values and \
                    _add_blur(slide, c.text, cx, cy, size, style, fname):
                rep.blurred += 1
            if style.lock_values:
                pptlock._place_lock(slide, art, cx, cy, size, c.unit)
                rep.locked += 1
            loose.setdefault(slide_no, []).append(rect)

    # 1b. narrative cells - a sentence in a cell that carries figures
    #     ("Scale base (58.3% share), lead formulation ...").  `pptlock.scan`
    #     never sees these: its `is_value` wants a cell to START with a digit,
    #     and these start with a word.  A cell has no room to redraw a
    #     paragraph around an inline padlock, so the cell goes whole.
    if style.narrative_cells == "blank":
        already = {(c.slide, c.shape_path, c.ref) for c in cands
                   if c.kind == "table-cell"}
        for slide_no, slide in enumerate(slides, 1):
            for path, shp in pptlock._iter_shapes(slide.shapes):
                if not getattr(shp, "has_table", False):
                    continue
                tbl = shp.table
                rect = pptlock._abs_rect(shp)
                if rect is None:
                    continue
                touched = False
                for r in range(1, len(tbl.rows)):
                    for c in range(1, len(tbl.columns)):
                        if (slide_no, path, (r, c)) in already:
                            continue
                        cell = tbl.cell(r, c)
                        txt = cell.text.strip()
                        if not txt or pptlock._is_lone_figure(txt):
                            continue
                        if not pptlock._prose_marks(txt, policy):
                            continue          # no figures in it: nothing to hide
                        size, fname = _cell_font(cell)
                        pptlock._blank_cell(cell)
                        cx = rect[0] + pptlock._col_x(tbl, c) +                             tbl.columns[c].width / EMU_PER_PX / 2
                        cy = rect[1] + pptlock._row_y(tbl, r) +                             tbl.rows[r].height / EMU_PER_PX / 2
                        if style.blur_values and _add_blur(
                                slide, txt, cx, cy, size, style, fname,
                                max_w=tbl.columns[c].width / EMU_PER_PX * 0.92):
                            rep.blurred += 1
                        touched = True
                if touched:
                    boxed.setdefault(slide_no, []).append(rect)

    # 2. chart labels: blurred through pptlock's own label plumbing, which is
    #    the only way anything but a number reaches a native chart's label
    if style.blur_values:
        by_chart = {}
        for c in cands:
            if c.kind == "chart-label":
                by_chart.setdefault((c.slide, c.shape_path), []).append(c)
        for (slide_no, shape_path), group in by_chart.items():
            slide = slides[slide_no - 1]
            try:
                shp = pptlock._shape_at(slide, shape_path)
                chart = shp.chart
            except (IndexError, AttributeError):
                continue
            pairs = [(p, s) for p in chart.plots for s in p.series]
            for c in group:
                si, pi = c.ref
                if si >= len(pairs):
                    continue
                plot, series = pairs[si]
                try:
                    point = series.points[pi]
                except (IndexError, AttributeError):
                    continue
                png = _blur_file(c.text + (f" {c.unit}" if c.unit else ""),
                                 11.0, style, tmpdir)
                if png is None:
                    continue
                pptlock._ensure_series_dlbls(plot, series)
                if pptlock._lock_chart_label(chart, point, _BlurAsArt(png)):
                    rep.blurred += 1
            r = pptlock._abs_rect(shp)
            if r:
                boxed.setdefault(slide_no, []).append(r)

    # 3. prose: a figure inside a sentence.  "inline-lock" redraws the whole
    #    paragraph as fragments with a padlock where the figure stood - the
    #    only way the format allows one - and "xx" falls back to `pptlock`.
    if policy.lock_prose:
        if style.prose == "inline-lock":
            shapes_done = set()
            for c in cands:
                if c.kind != "prose":
                    continue
                key = (c.slide, c.shape_path)
                if key in shapes_done:
                    continue
                shapes_done.add(key)
                slide = slides[c.slide - 1]
                try:
                    shp = pptlock._shape_at(slide, c.shape_path)
                except (IndexError, AttributeError):
                    continue
                lines = _layout_paragraphs(shp, policy, style)
                if lines is None:                 # unmeasurable: fall back
                    fallback.add(c.key)
                    continue
                rep.prose += _draw_inline(slide, shp, lines, art, style)
        else:
            fallback |= {c.key for c in cands if c.kind == "prose"}
    if fallback:
        sub = pptlock.apply_locks(prs, None, policy=policy,
                                  decisions=fallback, assets_dir=assets_dir)
        rep.locked += sub.locked
        rep.skipped_images = sub.skipped_images

    # 4. ONE padlock per container: every table and chart that lost a value,
    #    and every panel the loose figures cluster into
    for slide_no in set(list(boxed) + list(loose)):
        slide = slides[slide_no - 1]
        for rect in boxed.get(slide_no, []):
            _place_container_lock(slide, art, rect, style, rep)
        for cluster in _clusters(loose.get(slide_no, []), style.cluster_gap):
            _place_container_lock(slide, art, _bbox(cluster), style, rep)

    if dst:
        prs.save(str(dst))
    return rep


# ═════════════════════════════════ CLI ═══════════════════════════════
def _main(argv=None):
    ap = argparse.ArgumentParser(
        description="Lock a deck's values, showing a blur under the padlock.")
    ap.add_argument("src")
    ap.add_argument("dst", nargs="?")
    ap.add_argument("--open-years", default="")
    ap.add_argument("--lock-cagr", action="store_true")
    ap.add_argument("--totals-only", action="store_true",
                    help="in tables, keep the open years on the TOTAL row only")
    ap.add_argument("--no-prose", action="store_true")
    ap.add_argument("--protect", default="")
    ap.add_argument("--assets", default=None)
    ap.add_argument("--lock-values", action="store_true",
                    help="also put a padlock ON each blurred value "
                         "(default: one padlock per table / chart / panel)")
    ap.add_argument("--no-container-lock", action="store_true")
    a = ap.parse_args(argv)

    policy = LockPolicy(
        open_years={int(y) for y in a.open_years.replace(" ", "").split(",")
                    if y},
        open_cagr=not a.lock_cagr, totals_only=a.totals_only,
        lock_prose=not a.no_prose,
        protect=tuple(p.strip() for p in a.protect.split(";") if p.strip()))
    style = BlurStyle(lock_values=a.lock_values,
                      lock_containers=not a.no_container_lock)
    dst = a.dst or str(Path(a.src).with_name(
        Path(a.src).stem + " (locked).pptx"))
    rep = blur_pptx(a.src, dst, policy=policy, style=style, assets_dir=a.assets)
    print(f"open years: {sorted(policy.open_years) or '(none)'}")
    print(rep)
    print(f"saved: {dst}")
    return 0


if __name__ == "__main__":
    sys.exit(_main())
