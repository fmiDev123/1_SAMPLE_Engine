"""Native (fully editable) PowerPoint export.

Instead of page screenshots, every page is rebuilt from structured
specs into real PowerPoint objects at the same PBIX coordinates:

* titles, KPI cards, commentary  -> editable text boxes / shapes
* pivot tables                   -> native PowerPoint tables
* charts                         -> native charts (right-click 'Edit
                                    Data' works: bars, lines, donuts,
                                    stacked columns, bubbles, waterfall
                                    via the stacked-column technique)
* navigation                     -> a hyperlinked Contents slide plus
                                    Menu / Prev / Next buttons on every
                                    slide (like the PBIX action buttons)

No server or browser is needed - the deck builds straight from data.
"""
from __future__ import annotations

import functools
import math
import sys
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.chart.data import BubbleChartData, CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (XL_CHART_TYPE, XL_DATA_LABEL_POSITION,
                             XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION)
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu, Pt

from . import (components, config, data, exhibits, figures, geo, layout,
               subheaders, toc)
from .config import (ACCENT, ARROW_HIGH, ARROW_LOW, ARROW_MED, BAD,
                     BOX_COLORS, CHIP_EDGE, CHIP_FILL, CHIP_TEXT,
                     COVER_ACCENT, COVER_BAR, COVER_BG,
                     NAV_BTN_BG, NAV_BTN_EDGE, NAV_BTN_TEXT,
                     COVER_ON_LIGHT, COVER_TEXT, DIVIDER_BODY, DIVIDER_KICKER,
                     DIVIDER_MARK, DIVIDER_PANEL, DIVIDER_TITLE,
                     SECTION_BAND_TEXT,
                     COVER_DOT, COVER_GLOW, COVER_RING, COVER_SUB,
                     DATA_COLORS, DN_FILL, DN_LINE, FLOW_LEAF_EDGE,
                     FLOW_LEAF_TEXT, FLOW_LINE, FOOTER_BAND, FOOTER_BAND_H,
                     FOOTER_MARGIN, FOOTER_RULE_Y, FOOTER_TEXT_COLOR,
                     GOOD, GRID, HEADER_BG,
                     HEADER_FILL, HEADER_RULE, HEADER_RULE_SEGMENTS,
                     HEADER_STYLE, HEAT_STOPS,
                     AMBER_MARK, F3_BODY, F3_LBL, F3_RIM, F3_TOP,
                     GAUGE_TRACK, KPI_TINT, MUTED, NAVY_DEEP, NEUTRAL_LINE,
                     NOT_OK_COLOR, OK_COLOR, ON_DARK_MUTED, ORANGE,
                     OUTPACE_COLOR, BELOW_COLOR, DECLINE_COLOR,
                     SIG_HEAD_BG, SIG_ZEBRA, TOC_TILE_BG, TOC_TILE_TEXT,
                     TABLE_HEAD_BG,
                     TABLE_HEAD_TEXT, TABLE_TOTAL_BG,
                     TA_ACTUAL, TA_TARGET, WARN_TEXT,
                     PAGE_BG, PANEL_EDGE, PANEL_RADIUS, PANEL_ROUNDED,
                     PPT_FONT, PPT_THEME_COLORS, PPT_THEME_FONTS,
                     PPT_THEME_NAME,
                     REF_LINE, RULE_H, RULE_Y,
                     SECTION_BAND, SOURCE_TEXT, STAR_COLORS, STRIP_LINE,
                     STRIP_START, TEXT, TITLE_COLOR, UNIT_LABELS, UP_FILL,
                     UP_LINE, VC_NAVY, VC_PURPLE, YEARS, ZEBRA)
from . import theming
from .registry import (COMPETITOR_REPORT, COUNTRY_REPORT, DYNAMICS_REPORT,
                       HOME_SLUG, MAIN_REPORT, PAGES, SEC_PRICING,
                       SEC_SUPPLY, _DYN_FORCES, contents_buttons)

_EMU_PER_PX = 9525                     # 1280px canvas == 13.333in slide

# Horizontal room (px) a clustered bar needs before its value label is laid
# flat.  An 11pt "231.6K" is ~44px wide; below that the label is stood on
# end, which needs only its line height (~17px) - and below THAT the bars are
# too narrow for any 11pt label and carry none.  See _add_cat_chart.
_MIN_PX_PER_BAR_LABEL = 44
_MIN_PX_PER_ROTATED_LABEL = 17
# whether the chart _add_cat_chart last drew got its per-point numbers - a
# chart that did not must keep its value axis, or it has no reading at all
_LAST_LABELLED = {"v": True}
# A stacked segment shorter than this cannot hold its 11pt number inside it
# (a horizontal bar's segment needs the number's width instead).
_MIN_PX_PER_SEGMENT_LABEL = 16
_MIN_PX_PER_SEGMENT_LABEL_H = 40

def _hide_point_label(series, j):
    """Switch off the data label of point ``j`` of ``series`` alone."""
    try:
        dlbl = series.points[j].data_label._get_or_add_dLbl()
        for child in list(dlbl):
            if child.tag != qn("c:idx"):
                dlbl.remove(child)
        dlbl.append(parse_xml(f'<c:delete {nsdecls("c")} val="1"/>'))
    except Exception:
        pass

def _skip_thin_labels(chart, series, plot_px, pct100=False,
                      horizontal=False, width_px=None):
    """Blank the labels of stacked segments too small to hold an 11pt number.

    ``plot_px`` is the plot's extent along the value axis - its height for
    columns; for horizontal bars ``width_px`` is used instead."""
    need = _MIN_PX_PER_SEGMENT_LABEL
    if horizontal and width_px:
        plot_px, need = width_px, _MIN_PX_PER_SEGMENT_LABEL_H
    cols = list(zip(*[s[1] for s in series])) if series else []
    tots = [sum(abs(_num(v) or 0.0) for v in col) for col in cols]
    top = max(tots, default=0.0) or 1.0
    for si, s in enumerate(chart.series):
        if si >= len(series):
            break
        for j, v in enumerate(series[si][1]):
            base = (tots[j] if pct100 else top) if j < len(tots) else top
            extent = abs(_num(v) or 0.0) / (base or 1.0) * plot_px
            if extent < need:
                _hide_point_label(s, j)

# A font size is POINTS in PowerPoint but the canvas is measured in slide px
# (1 px = 0.75 pt), so text set at `s` pt occupies ~s * _PPT_FONT_K px of
# width in layout.py's CSS-px wrap model.  Any px-vs-pt measurement in this
# module goes through this factor -- it is the same constant
# templates_dynamics budgets its primitive blocks with.
_PPT_FONT_K = 1.35

# PowerPoint's NATURAL single line height, as a multiple of the font em.  A
# `line=` multiplier set on a paragraph is an lnSpc PERCENTAGE, and PowerPoint
# applies it to this natural height -- not to the em -- so a nominal 1.1
# spacing really occupies 1.1 * 1.2 ems.  Height budgets that missed this ran
# ~20% short and let long bullet text overrun its card.
_PPT_LINE_NAT = 1.2
SLIDE_W = Emu(1280 * _EMU_PER_PX)
SLIDE_H = Emu(720 * _EMU_PER_PX)
FONT = PPT_FONT

def _rgb(hexstr: str) -> RGBColor:
    return RGBColor.from_string(hexstr.lstrip("#"))

_C_HEADER, _C_ACCENT = _rgb(HEADER_BG), _rgb(ACCENT)
_C_TEXT, _C_MUTED = _rgb(TEXT), _rgb(MUTED)
_C_GOOD, _C_BAD = _rgb(GOOD), _rgb(BAD)
_C_EDGE, _C_PAGE = _rgb(PANEL_EDGE), _rgb(PAGE_BG)
_C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
_C_ZEBRA = _rgb(ZEBRA)
_C_GRID = _rgb(GRID)                    # light grey chart gridlines

# ── Brand (footer + logo on every slide) ─────────────────────────────
# Both wordmarks carry a white ground, so each is placed on a small white card
# on the header band (see _top_left_logo).  Which one, and the copyright line
# beside it, come from the ACTIVE template - `theming.apply()` refreshes these
# through `_retheme()`.
_LOGO_PATH = theming.logo_path()
_LOGO_AR = theming.theme().LOGO_AR              # wordmark w / h
_COPYRIGHT = theming.theme().COPYRIGHT

def _readable_on(col: RGBColor) -> RGBColor:
    """White or dark text, whichever reads better on a fill colour."""
    r, g, b = col[0], col[1], col[2]
    lum = 0.299 * r + 0.587 * g + 0.114 * b
    return _C_TEXT if lum > 150 else _C_WHITE

def E(px: float) -> Emu:
    return Emu(int(round(px * _EMU_PER_PX)))

def _num(v):
    try:
        f = float(v)
        return None if pd.isna(f) else f
    except (TypeError, ValueError):
        return None

def _palette(i: int) -> RGBColor:
    return _rgb(DATA_COLORS[i % len(DATA_COLORS)])

def _trace_color(trace, i):
    mk = getattr(trace, "marker", None)
    col = getattr(mk, "color", None) if mk is not None else None
    if isinstance(col, str) and col.startswith("#"):
        return _rgb(col)
    ln = getattr(trace, "line", None)
    lcol = getattr(ln, "color", None) if ln is not None else None
    if isinstance(lcol, str) and lcol.startswith("#"):
        return _rgb(lcol)
    return _palette(i)

# ── the content type standard ────────────────────────────────────────
# Every piece of CONTENT text in the deck - body copy, bullets, table cells,
# chart labels, captions - is set at 11pt or larger (client standard,
# 2026-09-11), so nothing has to be zoomed to be read.  Headings keep their
# own, larger sizes; only slide chrome (the copyright line, the slicer chips)
# may sit below it.  `_pt` is the one gate every content size goes through.
MIN_CONTENT_PT = 11.0

def _pt(size) -> float:
    """A content font size, raised to the deck's 11pt standard."""
    return max(float(size), MIN_CONTENT_PT)

# ── low-level shape helpers ──────────────────────────────────────────
def _set_text(tf, runs, align=PP_ALIGN.LEFT, anchor=None,
              pad=(3, 3, 1, 1), space_after=None, line=None, floor=True):
    """runs: list of paragraphs; each paragraph = list of
    (text, size, bold, color) tuples.

    ``pad`` = (left, right, top, bottom) internal margins in px - replaces
    python-pptx's chunky 0.1"/0.05" defaults so text aligns predictably with
    the shapes/rules placed around it (a big source of the 'raw' look).
    ``space_after`` (pt) adds breathing room between bullet paragraphs;
    ``line`` sets line spacing (multiplier).  ``floor`` raises every size to
    MIN_CONTENT_PT; only slide chrome passes False."""
    tf.word_wrap = True
    tf.margin_left, tf.margin_right = E(pad[0]), E(pad[1])
    tf.margin_top, tf.margin_bottom = E(pad[2]), E(pad[3])
    if anchor is not None:
        tf.vertical_anchor = anchor
    first = True
    for para in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if space_after is not None:
            p.space_after, p.space_before = Pt(space_after), Pt(0)
        if line is not None:
            p.line_spacing = line
        for text, size, bold, color in para:
            r = p.add_run()
            r.text = text
            r.font.size = Pt(_pt(size) if floor else size)
            r.font.bold = bold
            r.font.color.rgb = color
            r.font.name = FONT

def _white_panel(slide, x, y, w, h):
    """The card every panel sits in - rounded where the template asks for it.

    DROT builds its content out of rounded cards with a 0.75pt hairline;
    Fact.MR's PBIX cards are near-square.  `PANEL_RADIUS` decides, and the
    corner is set as a FRACTION of the short side (PowerPoint's own units)
    so a tall panel and a squat KPI card round by the same visual amount
    rather than the same percentage of themselves.
    """
    if PANEL_ROUNDED and PANEL_RADIUS > 0:
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     E(x), E(y), E(w), E(h))
        try:
            shp.adjustments[0] = min(0.5, (PANEL_RADIUS * 0.75)
                                     / max(1.0, min(w, h) * 0.75))
        except (IndexError, ValueError):
            pass
    else:
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                     E(x), E(y), E(w), E(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = _C_WHITE
    shp.line.color.rgb = _C_EDGE
    shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    return shp

def _textbox(slide, x, y, w, h):
    return slide.shapes.add_textbox(E(x), E(y), E(w), E(h))

_TITLE_H = 22

def _panel_frame(slide, spec):
    """White panel + optional title. Returns content rect (px)."""
    x, y, w, h = spec["x"], spec["y"], spec["w"], spec["h"]
    _white_panel(slide, x, y, w, h)
    if spec.get("title"):
        clean = _clean_title(spec["title"])
        # reserve up to 3 lines when a long title would wrap, so the chart/table
        # below never overlaps it.  Measured with the same wrap model the text
        # is actually set with - the old flat ~6.4px-per-char estimate
        # under-counted wide titles, which then wrapped onto the chart.
        lines = max(1, min(3, layout._wrap_count(clean, max(1, w - 24),
                                                 12 * _PPT_FONT_K)))
        th = _TITLE_H + (lines - 1) * 18
        tb = _textbox(slide, x + 10, y + 6, w - 20, th)
        # colour the YoY legend keywords (▲ green / ▶ orange / ▼ red); the same
        # split the dashboard uses.  Everything else stays black.
        runs = [(t, 12, True, _rgb(c))
                for t, c in components.title_runs(clean)]
        _set_text(tb.text_frame, [runs], pad=(0, 2, 1, 1), line=1.0)
        return x + 10, y + th + 8, w - 20, h - th - 14
    return x + 8, y + 8, w - 16, h - 16

# ── title / label helpers ────────────────────────────────────────────
import re as _re

# interactive-only hints that make no sense in a static deck
_HINT_RE = _re.compile(
    r"\s*[·|]\s*(click|hover|tap|drag|select|drill)[^·|]*", _re.IGNORECASE)

def _clean_title(t):
    """Strip dashboard-only interaction hints (e.g. '· click a bubble to
    filter') from a title for the PPT."""
    if not t:
        return t
    return _HINT_RE.sub("", str(t)).strip(" ·|")

def _abbrev_floor() -> float:
    """The magnitude below which numbers are printed in the workbook's own
    unit rather than abbreviated - see `config.abbrev_floor`.  Read per call,
    never cached: the units are rebound per workbook (and per session in the
    portal), so a value frozen at import time would describe the wrong deck."""
    try:
        return float(config.abbrev_floor())
    except Exception:
        return 1e3

def _num_format(values):
    """A PowerPoint number-format code for data labels, chosen from the value
    magnitudes: percentages as %, and big numbers abbreviated to K / M / B so
    labels stay short and don't overlap.  One consistent scale per chart.

    The abbreviation only goes as low as the unit labels allow: a chart of
    'US$ Mn' figures prints 50,531, not '50.5K' (see `config.abbrev_floor`)."""
    vals = [abs(v) for v in (_num(x) for x in values) if v is not None]
    mx = max(vals) if vals else 0
    floor = _abbrev_floor()
    if mx == 0:
        return "0"
    if mx <= 1.5:
        return "0.0%"                      # fractions shown as %
    if mx >= 1e9 and floor <= 1e9:
        return '#,##0.0,,,"B"'
    if mx >= 1e6 and floor <= 1e6:
        return '#,##0.0,,"M"'
    if mx >= 1e4 and floor <= 1e3:
        return '#,##0.0,"K"'
    if mx >= 100:
        return "#,##0"
    return "0.0"

def _axis_num_format(values):
    """Abbreviated number-format code for a numeric AXIS so big tick numbers
    read '30K' / '2M' / '1.5B' instead of 30000 / 2000000.

    Where the unit label already carries the magnitude the ticks keep the full
    number instead, grouped with separators - '970,124' under a caption that
    says US$ Mn, never '970.1K'."""
    vals = [abs(v) for v in (_num(x) for x in values) if v is not None]
    mx = max(vals) if vals else 0
    floor = _abbrev_floor()
    if mx >= 1e9 and floor <= 1e9:
        return '#,##0,,,"B"'
    if mx >= 1e6 and floor <= 1e6:
        return '#,##0,,"M"'
    if mx >= 1e3:
        return '#,##0,"K"' if floor <= 1e3 else "#,##0"
    return None                                # small numbers: leave as-is

def _pct_axis(fig, horizontal=False):
    """True when the figure's VALUE axis carries percentages.

    Read off the plotly figure rather than guessed from the numbers.  A CAGR
    printed as "6.1" and one printed as "6.1%" are not the same claim on a
    client slide, and the dashboard has always printed the sign - it was the
    deck that dropped it, because a native chart only ever saw the bare
    floats."""
    ax = getattr(fig.layout, "xaxis" if horizontal else "yaxis", None)
    if ax is not None:
        if str(getattr(ax, "ticksuffix", "") or "").strip() == "%":
            return True
        if "%" in str(getattr(ax, "tickformat", "") or ""):
            return True
    for t in fig.data:                      # every label ends in "%"
        txt = getattr(t, "text", None)
        if txt is None or isinstance(txt, str):
            continue
        seen = [str(v) for v in txt if v not in (None, "")]
        if seen and all(v.rstrip().endswith("%") for v in seen):
            return True
    return False

def _pct_format(values, decimals=1):
    """PowerPoint number format for a percentage series.  Values already
    scaled to 0-100 need a LITERAL per-cent sign; true fractions take the
    built-in % token, which scales by 100 itself."""
    vals = [abs(v) for v in (_num(x) for x in values) if v is not None]
    d = "0" * decimals
    dot = "." if decimals else ""
    return (f"0{dot}{d}%" if (max(vals) if vals else 0) <= 1.5
            else f'0{dot}{d}"%"')

def _val_axis_range(fig, horizontal=False):
    """The value axis's zero-based range as the dashboard sets it, or None.

    PowerPoint autoscales, and on a tight series that re-bases the axis: three
    CAGRs of 6.1 / 7.0 / 6.7 came out on a 5.6-7.2 axis, where a 0.9-point
    spread reads as a threefold gap.  Where the dashboard has pinned the axis
    to zero the deck follows it, so the bars stay proportional."""
    ax = getattr(fig.layout, "xaxis" if horizontal else "yaxis", None)
    rng = getattr(ax, "range", None) if ax is not None else None
    try:
        if rng and len(rng) == 2 and float(rng[0]) == 0.0 and rng[1]:
            return 0.0, float(rng[1])
    except (TypeError, ValueError):
        pass
    return None

def _sig_num(v, sig=3):
    """Format a number to ``sig`` significant digits, choosing the decimals
    from its magnitude.

    A fixed ``.1f`` is fine for values of order 1 but destroys small ones: a
    market priced at 0.00014 US$/unit rendered every axis tick and label as
    "0.0".  Deriving the decimals keeps the figure meaningful whatever the
    workbook's units."""
    v = _num(v)
    if v is None:
        return ""
    if v == 0:
        return "0"
    d = sig - 1 - int(math.floor(math.log10(abs(v))))
    return f"{v:,.{max(0, min(d, 8))}f}"

def _abbr(v):
    """Human-readable short number for an end-of-line label.

    Abbreviated only as far as the workbook's unit labels allow, so a legend
    beside an 'US$ Mn' caption reads 32,163 / 28,235 / 5,520 rather than
    mixing '32K' with '5,520' in the same six items."""
    v = _num(v)
    if v is None:
        return ""
    a = abs(v)
    floor = _abbrev_floor()
    if a >= 1e9 and floor <= 1e9:
        return f"{v / 1e9:.1f}B"
    if a >= 1e6 and floor <= 1e6:
        return f"{v / 1e6:.1f}M"
    if a >= 1e4 and floor <= 1e3:
        return f"{v / 1e3:.0f}K"
    if a >= 100:
        return f"{v:,.0f}"
    return f"{v:.1f}"

def _label_series_ends(chart, series):
    """Label ONLY the last point of each line with its value, coloured to match
    the series and pushed to the right (outside the plot).  Keeps a crowded
    multi-line chart readable - every value still ties to its segment by colour,
    and the full data table sits on the same slide."""
    for i, s in enumerate(chart.series):
        vals = list(series[i][1])
        li = next((j for j in range(len(vals) - 1, -1, -1)
                   if _num(vals[j]) is not None), None)
        if li is None:
            continue
        col = series[i][2] or _palette(i)
        try:
            dl = s.points[li].data_label
            tf = dl.text_frame
            tf.text = _abbr(vals[li])
            r = tf.paragraphs[0].runs[0]
            r.font.size = Pt(8)
            r.font.bold = True
            r.font.name = FONT
            r.font.color.rgb = col
            dl.position = XL_DATA_LABEL_POSITION.RIGHT
        except Exception:
            pass

def _label_all_line_points(chart, series, fmt):
    """Label every point of a few-line chart, coloured to its line and placed
    above/below alternately so a close pair of lines doesn't collide."""
    for i, s in enumerate(chart.series):
        col = series[i][2] or _palette(i)
        try:
            dls = s.data_labels
            dls.show_value = True
            dls.number_format = fmt
            dls.number_format_is_linked = False
            dls.font.size = Pt(6.5)
            dls.font.bold = True
            dls.font.name = FONT
            dls.font.color.rgb = col
            dls.position = (XL_DATA_LABEL_POSITION.ABOVE if i % 2 == 0
                            else XL_DATA_LABEL_POSITION.BELOW)
        except Exception:
            pass

def _label_big_stacked_segments(chart, series):
    """In a many-band stacked column, label only the segments that are tall
    enough to hold their number centred; leave the thin bands blank so labels
    never pile up (their values appear in the list beside the chart)."""
    cols = list(zip(*[s[1] for s in series])) if series else []
    axis_max = max((sum(abs(_num(v) or 0) for v in col) for col in cols),
                   default=0)
    thresh = 0.06 * axis_max                    # ~6% of the tallest bar
    for si, s in enumerate(chart.series):
        vals = series[si][1]
        col = series[si][2] or _palette(si)
        ink = _readable_on(col)
        for j, v in enumerate(vals):
            val = _num(v) or 0
            if abs(val) < thresh:
                continue                        # too thin -> no inside label
            try:
                dl = s.points[j].data_label
                tf = dl.text_frame
                tf.text = _abbr(val)
                r = tf.paragraphs[0].runs[0]
                r.font.size = Pt(7.5)
                r.font.name = FONT
                r.font.color.rgb = ink
                dl.position = XL_DATA_LABEL_POSITION.CENTER
            except Exception:
                pass

def _stacked_center_labels(chart, series, fmt, size):
    """Value labels centred in each stacked segment, coloured light or dark to
    contrast with that segment's own fill - so numbers on dark bars are white."""
    for si, s in enumerate(chart.series):
        col = series[si][2] if si < len(series) and series[si][2] else _palette(si)
        try:
            dls = s.data_labels
            dls.show_value = True
            dls.number_format = fmt
            dls.number_format_is_linked = False
            dls.font.size = Pt(size)
            dls.font.name = FONT
            dls.font.color.rgb = _readable_on(col)
            dls.position = XL_DATA_LABEL_POSITION.CENTER
            _dlbls_horizontal(dls)
        except Exception:
            pass

def _keep_negative_fill(chart):
    """Stop PowerPoint blanking the bars that fall below zero.

    ``invertIfNegative`` defaults to ON, so a diverging bar chart comes out
    with its negative bars drawn as white outlines - the share LOST, which is
    exactly the half a reader must see, disappears."""
    def switch_off(el, after):
        if el.find(qn("c:invertIfNegative")) is not None:
            return
        flag = parse_xml(f'<c:invertIfNegative {nsdecls("c")} val="0"/>')
        ref = None
        for tag in after:
            ref = el.find(qn(tag))
            if ref is not None:
                break
        if ref is None:
            el.append(flag)
        else:
            el.insert(list(el).index(ref), flag)

    for ser in chart._chartSpace.iter(qn("c:ser")):
        switch_off(ser, ("c:dPt", "c:dLbls", "c:cat", "c:val"))
        # a per-point colour writes its own <c:dPt>, which carries its OWN
        # copy of the flag - the series-level one alone leaves it inverted
        for dpt in ser.findall(qn("c:dPt")):
            switch_off(dpt, ("c:marker", "c:bubble3D", "c:explosion",
                             "c:spPr"))

def _add_data_labels(chart, fmt, pos=None, size=6, rotation=0):
    """Show clean numeric data labels on every point of a native chart.

    ``rotation`` in degrees; -90 stands the numbers on end so a dense clustered
    chart can label every bar without the values colliding."""
    try:
        for plot in chart.plots:
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_value = True
            dl.show_series_name = False
            dl.show_category_name = False
            dl.show_legend_key = False
            dl.number_format = fmt
            dl.number_format_is_linked = False
            dl.font.size = Pt(size)
            dl.font.name = FONT
            if pos is not None:
                try:
                    dl.position = pos
                except Exception:
                    pass
            _dlbls_rotation(dl, rotation)
    except Exception:
        pass

# ── chart building ───────────────────────────────────────────────────
def _style_chart(chart, n_series, legend=None, font=MIN_CONTENT_PT,
                 height_px=None):
    try:
        chart.font.size = Pt(_pt(font))
        chart.font.name = FONT
    except Exception:
        pass
    chart.has_title = False
    try:
        # varyColors on a single-series chart makes PowerPoint list
        # every category in the legend - always off for cat charts
        chart.plots[0].vary_by_categories = False
    except Exception:
        pass
    # Bar spacing.  A chart with one or two categories would otherwise draw
    # its bars as full-width slabs, so the gap is derived from the category
    # count by the same rule as figures.cap_bar_widths on the dashboard.
    try:
        n_cats = len(list(chart.plots[0].categories))
    except Exception:
        n_cats = 0
    _apply_bar_gap(chart, n_cats)
    show_legend = legend if legend is not None else n_series > 1
    chart.has_legend = show_legend
    if show_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(MIN_CONTENT_PT)
    for axis_name in ("category_axis", "value_axis"):
        try:
            ax = getattr(chart, axis_name)
            ax.tick_labels.font.size = Pt(MIN_CONTENT_PT)
            ax.format.line.color.rgb = _C_EDGE
            ax.has_major_gridlines = False       # no gridlines in the PPT charts
            ax.has_minor_gridlines = False
        except Exception:
            pass
    # Draw EVERY category label - WHEN THEY FIT.  Left to itself PowerPoint
    # decides how many a category axis has room for and silently draws every
    # Nth: on the Key Segments page that showed three sub-segment names out of
    # six, which reads as missing data rather than as a tidied axis.  Forcing
    # them unconditionally is worse - seventeen names in a 142px panel land on
    # top of each other - so the panel's own height decides, at the ~16px an
    # 11pt label needs.  Without a height the axis is left as it was.
    _fits = (height_px is not None and n_cats > 0
             and height_px / n_cats >= 16.0)
    if not _fits:
        return
    try:
        cat_ax = chart.category_axis._element
        for tag in ("c:tickLblSkip", "c:tickMarkSkip"):
            for old_el in cat_ax.findall(qn(tag)):
                cat_ax.remove(old_el)
        # both elements are last in the CT_CatAx sequence, after c:noMultiLvlLbl
        cat_ax.append(parse_xml(f'<c:tickLblSkip {nsdecls("c")} val="1"/>'))
        cat_ax.append(parse_xml(f'<c:tickMarkSkip {nsdecls("c")} val="1"/>'))
    except Exception:
        pass

_STACKED_100_TYPES = (XL_CHART_TYPE.COLUMN_STACKED_100,
                      XL_CHART_TYPE.BAR_STACKED_100)
_STACKED_TYPES = (XL_CHART_TYPE.COLUMN_STACKED,
                  XL_CHART_TYPE.BAR_STACKED) + _STACKED_100_TYPES

def _add_cat_chart(slide, rect, chart_type, categories, series,
                   point_colors=None, legend=None, label_ends=False,
                   stacked_threshold=False, line_multi=False,
                   pct=False, pct_decimals=1, val_range=None,
                   force_labels=False):
    """series: list of (name, values, RGBColor|None).

    ``pct`` marks the values as percentages: the per-point labels then carry
    the sign ("7.0%", not "7.0") and, where a value axis is on show, its ticks
    carry it too - bare axis numbers beside per-cent labels read as a second,
    different unit.  ``val_range`` pins the axis to the range the dashboard
    figure asked for, so the deck keeps its zero baseline."""
    x, y, w, h = rect
    cd = CategoryChartData()
    cd.categories = [str(c) for c in categories]
    for name, values, _ in series:
        cd.add_series(str(name), [_num(v) for v in values])
    gf = slide.shapes.add_chart(chart_type, E(x), E(y), E(w), E(h), cd)
    chart = gf.chart
    line_types = (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS)
    for i, (name, values, color) in enumerate(series):
        s = chart.series[i]
        col = color or _palette(i)
        try:
            if chart_type in line_types:
                s.format.line.color.rgb = col
                s.format.line.width = Pt(1.75)
                s.smooth = False
            else:
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = col
                s.format.line.fill.background()
        except Exception:
            pass
    if point_colors:
        try:
            s = chart.series[0]
            for j, col in enumerate(point_colors):
                if col is None:
                    s.points[j].format.fill.background()
                else:
                    s.points[j].format.fill.solid()
                    s.points[j].format.fill.fore_color.rgb = col
        except Exception:
            pass
    _style_chart(chart, len(series), legend=legend, height_px=h)
    all_vals = [v for _, vs, _ in series for v in vs]
    fmt = (_pct_format(all_vals, pct_decimals) if pct
           else _num_format(all_vals))
    # Every chart keeps EVERY data label the deck has always drawn.  A locked
    # deck withholds exactly those labels - pptlock padlocks / blurs the
    # labels a chart already shows and never adds one - so a label dropped
    # here is a figure the locked deck silently stops locking.  Only their
    # size (the 11pt standard) and, on horizontal bars, their orientation
    # changed on 2026-09-11.
    labelled = True
    horizontal = chart_type in (XL_CHART_TYPE.BAR_CLUSTERED,
                                XL_CHART_TYPE.BAR_STACKED,
                                XL_CHART_TYPE.BAR_STACKED_100)
    if stacked_threshold:
        # only label segments tall enough to hold a number (thin bands stay
        # blank here and are listed beside the chart instead)
        _label_big_stacked_segments(chart, series)
    elif label_ends:
        # crowded multi-line chart -> value only at each line's end
        _label_series_ends(chart, series)
    elif line_multi:
        # a few lines -> label every point, coloured per line
        _label_all_line_points(chart, series, fmt)
    else:
        # numeric data labels on every point (short K/M/B so they don't overlap)
        stacked = chart_type in _STACKED_TYPES
        n_labels = len(series) * max(1, len(categories))
        size = MIN_CONTENT_PT
        if stacked:
            # labels sit INSIDE the bars -> colour them light/dark per segment
            # so numbers on dark bars are white and stay readable
            _stacked_center_labels(chart, series, fmt, size)
        elif (chart_type not in line_types and not horizontal
              and w * 0.85 / max(1, n_labels) < _MIN_PX_PER_BAR_LABEL):
            # CLUSTERED columns too narrow to hold a HORIZONTAL number: laid
            # flat they collide into unreadable mush (11 years x 3 scenarios
            # once rendered as "143.0K43.0K43.0K").  Standing them on end costs
            # only the line height in width, so every bar keeps its value and
            # nothing overlaps.  figures.trend_bars angles the dashboard's
            # labels the same way.  (Horizontal bars keep their numbers flat
            # at the bar end - a number on end beside a horizontal bar read
            # sideways.)
            _add_data_labels(chart, fmt, XL_DATA_LABEL_POSITION.OUTSIDE_END,
                             size=size, rotation=-90)
        else:
            pos = (XL_DATA_LABEL_POSITION.ABOVE if chart_type in line_types
                   else XL_DATA_LABEL_POSITION.OUTSIDE_END)
            _add_data_labels(chart, fmt, pos, size=size)
    _LAST_LABELLED["v"] = labelled
    # percentages keep their sign on the axis too; everything else just gets
    # its big numbers abbreviated (30000 -> "30K")
    axfmt = _pct_format(all_vals, 0) if pct else _axis_num_format(all_vals)
    if chart_type in _STACKED_100_TYPES:
        # PowerPoint scales this axis itself, in fractions - a format derived
        # from 0-100 series values would print every tick as "0%"
        axfmt = None
    if axfmt:
        try:
            chart.value_axis.tick_labels.number_format = axfmt
            chart.value_axis.tick_labels.number_format_is_linked = False
        except Exception:
            pass
    if val_range:
        try:
            chart.value_axis.minimum_scale = val_range[0]
            chart.value_axis.maximum_scale = val_range[1]
        except Exception:
            pass
    return chart

# ── plotly figure -> native chart(s) ─────────────────────────────────
def _fig_to_native(slide, spec):
    """Draw one figure as native shapes.

    Wrapped in _fig_to_native_inner so that a panel with empty or degenerate
    data degrades to a caption instead of aborting the whole export: python-pptx
    raises ValueError("chart data contains no categories") for an empty series,
    and one such panel used to cost the entire deck.
    """
    rect = _panel_frame(slide, spec)
    # Recorded here, not at layout time, so the List of Figures describes the
    # charts the deck REALLY carries: a page the workbook cannot fill is
    # never drawn, and takes its figures out of the list with it.
    toc.capture("figure", spec.get("title"))
    try:
        return _fig_to_native_inner(slide, spec, rect)
    except ValueError as exc:
        if "categor" not in str(exc).lower() and "series" not in str(exc).lower():
            raise
        return _empty_panel_note(slide, rect)


def _fig_to_native_inner(slide, spec, rect):
    fig = spec["figure"]
    traces = list(fig.data)
    kinds = {t.type for t in traces}

    if "indicator" in kinds:
        return _render_gauge(slide, rect, traces[0], fig)
    if "waterfall" in kinds:
        return _render_waterfall(slide, rect, traces[0])
    if "pie" in kinds:
        _cmeta = getattr(fig, "_chart_meta", None)
        _center = _cmeta.get("center", "") if isinstance(_cmeta, dict) else ""
        return _render_pie(slide, rect, traces[0], center=_center)
    if "choropleth" in kinds or "scattergeo" in kinds:
        return _render_choropleth(slide, rect, traces[0], spec)
    meta = getattr(fig, "_chart_meta", None)
    if isinstance(meta, dict) and meta.get("chart") == "scenario_bars":
        return _render_scenario_bars(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "scenario_lines":
        return _render_scenario_lines(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "mekko":
        return _render_mekko(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "composition":
        return _render_composition(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "dumbbell":
        return _render_dumbbell(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "comp_bps":
        return _render_comp_bps(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "pct_stack":
        return _render_pct_stack(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "boxplot":
        return _render_boxplot(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "candlestick":
        return _render_candlestick(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "price_strip":
        return _render_price_strip(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "diverging":
        return _render_diverging(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "radar":
        return _render_radar(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "tier_funnel":
        return _render_tier_funnel(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "funnel3d":
        return _render_funnel3d(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "samfunnel":
        return _render_sam_funnel(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "posbubble":
        return _render_pos_bubble(slide, rect, meta)
    if isinstance(meta, dict) and meta.get("chart") == "target_actual":
        return _render_target_actual(slide, rect, meta)
    if "treemap" in kinds or "sunburst" in kinds:
        return _render_treemap(slide, rect, traces[0])
    if "heatmap" in kinds:
        return _render_heatmap(slide, rect, traces[0])
    bars = [t for t in traces if t.type == "bar"]
    lines = [t for t in traces if t.type == "scatter"
             and "lines" in (t.mode or "lines")]
    bubbles = [t for t in traces if t.type == "scatter"
               and "markers" in (t.mode or "") and t.mode
               and "lines" not in t.mode]
    if bars and lines:                       # PBIX combo -> two charts
        bh = rect[3] * 0.56
        _render_bars(slide, (rect[0], rect[1], rect[2], bh), bars,
                     fig, legend=True)
        pct = False
        try:
            pct = (fig.layout.yaxis2.tickformat or "").endswith("%")
        except Exception:
            pass
        _render_lines(slide, (rect[0], rect[1] + bh + 4, rect[2],
                              rect[3] - bh - 4), lines, pct=pct)
        return
    if bars and getattr(fig.layout, "barmode", None) == "overlay" \
            and len(bars) == 2:                  # 'bar inside a bar'
        return _render_nested_bars(slide, rect, bars)
    if bars:
        # a Historical/Forecast period bracket reserves a strip at the top
        pmeta = getattr(fig, "_chart_meta", None)
        pband = pmeta.get("period_band") if isinstance(pmeta, dict) else None
        chart_rect = rect
        if pband:
            bx0, by0, bw0, bh0 = rect
            chart_rect = (bx0, by0 + 30, bw0, bh0 - 30)
        # two absolute metrics on separate plotly axes (e.g. Volume + Value)
        # -> one clustered double-bar chart; hide the value-axis numbers since
        # the two series use different units (the per-bar labels carry values).
        dual = (len(bars) == 2 and
                getattr(bars[1], "yaxis", "y") == "y2")
        _render_bars(slide, chart_rect, bars, fig,
                     legend=True if dual else None, hide_val_axis=dual)
        _annotation_note(slide, spec, fig)
        if pband:
            _render_period_band(slide, rect, chart_rect, pband)
        return
    if bubbles:
        return _render_bubbles(slide, rect, bubbles)
    if lines:
        return _render_lines(slide, rect, lines, fig=fig)

def _dlbls_rotation(dls, degrees=0):
    """Set the rotation of a set of data labels.

    ``degrees`` 0 keeps numbers horizontal (the default everywhere); -90 stands
    them on end, reading bottom-to-top, which is how a clustered chart with
    many narrow bars can still carry a value on every bar without the numbers
    colliding.  OOXML measures rotation in 60000ths of a degree.
    """
    try:
        txPr = dls._element.find(qn("c:txPr"))
        if txPr is None:
            return
        bodyPr = txPr.find(qn("a:bodyPr"))
        if bodyPr is None:
            bodyPr = parse_xml(f'<a:bodyPr {nsdecls("a")}/>')
            txPr.insert(0, bodyPr)
        bodyPr.set("rot", str(int(round(degrees * 60000))))
        bodyPr.set("vert", "horz")
    except Exception:
        pass

def _dlbls_horizontal(dls):
    """Keep data labels horizontal (never rotated) so numbers read cleanly."""
    _dlbls_rotation(dls, 0)

def _pin_plot_area(chart, px, py, pw, ph):
    """Pin a chart's INNER plot area to a fixed fraction of the chart frame.

    PowerPoint otherwise sizes the plot area itself, and how much it reserves
    for a legend varies with the legend's content - so anything positioned by
    guessing where the plot ended up (e.g. a doughnut's centre caption) drifts.
    Pinning it makes the geometry knowable, so the caller can place things at
    the exact centre.  Returns the (x, y, w, h) fractions actually set.
    """
    try:
        plot_area = chart._chartSpace.find(qn("c:chart")).find(qn("c:plotArea"))
        for existing in plot_area.findall(qn("c:layout")):
            plot_area.remove(existing)
        layout = parse_xml(
            f'<c:layout {nsdecls("c")}>'
            f'<c:manualLayout>'
            f'<c:layoutTarget val="inner"/>'
            f'<c:xMode val="edge"/><c:yMode val="edge"/>'
            f'<c:x val="{px}"/><c:y val="{py}"/>'
            f'<c:w val="{pw}"/><c:h val="{ph}"/>'
            f'</c:manualLayout></c:layout>')
        plot_area.insert(0, layout)          # c:layout must precede the series
        return px, py, pw, ph
    except Exception:
        return None

def _series_labels(series, vals, pos, color, size, fmt=None):
    fmt = fmt or _num_format([v for v in vals if v is not None])
    try:
        dls = series.data_labels
        dls.show_value = True
        dls.number_format = fmt
        dls.number_format_is_linked = False
        dls.font.size = Pt(size)
        dls.font.name = FONT
        dls.font.color.rgb = color
        dls.position = pos
        _dlbls_horizontal(dls)
    except Exception:
        pass

def _set_gap(bar_el, val):
    gw = bar_el.find(qn("c:gapWidth"))
    if gw is None:
        gw = parse_xml(f'<c:gapWidth {nsdecls("c")} val="{val}"/>')
        axid = bar_el.find(qn("c:axId"))
        bar_el.insert(list(bar_el).index(axid), gw)
    else:
        gw.set("val", str(val))

def _set_axids(bar_el, cat_id, val_id):
    for a in bar_el.findall(qn("c:axId")):
        bar_el.remove(a)
    bar_el.append(parse_xml(f'<c:axId {nsdecls("c")} val="{cat_id}"/>'))
    bar_el.append(parse_xml(f'<c:axId {nsdecls("c")} val="{val_id}"/>'))

def _add_ser(bar_el, ser):
    """Insert a <c:ser> in schema order (before dLbls/gapWidth/axId), else the
    bar loses its fill in PowerPoint."""
    for tag in ("c:dLbls", "c:gapWidth", "c:overlap", "c:axId"):
        ref = bar_el.find(qn(tag))
        if ref is not None:
            bar_el.insert(list(bar_el).index(ref), ser)
            return
    bar_el.append(ser)

def _nmx(vals):
    m = [abs(v) for v in vals if v is not None]
    return max(m) if m else 1.0

def _secondary_axes(cat_id, val_id, num_fmt, side="r"):
    fmt = (num_fmt or "General").replace('"', "&quot;")
    axpos = "l" if side == "l" else "r"
    crosses = "autoZero" if side == "l" else "max"
    cat = parse_xml(
        f'<c:catAx {nsdecls("c")}><c:axId val="{cat_id}"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        f'<c:delete val="1"/><c:axPos val="b"/><c:crossAx val="{val_id}"/>'
        f'<c:crosses val="{crosses}"/></c:catAx>')
    val = parse_xml(
        f'<c:valAx {nsdecls("c")}><c:axId val="{val_id}"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        f'<c:delete val="0"/><c:axPos val="{axpos}"/>'
        f'<c:numFmt formatCode="{fmt}" sourceLinked="0"/>'
        '<c:majorTickMark val="out"/><c:minorTickMark val="none"/>'
        '<c:tickLblPos val="nextTo"/>'
        f'<c:crossAx val="{cat_id}"/><c:crosses val="max"/></c:valAx>')
    return cat, val

def _axis_to_side(plot_area, cat_id, val_id, side):
    """Put a value axis (and where its category axis crosses it) on left/right."""
    crosses = "max" if side == "r" else "autoZero"
    for ax in plot_area.findall(qn("c:catAx")):
        if ax.find(qn("c:axId")).get("val") == cat_id:
            cr = ax.find(qn("c:crosses"))
            if cr is None:
                ax.append(parse_xml(f'<c:crosses {nsdecls("c")} '
                                    f'val="{crosses}"/>'))
            else:
                cr.set("val", crosses)
    for ax in plot_area.findall(qn("c:valAx")):
        if ax.find(qn("c:axId")).get("val") == val_id:
            ap = ax.find(qn("c:axPos"))
            if ap is not None:
                ap.set("val", side)

def _label_points_real(series, real_vals, pos, color, size):
    """Per-point labels showing REAL numbers (used when a bar is plotted at a
    scaled height but must still display its true value)."""
    for j, v in enumerate(real_vals):
        if v is None:
            continue
        try:
            dl = series.points[j].data_label
            tf = dl.text_frame
            tf.text = _abbr(v)
            r = tf.paragraphs[0].runs[0]
            r.font.size = Pt(size)
            r.font.name = FONT
            r.font.color.rgb = color
            dl.position = pos
        except Exception:
            pass

def _render_nested_bars(slide, rect, bars):
    """'Bar inside a bar': a WIDE bar with a NARROWER one nested inside it -
    two overlaid column plots of different widths on ONE value axis, the wide
    one drawn first (behind) and the narrow one second (IN FRONT).

    Both series carry their REAL numbers.  The chart used to rescale the inner
    series so it always fitted inside the outer one, which made the picture
    right and the underlying data wrong - "Edit Data" read back figures that
    were nowhere in the report.  Instead, when the two series use different
    units (Volume vs Value) the LARGER one takes the wide bar, so the other
    nests inside it as it stands, and the axis numbers are dropped: one axis
    cannot label two units, and every bar carries its own value.  The wide bar
    keeps the pale 'background' colour and the narrow bar the saturated one,
    whichever series each turns out to be - the legend names them.

    With two units (Volume vs Value) the narrow bar reads on its OWN hidden
    secondary axis, scaled so that every inner bar stands just inside its
    outer bar - the reference 'bar inside a bar' look, whatever the two
    units' magnitudes - while both series still carry their real numbers."""
    from copy import deepcopy
    x, y, w, h = rect
    wide, narrow = sorted(bars, key=lambda t: -(t.width or 0))   # wide, narrow
    pale, deep = _trace_color(wide, 0), _trace_color(narrow, 1)  # by ROLE
    dual = any(getattr(b, "yaxis", "y") == "y2" for b in bars)
    cats = [str(c) for c in wide.x]
    w_vals = [_num(v) for v in wide.y]
    n_vals = [_num(v) for v in narrow.y]
    names = (str(wide.name or "Outer"), str(narrow.name or "Inner"))

    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series(names[0], w_vals)                              # ser0 = wide
    cd.add_series(names[1], n_vals)                              # ser1 = narrow
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                                E(x), E(y), E(w), E(h), cd)
    chart = gf.chart
    _style_chart(chart, 2, legend=True)
    for s, col in ((chart.series[0], pale), (chart.series[1], deep)):
        try:
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = col
            s.format.line.fill.background()
        except Exception:
            pass
    # with two units the axis carries no numbers, so the wide bar shows its
    # full value rather than the axis's abbreviated one
    o_fmt = (_num_format(w_vals) if dual
             else (_axis_num_format(w_vals) or _num_format(w_vals)))
    # both series keep their labels - a locked deck padlocks exactly these
    _series_labels(chart.series[0], w_vals,
                   XL_DATA_LABEL_POSITION.OUTSIDE_END, _C_MUTED,
                   MIN_CONTENT_PT, fmt=o_fmt)
    _series_labels(chart.series[1], n_vals,      # data-linked, so the label and
                   XL_DATA_LABEL_POSITION.CENTER, _C_WHITE, MIN_CONTENT_PT,
                   fmt=_num_format(n_vals))      # the series can never diverge

    # two overlaid barCharts of different widths, BOTH on the primary axis:
    # wide drawn first (behind), narrow second (in front) -> narrow is in FRONT
    bar_el = chart.plots[0]._element
    plot_area = bar_el.getparent()
    sers = bar_el.findall(qn("c:ser"))
    wide_ser, narrow_ser = sers[0], sers[1]
    bar_el.remove(wide_ser)                         # bar_el keeps the narrow one
    # the OUTER bar carries the deck-wide 50% gap; the inner one has to stay
    # narrower than it or the nesting disappears, so it keeps the same
    # inner/outer width ratio (~0.47) against the new outer width
    _set_gap(bar_el, 219)                            # narrow (nests inside)
    wide_bar = deepcopy(bar_el)
    for s in wide_bar.findall(qn("c:ser")):
        wide_bar.remove(s)
    _add_ser(wide_bar, deepcopy(wide_ser))
    _set_gap(wide_bar, BAR_GAP_WIDTH)                # wide = deck-wide gap
    plot_area.insert(list(plot_area).index(bar_el), wide_bar)

    # the primary value axis is scaled to the wide series: min 0, max = its
    # own maximum x 1.18, abbreviated
    w_max = _nmx(w_vals) * 1.18
    for vax in plot_area.findall(qn("c:valAx")):
        _scale_val_axis(vax, w_max, _axis_num_format(w_vals))
    if dual:
        # two units - the tick numbers would label only one of them, so the
        # per-bar values carry the reading instead.  (Done BEFORE the second
        # axis exists: with two, python-pptx's chart.value_axis is the 2nd.)
        _hide_value_axis_labels(chart)
        # the narrow series moves to its own (hidden) secondary axis whose
        # maximum is chosen so the tallest inner/outer ratio fills 88% of its
        # outer bar - every other inner bar then sits inside its outer one
        ratios = [nv / wv for nv, wv in zip(n_vals, w_vals)
                  if nv is not None and wv and nv > 0 and wv > 0]
        n_max = (max(ratios) * w_max / 0.88) if ratios else _nmx(n_vals) * 1.18
        cat_id, val_id = "90210001", "90210002"
        _set_axids(bar_el, cat_id, val_id)
        cat_ax, val_ax = _secondary_axes(cat_id, val_id,
                                         _axis_num_format(n_vals))
        last_ax = plot_area.findall(qn("c:valAx"))[-1]
        at = list(plot_area).index(last_ax) + 1
        plot_area.insert(at, val_ax)
        plot_area.insert(at, cat_ax)
        _scale_val_axis(val_ax, n_max, None)
        # hidden, NOT deleted: PowerPoint plots a series whose secondary axis
        # is deleted against the primary scale, which flattens it to nothing
        val_ax.find(qn("c:tickLblPos")).set("val", "none")
        val_ax.find(qn("c:majorTickMark")).set("val", "none")
        val_ax.insert(list(val_ax).index(val_ax.find(qn("c:tickLblPos"))) + 1,
                      parse_xml(f'<c:spPr {nsdecls("c", "a")}><a:ln>'
                                '<a:noFill/></a:ln></c:spPr>'))
        # the bars grow up from zero, not down from the axis maximum
        cat_ax.find(qn("c:crosses")).set("val", "autoZero")
    return chart

def _scale_val_axis(vax, vmax, num_fmt):
    """Pin a value axis to [0, vmax] and (optionally) give it an abbreviated
    number format, keeping <c:scaling> children in schema order."""
    scaling = vax.find(qn("c:scaling"))
    for tag in ("c:max", "c:min"):
        e = scaling.find(qn(tag))
        if e is not None:
            scaling.remove(e)
    orient = scaling.find(qn("c:orientation"))
    i0 = (list(scaling).index(orient) + 1) if orient is not None else 0
    scaling.insert(i0, parse_xml(f'<c:min {nsdecls("c")} val="0"/>'))
    vtxt = f"{vmax:.0f}" if abs(vmax) >= 10 else f"{vmax:.6g}"
    scaling.insert(i0, parse_xml(f'<c:max {nsdecls("c")} val="{vtxt}"/>'))
    if not num_fmt:
        return
    old = vax.find(qn("c:numFmt"))
    if old is not None:
        vax.remove(old)
    fmt = num_fmt.replace('"', "&quot;")
    nf = parse_xml(f'<c:numFmt {nsdecls("c")} formatCode="{fmt}" '
                   'sourceLinked="0"/>')
    anchor = vax.find(qn("c:majorTickMark"))
    if anchor is None:
        anchor = vax.find(qn("c:crossAx"))
    vax.insert(list(vax).index(anchor), nf)

def _lerp_rgb(c1, c2, f):
    a = _rgb(c1) if isinstance(c1, str) else c1
    b = _rgb(c2) if isinstance(c2, str) else c2
    return RGBColor(int(round(a[0] + (b[0] - a[0]) * f)),
                    int(round(a[1] + (b[1] - a[1]) * f)),
                    int(round(a[2] + (b[2] - a[2]) * f)))

# Single-hue sequential meter gradient - mirrors figures.HEAT_SCALE
# (white = lowest → FMI deep = highest).
_HEAT_STOPS = [tuple(stop) for stop in HEAT_STOPS]

def _heat_color(t):
    """Multi-stop sequential meter (t in [0,1]); matches figures.HEAT_SCALE -
    white = lowest value, deep teal = highest."""
    t = max(0.0, min(1.0, t))
    for i in range(1, len(_HEAT_STOPS)):
        p0, c0 = _HEAT_STOPS[i - 1]
        p1, c1 = _HEAT_STOPS[i]
        if t <= p1:
            f = (t - p0) / (p1 - p0) if p1 > p0 else 0.0
            return _lerp_rgb(c0, c1, f)
    return _rgb(_HEAT_STOPS[-1][1])

def _render_heatmap(slide, rect, t):
    """Native colour-coded intensity matrix - a white→deep 'meter' (no numbers
    on the tiles) with a colour-bar legend on the right mapping shade→value.
    A faint tile border keeps the near-white (lowest) cells visible."""
    x, y, w, h = rect
    xs = [str(v) for v in t.x]
    ys = [str(v) for v in t.y]
    z = [[_num(v) for v in row] for row in t.z]
    nr, nc = len(ys), len(xs)
    flat = [v for row in z for v in row if v is not None]
    if not flat or nr == 0 or nc == 0:
        return
    zlo, zhi = min(flat), max(flat)
    zrange = max(zhi - zlo, 1e-9)
    grid = _C_GRID          # was a hex literal: a renderer must not
                            # pick a colour, the theme does
    suffix = "%"
    cbar_w = 46.0                                  # reserve right strip for legend
    lab_w = min(170.0, max(70.0, (w - cbar_w) * 0.22))
    top_h = 15.0
    gx, gy = x + lab_w, y + top_h
    grid_w = w - lab_w - cbar_w
    cw, ch = grid_w / nc, (h - top_h) / nr
    for j, xv in enumerate(xs):                    # column headers on top
        tb = _textbox(slide, gx + j * cw, y, cw, top_h)
        _set_text(tb.text_frame, [[(xv, 7.5, True, _C_MUTED)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    for i in range(nr):
        tb = _textbox(slide, x, gy + i * ch, lab_w - 3, ch)
        _set_text(tb.text_frame, [[(ys[i], 7.5, False, _C_TEXT)]],
                  anchor=MSO_ANCHOR.MIDDLE)
        for j in range(nc):
            v = z[i][j]
            if v is None:
                continue
            col = _heat_color((v - zlo) / zrange)
            shp = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, E(gx + j * cw), E(gy + i * ch),
                E(cw), E(ch))
            shp.fill.solid()
            shp.fill.fore_color.rgb = col
            shp.line.color.rgb = grid
            shp.line.width = Pt(0.5)
            shp.shadow.inherit = False
    # colour-bar legend (the 'meter'): vertical gradient + low/mid/high labels
    bar_x = x + w - cbar_w + 6
    bar_w = 11.0
    bar_top, bar_bot = gy + 4, gy + (h - top_h) - 4
    steps = 40
    seg = (bar_bot - bar_top) / steps
    for k in range(steps):
        frac_top = 1.0 - k / steps                 # top = highest value = deep
        cell = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, E(bar_x), E(bar_top + k * seg),
            E(bar_w), E(seg + 0.8))
        cell.fill.solid()
        cell.fill.fore_color.rgb = _heat_color(frac_top)
        cell.line.fill.background()
        cell.shadow.inherit = False
    frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(bar_x), E(bar_top),
                                   E(bar_w), E(bar_bot - bar_top))
    frame.fill.background()
    frame.line.color.rgb = grid
    frame.line.width = Pt(0.5)
    frame.shadow.inherit = False
    for frac, val in ((1.0, zhi), (0.5, (zlo + zhi) / 2), (0.0, zlo)):
        ly = bar_bot - frac * (bar_bot - bar_top)
        tl = _textbox(slide, bar_x + bar_w + 1, ly - 7, cbar_w - bar_w - 6, 14)
        _set_text(tl.text_frame,
                  [[(f"{val:.0f}{suffix}", 7, False, _C_MUTED)]],
                  align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

def _render_bars(slide, rect, bars, fig, legend=None, hide_val_axis=False):
    horizontal = getattr(bars[0], "orientation", None) == "h"
    stacked = getattr(fig.layout, "barmode", None) == "stack"
    # what the dashboard says about the value axis: is it a percentage, has it
    # been pinned to zero, is it hidden altogether?
    pct = _pct_axis(fig, horizontal)
    val_range = _val_axis_range(fig, horizontal)
    vax = getattr(fig.layout, "xaxis" if horizontal else "yaxis", None)
    if getattr(vax, "visible", None) is False:
        hide_val_axis = True
    if horizontal:
        chart_type = (XL_CHART_TYPE.BAR_STACKED if stacked
                      else XL_CHART_TYPE.BAR_CLUSTERED)
        categories = list(bars[0].y)
    else:
        chart_type = (XL_CHART_TYPE.COLUMN_STACKED if stacked
                      else XL_CHART_TYPE.COLUMN_CLUSTERED)
        categories = list(bars[0].x)
    series, point_colors = [], None
    for i, t in enumerate(bars):
        vals = list(t.x) if horizontal else list(t.y)
        mk_col = getattr(t.marker, "color", None)
        if len(bars) == 1 and isinstance(mk_col, (list, tuple)):
            point_colors = [_rgb(c) if isinstance(c, str)
                            and c.startswith("#") else _palette(0)
                            for c in mk_col]
            series.append((t.name or "Series", vals, _palette(0)))
        else:
            series.append((t.name or f"S{i + 1}", vals,
                           _trace_color(t, i)))
    # dense vertical stack (many thin segments) -> label only the segments big
    # enough to hold a number, and move every series' value to a colour-keyed
    # list beside the chart so the thin ones never overlap
    dense = (stacked and not horizontal and len(series) >= 4)
    if dense:
        x, y, w, h = rect
        list_w = min(190, max(150, int(w * 0.16)))
        chart = _add_cat_chart(slide, (x, y, w - list_w - 8, h), chart_type,
                               categories, series, legend=False,
                               stacked_threshold=True, pct=pct,
                               val_range=val_range)
        _stacked_value_list(slide, (x + w - list_w, y, list_w, h),
                            categories, series)
        return
    chart = _add_cat_chart(slide, rect, chart_type, categories, series,
                           point_colors=point_colors, legend=legend,
                           pct=pct, val_range=val_range,
                           force_labels=hide_val_axis)
    # the axis goes only when the bars carry their numbers instead
    if hide_val_axis and _LAST_LABELLED["v"]:
        _hide_value_axis_labels(chart)

# ── Side legend beside a stacked chart ───────────────────────────────
# Type scales the legend is allowed to shrink through: (value, chip, name,
# space-before), all in points, largest first.  The list used to be drawn at a
# fixed size, so a page with one entry per COUNTRY (a region-scoped deck: 17
# of them) simply ran off the bottom of its panel.  It now picks the largest
# scale that fits, and only then breaks into two columns - so a short legend
# still looks exactly as it always did.
_LEG_STEPS = ((9.5, 9.0, 7.5, 5.0), (9.0, 8.5, 7.2, 4.0),
              (8.5, 8.0, 7.0, 3.0), (8.0, 7.5, 6.8, 2.2),
              (7.5, 7.0, 6.4, 1.6), (7.0, 6.5, 6.0, 1.2),
              (6.5, 6.0, 5.6, 0.9), (6.0, 5.5, 5.2, 0.6))

def _leg_row_h(step) -> float:
    """A legend row's drawn height in slide px (pt -> px, then line spacing)."""
    return max(step[0], step[1], step[2]) * 1.333 * 1.24 + step[3] * 1.333

def _leg_fit(n_rows: int, avail_h: float) -> tuple:
    """(type scale, columns) for ``n_rows`` entries in ``avail_h`` px.

    Shrinking is preferred to splitting: a single column keeps the full panel
    width for the names, which is what stops "Rest of Europe" wrapping."""
    for step in _LEG_STEPS:
        if _leg_row_h(step) * n_rows <= avail_h:
            return step, 1
    for step in _LEG_STEPS:
        if _leg_row_h(step) * -(-n_rows // 2) <= avail_h:
            return step, 2
    return _LEG_STEPS[-1], 2

def _stacked_value_list(slide, rect, categories, series):
    """A tidy colour-keyed panel beside a stacked chart: every segment (incl. the
    thin ones that can't be labelled inside) shows its latest value, marked with
    the series colour so the number ties back to its band."""
    x, y, w, h = rect
    panel = _white_panel(slide, x, y + 4, w, h - 8)
    panel.line.color.rgb = _C_EDGE
    latest = str(categories[-1]) if categories else ""
    rows = sorted(series, key=lambda s: (_num(s[1][-1]) or 0), reverse=True)

    # header first, then whatever height is left belongs to the entries
    head_h = 18.0
    htb = _textbox(slide, x + 8, y + 10, w - 14, head_h)
    hr = htb.text_frame.paragraphs[0].add_run()
    hr.text = f"By segment  ·  {latest}"
    hr.font.size = Pt(8.5)
    hr.font.bold = True
    hr.font.name = FONT
    hr.font.color.rgb = _C_MUTED

    body_y = y + 10 + head_h
    avail = (y + h - 10) - body_y
    (val_fs, chip_fs, nm_fs, sp), ncols = _leg_fit(len(rows), avail)
    per = -(-len(rows) // ncols)
    col_w = (w - 14 - (ncols - 1) * 6) / ncols

    for c in range(ncols):
        part = rows[c * per:(c + 1) * per]
        if not part:
            continue
        tb = _textbox(slide, x + 8 + c * (col_w + 6), body_y, col_w, avail)
        tf = tb.text_frame
        tf.word_wrap = True
        for i, (name, vals, col) in enumerate(part):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_before = Pt(sp)
            chip = p.add_run()
            chip.text = "■  "
            chip.font.size = Pt(chip_fs)
            chip.font.name = FONT
            chip.font.color.rgb = col or _palette(0)
            val = p.add_run()
            val.text = f"{_abbr(_num(vals[-1]) or 0)}"
            val.font.size = Pt(val_fs)
            val.font.bold = True
            val.font.name = FONT
            val.font.color.rgb = col or _palette(0)
            nm = p.add_run()
            nm.text = f"   {name}"
            nm.font.size = Pt(nm_fs)
            nm.font.name = FONT
            nm.font.color.rgb = _C_TEXT

def _hide_value_axis_labels(chart):
    """Blank the value-axis tick numbers, and the tick marks with them (used
    when two different-unit series share one axis, or when the per-point
    labels carry the reading).  Leaving the marks behind put a short ruled
    stub under the plot with nothing to measure."""
    try:
        val_ax = chart.value_axis._element
        tlp = val_ax.find(qn("c:tickLblPos"))
        if tlp is None:
            tlp = val_ax.makeelement(qn("c:tickLblPos"), {})
            val_ax.append(tlp)
        tlp.set("val", "none")
        for tag in ("c:majorTickMark", "c:minorTickMark"):
            tm = val_ax.find(qn(tag))
            if tm is not None:
                tm.set("val", "none")
    except Exception:
        pass
    try:
        chart.value_axis.format.line.fill.background()
    except Exception:
        pass

def _render_lines(slide, rect, lines, pct=False, fig=None):
    categories = list(lines[0].x)
    val_range = None
    if fig is not None:
        pct = pct or _pct_axis(fig)
        val_range = _val_axis_range(fig)
    series = [(t.name or f"S{i + 1}", list(t.y), _trace_color(t, i))
              for i, t in enumerate(lines)]
    # 1 line: every point (default).  2-3 lines: every point, coloured &
    # staggered.  4+ lines: only each line's end (avoids the mid-chart pile-up).
    # A 2-column chart is a SLOPE graph -> label both ends whatever the count.
    n = len(series)
    slope = len(categories) == 2
    chart = _add_cat_chart(slide, rect, XL_CHART_TYPE.LINE_MARKERS,
                           categories, series,
                           label_ends=(n >= 4 and not slope),
                           line_multi=(slope or 2 <= n <= 3),
                           pct=pct, val_range=val_range)

def _render_pie(slide, rect, t, center=""):
    labels = list(t.labels) if t.labels is not None else \
        [f"S{i}" for i in range(len(t.values))]
    cd = CategoryChartData()
    cd.categories = [str(x) for x in labels]
    cd.add_series("Share", [_num(v) for v in t.values])
    x, y, w, h = rect
    gf = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT,
                                E(x), E(y), E(w), E(h), cd)
    chart = gf.chart
    mk_cols = getattr(t.marker, "colors", None)
    slice_cols = []
    try:
        s = chart.series[0]
        for j in range(len(labels)):
            col = None
            if mk_cols is not None and j < len(mk_cols) and \
                    isinstance(mk_cols[j], str) and \
                    mk_cols[j].startswith("#"):
                col = _rgb(mk_cols[j])
            col = col or _palette(j)
            slice_cols.append(col)
            s.points[j].format.fill.solid()
            s.points[j].format.fill.fore_color.rgb = col
    except Exception:
        pass
    _style_chart(chart, 2, legend=True)
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    # Pin the plot area so the doughnut's centre is a known point rather than
    # wherever PowerPoint happened to leave it after sizing the legend - the
    # centre caption below is placed from these same numbers, so it can no
    # longer drift sideways.  The right ~30% is left for the legend.
    _PIE_PLOT = (0.01, 0.04, 0.68, 0.92)
    pinned = _pin_plot_area(chart, *_PIE_PLOT)

    # put the % share on every slice (was blank).  Doughnut charts don't
    # honour show_percentage reliably, so write explicit label text and
    # colour it per slice for readable contrast.
    try:
        vals = [abs(_num(v) or 0.0) for v in t.values]
        total = sum(vals) or 1.0
        plot = chart.plots[0]
        plot.has_data_labels = True
        for j, col in enumerate(slice_cols or [_palette(i)
                                for i in range(len(vals))]):
            try:
                dl = chart.series[0].points[j].data_label
                dl.text_frame.text = f"{vals[j] / total * 100:.1f}%"
                for para in dl.text_frame.paragraphs:
                    para.font.size = Pt(8.5)
                    para.font.bold = True
                    para.font.name = FONT
                    para.font.color.rgb = _readable_on(col)
            except Exception:
                pass
    except Exception:
        pass
    if center:
        # Centre caption goes at the exact middle of the pinned plot area (the
        # doughnut is drawn centred inside it), so it sits in the hole instead
        # of drifting to one side.
        # '<br>' (plotly's line break) becomes real lines here so a caption
        # like "100% =<br>143,010 US$ Th" renders the same on both surfaces.
        px, py, pw, ph = pinned or _PIE_PLOT
        lines = [ln for ln in str(center).split("<br>") if ln != ""]
        bh = max(24, 18 * len(lines) + 6)
        cxp = x + w * (px + pw / 2)
        cyp = y + h * (py + ph / 2)
        # keep the caption inside the doughnut hole: the ring's diameter is the
        # smaller of the plot area's sides, and the hole is roughly half of it
        bw = min(160.0, max(60.0, min(w * pw, h * ph) * 0.52))
        tb = _textbox(slide, cxp - bw / 2, cyp - bh / 2, bw, bh)
        _set_text(tb.text_frame, [[(ln, 13, True, _C_TEXT)] for ln in lines],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def _render_waterfall(slide, rect, t):
    """Native editable waterfall via the stacked-column technique."""
    labels = [str(v) for v in t.x]
    values = [float(v) for v in t.y]
    measures = list(t.measure)
    base, delta, colors = [], [], []
    cum = 0.0
    for v, m in zip(values, measures):
        if m in ("absolute", "total"):
            base.append(0.0)
            delta.append(v)
            colors.append(_C_ACCENT)
            cum = v
        else:
            nxt = cum + v
            base.append(min(cum, nxt))
            delta.append(abs(v))
            colors.append(_C_GOOD if v >= 0 else _C_BAD)
            cum = nxt
    x, y, w, h = rect
    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("base", base)
    cd.add_series("value", delta)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED,
                                E(x), E(y), E(w), E(h), cd)
    chart = gf.chart
    try:
        chart.series[0].format.fill.background()
        chart.series[0].format.line.fill.background()
        s = chart.series[1]
        for j, col in enumerate(colors):
            s.points[j].format.fill.solid()
            s.points[j].format.fill.fore_color.rgb = col
    except Exception:
        pass
    _style_chart(chart, 1, legend=False)
    _apply_bar_gap(chart, len(delta))              # deck-wide 50% gap
    # value labels on the visible (delta) series only, at 11pt - every step
    # keeps its label, which is what a locked deck padlocks
    try:
        dls = chart.series[1].data_labels
        dls.show_value = True
        dls.number_format = _num_format(delta)
        dls.number_format_is_linked = False
        dls.font.size = Pt(MIN_CONTENT_PT)
        dls.font.name = FONT
        dls.position = XL_DATA_LABEL_POSITION.INSIDE_END
    except Exception:
        pass
    # abbreviate big axis numbers (150000 -> "150K")
    axfmt = _axis_num_format([b + d for b, d in zip(base, delta)])
    if axfmt:
        try:
            chart.value_axis.tick_labels.number_format = axfmt
            chart.value_axis.tick_labels.number_format_is_linked = False
        except Exception:
            pass

def _render_bubbles(slide, rect, traces):
    x, y, w, h = rect
    cd = BubbleChartData()
    for i, t in enumerate(traces):
        s = cd.add_series(str(t.name or f"S{i + 1}"))
        sizes = getattr(t.marker, "size", None)
        sizes = list(sizes) if isinstance(sizes, (list, tuple)) \
            else [10] * len(t.x)
        for xi, yi, si in zip(t.x, t.y, sizes):
            s.add_data_point(_num(xi) or 0, _num(yi) or 0,
                             max(_num(si) or 1, 1))
    gf = slide.shapes.add_chart(XL_CHART_TYPE.BUBBLE,
                                E(x), E(y), E(w), E(h), cd)
    chart = gf.chart
    for i, t in enumerate(traces):
        try:
            chart.series[i].format.fill.solid()
            chart.series[i].format.fill.fore_color.rgb = \
                _trace_color(t, i)
        except Exception:
            pass
    _style_chart(chart, len(traces), legend=True)

def _render_choropleth(slide, rect, t, spec):
    """Maps have no native PPT equivalent -> editable top-N bar chart.
    Handles both choropleth (value in ``z``) and the scattergeo bubble
    map (value in ``marker.color``)."""
    names = list(t.text) if t.text is not None else []
    if t.type == "choropleth":
        vals = list(t.z) if t.z is not None else []
    else:                                        # scattergeo bubble map
        mc = getattr(t.marker, "color", None)
        vals = list(mc) if mc is not None and not isinstance(mc, str) else []
    seen: dict[str, float] = {}
    for name, z in zip(names, vals):
        try:
            seen.setdefault(str(name), float(z))
        except (TypeError, ValueError):
            continue
    items = sorted(seen.items(), key=lambda kv: kv[1], reverse=True)[:10]
    items.reverse()                       # largest on top of bar chart
    series = [("Index", [v for _, v in items], _C_ACCENT)]
    _add_cat_chart(slide, rect, XL_CHART_TYPE.BAR_CLUSTERED,
                   [k for k, _ in items], series, legend=False)

def _sq_worst(row, side):
    s = sum(row)
    thick = s / side
    return max(max(thick * thick / a, a / (thick * thick)) for a in row)

def _squarify(sizes, x, y, dx, dy):
    """Classic squarified-treemap layout: nested rectangles filling (dx,dy)."""
    sizes = [float(s) for s in sizes if s > 0]
    if not sizes or dx <= 0 or dy <= 0:
        return []
    total = sum(sizes)
    scale = (dx * dy) / total
    return _sq([s * scale for s in sizes], x, y, dx, dy)

def _sq(sizes, x, y, dx, dy):
    if not sizes:
        return []
    if len(sizes) == 1:
        return [{"x": x, "y": y, "dx": dx, "dy": dy}]
    side = dy if dx >= dy else dx
    i = 1
    while (i < len(sizes)
           and _sq_worst(sizes[:i], side) >= _sq_worst(sizes[:i + 1], side)):
        i += 1
    row, rest = sizes[:i], sizes[i:]
    thick = sum(row) / side
    rects, off = [], 0.0
    for a in row:
        seg = a / thick
        if dx >= dy:
            rects.append({"x": x, "y": y + off, "dx": thick, "dy": seg})
        else:
            rects.append({"x": x + off, "y": y, "dx": seg, "dy": thick})
        off += seg
    if dx >= dy:
        return rects + _sq(rest, x + thick, y, dx - thick, dy)
    return rects + _sq(rest, x, y + thick, dx, dy - thick)

def _render_treemap(slide, rect, t):
    """Real squarified treemap (nested rectangles) - also used for the sunburst
    hierarchy in the PPT.  Leaves sized by value, coloured, with in-tile labels."""
    x, y, w, h = rect
    labels = [str(l) for l in t.labels]
    parents = [str(p) for p in (t.parents if t.parents is not None
                                else [""] * len(labels))]
    vals = [_num(v) or 0.0 for v in t.values]
    cols = None
    try:
        if t.marker.colors is not None:
            cols = list(t.marker.colors)
    except Exception:
        pass
    pset = set(parents)
    leaves = [(labels[i], vals[i], cols[i] if cols and i < len(cols) else None)
              for i in range(len(labels))
              if labels[i] not in pset and vals[i] > 0]
    if not leaves:
        leaves = [(labels[i], vals[i], cols[i] if cols and i < len(cols)
                   else None) for i in range(len(labels))
                  if parents[i] and vals[i] > 0]
    leaves.sort(key=lambda kv: kv[1], reverse=True)
    rects = _squarify([v for _, v, _ in leaves], x, y, w, h)
    for (lab, val, c), r in zip(leaves, rects):
        col = _rgb(c) if isinstance(c, str) and c.startswith("#") else _palette(0)
        shp = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, E(r["x"]), E(r["y"]),
            E(max(1.0, r["dx"])), E(max(1.0, r["dy"])))
        shp.fill.solid()
        shp.fill.fore_color.rgb = col
        shp.line.color.rgb = _C_WHITE
        shp.line.width = Pt(1)
        shp.shadow.inherit = False
        if r["dx"] > 46 and r["dy"] > 24:
            tf = shp.text_frame
            tf.word_wrap = True
            for m in ("margin_left", "margin_right", "margin_top",
                      "margin_bottom"):
                setattr(tf, m, Emu(20000))
            ink = _readable_on(col)
            _set_text(tf, [[(lab, 7.5, True, ink)], [(_abbr(val), 7, False, ink)]],
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def _empty_panel_note(slide, rect, msg="No data for this view."):
    """Caption drawn where a chart would go when there is nothing to plot.

    A panel with no data is a legitimate outcome (a workbook can size a
    breakdown in one unit only); it must read as an empty panel, not take the
    export down with it.
    """
    x, y, w, h = rect
    tb = _textbox(slide, x + 8, y + max(10, h / 2 - 14), max(40, w - 16), 28)
    tb.text_frame.word_wrap = True
    _set_text(tb.text_frame, [[(msg, 9.5, False, _C_MUTED)]],
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return tb


def _render_composition(slide, rect, meta):
    """Native 100%-stacked horizontal bars: one bar per group (largest first,
    with its size in the row label), split into its parts' share; % labels sit
    centred in each segment (thin ones stay blank), legend below."""
    cats, parts, shares = meta["cats"], meta["parts"], meta["shares"]
    if not cats or not parts:
        # e.g. the volume composition of a segment the workbook sizes in value
        # only - python-pptx raises "chart data contains no categories"
        return _empty_panel_note(slide, rect)
    colors = [_rgb(c) for c in meta["colors"]]
    # PowerPoint plots the first category at the bottom - reverse so the largest
    # group (first in cats) ends up on top, matching the dashboard.
    idx = list(range(len(cats)))[::-1]
    x, y, w, h = rect
    cd = CategoryChartData()
    cd.categories = [cats[i] for i in idx]
    for pi, p in enumerate(parts):
        cd.add_series(p, [shares[pi][i] for i in idx])
    gf = slide.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED,
                                E(x), E(y), E(w), E(h), cd)
    chart = gf.chart
    for pi in range(len(parts)):
        s = chart.series[pi]
        try:
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = colors[pi]
            s.format.line.fill.background()
        except Exception:
            pass
    _style_chart(chart, len(parts), legend=True)
    try:                                            # lock the axis to 0..100 %
        va = chart.value_axis
        va.minimum_scale, va.maximum_scale = 0, 100
        va.tick_labels.number_format = '0"%"'
        va.tick_labels.number_format_is_linked = False
    except Exception:
        pass
    # % labels centred in each segment; skip segments too thin to hold a number
    for pi in range(len(parts)):
        s = chart.series[pi]
        ink = _readable_on(colors[pi])
        for j, i in enumerate(idx):
            v = shares[pi][i]
            if v < 7:
                continue
            try:
                dl = s.points[j].data_label
                tf = dl.text_frame
                tf.text = f"{v:.0f}%"
                r = tf.paragraphs[0].runs[0]
                r.font.size = Pt(7.5)
                r.font.name = FONT
                r.font.color.rgb = ink
                dl.position = XL_DATA_LABEL_POSITION.CENTER
                _dlbls_horizontal(dl)
            except Exception:
                pass
    return chart


def _dot(slide, cx, cy, d, color):
    """A filled circle centred at (cx, cy) with diameter d (px)."""
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, E(cx - d / 2), E(cy - d / 2),
                                 E(d), E(d))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def _tri(slide, cx, cy, d, color, up=True):
    """A filled triangle centred at (cx, cy), ``d`` px wide - pointing UP for a
    gain and DOWN for a loss.  PowerPoint has no down-triangle autoshape, so a
    down marker is the isosceles triangle turned through 180 degrees."""
    shp = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                 E(cx - d / 2), E(cy - d / 2), E(d), E(d))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    if not up:
        shp.rotation = 180.0
    return shp

def _render_dumbbell(slide, rect, meta):
    """Native connected-dot ('dumbbell') chart: per category a label, a start
    dot, a connector bar, an end dot and the end value - a distinctive
    consultant-style view rebuilt entirely from shapes (editable)."""
    x, y, w, h = rect
    cats, start, end = meta["cats"], meta["start"], meta["end"]
    n = len(cats)
    if n == 0:
        return
    lab_w = min(0.42 * w, 180)
    ax_l, ax_r = x + lab_w + 8, x + w - 56
    top, bot = y + 24, y + h - 6
    row_h = (bot - top) / n
    x0 = float(meta.get("x0", 0.0))               # axis floor (zoom band)
    mx = float(meta.get("xmax") or ((max(end) if end else 1) * 1.18)) or 1
    span = (mx - x0) or 1
    c_start, c_line = _rgb(STRIP_START), _rgb(STRIP_LINE)

    def sx(v):
        return ax_l + ((v - x0) / span) * (ax_r - ax_l)

    # meta is sorted ascending (largest last) -> draw the largest on top
    for k, i in enumerate(reversed(range(n))):
        cy = top + (k + 0.5) * row_h
        tb = _textbox(slide, x, cy - 9, lab_w, 18)
        _set_text(tb.text_frame, [[(cats[i], 8.5, False, _C_TEXT)]],
                  anchor=MSO_ANCHOR.MIDDLE)
        cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                        E(sx(start[i])), E(cy),
                                        E(sx(end[i])), E(cy))
        cn.line.color.rgb = c_line
        cn.line.width = Pt(2.5)
        cn.shadow.inherit = False
        _dot(slide, sx(start[i]), cy, 9, c_start)
        _dot(slide, sx(end[i]), cy, 11, _C_ACCENT)
        tb2 = _textbox(slide, sx(end[i]) + 8, cy - 8, 66, 16)
        _set_text(tb2.text_frame, [[(_abbr(end[i]), 8, True, _C_HEADER)]],
                  anchor=MSO_ANCHOR.MIDDLE)
        tb2.text_frame.word_wrap = False          # keep the value on one line
    # legend: start-year (light dot) and end-year (accent dot)
    lx = x
    for lab, col in ((meta["y0"], c_start), (meta["y1"], _C_ACCENT)):
        _dot(slide, lx + 5, y + 9, 9, col)
        tb = _textbox(slide, lx + 12, y + 1, 56, 16)
        _set_text(tb.text_frame, [[(str(lab), 8.5, False, _C_MUTED)]],
                  anchor=MSO_ANCHOR.MIDDLE)
        lx += 74

def _arrow_line(slide, x1, y1, x2, y2, color, width=1.25, both=True):
    """A straight connector with a triangle arrowhead at its end (and, with
    ``both``, at its start too)."""
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                    E(x1), E(y1), E(x2), E(y2))
    cn.line.color.rgb = color
    cn.line.width = Pt(width)
    cn.shadow.inherit = False
    ln = cn.line._get_or_add_ln()
    for tag in (("a:headEnd", "a:tailEnd") if both else ("a:tailEnd",)):
        for old in ln.findall(qn(tag)):
            ln.remove(old)
        ln.append(parse_xml(f'<{tag} {nsdecls("a")} type="triangle" '
                            f'w="med" len="med"/>'))
    return cn

def _render_scenario_bars(slide, rect, meta):
    """Scenario forecast (reference 'Scenario Outlook' left chart): the
    base-year bar on the left, the three horizon-year scenario bars on the
    right, and one "CAGR: x%" pill per scenario riding a double-headed arrow
    that runs from the base bar to its scenario bar.  Drawn from shapes so the
    arrows can be placed exactly; every height and label comes from the data,
    and each pill's CAGR is a figure the locked deck padlocks."""
    x, y, w, h = rect
    bars = meta["bars"]                                 # [(name, end, cagr)]
    cols = meta["colors"]
    leg_h, yr_h = 26, 22
    base_y = y + h - leg_h - yr_h                       # the bars' floor
    plot_top = y + 4
    plot_h = max(40.0, base_y - plot_top)
    u = (w - 8) / 6.0
    bw = u * 0.68
    cx_base = x + 4 + u * 0.8
    cxs = [x + 4 + u * (3.3 + 1.05 * i) for i in range(len(bars))]
    vals = [meta["base"]] + [b[1] for b in bars]
    vmax = max([v for v in vals if v == v and v is not None] + [1e-9])
    pill_h, row_gap = 22, 27

    def rows_for(frac):
        tops = [base_y - max(0.0, b[1]) / vmax * plot_h * frac for b in bars]
        ys, lim = [], None
        for k, top in enumerate(tops):
            yk = top - 14                                # just above its bar
            if k:
                yk = min(yk, ys[-1] - row_gap,           # clear of the row
                         min(tops[:k]) - 14)             # and bars it crosses
            ys.append(yk)
        return tops, ys

    frac = 0.95
    tops, ys = rows_for(frac)
    while frac > 0.35 and ys and min(ys) < plot_top + pill_h / 2:
        frac -= 0.05
        tops, ys = rows_for(frac)

    # the floor
    fl = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x), E(base_y),
                                    E(x + w), E(base_y))
    fl.line.color.rgb = _rgb("#BFBFBF")
    fl.line.width = Pt(0.75)

    def bar(cx, top, color):
        hh = max(1.0, base_y - top)
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(cx - bw / 2),
                                     E(base_y - hh), E(bw), E(hh))
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(color)
        shp.line.fill.background()
        shp.shadow.inherit = False

    base_top = base_y - max(0.0, meta["base"]) / vmax * plot_h * frac
    bar(cx_base, base_top, cols["Base"])
    for (name, _, _), cx, top in zip(bars, cxs, tops):
        bar(cx, top, cols[name])

    # CAGR arrows + pills
    ax0 = cx_base + bw / 2 + 12
    pw = min(150.0, max(90.0, (cxs[0] - bw / 2) - ax0 - 24))
    pcx = ax0 + ((cxs[0] - bw / 2) - ax0) / 2
    for (name, _, cagr), cx, yk in zip(bars, cxs, ys):
        col = _rgb(cols[name])
        _arrow_line(slide, ax0, yk, cx - bw / 2 - 2, yk, col)
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      E(pcx - pw / 2), E(yk - pill_h / 2),
                                      E(pw), E(pill_h))
        pill.fill.solid()
        pill.fill.fore_color.rgb = col
        pill.line.fill.background()
        pill.shadow.inherit = False
        txt = f"CAGR: {components.fmt_pct(cagr)}"
        _set_text(pill.text_frame, [[(txt, 11, True, _C_WHITE)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                  pad=(2, 2, 0, 0))

    # year captions under the base bar and under the middle scenario bar
    for cap, cx in ((meta["base_year"], cx_base),
                    (meta["end_year"], cxs[len(cxs) // 2])):
        tb = _textbox(slide, cx - 40, base_y + 3, 80, yr_h - 3)
        _set_text(tb.text_frame, [[(str(cap), 11, True, _C_TEXT)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # legend: a dot + bold name per scenario, spread across the width
    ly = y + h - leg_h / 2
    slot = w / max(1, len(bars))
    for i, (name, _, _) in enumerate(bars):
        lx = x + i * slot + 6
        _dot(slide, lx + 8, ly, 14, _rgb(cols[name]))
        tb = _textbox(slide, lx + 20, ly - 10, slot - 26, 20)
        _set_text(tb.text_frame, [[(f"{name} Scenario", 11, True, _C_TEXT)]],
                  anchor=MSO_ANCHOR.MIDDLE)
        tb.text_frame.word_wrap = False

def _render_scenario_lines(slide, rect, meta):
    """Scenario growth lines (reference 'Scenario Outlook' right chart): one
    marker line per scenario across the forecast years on a 0.00% axis, no
    point labels - the axis carries the reading, so nothing here is locked."""
    x, y, w, h = rect
    cd = CategoryChartData()
    cd.categories = list(meta["years"])
    for name, vals, _ in meta["series"]:
        cd.add_series(str(name), [_num(v) for v in vals])
    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
                                E(x), E(y), E(w), E(h), cd)
    chart = gf.chart
    _style_chart(chart, len(meta["series"]), legend=True, height_px=h)
    for s, (_, _, color) in zip(chart.series, meta["series"]):
        col = _rgb(color)
        s.smooth = False
        s.format.line.color.rgb = col
        s.format.line.width = Pt(2.25)
        try:
            from pptx.enum.chart import XL_MARKER_STYLE
            s.marker.style = XL_MARKER_STYLE.CIRCLE
            s.marker.size = 6
            s.marker.format.fill.solid()
            s.marker.format.fill.fore_color.rgb = col
            s.marker.format.line.color.rgb = col
        except Exception:
            pass
    allv = [v for _, vs, _ in meta["series"] for v in vs if v is not None]
    va = chart.value_axis
    va.tick_labels.number_format = "0.00%"
    va.tick_labels.number_format_is_linked = False
    if allv and min(allv) >= 0:
        va.minimum_scale = 0
    return chart

def _render_diverging(slide, rect, meta):
    """Native diverging column chart - bars grow up (green) / down (cyan) from a
    dashed, centred zero baseline; value labels + category labels below.  The
    editable twin of figures.diverging_bars."""
    x, y, w, h = rect
    names, vals, colors = meta["names"], meta["vals"], meta["colors"]
    m = float(meta.get("m") or 1.0)
    suffix = meta.get("suffix", "")
    n = len(names)
    if n == 0:
        return
    lab_h = 30.0
    plot_l, plot_r = x + 40, x + w - 8
    top, bot = y + 12, y + h - lab_h
    slot = (plot_r - plot_l) / n
    bar_w = min(slot * 0.66, 48.0)          # gap ≈ 50% of bar width
    zero_y = (top + bot) / 2.0                     # symmetric range -> centred

    def sy(v):
        return bot - ((v + m) / (2 * m)) * (bot - top)

    # left value axis: +m / baseline / -m.  When the caller named the baseline
    # (the pricing page passes "Global"), the middle tick shows that word
    # instead of "0" - the line means "same as Global", and the deck must not
    # keep reading as a zero baseline.
    blabel = meta.get("blabel") or ""
    for v in (round(m * 0.85, 1), 0.0, -round(m * 0.85, 1)):
        if v:
            txt, bold = f"{v:+.0f}{suffix}", False
        else:
            txt, bold = (blabel, True) if blabel else (f"0{suffix}", False)
        tb = _textbox(slide, x - 2, sy(v) - 7, 38, 14)
        _set_text(tb.text_frame, [[(txt, 7.5, bold, _C_MUTED)]],
                  align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    # dashed baseline
    zl = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(plot_l),
                                    E(zero_y), E(plot_r), E(zero_y))
    zl.line.color.rgb = _C_MUTED
    zl.line.width = Pt(1.1)
    zl.shadow.inherit = False
    try:
        zl.line._get_or_add_ln().append(
            parse_xml(f'<a:prstDash {nsdecls("a")} val="dash"/>'))
    except Exception:
        pass
    for i, name in enumerate(names):
        cx = plot_l + (i + 0.5) * slot
        v = vals[i]
        yv = sy(v)
        col = _rgb(colors[i])
        y_top, y_bot = min(zero_y, yv), max(zero_y, yv)
        if y_bot - y_top >= 1:
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(cx - bar_w / 2),
                                         E(y_top), E(bar_w), E(y_bot - y_top))
            bar.fill.solid()
            bar.fill.fore_color.rgb = col
            bar.line.fill.background()
            bar.shadow.inherit = False
        # value label just past the bar end (above if +, below if −)
        vy = (yv - 15) if v >= 0 else (yv + 1)
        vl = _textbox(slide, cx - slot / 2, vy, slot, 14)
        _set_text(vl.text_frame, [[(f"{v:+.1f}{suffix}", 8, True, _C_TEXT)]],
                  align=PP_ALIGN.CENTER)
        # category label at the bottom
        cl = _textbox(slide, cx - slot / 2, bot + 3, slot, lab_h - 4)
        cl.text_frame.word_wrap = True
        _set_text(cl.text_frame, [[(str(name), 7.5, False, _C_TEXT)]],
                  align=PP_ALIGN.CENTER)

def _render_boxplot(slide, rect, meta):
    """Native VERTICAL box plot - per category a whisker (min→max) with a box
    (Q1→Q3), a median line and end caps, standing on a shared value axis.  The
    editable twin of figures.price_box."""
    x, y, w, h = rect
    stats = meta["stats"]
    n = len(stats)
    if n == 0:
        return
    y0 = float(meta.get("y0", 0.0))
    y1 = float(meta.get("y1", 1.0))
    span = (y1 - y0) or 1.0
    lab_h = 34.0                                   # category labels at bottom
    # the left gutter is sized to the WIDEST tick text: a small-unit market
    # prints "0.000145", which wrapped onto two lines in the old fixed 40px box
    ticks = [(y0 + f * span, bot_f) for f, bot_f in
             ((0.0, 0.0), (0.5, 0.5), (1.0, 1.0))]
    tick_txt = [_sig_num(v) for v, _ in ticks]
    gut = max(40.0, min(78.0, max((layout.text_width(t, 7.5 * _PPT_FONT_K)
                                   for t in tick_txt), default=40.0) + 8))
    plot_l, plot_r = x + gut + 4, x + w - 8        # left value-axis gutter
    top, bot = y + 12, y + h - lab_h
    slot = (plot_r - plot_l) / n
    box_w = min(slot * 0.5, 46.0)
    box_fill, box_line = _rgb("#CFE7F5"), _C_HEADER
    whisk = _rgb("#8FB6CC")

    def sy(v):
        return bot - ((v - y0) / span) * (bot - top)

    # left value axis (scale)
    for (v, frac), txt in zip(ticks, tick_txt):
        ty = bot - frac * (bot - top)
        tb = _textbox(slide, x - 2, ty - 7, gut, 14)
        tb.text_frame.word_wrap = False
        _set_text(tb.text_frame, [[(txt, 7.5, False, _C_MUTED)]],
                  align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    for i, st in enumerate(stats):
        cx = plot_l + (i + 0.5) * slot
        # whisker (low -> high)
        wl = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(cx),
                                        E(sy(st["low"])), E(cx), E(sy(st["high"])))
        wl.line.color.rgb = whisk
        wl.line.width = Pt(1.4)
        wl.shadow.inherit = False
        for v in (st["low"], st["high"]):          # end caps
            cap = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT, E(cx - box_w * 0.3), E(sy(v)),
                E(cx + box_w * 0.3), E(sy(v)))
            cap.line.color.rgb = whisk
            cap.line.width = Pt(1.4)
            cap.shadow.inherit = False
        # box (Q1 -> Q3)
        by0, by1 = sy(st["q3"]), sy(st["q1"])       # q3 is higher (smaller y px)
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(cx - box_w / 2),
                                     E(by0), E(box_w), E(max(2.0, by1 - by0)))
        box.fill.solid()
        box.fill.fore_color.rgb = box_fill
        box.line.color.rgb = box_line
        box.line.width = Pt(1.25)
        box.shadow.inherit = False
        # median line
        med = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, E(cx - box_w / 2), E(sy(st["med"])),
            E(cx + box_w / 2), E(sy(st["med"])))
        med.line.color.rgb = _C_ACCENT
        med.line.width = Pt(2)
        med.shadow.inherit = False
        # median value label just above the box
        vl = _textbox(slide, cx - slot / 2, sy(st["high"]) - 16, slot, 14)
        _set_text(vl.text_frame, [[(_sig_num(st["med"]), 7.5, True, _C_HEADER)]],
                  align=PP_ALIGN.CENTER)
        # category label below
        cl = _textbox(slide, cx - slot / 2, bot + 3, slot, lab_h - 4)
        cl.text_frame.word_wrap = True
        _set_text(cl.text_frame, [[(st["name"], 7.5, False, _C_TEXT)]],
                  align=PP_ALIGN.CENTER)

def _render_candlestick(slide, rect, meta):
    """Native OHLC candlestick - per category a thin wick (low→high) and a body
    (open→close): green when close≥open, red when it falls, on a shared value
    axis.  The editable twin of figures.candlestick."""
    x, y, w, h = rect
    candles = meta["candles"]
    n = len(candles)
    if n == 0:
        return
    y0 = float(meta.get("y0", 0.0))
    y1 = float(meta.get("y1", 1.0))
    span = (y1 - y0) or 1.0
    lab_h = 34.0
    plot_l, plot_r = x + 44, x + w - 8
    top, bot = y + 12, y + h - lab_h
    slot = (plot_r - plot_l) / n
    body_w = min(slot * 0.5, 46.0)
    uni = meta.get("uniform", "")
    if uni:
        up_f = dn_f = _rgb(uni)
        up_l = dn_l = _rgb(NEUTRAL_LINE)
    else:
        up_f, up_l = _rgb(UP_FILL), _rgb(UP_LINE)
        dn_f, dn_l = _rgb(DN_FILL), _rgb(DN_LINE)

    def sy(v):
        return bot - ((v - y0) / span) * (bot - top)

    for frac in (0.0, 0.5, 1.0):                    # left value axis
        v = y0 + frac * span
        ty = bot - frac * (bot - top)
        tb = _textbox(slide, x - 2, ty - 7, 40, 14)
        _set_text(tb.text_frame, [[(f"{v:.1f}", 7.5, False, _C_MUTED)]],
                  align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    for i, cd in enumerate(candles):
        cx = plot_l + (i + 0.5) * slot
        fill = up_f if cd["up"] else dn_f
        line = up_l if cd["up"] else dn_l
        # wick low -> high
        wk = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(cx),
                                        E(sy(cd["high"])), E(cx), E(sy(cd["low"])))
        wk.line.color.rgb = line
        wk.line.width = Pt(1.4)
        wk.shadow.inherit = False
        # body open -> close
        top_v, bot_v = max(cd["open"], cd["close"]), min(cd["open"], cd["close"])
        by0 = sy(top_v)
        body = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(cx - body_w / 2),
                                      E(by0), E(body_w),
                                      E(max(2.0, sy(bot_v) - by0)))
        body.fill.solid()
        body.fill.fore_color.rgb = fill
        body.line.color.rgb = line
        body.line.width = Pt(1.25)
        body.shadow.inherit = False
        # close value label just above the candle
        vl = _textbox(slide, cx - slot / 2, sy(cd["high"]) - 16, slot, 14)
        _set_text(vl.text_frame, [[(f"{cd['close']:.1f}", 7.5, True, _C_HEADER)]],
                  align=PP_ALIGN.CENTER)
        # category label below
        cl = _textbox(slide, cx - slot / 2, bot + 3, slot, lab_h - 4)
        cl.text_frame.word_wrap = True
        _set_text(cl.text_frame, [[(cd["name"], 7.5, False, _C_TEXT)]],
                  align=PP_ALIGN.CENTER)

def _render_price_strip(slide, rect, meta):
    """Native price-point strip - per row a light-blue price-RANGE band on a
    shared price axis plus an orange average dot (avg_only) or coloured product
    dots.  The editable twin of figures.price_strip."""
    x, y, w, h = rect
    rows = meta["rows"]
    n = len(rows)
    if n == 0:
        return
    xmin, xmax = float(meta["x_min"]), float(meta["x_max"])
    span = (xmax - xmin) or 1.0
    lab_w = 118.0
    plot_l, plot_r = x + lab_w, x + w - 14
    top, bot = y + 34, y + h - 24
    grid = _rgb("#E3E8EC")
    band_fill, band_line = _rgb(meta["band_fill"]), _rgb(meta["band_line"])
    avg_col, avg_edge = _rgb(meta["avg_color"]), _rgb(meta["avg_edge"])
    white = _rgb("#FFFFFF")

    def sx(v):
        return plot_l + (v - xmin) / span * (plot_r - plot_l)

    row_h = (bot - top) / n
    tick_fmt = "{:,.1f}" if span < 20 else "{:,.0f}"
    for k in range(6):                              # axis gridlines + ticks
        v = xmin + k / 5.0 * span
        gx = sx(v)
        gl = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(gx),
                                        E(top - 4), E(gx), E(bot))
        gl.line.color.rgb = grid
        gl.line.width = Pt(0.5)
        gl.shadow.inherit = False
        tb = _textbox(slide, gx - 26, bot + 2, 52, 12)
        _set_text(tb.text_frame, [[(tick_fmt.format(v), 7, False, _C_MUTED)]],
                  align=PP_ALIGN.CENTER)
    for i, r in enumerate(rows):
        cy = top + (i + 0.5) * row_h
        lb = _textbox(slide, x, cy - 9, lab_w - 6, 18)
        lb.text_frame.word_wrap = True
        _set_text(lb.text_frame, [[(r["name"], 8.5, True, _C_HEADER)]],
                  anchor=MSO_ANCHOR.MIDDLE)
        bx0, bx1 = sx(r["lo"]), sx(r["hi"])
        band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(bx0),
                                      E(cy - row_h * 0.26),
                                      E(max(3.0, bx1 - bx0)), E(row_h * 0.52))
        band.fill.solid()
        band.fill.fore_color.rgb = band_fill
        band.line.color.rgb = band_line
        band.line.width = Pt(1)
        band.shadow.inherit = False
        if meta["avg_only"]:
            _edged_dot(slide, sx(r["avg"]), cy, 13, avg_col, avg_edge)
        else:
            cols = meta["colors"]
            for si, v in enumerate(r["dots"]):
                _edged_dot(slide, sx(v), cy, 11,
                           _rgb(cols[si]) if si < len(cols) else avg_col, white)
    # top legend
    lx = plot_l
    if meta["avg_only"]:
        _edged_dot(slide, lx + 6, top - 15, 12, avg_col, avg_edge)
        t = _textbox(slide, lx + 14, top - 22, 88, 14)
        _set_text(t.text_frame, [[("Average Price", 8, False, _C_TEXT)]])
        sw = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(lx + 108),
                                    E(top - 20), E(20), E(12))
        sw.fill.solid(); sw.fill.fore_color.rgb = band_fill
        sw.line.color.rgb = band_line; sw.line.width = Pt(1)
        sw.shadow.inherit = False
        t2 = _textbox(slide, lx + 132, top - 22, 84, 14)
        _set_text(t2.text_frame, [[("Price Range", 8, False, _C_TEXT)]])
    else:
        cx = lx
        for si, sname in enumerate(meta["series"]):
            _edged_dot(slide, cx + 6, top - 15, 10,
                       _rgb(meta["colors"][si]) if si < len(meta["colors"])
                       else avg_col, white)
            tw = 12 + len(sname) * 5.0
            t = _textbox(slide, cx + 13, top - 22, tw, 14)
            _set_text(t.text_frame, [[(sname, 7.5, False, _C_TEXT)]])
            cx += 13 + tw + 6
        sw = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(cx + 2), E(top - 20),
                                    E(18), E(12))
        sw.fill.solid(); sw.fill.fore_color.rgb = band_fill
        sw.line.color.rgb = band_line; sw.line.width = Pt(1)
        sw.shadow.inherit = False
        t2 = _textbox(slide, cx + 22, top - 22, 80, 14)
        _set_text(t2.text_frame, [[("Price Range", 7.5, False, _C_TEXT)]])

def _polyline(slide, pts, color, width, closed=False):
    """Draw an (optionally closed) unfilled polyline through ``pts`` (px).
    ``color`` may be a hex string or an RGBColor.  Uses a freeform shape when
    available; falls back to straight connectors."""
    col = _rgb(color) if isinstance(color, str) else color
    try:
        fb = slide.shapes.build_freeform(float(pts[0][0]), float(pts[0][1]),
                                         scale=_EMU_PER_PX)
        fb.add_line_segments([(float(px), float(py)) for px, py in pts[1:]],
                             close=closed)
        shp = fb.convert_to_shape()
        shp.fill.background()
        shp.line.color.rgb = col
        shp.line.width = Pt(width)
        shp.shadow.inherit = False
        return shp
    except Exception:
        seq = list(pts) + ([pts[0]] if closed else [])
        for a, b in zip(seq, seq[1:]):
            ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(a[0]),
                                            E(a[1]), E(b[0]), E(b[1]))
            ln.line.color.rgb = col
            ln.line.width = Pt(width)
            ln.shadow.inherit = False
        return None

def _edged_dot(slide, cx, cy, d, color, edge):
    o = slide.shapes.add_shape(MSO_SHAPE.OVAL, E(cx - d / 2), E(cy - d / 2),
                               E(d), E(d))
    o.fill.solid()
    o.fill.fore_color.rgb = color if not isinstance(color, str) else _rgb(color)
    o.line.color.rgb = edge
    o.line.width = Pt(1)
    o.shadow.inherit = False
    return o

def _render_radar(slide, rect, meta):
    """Native radar/spider chart - light-grey concentric decagon rings + spokes,
    a blue-filled data polygon, optional vertex dots + a faint 0..max scale, and
    category labels.  The editable twin of figures.radar."""
    x, y, w, h = rect
    cats = meta["cats"]
    vals = meta["vals"]
    smax = float(meta.get("scale_max") or 10.0) or 10.0
    markers = meta.get("markers", True)
    scale_labels = meta.get("scale_labels", True)
    tickvals = meta.get("tickvals") or []
    n = len(cats)
    if n == 0:
        return
    line_col = _rgb(meta.get("color") or ACCENT)
    cx, cy = x + w / 2.0, y + h / 2.0
    R = max(24.0, min(w, h) / 2.0 - 58.0)          # leave room for labels

    def pt(i, rad):
        a = math.radians(-90.0 + i * 360.0 / n)    # start at top, clockwise
        return (cx + rad * math.cos(a), cy + rad * math.sin(a))

    rings = [tv / smax for tv in tickvals] if tickvals else [.2, .4, .6, .8, 1.]
    for frac in rings:                             # concentric decagon rings
        ring = [pt(i, R * frac) for i in range(n)]
        _polyline(slide, ring, _C_GRID, 0.75, closed=True)
    for i in range(n):                             # radial spokes
        ex, ey = pt(i, R)
        sp = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(cx), E(cy),
                                        E(ex), E(ey))
        sp.line.color.rgb = _C_GRID
        sp.line.width = Pt(0.75)
        sp.shadow.inherit = False
    # filled data polygon
    dpts = [pt(i, R * max(0.0, min(1.0, vals[i] / smax))) for i in range(n)]
    try:
        fb = slide.shapes.build_freeform(float(dpts[0][0]), float(dpts[0][1]),
                                         scale=_EMU_PER_PX)
        fb.add_line_segments([(float(px), float(py)) for px, py in dpts[1:]],
                             close=True)
        poly = fb.convert_to_shape()
        poly.fill.solid()
        poly.fill.fore_color.rgb = line_col
        _set_fill_alpha(poly, 62)
        poly.line.color.rgb = line_col
        poly.line.width = Pt(2.2)
        poly.shadow.inherit = False
    except Exception:
        _polyline(slide, dpts, line_col, 2.2, closed=True)
    if markers:                                    # vertex dots
        for px, py in dpts:
            _edged_dot(slide, px, py, 7.0, line_col, _C_WHITE)
    if scale_labels:                               # faint scale up the top spoke
        for tv in tickvals:
            ty = cy - R * (tv / smax)
            lb = _textbox(slide, cx + 3, ty - 7, 30, 13)
            _set_text(lb.text_frame, [[(f"{tv:g}", 7, False, _C_MUTED)]],
                      align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    # category labels just outside each spoke
    for i, name in enumerate(cats):
        lx, ly = pt(i, R + 22)
        bw, bh = 138.0, 30.0
        tb = _textbox(slide, lx - bw / 2, ly - bh / 2, bw, bh)
        tb.text_frame.word_wrap = True
        _set_text(tb.text_frame, [[(str(name), 8, False, _C_TEXT)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def _render_tier_funnel(slide, rect, meta):
    """Native centred tier funnel - an ordered set of monotonically narrowing
    bands, each centred with its label inside.  Editable twin of
    figures.tier_funnel."""
    x, y, w, h = rect
    bands = meta["bands"]
    u = meta.get("unit", "")
    n = len(bands)
    if n == 0:
        return
    gap = 14.0
    band_h = (h - (n - 1) * gap) / n
    maxw = w * 0.98
    for i, b in enumerate(bands):
        frac = 1.0 - i * (0.66 / max(1, n - 1))
        bw = max(110.0, maxw * frac)
        bx = x + (w - bw) / 2.0
        by = y + i * (band_h + gap)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(bx), E(by),
                                     E(bw), E(band_h))
        bar.fill.solid()
        bar.fill.fore_color.rgb = _rgb(b.get("color", "#1668A8"))
        bar.line.color.rgb = _C_WHITE
        bar.line.width = Pt(1.5)
        bar.shadow.inherit = False
        head = (f"{b['name']}  ·  {b.get('sharetext', '')}  ·  "
                f"{_abbr(b.get('size', 0))} {u}".rstrip(" ·").rstrip())
        paras = [[(head, 11.5, True, _C_WHITE)]]
        if b.get("label"):
            paras.append([(b["label"], 8.5, False, _rgb("#E6EEF5"))])
        _set_text(bar.text_frame, paras, align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE, line=1.05)


def _render_funnel3d(slide, rect, meta):
    """Native 3-D funnel - a light-green cone with a blue cylinder disc per
    segment (light lid / body / dark rim), tier + size above and % on the disc.
    Editable twin of figures.funnel3d."""
    x, y, w, h = rect
    segs = meta["segments"]
    n = len(segs)
    if n == 0:
        return
    cx = x + w / 2.0
    top_m, bot_m = 10, 8
    avail = h - top_m - bot_m
    slot = avail / n
    maxhw = w * 0.42
    ry = min(9.0, slot * 0.15)
    bh = slot * 0.34                                   # full body height

    def hw(i):
        return segs[i]["w"] * maxhw

    # tinted cone behind - hugs each disc (constant width across the disc's full
    # height, tapering only in the gaps between discs) for a clean funnel
    top_y, bot_y = y + top_m, y + h - bot_m
    right = [(cx + hw(0), top_y)]
    for i in range(n):
        cyi = y + top_m + (i + 0.5) * slot
        right.append((cx + hw(i), cyi - bh / 2 - ry))     # disc top (lid)
        right.append((cx + hw(i), cyi + bh / 2 + ry))     # disc bottom (rim)
    right.append((cx + hw(n - 1), bot_y))
    conepts = right + [(2 * cx - px, py) for px, py in reversed(right)]
    try:
        fb = slide.shapes.build_freeform(float(conepts[0][0]),
                                         float(conepts[0][1]), scale=_EMU_PER_PX)
        fb.add_line_segments([(float(a), float(b)) for a, b in conepts[1:]],
                             close=True)
        cone = fb.convert_to_shape()
        cone.fill.solid()
        cone.fill.fore_color.rgb = _rgb(UP_FILL)
        cone.line.fill.background()
        cone.shadow.inherit = False
    except Exception:
        pass
    for i, s in enumerate(segs):
        w_i = hw(i)
        cy = y + top_m + (i + 0.5) * slot
        body_top, body_bot = cy - bh / 2, cy + bh / 2
        _wrect(slide, cx - w_i, body_top, 2 * w_i, bh, _rgb(F3_BODY))
        _wovl(slide, cx - w_i, body_bot - ry, 2 * w_i, 2 * ry, _rgb(F3_RIM))
        lid = _wovl(slide, cx - w_i, body_top - ry, 2 * w_i, 2 * ry,
                    _rgb(F3_TOP))
        lid.line.color.rgb = _C_WHITE
        lid.line.width = Pt(1)
        nb = _textbox(slide, cx - maxhw, body_top - ry - 34, 2 * maxhw, 32)
        _set_text(nb.text_frame,
                  [[(s["name"], 10.5, True, _rgb(F3_LBL))],
                   [(s["size"], 9.5, False, _rgb(F3_LBL))]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, line=1.0)
        pb = _textbox(slide, cx - w_i, body_top, 2 * w_i, bh)
        _set_text(pb.text_frame, [[(s["pct"], 9.5, True, _C_WHITE)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def _render_sam_funnel(slide, rect, meta):
    """Native left-aligned funnel - descending-width bars with a single-line
    white label inside.  Editable twin of figures.sam_funnel."""
    x, y, w, h = rect
    rows = meta["rows"]
    n = len(rows)
    if n == 0:
        return
    gap = 12.0
    band_h = (h - (n - 1) * gap) / n
    for i, r in enumerate(rows):
        frac = 1.0 - i * (0.46 / max(1, n - 1))
        bw = max(120.0, w * frac)
        by = y + i * (band_h + gap)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(x), E(by), E(bw),
                                     E(band_h))
        bar.fill.solid()
        bar.fill.fore_color.rgb = _rgb(r["color"])
        bar.line.color.rgb = _C_WHITE
        bar.line.width = Pt(1.5)
        bar.shadow.inherit = False
        _set_text(bar.text_frame, [[(r["text"], 11.5, True, _C_WHITE)]],
                  align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
                  pad=(12, 6, 2, 2))


def _render_pos_bubble(slide, rect, meta):
    """Native positioning bubble scatter - x/y axes + a few labelled bubbles.
    Editable twin of figures.pos_bubble."""
    x, y, w, h = rect
    pts = meta["points"]
    xmax = float(meta.get("xmax") or 3.5)
    ymax = float(meta.get("ymax") or 4.0)
    plot_l, plot_r = x + 34, x + w - 12
    plot_t, plot_b = y + 10, y + h - 30
    # axes
    for ax in ((plot_l, plot_b, plot_r, plot_b), (plot_l, plot_t, plot_l,
                                                  plot_b)):
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(ax[0]),
                                        E(ax[1]), E(ax[2]), E(ax[3]))
        ln.line.color.rgb = _C_GRID
        ln.line.width = Pt(1)
        ln.shadow.inherit = False

    def sx(vx):
        return plot_l + (vx / xmax) * (plot_r - plot_l)

    def sy(vy):
        return plot_b - (vy / ymax) * (plot_b - plot_t)

    # x tick labels
    for t in range(0, int(xmax) + 1):
        tb = _textbox(slide, sx(t) - 12, plot_b + 2, 24, 12)
        _set_text(tb.text_frame, [[(str(t), 7, False, _C_MUTED)]],
                  align=PP_ALIGN.CENTER)
    xl = _textbox(slide, plot_l, plot_b + 14, plot_r - plot_l, 12)
    _set_text(xl.text_frame, [[(meta.get("xlab", ""), 8, False, _C_MUTED)]],
              align=PP_ALIGN.CENTER)
    yl = _textbox(slide, x - 24, plot_t + (plot_b - plot_t) / 2 - 8, 60, 16)
    yl.rotation = 270
    _set_text(yl.text_frame, [[(meta.get("ylab", ""), 8, False, _C_MUTED)]],
              align=PP_ALIGN.CENTER)
    for pt in pts:                                  # bubbles + labels
        cx, cy = sx(pt["x"]), sy(pt["y"])
        d = pt.get("size", 26)
        _circle(slide, cx, cy, d, fill=_rgb(pt["color"]), line=_C_WHITE, lw=1.5)
        lb = _textbox(slide, cx + d / 2 + 4, cy - 9, 150, 18)
        _set_text(lb.text_frame, [[(pt["name"], 9.5, True, _C_TEXT)]],
                  anchor=MSO_ANCHOR.MIDDLE)


def _render_target_actual(slide, rect, meta):
    """Native horizontal grouped bars - Target (dark) vs Actual (light) per
    metric.  Editable twin of figures.target_actual_bars."""
    x, y, w, h = rect
    rows = meta["rows"]
    vmax = float(meta.get("vmax") or 5.0)
    n = len(rows)
    if n == 0:
        return
    lab_w = 118
    plot_l, plot_r = x + lab_w, x + w - 14
    top, bot = y + 20, y + h - 20
    band = (bot - top) / n
    bar_h = min(11.0, band * 0.30)

    def sx(v):
        return plot_l + (v / vmax) * (plot_r - plot_l)

    # legend
    for j, (nm, col) in enumerate((("Target", _rgb(TA_TARGET)),
                                   ("Actual", _rgb(TA_ACTUAL)))):
        lx = plot_r - 150 + j * 78
        sw = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(lx), E(y + 2),
                                    E(10), E(10))
        sw.fill.solid()
        sw.fill.fore_color.rgb = col
        sw.line.fill.background()
        sw.shadow.inherit = False
        lt = _textbox(slide, lx + 12, y - 1, 60, 14)
        _set_text(lt.text_frame, [[(nm, 8, False, _C_TEXT)]])
    for i, r in enumerate(rows):
        cy = top + i * band + band / 2
        cl = _textbox(slide, x + 2, cy - 12, lab_w - 6, 24)
        _set_text(cl.text_frame, [[(r["name"], 8.5, False, _C_TEXT)]],
                  anchor=MSO_ANCHOR.MIDDLE)
        for j, (key, col) in enumerate((("target", _rgb(TA_TARGET)),
                                        ("actual", _rgb(TA_ACTUAL)))):
            v = float(r.get(key, 0))
            by = cy - bar_h - 1 + j * (bar_h + 2)
            bw = max(1.0, sx(v) - plot_l)
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(plot_l), E(by),
                                         E(bw), E(bar_h))
            bar.fill.solid()
            bar.fill.fore_color.rgb = col
            bar.line.fill.background()
            bar.shadow.inherit = False

def _render_period_band(slide, panel_rect, chart_rect, pb):
    """A native Historical/Forecast bracket in the strip above the chart,
    aligned to the year categories via an estimated plot x-range."""
    hc, n = pb["hc"], pb["n"]
    if hc <= 0 or hc >= n:
        return
    px, py, pw, ph = panel_rect
    cx, cy, cw, ch = chart_rect
    plot_l = cx + 42                              # est. left inset (y-axis nums)
    plot_r = cx + cw - 12
    bx = plot_l + (hc / n) * (plot_r - plot_l)    # boundary x
    yline, ytick = py + 20, py + 26

    def bracket(x0, x1, color):
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                        E(x0), E(yline), E(x1), E(yline))
        ln.line.color.rgb = color
        ln.line.width = Pt(2)
        ln.shadow.inherit = False
        for xc in (x0, x1):
            t = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                           E(xc), E(yline), E(xc), E(ytick))
            t.line.color.rgb = color
            t.line.width = Pt(2)
            t.shadow.inherit = False

    bracket(plot_l + 2, bx - 6, _C_MUTED)
    bracket(bx + 6, plot_r - 2, _C_ACCENT)
    dv = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                    E(bx), E(cy + 4), E(bx), E(cy + ch * 0.8))
    dv.line.color.rgb = _rgb("#C7CFD6")
    dv.line.width = Pt(1)
    dv.shadow.inherit = False
    try:
        dv.line._get_or_add_ln().append(
            parse_xml(f'<a:prstDash {nsdecls("a")} val="sysDot"/>'))
    except Exception:
        pass
    hlab = _textbox(slide, plot_l, py + 3, max(20, bx - plot_l), 14)
    _set_text(hlab.text_frame,
              [[(f"Historical  {pb['first']}–{pb['hist_end']}", 8, False,
                 _C_MUTED)]], align=PP_ALIGN.CENTER)
    flab = _textbox(slide, bx, py + 3, max(20, plot_r - bx), 14)
    _set_text(flab.text_frame,
              [[(f"Forecast  {pb['hist_end'] + 1}–{pb['last']}", 8, False,
                 _C_ACCENT)]], align=PP_ALIGN.CENTER)

def _stacked_share_chart(slide, rect, years, series, names, pct, legend=True):
    """A REAL 100%-stacked column chart for a share-mix panel.

    These panels used to be drawn as a pile of rectangles - a picture of a
    chart, with no series, no categories and nothing behind it to edit.  As a
    native chart the deck carries the numbers: click it in PowerPoint and
    "Edit Data" opens the shares, the legend wraps itself, and the segment
    labels are linked to the values instead of being painted on."""
    key = "shares" if pct else "vals"
    stacked = [(str(nm), [float(v) for v in s[key]], _rgb(s["color"]))
               for nm, s in zip(names, series)]
    return _add_cat_chart(slide, rect, XL_CHART_TYPE.COLUMN_STACKED_100,
                          [str(y) for y in years], stacked,
                          legend=legend, pct=pct, pct_decimals=0)

# What a native column chart reserves BELOW its plot area, in slide px:
# one band for the category labels and one per wrapped legend row.  Nothing in
# the file reports the plot area, so the BPS gutter beside the chart has to
# predict it; both numbers are measured off a rendered deck.
_CAT_BAND, _LEG_ROW, _PLOT_TOP = 30.0, 30.0, 0.04
_LEG_ITEM_W = 175.0                       # legend entry width before it wraps

def _plot_band(y, h, width, n_series):
    """(top, bottom) of a native column chart's plot area inside (y, h)."""
    per_row = max(1, int(width // _LEG_ITEM_W))
    rows = max(1, -(-n_series // per_row))
    drop = _CAT_BAND + rows * _LEG_ROW
    return y + h * _PLOT_TOP, y + max(h * 0.5, h - drop)

def _render_comp_bps(slide, rect, meta):
    """Native twin of figures.composition_bps - the share mix as a REAL
    100%-stacked column chart, with the BPS share shift beside it as green /
    red markers read against their own scale, exactly as the dashboard draws
    them.

    Only the columns became a chart.  The BPS half stays a marker scale
    because that is what it is: seven one-point readings sharing a gutter, not
    a series of bars.  ``BPS_BAND`` lines the gutter up with the chart's plot
    area - PowerPoint decides that area itself, so the fractions below are
    measured against a rendered deck rather than derived."""
    x, y, w, h = rect
    years, series = meta["years"], meta["series"]
    if not years or not series:
        return
    bvals = [float(v) for v in (meta.get("bvals") or [])]
    blim = float(meta.get("blim") or 1.0) or 1.0
    bcolors = meta.get("bcolors") or []
    btext = meta.get("btext") or []
    pct = bool(meta.get("pct"))
    names = meta.get("snames") or [s["name"] for s in series]
    show_bps = bool(bvals) and len(bvals) == len(btext)
    cap = 15.0                                   # caption strip over the pair
    gut = min(max(w * 0.22, 120.0), 190.0) if show_bps else 0.0
    chart = _stacked_share_chart(slide, (x, y + cap, w - gut, h - cap),
                                 years, series, names, pct)
    if not show_bps:
        return chart

    # the gutter's scale spans the same band as the chart's plot area
    plot_t, plot_b = _plot_band(y + cap, h - cap, w - gut, len(series))
    axis_x = x + w - 46                          # the BPS axis line
    dot_x = x + w - gut + 16                     # the markers

    def yb(v):
        return plot_b - (v + blim) / (2 * blim) * (plot_b - plot_t)

    # A bold "BPS" caption, centred over the WHOLE gutter and sized to read
    # in the 2-3 seconds a viewer actually spends on it - the small print
    # this replaces ("BPS Δ", 7.5pt, tucked at the axis's top-left corner)
    # is easy to miss even here, and on a locked export it is the only thing
    # left standing once the axis ticks beside it are blurred: without it,
    # a blurred scale is a scale for nothing.  It fills the blank strip
    # above the markers instead, so the UNIT reads before a single value
    # does, locked or not.
    gut_l = x + w - gut
    hdr_h = max(14.0, plot_t - (y + cap) - 3)
    hdr = _textbox(slide, gut_l, y + cap + 1, (x + w) - gut_l, hdr_h)
    _set_text(hdr.text_frame,
              [[(f"BPS Δ  ({years[0]}→{years[-1]})", 13.0, True,
                 _C_ACCENT)]],
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    ax = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                    E(axis_x), E(plot_t), E(axis_x), E(plot_b))
    ax.line.color.rgb = _C_MUTED
    ax.line.width = Pt(1)
    ax.shadow.inherit = False
    z0 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(dot_x - 10),
                                    E(yb(0)), E(axis_x), E(yb(0)))
    z0.line.color.rgb = _rgb("#C7CFD6")
    z0.line.width = Pt(0.75)
    z0.shadow.inherit = False
    for bt in (-blim / 1.3, 0.0, blim / 1.3):
        tk = _textbox(slide, axis_x + 3, yb(bt) - 7, 44, 14)
        _set_text(tk.text_frame, [[(f"{bt:+,.0f}", 7.5, False, _C_MUTED)]],
                  align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    # markers first, then labels nudged apart so two close readings stay legible
    ys = [yb(v) for v in bvals]
    for i in sorted(range(len(bvals)), key=lambda k: ys[k]):
        # a green ▲ for a share gain, a red ▼ for a loss - the read-out
        # bullets beside this chart name exactly these two marks
        _tri(slide, dot_x, ys[i], 10, _rgb(bcolors[i]), up=bvals[i] >= 0)
    placed = []
    for i in sorted(range(len(bvals)), key=lambda k: ys[k]):
        ty = ys[i]
        if placed and ty - placed[-1] < 13:
            ty = placed[-1] + 13
        placed.append(ty)
        lb = _textbox(slide, dot_x + 7, ty - 8, 62, 16)
        lb.text_frame.word_wrap = False
        # the marker IS the arrow now, so the label never repeats the glyph
        txt = str(btext[i]).lstrip("▲▼ ").strip() or btext[i]
        _set_text(lb.text_frame, [[(txt, 8, True, _rgb(bcolors[i]))]],
                  align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    return chart

def _render_pct_stack(slide, rect, meta):
    """Native 100%-stacked columns - one full-height bar per year, split by
    series share.  The editable twin of figures.region_stack: a real chart,
    not a stack of rectangles, so the shares can be read back and edited."""
    years, series = meta["years"], meta["series"]
    if not years or not series:
        return
    names = meta.get("snames") or [s["name"] for s in series]
    return _stacked_share_chart(slide, rect, years, series, names, True)

def _render_mekko(slide, rect, meta):
    """Native Marimekko: columns (width ∝ size) stacked by segment (height ∝
    share), coloured legend below."""
    x, y, w, h = rect
    cols, rows = meta["cols"], meta["rows"]
    widths, col_tot = meta["widths"], meta["col_tot"]
    heights, colors = meta["heights"], meta["colors"]
    lab_h, leg_h = 24.0, 20.0
    plot_h = max(30.0, h - lab_h - leg_h)
    accx = 0.0
    for ci, c in enumerate(cols):
        cw = widths[ci] / 100.0 * w
        cx = x + accx
        accy = 0.0
        for ri in range(len(rows)):
            frac = heights[ri][ci] / 100.0
            sh = frac * plot_h
            if sh <= 0:
                continue
            yy = y + plot_h - accy - sh
            col = _rgb(colors[ri])
            shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(cx), E(yy),
                                         E(max(1.0, cw)), E(sh))
            shp.fill.solid()
            shp.fill.fore_color.rgb = col
            shp.line.color.rgb = _C_WHITE
            shp.line.width = Pt(0.75)
            shp.shadow.inherit = False
            if sh >= 15 and cw >= 30:
                tf = shp.text_frame
                tf.word_wrap = False
                for m in ("margin_left", "margin_right", "margin_top",
                          "margin_bottom"):
                    setattr(tf, m, Emu(0))
                _set_text(tf, [[(f"{heights[ri][ci]:.0f}%", 7,
                                 False, _readable_on(col))]],
                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            accy += sh
        tb = _textbox(slide, cx, y + plot_h + 2, cw, lab_h)
        _set_text(tb.text_frame, [[(c, 7, True, _C_TEXT)],
                                  [(_abbr(col_tot[ci]), 7, False, _C_MUTED)]],
                  align=PP_ALIGN.CENTER)
        accx += cw
    lx = x
    for ri, r in enumerate(rows):
        sw = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(lx),
                                    E(y + plot_h + lab_h + 4), E(10), E(10))
        sw.fill.solid()
        sw.fill.fore_color.rgb = _rgb(colors[ri])
        sw.line.fill.background()
        sw.shadow.inherit = False
        tw = 8 + len(r) * 4.6
        tb = _textbox(slide, lx + 13, y + plot_h + lab_h + 1, tw, leg_h)
        _set_text(tb.text_frame, [[(r, 7.5, False, _C_TEXT)]],
                  anchor=MSO_ANCHOR.MIDDLE)
        lx += 13 + tw + 10

def _render_gauge(slide, rect, t, fig):
    """No native gauge in PPT -> editable KPI-style deviation card.  Reads the
    value AND its unit suffix off the plotly trace, so a gauge built in basis
    points (BPS) shows BPS here too."""
    x, y, w, h = rect
    val = float(t.value)
    col = _C_GOOD if val >= 0 else _C_BAD
    suffix = ""
    try:
        suffix = (t.number.suffix or "").strip()
    except Exception:
        pass
    suffix = suffix or "BPS"
    valstr = (f"{val:+,.0f} {suffix}" if suffix.lower() == "bps"
              else f"{val:+.2f} {suffix}")
    tb = _textbox(slide, x, y + h * 0.18, w, h * 0.6)
    subtitle = ""
    if fig.layout.title and fig.layout.title.text:
        subtitle = str(fig.layout.title.text)
    # the headline figure keeps the GOOD/BAD chart tokens; the ▲/▼ line under
    # it is a classification, so it takes the classification colours
    mark_col = _rgb(OUTPACE_COLOR if val >= 0 else DECLINE_COLOR)
    _set_text(tb.text_frame, [
        [(valstr, 24, True, col)],
        [(subtitle or "vs regional benchmark", 9, False, _C_MUTED)],
        [("▲ above benchmark" if val >= 0 else "▼ below benchmark",
          10, True, mark_col)],
    ], align=PP_ALIGN.CENTER)

def _annotation_note(slide, spec, fig):
    """Keep reference-line labels (e.g. 'Global: 12.2%') as small text."""
    try:
        notes = [a.text for a in fig.layout.annotations or []
                 if a.text and "%" in str(a.text)]
    except Exception:
        notes = []
    if notes:
        tb = _textbox(slide, spec["x"] + spec["w"] - 170,
                      spec["y"] + 3, 162, 16)
        _set_text(tb.text_frame,
                  [[(notes[0], 8, True, _rgb(REF_LINE))]],
                  align=PP_ALIGN.RIGHT)

# ── other spec renderers ─────────────────────────────────────────────
def _render_kpi(slide, spec):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, E(spec["x"]), E(spec["y"]),
        E(spec["w"]), E(spec["h"]))
    shp.adjustments[0] = 0.08
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(KPI_TINT)
    shp.line.color.rgb = _C_EDGE
    shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    big = 20 if spec["h"] > 80 else 15
    # label and caption at the 11pt content standard; the figure keeps its
    # display size.  Tight insets and spacing keep all three lines inside a
    # card as short as the 56px KPI band.
    paras = [
        [(spec["label"].upper(), MIN_CONTENT_PT, True, _C_MUTED)],
        [(spec["value"], big, True, _C_ACCENT)],
    ]
    if spec.get("sub"):
        paras.append([(spec["sub"], MIN_CONTENT_PT, False, _C_MUTED)])
    _set_text(shp.text_frame, paras, align=PP_ALIGN.CENTER,
              anchor=MSO_ANCHOR.MIDDLE, pad=(6, 6, 1, 1), line=0.95)

def _render_writeup(slide, spec):
    shp = _white_panel(slide, spec["x"], spec["y"], spec["w"], spec["h"])
    # 11pt content standard: a write-up no longer shrinks to fit its strip
    # (the layout sizes the strip for 11pt instead - see _block_need)
    fs, _sp, ln = _fit_text_fs([spec["text"]], spec["w"] - 26, spec["h"] - 12,
                               sizes=(MIN_CONTENT_PT,))
    _set_text(shp.text_frame,
              [[(t, fs, False, _rgb(c))
                for t, c in components.mark_runs(spec["text"])]],
              anchor=MSO_ANCHOR.MIDDLE, pad=(12, 12, 6, 6),
              line=max(1.12, ln))

def _render_text_panel(slide, spec):
    rect = _panel_frame(slide, spec)
    x, y, w, h = rect
    tb = _textbox(slide, x, y, w, h)
    lines = spec["lines"]
    # size the font so every line fits the panel height at a readable size -
    # avoids the cramped/auto-shrunk look when a tree has many rows
    n = max(1, len(lines))
    line_px = (h - 6) / n
    child_fs = max(6.0, min(10.0, line_px / 1.5))
    parent_fs = child_fs + 1.5
    paras = []
    for text, level in lines:
        if level == 0:
            paras.append([("▸  " + text, parent_fs, True, _C_ACCENT)])
        else:
            paras.append([("      ☐  " + text, child_fs, False,
                           _C_HEADER)])
    _set_text(tb.text_frame, paras)

# ── tables: ONE metric model, to draw a table AND to size its split ──
# Every cell - header, body, total - is set at the 11pt content standard.
# Column widths come from the cells' own text measured in the deck face, so
# a figure never runs out of its column; row heights follow the model the
# padlock pass (pptlock) re-measures a locked table with - line box x 1.2
# lead + the cell margins, x 1.15 safety - so a table that is split to fit
# its slides fits in the open deck and in the locked one alike.
_TBL_PAD_L, _TBL_PAD_R, _TBL_PAD_V = 5.0, 4.0, 2.0     # cell margins, px
_TBL_LEAD = 1.2                   # pptlock.LINE_LEAD
_TBL_SAFETY = 1.15                # pptlock._ROW_SAFETY
_TBL_WRAP_FILL = 0.80             # pptlock._TABLE_FILL: a wrapping cell breaks early
_TBL_TEXT_K = 1.03                # PowerPoint's line vs the font file's measure
_TBL_WRAP_LEAD = 0.9              # line spacing of a wrapping label (open decks)
_TBL_ROW_GROW = 1.35              # the most a row grows to fill its panel
_TBL_CAPTION = 16.0               # room for the source caption under the rows
_TBL_BOTTOM = 6.0 + _TBL_CAPTION  # _panel_frame's bottom pad + the caption
_TBL_TITLE_TUCK = 3.0              # a table sits this much closer under its
                                   # panel title than a chart (the 12pt title
                                   # ends ~6px above the chart gap)

_FACE_FILES = {                   # regular, bold, italic, bold italic
    "calibri": ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
    "segoe ui": ("segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf",
                 "segoeuiz.ttf"),
}
_FACE_CACHE: dict = {}
# glyphs the deck face lacks, which PowerPoint takes from a symbol font
_SYMBOLS = frozenset("▲▼▶◀►◄●○■□▪▫◆◇✓✔✗✘★☆➤➜↑↓→←")

@functools.lru_cache(maxsize=200_000)
def _text_px(text, size_pt, bold=False, italic=False, face=None):
    """Width in slide px of ``text`` in the deck face at ``size_pt`` -
    measured with the font file when it is on this machine, else estimated
    with layout's width model."""
    px = float(size_pt) * 4.0 / 3.0
    s = str(text)
    if any(ch in _SYMBOLS for ch in s):
        # Calibri has no ▲ / ▼ / ▶ glyphs: PowerPoint sets them in a symbol
        # face ~0.95em wide, where the font file would measure its narrow
        # missing-glyph box - and a "▲ 19.1%" sized to that wrapped
        n_sym = sum(ch in _SYMBOLS for ch in s)
        plain = "".join(ch for ch in s if ch not in _SYMBOLS)
        return _text_px(plain, size_pt, bold, italic, face) + n_sym * px * 0.95
    files = _FACE_FILES.get(str(face or FONT).strip().lower())
    key = (files, bold, italic, round(px * 4))
    font = _FACE_CACHE.get(key)
    if font is None:
        font = False
        if files:
            try:
                from PIL import ImageFont
                path = Path(r"C:\Windows\Fonts") / files[bold + 2 * italic]
                if path.exists():
                    font = ImageFont.truetype(str(path), key[3])
            except Exception:
                font = False
        _FACE_CACHE[key] = font
    if font:
        return font.getlength(str(text)) / 4.0
    return layout.text_width(str(text), px) * (1.02 if bold else 0.95)

def _wrap_px(text, avail, size_pt, bold=False):
    """Lines ``text`` takes in ``avail`` px at ``size_pt``: greedy word wrap,
    a word longer than the line broken where PowerPoint breaks it."""
    total = 0
    space = _text_px(" ", size_pt, bold)
    for para in str(text).split("\n"):
        n, cur = 1, 0.0
        for word in para.split():
            for j, piece in enumerate(_pieces(word)):
                ww = _text_px(piece, size_pt, bold)
                gap = space if j == 0 else 0.0   # a dash break takes no space
                if cur <= 0:
                    cur = ww
                elif cur + gap + ww <= avail:
                    cur += gap + ww
                else:
                    n += 1
                    cur = ww
                while avail > 0 and cur > avail:
                    n += 1
                    cur -= avail
        total += n
    return max(1, total)

def _pieces(word):
    """A word's pieces at the points PowerPoint may break it without a
    space: after a dash ("615.18–627.60" -> "615.18–", "627.60").  Not after
    a slash - PowerPoint keeps "Adjustable/Detachable-" together."""
    return [p for p in _re.split(r"(?<=[–—\-])", str(word)) if p] or [word]

def _break_lines(text, avail, size_pt, bold=False):
    """``text`` cut into the lines a table cell ``avail`` px wide sets it
    on: greedy, at a space or after a dash, a piece wider than the line cut
    where it fills it.  Each line measures within ``avail`` with the table's
    width allowance, so none of them wraps again."""
    def fits(s):
        return _text_px(s, size_pt, bold) * _TBL_TEXT_K + 2.0 <= avail

    lines = []
    for para in str(text).split("\n"):
        cur = ""
        for word in para.split():
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

def _tbl_row_need(lines, spacing=1.0):
    """A row's height in px holding ``lines`` lines of 11pt at line
    ``spacing``, margins included (pptlock's measure at 1.0, before its
    safety factor)."""
    return (lines * MIN_CONTENT_PT * _TBL_LEAD * spacing * 4.0 / 3.0
            + 2 * _TBL_PAD_V)

def _table_rows(df, index_name, total_row):
    """The table's cells, row by row, as (text, bold) - header first, then
    the body, then the total row if there is one."""
    rows = [[(str(index_name or ""), True)]
            + [(_col_label(c), True) for c in df.columns]]
    for idx, row in zip(df.index, df.itertuples(index=False)):
        rows.append([(str(idx), True)]
                    + [((str(v[0]), True) if isinstance(v, tuple)
                        else (str(v), False)) for v in row])
    if total_row is not None:
        rows.append([("Total", True)] + [(str(v), True) for v in total_row])
    return rows

# Cell margins (left, right) a table may use, roomiest first: a table whose
# figures do not fit on one line at the first steps down until they do.
_TBL_PADS = ((5.0, 4.0), (4.0, 3.0), (3.0, 2.0), (2.0, 2.0), (2.0, 1.5))

def _table_layout(df, index_name, total_row, content_w):
    """How a table sits in ``content_w`` px at 11pt.

    Returns {"widths": px per column, "heights": px per row, "tight": the
    rows at their bare text height, "wrap": per cell, whether its text has
    to wrap, "pad_l"/"pad_r": the cell margins it is drawn with}.

    The rules, in order (client, 2026-09-11):
      * every figure and every label on ONE line if the table can hold it -
        the cell margins tighten before anything wraps;
      * a column is as wide as its content, the value columns sharing one
        width so a year table reads as an even grid; width the text does not
        need is shared out so the table spans its frame, edge to edge;
      * when something must wrap, headers wrap first, then the columns of
        words - which take whatever room is left, the figures keeping
        theirs; a figure is never broken.
    A row is its text height in an open deck; a locked deck adds pptlock's
    15%, because the padlock pass re-grows every row to that."""
    rows = _table_rows(df, index_name, total_row)
    fs = MIN_CONTENT_PT
    n = len(rows[0])
    body = rows[1:] or rows
    W = float(content_w)

    def tw(t, b):
        # PowerPoint sets a line a hair wider than the font file measures
        # it, and breaks a cell whose text fills it exactly - so every width
        # carries a small allowance (a figure sized to the pixel wrapped)
        return _text_px(t, fs, b) * _TBL_TEXT_K + 2.0

    def plan(pads):
        def low_of(t, b):
            # the longest piece a line can break down to - PowerPoint breaks
            # at a space and after a dash or a slash ("615.18–" / "627.60")
            return max((tw(p, b) for wd in t.split() for p in _pieces(wd)),
                       default=0.0) + pads

        # Three widths per column:
        #   nat   - every cell on one line, header included
        #   want  - every BODY cell on one line; the header may wrap
        #   floor - the least it can take: a column of FIGURES keeps its
        #           widest figure whole; a column of words its longest piece
        nat, want, floor = [], [], []
        for c in range(n):
            cells = [r[c] for r in body if c < len(r)]
            hdr_t, hdr_b = rows[0][c]
            body_nat = max((tw(t, b) for t, b in cells), default=0.0) + pads
            head_low = low_of(hdr_t, hdr_b)
            nat.append(max(body_nat, tw(hdr_t, hdr_b) + pads))
            want.append(max(body_nat, head_low))
            if c > 0 and all(len(t.strip()) <= 12 for t, _b in cells):
                floor.append(want[-1])
            else:
                floor.append(max([low_of(t, b) for t, b in cells]
                                 + [head_low]))
        return nat, want, floor

    for pad_l, pad_r in _TBL_PADS:
        nat, want, floor = plan(pad_l + pad_r)
        if sum(want) <= W:
            break
    pads = pad_l + pad_r
    if sum(nat) <= W:
        # everything on one line - each column at its content width, no
        # wider (the value columns evened out where that still fits)
        widths = list(nat)
        if n > 1:
            uni = max(nat[1:])
            if nat[0] + uni * (n - 1) <= W:
                widths = [nat[0]] + [uni] * (n - 1)
    elif sum(want) <= W:
        # every body cell on one line; the headers wrap, sharing what is left
        slack = [max(0.0, a - b) for a, b in zip(nat, want)]
        room, tot = W - sum(want), (sum(slack) or 1.0)
        widths = [b + room * s / tot for b, s in zip(want, slack)]
    elif sum(floor) <= W:
        # some body cells must wrap too: the room above each column's floor
        # goes to the columns in proportion to how much more they want
        slack = [max(0.0, a - b) for a, b in zip(want, floor)]
        room, tot = W - sum(floor), (sum(slack) or 1.0)
        widths = [b + room * s / tot for b, s in zip(floor, slack)]
    else:
        # even the floors do not fit: the figures keep theirs - a number is
        # never broken - and the columns of words share what is left; only
        # when that would leave them next to nothing does everything shrink
        rigid = [c > 0 and all(len(r[c][0].strip()) <= 12
                               for r in body if c < len(r)) for c in range(n)]
        fixed = sum(f for f, rg in zip(floor, rigid) if rg)
        words = sum(f for f, rg in zip(floor, rigid) if not rg)
        if words > 0 and W - fixed >= 0.5 * words:
            k = (W - fixed) / words
            widths = [f if rg else f * k for f, rg in zip(floor, rigid)]
        else:
            k = W / sum(floor)           # last resort: words break
            widths = [b * k for b in floor]
    # Width the text does not need is not left as a white margin beside the
    # table: the columns share it in proportion, so the table spans its
    # frame and a year grid stays an even grid.
    if 0 < sum(widths) < W:
        k = W / sum(widths)
        widths = [v * k for v in widths]
    # An open deck is drawn as PowerPoint breaks it: at the full width.
    # A locked deck is re-measured by pptlock, which rewrites every row to
    # its own estimate plus 15%.  It measures a WRAPPING cell at 80% of its
    # width in a face ~8% wider than Calibri - a row label it counted at
    # four or five lines where PowerPoint sets three, the rows grew and the
    # table ran over its source line.  So in a locked deck a row label that
    # has to wrap is written as its own lines, each set not to wrap: pptlock
    # measures such a cell line by line in its real face at the full width,
    # which makes the row it writes the row PowerPoint draws.  (Only the
    # label column: pptlock reads it for nothing but the "Total" row, while
    # a header or a figure's text decides what it locks.)  Any other cell
    # that wraps is sized to the taller of pptlock's count and PowerPoint's.
    locked = bool(_FIT_TO_SLIDE.get("locked"))
    safety = _TBL_SAFETY if locked else 1.0
    ppl = _pptlock_mod() if locked else None
    heights, tight, wrap, cuts = [], [], [], {}
    for ri, row in enumerate(rows):
        need, flags = _tbl_row_need(1), []
        for c, (t, b) in enumerate(row[:n]):
            avail = max(16.0, widths[c] - pads)
            one = tw(t, b) <= avail
            if not t.strip():
                flags.append(False)
                continue
            if one:
                lines = 1
            elif locked and c == 0:
                cuts[(ri, c)] = _break_lines(t, avail, fs, b)
                lines = len(cuts[(ri, c)])
            elif locked:
                lines = _wrap_px(t, avail * 0.97, fs, b)
                if ppl is not None:
                    lines = max(lines, len(ppl._wrap_lines(
                        t, avail * ppl._TABLE_FILL, fs * ppl.PT_TO_PX)))
                else:
                    lines = max(lines, _wrap_px(t, avail * _TBL_WRAP_FILL,
                                                fs, b))
            else:
                lines = _wrap_px(t, avail * 0.97, fs, b)
            flags.append(not one and (ri, c) not in cuts)
            need = max(need, _tbl_row_need(
                lines, 1.0 if (locked or one) else _TBL_WRAP_LEAD))
        heights.append(need * safety + 1.0)
        tight.append(need + 1.0)
        wrap.append(flags)
    return {"widths": widths, "heights": heights, "tight": tight,
            "wrap": wrap, "cuts": cuts, "pad_l": pad_l, "pad_r": pad_r}

def _render_table(slide, spec):
    rect = _panel_frame(slide, spec)
    toc.capture("table", spec.get("title"))       # for the List of Tables
    x, y, w, h = rect
    # The table FILLS its frame (client, 2026-09-11): edge to edge between
    # the frame's hairlines, tucked up under the title, its rows reaching
    # down to the source line - no white margin on any side.  (The frame's
    # rounded corners are clear of it: the title sits above, the source
    # line below.)
    x, w = spec["x"] + 1, spec["w"] - 2
    if spec.get("title"):
        y, h = y - _TBL_TITLE_TUCK, h + _TBL_TITLE_TUCK
    df: pd.DataFrame = spec["df"]
    tot = spec.get("total_row")
    lay = _table_layout(df, spec.get("index_name", ""), tot, w)
    widths, wraps = lay["widths"], lay["wrap"]
    heights, tight = lay["heights"], lay["tight"]
    avail = max(1.0, h - _TBL_CAPTION)
    if sum(heights) <= avail:
        # a frame taller than the table: the body rows share the spare
        # height evenly (the header keeps its own), so the table reaches
        # the source line instead of leaving a white band above it
        hs = list(heights)
        if len(hs) > 1:
            extra = (avail - sum(hs)) / (len(hs) - 1)
            hs = [hs[0]] + [v + extra for v in hs[1:]]
    elif sum(tight) < avail:
        # a few rows too many for the panel: every row keeps its text and
        # gives up part of its breathing room, so the table stays whole
        share = (avail - sum(tight)) / (sum(heights) - sum(tight))
        hs = [t + (v - t) * share for t, v in zip(tight, heights)]
    else:
        hs = tight
    rows = _table_rows(df, spec.get("index_name", ""), tot)
    n_rows, n_cols = len(rows), len(widths)
    # as wide as its frame: the columns share its full width
    gf = slide.shapes.add_table(n_rows, n_cols, E(x), E(y),
                                E(sum(widths)), E(sum(hs)))
    tbl = gf.table
    tbl.first_row = False
    tbl.horz_banding = False
    for c in range(n_cols):
        tbl.columns[c].width = E(widths[c])
    pad_l, pad_r = lay["pad_l"], lay["pad_r"]
    cuts = lay.get("cuts") or {}

    def cell_text(cell, text, bold, color, fill, wrap, lines=None):
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
        cell.margin_left = E(pad_l)
        cell.margin_right = E(pad_r)
        cell.margin_top = E(_TBL_PAD_V)
        cell.margin_bottom = E(_TBL_PAD_V)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = bool(wrap)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        # An EMPTY cell still sets its line's height: PowerPoint ignores an
        # empty run's formatting, falls back to its 18pt default and grows
        # the row - a table with blank label cells ran off its panel.  The
        # paragraph's own end-of-line size pins every cell to 11pt.
        p._p.get_or_add_endParaRPr().set("sz",
                                         str(int(MIN_CONTENT_PT * 100)))
        if str(text) == "":
            return
        if wrap and not _FIT_TO_SLIDE.get("locked"):
            # a label that wraps sits on slightly tighter lines (the row
            # height model assumes the same) - open decks only, since the
            # padlock pass re-measures a locked row at full spacing
            p.line_spacing = _TBL_WRAP_LEAD
        # a label cut into lines (locked decks, see _table_layout) is one
        # paragraph per line, none of them wrapping
        for i, part in enumerate(lines or [str(text)]):
            if i:
                p = tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p._p.get_or_add_endParaRPr().set(
                    "sz", str(int(MIN_CONTENT_PT * 100)))
            r = p.add_run()
            r.text = part
            r.font.size = Pt(MIN_CONTENT_PT)
            r.font.bold = bold
            r.font.color.rgb = color
            r.font.name = FONT

    head_bg, head_ink = _rgb(TABLE_HEAD_BG), _rgb(TABLE_HEAD_TEXT)
    for c, (t, _b) in enumerate(rows[0]):
        cell_text(tbl.cell(0, c), t, True, head_ink, head_bg, wraps[0][c],
                  cuts.get((0, c)))
    for r, (idx, row) in enumerate(df.iterrows(), start=1):
        zebra = _C_ZEBRA if r % 2 == 0 else _C_WHITE
        cell_text(tbl.cell(r, 0), idx, True, _C_TEXT, zebra, wraps[r][0],
                  cuts.get((r, 0)))
        for c, v in enumerate(row, start=1):
            if isinstance(v, tuple):
                text, color = v
                cell_text(tbl.cell(r, c), text, True, _rgb(color), zebra,
                          wraps[r][c])
            else:
                cell_text(tbl.cell(r, c), v, False, _C_TEXT, zebra,
                          wraps[r][c])
    if tot is not None:
        r = n_rows - 1
        fill = _rgb(TABLE_TOTAL_BG)
        cell_text(tbl.cell(r, 0), "Total", True, _C_TEXT, fill, wraps[r][0],
                  cuts.get((r, 0)))
        for c, v in enumerate(tot, start=1):
            cell_text(tbl.cell(r, c), v, True, _C_TEXT, fill, wraps[r][c])
    for i, row in enumerate(tbl.rows):
        row.height = E(hs[i])
    # 'Source: ...' caption pinned to the panel bottom (the table's height
    # budget leaves _TBL_CAPTION px for it)
    # the source line sits right under the table (never below its panel)
    src = _textbox(slide, spec["x"] + 8,
                   min(spec["y"] + spec["h"] - 21, y + sum(hs) + 1),
                   spec["w"] - 16, 18)
    _set_text(src.text_frame,
              [[(components.SOURCE_TEXT, MIN_CONTENT_PT, False, _C_MUTED)]],
              anchor=MSO_ANCHOR.BOTTOM, pad=(0, 0, 0, 0))
    for p in src.text_frame.paragraphs:
        for run in p.runs:
            run.font.italic = True

# Every bar / stacked column in the deck uses ONE gap width, so a column on
# the Key Segments page is exactly as wide as a column on the Value & Share
# Shift page.  gapWidth is a percentage of the BAR width, so 50 means each bar
# fills 1 / (1 + 50/100) = two thirds of its category slot.
BAR_GAP_WIDTH = 50

def _apply_bar_gap(chart, n_cats=None, base=None):
    """Set PowerPoint's gapWidth to the deck-wide BAR_GAP_WIDTH.

    The gap used to be derived from the category count (the same rule as
    figures.cap_bar_widths), which made a two-category chart draw thin bars and
    a twelve-category chart draw fat ones.  One fixed gap keeps every chart in
    the deck consistent; ``n_cats`` / ``base`` are kept so the existing call
    sites stay valid.
    """
    try:
        chart.plots[0].gap_width = BAR_GAP_WIDTH
    except Exception:
        pass

def _fit_text_fs(chunks, text_w, avail_h,
                 sizes=(11.5, 11, 10.5, 10, 9.5, 9, 8.5, 8, 7.5)):
    """Largest font size (from ``sizes``) at which the wrapped ``chunks`` (list
    of paragraph strings) fit inside ``avail_h`` px within a ``text_w`` px column
    - so bullet/writeup text auto-shrinks to its card instead of overflowing.
    Returns (font_size, space_after_pt, line_multiplier).

    ``sizes`` are POINTS but ``text_w``/``avail_h`` are slide px, and 1 px =
    0.75 pt - so the size has to be converted before it is measured against a
    px width.  ``layout._wrap_count`` is a CSS-px model, hence the
    ``_PPT_FONT_K`` factor (the same one templates_dynamics budgets with).
    Measuring the width in raw points under-counted the wrapped lines, which is
    how bullet text could still overflow a card the auto-shrink had "fitted"."""
    for fs in sizes:
        line = 1.1 if fs >= 10 else 1.05
        sp = 5.0 if fs >= 10 else 3.0
        line_px = fs * line * 1.333 * _PPT_LINE_NAT
        sp_px = sp * 1.333
        total = 0.0
        for c in chunks:
            lines = max(1, layout._wrap_count(str(c), text_w,
                                              fs * _PPT_FONT_K))
            total += lines * line_px + sp_px
        if total <= avail_h - 2:
            return fs, sp, line
    return sizes[-1], 3.0, 1.0

def _render_summary_panel(slide, spec):
    """Editorial summary card -> native white panel + accent rule + bullet
    text (all editable).  Bullet font AUTO-SHRINKS to fit the card height so the
    text never spills past the panel."""
    x, y, w, h = spec["x"], spec["y"], spec["w"], spec["h"]
    _white_panel(slide, x, y, w, h)
    px, pw = x + 14, w - 28                        # consistent inner padding
    tb = _textbox(slide, px, y + 9, pw, 22)
    _set_text(tb.text_frame, [[(spec["title"], 13, True, _C_HEADER)]],
              pad=(0, 2, 1, 1))
    rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(px), E(y + 34),
                                  E(pw), E(2))
    rule.fill.solid()
    rule.fill.fore_color.rgb = _C_ACCENT
    rule.line.fill.background()
    rule.shadow.inherit = False
    body_top, body_h = y + 43, h - 51
    bullets = spec["bullets"] or [""]
    # 11.5pt where the card has room, never below the 11pt content standard
    fs, sp, ln = _fit_text_fs(bullets, pw - 16, body_h,
                              sizes=(11.5, MIN_CONTENT_PT))
    body = _textbox(slide, px, body_top, pw, body_h)
    # a bullet is running text, so its ▲/▼ (and the "green"/"red" the sentence
    # names them with) are split out and coloured - components.mark_runs is the
    # same helper the dashboard's bullets go through
    paras = [[("▪  ", min(10.0, fs * 0.9), False, _C_ACCENT),
              *((t, fs, False, _rgb(c)) for t, c in components.mark_runs(b))]
             for b in bullets]
    _set_text(body.text_frame, paras, pad=(0, 2, 1, 1), space_after=sp, line=ln)

_RENDERERS = {
    "kpi": _render_kpi,
    "writeup": _render_writeup,
    "table": _render_table,
    "figure": _fig_to_native,
    "text": _render_text_panel,
    "sumpanel": _render_summary_panel,
}

# ── slide assembly ───────────────────────────────────────────────────
def _is_round(shape) -> bool:
    """True for an ellipse/circle - its text has less room than its box."""
    try:
        return 'prst="ellipse"' in shape._element.xml
    except Exception:
        return False


def _tidy_slide(slide, margin=6.0):
    """Last pass over a finished slide: nothing sticking out, nothing clipped.

    The layout engine sizes every BLOCK, but a renderer can still place a
    label from a measurement that turns out short - a long geography name in
    a narrow gutter, a value tag on the last category of a chart.  Two rules,
    applied to small text shapes only, so a deliberate full-bleed graphic (the
    cover's rings) is never touched:

      * a label that runs past the right or bottom edge is pulled back inside
      * a label whose text cannot fit its box on one line is set smaller,
        down to a floor, rather than being clipped by PowerPoint

    Runs from `_footer`, which every content slide calls once its content is
    drawn, so it costs one walk per slide and applies to both templates.
    """
    for sh in slide.shapes:
        if sh.shape_type == 13 or not sh.has_text_frame:
            continue
        if sh.width is None or sh.height is None:
            continue
        w, h = sh.width / _EMU_PER_PX, sh.height / _EMU_PER_PX
        if w > 420 or h > 200:              # a panel, not a label
            continue
        txt = sh.text_frame.text.strip()
        if not txt:
            continue

        # ── a BADGE: a short label centred in a small shape ──────────
        # "10" in a 34px circle came out stacked as "1" over "0".  Nothing
        # was oversized: PowerPoint's default side insets are 7.2pt EACH, so
        # a 25.5pt-wide circle offered ~11pt of usable width to a 16pt
        # string, and word-wrap did the rest.  A number is never allowed to
        # break, so the insets go, wrapping goes, and only then - if the
        # glyphs still do not fit the circle - the size comes down.
        if len(txt) <= 4 and " " not in txt and w <= 72 and h <= 72:
            tf = sh.text_frame
            tf.word_wrap = False
            tf.margin_left = tf.margin_right = 0
            tf.margin_top = tf.margin_bottom = 0
            size = next((r.font.size.pt for p in tf.paragraphs
                         for r in p.runs if r.font.size), None)
            if size:
                # a circle's usable width is its INSCRIBED square, ~0.72 of
                # the diameter, so a 2-digit badge is sized against that
                inner = w * (0.72 if _is_round(sh) else 0.94)
                need = layout.text_width(txt, size / 0.75)
                if need > inner:
                    for p in tf.paragraphs:
                        for r in p.runs:
                            # never below the 11pt content standard
                            r.font.size = Pt(max(MIN_CONTENT_PT,
                                                 size * inner / need))
            continue

        x, y = sh.left / _EMU_PER_PX, sh.top / _EMU_PER_PX
        if x + w > 1280 - margin:
            sh.left = E(max(margin, 1280 - margin - w))
        if y + h > 720 - margin:
            sh.top = E(max(margin, 720 - margin - h))
        # one line that cannot fit: shrink it rather than let it be cut
        # A frame that WRAPS cannot be cut off - the words move to the next
        # line, and whatever placed the shape sized it for that.  Shrinking
        # those here gave every long label a size of its own, computed from
        # its own width: on the taxonomy page 'Downstream Pathway Modulation'
        # came out at 7.33pt and 'Autoimmune & Inflammatory Diseases' at 6.5,
        # beside 'Oncology' at the 11 the layout asked for.  Only a label that
        # CANNOT wrap - one word, or wrapping switched off - risks the clip
        # this pass exists to prevent.
        _wraps = bool(getattr(sh.text_frame, "word_wrap", False)) and " " in txt
        if "\n" not in txt and len(txt) > 3 and not _wraps:
            size = next((r.font.size.pt for p in sh.text_frame.paragraphs
                         for r in p.runs if r.font.size), None)
            if size:
                # everything in this pass is in canvas PX; `size` is in pt, so
                # the measurement is taken at the px font size (pt / 0.75) and
                # compared against the box's px width - mixing the two units
                # is what let a 97pt label sit unshrunk in a 112px box.
                need = layout.text_width(txt, size / 0.75)
                if need > w - 3:
                    new = max(MIN_CONTENT_PT, size * (w - 3) / need)
                    if new < size - 0.2:
                        for p in sh.text_frame.paragraphs:
                            for r in p.runs:
                                r.font.size = Pt(new)
                    # still wider than its box at the 11pt content standard:
                    # the box grows (about its own alignment) rather than
                    # the label dropping below the standard
                    bold = bool(next((r.font.bold
                                      for p in sh.text_frame.paragraphs
                                      for r in p.runs), False))
                    need = _text_px(txt, max(new, MIN_CONTENT_PT), bold)
                    if need > w - 8:
                        _widen_label(sh, need + 10, margin)


def _is_dark_slide(slide) -> bool:
    """True for the cover and the section dividers - full-bleed dark art whose
    shapes are MEANT to bleed off the canvas, so the tidy pass leaves them be."""
    for sh in slide.shapes:
        if sh.shape_type != 1 or sh.width is None:
            continue
        if (sh.width / _EMU_PER_PX > 1150 and sh.height / _EMU_PER_PX > 640):
            try:
                if sh.fill.type == 1:
                    c = str(sh.fill.fore_color.rgb)
                    return sum(int(c[i:i + 2], 16) for i in (0, 2, 4)) < 380
            except Exception:
                return False
    return False


def _install_theme(prs):
    """Put the template's OWN colour and font scheme into the deck's theme part.

    Every colour this project draws is explicit, so the theme part is normally
    invisible - but it is what PowerPoint hands to anything added AFTER the
    file is opened: a new shape, a pasted chart, SmartArt.  Left alone, a deck
    built by python-pptx carries Microsoft's "Office" scheme (accent1 4F81BD,
    accent2 C0504D ...), so a client's own edits come out in Office brights
    beside the template's muted accents.

    A theme with no colours declared (`PPT_THEME_COLORS is None`) means "leave
    the package's theme as it is", which is what Fact.MR wants.
    """
    if not PPT_THEME_COLORS:
        return
    from pptx.oxml.ns import qn
    for master in prs.slide_masters:
        try:
            theme = master.part.part_related_by(
                "http://schemas.openxmlformats.org/officeDocument/2006/"
                "relationships/theme")
            root = theme._element if hasattr(theme, "_element") else None
            if root is None:
                from lxml import etree
                root = etree.fromstring(theme.blob)
            scheme = root.find(qn("a:themeElements"))
            if scheme is None:
                continue
            clr = scheme.find(qn("a:clrScheme"))
            if clr is not None:
                if PPT_THEME_NAME:
                    clr.set("name", PPT_THEME_NAME)
                for slot, hexval in PPT_THEME_COLORS.items():
                    node = clr.find(qn(f"a:{slot}"))
                    if node is None:
                        continue
                    for child in list(node):
                        node.remove(child)
                    srgb = node.makeelement(qn("a:srgbClr"), {"val": hexval})
                    node.append(srgb)
            fonts = scheme.find(qn("a:fontScheme"))
            if fonts is not None and PPT_THEME_FONTS:
                for tag, face in zip(("a:majorFont", "a:minorFont"),
                                     PPT_THEME_FONTS):
                    grp = fonts.find(qn(tag))
                    if grp is None:
                        continue
                    latin = grp.find(qn("a:latin"))
                    if latin is not None:
                        latin.set("typeface", face)
            if not hasattr(theme, "_element"):
                from lxml import etree
                theme._blob = etree.tostring(root, xml_declaration=True,
                                             encoding="UTF-8",
                                             standalone=True)
        except (KeyError, AttributeError):
            continue


def _tidy_deck(prs):
    """Run `_tidy_slide` over every content slide, once the deck is complete.

    Deliberately NOT hooked into `_footer`: several full-page renderers draw
    their labels AFTER the footer, so tidying there would run before the
    labels existed and miss exactly the cases it is meant to catch.
    """
    for slide in prs.slides:
        if not _is_dark_slide(slide):
            _tidy_slide(slide)


def _widen_label(sh, need_w, margin=6.0):
    """Grow a one-line label's box to ``need_w`` px about its own alignment
    (left edge, centre or right edge held), staying on the slide."""
    w = sh.width / _EMU_PER_PX
    if need_w <= w:
        return
    x = sh.left / _EMU_PER_PX
    al = next((p.alignment for p in sh.text_frame.paragraphs
               if p.alignment is not None), None)
    grow = need_w - w
    if al == PP_ALIGN.CENTER:
        x -= grow / 2
    elif al == PP_ALIGN.RIGHT:
        x -= grow
    new_w = min(need_w, 1280 - 2 * margin)
    x = max(margin, min(x, 1280 - margin - new_w))
    sh.left, sh.width = E(x), E(new_w)


def _enforce_content_floor(prs) -> int:
    """The last word on the 11pt content standard.

    Every renderer sets its content at MIN_CONTENT_PT or more, but a deck is
    assembled from dozens of them; this pass guarantees the rule for the
    whole file.  Any run under 11pt in a slide's CONTENT area - text boxes,
    shapes, table cells, chart text - is raised to it.  The slide chrome is
    left alone: the header zone above the rule (title, sub header, chips,
    buttons) and the copyright strip under the content box.

    Returns how many runs had to be raised - 0 when every renderer already
    complied.  Set FMI_FLOOR_REPORT to a file path to have each one listed,
    which is how a renderer still setting small type is found."""
    import os
    floor = int(round(MIN_CONTENT_PT * 100))      # OOXML sz: 1/100 pt
    top, bottom = RULE_Y * _EMU_PER_PX, layout.CONTENT_BOTTOM * _EMU_PER_PX
    tags = (qn("a:rPr"), qn("a:defRPr"), qn("a:endParaRPr"))
    report = [] if os.environ.get("FMI_FLOOR_REPORT") else None
    raised = 0

    def shapes(coll):
        for sh in coll:
            yield sh
            if sh.shape_type == 6:                  # a group
                yield from shapes(sh.shapes)

    def lift(root, where):
        nonlocal raised
        for node in root.iter(*tags):
            sz = node.get("sz")
            if sz is not None and int(sz) < floor:
                if report is not None:
                    run = node.getparent()
                    txt = "".join(run.itertext()) if run is not None else ""
                    report.append(f"{where}\t{int(sz) / 100:g}pt\t{txt[:60]}")
                node.set("sz", str(floor))
                raised += 1

    for i, slide in enumerate(prs.slides, 1):
        for sh in shapes(slide.shapes):
            if sh.top is None or not top <= sh.top < bottom:
                continue                            # header / footer chrome
            if sh.name == _TOC_SHAPE:
                # The contents pages are a transcript of the research team's
                # Word file, pasted the way they paste it themselves - 10pt
                # for a heading, 9pt for the levels under it, 7pt for the
                # tab.  Raising those to 11pt would re-flow the file into
                # something that is no longer what the analyst wrote, and
                # thirty rows would no longer fit a table.
                continue
            lift(sh._element, f"slide {i}\t{sh.name}")
            if getattr(sh, "has_chart", False):
                lift(sh.chart._chartSpace, f"slide {i}\t{sh.name} (chart)")
    if report:
        with open(os.environ["FMI_FLOOR_REPORT"], "a", encoding="utf-8") as fh:
            fh.write("\n".join(report) + "\n")
    return raised


# Slides whose content is drawn at the fixed coordinates of the old 84px
# header: id(slide) -> (slide, index of its first content shape, remap?).
# Recorded by _base_slide, consumed by _footer once the content is drawn.
_CONTENT_MARK: dict = {}
_REMAP_T0 = 84.0          # where fixed-coordinate content may begin
_REMAP_B = 706.0          # ...and the footer line it must stay above

def _remap_content(slide, start):
    """Move a finished page's content under the taller header.

    The bespoke pages (dynamics, pricing, competitor, segment-outlook, the
    fixed-layout country pages ...) place their shapes at literal coordinates
    written for a header that ended at 84px.  Rather than re-cut each one,
    the page's content box is mapped linearly from [84, 706] onto
    [CONTENT_TOP - 3, 706]: every shape keeps its place relative to its
    neighbours and the page gives up ~2.5% of its height."""
    t1 = layout.CONTENT_TOP - 3.0
    k = (_REMAP_B - t1) / (_REMAP_B - _REMAP_T0)
    for sh in list(slide.shapes)[start:]:
        if sh.top is None or sh.height is None:
            continue
        y = sh.top / _EMU_PER_PX
        ny = (t1 + (y - _REMAP_T0) * k if y >= _REMAP_T0
              else y + (t1 - _REMAP_T0))
        sh.top = E(ny)
        sh.height = Emu(int(round(sh.height * k)))
        if getattr(sh, "has_table", False):
            for row in sh.table.rows:
                row.height = Emu(int(round(row.height * k)))

def _footer(slide, dark=False, logo=True):
    """The page's closing line, on the template's bar where it has one.

    Called once a slide's content is complete, so it is also where a page
    drawn at the old fixed coordinates is moved under the header
    (`_remap_content`).

    FMI used to close every page with a solid blue bar carrying the copyright
    in white.  The bar was removed on 2026-09-11 (client request): the
    copyright now sits on the white page in FOOTER_TEXT_COLOR.  Fact.MR
    (FOOTER_TEXT_COLOR None) keeps its muted line on the right.
    """
    mark = _CONTENT_MARK.pop(id(slide), None)
    if mark is not None and mark[2]:
        _remap_content(slide, mark[1])
    if FOOTER_BAND and not dark:
        band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(0),
                                      E(FOOTER_RULE_Y), E(1280),
                                      E(FOOTER_BAND_H))
        band.fill.solid()
        band.fill.fore_color.rgb = _rgb(FOOTER_BAND)
        band.line.fill.background()
        band.shadow.inherit = False
        tb = _textbox(slide, FOOTER_MARGIN, FOOTER_RULE_Y + 1,
                      1280 - 2 * FOOTER_MARGIN, FOOTER_BAND_H - 2)
        _set_text(tb.text_frame,
                  [[(_COPYRIGHT, 7, False, _rgb(FOOTER_TEXT_COLOR))]],
                  align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        return
    if FOOTER_TEXT_COLOR and not dark:
        tb = _textbox(slide, FOOTER_MARGIN, FOOTER_RULE_Y,
                      1280 - 2 * FOOTER_MARGIN, 720 - FOOTER_RULE_Y)
        # slide chrome, not content: it keeps its small size (floor=False)
        # so it fits the strip under the content box
        _set_text(tb.text_frame,
                  [[(_COPYRIGHT, 8, False, _rgb(FOOTER_TEXT_COLOR))]],
                  align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
                  pad=(3, 3, 0, 0), floor=False)
        return
    tb = _textbox(slide, 420, 707, 842, 13)
    col = _rgb(ON_DARK_MUTED) if dark else _C_MUTED
    _set_text(tb.text_frame, [[(_COPYRIGHT, 7, False, col)]],
              align=PP_ALIGN.RIGHT, floor=False)

# top-left logo geometry (also fixes where the header title starts).  The
# wordmark sits on a small WHITE CARD so it reads on any header fill.
# 2026-09-11: it was 28px (21pt) tall, too small to read as a brand mark.
# 46px (~34pt) is ~1.6x that - a little over half the FMI master's own 55pt
# placement, which is as large as the 96px header can hold beside a two-line
# 21pt title - and it is centred vertically in the header zone.
_LOGO_IMG_H = 46
_LOGO_CARD_X, _LOGO_CARD_Y = 24, (96 - _LOGO_IMG_H) // 2
_LOGO_CARD_W = _LOGO_IMG_H * _LOGO_AR + 14  # card width, used to place the title

def _top_left_logo(slide, x=None, y=None):
    """The template's wordmark on a rounded white card - top-left of the
    header, unless ``x``/``y`` place it elsewhere (the section dividers)."""
    if not _LOGO_PATH.exists():
        return
    x = _LOGO_CARD_X if x is None else x
    y = _LOGO_CARD_Y if y is None else y
    lw = _LOGO_IMG_H * _LOGO_AR
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, E(x - 7), E(y - 5),
        E(lw + 14), E(_LOGO_IMG_H + 10))
    card.adjustments[0] = 0.16
    card.fill.solid()
    card.fill.fore_color.rgb = _C_WHITE
    card.line.fill.background()
    card.shadow.inherit = False
    try:
        slide.shapes.add_picture(
            str(_LOGO_PATH), E(x), E(y),
            height=E(_LOGO_IMG_H), width=E(lw))
    except Exception:
        pass

def _nav_button(slide, x, w, label, size=11):
    btn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 E(x), E(24), E(w), E(44))
    btn.adjustments[0] = 0.22
    btn.fill.solid()
    btn.fill.fore_color.rgb = _rgb(NAV_BTN_BG)
    btn.line.color.rgb = _rgb(NAV_BTN_EDGE)
    btn.line.width = Pt(0.5)
    btn.shadow.inherit = False
    _set_text(btn.text_frame, [[(label, size, True, _rgb(NAV_BTN_TEXT))]],
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return btn

# The sub header drawn under the title of the slide being built.  Set by the
# page loop (see emit() in export_native_ppt) from subheaders.for_page, so
# every slide a page produces - including its numbered "(j/k)" spills -
# carries the same one-line description of what it shows.
_SUBHEADER = {"text": ""}
SUBHEADER_COLOR = TEXT           # body ink, italic - as in the FMI reference decks

_SUB_FONT_FILES = (r"C:\Windows\Fonts\calibrii.ttf",      # Calibri Italic
                   r"C:\Windows\Fonts\calibri.ttf")
_SUB_FONT_CACHE: dict = {}

def _sub_text_width(text, size_pt):
    """Rendered width in slide px of ``text`` at ``size_pt`` in the header's
    italic face.  Measured with the real font file when it is on this machine
    (Pillow), else estimated from Calibri's average advance (0.42 em).  The
    generic wrap model used for card text (_fit_text_fs) over-counts a single
    long line by ~15%, which is what made some sub headers shrink for no
    reason."""
    px = size_pt * 96.0 / 72.0
    try:
        from PIL import ImageFont
        key = round(px * 4)
        font = _SUB_FONT_CACHE.get(key)
        if font is None:
            for path in _SUB_FONT_FILES:
                if Path(path).exists():
                    font = ImageFont.truetype(path, key)   # 4x for precision
                    break
            _SUB_FONT_CACHE[key] = font or False
        if font:
            return font.getlength(text) / 4.0
    except Exception:
        pass
    return len(text) * 0.42 * px

def _subheader_size(text, avail_px, sizes=(14, 13.5, 13, 12.5, 12)):
    """Largest of ``sizes`` (pt) at which ``text`` stays on one line inside
    ``avail_px``."""
    for fs in sizes:
        if _sub_text_width(text, fs) <= avail_px:
            return fs
    return sizes[-1]

# A page spread over several slides numbers them in its heading - "(1/2)",
# "(2/2)".  On a long title that suffix is what pushes the heading onto a
# second line, so a numbered heading steps its size down until it fits one.
_PART_RE = _re.compile(r"\(\d+/\d+\)\s*$")
_TITLE_FONT_FILES = {"calibri": r"C:\Windows\Fonts\calibrib.ttf",
                     "segoe ui": r"C:\Windows\Fonts\segoeuib.ttf"}
_TITLE_FONT_CACHE: dict = {}

def _title_text_width(text, size_pt):
    """Width in slide px of ``text`` set bold at ``size_pt`` in the deck face
    (measured with the font file when it is on this machine)."""
    px = size_pt * 96.0 / 72.0
    try:
        from PIL import ImageFont
        path = _TITLE_FONT_FILES.get(str(FONT).strip().lower())
        key = (path, round(px * 4))
        font = _TITLE_FONT_CACHE.get(key)
        if font is None:
            font = (ImageFont.truetype(path, key[1])
                    if path and Path(path).exists() else False)
            _TITLE_FONT_CACHE[key] = font
        if font:
            return font.getlength(text) / 4.0
    except Exception:
        pass
    return len(text) * 0.55 * px

# Slide heading type, per the client's standard (2026-09-11; title 22 -> 21pt
# on 2026-09-15): the bold title is 21pt and the sub header under it 14pt,
# not bold.  The title may take two lines; the header zone above RULE_Y is
# sized for exactly that.
_HDR_TITLE_PT = 21.0
_HDR_SUB_PT = 14.0
_CHIP_W, _CHIP_STEP = 150, 158    # a slicer chip, and chip + gap


def _chips_left(n):
    """x of the first of `n` slicer chips, right-aligned to the margin."""
    return 1265 - _CHIP_W - (n - 1) * _CHIP_STEP


_HDR_TITLE_RIGHT_NAV = 1265       # the right margin (no nav buttons since
                                  # 2026-09-25), as the sub header
_HDR_SUB_Y = 68                   # the sub header's line, under the buttons

def _title_lines(text, avail_px, size_pt):
    """How many lines ``text`` wraps to, set bold at ``size_pt`` in a text
    frame ``avail_px`` wide - greedy word wrap, the way PowerPoint breaks it,
    measured with the deck's own font file."""
    space = _title_text_width(" ", size_pt)
    lines, cur = 1, 0.0
    for word in str(text).split():
        ww = _title_text_width(word, size_pt)
        if cur <= 0:
            cur = ww
        elif cur + space + ww <= avail_px:
            cur += space + ww
        else:
            lines += 1
            cur = ww
    return lines

def _header_title_fit(text, tx, chips):
    """(title size, chips to draw) so the heading holds to two lines at 21pt.

    Beside the slicer chips when it fits there; otherwise the chips give way
    - the heading already names the geography they repeat - and the title
    takes the width up to the right margin.  Only a heading too long
    for two lines even then steps down, half a point at a time."""
    inner = 8                                    # the frame's side insets
    beside = _chips_left(len(chips) or 1) - 10 - tx - inner
    full = _HDR_TITLE_RIGHT_NAV - tx - inner
    if chips and _title_lines(text, beside, _HDR_TITLE_PT) <= 2:
        return _HDR_TITLE_PT, chips
    fs = _HDR_TITLE_PT
    while fs > 14 and _title_lines(text, full, fs) > 2:
        fs -= 0.5
    return fs, ()

def _base_slide(prs, title, chip_vals=(), remap=True):
    """A page background + header zone + title + sub header + slicer chips.

    ``remap`` marks a page whose content is drawn at the fixed coordinates
    written for the old 84px header; `_footer` moves it under this header
    once it is drawn (see `_remap_content`).  Pages placed by the layout
    engine already start at layout.CONTENT_TOP and pass ``remap=False``."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = _C_PAGE
    bg.line.fill.background()
    bg.shadow.inherit = False
    # THE ONE PLACE the two templates' slide chrome differs (mirrors
    # components._band_style on the dashboard):
    #   Fact.MR "band" - the zone is filled with the brand colour, title white
    #   FMI     "rule" - the zone stays white, title in dark ink, and the
    #                    template's light-blue rule is drawn underneath
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(13), E(9),
                                  E(1254), E(layout.CONTENT_TOP - 16))
    band.fill.solid()
    band.fill.fore_color.rgb = _rgb(HEADER_FILL)
    band.line.fill.background()
    band.shadow.inherit = False
    if HEADER_STYLE != "band" and RULE_H:
        # The FMI template divides this rule into one coloured segment per
        # content column; a template with no segments draws it whole.  Same
        # geometry either way, so the two surfaces stay in step with
        # components._rule_segments().
        segs = (HEADER_RULE_SEGMENTS or
                ((0.0, 1.0, HEADER_RULE),) if HEADER_RULE else ())
        for start, width, colour in segs:
            rule = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, E(13 + round(start * 1254)), E(RULE_Y),
                E(round(width * 1254)), E(RULE_H))
            rule.fill.solid()
            rule.fill.fore_color.rgb = _rgb(colour)
            rule.line.fill.background()
            rule.shadow.inherit = False
    _top_left_logo(slide)                        # logo now top-left, not footer
    tx = int(_LOGO_CARD_X + _LOGO_CARD_W + 14)   # title starts after the logo
    sub = _SUBHEADER["text"]
    clean = _clean_title(title)
    tfs, chip_vals = _header_title_fit(clean, tx, tuple(chip_vals))
    right = (_chips_left(len(chip_vals)) - 10 if chip_vals
             else _HDR_TITLE_RIGHT_NAV)
    if sub:
        # The title (one or two lines, bottom-aligned) sits beside the chips
        # and buttons; the sub header runs the full width on its own line
        # under them, so a long sub header never has to shrink or wrap.
        tb = _textbox(slide, tx, 2, right - tx, _HDR_SUB_Y - 2)
        _set_text(tb.text_frame, [[(clean, tfs, True, _rgb(TITLE_COLOR))]],
                  anchor=MSO_ANCHOR.BOTTOM, pad=(3, 3, 0, 0), line=0.9)
        sb = _textbox(slide, tx, _HDR_SUB_Y, 1265 - tx, 26)
        sb.text_frame.word_wrap = False
        # Reference-deck style: italic, not bold, in body ink - 14pt, stepping
        # down only for a line too long for the full width.
        sfs = _subheader_size(sub, 1265 - tx - 8)
        _set_text(sb.text_frame, [[(sub, sfs, False, _rgb(SUBHEADER_COLOR))]],
                  anchor=MSO_ANCHOR.TOP, pad=(3, 3, 0, 0))
        for para in sb.text_frame.paragraphs:
            for run in para.runs:
                run.font.italic = True
    else:
        tb = _textbox(slide, tx, 8, right - tx, 80)
        _set_text(tb.text_frame, [[(clean, tfs, True, _rgb(TITLE_COLOR))]],
                  anchor=MSO_ANCHOR.MIDDLE, pad=(3, 3, 0, 0), line=0.9)
    chip_x = _chips_left(len(chip_vals))
    for v in chip_vals:
        chip = _nav_button(slide, chip_x, _CHIP_W, v, size=10)
        chip.fill.fore_color.rgb = _rgb(CHIP_FILL)
        chip.line.color.rgb = _rgb(CHIP_EDGE)
        for p in chip.text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = _rgb(CHIP_TEXT)
        chip_x += _CHIP_STEP
    _CONTENT_MARK[id(slide)] = (slide, len(slide.shapes), remap)
    return slide

def _draw_specs(slide, specs):
    for spec in specs:
        fn = _RENDERERS.get(spec["kind"])
        if fn:
            fn(slide, spec)

def _page_note(page) -> str:
    """The speaker note identifying a slide: report ▸ group ▸ page.

    Every part is a registry LABEL, so it follows the geography the deck is
    being built for - a Europe deck's notes say "Europe Overview", not "Global
    Overview".  The slug is untouched (see geo.label)."""
    return (f"{geo.label(page['report'])} {geo.label(page['group'])} "
            f"{geo.label(page['name'])}")

# ── fixed-coordinate pages ───────────────────────────────────────────
def _remap_spec(s):
    """A collected block's box moved from the old 84px header's coordinates
    into the content box under the current header - the same linear map
    `_remap_content` applies to finished shapes, applied BEFORE drawing so
    every renderer fits its text to the box it will really have."""
    t1 = layout.CONTENT_TOP - 3.0
    k = (_REMAP_B - t1) / (_REMAP_B - _REMAP_T0)
    y = float(s.get("y", 0))
    ny = (t1 + (y - _REMAP_T0) * k if y >= _REMAP_T0
          else y + (t1 - _REMAP_T0))
    return {**s, "y": ny, "h": float(s.get("h", 0)) * k}

def _widen_text_columns(specs):
    """Give a column of summary cards that cannot hold its text at 11pt a
    little of the width of the chart beside it - up to 20% of the chart -
    so its bullets wrap onto fewer lines and fit where the template put
    them, rather than one or two spilling onto a slide of their own."""
    specs = [dict(s) for s in specs]
    cols: dict = {}
    for i, s in enumerate(specs):
        if s.get("kind") == "sumpanel":
            cols.setdefault((round(s["x"]), round(s["w"])), []).append(i)
    for (cx, cw), idx in cols.items():
        top = min(specs[i]["y"] for i in idx)
        bot = max(specs[i]["y"] + specs[i]["h"] for i in idx)
        avail = sum(specs[i]["h"] for i in idx)

        def need(extra, idx=idx, cw=cw):
            return sum(_sumpanel_need(specs[i].get("bullets"), cw + extra,
                                      compact=True) for i in idx)

        if need(0.0) <= avail:
            continue
        # the chart(s) whose right edge meets this column, alongside it
        figs = [j for j, s in enumerate(specs) if s.get("kind") == "figure"
                and abs((s["x"] + s["w"]) - cx) <= 24
                and s["y"] < bot and s["y"] + s["h"] > top]
        if not figs:
            continue
        room = min(specs[j]["w"] for j in figs) * 0.20
        extra = 0.0
        while extra < room and need(extra) > avail:
            extra += 10.0
        extra = min(extra, room)
        for j in figs:
            specs[j]["w"] -= extra
        for i in idx:
            specs[i]["x"] -= extra
            specs[i]["w"] += extra
    return specs

def _rebalance_cards(specs):
    """Summary cards stacked in one column of a fixed-layout page share the
    column's height by what their text needs at 11pt - so a long card
    borrows from a short one beside it instead of spilling a bullet or two
    onto a slide of its own.  A column whose cards all fit is left as the
    template laid it out."""
    specs = [dict(s) for s in specs]
    cols: dict = {}
    for i, s in enumerate(specs):
        if s.get("kind") == "sumpanel":
            cols.setdefault((round(s["x"]), round(s["w"])), []).append(i)
    for idx in cols.values():
        if len(idx) < 2:
            continue
        idx.sort(key=lambda i: specs[i]["y"])
        w = specs[idx[0]]["w"]
        needs = [_sumpanel_need(specs[i].get("bullets"), w) for i in idx]
        if all(n <= specs[i]["h"] for n, i in zip(needs, idx)):
            continue
        gaps = [specs[b]["y"] - (specs[a]["y"] + specs[a]["h"])
                for a, b in zip(idx, idx[1:])]
        avail = (specs[idx[-1]]["y"] + specs[idx[-1]]["h"]
                 - specs[idx[0]]["y"] - sum(gaps))
        if sum(needs) > avail:
            needs = [_sumpanel_need(specs[i].get("bullets"), w, compact=True)
                     for i in idx]
            if sum(needs) > avail:
                continue                  # too much even so: it continues
        spare = (avail - sum(needs)) / len(idx)
        y = specs[idx[0]]["y"]
        for j, i in enumerate(idx):
            specs[i]["y"], specs[i]["h"] = y, needs[j] + spare
            y += needs[j] + spare + (gaps[j] if j < len(gaps) else 0)
    return specs

def _panel_need(s, w=None):
    """The height a panel's content actually wants, or None for one whose
    height is its own business (a chart is as tall as its box)."""
    w = s.get("w") if w is None else w
    if s.get("kind") == "table" and s.get("df") is not None and len(s["df"]):
        return _table_need(_table_geo(s, w), 0, len(s["df"]),
                           s.get("total_row") is not None)
    if s.get("kind") == "sumpanel" and s.get("bullets"):
        return _sumpanel_need(s.get("bullets"), w)
    return None


def _reclaim_column(specs):
    """In a column, a panel with height to spare hands it to one below that
    has not got enough.

    The template gives every panel a fixed box, and a table whose rows
    outgrew its box was cut there even when the panel ABOVE it in the same
    column was half empty - the white space under one table, and the two or
    three rows of the next one pushed onto a slide of their own.  Panels
    keep their column, their order and the gaps between them; only the
    share-out of the column's height changes, and only when something in it
    is short of room.  A chart is not asked to give here - `_block_give`
    decides that, on the pages laid out in rows.
    """
    specs = [dict(s) for s in specs]
    cols: dict = {}
    for i, s in enumerate(specs):
        if s.get("kind") not in ("table", "sumpanel", "figure"):
            continue
        if any(s.get(k) is None for k in ("x", "y", "w", "h")):
            continue
        cols.setdefault((round(s["x"]), round(s["w"])), []).append(i)
    for idx in cols.values():
        if len(idx) < 2:
            continue
        idx.sort(key=lambda i: specs[i]["y"])
        needs = [_panel_need(specs[i]) for i in idx]
        if not any(n is not None and n > specs[i]["h"]
                   for n, i in zip(needs, idx)):
            continue                     # nothing in this column is short
        gaps = [specs[b]["y"] - (specs[a]["y"] + specs[a]["h"])
                for a, b in zip(idx, idx[1:])]
        span = (specs[idx[-1]]["y"] + specs[idx[-1]]["h"]
                - specs[idx[0]]["y"] - sum(gaps))
        # a panel with no need of its own (a chart) holds the height it has
        want = [specs[i]["h"] if n is None else n
                for n, i in zip(needs, idx)]
        if sum(want) > span:
            continue                     # the column cannot hold it all
        spare = (span - sum(want)) / len(idx)
        y = specs[idx[0]]["y"]
        for j, i in enumerate(idx):
            specs[i]["y"], specs[i]["h"] = int(y), int(want[j] + spare)
            y += want[j] + spare + (gaps[j] if j < len(gaps) else 0)
    return specs


def _grow_panels(specs):
    """An over-full table or card takes the empty space under it before
    anything is split off to a slide of its own.

    A fixed-layout page carries the box the template drew, and a panel whose
    content outgrew that box used to be cut there - even when the space
    below it on the slide was empty.  That is what left a slide with white
    space under a table and put the rest of that table on a slide of its
    own.  A panel grows only into space nothing else occupies: the next
    thing below it in its own column, or the foot of the content box.
    """
    specs = [dict(s) for s in specs]
    boxes = [(s.get("x"), s.get("y"), s.get("w"), s.get("h")) for s in specs]
    for i, s in enumerate(specs):
        x, y, w, h = boxes[i]
        if None in (x, y, w, h):
            continue
        kind = s.get("kind")
        if kind == "table" and s.get("df") is not None and len(s["df"]):
            need = _table_need(_table_geo(s, w), 0, len(s["df"]),
                               s.get("total_row") is not None)
        elif kind == "sumpanel" and s.get("bullets"):
            need = _sumpanel_need(s.get("bullets"), w)
        else:
            continue
        if need <= h:
            continue
        # the first thing below this panel that stands in its column
        floor = layout.CONTENT_BOTTOM
        for j, (ox, oy, ow, oh) in enumerate(boxes):
            if j == i or None in (ox, oy, ow, oh):
                continue
            if ox < x + w and x < ox + ow and oy >= y + h - 1:
                floor = min(floor, oy - layout.GUTTER_Y)
        grown = min(need, floor - y)
        if grown > h:
            s["h"] = grown
    return specs


def _overflow_specs(specs):
    """Split a fixed-layout page's over-full panels at the 11pt standard.

    A table or summary card whose content no longer fits its fixed box
    keeps what fits there - a table only a few rows over is squeezed
    instead, per the rule in `_split_table` - and the rest becomes a block
    that continues on numbered slides after this one.  Returns (specs to
    draw here, blocks to continue)."""
    # Adjust on the slide first, and split only what is still over: a text
    # column takes a little width from the chart beside it, its cards share
    # the column's height, a panel with room to spare hands it to one below
    # that is short, and anything still over-full grows into the empty space
    # under it.  Only then is a table cut.
    specs = _grow_panels(_reclaim_column(
        _rebalance_cards(_widen_text_columns(specs))))
    keep, more = [], []
    for s in specs:
        kind = s.get("kind")
        if kind == "table" and s.get("df") is not None and len(s["df"]):
            head, tail = _split_table(s, _table_geo(s, s["w"]), s["h"], True)
            keep.append(head if head is not None else s)
            if tail is not None:
                more.append({"kind": "table", "df": tail["df"],
                             "index_name": s.get("index_name", ""),
                             "title": s.get("title"),
                             "total_row": tail.get("total_row"),
                             "frac": 1.0, "_chunk": True,
                             # where it sat, so the continuation can put
                             # it back beside whatever shared its row
                             "_x": s.get("x"), "_y": s.get("y"),
                             "_w": s.get("w")})
            continue
        if kind == "sumpanel":
            bl = list(s.get("bullets") or [])
            if (len(bl) > 1 and
                    _sumpanel_need(bl, s["w"], compact=True) > s["h"]):
                j = 0
                while (j < len(bl) and
                       _sumpanel_need(bl[:j + 1], s["w"]) <= s["h"]):
                    j += 1
                j = max(1, min(j, len(bl) - 1))
                keep.append({**s, "bullets": bl[:j]})
                more.append({"kind": "sumpanel", "title": s.get("title"),
                             "bullets": bl[j:], "frac": 1.0, "_chunk": True,
                             "ch": _sumpanel_need(bl[j:], layout.CONTENT_W),
                             "_x": s.get("x"), "_y": s.get("y"),
                             "_w": s.get("w")})
                continue
        keep.append(s)
    return keep, more

_CONT_ROW_BAND = 24       # px: panels this close in y shared a row


def _continuation_rows(more):
    """How the overflow is laid out on the slides that continue a page.

    One block per row, each at the full width, reads well and fills the
    slide down the page - so that is kept while it costs nothing.  But two
    tables that sat SIDE BY SIDE come back stacked that way, and once
    enough of them spill, stacking runs off the foot and opens a SECOND
    continuation slide that need not exist.  So the stack is measured
    first, and only when it will not fit one slide are the blocks put back
    into the rows they came from - side by side, sharing the width as they
    did on the page.
    """
    stacked = [[dict(b)] for b in more]
    need = 0.0
    for b in more:
        h = _panel_need(b, layout.CONTENT_W)
        need += (h if h is not None else b.get("ch") or 0) + layout.GUTTER_Y
    if need - layout.GUTTER_Y <= layout.CONTENT_BOTTOM - layout.CONTENT_TOP:
        return stacked

    more = [dict(b) for b in more]
    rows, cur = [], []
    for b in sorted(more, key=lambda b: (b.get("_y") or 0, b.get("_x") or 0)):
        if cur and abs((b.get("_y") or 0)
                       - (cur[0].get("_y") or 0)) > _CONT_ROW_BAND:
            rows.append(cur)
            cur = []
        cur.append(b)
    if cur:
        rows.append(cur)
    for row in rows:
        span = sum(b.get("_w") or 0 for b in row)
        for b in row:
            b["frac"] = ((b.get("_w") or 0) / span if span
                         else 1.0 / len(row))
    return rows


def _spec_page_slides(prs, page, title, specs, chips):
    """A page drawn from collected specs at fixed coordinates: its slide,
    plus a numbered continuation - "(1/2)", "(2/2)" - for whatever no longer
    fits it at the 11pt content standard."""
    keep, more = _overflow_specs([_remap_spec(s) for s in specs])
    extra = _flow_fill(_continuation_rows(more)) if more else []
    k = 1 + len(extra)
    note = _page_note(page)
    slide = _base_slide(prs, title if k == 1 else f"{title} (1/{k})",
                        chips, remap=False)
    _draw_specs(slide, keep)
    _footer(slide)
    slide.notes_slide.notes_text_frame.text = \
        note + (f"  (1/{k})" if k > 1 else "")
    made = [slide]
    for j, dp in enumerate(extra, start=2):
        cont = _base_slide(prs, f"{title} ({j}/{k})", (), remap=False)
        cspecs: list[dict] = []
        components.set_collector(cspecs)
        try:
            layout._components(dp["blocks"])
        finally:
            components.set_collector(None)
        _draw_specs(cont, cspecs)
        _footer(cont)
        cont.notes_slide.notes_text_frame.text = f"{note}  ({j}/{k})"
        made.append(cont)
    return made

def _page_slide(prs, page, region, country):
    """Non-flow (fixed-layout) page - its slide from collected specs, plus a
    numbered continuation for any panel that overflows.  Returns a list."""
    values = {"region": region, "country": country}
    args = [values[r] for r in page["slicers"]]
    specs: list[dict] = []
    components.set_collector(specs)
    try:
        title, _ = page["render"](*args)
        title = geo.label(title)
    finally:
        components.set_collector(None)
    return _spec_page_slides(prs, page, title, specs,
                             [values[r] for r in page["slicers"]])

def _taxonomy_one_slide(prs, page, title, spec, values):
    """Collapse the taxonomy's split trees onto ONE PPT slide WHEN they fit at a
    readable font; if the merged trees are too tall for one slide, fall back to
    the dashboard's own two-slide split so nothing overflows.  (Export-only -
    the live dashboard is untouched either way.)"""
    # merge the per-slide panels by title (Region / Segment / About)
    merged, order = {}, []
    for sl in spec["slides"]:
        for b in sl:
            if b.get("kind") != "panel":
                continue
            t = b.get("title") or ""
            if t not in merged:
                merged[t] = {"x": b["x"], "y": b["y"], "w": b["w"],
                             "h": b["h"], "lines": []}
                order.append(t)
            merged[t]["lines"].extend(b.get("export_lines") or [])

    def _note(slide):
        slide.notes_slide.notes_text_frame.text = \
            _page_note(page)

    def _draw(slide, panels):
        for p in panels:
            _render_text_panel(slide, {"kind": "text", **p})
        _footer(slide)
        _note(slide)

    # would the tallest merged column fit on one slide at a readable size?
    # (this is exactly the font _render_text_panel would pick for that height)
    top_y = min((m["y"] for m in merged.values()), default=96)
    avail_h = 720 - top_y - 34                 # slide is 720px tall; leave footer
    max_n = max((len(m["lines"]) for m in merged.values()), default=1)
    child_fs = (avail_h - 6) / max(1, max_n) / 1.5

    if child_fs >= 6.8:                         # fits comfortably -> one slide
        slide = _base_slide(prs, title, [values[r] for r in page["slicers"]])
        _draw(slide, [{"title": t, "x": merged[t]["x"], "y": merged[t]["y"],
                       "w": merged[t]["w"], "h": avail_h,
                       "lines": merged[t]["lines"]} for t in order])
        return [slide]

    # too tall -> two slides, mirroring the dashboard's own split
    made, k = [], len(spec["slides"])
    for j, sl in enumerate(spec["slides"], start=1):
        heading = title if k == 1 else f"{title} ({j}/{k})"
        chips = [values[r] for r in page["slicers"]] if j == 1 else ()
        slide = _base_slide(prs, heading, chips)
        _draw(slide, [{"title": b.get("title") or "", "x": b["x"], "y": b["y"],
                       "w": b["w"], "h": b["h"],
                       "lines": b.get("export_lines") or []}
                      for b in sl if b.get("kind") == "panel"])
        made.append(slide)
    return made

# ── Taxonomy flow chart -> native shapes ─────────────────────────────
_C_FLOW_LINE = _rgb(FLOW_LINE)
_C_FLOW_LEAF_EDGE = _rgb(FLOW_LEAF_EDGE)
_C_FLOW_LEAF_TEXT = _rgb(FLOW_LEAF_TEXT)

def _flow_connector(slide, x1, y1, x2, y2):
    """One straight connector segment matching the dashboard's thin lines."""
    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                    E(x1), E(y1), E(x2), E(y2))
    cn.line.color.rgb = _C_FLOW_LINE
    cn.line.width = Pt(1.25)
    cn.shadow.inherit = False
    return cn

def _flow_box_ppt(slide, n):
    """One taxonomy box as a native, editable rounded rectangle + text."""
    x, y, w, h = n["x"], n["y"], n["w"], n["h"]
    lvl = n["level"]
    if lvl == 0:
        fill, line, ink, bold = _C_HEADER, _C_HEADER, _C_WHITE, True
    elif lvl == 1:
        fill, line, ink, bold = _C_ACCENT, _C_ACCENT, _C_WHITE, True
    else:
        fill, line, ink, bold = _C_WHITE, _C_FLOW_LEAF_EDGE, \
            _C_FLOW_LEAF_TEXT, False
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 E(x), E(y), E(w), E(h))
    shp.adjustments[0] = min(0.4, 7.0 / max(1, min(w, h)))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = line
    shp.line.width = Pt(1)
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.word_wrap = True
    for m in ("margin_top", "margin_bottom"):
        setattr(tf, m, Emu(9000))
    left = n.get("align", "center") == "left"
    tf.margin_left = Emu(70000 if left else 40000)
    tf.margin_right = Emu(40000)
    _set_text(tf, [[(str(n["text"]), n["fs"], bold, ink)]],
              align=PP_ALIGN.LEFT if left else PP_ALIGN.CENTER,
              anchor=MSO_ANCHOR.MIDDLE)
    return shp

def _reflow_taxonomy_ppt(fslide):
    """Re-measure the taxonomy boxes for PowerPoint's (taller) text metrics so
    no label overflows: each box grows to fit its wrapped text, leaf columns
    are re-stacked within their existing vertical span (gaps shrink to make
    room - the column keeps its footprint, so nothing spills off the slide),
    and the connectors are re-routed to the new leaf centres.  Fully data-
    driven - longer labels simply get taller boxes.  PPT-only; the dashboard
    geometry is untouched."""
    nodes = fslide.get("nodes") or []
    root = next((n for n in nodes if n["level"] == 0), None)
    branches = sorted((n for n in nodes if n["level"] == 1),
                      key=lambda n: n["x"])
    leaves = [n for n in nodes if n["level"] == 2]

    def need_h(n):
        left = n.get("align", "center") == "left"
        usable = n["w"] - (12 if left else 9)     # matches the PPT box margins
        # the size is POINTS here but _wrap_count is a CSS-px model, so scale it
        # -- without this the box is sized for fewer lines than PowerPoint
        # actually draws and the last line sits on (or through) the border
        lines = layout._wrap_count(n["text"], usable, n["fs"] * _PPT_FONT_K)
        # a PPT line is ~1.6-1.7x the FONT SIZE in slide px (pt->px 1.333 x
        # ~1.2-1.3 line spacing), so size boxes on that, plus breathing room
        return int(round(lines * n["fs"] * 1.72 + 14))

    if root:
        root["h"] = need_h(root)
    if branches:
        bh = max(need_h(b) for b in branches)     # uniform, aligned branch row
        for b in branches:
            b["h"] = bh

    edges = []
    # trunk: root -> horizontal bus -> each branch top
    if root and branches:
        rcx = root["x"] + root["w"] / 2
        rbot = root["y"] + root["h"]
        btop = min(b["y"] for b in branches)
        busy = (rbot + btop) / 2
        cxs = [b["x"] + b["w"] / 2 for b in branches]
        edges.append({"pts": [(rcx, rbot), (rcx, busy)]})
        if len(cxs) > 1:
            edges.append({"pts": [(min(cxs), busy), (max(cxs), busy)]})
        for cx in cxs:
            edges.append({"pts": [(cx, busy), (cx, btop)]})

    # re-stack each leaf SUB-COLUMN within its original vertical span.  A branch
    # normally has one, but a page with few branches lays its leaves out in
    # several (layout._leaf_grid), and re-stacking those into a single tower
    # would push the tail of the list off the slide.  Grouping by the leaf's own
    # x keeps whatever grid the geometry chose.
    for b in branches:
        mine = [l for l in leaves
                if b["x"] <= l["x"] + l["w"] / 2 <= b["x"] + b["w"]]
        cols = [sorted((l for l in mine if l["x"] == lx), key=lambda l: l["y"])
                for lx in sorted({l["x"] for l in mine})]
        for col in cols:
            if not col:
                continue
            top0 = max(min(l["y"] for l in col), b["y"] + b["h"] + 16)
            orig_bottom = max(l["y"] + l["h"] for l in col)
            hs = [need_h(l) for l in col]
            n = len(col)
            total = sum(hs)
            # keep the original footprint when the taller boxes still fit; only
            # extend downward (up to just above the footer) if they need it
            span = max(orig_bottom - top0, total + (n - 1) * 8)
            span = min(span, 702 - top0)
            gap = (span - total) / (n - 1) if n > 1 else 0
            gap = max(6.0, gap)
            y = float(top0)
            for l, hh in zip(col, hs):
                l["y"], l["h"] = int(round(y)), int(hh)
                y += hh + gap
        # the connectors are rebuilt by the SAME routine the dashboard uses, so
        # the two surfaces cannot draw the tree differently
        edges.extend(layout.flow_edges(b, cols))

    for e in edges:
        e["pts"] = [(int(round(x)), int(round(yy))) for x, yy in e["pts"]]
    return {"nodes": nodes, "edges": edges} if root else fslide

def _render_flow_ppt(slide, fslide):
    """Draw connectors (behind) then boxes for one taxonomy flow slide."""
    fslide = _reflow_taxonomy_ppt(fslide)
    for e in fslide["edges"]:
        pts = e["pts"]
        for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
            _flow_connector(slide, x1, y1, x2, y2)
    for n in fslide["nodes"]:
        _flow_box_ppt(slide, n)

def _taxonomy_flow_slides(prs, page, title, spec, values):
    """Each flow slide -> one native PPT slide, mirroring the dashboard."""
    fslides = spec["flow"] or [{"nodes": [], "edges": []}]
    k = len(fslides)
    made = []
    for j, fsl in enumerate(fslides, start=1):
        heading = title if k == 1 else f"{title} ({j}/{k})"
        chips = [values[r] for r in page["slicers"]] if j == 1 else ()
        slide = _base_slide(prs, heading, chips)
        _render_flow_ppt(slide, fsl)
        _footer(slide)
        note = _page_note(page)
        if k > 1:
            note += f"  ({j}/{k})"
        slide.notes_slide.notes_text_frame.text = note
        made.append(slide)
    return made

# ── Market Signals -> native editable table ──────────────────────────
def _sig_cell(cell, text, size, bold, color, fill):
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill
    cell.margin_left = Emu(70000)
    cell.margin_right = Emu(70000)
    cell.margin_top = Emu(45000)
    cell.margin_bottom = Emu(45000)
    cell.vertical_anchor = MSO_ANCHOR.TOP
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = str(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = FONT

def _render_signals_native(slide, sig):
    """The 3-column Market Signals editorial as a native, editable PPT table."""
    cols, rows = sig["cols"], sig["rows"]
    ncols = max(1, len(cols))
    nrows = len(rows) + 1
    x, y, w, h = 15, 96, 1250, 610
    gf = slide.shapes.add_table(nrows, ncols, E(x), E(y), E(w), E(h))
    tbl = gf.table
    tbl.first_row = False
    tbl.horz_banding = False
    fr = [0.30, 0.42, 0.28][:ncols]
    fr = fr + [1.0 / ncols] * (ncols - len(fr))
    s = sum(fr) or 1.0
    for c in range(ncols):
        tbl.columns[c].width = E(w * fr[c] / s)
    head_bg, zebra_bg = _rgb(SIG_HEAD_BG), _rgb(SIG_ZEBRA)
    for c, name in enumerate(cols):
        _sig_cell(tbl.cell(0, c), str(name).upper(), 11, True,
                  _C_HEADER, head_bg)
    for r, row in enumerate(rows, start=1):
        bg = zebra_bg if r % 2 == 0 else _C_WHITE
        for c, v in enumerate(row):
            first = c == 0
            _sig_cell(tbl.cell(r, c), v, 10.5, first,
                      _C_HEADER if first else _C_TEXT, bg)
    tbl.rows[0].height = E(38)
    body_h = max(1, (h - 38) // max(1, len(rows)))
    for r in range(1, nrows):
        tbl.rows[r].height = E(body_h)

def _render_segstart_native(slide, spec, continued=False):
    """Segment-opening analysis page as native, editable shapes - intro line, a
    themed sub-segment value matrix and a three-panel description row.  Mirrors
    components.seg_start_view on the dashboard.

    A page dealt over several slides (_segstart_slides) passes the columns of
    this slide only, with ``first`` = the index of its first sub-segment so the
    badges keep counting; a ``continued`` slide gives its description row to
    the insights of its own sub-segments."""
    cols = spec["cols"]
    rl = spec["row_labels"]
    n = len(cols) or 1
    first = int(spec.get("first", 0))
    MX = 34
    NAVY, ACC, GRD = _rgb(VC_NAVY), _C_ACCENT, _rgb("#E1E8EE")
    intro = _textbox(slide, MX, 90, 1280 - 2 * MX, 34)
    intro.text_frame.word_wrap = True
    _set_text(intro.text_frame, [[(spec["intro"], 11.5, False, _C_TEXT)]])
    # ── value matrix ──
    lab_w = 150
    x0 = MX + lab_w + 12
    col_w = (1280 - MX - x0 - (n - 1) * 12) / n
    badge_d, badge_y, head_y, head_h, r_h = 34, 138, 174, 48, 52
    # under fit_to_slide a long sub-segment name steps its header size down
    # until it fits the column, rather than running out of it
    head_fs = 11
    if _FIT_TO_SLIDE["on"] and cols:
        head_fs = (_segstart_head_fs([c["name"] for c in cols], col_w)
                   or MIN_CONTENT_PT)
    rows_y = [head_y + head_h + 2 + k * (r_h + 2) for k in range(3)]
    lab_cells = [(rl[0], head_y, head_h, "#EEF3F7", True),
                 (rl[1], rows_y[0], r_h, "#F7FAFC", False),
                 (rl[2], rows_y[1], r_h, "#F7FAFC", False),
                 (rl[3], rows_y[2], r_h, "#F7FAFC", False)]
    for txt, ly, lh, bg, is_head in lab_cells:
        cell = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(MX), E(ly),
                                      E(lab_w), E(lh))
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(bg)
        cell.line.color.rgb = GRD
        cell.line.width = Pt(0.5)
        cell.shadow.inherit = False
        tb = _textbox(slide, MX + 9, ly, lab_w - 14, lh)
        tb.text_frame.word_wrap = True
        _set_text(tb.text_frame, [[(txt, 10.5, is_head,
                                    NAVY if is_head else _C_MUTED)]],
                  anchor=MSO_ANCHOR.MIDDLE)
    for i, c in enumerate(cols):
        cx = x0 + i * (col_w + 12)
        color = _rgb(c["color"])
        hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, E(cx),
                                     E(head_y), E(col_w), E(head_h))
        hdr.adjustments[0] = 0.18
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = color
        hdr.line.fill.background()
        hdr.shadow.inherit = False
        flat = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(cx),
                                      E(head_y + head_h / 2), E(col_w),
                                      E(head_h / 2))
        flat.fill.solid()
        flat.fill.fore_color.rgb = color
        flat.line.fill.background()
        flat.shadow.inherit = False
        hl = _textbox(slide, cx + 4, head_y, col_w - 8, head_h)
        hl.text_frame.word_wrap = True
        _set_text(hl.text_frame, [[(c["name"], head_fs, True, _C_WHITE)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        for k, val in enumerate((c["v26"], c["v36"], c["cagr"])):
            vc = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(cx), E(rows_y[k]),
                                        E(col_w), E(r_h))
            vc.fill.solid()
            vc.fill.fore_color.rgb = _C_WHITE if k % 2 == 0 else _rgb("#F5F8FB")
            vc.line.color.rgb = _rgb("#E6ECF2")
            vc.line.width = Pt(0.5)
            vc.shadow.inherit = False
            _set_text(vc.text_frame, [[(val, 15, True, color)]],
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        bd = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                    E(cx + col_w / 2 - badge_d / 2), E(badge_y),
                                    E(badge_d), E(badge_d))
        bd.fill.solid()
        bd.fill.fore_color.rgb = color
        bd.line.color.rgb = _rgb("#EAF2F8")
        bd.line.width = Pt(2.5)
        bd.shadow.inherit = False
        _set_text(bd.text_frame, [[(str(first + i + 1), 15, True, _C_WHITE)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # ── three description panels ──
    dy, ph = 402, 282

    def panel(x, w, title, paras, bg="#FFFFFF", edge="#DCE6EF"):
        pnl = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, E(x), E(dy),
                                     E(w), E(ph))
        pnl.adjustments[0] = 0.03
        pnl.fill.solid()
        pnl.fill.fore_color.rgb = _rgb(bg)
        pnl.line.color.rgb = _rgb(edge)
        pnl.line.width = Pt(0.75)
        pnl.shadow.inherit = False
        tt = _textbox(slide, x + 14, dy + 10, w - 28, 20)
        _set_text(tt.text_frame, [[(title, 12.5, True, NAVY)]])
        rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(x + 14), E(dy + 32),
                                      E(w - 28), E(2))
        rule.fill.solid()
        rule.fill.fore_color.rgb = ACC
        rule.line.fill.background()
        rule.shadow.inherit = False
        body = _textbox(slide, x + 14, dy + 40, w - 28, ph - 50)
        body.text_frame.word_wrap = True
        # AUTO-SHRINK to the card, like _render_summary_panel: a segment with
        # many sub-segments used to push its last bullets straight through the
        # panel's bottom border.  Every run is scaled by the same ratio, so the
        # bullet glyph / name / body size relationship is preserved.
        chunks = ["".join(str(r[0]) for r in p) for p in paras] or [""]
        base = max((r[1] for p in paras for r in p), default=10.5)
        fs, sp, ln = _fit_text_fs(
            chunks, w - 34, ph - 50,
            sizes=tuple(round(base * k, 2)
                        for k in (1.0, .95, .9, .85, .8, .75, .7, .65, .6)))
        k = (fs / base) if base else 1.0
        _set_text(body.text_frame,
                  [[(r[0], r[1] * k, r[2], r[3]) for r in p] for p in paras],
                  space_after=sp, line=ln)
        return x, w

    b_paras = [[("●  ", 9, False, _rgb(it["color"])),
                (it["name"] + ": ", 10, True, _rgb(it["color"])),
                (it["text"], 10, False, _C_TEXT)]
               for it in spec["insights"]]
    if continued:
        # the segment-level panels were drawn on the first slide; this one
        # carries the insights of its own sub-segments, across the row
        panel(34, 1246 - 34, "Sub-Segment Insights", b_paras)
        srcb = _textbox(slide, 34, 686, 360, 18)     # 11pt: clear of the footer
        sr = srcb.text_frame.paragraphs[0].add_run()
        sr.text = components.SOURCE_TEXT
        sr.font.size, sr.font.italic, sr.font.name = Pt(8), True, FONT
        sr.font.color.rgb = _C_MUTED
        return
    a_paras = [[("▪  ", 10, False, ACC), (b, 10.5, False, _C_TEXT)]
               for b in spec["highlights"]]
    # Same proportions as components._seg_desc_columns: the widest box
    # goes to the panel carrying one bullet per sub-segment.
    panel(34, 340, "Key Highlights", a_paras)
    panel(386, 470, "Sub-Segment Insights", b_paras)
    cx, cw = 868, 1246 - 868
    c_paras = [[("✓  ", 10, True, _rgb(GOOD)), (p, 10.5, False, _C_TEXT)]
               for p in spec["overview"]["points"]]
    panel(cx, cw, "Segment Overview", c_paras, bg="#EAF3FB", edge="#CFE2F2")
    # highlighted stat box inside the overview panel
    sb_y = dy + ph - 74
    sbox = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(cx + 14), E(sb_y),
                                  E(cw - 28), E(58))
    sbox.fill.solid()
    sbox.fill.fore_color.rgb = _C_WHITE
    sbox.line.color.rgb = _rgb("#CFE2F2")
    sbox.line.width = Pt(0.5)
    sbox.shadow.inherit = False
    # (a 4px accent bar used to sit on the box's left edge - removed
    #  2026-09-07, the stat now reads as plain text in a plain box)
    st = _textbox(slide, cx + 24, sb_y + 4, cw - 40, 50)
    st.text_frame.word_wrap = True
    _set_text(st.text_frame,
              [[(spec["overview"]["stat"], 11.5, True, NAVY)]],
              anchor=MSO_ANCHOR.MIDDLE)
    srcb = _textbox(slide, 34, 686, 360, 18)     # 11pt: clear of the footer
    sr = srcb.text_frame.paragraphs[0].add_run()
    sr.text = components.SOURCE_TEXT
    sr.font.size, sr.font.italic, sr.font.name = Pt(8), True, FONT
    sr.font.color.rgb = _C_MUTED

# The outlook page's value matrix gives every sub-segment a column.  A cross
# segmentation can carry 30+ of them, and at that count the columns are too
# narrow to name - so, under fit_to_slide, the matrix is dealt out over as
# many slides as it takes for every name to fit its column header.
_SEGSTART_MAX_COLS = 12
_SEGSTART_HEAD_H = 48                      # the column header's height, px

def _segstart_col_w(n):
    """Matrix column width for ``n`` columns (the geometry
    _render_segstart_native lays out)."""
    n = max(1, n)
    return (1280 - 34 - (34 + 150 + 12) - (n - 1) * 12) / n

def _segstart_head_fs(names, col_w, sizes=(MIN_CONTENT_PT,)):
    """The size at which every name fits its column header, or None.

    Only the 11pt content standard is offered (2026-09-11): a matrix with
    more sub-segments than can be named at 11pt is dealt over more slides
    (`_segstart_groups`) rather than having its headers shrunk."""
    usable = max(1.0, col_w - 8 - 6)       # header box inset + text margins
    for fs in sizes:
        lines = max((layout._wrap_count(str(nm), usable, fs * _PPT_FONT_K)
                     for nm in names), default=1)
        if lines * fs * _PPT_LINE_NAT * 4 / 3 <= _SEGSTART_HEAD_H - 4:
            return fs
    return None

def _segstart_groups(n, names):
    """How many slides the matrix is dealt over: the fewest whose columns
    can all be named legibly."""
    for k in range(1, max(1, n) + 1):
        per = -(-n // k)
        if per <= _SEGSTART_MAX_COLS and \
                _segstart_head_fs(names, _segstart_col_w(per)) is not None:
            return k
    return max(1, -(-n // _SEGSTART_MAX_COLS))

def _segstart_slides(prs, page, title, spec, values):
    """Segment-opening analysis page -> one native PPT slide - or, under
    fit_to_slide, one per group of sub-segments when there are more of them
    than the value matrix can name (the rest on numbered "(1/2)", "(2/2)" slides)."""
    chips = [values[r] for r in page["slicers"]]
    d = spec["segstart"]
    cols = list(d.get("cols") or [])
    k = (_segstart_groups(len(cols), [c["name"] for c in cols])
         if _FIT_TO_SLIDE["on"] and cols else 1)
    if k <= 1:
        slide = _base_slide(prs, title, chips)
        _render_segstart_native(slide, d)
        _footer(slide)
        slide.notes_slide.notes_text_frame.text = \
            _page_note(page)
        return [slide]
    # cols and insights are built from the same ranking (templates_main.
    # seg_start), so one slice of each belongs to the same sub-segments
    per = -(-len(cols) // k)
    insights = list(d.get("insights") or [])
    made = []
    for j in range(k):
        a, b = j * per, (j + 1) * per
        part = {**d, "cols": cols[a:b], "insights": insights[a:b],
                "first": a}
        heading = f"{title} ({j + 1}/{k})"
        slide = _base_slide(prs, heading, chips if j == 0 else ())
        _render_segstart_native(slide, part, continued=j > 0)
        _footer(slide)
        slide.notes_slide.notes_text_frame.text = \
            f"{_page_note(page)}  ({j + 1}/{k})"
        made.append(slide)
    return made

_DASH_MAP = {"dashed": "dash", "dotted": "sysDot"}

def _draw_primitives(slide, prims):
    """Draw a shared primitive list (text/rect/rrect/line/poly/shape/dot) as
    native, editable shapes.  The PPT twin of components.draw_primitives, so any
    primitive-based custom slide matches the dashboard exactly."""
    _AL = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
    for p in prims:
        if p.get("web_only"):                      # dashboard-only hover target
            continue
        t = p["t"]
        if t in ("rect", "rrect"):
            shp = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE if t == "rrect" else MSO_SHAPE.RECTANGLE,
                E(p["x"]), E(p["y"]), E(p["w"]), E(p["h"]))
            if t == "rrect":
                try:
                    shp.adjustments[0] = min(0.5, p.get("r", 8) / max(p["w"], p["h"]))
                except Exception:
                    pass
            fill = p.get("fill")
            if fill in (None, "none", "transparent"):
                shp.fill.background()
            else:
                shp.fill.solid()
                shp.fill.fore_color.rgb = _rgb(fill)
            if p.get("line") in (None, "none"):
                shp.line.fill.background()
            else:
                shp.line.color.rgb = _rgb(p["line"])
                shp.line.width = Pt(p.get("lw", 0.75))
                if p.get("dash") in _DASH_MAP:
                    _set_line_dash(shp, _DASH_MAP[p["dash"]])
            shp.shadow.inherit = False
        elif t == "line":
            ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(p["x1"]),
                                            E(p["y1"]), E(p["x2"]), E(p["y2"]))
            ln.line.color.rgb = _rgb(p["color"])
            ln.line.width = Pt(max(0.5, p["width"]))
            if p.get("dash") in _DASH_MAP:
                _set_line_dash(ln, _DASH_MAP[p["dash"]])
            ln.shadow.inherit = False
        elif t == "poly":
            _polyline(slide, p["pts"], p["color"], p["width"])
        elif t == "shape":                          # filled polygon
            _filled_poly(slide, p["pts"], p["fill"], p.get("line"),
                         p.get("lw", 0.75))
        elif t == "dot":
            _edged_dot(slide, p["cx"], p["cy"], p["d"], _rgb(p["color"]),
                       _rgb(p["edge"]))
        elif t == "text":
            tb = _textbox(slide, p["x"], p["y"], p["w"], p["h"])
            tf = tb.text_frame
            tf.word_wrap = True
            _set_text(tf,
                      [[(p["s"], p["size"], p["bold"], _rgb(p["color"]))]],
                      align=_AL.get(p["align"], PP_ALIGN.LEFT),
                      anchor=(MSO_ANCHOR.TOP if p.get("anchor") == "t"
                              else MSO_ANCHOR.MIDDLE),
                      line=p.get("line", 1.15 if p.get("wrap") else None))
            if p.get("underline"):
                for para in tf.paragraphs:
                    for run in para.runs:
                        run.font.underline = True

def _set_line_dash(shape, val):
    try:
        ln = shape.line._get_or_add_ln()
        pd = ln.find(qn("a:prstDash"))
        if pd is None:
            pd = parse_xml(f'<a:prstDash {nsdecls("a")} val="{val}"/>')
            ln.append(pd)
        else:
            pd.set("val", val)
    except Exception:
        pass

def _filled_poly(slide, pts, fill, line=None, lw=0.75):
    try:
        fb = slide.shapes.build_freeform(float(pts[0][0]), float(pts[0][1]),
                                         scale=_EMU_PER_PX)
        fb.add_line_segments([(float(px), float(py)) for px, py in pts[1:]],
                             close=True)
        shp = fb.convert_to_shape()
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(fill)
        if line in (None, "none"):
            shp.line.fill.background()
        else:
            shp.line.color.rgb = _rgb(line)
            shp.line.width = Pt(lw)
        shp.shadow.inherit = False
        return shp
    except Exception:
        return None

def _render_pricepattern_native(slide, data):
    """Per-region pricing-pattern page as native shapes (shared primitive list)."""
    _draw_primitives(slide, layout.pricepattern_primitives(data))

def _render_primitives_native(slide, prims):
    _draw_primitives(slide, prims)

def _primitives_slides(prs, page, title, spec, values):
    """Generic primitive-drawn page -> one native PPT slide."""
    chips = [values[r] for r in page["slicers"]]
    slide = _base_slide(prs, title, chips)
    _render_primitives_native(slide, spec["primitives"])
    _footer(slide)
    slide.notes_slide.notes_text_frame.text = \
        _page_note(page)
    return [slide]

def _primpages_slides(prs, page, title, spec, values):
    """A primitive page that runs to several slides (the taxonomy's segment
    grid) -> one native slide per page, numbered (j/k) like the flow pages.
    Drawn at the layout engine's own coordinates, under CONTENT_TOP, so the
    old-header remap (which also squashes heights) is skipped."""
    pages = spec["primpages"] or [[]]
    k = len(pages)
    made = []
    for j, prims in enumerate(pages, start=1):
        heading = title if k == 1 else f"{title} ({j}/{k})"
        chips = [values[r] for r in page["slicers"]] if j == 1 else ()
        slide = _base_slide(prs, heading, chips, remap=False)
        _render_primitives_native(slide, prims)
        _footer(slide)
        note = _page_note(page)
        if k > 1:
            note += f"  ({j}/{k})"
        slide.notes_slide.notes_text_frame.text = note
        made.append(slide)
    return made

def _pricepattern_slides(prs, page, title, spec, values):
    """Pricing-pattern page -> one native PPT slide PER page of region rows
    (spills with the same formatting when there are more regions than fit)."""
    pages = layout.pricepattern_pages(spec["pricepattern"])
    k = len(pages)
    made = []
    for j, pdata in enumerate(pages, start=1):
        chips = [values[r] for r in page["slicers"]] if j == 1 else ()
        # a page set spread over several slides is numbered in its heading,
        # like every other continuation in the deck
        heading = title if k == 1 else f"{title} ({j}/{k})"
        slide = _base_slide(prs, heading, chips)
        _render_pricepattern_native(slide, pdata)
        _footer(slide)
        note = _page_note(page)
        if k > 1:
            note += f"  ({j}/{k})"
        slide.notes_slide.notes_text_frame.text = note
        made.append(slide)
    return made

_ENDCARD_LINKS: list = []      # (pill shape, target slug) wired after slide build

def _render_endcards_native(slide, spec):
    """Segment-section closing card as native, editable shapes - four colour-
    headed panels (numbered badges + chevrons), a concluding line and a Next
    pill.  Mirrors components.end_cards on the dashboard."""
    cards = spec["cards"]
    n = len(cards) or 1
    MX, GAP = 34, 30
    cw = (1280 - 2 * MX - (n - 1) * GAP) / n
    badge_d, card_top, head_h, card_bot = 54, 172, 50, 502
    NAVY = _rgb(VC_NAVY)
    for i, c in enumerate(cards):
        cx = MX + i * (cw + GAP)
        color = _rgb(c["color"])
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, E(cx),
                                      E(card_top), E(cw), E(card_bot - card_top))
        card.adjustments[0] = 0.045
        card.fill.solid()
        card.fill.fore_color.rgb = _C_WHITE
        card.line.color.rgb = _rgb("#DCE4EC")
        card.line.width = Pt(0.75)
        card.shadow.inherit = False
        hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, E(cx),
                                     E(card_top), E(cw), E(head_h))
        hdr.adjustments[0] = 0.22
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = color
        hdr.line.fill.background()
        hdr.shadow.inherit = False
        flat = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(cx),
                                      E(card_top + head_h / 2), E(cw),
                                      E(head_h / 2))
        flat.fill.solid()
        flat.fill.fore_color.rgb = color
        flat.line.fill.background()
        flat.shadow.inherit = False
        lbl = _textbox(slide, cx, card_top + head_h - 26, cw, 24)
        _set_text(lbl.text_frame, [[(c["label"], 15, True, _C_WHITE)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        body = _textbox(slide, cx + 14, card_top + head_h + 6, cw - 28,
                        card_bot - (card_top + head_h) - 14)
        body.text_frame.word_wrap = True
        _set_text(body.text_frame, [[(c["body"], 11.5, False, _C_TEXT)]],
                  anchor=MSO_ANCHOR.MIDDLE)
        bd = slide.shapes.add_shape(MSO_SHAPE.OVAL, E(cx + cw / 2 - badge_d / 2),
                                    E(card_top - badge_d + 14), E(badge_d),
                                    E(badge_d))
        bd.fill.solid()
        bd.fill.fore_color.rgb = color
        bd.line.color.rgb = _rgb("#EAF2F8")
        bd.line.width = Pt(3)
        bd.shadow.inherit = False
        _set_text(bd.text_frame, [[(c["num"], 20, True, _C_WHITE)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if i < n - 1:
            ch = _textbox(slide, cx + cw + GAP / 2 - 14,
                          card_top + head_h / 2 - 20, 28, 40)
            _set_text(ch.text_frame, [[("›", 30, True, color)]],
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    ft = _textbox(slide, 0, 542, 1280, 26)
    _set_text(ft.text_frame, [[(spec["footer"], 16, True, NAVY)]],
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    for sx in (175, 1025):
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(sx), E(556),
                                        E(sx + 80), E(556))
        ln.line.color.rgb = _rgb("#9FB6CC")
        ln.line.width = Pt(1.5)
        ln.shadow.inherit = False
    pill_w, pill_h = 372, 44
    px = 1280 / 2 - pill_w / 2
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, E(px), E(590),
                                  E(pill_w), E(pill_h))
    pill.adjustments[0] = 0.5
    pill.fill.solid()
    pill.fill.fore_color.rgb = NAVY
    pill.line.fill.background()
    pill.shadow.inherit = False
    _set_text(pill.text_frame, [[("➜   " + spec["next"], 13, True,
                                  _C_WHITE)]],
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if spec.get("next_slug"):                     # wired after all slides built
        _ENDCARD_LINKS.append((pill, spec["next_slug"]))
    srcb = _textbox(slide, 30, 686, 360, 18)     # 11pt: clear of the footer
    sr = srcb.text_frame.paragraphs[0].add_run()
    sr.text = components.SOURCE_TEXT
    sr.font.size, sr.font.italic, sr.font.name = Pt(8), True, FONT
    sr.font.color.rgb = _C_MUTED

def _endcards_slides(prs, page, title, spec, values):
    """Segment-section closing card -> one native PPT slide."""
    chips = [values[r] for r in page["slicers"]]
    slide = _base_slide(prs, title, chips)
    _render_endcards_native(slide, spec["endcards"])
    _footer(slide)
    slide.notes_slide.notes_text_frame.text = \
        _page_note(page)
    return [slide]

def _signals_slides(prs, page, title, spec, values):
    """Market Signals -> one native PPT slide with an editable table."""
    chips = [values[r] for r in page["slicers"]]
    slide = _base_slide(prs, title, chips)
    _render_signals_native(slide, spec["signals"])
    _footer(slide)
    slide.notes_slide.notes_text_frame.text = \
        _page_note(page)
    return [slide]

_ARROW_RGB = {"high": ARROW_HIGH, "med": ARROW_MED, "low": ARROW_LOW}


def _circle(slide, cx, cy, d, fill=None, line=None, lw=1.0):
    o = slide.shapes.add_shape(MSO_SHAPE.OVAL, E(cx - d / 2), E(cy - d / 2),
                               E(d), E(d))
    if fill is None:
        o.fill.background()
    else:
        o.fill.solid()
        o.fill.fore_color.rgb = fill if not isinstance(fill, str) else _rgb(fill)
    if line is None:
        o.line.fill.background()
    else:
        o.line.color.rgb = line if not isinstance(line, str) else _rgb(line)
        o.line.width = Pt(lw)
    o.shadow.inherit = False
    return o


def _mini_gauge(slide, cx, cy, R, frac, color):
    """A small donut gauge - grey ring with a coloured wedge for ``frac`` and a
    white hole.  Wedge drawn as a freeform so no arc-angle maths is needed."""
    frac = max(0.0, min(1.0, frac))
    _circle(slide, cx, cy, 2 * R, fill=_rgb(GAUGE_TRACK))
    if frac > 0.001:
        pts = [(cx, cy)]
        steps = max(2, int(frac * 48))
        for s in range(steps + 1):
            a = math.radians(-90 + 360.0 * frac * s / steps)
            pts.append((cx + R * math.cos(a), cy + R * math.sin(a)))
        try:
            fb = slide.shapes.build_freeform(float(pts[0][0]), float(pts[0][1]),
                                             scale=_EMU_PER_PX)
            fb.add_line_segments([(float(x), float(y)) for x, y in pts[1:]],
                                 close=True)
            wedge = fb.convert_to_shape()
            wedge.fill.solid()
            wedge.fill.fore_color.rgb = _rgb(color)
            wedge.line.fill.background()
            wedge.shadow.inherit = False
        except Exception:
            pass
    _circle(slide, cx, cy, 2 * R * 0.56, fill=_C_WHITE)


def _draw_strategy_badge(slide, cx, cy, strat):
    """7 lettered circles (active = blue, inactive = grey), two rows (4 + 3)."""
    d, gap = 19.0, 3.0                   # holds an 11pt letter

    def row(items, ry):
        total = len(items) * d + (len(items) - 1) * gap
        sx = cx - total / 2
        for j, (letter, active) in enumerate(items):
            x = sx + j * (d + gap)
            _circle(slide, x + d / 2, ry, d,
                    fill=_rgb(NAVY_DEEP) if active else _rgb("#C9CDD1"))
            t = _textbox(slide, x, ry - d / 2, d, d)
            _set_text(t.text_frame, [[(letter, MIN_CONTENT_PT, True,
                                       _C_WHITE)]],
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                      pad=(0, 0, 0, 0))
    row(strat[:4], cy - 12)
    row(strat[4:], cy + 12)


def _draw_matrix_legend(slide, legend, note):
    # 11pt content standard: each entry is spaced by its label's measured
    # width, and the note lines by an 11pt line
    x, y, d = 30.0, 632.0, 18.0
    for k, v in legend:
        _circle(slide, x + d / 2, y + d / 2, d, fill=_rgb(NAVY_DEEP))
        t = _textbox(slide, x, y, d, d)
        _set_text(t.text_frame, [[(k, MIN_CONTENT_PT, True, _C_WHITE)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                  pad=(0, 0, 0, 0))
        tw = _text_px(v, MIN_CONTENT_PT)
        lb = _textbox(slide, x + d + 4, y - 1, tw + 8, d + 2)
        _set_text(lb.text_frame, [[(v, MIN_CONTENT_PT, False, _C_TEXT)]],
                  anchor=MSO_ANCHOR.MIDDLE, pad=(0, 0, 0, 0))
        x += d + 4 + tw + 18
    # ``note`` may be a single string or a LIST of lines (mirrors
    # components.competitor_matrix_view)
    if note:
        lines = [note] if isinstance(note, str) else list(note)
        for li, ln in enumerate(lines):
            nb = _textbox(slide, 30, 653 + li * 17, 1220, 17)
            _set_text(nb.text_frame, [[(ln, MIN_CONTENT_PT, True,
                                        _rgb(WARN_TEXT))]],
                      pad=(0, 0, 0, 0))


def _render_competitor_matrix_native(slide, data):
    """Creative competitor dashboard - companies as columns with visual cells
    (sparkles, ✓/✗, colour boxes, revenue, a share donut gauge, presence arrows
    and a strategy badge).  Mirrors components.competitor_matrix_view."""
    rows = data["rows"]
    n = len(rows) or 1
    star_cols = data.get("star_colors", list(STAR_COLORS))
    box_cols = data.get("box_colors", list(BOX_COLORS))
    labels = data.get("row_labels", [])
    if data.get("subtitle"):
        sb = _textbox(slide, 30, 86, 1100, 18)
        _set_text(sb.text_frame, [[(data["subtitle"], 9.5, False, _C_MUTED)]],
                  pad=(0, 2, 0, 0))
    LBLX, LBLW = 14, 120
    COLX0 = LBLX + LBLW + 2
    GRIGHT = 1266
    colw = (GRIGHT - COLX0) / n
    # (2026-09-11) re-spaced for the 11pt content standard: regional presence
    # gets the room its wrapped 11pt names need, and the matrix ends higher
    # so the legend and notes fit under it at 11pt
    bands = [("header", 108, 40), ("brand", 148, 28), ("portfolio", 176, 92),
             ("overall", 268, 36), ("target", 304, 36), ("share", 340, 84),
             ("presence", 424, 122), ("strategies", 546, 80)]
    band = {k: (y, h) for k, y, h in bands}
    top, bot = bands[0][1], bands[-1][1] + bands[-1][2]
    # white card behind the whole matrix so it reads clean (not on the pale-blue
    # page) - plus a soft tint on the header row and the left label column
    _white_panel(slide, LBLX, top - 2, GRIGHT - LBLX, bot - top + 4)
    hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(LBLX), E(top),
                                 E(GRIGHT - LBLX), E(bands[0][2]))
    hdr.fill.solid()
    hdr.fill.fore_color.rgb = _rgb("#F1F7FB")
    hdr.line.fill.background()
    hdr.shadow.inherit = False
    lcol = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(LBLX), E(top),
                                  E(LBLW), E(bot - top))
    lcol.fill.solid()
    lcol.fill.fore_color.rgb = _rgb("#F7FAFC")
    lcol.line.fill.background()
    lcol.shadow.inherit = False
    for _k, y, _h in bands[1:]:                    # horizontal separators
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(LBLX), E(y),
                                        E(GRIGHT), E(y))
        ln.line.color.rgb = _C_GRID
        ln.line.width = Pt(0.75)
        ln.shadow.inherit = False
    for gx in [LBLX, GRIGHT] + [COLX0 + c * colw for c in range(n + 1)]:
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(gx), E(top),
                                        E(gx), E(bot))
        ln.line.color.rgb = _C_GRID
        ln.line.width = Pt(0.75)
        ln.shadow.inherit = False
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(LBLX), E(bot),
                                    E(GRIGHT), E(bot))
    ln.line.color.rgb = _C_GRID
    ln.line.width = Pt(0.75)
    ln.shadow.inherit = False
    lbl_bands = ["brand", "portfolio", "overall", "target", "share",
                 "presence", "strategies"]
    for lb, key in zip(labels, lbl_bands):
        y, h = band[key]
        tb = _textbox(slide, LBLX + 4, y, LBLW - 6, h)
        _set_text(tb.text_frame, [[(lb, 8, True, _C_HEADER)]],
                  anchor=MSO_ANCHOR.MIDDLE, pad=(2, 2, 1, 1), line=1.0)
    maxshare = max((r["share"] for r in rows), default=1.0) or 1.0
    for i, r in enumerate(rows):
        cx0 = COLX0 + i * colw
        ccx = cx0 + colw / 2
        y, h = band["header"]
        hb = _textbox(slide, cx0 + 3, y + 2, colw - 6, h - 4)
        _set_text(hb.text_frame, [[(r["company"], 8.5, True, _C_TEXT)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                  pad=(1, 1, 1, 1), line=0.95)
        ul = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(ccx - 24),
                                        E(y + h - 3), E(ccx + 24), E(y + h - 3))
        ul.line.color.rgb = _C_ACCENT
        ul.line.width = Pt(1.5)
        ul.shadow.inherit = False
        y, h = band["brand"]
        bb = _textbox(slide, cx0 + 3, y, colw - 6, h)
        _set_text(bb.text_frame, [[(r["brand"], 7.5, False, _rgb("#AFB4B9"))]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        y, h = band["portfolio"]
        star_runs = [("✦ ", 13, False,
                      _rgb(star_cols[k % len(star_cols)] if k < r["stars"]
                           else "#D4D7DA")) for k in range(4)]
        sr = _textbox(slide, cx0 + 2, y + 4, colw - 4, 22)
        _set_text(sr.text_frame, [star_runs], align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE, pad=(0, 0, 0, 0))
        chk = [("✔ " if ok else "✘ ", 10, True,
                _rgb(OK_COLOR) if ok else _rgb(NOT_OK_COLOR)) for ok in r["apps"]]
        cr = _textbox(slide, cx0 + 2, y + 28, colw - 4, 20)
        _set_text(cr.text_frame, [chk], align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE)
        nb = len(r["boxes"])
        bxw, bxgap = 12.0, 6.0
        bx0 = ccx - (nb * bxw + (nb - 1) * bxgap) / 2
        for k, on in enumerate(r["boxes"]):
            rx = bx0 + k * (bxw + bxgap)
            box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(rx), E(y + 56),
                                         E(bxw), E(bxw))
            col = _rgb(box_cols[k % len(box_cols)])
            box.fill.solid()
            if on:
                box.fill.fore_color.rgb = col
                box.line.fill.background()
            else:
                box.fill.fore_color.rgb = _C_WHITE
                box.line.color.rgb = col
                box.line.width = Pt(1)
            box.shadow.inherit = False
        for key, val in (("overall", r["overall_rev"]),
                         ("target", r["target_rev"])):
            y, h = band[key]
            vb = _textbox(slide, cx0 + 3, y, colw - 6, h)
            _set_text(vb.text_frame,
                      [[(components.fmt_num(val), 8.5, False, _C_TEXT)]],
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        y, h = band["share"]
        R = min(26.0, h / 2 - 8)
        gcy = y + h / 2 - 2
        _mini_gauge(slide, ccx, gcy, R, r["share"] / maxshare, OK_COLOR)
        lab = _textbox(slide, ccx - R, gcy - 7, 2 * R, 14)
        _set_text(lab.text_frame, [[(f"{r['share']:.1f}%", 7.5, True,
                                     _C_HEADER)]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        y, h = band["presence"]
        paras = [[(nm + "  ", 7.5, False, _C_TEXT),
                  ("▲", 8.5, True, _rgb(_ARROW_RGB.get(st, ARROW_LOW)))]
                 for nm, st in r["presence"]]
        pb = _textbox(slide, cx0 + 5, y + 4, colw - 9, h - 8)
        _set_text(pb.text_frame, paras or [[("", 7.5, False, _C_TEXT)]],
                  align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
                  space_after=3, line=1.0)
        y, h = band["strategies"]
        _draw_strategy_badge(slide, ccx, y + h / 2, r["strategies"])
    _draw_matrix_legend(slide, data.get("legend", []), data.get("note", ""))


def _compmatrix_slides(prs, page, title, spec, values):
    """Creative competitor dashboard -> one native PPT slide."""
    chips = [values[r] for r in page["slicers"]]
    slide = _base_slide(prs, title, chips)
    _render_competitor_matrix_native(slide, spec["compmatrix"])
    _footer(slide)
    slide.notes_slide.notes_text_frame.text = \
        _page_note(page)
    return [slide]


_VC_PURPLE = _rgb(VC_PURPLE)
_VC_NAVY = _rgb(VC_NAVY)


def _wrect(slide, x, y, w, h, color):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(x), E(y), E(w), E(h))
    r.fill.solid()
    r.fill.fore_color.rgb = color
    r.line.fill.background()
    r.shadow.inherit = False
    return r


def _wovl(slide, x, y, w, h, color):
    o = slide.shapes.add_shape(MSO_SHAPE.OVAL, E(x), E(y), E(w), E(h))
    o.fill.solid()
    o.fill.fore_color.rgb = color
    o.line.fill.background()
    o.shadow.inherit = False
    return o


def _wtri(slide, x, y, w, h, color):
    t = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, E(x), E(y), E(w),
                               E(h))
    t.fill.solid()
    t.fill.fore_color.rgb = color
    t.line.fill.background()
    t.shadow.inherit = False
    return t


def _vc_icon(slide, kind, cx, cy, col):
    """A small white pictogram centred at (cx, cy); ``col`` (node colour) is
    used for cut-out details (wheels/door) so they read against the white."""
    W = _C_WHITE
    if kind == "sprout":
        _wrect(slide, cx - 1.5, cy - 2, 3, 17, W)
        lf = _wovl(slide, cx - 15, cy - 9, 15, 9, W)
        lf.rotation = -35
        rf = _wovl(slide, cx, cy - 9, 15, 9, W)
        rf.rotation = 35
    elif kind == "factory":
        _wrect(slide, cx - 18, cy - 2, 36, 18, W)
        _wrect(slide, cx - 15, cy - 13, 7, 12, W)
        _wrect(slide, cx - 4, cy - 9, 7, 8, W)
        _wrect(slide, cx + 7, cy - 12, 7, 11, W)
    elif kind == "truck":
        _wrect(slide, cx - 19, cy - 11, 22, 18, W)
        _wrect(slide, cx + 4, cy - 3, 13, 10, W)
        _wovl(slide, cx - 13, cy + 5, 9, 9, col)
        _wovl(slide, cx + 6, cy + 5, 9, 9, col)
    elif kind == "shop":
        _wtri(slide, cx - 19, cy - 15, 38, 12, W)
        _wrect(slide, cx - 15, cy - 4, 30, 20, W)
        _wrect(slide, cx - 4, cy + 4, 8, 12, col)          # door cut-out
    elif kind == "person":
        _wovl(slide, cx - 7, cy - 16, 14, 14, W)
        try:
            pts = [(cx - 13, cy + 15), (cx - 9, cy + 1), (cx + 9, cy + 1),
                   (cx + 13, cy + 15)]
            fb = slide.shapes.build_freeform(float(pts[0][0]),
                                             float(pts[0][1]), scale=_EMU_PER_PX)
            fb.add_line_segments([(float(x), float(y)) for x, y in pts[1:]],
                                 close=True)
            body = fb.convert_to_shape()
            body.fill.solid()
            body.fill.fore_color.rgb = W
            body.line.fill.background()
            body.shadow.inherit = False
        except Exception:
            _wrect(slide, cx - 11, cy + 1, 22, 14, W)


def _render_value_chain_native(slide, data):
    """Creative value-chain flow - colour-coded icon nodes on a connector, with
    the stage name and value-add % above, a stakeholder card below each, and a
    value-added read-out.  Mirrors components.value_chain_view."""
    stages = data["stages"]
    colors = data.get("colors", ["#1668A8"] * len(stages))
    n = len(stages) or 1
    if data.get("subtitle"):
        sb = _textbox(slide, 30, 86, 1120, 18)
        _set_text(sb.text_frame, [[(data["subtitle"], 9.5, False, _C_MUTED)]],
                  pad=(0, 2, 0, 0))
    MX = 40
    usable = 1280 - 2 * MX
    step = usable / n
    node_d = 84
    node_cy = 210
    name_y, name_h, pill_y = 118, 38, 264
    card_y, card_h = 298, 250
    centers = [MX + i * step + step / 2 for i in range(n)]
    # connector arrows between node centres (behind the nodes)
    for i in range(n - 1):
        ax0 = centers[i] + node_d / 2 - 4
        ax1 = centers[i + 1] - node_d / 2 + 4
        arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, E(ax0),
                                     E(node_cy - 9), E(ax1 - ax0), E(18))
        arr.fill.solid()
        arr.fill.fore_color.rgb = _rgb("#CBD5DE")
        arr.line.fill.background()
        arr.shadow.inherit = False
        arr.adjustments[1] = 0.55
    for i, st in enumerate(stages):
        cx = centers[i]
        col = _rgb(colors[i % len(colors)])
        # stage name above
        nm = _textbox(slide, cx - step / 2 + 4, name_y, step - 8, name_h)
        _set_text(nm.text_frame,
                  [[(ln, 11, True, _VC_NAVY)]
                   for ln in str(st["name"]).split("\n")],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, line=1.0)
        # node circle + halo + icon
        _circle(slide, cx, node_cy, node_d + 10, fill=_rgb("#EEF3F7"))
        _circle(slide, cx, node_cy, node_d, fill=col)
        _vc_icon(slide, st.get("icon", ""), cx, node_cy, col)
        # value-add % pill below
        pw = 110                  # holds "▲ 20–40%" on one line at 11pt
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      E(cx - pw / 2), E(pill_y), E(pw), E(24))
        pill.adjustments[0] = 0.5
        pill.fill.solid()
        pill.fill.fore_color.rgb = _rgb("#F0E8F6")
        pill.line.color.rgb = _VC_PURPLE
        pill.line.width = Pt(0.75)
        pill.shadow.inherit = False
        _set_text(pill.text_frame, [[("▲ " + st.get("valueadd", ""), 9.5, True,
                                      _VC_PURPLE)]], align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE)
        # stakeholder card
        cw = step - 16
        cxl = cx - cw / 2
        _white_panel(slide, cxl, card_y, cw, card_h)
        bar = _wrect(slide, cxl, card_y, cw, 5, col)
        tt = _textbox(slide, cxl + 9, card_y + 12, cw - 18, 42)
        _set_text(tt.text_frame, [[(st["title"], 11.5, True, _VC_NAVY)]],
                  line=1.05)
        nt = _textbox(slide, cxl + 9, card_y + 56, cw - 18, card_h - 66)
        _set_text(nt.text_frame, [[(st["note"], 10, False, _C_TEXT)]],
                  line=1.22)
    _render_summary_panel(slide, {
        "x": 40, "y": card_y + card_h + 14, "w": 1200,
        "h": 704 - (card_y + card_h + 14), "title": "Value-Added Read-out",
        "bullets": data.get("readout", [])})


def _valuechain_slides(prs, page, title, spec, values):
    chips = [values[r] for r in page["slicers"]]
    slide = _base_slide(prs, title, chips)
    _render_value_chain_native(slide, spec["valuechain"])
    _footer(slide)
    slide.notes_slide.notes_text_frame.text = \
        _page_note(page)
    return [slide]


# ── fit every page to its slide (locked exports) ─────────────────────
# A page is laid out in WHOLE rows, and a row taller than the slide - a
# 33-row cross-segment table - is placed anyway and runs off the bottom.
# While `fit_to_slide` is on, such a table is cut into row chunks that each
# fit a slide, the chunks after the first flowing onto numbered "(1/2)", "(2/2)" slides,
# and a segment-outlook page with more sub-segments than its matrix can label
# is cut by columns the same way (see _segstart_slides).  Nothing is dropped:
# every row, column and total is still drawn, only spread over more slides.
# The locked exports switch it on (locking.export_pptx_locked); the unlocked
# ones are built exactly as before.
# 2026-09-11: EVERY deck is now built this way, the open exports included -
# at the 11pt content standard a long table no longer fits where a 7pt one
# did, and overflowing the slide is exactly what the client asked to stop.
# The locked exports enter `fit_to_slide(locked=True)`, which only changes
# one thing: a locked table is never squeezed to stay whole (the padlock
# pass re-grows squeezed rows), it continues on the next slide instead.
_FIT_TO_SLIDE = {"on": True, "locked": False}

@contextmanager
def fit_to_slide(on: bool = True, locked: bool = True):
    """Build decks with every oversize page split to fit its slides."""
    was = dict(_FIT_TO_SLIDE)
    _FIT_TO_SLIDE.update(on=on, locked=locked)
    try:
        yield
    finally:
        _FIT_TO_SLIDE.clear()
        _FIT_TO_SLIDE.update(was)

# The client's rule for a table that overruns its slide (2026-09-11):
#   * _SQUEEZE_ROWS rows over or fewer - it stays on this slide, its rows
#     giving up part of their breathing room (never their text) to fit
#   * more than that - it continues on a new slide, numbered "(1/2)", "(2/2)"
# A continuation always carries MORE than _SQUEEZE_ROWS rows, so no slide is
# ever added for two or three rows; a table that short moves whole instead.
_SQUEEZE_ROWS = 4
_MIN_SPLIT_ROWS = _SQUEEZE_ROWS + 1

def _pptlock_mod():
    """The padlock module, whose row sizing a locked table is drawn with -
    or None, in which case the layout's own width model stands in."""
    try:
        from .locking import _pptlock
        return _pptlock()
    except Exception:
        return None

def _col_label(colname):
    """A column header as _render_table prints it (years as 2021A/E/F)."""
    try:
        y = int(str(colname).strip())
    except ValueError:
        return str(colname)
    if not 2000 <= y <= 2100:
        return str(colname)
    return f"{y}{'A' if y <= config.HIST_END else 'E' if y == config.EST_YEAR else 'F'}"

def _table_geo(block, w):
    """A table panel of width ``w`` as _render_table draws it, in px:
    (top, header row, [body rows], total row, bottom, tight), ``tight`` being
    (header, [body], total) at the rows' bare text height - as far as a
    squeezed table (`_split_table`) may shrink.

    ``top`` is _panel_frame's title strip and gap, ``bottom`` its bottom pad
    plus the source caption.  The rows come from the model the table is
    drawn with (`_table_layout`), which matches the heights the padlock pass
    writes into a locked table - so a split sized here fills its slide to
    the foot in the open deck and the locked one alike."""
    df = block["df"]
    title = _clean_title(block.get("title") or "")
    if title:
        lines = max(1, min(3, layout._wrap_count(title, max(1, w - 24),
                                                 12 * _PPT_FONT_K)))
        top = _TITLE_H + (lines - 1) * 18 + 8 - _TBL_TITLE_TUCK
        content_w, pad_b = w - 2, 6.0
    else:
        top, content_w, pad_b = 8, w - 2, 8.0
    tot = block.get("total_row")
    lay = _table_layout(df, block.get("index_name", ""), tot,
                        max(40.0, content_w))
    n, has_tot = len(df), tot is not None
    hs, ts = lay["heights"], lay["tight"]
    return (top, hs[0], hs[1:1 + n], hs[-1] if has_tot else 0.0,
            pad_b + _TBL_CAPTION,
            (ts[0], ts[1:1 + n], ts[-1] if has_tot else 0.0))

def _table_need(geo, a, b, with_total, tight=False):
    """Panel height showing body rows a..b (and the total row); ``tight``
    measures the rows squeezed to their bare text height."""
    top, head, body, total, bottom, squeezed = geo
    if tight:
        head, body, total = squeezed
    return int(math.ceil(top + head + sum(body[a:b])
                         + (total if with_total else 0) + bottom))

def _split_table(block, geo, room, fresh):
    """(head, tail): as many rows of ``block`` as fill ``room`` px, and the
    rest - with the total row - to continue on the next slide.

    Returns (block, None) when it all fits - including a table that overruns
    by no more than _SQUEEZE_ROWS rows and fits once its rows are squeezed
    (open decks only; `_render_table` does the squeezing) - and (None, None)
    when the slide already holds something and only a stub would fit here,
    or the table is too short to be worth splitting: it then starts on the
    next slide.  A continuation always carries _MIN_SPLIT_ROWS rows or more."""
    df, tot = block["df"], block.get("total_row")
    n = len(df)

    def fits(k, tight=False):
        return _table_need(geo, 0, k, k == n and tot is not None,
                           tight) <= room

    k = 0
    while k < n and fits(k + 1):
        k += 1
    if k >= n:
        return block, None
    if (not _FIT_TO_SLIDE.get("locked") and n - k <= _SQUEEZE_ROWS
            and fits(n, tight=True)):
        return block, None                   # a few rows over: squeeze them in
    if not fresh and (k < _MIN_SPLIT_ROWS or n < 2 * _MIN_SPLIT_ROWS):
        return None, None
    if n > _MIN_SPLIT_ROWS:
        k = min(k, n - _MIN_SPLIT_ROWS)       # the continuation is never a stub
    k = max(1, k)
    return ({**block, "df": df.iloc[:k], "total_row": None, "_chunk": True},
            {**block, "df": df.iloc[k:], "total_row": tot, "_chunk": True})

# ── how tall a text card is at the 11pt content standard ─────────────
# Mirrors _render_summary_panel / _render_writeup and the line model of
# _fit_text_fs: 11pt on 1.1 line spacing with 5pt after each paragraph, or
# the 1.0 / 3pt "compact" spacing it falls back to when a card is tight.
# The most a chart gives up to keep a page together.  Raised from 0.12 after
# measuring the page it decides: "Absolute $ Opportunity" carries two
# waterfalls over a third and a per-sub-segment table, and needs 574 of the
# 603px it has.  Once a segment runs to about fourteen sub-segments that
# table drives the second row past the foot, and at 0.12 the whole page could
# squeeze to only 610px - seven short - so it split, putting the two
# waterfalls on separate slides.  At 0.18 it comes to 594 and holds together.
# A 280px chart giving up 18% still stands 230px tall, which reads fine.
_FIG_SQUEEZE = 0.18

def _paras_h(chunks, text_w, compact=False):
    """Height in px of the paragraphs ``chunks`` set at 11pt in a column
    ``text_w`` px wide."""
    fs = MIN_CONTENT_PT
    line, sp = (1.0, 3.0) if compact else (1.1, 5.0)
    line_px = fs * line * 1.333 * _PPT_LINE_NAT
    total = 0.0
    for c in chunks:
        # measured in the deck's own face (layout's model is Segoe UI, ~8%
        # wider than Calibri, and split cards whose text fitted), less 3%
        # because PowerPoint breaks a line a little before the frame's edge
        n = _wrap_px(str(c), max(20.0, text_w) * 0.97, fs)
        total += n * line_px + sp * 1.333
    return total + 4.0

def _sumpanel_need(bullets, w, compact=False):
    """Height of a summary card ``w`` px wide holding ``bullets`` at 11pt
    (its title strip and rule are the 51px _render_summary_panel keeps)."""
    return int(math.ceil(51 + _paras_h(list(bullets or [""]), w - 44,
                                       compact)))

def _writeup_need(text, w, compact=False):
    """Height of a write-up strip ``w`` px wide holding ``text`` at 11pt."""
    return int(math.ceil(12 + _paras_h([text], w - 26, compact)))

def _block_need(b, w):
    """Height a non-table block needs: its template height, or more where
    its text needs it at the 11pt content standard."""
    h = layout._block_h(b, w)
    if b["kind"] == "sumpanel":
        return max(h, _sumpanel_need(b.get("bullets"), w))
    if b["kind"] == "writeup":
        return max(h, _writeup_need(b.get("text", ""), w))
    return h

def _block_give(b, w, geo, row_h):
    """How much height a block can yield when its page is a little short,
    before anything moves to a new slide: a chart up to _FIG_SQUEEZE of
    itself, a table the breathing room of its rows (open decks only - the
    padlock pass re-grows squeezed rows), a text card its paragraph
    spacing.  Anything else holds its height."""
    kind = b["kind"]
    if kind == "figure":
        return row_h * _FIG_SQUEEZE
    if kind == "table" and geo is not None:
        if _FIT_TO_SLIDE.get("locked"):
            return 0.0
        n, tot = len(b["df"]), b.get("total_row") is not None
        return max(0.0, _table_need(geo, 0, n, tot)
                   - _table_need(geo, 0, n, tot, tight=True))
    if kind == "sumpanel":
        return max(0.0, _block_need(b, w)
                   - _sumpanel_need(b.get("bullets"), w, compact=True))
    return 0.0

def _split_sumpanel(block, w, room, fresh):
    """(head, tail) of a summary card that overruns ``room`` px: as many
    bullets as fit here, the rest to continue on the next slide.

    Returns (block, None) when it fits - with its spacing tightened if need
    be, the card version of squeezing a table - and (None, None) when the
    slide already holds something and fewer than two bullets would fit
    here: the card then starts on the next slide."""
    bl = list(block.get("bullets") or [])
    if _sumpanel_need(bl, w, compact=True) <= room:
        return {**block, "ch": min(room, _sumpanel_need(bl, w))}, None
    k = 0
    while k < len(bl) and _sumpanel_need(bl[:k + 1], w) <= room:
        k += 1
    if not fresh and k < 2:
        return None, None
    k = max(1, min(k, len(bl) - 1))
    return ({**block, "bullets": bl[:k], "ch": _sumpanel_need(bl[:k], w),
             "_chunk": True},
            {**block, "bullets": bl[k:], "frac": 1.0, "_chunk": True,
             "ch": _sumpanel_need(bl[k:], layout.CONTENT_W)})

def _flow_fill(rows):
    """Lay 'rows' out on slides, filling each one to its foot.

    Rows go down the slide as layout.paginate_rows places them, but a table
    reaching the foot of a slide is cut at the last row that fits and its
    remaining rows continue at the top of the next slide - so no slide is
    left with white space under a table that goes on elsewhere.  What comes
    after it follows straight on below.  A continuation that stood beside a
    chart spans the full width, rather than leave half its slide empty.
    Spare height goes to the rows that are not cut tables, so a cut table's
    panel ends where its rows do."""
    top, bottom = layout.CONTENT_TOP, layout.CONTENT_BOTTOM
    gx, gy = layout.GUTTER_X, layout.GUTTER_Y
    pages = []
    # placed: [blocks, height, is_cut, give] - ``give`` is how much height
    # the row can still yield to keep its page together (see squeeze)
    placed, raws, y = [], [], top

    def flush():
        nonlocal placed, raws, y
        if placed:
            used = sum(p[1] for p in placed) + gy * (len(placed) - 1)
            spare = max(0.0, bottom - top - used)
            # spare height goes to charts and text cards, never to a table -
            # a table is as tall as its rows; a page of tables alone keeps
            # its spare room at the foot
            takers = [i for i, p in enumerate(placed) if not p[2]
                      and not any(b["kind"] == "table" for b in p[0])]
            add = spare / len(takers) if takers else 0.0
            ny = top
            for i, (blocks, h, _c, _f) in enumerate(placed):
                nh = int(h + (add if i in takers else 0))
                for b in blocks:
                    b["y"], b["h"] = int(ny), nh
                ny += nh + gy
        if placed or raws:
            pages.append({"blocks": raws + [b for p in placed
                                            for b in p[0]],
                          "height": layout.CANVAS_H})
        placed, raws, y = [], [], top

    def place(blocks, widths, h, cut, row_raws, give=0.0):
        nonlocal y
        raws.extend(row_raws)
        x, out = layout.PAGE_LEFT, []
        for b, w in zip(blocks, widths):
            out.append({**b, "x": int(round(x)), "y": int(y), "w": int(w),
                        "h": int(h)})
            x += w + gx
        placed.append([out, h, cut, max(0.0, give)])
        y += h + gy

    def squeeze(row_h, row_give):
        """Make room for a row that overruns the foot by a little: the rows
        already on the page - and this one - yield what they can, in
        proportion (charts up to _FIG_SQUEEZE of their height, tables the
        breathing room of their rows, text cards their paragraph spacing;
        see _block_give).  This is the "adjust it on the same slide" step
        that comes before any new slide.  Returns the row's height after the
        squeeze, or None when everything together would not be enough."""
        nonlocal y
        over = (y + row_h) - bottom
        total = sum(p[3] for p in placed) + row_give
        if over <= 0 or total <= 0 or over > total:
            return None
        k = over / total
        for p in placed:
            p[1] -= p[3] * k
            p[3] -= p[3] * k
        y = top + sum(p[1] for p in placed) + gy * len(placed)
        return row_h - row_give * k

    def pull_table_tail(row):
        """A text card about to leave only a bullet or two for the next
        slide: when the row above it is a whole table, the table's last
        rows go on ahead instead - the next slide opens with the table's
        continuation and the card follows whole, so no slide is ever added
        for two bullets.  Returns False when there is no such table."""
        nonlocal y
        if not placed or placed[-1][2] or len(placed[-1][0]) != 1:
            return False
        tb = placed[-1][0][0]
        if tb["kind"] != "table" or len(tb["df"]) < 2 * _MIN_SPLIT_ROWS:
            return False
        k = len(tb["df"]) - _MIN_SPLIT_ROWS
        head_h = _table_need(_table_geo(tb, tb["w"]), 0, k, False)
        head = {**tb, "df": tb["df"].iloc[:k], "total_row": None,
                "_chunk": True, "h": int(head_h)}
        tail = {key: v for key, v in tb.items()
                if key not in ("x", "y", "w", "h")}
        tail.update(df=tb["df"].iloc[k:], _chunk=True)   # keeps its total
        placed[-1] = [[head], head_h, True, 0.0]
        y = top + sum(p[1] for p in placed) + gy * len(placed)
        queue.insert(0, row)
        queue.insert(0, [tail])
        flush()
        return True

    queue = [list(r) for r in rows]
    while queue:
        row = queue.pop(0)
        row_raws = [b for b in row if b["kind"] == "raw"]
        real = [b for b in row if b["kind"] != "raw"]
        if not real:
            raws.extend(row_raws)
            continue
        n = len(real)
        avail_w = layout.CONTENT_W - (n - 1) * gx
        widths = [max(120, int(b.get("frac", 1.0 / n) * avail_w))
                  for b in real]
        geos = [_table_geo(b, w) if b["kind"] == "table" else None
                for b, w in zip(real, widths)]
        # text cards are measured at the 11pt content standard, not at the
        # height their template guessed for them
        hs = [(_table_need(g, 0, len(b["df"]), b.get("total_row") is not None)
               if g else _block_need(b, w))
              for b, w, g in zip(real, widths, geos)]
        room = bottom - y
        fresh = not placed
        cut = any(b.get("_chunk") for b in real)
        rh = max(hs)
        # The row shrinks as ONE - every block in it is placed at the row's
        # height - so it can come down until the first block reaches its own
        # floor: its natural height less whatever it can give.  A block
        # SHORTER than the row is being stretched to fill it and loses
        # nothing until the row reaches its natural height.
        #
        # This used to take the smallest give across the row, which is only
        # the same answer when every block is exactly as tall as the row.
        # Otherwise it threw away room that was there: one table with no give
        # of its own zeroed the whole row, even when the chart beside it
        # stood 40px taller and could have yielded 35 without the table
        # noticing.  The page then went to a second slide - and that is the
        # continuation that arrives holding five rows and half a slide of
        # white space.
        give = 0.0 if cut else max(0.0, rh - max(
            h - _block_give(b, w, g, rh)
            for b, h, w, g in zip(real, hs, widths, geos)))
        if rh <= room:
            place(real, widths, rh, cut, row_raws, give)
            continue
        sq = squeeze(rh, give)
        if sq is not None:
            place(real, widths, sq, cut, row_raws, give - (rh - sq))
            continue
        room = bottom - y
        split = [i for i, b in enumerate(real)
                 if b["kind"] in ("table", "sumpanel")]
        rest = max((hs[i] for i in range(n) if i not in split), default=0)
        if split and rest <= room:
            parts = {i: (_split_table(real[i], geos[i], room, fresh)
                         if real[i]["kind"] == "table"
                         else _split_sumpanel(real[i], widths[i], room, fresh))
                     for i in split}
            if (n == 1 and real[0]["kind"] == "sumpanel"
                    and parts[0][1] is not None
                    and len(parts[0][1].get("bullets") or []) <= 2
                    and pull_table_tail(row)):
                continue                 # the table's tail leads the next slide
            if all(p[0] is not None for p in parts.values()):
                heads = [parts[i][0] if i in parts else b
                         for i, b in enumerate(real)]
                # a block kept whole by squeezing is placed at the room it
                # has, and its renderer squeezes its rows / lines into that
                hh = []
                for i in range(n):
                    if i not in parts:
                        hh.append(hs[i])
                    elif real[i]["kind"] == "table":
                        hh.append(min(room, _table_need(
                            geos[i], 0, len(parts[i][0]["df"]),
                            parts[i][0].get("total_row") is not None)))
                    else:
                        hh.append(min(room, parts[i][0]["ch"]))
                place(heads, widths, max(hh), True, row_raws)
                tails = [parts[i][1] for i in split
                         if parts[i][1] is not None]
                if tails:
                    if len(tails) == 1 and n > 1:
                        tails = [{**tails[0], "frac": 1.0}]
                    queue.insert(0, tails)
                    flush()
                continue
        if not fresh:
            # a text card that would open a slide of its own at the very end
            # of the page: the table above sends its last rows along with it
            # instead, so the new slide is never just the card
            if (n == 1 and real[0]["kind"] == "sumpanel" and not queue
                    and pull_table_tail(row)):
                continue
            flush()                      # start the row on the next slide
            queue.insert(0, row)
            continue
        place(real, widths, rh, cut, row_raws)   # nothing smaller to do
    flush()
    return pages

def _fit_pages(spec):
    """Slides for ``spec`` with every table and text card flowed to fill
    them (see `_flow_fill`), or None when the ordinary pagination already
    fits: a table-packing page whose tables all fit, or a page that is not
    laid out in rows at all."""
    if "rows" in spec:
        rows = spec["rows"]
        return _flow_fill(rows)
    elif spec.get("tables"):
        fixed = spec.get("fixed", [])
        top = layout.CONTENT_TOP + (max(f["ch"] for f in fixed)
                                    + layout.GUTTER_Y if fixed else 0)

        # (always flowed here: the table packer sizes tables with the
        # dashboard's smaller rows, so at 11pt stacked tables ran into each
        # other - every table is placed from its real deck height instead)
        rows = ([[{**f, "frac": 1.0 / len(fixed)} for f in fixed]]
                if fixed else []) + [[{**t, "frac": 1.0}]
                                     for t in spec["tables"]]
    else:
        return None
    if not any(b.get("kind") == "table" for r in rows for b in r):
        return None
    return _flow_fill(rows)

def _flow_slides(prs, page, region, country):
    """A flow page renders to one OR MORE slides - mirrors the dashboard by
    paginating the layout spec and drawing each physical page's blocks (the
    same layout._components call the dashboard uses, under the export
    collector)."""
    values = {"region": region, "country": country}
    args = [values[r] for r in page["slicers"]]
    title, spec = page["render"](*args)
    title = geo.label(title)
    if "flow" in spec:                         # taxonomy flow chart (native)
        return _taxonomy_flow_slides(prs, page, title, spec, values)
    if "signals" in spec:                      # market-signals editorial table
        return _signals_slides(prs, page, title, spec, values)
    if "endcards" in spec:                     # segment-section closing card
        return _endcards_slides(prs, page, title, spec, values)
    if "segstart" in spec:                     # segment-section opening page
        return _segstart_slides(prs, page, title, spec, values)
    if "pricepattern" in spec:                 # per-region pricing-pattern page
        return _pricepattern_slides(prs, page, title, spec, values)
    if "primitives" in spec:                   # generic primitive-drawn page
        return _primitives_slides(prs, page, title, spec, values)
    if "primpages" in spec:                    # primitive pages (taxonomy grid)
        return _primpages_slides(prs, page, title, spec, values)
    if "compmatrix" in spec:                   # creative competitor dashboard
        return _compmatrix_slides(prs, page, title, spec, values)
    if "valuechain" in spec:                   # value-chain chevron flow
        return _valuechain_slides(prs, page, title, spec, values)
    if "slides" in spec:                       # legacy taxonomy -> one PPT slide
        return _taxonomy_one_slide(prs, page, title, spec, values)
    if "positioned" in spec:                   # fixed coordinates: its slide
        specs: list[dict] = []                 # plus numbered continuations
        components.set_collector(specs)
        try:
            layout._components(list(spec["positioned"]))
        finally:
            components.set_collector(None)
        return _spec_page_slides(prs, page, title, specs,
                                 [values[r] for r in page["slicers"]])
    doc_pages = ((_fit_pages(spec) if _FIT_TO_SLIDE["on"] else None)
                 or layout.paginate(spec))
    k = len(doc_pages)
    made = []
    for j, dp in enumerate(doc_pages, start=1):
        heading = title if k == 1 else f"{title} ({j}/{k})"
        chips = [values[r] for r in page["slicers"]] if j == 1 else ()
        # a layout-engine page already starts under the header; a
        # "positioned" page carries fixed coordinates and is remapped
        slide = _base_slide(prs, heading, chips,
                            remap="positioned" in spec)
        specs: list[dict] = []
        components.set_collector(specs)
        try:
            layout._components(dp["blocks"])
        finally:
            components.set_collector(None)
        _draw_specs(slide, specs)
        _footer(slide)
        note = _page_note(page)
        if k > 1:
            note += f"  ({j}/{k})"
        slide.notes_slide.notes_text_frame.text = note
        made.append(slide)
    return made

def _set_fill_alpha(shape, pct):
    """Make a shape's solid fill translucent (pct = 0..100 opacity)."""
    try:
        srgb = shape._element.spPr.find(
            qn("a:solidFill") + "/" + qn("a:srgbClr"))
        if srgb is not None:
            srgb.append(parse_xml(
                f'<a:alpha {nsdecls("a")} val="{int(pct * 1000)}"/>'))
    except Exception:
        pass

def _cover_shape(slide, kind, x, y, w, h, fill=None, line=None, lw=1.75,
                 alpha=None, rot=0):
    shp = slide.shapes.add_shape(kind, E(x), E(y), E(w), E(h))
    shp.shadow.inherit = False
    if rot:
        shp.rotation = rot
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(fill)
        if alpha is not None:
            _set_fill_alpha(shp, alpha)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = _rgb(line)
        shp.line.width = Pt(lw)
    return shp

def _title_slide(prs, region, country):
    """Decorative cover - layered blue background, concentric brand rings, a
    subtle ascending-bar motif; the title stands alone (no accent rules)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _cover_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 1280, 720, fill=COVER_BG)
    # soft glow bottom-left + concentric rings top-right (echo the FMI globe)
    _cover_shape(slide, MSO_SHAPE.OVAL, -200, 470, 560, 560, fill=COVER_GLOW,
                 alpha=45)
    for d in (560, 410, 270):
        _cover_shape(slide, MSO_SHAPE.OVAL, 1120 - d / 2, 70 - d / 2, d, d,
                     line=COVER_RING, lw=1.5)
    _cover_shape(slide, MSO_SHAPE.OVAL, 1120 - 60, 70 - 60, 120, 120,
                 fill=COVER_DOT, alpha=22)
    # ascending bar motif, bottom-right (a nod to the report's charts)
    for i, hh in enumerate((26, 44, 34, 60, 82, 52, 96)):
        _cover_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 946 + i * 40,
                     640 - hh, 24, hh, fill=COVER_BAR, alpha=70)
    # (the thin accent bar beside the title and the short rule under it were
    #  removed 2026-09-07 - the cover carries the heading text alone)
    # Fact.MR wordmark on a white card, top-left of the cover
    if _LOGO_PATH.exists():
        lh = 54.0
        lw = lh * _LOGO_AR
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, E(96), E(64),
                                      E(lw + 22), E(lh + 18))
        card.adjustments[0] = 0.14
        card.fill.solid()
        card.fill.fore_color.rgb = _C_WHITE
        card.line.fill.background()
        card.shadow.inherit = False
        try:
            slide.shapes.add_picture(str(_LOGO_PATH), E(96 + 11), E(64 + 9),
                                     height=E(lh), width=E(lw))
        except Exception:
            pass
    tb = _textbox(slide, 120, 280, 980, 190)
    _set_text(tb.text_frame,
              [[(f"{geo.word()} {data.MARKET_LABEL}", 34, True,
                 _rgb(COVER_TEXT))]],
              anchor=MSO_ANCHOR.MIDDLE)
    tb2 = _textbox(slide, 122, 492, 980, 90)
    _set_text(tb2.text_frame, [
        [(f"Value ({UNIT_LABELS['Value']}) and Volume "
          f"({UNIT_LABELS['Volume']}) Analysis and Forecast, "
          f"{YEARS[0]}–{YEARS[-1]}", 16, False, _rgb(COVER_SUB))],
        [(f"Region focus: {region}   ·   Country focus: {country}   ·   "
          f"Generated {datetime.now():%d %b %Y}   ·   {SOURCE_TEXT}",
          10, False, _rgb("#8FCBE6"))],
    ])
    _footer(slide, dark=not COVER_ON_LIGHT, logo=False)
    return slide

def _menu_slide(prs):
    """The Contents slide (slide 1) - light-blue page + teal FMI header band.
    The clickable navigation buttons are drawn by _render_contents_ppt once
    every target slide exists (so internal hyperlinks resolve correctly)."""
    # the "click a section to jump" hint now lives in the sub header
    return _base_slide(prs, "Agenda", ())

def _render_contents_ppt(slide, items, slug_of):
    """Draw the shared button-grid items as native teal navigation cards, each
    carrying an internal slide hyperlink resolved from the slug -> slide map."""
    for it in items:
        btn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     E(it["x"]), E(it["y"]),
                                     E(it["w"]), E(it["h"]))
        btn.adjustments[0] = min(0.5, 8.0 / max(1, it["h"]))
        btn.fill.solid()
        btn.fill.fore_color.rgb = _rgb(TOC_TILE_BG)
        btn.line.color.rgb = _C_EDGE
        btn.line.width = Pt(0.5)
        btn.shadow.inherit = False
        btn.text_frame.word_wrap = True
        _set_text(btn.text_frame, [[(it["label"], 13, True,
                                     _rgb(TOC_TILE_TEXT))]],
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        tgt = slug_of.get(it.get("slug"))
        if tgt is not None:
            try:
                btn.click_action.target_slide = tgt
            except Exception:
                pass
    _footer(slide)

# ══════ Table of Contents / List of Figures / List of Tables ════════
# The research team writes the contents in Word and used to paste them into
# the deck by hand: thirty rows to a table, two tables to a slide, with "Keep
# Source Formatting" so Word's numbering, indent and colour survived.
# `toc.py` reads that file and resolves the numbering (the numbers are fields,
# not text); what follows lays the result out the same way, inside THIS
# deck's chrome - its header rule, its logo, its footer, its section divider.
_TOC_ROWS = 30                    # rows per table - the team's own unit
_TOC_TABLES = 2                   # tables per slide, side by side
_TOC_PER_SLIDE = _TOC_ROWS * _TOC_TABLES
_TOC_ROW_H = 18.4                 # px - the 0.19" row of the reference deck
# The reference deck's table style: no fill and no gridlines, just a rule
# above and below the block, which takes its colour from the theme - so on
# this deck it comes out in the FMI accent rather than the Word original's.
_TOC_TBL_STYLE = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"
_TOC_SHAPE = "FMI Contents Table"  # marks the tables the 11pt floor skips
_TOC_INSET = 9525                 # EMU - the cell's other three insets
_LIST_PT = 11.0                   # List of Figures / Tables: Calibri 11
_SECTION_BAR = "FMI Section Bar"   # the 'Section NN' bar, so it can renumber

TOC_TITLE = "Table of Contents"
LIST_TITLE = "List of Figures, List of Tables"


def _toc_cell(cell, runs, marL, size):
    """One contents row: a cell of plain runs at Word's own size and colour."""
    cell.margin_left = Emu(int(marL))
    cell.margin_right = Emu(_TOC_INSET)
    cell.margin_top = Emu(_TOC_INSET)
    cell.margin_bottom = Emu(0)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    # PowerPoint ignores an EMPTY run's formatting, falls back to its 18pt
    # default and grows the row - which on a part-filled last table pushed
    # the blank tail rows off the slide.  Pinning the paragraph's own
    # end-of-line size keeps a blank row exactly as tall as a full one.
    p._p.get_or_add_endParaRPr().set("sz", str(int(round(size * 100))))
    for text, font, sz, bold, italic, color in runs:
        r = p.add_run()
        r.text = text
        f = r.font
        f.name = font
        f.size = Pt(sz)
        f.bold = bold
        f.italic = italic
        f.color.rgb = RGBColor.from_string(color)


def _toc_row_cost(row, w) -> int:
    """How many 18.4px row slots a contents line fills in a `w`-px column:
    1, or 2 when its text wraps.  Greedy word wrap measured on the run's own
    font file (Calibri), the advance PowerPoint lays the text out on."""
    runs, marL, _ = row
    if not runs:
        return 1
    text = "".join(r[0] for r in runs)
    face, size = runs[0][1], runs[0][2]
    font_px = size * 96.0 / 72.0
    avail = w - (marL + _TOC_INSET) / _EMU_PER_PX - 3
    lock = _pptlock_mod()

    def tw(t):
        try:
            return lock._tw(t, font_px, face=face)
        except Exception:
            return len(t) * font_px * 0.48
    lines, cur = 1, ""
    for word in text.split():
        trial = f"{cur} {word}" if cur else word
        if cur and tw(trial) > avail:
            lines, cur = lines + 1, word
        else:
            cur = trial
    return min(2, lines)


def _toc_columns(rows, w):
    """Split contents rows into columns of `_TOC_ROWS` slots, a wrapping line
    taking two, so a long figure title can never push the table off the
    slide.  Returns [[(row, cost), ...], ...]."""
    cols, cur, used = [], [], 0
    for row in rows:
        cost = _toc_row_cost(row, w)
        if used + cost > _TOC_ROWS and cur:
            cols.append(cur)
            cur, used = [], 0
        cur.append((row, cost))
        used += cost
    if cur:
        cols.append(cur)
    return cols


def _toc_table(slide, x, y, w, rows):
    """A 30-slot, single-column contents table.

    `rows` is [(row, cost)]: a line whose text wraps is one table row two
    slots tall.  Short lists leave their tail slots as blank rows rather than
    shrinking the table, so the last page's two columns line up with every
    page before it.
    """
    blanks = max(0, _TOC_ROWS - sum(c for _, c in rows))
    n = len(rows) + blanks
    gf = slide.shapes.add_table(n, 1, E(x), E(y), E(w),
                                E(_TOC_ROWS * _TOC_ROW_H))
    gf.name = _TOC_SHAPE
    tbl = gf.table
    tbl.first_row = False
    tbl.horz_banding = False
    style = tbl._tbl.tblPr.find(qn("a:tableStyleId"))
    if style is None:
        style = parse_xml(f'<a:tableStyleId {nsdecls("a")}/>')
        tbl._tbl.tblPr.append(style)
    style.text = _TOC_TBL_STYLE
    tbl.columns[0].width = E(w)
    for i in range(n):
        if i < len(rows):
            (runs, marL, size), cost = rows[i]
        else:
            (runs, marL, size), cost = ([], _TOC_INSET, _LIST_PT), 1
        # One line tall, whatever `cost` says: a row height in PowerPoint is
        # a MINIMUM, so a line that really wraps grows its own row, and one
        # that fits (PowerPoint's metrics, or a table widened by hand) stays
        # a single line with no dead space above and below.  `cost` only
        # paginates, so a column can never run off the slide.
        tbl.rows[i].height = E(_TOC_ROW_H)
        _toc_cell(tbl.cell(i, 0), runs, marL, size)
    return gf


def _toc_entry_row(entry):
    """(runs, marL, size) for one line of the Word contents.

    Three runs, as the paste made them: the number in its OWN face - a Symbol
    bullet is a private-use codepoint and comes out as a hollow box in any
    other font - then the padding that stands in for the tab, then the text.
    """
    runs = []
    style = (entry.size, entry.bold, entry.italic, entry.color)
    if entry.marker:
        runs.append((entry.marker, entry.marker_font) + style)
        pad = entry.tab_pad()
        runs.append((" " * pad + " ", toc._PAD_FONT, toc._PAD_SIZE,
                     entry.bold, entry.italic, entry.color))
    runs.append((entry.text, "Calibri") + style)
    return runs, entry.marL, entry.size


def _listing_row(text):
    """(runs, marL, size) for a List of Figures / Tables line: flush left,
    Calibri 11, black - no numbering to indent under."""
    return ([(text, "Calibri", _LIST_PT, False, False, "000000")],
            _TOC_INSET, _LIST_PT)


def _contents_col_w():
    """Width of one of the two side-by-side contents tables, in px."""
    return (layout.CONTENT_W - layout.GUTTER_X) / 2.0


def _contents_slides(prs, title, rows, note):
    """Paginate `rows` over as many two-table slides as they need."""
    if not rows:
        return []
    w = _contents_col_w()
    xs = (layout.PAGE_LEFT, layout.PAGE_LEFT + w + layout.GUTTER_X)
    cols = _toc_columns(rows, w)
    pages = [cols[i:i + _TOC_TABLES]
             for i in range(0, len(cols), _TOC_TABLES)]
    made = []
    for j, page in enumerate(pages, start=1):
        heading = title if len(pages) == 1 else \
            f"{title} ({j}/{len(pages)})"
        slide = _base_slide(prs, heading, (), remap=False)
        for t, x in enumerate(xs):
            _toc_table(slide, x, layout.CONTENT_TOP, w,
                       page[t] if t < len(page) else [])
        _footer(slide)
        slide.notes_slide.notes_text_frame.text = (
            f"{note}  ({j}/{len(pages)})" if len(pages) > 1 else note)
        made.append(slide)
    return made


def _contents_section(prs, title, groups, bullets):
    """A section front page plus its contents slides.

    `groups` is [(page title, rows)] - one for the contents, two for the
    figure and table lists, which share a section but head their own pages.
    This is what `emit` does for a registry page, done by hand because these
    sections are assembled after the deck they describe.
    """
    slides = []
    for page_title, rows in groups:
        slides += _contents_slides(prs, page_title, rows,
                                   f"{geo.label(MAIN_REPORT)} {page_title}")
    if not slides:
        return []
    # numbered 0 here and renumbered by `_renumber_sections` once the deck
    # is in its final order
    div = _section_divider_slide(prs, 0, title, bullets,
                                 kicker=geo.label(MAIN_REPORT))
    return [div] + slides


def _move_slides(prs, movers, index):
    """Move `movers`, in order, so they sit at `index` in the finished deck.

    They are built last - the List of Figures cannot be written until the
    deck has drawn its figures - but they belong early, so the slide-id list
    is re-ordered once they exist.  Every mover is at the END of the deck,
    so removing them leaves the positions before `index` untouched.
    """
    lst = prs.slides._sldIdLst
    # Keyed on each slide's own XML element, not on the Slide object:
    # python-pptx builds those on demand, so two walks of prs.slides need not
    # hand back the same instance and an identity map can silently come up
    # empty - leaving the pages built but never moved.
    pos = {s._element: e for s, e in zip(prs.slides, list(lst))}
    moving = [pos[s._element] for s in movers if s._element in pos]
    for el in moving:
        lst.remove(el)
    for off, el in enumerate(moving):
        lst.insert(index + off, el)


def _renumber_sections(prs):
    """Rewrite every 'Section NN' bar to its place in the FINAL order.

    The dividers are numbered as they are built; inserting two sections in
    the middle afterwards would leave the rest counting from the wrong
    place, so the bars are renumbered once the deck is in its final order.
    """
    n = 0
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.name != _SECTION_BAR or not sh.has_text_frame:
                continue
            n += 1
            for para in sh.text_frame.paragraphs:
                for i, run in enumerate(para.runs):
                    run.text = f"Section {n:02d}" if i == 0 else ""


TOC_SLUG = "main-table-of-contents"
LIST_SLUG = "main-list-of-figures-tables"


def _build_contents(prs, slides, anchor, slug_of_slide, log):
    """Build the two contents sections and drop them in after `anchor`.

    Splices them into `slides` so the Prev/Next buttons thread through them,
    and returns their Agenda cards so the menu lists them where they sit.
    """
    made, buttons = [], []

    def section(title, slug, groups, bullets):
        block = _contents_section(prs, title, groups, bullets)
        if not block:
            return
        made.extend(block)
        slug_of_slide[slug] = block[1]      # [0] is the section front page
        buttons.append({"label": title, "slug": slug})

    src = toc.find_toc_file()
    if src is None:
        log("table of contents: no TOC_*.docx found; section skipped")
    else:
        try:
            entries = toc.parse_toc(src)
        except Exception as exc:            # a stray .docx, a file Word holds
            log(f"table of contents: could not read {src.name} ({exc})")
            entries = []
        if entries:
            log(f"table of contents: {len(entries)} lines from {src.name}")
            # No COUNT in a divider bullet: a locked deck withholds figures
            # wherever it finds them, and "351 lines" came out as "🔒 lines"
            # - a padlock over a number that was never a market figure.
            section(TOC_TITLE, TOC_SLUG,
                    [(TOC_TITLE, [_toc_entry_row(e) for e in entries])],
                    ["The full contents of this study, as delivered"])
        else:
            log(f"table of contents: no numbered lines in {src.name}")

    # The lists describe the full study (weighted exhibit plan, see
    # exhibits.py); the deck's own captions are the fallback if the plan
    # cannot be drawn up from this workbook.
    try:
        figs, tabs = exhibits.figure_entries(), exhibits.table_entries()
    except Exception as exc:
        log(f"list of figures/tables: exhibit plan failed ({exc}); "
            f"listing the deck's own charts and tables instead")
        figs, tabs = toc.figure_entries(), toc.table_entries()
    groups, bullets = [], []
    if figs:
        groups.append(("List of Figures", [_listing_row(t) for t in figs]))
        bullets.append("Every figure in the full study, by analysis block")
    if tabs:
        groups.append(("List of Tables", [_listing_row(t) for t in tabs]))
        bullets.append("Every table in the full study, by analysis block")
    if groups:
        log(f"list of figures/tables: {len(figs)} figures, {len(tabs)} tables")
        section(LIST_TITLE, LIST_SLUG, groups, bullets)

    if made:
        _move_slides(prs, made, 2 + anchor)     # 2 = cover + Agenda menu
        slides[anchor:anchor] = made
    return buttons


def _agenda_buttons(extra):
    """The Agenda's cards, with the contents sections slotted in where their
    slides are - straight after the Market Taxonomy pages, so the menu
    reads in the order the deck actually runs."""
    buttons = list(contents_buttons())
    if not extra:
        return buttons
    # after the LAST taxonomy page (Company Coverage), so the whole Market
    # Taxonomy block reads through before the contents begin
    tax = {p["slug"] for p in PAGES if p["report"] == MAIN_REPORT
           and p["group"] == "Market Taxonomy"}
    at = max((i + 1 for i, b in enumerate(buttons) if b["slug"] in tax),
             default=len(buttons))
    return buttons[:at] + list(extra) + buttons[at:]


def _divider_slide(prs, title, subtitle=""):
    """A full-bleed section-title slide (heads each country in a region deck)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = _C_HEADER
    bg.line.fill.background()
    bg.shadow.inherit = False
    tb = _textbox(slide, 80, 300, 1120, 130)
    _set_text(tb.text_frame, [[(title, 34, True, _C_WHITE)],
                              [(subtitle, 16, False, _rgb(ON_DARK_MUTED))]])
    _footer(slide, dark=True, logo=False)
    return slide

_SECTION_BAND = _rgb(SECTION_BAND)                 # 'Section N' bar

def _section_divider_slide(prs, section_no, title, bullets, kicker=""):
    """A section front page (export-only) modelled on the reference deck: a big
    blue panel with the section title + italic bullet blurb, and a yellow
    'Section N' bar along the bottom.  Not shown in the dashboard."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    # light blue from the template's palette (was white until 2026-09-25)
    bg.fill.fore_color.rgb = _rgb(getattr(theming.theme(), "DIVIDER_BG",
                                          PAGE_BG))
    bg.line.fill.background()
    bg.shadow.inherit = False
    # left blue panel
    panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(96), E(24),
                                   E(760), E(612))
    panel.fill.solid()
    # the panel used to be the brand colour outright, which is what made a
    # divider half a page of solid navy under the FMI template
    panel.fill.fore_color.rgb = _rgb(DIVIDER_PANEL)
    panel.line.fill.background()
    panel.shadow.inherit = False
    # the enlarged wordmark no longer fits the 96px margin left of the panel,
    # so on a divider it sits inside the panel's top-left corner
    _top_left_logo(slide, x=128, y=52)
    # stack kicker -> title (may wrap to 2 lines) -> bullets as one block,
    # vertically centred in the panel, so nothing overlaps at any title length
    tw = 700
    tlen = len(str(title))
    title_fs = 30 if tlen <= 30 else (25 if tlen <= 44 else 21)
    cpl = max(8, int((tw - 6) / (title_fs * 0.60)))     # chars per line
    title_lines = min(2, max(1, int(math.ceil(tlen / cpl))))
    tlh = title_fs * 1.34
    title_h = int(title_lines * tlh + 6)
    kick_h = 22 if kicker else 0
    gap1, gap2, bull_lh = 4, 16, 27
    bull_h = max(1, len(bullets)) * bull_lh
    block_h = kick_h + gap1 + title_h + gap2 + bull_h
    y = 24 + max(24, (612 - block_h) / 2.0)
    if kicker:
        kb = _textbox(slide, 128, y, tw, kick_h)
        _set_text(kb.text_frame, [[(kicker, 12.5, True,
                                    _rgb(DIVIDER_KICKER))]],
                  pad=(0, 2, 0, 0))
        y += kick_h + gap1
    tb = _textbox(slide, 128, y, tw, title_h)
    _set_text(tb.text_frame, [[(title, title_fs, True, _rgb(DIVIDER_TITLE))]],
              anchor=MSO_ANCHOR.TOP, line=1.02, pad=(0, 2, 0, 0))
    y += title_h + gap2
    bl = _textbox(slide, 128, y, tw, bull_h + 8)
    paras = [[("➤  ", 12, True, _rgb(DIVIDER_MARK)),
              (b, 12.5, False, _rgb(DIVIDER_BODY))] for b in bullets]
    _set_text(bl.text_frame, paras or [[("", 12, False, _rgb(DIVIDER_BODY))]],
              space_after=8, line=1.1)
    # yellow 'Section N' bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(40), E(648),
                                 E(1200), E(46))
    bar.name = _SECTION_BAR          # so `_renumber_sections` can find it
    bar.fill.solid()
    bar.fill.fore_color.rgb = _SECTION_BAND
    bar.line.fill.background()
    bar.shadow.inherit = False
    _set_text(bar.text_frame, [[(f"Section {section_no:02d}", 15, True,
                                 _rgb(SECTION_BAND_TEXT))]],
              anchor=MSO_ANCHOR.MIDDLE, pad=(28, 8, 2, 2))
    _footer(slide, dark=not COVER_ON_LIGHT, logo=False)
    return slide

# ── Section breakers ─────────────────────────────────────────────────
# A full-bleed slide in the template's darkest brand shade that opens a run
# of sections: an unnumbered title and an arrow list of what follows, with
# segment names indented under their heading.  Export-only, like the
# numbered section dividers, which still head each section after it.
_BREAKER_SHAPE = "FMI Section Breaker"


def _breaker_slide(prs, title, items):
    """`items` is [(text, level)], level 0 for a main point, 1 for a sub.

    The whole page is shaded from the template's darkest brand tone into its
    accent, with the template's logo and copyright laid over it and the
    text in white Calibri: title 30pt bold, points 20pt, sub-points 14pt."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.name = _BREAKER_SHAPE
    bg.line.fill.background()
    bg.shadow.inherit = False
    bg.fill.gradient()
    bg.fill.gradient_angle = 0                       # left -> right
    stops = [(0, theming.theme().BRAND_DK), (55000, HEADER_BG),
             (100000, ACCENT)]
    gs_lst = bg.fill._fill._gradFill.gsLst             # replace the 2 stops
    for gs in list(gs_lst):
        gs_lst.remove(gs)
    for pos, colour in stops:
        gs_lst.append(parse_xml(
            f'<a:gs {nsdecls("a")} pos="{pos}">'
            f'<a:srgbClr val="{str(colour).lstrip("#").upper()}"/></a:gs>'))
    # no header rule on a breaker (removed 2026-09-25) - just the logo
    _top_left_logo(slide)

    top, bottom = layout.CONTENT_TOP, FOOTER_RULE_Y - 4
    title_pt, main_pt, sub_pt = 30.0, 20.0, 14.0
    title_h = 56

    def list_h(m, s_):
        return sum((m if lvl == 0 else s_) * 1.6 * 96 / 72
                   for _, lvl in items)
    # only a market with a very long segment list steps the points down
    while list_h(main_pt, sub_pt) > bottom - top - title_h - 20             and sub_pt > 11:
        main_pt, sub_pt = main_pt - 1, sub_pt - 0.5
    lh = list_h(main_pt, sub_pt)
    y = top + max(20, (bottom - top - (title_h + lh)) / 2)

    def runs(para, parts):
        for text, size, bold, color in parts:
            r = para.add_run()
            r.text = text
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
            r.font.name = "Calibri"

    tb = _textbox(slide, 96, y, 1100, title_h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = E(0)
    runs(tf.paragraphs[0], [(title, title_pt, True, _C_WHITE)])
    y += title_h

    lb = _textbox(slide, 96, y, 1100, lh + 10)
    tf = lb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = E(0)
    tf.margin_top = tf.margin_bottom = E(0)
    for i, (text, lvl) in enumerate(items):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        size = main_pt if lvl == 0 else sub_pt
        para.space_after = Pt(size * 0.45)
        para.line_spacing = 1.0
        pPr = para._p.get_or_add_pPr()
        pPr.set("marL", str(E(40 + 40 * lvl)))
        pPr.set("indent", str(-E(36)))
        color = _C_WHITE if lvl == 0 else _rgb(ON_DARK_MUTED)
        mark = "➤" if lvl == 0 else "›"
        runs(para, [(mark + "	", size, False, color),
                    (text, size, False, color)])
    _footer(slide, dark=True, logo=False)
    slide.notes_slide.notes_text_frame.text = f"{title} (section breaker)"
    return slide


def _breaker_for(report, group, main_pages, country_pages):
    """(title, items) for the breaker that opens `group`, or None."""
    if report == MAIN_REPORT and group == "Market Taxonomy":
        items = [("Market Taxonomy", 0)]
        if toc.find_toc_file() is not None:
            items.append((TOC_TITLE, 0))
        items.append((LIST_TITLE, 0))
        return "Market Overview", items
    if report == MAIN_REPORT and group == "Global Overview":
        groups = list(dict.fromkeys(p["group"] for p in main_pages))
        segs = [g for g in groups if g in set(data.SEGMENTS)]
        items = [(geo.label("Global Overview"), 0)]
        if "By Region" in groups:
            items.append((geo.label("Global Market Analysis By Region"), 0))
        if segs:
            items.append((geo.label("Global Market Analysis By Segments"), 0))
            items += [(f"By {g}", 1) for g in segs]
        return geo.label("Global Market Outlook"), items
    if report == COUNTRY_REPORT and group == "Overview":
        segs = [g for g in dict.fromkeys(p["group"] for p in country_pages)
                if g != "Overview"]
        items = [("Overview (Volume & Value)", 0)]
        if segs:
            items.append(("Segment Overview (Volume & Value)", 0))
            items += [(g, 1) for g in segs]
        return "Country Market Outlook", items
    if report == DYNAMICS_REPORT:
        groups = list(dict.fromkeys(p["group"] for p in PAGES
                                    if p["report"] == DYNAMICS_REPORT))
        groups.sort(key=lambda g: g != _DYN_FORCES)   # the deck's order
        return "Market Dynamics", [(geo.label(g), 0) for g in groups]
    if report == MAIN_REPORT and group == "Pricing":
        return "Pricing Assessment", [("Market Pricing", 0)]
    return None


def _section_blurb(report, group):
    """(kicker, title, bullet-lines) for a section divider - derived from the
    section's pages so it adapts to any workbook.  Bullets list what's inside."""
    pgs = [p for p in PAGES if p["report"] == report and p["group"] == group]
    names = []
    for p in pgs:                                  # dedupe, keep order
        nm = p["name"]
        if nm not in names:
            names.append(nm)
    kicker = geo.label(report)
    bullets = [geo.label(n) for n in names[:5]] if names else [geo.label(group)]
    return kicker, geo.label(group), bullets

def export_pptx_native(out_path: str | Path, region: str = "Global",
                       country: str | None = None,
                       reports: tuple = ("main", "country", "dynamics",
                                         "competitor"),
                       countries=None, progress=None,
                       section_dividers: bool = True,
                       template: str | None = None) -> Path:
    """Build one deck: the main report for ``region`` + the country report for
    every country in ``countries`` (default: just ``country``).

    ``template`` selects the visual template - "factmr" or "fmi".  It is
    applied for the duration of the build and restored afterwards, so a deck
    can be generated in one template while the dashboard stays on the other.
    Passing None keeps whatever is active, which is what every existing caller
    gets: the Fact.MR deck is produced exactly as before.
    """
    # The whole deck is built AS this geography: every label that describes the
    # report rather than its data ("Global Pricing Assessment", the Agenda, the
    # section dividers, the speaker notes) re-points at it - see geo.py.  Set
    # here, on the validated value, so it covers the cover slide too.
    with theming.use(template) as _tpl, \
            geo.active(data.valid_geos(region, country)[0]):
        return _export_pptx_native(out_path, region, country, reports,
                                   countries, progress, section_dividers)


def _export_pptx_native(out_path: str | Path, region: str = "Global",
                        country: str | None = None,
                        reports: tuple = ("main", "country", "dynamics",
                                          "competitor"),
                        countries=None, progress=None,
                        section_dividers: bool = True) -> Path:
    """The deck builder itself.  Always runs with a template already applied
    (see `export_pptx_native`)."""
    region, country = data.valid_geos(region, country)
    log = progress or (lambda msg: None)
    out_path = Path(out_path)
    # The Agenda/Contents page is realised as the native menu slide below
    # (with internal hyperlinks), so it must be skipped in the ordinary page
    # loop - otherwise it emits a BLANK slide (its content is Dash-only) plus
    # a redundant section divider.  Matched on HOME_SLUG, not on a group name:
    # this filter used to read `group != "Contents"` and silently stopped
    # working the moment that group was renamed to "Agenda".
    def _deck_pages(report, key):
        return [p for p in PAGES if p["report"] == report and key in reports
                and p["slug"] != HOME_SLUG]

    main_pages = _deck_pages(MAIN_REPORT, "main")
    country_pages = _deck_pages(COUNTRY_REPORT, "country")
    dynamics_pages = _deck_pages(DYNAMICS_REPORT, "dynamics")
    competitor_pages = _deck_pages(COMPETITOR_REPORT, "competitor")
    ctys = [c for c in (countries or [country]) if c] if country_pages else []

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    _CONTENT_MARK.clear()                    # no stale marks from a failed build
    # slide 1 = decorative cover / front page, then slide 2 = Contents
    _title_slide(prs, region, country)
    home = next((p for p in PAGES if p["slug"] == HOME_SLUG), None)
    _SUBHEADER["text"] = (subheaders.for_page(home, region, country)
                          if home else "")
    menu = _menu_slide(prs)
    _SUBHEADER["text"] = ""

    slides = []
    slug_of_slide = {}                       # page slug -> its first PPT slide
    _ENDCARD_LINKS.clear()                    # reset per-export deferred links
    _sect = {"n": 0, "seen": set(),          # section-divider state
             "breakers": set()}              # breakers already placed
    # Where the contents sections go, once they can be built: straight after
    # the last Market Taxonomy page.  -1 until one has been emitted.
    _anchor = {"at": -1}
    toc.start_capture()                      # collect figure/table captions

    def emit(page, reg, ctry):
        # export-only section front page on the FIRST page of each (report,
        # group) - once per section, even if a group is split in page order
        if section_dividers:
            key = (page["report"], page["group"])
            if key not in _sect["seen"]:
                _sect["seen"].add(key)
                brk = _breaker_for(page["report"], page["group"],
                                   main_pages, country_pages)
                if brk and brk[0] not in _sect["breakers"]:
                    _sect["breakers"].add(brk[0])
                    slides.append(_breaker_slide(prs, *brk))
                _sect["n"] += 1
                kicker, title, bullets = _section_blurb(page["report"],
                                                        page["group"])
                slides.append(_section_divider_slide(
                    prs, _sect["n"], title, bullets, kicker=kicker))
        _SUBHEADER["text"] = subheaders.for_page(page, reg, ctry)
        try:
            made = (_flow_slides(prs, page, reg, ctry) if page.get("flow")
                    else _page_slide(prs, page, reg, ctry))
        finally:
            _SUBHEADER["text"] = ""
        slides.extend(made)
        if made:
            slug_of_slide.setdefault(page["slug"], made[0])
        # moves on with each taxonomy page, so it ends up after the last one
        # (Company Coverage) and the contents open straight after it
        if (page["report"], page["group"]) == (MAIN_REPORT, "Market Taxonomy"):
            _anchor["at"] = len(slides)
        return made[0] if made else None

    multi = len(ctys) > 1

    def emit_countries():
        # the breaker opens the whole country block, ahead of the first
        # country's own divider in a multi-country deck
        brk = _breaker_for(COUNTRY_REPORT, "Overview", main_pages,
                           country_pages)
        if section_dividers and ctys and country_pages and brk:
            _sect["breakers"].add(brk[0])
            slides.append(_breaker_slide(prs, *brk))
        for cty in ctys:
            if multi:
                div = _divider_slide(prs, cty, COUNTRY_REPORT)
                slides.append(div)
            for page in country_pages:
                emit(page, region, cty)

    # Market Dynamics runs Forces first, then Driving Factors, and sits
    # between Pricing and Supply Chain.
    dynamics_pages = sorted(dynamics_pages,
                            key=lambda p: p["group"] != _DYN_FORCES)

    def emit_dynamics():
        for page in dynamics_pages:
            emit(page, region, country)

    # The country report sits inside the main report, just ahead of Pricing
    # (or Supply Chain, when a workbook has no prices) - or at its end if
    # it has neither.  Dynamics goes in just ahead of Supply Chain, or at
    # the end of the main report without one.
    pending, dyn_pending = True, True
    for page in main_pages:
        if pending and page.get("section") in (SEC_PRICING, SEC_SUPPLY):
            emit_countries()
            pending = False
        if dyn_pending and page.get("section") == SEC_SUPPLY:
            emit_dynamics()
            dyn_pending = False
        emit(page, region, country)
    if pending:
        emit_countries()
    if dyn_pending:
        emit_dynamics()

    for page in competitor_pages:
        emit(page, region, country)

    # The contents sections come last and move into place: the List of
    # Figures cannot be written until the deck has drawn its figures.
    toc.stop_capture()
    extra_buttons = []
    if _anchor["at"] >= 0:
        extra_buttons = _build_contents(prs, slides, _anchor["at"],
                                        slug_of_slide, log)
        _renumber_sections(prs)

    # draw the dynamic, hyperlinked navigation buttons onto the Contents slide
    _render_contents_ppt(menu,
                         layout.contents_buttons_layout(
                             _agenda_buttons(extra_buttons)),
                         slug_of_slide)

    # wire each 'End of Segment' Next-pill to the next section's first slide
    for shape, slug in _ENDCARD_LINKS:
        tgt = slug_of_slide.get(slug)
        if tgt is not None:
            try:
                shape.click_action.target_slide = tgt
            except Exception:
                pass

    # (the per-slide Menu / Prev / Next buttons were removed on 2026-09-25;
    # the title now runs to the right margin instead)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    _tidy_deck(prs)
    n_raised = _enforce_content_floor(prs)      # the 11pt content standard
    if n_raised:
        log(f"content floor: {n_raised} runs raised to {MIN_CONTENT_PT:g}pt")
    _install_theme(prs)
    # python-pptx walks the package's relationship graph by RECURSION, once
    # per related part, when it saves.  A deck carrying every country's report
    # - 38 of them on the cGAS-STING workbook - puts a four-figure number of
    # parts in one file and the walk runs past CPython's 1,000-frame default,
    # dying with a RecursionError AFTER the whole deck has been built.  The
    # locked path has had this guard since 2026-08-18; the unlocked save needs
    # it for the same reason.  Raised for this call only.
    _was = sys.getrecursionlimit()
    try:
        sys.setrecursionlimit(max(_was, 10_000))
        prs.save(str(out_path))
    finally:
        sys.setrecursionlimit(_was)
    log(f"Saved: {out_path}")
    return out_path

def _safe_name(s):
    return "".join(c if (c.isalnum() or c in " -_&") else "_" for c in str(s))

def export_regions_zip(out_zip: str | Path, country: str | None = None,
                       reports: tuple = ("main", "country"),
                       progress=None, template: str | None = None,
                       regions: list | None = None) -> Path:
    """A ZIP with ONE self-contained PPT per region - each = that region's
    main report + the country report for the countries in that region
    (Global's deck covers all countries).

    ``template`` selects the visual template for every deck in the ZIP; None
    keeps whatever is active (the pre-existing behaviour).

    ``regions`` narrows the ZIP to the named regions, in the dataset's own
    order; None (the default) keeps every region, as before.  Names are
    checked against the dataset so a typo fails here rather than yielding a
    silently empty ZIP."""
    import tempfile
    import zipfile
    out_zip = Path(out_zip)
    log = progress or (lambda msg: None)
    tmpdir = Path(tempfile.mkdtemp())
    if regions is None:
        regions = list(data.GEOS_TOP)
    else:
        unknown = [r for r in regions if r not in data.GEOS_TOP]
        if unknown:
            raise ValueError(
                f"unknown region(s) {unknown}; this dataset has "
                f"{list(data.GEOS_TOP)}")
        regions = [r for r in data.GEOS_TOP if r in set(regions)]
    out_zip.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, reg in enumerate(regions, 1):
            if "country" in reports:
                ctys = (list(data.ALL_COUNTRIES) if reg == "Global"
                        else [c for c in data.COUNTRIES.get(reg, [])
                              if c in data.ALL_COUNTRIES])
            else:
                ctys = []
            log(f"  [{i}/{len(regions)}] Region: {reg} "
                f"({len(ctys)} countries)")
            tmp = tmpdir / f"{reg}.pptx"
            export_pptx_native(tmp, region=reg, country=country,
                               reports=reports, countries=ctys or None,
                               template=template)
            zf.write(tmp, f"{data.MARKET_LABEL} - {_safe_name(reg)}.pptx")
    log(f"Saved ZIP: {out_zip}")
    return out_zip


# ── Template switch support ──────────────────────────────────────────
def _retheme() -> None:
    """Re-derive every module-level constant that bakes a theme token.

    `theming.apply()` rebinds the tokens imported at the top of this module and
    then calls this hook.  RGBColor objects and the brand assets are built ONCE
    at import, so each has to be rebuilt here or the deck keeps the previous
    template's colours, logo and copyright line.
    """
    global FONT, _C_HEADER, _C_ACCENT, _C_TEXT, _C_MUTED, _C_GOOD, _C_BAD
    global _C_EDGE, _C_PAGE, _C_ZEBRA, _C_GRID, _HEAT_STOPS
    global _C_FLOW_LINE, _C_FLOW_LEAF_EDGE, _C_FLOW_LEAF_TEXT
    global _ARROW_RGB, _VC_PURPLE, _VC_NAVY, _SECTION_BAND
    global _LOGO_PATH, _LOGO_AR, _LOGO_CARD_W, _COPYRIGHT

    FONT = PPT_FONT
    _C_HEADER, _C_ACCENT = _rgb(HEADER_BG), _rgb(ACCENT)
    _C_TEXT, _C_MUTED = _rgb(TEXT), _rgb(MUTED)
    _C_GOOD, _C_BAD = _rgb(GOOD), _rgb(BAD)
    _C_EDGE, _C_PAGE = _rgb(PANEL_EDGE), _rgb(PAGE_BG)
    _C_ZEBRA, _C_GRID = _rgb(ZEBRA), _rgb(GRID)
    _HEAT_STOPS = [tuple(stop) for stop in HEAT_STOPS]
    _C_FLOW_LINE = _rgb(FLOW_LINE)
    _C_FLOW_LEAF_EDGE = _rgb(FLOW_LEAF_EDGE)
    _C_FLOW_LEAF_TEXT = _rgb(FLOW_LEAF_TEXT)
    _ARROW_RGB = {"high": ARROW_HIGH, "med": ARROW_MED, "low": ARROW_LOW}
    _VC_PURPLE, _VC_NAVY = _rgb(VC_PURPLE), _rgb(VC_NAVY)
    _SECTION_BAND = _rgb(SECTION_BAND)
    # brand assets: the wordmark, its aspect ratio and the card width that
    # decides where the slide title starts
    _LOGO_PATH = theming.logo_path()
    _LOGO_AR = theming.theme().LOGO_AR
    _LOGO_CARD_W = _LOGO_IMG_H * _LOGO_AR + 14
    _COPYRIGHT = theming.theme().COPYRIGHT
