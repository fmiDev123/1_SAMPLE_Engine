"""The Pricing Assessment slides (redesign of 2026-10-07).

Three slide types, built from the loaded workbook and drawn straight onto the
deck inside the engine's own header and footer:

  1. single-segment pricing    - price trend, 2026 index, KPI table, 3 cards
  2. cross-sectional matrix    - price by A x B cell, shaded by index, range
                                 chart; ONLY when the workbook carries the
                                 cross segment (never synthesised)
  3. country benchmarking      - price range and index across the geography's
                                 children; only when there are at least two

Average price = Value x price scale / Volume.  A parent, region or average
price is always sum(value) / sum(volume), never a mean of child prices.

The slides follow the client's reference build (`build_pricing_slides.py`);
what changed is where the numbers come from (the engine's own data layer) and
that every price figure is a plain number here - the lock pass withholds them
in a sample (pptlock: "FMI Price Chart" charts, the table and prose rules for
the pricing pages) while indices and CAGRs stay readable.
"""
from __future__ import annotations

import hashlib
import math
import random

import pandas as pd
from . import config as _config


def _vu() -> str:
    """The market's own volume unit ('Units', 'MT', 'Doses') for price text."""
    try:
        return _config.volume_noun() or "unit"
    except Exception:
        return "unit"

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (XL_CHART_TYPE, XL_LABEL_POSITION,
                             XL_LEGEND_POSITION, XL_MARKER_STYLE,
                             XL_TICK_LABEL_POSITION)
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from . import config, data, geo
from . import trend_patterns as _tp
from .data import MARKET_LABEL as MKT

# ── switches (Step 9 of the brief) ────────────────────────────────────
PRICING_PRIMARY_SEGMENT = "Product Type"   # drives the country slide
PRICING_ALL_SEGMENTS = False               # True: one slide per single segment
PRICING_MAX_CROSS = 1                      # cross matrices to build, at most

BASE, END = config.EST_YEAR, config.END_YEAR
START = config.YEARS[0]
YEARS = list(config.YEARS)

# ── palette ───────────────────────────────────────────────────────────
def _c(h):
    return RGBColor.from_string(h)

NAVY, TEAL, MUTED = _c("0B2E4A"), _c("008189"), _c("59636E")
INK, PANEL, GRID, WHITE = _c("1F1F1F"), _c("F3F6F9"), _c("D9DEE4"), _c("FFFFFF")
ORANGE, BLUE = "C55A11", "5B7DB1"
SERIES = ["C55A11", "0B2E4A", "7030A0", "008189", "9E7E0A", "5B7DB1",
          "A23B5C", "6E8B3D"]
FONT = "Calibri"
_BANDS = [(2.5, "F4B183"), (1.0, "FBE3D3"), (-1.0, "F7F7F7"),
          (-2.5, "D9E6F2")]
_LOW = "9DC0E0"

# ── geometry: the reference's inches, placed inside the engine's body ──
_TOP = 103.0                 # layout.CONTENT_TOP: under the header rule


def X(inch):
    return inch * 96.0


def Y(inch):
    return _TOP + (inch - 1.05) * 96.0


def S(inch):
    return inch * 96.0


def _emu(px):
    return Emu(int(round(px * 9525)))


# ═════════════════════════════ data ══════════════════════════════════
def _scale() -> float:
    return config.price_scale()


def _prices(geo_name, segment, level_rows=None):
    """DataFrame parents x YEARS of price (Value / Volume), NaN where there
    is no volume."""
    v = data.segment_table(geo_name, "Value", segment)
    q = data.segment_table(geo_name, "Volume", segment)
    idx = [i for i in v.index if i in q.index]
    v, q = v.loc[idx], q.loc[idx]
    p = (v[YEARS] * _scale()) / q[YEARS].where(q[YEARS] != 0)
    return p.astype(float), v, q


def _avg(v, q):
    """Volume-weighted average price per year: sum(value) / sum(volume)."""
    qs = q[YEARS].sum()
    return (v[YEARS].sum() * _scale()) / qs.where(qs != 0)


def _chartable(p) -> bool:
    ok = p.dropna(how="any")
    return len(ok) >= 2


def _top_with_others(p, v, q, keep=7):
    """More than 8 series: the top 7 by value plus an aggregated 'Others'."""
    if len(p) <= 8:
        return p, v, q
    order = v[BASE].sort_values(ascending=False).index
    top, rest = list(order[:keep]), list(order[keep:])
    vo, qo = v.loc[rest, YEARS].sum(), q.loc[rest, YEARS].sum()
    po = (vo * _scale()) / qo.where(qo != 0)
    p2 = pd.concat([p.loc[top], pd.DataFrame([po], index=["Others"])])
    v2 = pd.concat([v.loc[top], pd.DataFrame([vo], index=["Others"])])
    q2 = pd.concat([q.loc[top], pd.DataFrame([qo], index=["Others"])])
    return p2, v2, q2


def _cagr(a, b, n):
    try:
        return (b / a) ** (1.0 / n) - 1.0
    except Exception:
        return float("nan")


def _is_plural(name) -> bool:
    return str(name).strip().lower().endswith("s")


def _fmt(x, dp=0) -> str:
    return f"{x:,.{dp}f}"


def _pct(x, dp=1, sign=False) -> str:
    return f"{x * 100:+.{dp}f}%" if sign else f"{x * 100:.{dp}f}%"


def _band(idx):
    d = idx - 100.0
    for lim, col in _BANDS:
        if d >= lim:
            return _c(col)
    return _c(_LOW)


# ═══════════════════════ drawing helpers (px) ════════════════════════
def _style_off(shape):
    st = shape._element.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}style")
    if st is not None:
        shape._element.remove(st)


def _rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, _emu(x), _emu(y), _emu(w), _emu(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(0.75)
    s.shadow.inherit = False
    _style_off(s)
    return s


def _text(slide, x, y, w, h, text, size=11, bold=False, color=INK,
          italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    s = slide.shapes.add_textbox(_emu(x), _emu(y), _emu(w), _emu(h))
    tf = s.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(27432)
    tf.margin_top = tf.margin_bottom = Emu(27432)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name, r.font.size = FONT, Pt(size)
    r.font.bold, r.font.italic = bold, italic
    r.font.color.rgb = color
    return s


def _headline(slide, text):
    box = _rect(slide, X(0.3), Y(1.05), S(12.73), S(0.5), NAVY)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(137160)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    r.font.name, r.font.size, r.font.bold = FONT, Pt(13), True
    r.font.color.rgb = WHITE
    return box


def _panel_title(slide, x, y, w, text):
    _text(slide, x, y, w, 27, text, size=11.5, bold=True, color=NAVY)


def _cards(slide, y, h, items):
    n = len(items)
    gap = 14.4
    w = (S(12.73) - gap * (n - 1)) / n
    for i, (kicker, sentence) in enumerate(items):
        x = X(0.3) + i * (w + gap)
        _rect(slide, x, y, w, h, PANEL)
        c = _rect(slide, x + 12, y + 13, 33, 33, NAVY, shape=MSO_SHAPE.OVAL)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(i + 1)
        r.font.size, r.font.bold, r.font.name = Pt(12), True, FONT
        r.font.color.rgb = WHITE
        _text(slide, x + 53, y + 9, w - 62, 28, kicker, size=11.5, bold=True,
              color=NAVY)
        _text(slide, x + 53, y + 38, w - 62, h - 42, sentence, size=10.5)


def _table(slide, x, y, w, h, header, rows, col_w_in, fills=None, font=8.5,
           head_font=8.5):
    nr, nc = len(rows) + 1, len(header)
    gt = slide.shapes.add_table(nr, nc, _emu(x), _emu(y), _emu(w), _emu(h))
    t = gt.table
    t.first_row = True
    for j, cw in enumerate(col_w_in):
        t.columns[j].width = _emu(S(cw))
    for i in range(nr):
        t.rows[i].height = _emu(h / nr)
    for i in range(nr):
        for j in range(nc):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Emu(36576)
            c.margin_top = c.margin_bottom = Emu(9144)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            txt = header[j] if i == 0 else rows[i - 1][j]
            tf = c.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = txt
            r.font.name = FONT
            if i == 0:
                r.font.size, r.font.bold = Pt(head_font), True
                r.font.color.rgb = WHITE
                c.fill.solid()
                c.fill.fore_color.rgb = NAVY
            else:
                r.font.size = Pt(font)
                r.font.bold = (j == 0)
                r.font.color.rgb = INK
                c.fill.solid()
                f = fills[i - 1][j] if fills else None
                c.fill.fore_color.rgb = (f if f is not None
                                         else (WHITE if i % 2 else PANEL))
    return t


def _axis_text(axis, size=9, color=MUTED):
    axis.tick_labels.font.size = Pt(size)
    axis.tick_labels.font.name = FONT
    axis.tick_labels.font.color.rgb = color


def _skip(chart):
    """Every category label prints: no auto-skipping."""
    cax = chart._chartSpace.find(".//" + qn("c:catAx"))
    for tag in ("c:tickLblSkip", "c:noMultiLvlLbl"):
        old = cax.find(qn(tag))
        if old is not None:
            cax.remove(old)
    a = cax.makeelement(qn("c:tickLblSkip"), {"val": "1"})
    b = cax.makeelement(qn("c:noMultiLvlLbl"), {"val": "0"})
    cax.append(a)
    cax.append(b)


def _chart(slide, kind, x, y, w, h, cd, name):
    gf = slide.shapes.add_chart(kind, _emu(x), _emu(y), _emu(w), _emu(h), cd)
    gf.name = name
    return gf.chart


def _shape_series(key, vals, k=2.6):
    """The yearly path of a price line, bent along one of the client's growth
    PATTERNS (trend_patterns) instead of a ruler-straight rise.

    The anchor years are untouched - the first year, the base year and the
    last year keep their real prices, so the table beside the chart still
    reads off the line - and between them each year's growth follows the
    pattern's shape (more growth where the curve is high), re-scaled so the
    interval compounds to exactly its real change.  ``key`` seeds a stable
    choice of pattern per series."""
    ys = [y for y in YEARS if not pd.isna(vals[y])]
    if len(ys) < 4:
        return [None if pd.isna(vals[y]) else float(vals[y]) for y in YEARS]
    seed = int(hashlib.md5(str(key).encode("utf-8")).hexdigest()[:8], 16)
    pat = _tp.PATTERNS[random.Random(seed).choice(sorted(_tp.PATTERNS))]

    def pv(y):
        return pat[min(max(y - _tp.YEARS[0], 0), len(pat) - 1)]

    anchors = sorted({ys[0], min(max(BASE, ys[0]), ys[-1]), ys[-1]})
    out = {anchors[0]: float(vals[anchors[0]])}
    for a, b in zip(anchors, anchors[1:]):
        pa, pb = float(vals[a]), float(vals[b])
        n = b - a
        if n < 2 or pa <= 0 or pb <= 0:
            for y in range(a + 1, b + 1):
                out[y] = float(vals[y])
            continue
        R = pb / pa
        g = R ** (1.0 / n) - 1.0
        pts = [pv(y) for y in range(a + 1, b + 1)]
        m = sum(pts) / len(pts) or 1.0
        w = [min(3.5, max(-0.5, 1.0 + k * (p_ / m - 1.0))) for p_ in pts]
        base = [g * x for x in w]

        def grown(shift):
            t = 1.0
            for r in base:
                t *= max(1e-6, 1.0 + r + shift)
            return t
        lo, hi = -0.5, 0.5
        for _ in range(60):
            mid = (lo + hi) / 2.0
            if grown(mid) < R:
                lo = mid
            else:
                hi = mid
        shift = (lo + hi) / 2.0
        lvl = pa
        for y, r in zip(range(a + 1, b + 1), base):
            lvl *= 1.0 + r + shift
            out[y] = lvl
        out[b] = pb
    return [None if y not in out else round(out[y], 0) for y in YEARS]


# ═════════════════════ 1. single-segment slide ═══════════════════════
def _single(new_slide, finish, region, segment):
    p, v, q = _prices(region, segment)
    if not _chartable(p):
        return False
    p, v, q = _top_with_others(p, v, q)
    avg = _avg(v, q)
    a26, a36 = float(avg[BASE]), float(avg[END])
    p26 = p[BASE].dropna().sort_values(ascending=False)
    if len(p26) < 2:
        return False
    idx = p[BASE] / a26 * 100.0
    g = pd.Series({n: _cagr(p.loc[n, BASE], p.loc[n, END], END - BASE)
                   for n in p.index})
    hi, lo = p26.index[0], p26.index[-1]
    spread = p26.iloc[0] / p26.iloc[-1] - 1.0
    fast, slow = g.idxmax(), g.idxmin()
    disp0 = p[START].max() / p[START].min() - 1.0
    disp1 = p[END].max() / p[END].min() - 1.0
    label = segment
    pl = _is_plural(hi)
    v1, v2, v3 = (("command", "sit", "post") if pl
                  else ("commands", "sits", "posts"))
    pt_lo = _is_plural(lo)
    headline = (f"{hi} {v1} the highest {BASE} price point at US$ "
                f"{_fmt(p26.iloc[0])}/{_vu()}, a {_pct(spread)} premium over "
                f"{lo}.")
    s = new_slide(
        f"{region} {MKT} Pricing Assessment: By {label}",
        f"Average price ({_config.price_unit()}) by {label.lower()}, {START} to {END}, with "
        f"{BASE} price index and growth outlook")
    _headline(s, headline)

    # left: price trend
    _rect(s, X(0.3), Y(1.65), S(7.05), S(3.85), PANEL)
    _panel_title(s, X(0.42), Y(1.7), S(6.8),
                 f"Price Trend by {label}, {START} to {END} ({_config.price_unit()})")
    cd = CategoryChartData()
    cd.categories = [str(y) for y in YEARS]
    order = list(p26.index)
    for n in order:
        cd.add_series(n, _shape_series(f"{MKT}|{region}|{segment}|{n}",
                                       p.loc[n]))
    ch = _chart(s, XL_CHART_TYPE.LINE, X(0.38), Y(1.98), S(6.9), S(3.48), cd,
                "FMI Price Chart")
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size, ch.legend.font.name = Pt(9), FONT
    va = ch.value_axis
    lo_v, hi_v = float(p[YEARS].min().min()), float(p[YEARS].max().max())
    step = 500.0 if (hi_v - lo_v) < 8000 else 2000.0
    va.minimum_scale = math.floor(lo_v / step) * step
    va.maximum_scale = (math.floor(hi_v / step) + 1) * step
    va.major_unit = step
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GRID
    va.tick_labels.number_format = "#,##0"
    va.tick_labels.number_format_is_linked = False
    va.format.line.fill.background()
    _axis_text(va)
    ca = ch.category_axis
    _axis_text(ca)
    ca.format.line.color.rgb = GRID
    ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    for i, ser in enumerate(ch.plots[0].series):
        ser.smooth = False
        ser.format.line.width = Pt(2)
        ser.format.line.color.rgb = _c(SERIES[i % len(SERIES)])
        ser.marker.style = XL_MARKER_STYLE.NONE
    _text(s, X(5.0), Y(1.72), S(2.3), 24, f"Forecast period: {BASE} to {END}",
          size=8.5, italic=True, color=MUTED, align=PP_ALIGN.RIGHT)

    # right top: 2026 index
    _rect(s, X(7.5), Y(1.65), S(5.53), S(1.75), PANEL)
    _panel_title(s, X(7.62), Y(1.7), S(5.3),
                 f"{BASE} Price Index vs {region} Average (= 100)")
    isort = idx.dropna().sort_values()
    cd2 = CategoryChartData()
    cd2.categories = list(isort.index)
    cd2.add_series("Index", [round(float(x), 1) for x in isort.values])
    c2 = _chart(s, XL_CHART_TYPE.BAR_CLUSTERED, X(7.55), Y(1.93), S(5.43),
                S(1.45), cd2, "FMI Open Chart")
    c2.has_title = False
    c2.has_legend = False
    _skip(c2)
    pl2 = c2.plots[0]
    pl2.gap_width = 45
    pl2.has_data_labels = True
    pl2.data_labels.number_format = "0.0"
    pl2.data_labels.number_format_is_linked = False
    pl2.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    pl2.data_labels.font.size = Pt(9)
    pl2.data_labels.font.bold = True
    pl2.data_labels.font.color.rgb = INK
    for i, val in enumerate(isort.values):
        pt = pl2.series[0].points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = _c(ORANGE if val >= 100 else BLUE)
    v2a = c2.value_axis
    v2a.minimum_scale = round(min(isort.min(), 100) - 3)
    v2a.maximum_scale = round(max(isort.max(), 100) + 2)
    v2a.visible = False
    v2a.has_major_gridlines = False
    _axis_text(c2.category_axis, size=8.5, color=INK)
    c2.category_axis.format.line.color.rgb = GRID

    # right bottom: KPI table
    header = [label, f"{START}", f"{BASE}", f"{END}", f"CAGR {BASE} to {END}",
              "Index"]
    rows, fills = [], []
    for n in order:
        rows.append([n, _fmt(p.loc[n, START]), _fmt(p.loc[n, BASE]),
                     _fmt(p.loc[n, END]), _pct(g[n], 2),
                     f"{idx[n]:.1f}"])
        fills.append([None] * 5 + [_band(idx[n])])
    rows.append([f"{region} Average", _fmt(avg[START]), _fmt(a26),
                 _fmt(a36), _pct(_cagr(a26, a36, END - BASE), 2), "100.0"])
    fills.append([PANEL] * 6)
    _table(s, X(7.5), Y(3.48), S(5.53), S(2.02), header, rows,
           [2.33, 0.6, 0.6, 0.6, 0.85, 0.55], fills=fills)

    _cards(s, Y(5.6), S(1.35), [
        ("Price Hierarchy",
         f"{hi} {v2} at the top of the {BASE} price ladder at an index of "
         f"{idx[hi]:.1f}, while {lo} {'trade' if pt_lo else 'trades'} at "
         f"{idx[lo]:.1f}, which keeps the full {label.lower()} spread within "
         f"{_pct(spread)}."),
        ("Price Momentum",
         f"{fast} {v3} the fastest price growth at a {_pct(g[fast], 2)} CAGR "
         f"between {BASE} and {END}, compared with {_pct(g[slow], 2)} for "
         f"{slow}."),
        ("Outlook",
         f"The {region} average price rises from US$ {_fmt(a26)}/{_vu()} in "
         f"{BASE} to US$ {_fmt(a36)}/{_vu()} in {END}, and the gap between the "
         f"highest and lowest priced {label.lower()} moves from "
         f"{_pct(disp0)} in {START} to {_pct(disp1)} in {END}.")])
    _text(s, X(0.3), Y(7.02) - 8, S(12.7), 22,
          "Source: FMI analysis. Average price = market value (US$) / market "
          f"volume ({_vu()}). {START} to {BASE - 1} historical, {BASE} to {END} "
          "forecast.", size=8.5, color=MUTED)
    finish(s)
    return True


# ═══════════════════ 2. cross-sectional matrix slide ═════════════════
def _cross(new_slide, finish, region, cross_seg):
    a_name, b_name = data.cross_axes(cross_seg)
    mv = data.cross_matrix_main(region, "Value", cross_seg, BASE)
    mq = data.cross_matrix_main(region, "Volume", cross_seg, BASE)
    if mv.empty or mq.empty:
        return False
    cols = [c for c in mv.columns if c in mq.columns]
    rows_ = [r for r in mv.index if r in mq.index]
    mv, mq = mv.loc[rows_, cols], mq.loc[rows_, cols]
    if len(cols) > 9:                      # top 8 columns by value + Others
        order = list(mv.sum().sort_values(ascending=False).index)
        keep, rest = order[:8], order[8:]
        mv = pd.concat([mv[keep], mv[rest].sum(axis=1).rename("Others")],
                       axis=1)
        mq = pd.concat([mq[keep], mq[rest].sum(axis=1).rename("Others")],
                       axis=1)
        cols = list(mv.columns)
    sc = _scale()
    M = (mv * sc) / mq.where(mq != 0)
    rowavg = (mv.sum(axis=1) * sc) / mq.sum(axis=1).where(mq.sum(axis=1) != 0)
    colavg = (mv.sum(axis=0) * sc) / mq.sum(axis=0).where(mq.sum(axis=0) != 0)
    tot = float(mv.values.sum() * sc / mq.values.sum())
    row_order = list(rowavg.sort_values(ascending=False).index)
    M = M.loc[row_order]
    I = M / tot * 100.0
    stack = I.stack()
    if len(stack) < 4:
        return False
    best, worst = stack.idxmax(), stack.idxmin()

    # finest level the engine keeps: the row parent x the column's sub-types
    fv = data.cross_matrix(region, "Value", cross_seg, BASE)
    fq = data.cross_matrix(region, "Volume", cross_seg, BASE)
    fp = (fv * sc) / fq.where(fq != 0)
    rng = pd.DataFrame({"min": fp.min(axis=1), "max": fp.max(axis=1)})
    rng = rng.loc[[r for r in row_order if r in rng.index]].dropna()
    if rng.empty:
        return False
    widest = (rng["max"] / rng["min"] - 1.0).idxmax()
    widest_pct = float(rng.loc[widest, "max"] / rng.loc[widest, "min"] - 1.0)

    headline = (f"{best[0]} for {best[1]} is the highest priced combination "
                f"in {BASE} at US$ {_fmt(M.loc[best])}/{_vu()}, "
                f"{_pct(I.loc[best] / 100 - 1, sign=True)} above the "
                f"{region} average.")
    s = new_slide(
        f"{region} {MKT} Price Matrix: {a_name} × {b_name}",
        f"Average price ({_config.price_unit()}) for every {a_name.lower()} and "
        f"{b_name.lower()} combination, {BASE}")
    _headline(s, headline)

    _rect(s, X(0.3), Y(1.65), S(8.25), S(3.85), PANEL)
    _panel_title(s, X(0.42), Y(1.7), S(8.0),
                 f"{BASE} Price Matrix ({_config.price_unit()}), Shaded by Index vs "
                 f"{region} Average")
    header = [a_name] + cols + ["Row Avg."]
    body, fills = [], []
    for r in row_order:
        body.append([r] + [_fmt(M.loc[r, c]) for c in cols]
                    + [_fmt(rowavg[r])])
        fills.append([WHITE] + [_band(I.loc[r, c]) for c in cols] + [PANEL])
    body.append(["Column Avg."] + [_fmt(colavg[c]) for c in cols]
                + [_fmt(tot)])
    fills.append([PANEL] * (len(cols) + 2))
    first_w, last_w = 1.7, 0.66
    cw = (8.05 - first_w - last_w) / len(cols)
    _table(s, X(0.4), Y(2.02), S(8.05), S(3.1), header, body,
           [first_w] + [cw] * len(cols) + [last_w], fills=fills, font=9,
           head_font=8)
    lx = X(0.42)
    for col, txt in [(_LOW, "Index below 97.5"), ("D9E6F2", "97.5 to 99"),
                     ("F7F7F7", "Parity: 99 to 101"),
                     ("FBE3D3", "101 to 102.5"),
                     ("F4B183", "Index above 102.5")]:
        _rect(s, lx, Y(5.22), 19, 15, _c(col), line=GRID)
        lab = _text(s, lx + 22, Y(5.16), 134, 25, txt, size=8.5, color=MUTED)
        lab.name = "FMI Scale legend"          # a key, not a figure
        lx += 154

    _rect(s, X(8.7), Y(1.65), S(4.33), S(3.85), PANEL)
    ptitle = f"Price Range Across All {b_name} Sub-Types, {BASE}"
    _panel_title(s, X(8.82), Y(1.7), S(4.1), ptitle)
    two_line = len(ptitle) > 42
    _text(s, X(8.82), Y(1.96) + (16 if two_line else 0), S(4.1), 23,
          f"Bar spans lowest to highest price ({_config.price_unit()}) at the finest data "
          "level.", size=8.5, italic=True, color=MUTED)
    rr = rng.iloc[::-1]
    cd = CategoryChartData()
    cd.categories = list(rr.index)
    floor = int(rng["min"].min() // 2000 * 2000)
    cd.add_series("Base", [round(float(x)) for x in rr["min"]])
    cd.add_series("Range", [round(float(b - a))
                            for a, b in zip(rr["min"], rr["max"])])
    c = _chart(s, XL_CHART_TYPE.BAR_STACKED, X(8.75), Y(2.2), S(4.23),
               S(3.25), cd, "FMI Price Chart")
    c.has_legend = False
    _skip(c)
    pl = c.plots[0]
    pl.gap_width, pl.overlap = 55, 100
    base_s, rng_s = pl.series[0], pl.series[1]
    base_s.format.fill.background()
    rng_s.format.fill.solid()
    rng_s.format.fill.fore_color.rgb = TEAL
    for i, (a, b) in enumerate(zip(rr["min"], rr["max"])):
        dl = base_s.points[i].data_label
        dl.has_text_frame = True
        dl.text_frame.text = f"{_fmt(a)} to {_fmt(b)}"
        dl.position = XL_LABEL_POSITION.INSIDE_END
        for para in dl.text_frame.paragraphs:
            for r in para.runs:
                r.font.size, r.font.bold = Pt(8.5), True
                r.font.color.rgb, r.font.name = INK, FONT
    va = c.value_axis
    va.minimum_scale = floor
    va.maximum_scale = int((rng["max"].max() // 2000 + 1) * 2000)
    va.major_unit = 2000
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GRID
    va.format.line.fill.background()
    va.tick_labels.number_format = "#,##0"
    va.tick_labels.number_format_is_linked = False
    _axis_text(va, size=8)
    _axis_text(c.category_axis, size=8.5, color=INK)
    c.category_axis.format.line.color.rgb = GRID

    gap_lo = 1 - I.loc[worst] / 100
    _cards(s, Y(5.6), S(1.35), [
        ("Premium Pockets",
         f"{best[0]} positioned on {best[1]} carry the strongest premium at "
         f"an index of {I.loc[best]:.1f}, which signals room for value-added "
         f"claims in this combination."),
        ("Value Pockets",
         f"{worst[0]} for {worst[1]} is the lowest priced combination at an "
         f"index of {I.loc[worst]:.1f}, sitting {_pct(gap_lo)} below the "
         f"{region} average of US$ {_fmt(tot)}/{_vu()}."),
        ("Within-Type Spread",
         f"{widest} show the widest internal price spread at "
         f"{_pct(widest_pct)}, ranging from US$ "
         f"{_fmt(rng.loc[widest, 'min'])}/{_vu()} to US$ "
         f"{_fmt(rng.loc[widest, 'max'])}/{_vu()} across sub-types.")])
    _text(s, X(0.3), Y(7.02) - 8, S(12.7), 22,
          f"Source: FMI analysis. Cell price = combined value (US$) / "
          f"combined volume ({_vu()}); shading shows the index vs the {region} "
          f"average of US$ {_fmt(tot)}/{_vu()}.", size=8.5, color=MUTED)
    finish(s)
    return True


# ═══════════════════ 3. country benchmarking slide ═══════════════════
def _country(new_slide, finish, region, segment):
    kids = [k for k in data.children(region)
            if not str(k).lower().startswith(("other ", "rest of"))]
    recs = []
    for k in kids:
        try:
            p, v, q = _prices(k, segment)
        except Exception:
            continue
        if p.empty or p[BASE].dropna().empty:
            continue
        a = _avg(v, q)
        recs.append(dict(country=k, mn=float(p[BASE].min()),
                         mx=float(p[BASE].max()), avg=float(a[BASE]),
                         g=_cagr(a[BASE], a[END], END - BASE)))
    if len(recs) < 3:
        return False
    D = pd.DataFrame(recs).set_index("country").sort_values(
        "avg", ascending=False)
    _p, rv, rq = _prices(region, segment)
    rav = _avg(rv, rq)
    eu, eu_g = float(rav[BASE]), _cagr(rav[BASE], rav[END], END - BASE)
    D["idx"] = D["avg"] / eu * 100.0
    top, bot = D.index[0], D.index[-1]
    plural = f"{segment}s" if not _is_plural(segment) else segment
    child_word = "market" if region != "Global" else "region"
    headline = (f"{top} is the most premium {region} {child_word} at US$ "
                f"{_fmt(D.loc[top, 'avg'])}/{_vu()} in {BASE}, "
                f"{_pct(D.loc[top, 'idx'] / 100 - 1, sign=True)} above the "
                f"{'regional' if region != 'Global' else 'global'} average, "
                f"while {bot} is the lowest priced.")
    kind = "Country" if region != "Global" else "Region"
    s = new_slide(
        f"{region} {MKT} Pricing Assessment: {kind} Benchmarking",
        f"{kind} average price and price range across {plural.lower()}, "
        f"{BASE} ({_config.price_unit()})")
    _headline(s, headline)

    _rect(s, X(0.3), Y(1.65), S(8.25), S(3.85), PANEL)
    _panel_title(s, X(0.42), Y(1.7), S(8.0),
                 f"{BASE} Price Range Across {plural} by {kind} ({_config.price_unit()}), "
                 f"Ranked by Average")
    # The bars are drawn on the PRICE INDEX (region average = 100), so one
    # left axis lets every country be compared at a glance and no figure is
    # written on the columns (client, 2026-10-07).  An index is not a price:
    # the axis stays readable in a sample.
    imn, imx = D["mn"] / eu * 100.0, D["mx"] / eu * 100.0
    cd = CategoryChartData()
    cd.categories = list(D.index)
    cd.add_series("Base", [round(float(x), 1) for x in imn])
    cd.add_series("Range", [round(float(b - a), 1)
                            for a, b in zip(imn, imx)])
    c = _chart(s, XL_CHART_TYPE.COLUMN_STACKED, X(0.38), Y(2.18), S(8.1),
               S(3.3), cd, "FMI Open Chart")
    c.has_legend = False
    pl = c.plots[0]
    pl.gap_width, pl.overlap = 40, 100
    pl.series[0].format.fill.background()
    rs = pl.series[1]
    rs.format.fill.solid()
    rs.format.fill.fore_color.rgb = TEAL
    for i, cn in enumerate(D.index):
        if D.loc[cn, "idx"] < 100:
            pt = rs.points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = _c(BLUE)
    va = c.value_axis
    lo_i = math.floor(float(imn.min()) / 5.0) * 5.0
    hi_i = math.ceil(float(imx.max()) / 5.0) * 5.0
    va.minimum_scale, va.maximum_scale = lo_i, hi_i
    va.major_unit = 5.0 if (hi_i - lo_i) <= 50 else 10.0
    va.tick_labels.number_format = "0"
    va.tick_labels.number_format_is_linked = False
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GRID
    va.format.line.fill.background()
    _axis_text(va, size=8.5)
    ca = c.category_axis
    _axis_text(ca, size=8.5, color=INK)
    ca.format.line.color.rgb = GRID
    _skip(c)
    _text(s, X(0.42), Y(1.95), S(7.9), 23,
          f"Bar spans the lowest to highest {segment.lower()} price in each "
          f"{kind.lower()}, read on the price index axis ({region} average = "
          f"100). Teal = above {region} average, blue = below.", size=8.5,
          italic=True, color=MUTED)

    _rect(s, X(8.7), Y(1.65), S(4.33), S(3.85), PANEL)
    _panel_title(s, X(8.82), Y(1.7), S(4.1), "Premium and Discount Markets")
    n_side = min(5, len(D) // 2)
    sel = list(D.index[:n_side]) + list(D.index[-n_side:])
    body, fills = [], []
    for cn in sel:
        body.append([cn, _fmt(D.loc[cn, "avg"]), f"{D.loc[cn, 'idx']:.1f}",
                     _pct(D.loc[cn, "g"], 2)])
        fills.append([None, None, _band(D.loc[cn, "idx"]), None])
    _table(s, X(8.8), Y(2.02), S(4.13), S(3.38),
           [kind, f"Avg. {BASE}", "Index", f"CAGR {BASE} to {END}"], body,
           [1.55, 0.86, 0.72, 1.0], fills=fills, font=9, head_font=9)

    spread_c = D["avg"].max() / D["avg"].min() - 1.0
    fastest = D["g"].idxmax()
    within = D["mx"] / D["mn"] - 1.0
    wid = within.idxmax()
    _cards(s, Y(5.6), S(1.35), [
        (f"{kind} Spread",
         f"Average prices across the ranked "
         f"{'countries' if kind == 'Country' else 'regions'} differ by "
         f"{_pct(spread_c)} in "
         f"{BASE}, from US$ {_fmt(D['avg'].min())}/{_vu()} in {bot} to US$ "
         f"{_fmt(D['avg'].max())}/{_vu()} in {top}."),
        ("Growth Leader",
         f"{fastest} records the fastest average price growth at a "
         f"{_pct(D.loc[fastest, 'g'], 2)} CAGR between {BASE} and {END}, "
         f"against {_pct(eu_g, 2)} for {region} as a whole."),
        ("Mix Exposure",
         f"{wid} shows the widest gap between its cheapest and most "
         f"expensive {segment.lower()} at {_pct(within[wid])}, so product "
         f"mix shifts move its average price the most.")])
    _text(s, X(0.3), Y(7.02) - 8, S(12.7), 22,
          f"Source: FMI analysis. Range = lowest to highest {segment.lower()} "
          f"price in each {kind.lower()}; ranked by average price. Other "
          f"geographies grouped as 'Other' are excluded from the ranking.",
          size=8.5, color=MUTED)
    finish(s)
    return True


# ═══════════════════════════ entry point ═════════════════════════════
def single_segments(region):
    segs = [s for s in data.SEGMENTS
            if not data.is_cross(s) and s not in ("Region", "Country")]
    out = []
    for sg in segs:
        try:
            p, _v, _q = _prices(region, sg)
        except Exception:
            continue
        if _chartable(p):
            out.append(sg)
    return out


def build(new_slide, finish, region):
    """Build the pricing slides for ``region``; returns how many were drawn.
    ``new_slide(title, subtitle)`` hands back a slide inside the deck's own
    header, ``finish(slide)`` closes it (footer, notes)."""
    if not (data.has_unit("Value") and data.has_unit("Volume")):
        return 0
    singles = single_segments(region)
    if not singles:
        return 0
    primary = (PRICING_PRIMARY_SEGMENT if PRICING_PRIMARY_SEGMENT in singles
               else singles[0])
    made = 0
    for sg in (singles if PRICING_ALL_SEGMENTS else [primary]):
        made += bool(_single(new_slide, finish, region, sg))
    for cs in [c for c in data.SEGMENTS if data.is_cross(c)][:PRICING_MAX_CROSS]:
        made += bool(_cross(new_slide, finish, region, cs))
    made += bool(_country(new_slide, finish, region, primary))
    return made
