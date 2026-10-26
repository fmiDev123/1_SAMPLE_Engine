"""Reference build for the redesigned FMI pricing slides.

Builds three slide types from the Formatted Dataset 'Master Sheet':
  1. Single-segment pricing slide   (e.g. Product Type)
  2. Cross-sectional price matrix   (e.g. Product Type x Primary Function)
  3. Country price benchmarking     (country spread across one segment)

Average price (US$/MT) = Value (US$ Mn) * 1,000,000 / Volume (MT).
Parent prices are re-derived as sum(value) / sum(volume), never as an
average of child prices.
"""
import sys
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

DATA = sys.argv[1]
LOGO = sys.argv[2]
OUT = sys.argv[3]

MARKET = "Pet Cosmetic Products"
REGION = "Europe"
BASE, START, END = 2026, 2021, 2036
YEARS = list(range(START, END + 1))
MASK = False  # True for sample decks: market-level prices print as XX

# ---------- palette (taken from the existing FMI slide) ----------
NAVY = RGBColor(0x0B, 0x2E, 0x4A)
TEAL = RGBColor(0x00, 0x81, 0x89)
GOLD = RGBColor(0x9E, 0x7E, 0x0A)
PLUM = RGBColor(0x40, 0x31, 0x5E)
INK = RGBColor(0x1F, 0x1F, 0x1F)
MUTED = RGBColor(0x59, 0x63, 0x6E)
PANEL = RGBColor(0xF3, 0xF6, 0xF9)
GRID = RGBColor(0xD9, 0xDE, 0xE4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SERIES = ["C55A11", "0B2E4A", "7030A0", "008189", "9E7E0A", "5B7DB1", "A23B5C", "6E8B3D"]
ABOVE = [RGBColor(0xFB, 0xE3, 0xD3), RGBColor(0xF4, 0xB1, 0x83)]   # >+1%, >+2.5%
BELOW = [RGBColor(0xD9, 0xE6, 0xF2), RGBColor(0x9D, 0xC0, 0xE0)]   # <-1%, <-2.5%
PARITY = RGBColor(0xF7, 0xF7, 0xF7)
FONT = "Calibri"


def hexrgb(h):
    return RGBColor.from_string(h)


# =================================================================
# DATA LAYER
# =================================================================
def load(path):
    df = pd.read_excel(path, sheet_name="Master Sheet", header=5)
    keys = ["Region", "Segment", "Sub-segment", "Sub-segment 1", "Sub-segment 2"]
    v = df[df.Unit == "Value"].set_index(keys)[YEARS]
    q = df[df.Unit == "Volume"].set_index(keys)[YEARS]
    return v, q


def price(vsum, qsum):
    p = vsum * 1e6 / qsum.replace(0, pd.NA)
    return p.astype(float)


def segment_prices(v, q, region, segment, level):
    """level: 'Sub-segment' (parent) or 'Sub-segment 1' (child)."""
    vs = v.xs((region, segment), level=("Region", "Segment")).groupby(level=level, sort=False).sum()
    qs = q.xs((region, segment), level=("Region", "Segment")).groupby(level=level, sort=False).sum()
    return price(vs, qs)


def region_total(v, q, region, segment):
    vs = v.xs((region, segment), level=("Region", "Segment")).sum()
    qs = q.xs((region, segment), level=("Region", "Segment")).sum()
    return vs * 1e6 / qs


def cagr(a, b, n):
    return (b / a) ** (1 / n) - 1


def fmt(x, dp=0):
    if MASK:
        return "XX"
    return f"{x:,.{dp}f}"


def pct(x, dp=1, sign=True):
    s = f"{x*100:+.{dp}f}%" if sign else f"{x*100:.{dp}f}%"
    return s


# =================================================================
# DRAWING HELPERS
# =================================================================
def tb(slide, x, y, w, h, text, size=11, bold=False, color=INK, italic=False,
       align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, wrap=True):
    s = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = s.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, Inches(0.03))
    runs = text if isinstance(text, list) else [[(text, {})]]
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for t, o in para:
            r = p.add_run()
            r.text = t
            f = r.font
            f.name = FONT
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", italic)
            f.color.rgb = o.get("color", color)
    return s


def rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(0.75)
    s.shadow.inherit = False
    st = s._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}style')
    if st is not None:
        s._element.remove(st)
        s.text_frame.paragraphs[0].font.color.rgb = WHITE
    return s


def hline(slide, x, y, w, color, weight=2.25):
    c = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x + w), Inches(y))
    c.line.color.rgb = color
    c.line.width = Pt(weight)


def frame(prs, title, subtitle, headline, source):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(LOGO, Inches(0.3), Inches(0.12), height=Inches(0.72))
    tb(s, 1.75, 0.14, 11.3, 0.42, title, size=22, bold=True, color=INK)
    tb(s, 1.75, 0.53, 11.3, 0.32, subtitle, size=12.5, italic=True, color=INK)
    hline(s, 0.3, 0.95, 4.2, TEAL)
    hline(s, 4.6, 0.95, 4.2, GOLD)
    hline(s, 8.9, 0.95, 4.13, PLUM)
    box = rect(s, 0.3, 1.05, 12.73, 0.5, NAVY)
    tf = box.text_frame
    tf.margin_left = Inches(0.15)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = headline
    r.font.name, r.font.size, r.font.bold, r.font.color.rgb = FONT, Pt(13), True, WHITE
    tb(s, 0.3, 7.02, 9.5, 0.22, source, size=8.5, color=MUTED)
    tb(s, 9.8, 7.02, 3.23, 0.22, "©2026 Future Market Insights, All Rights Reserved",
       size=8.5, color=MUTED, align=PP_ALIGN.RIGHT)
    return s


def panel_title(slide, x, y, w, text):
    tb(slide, x, y, w, 0.28, text, size=11.5, bold=True, color=NAVY)


def style_axis_text(axis, size=9, color=MUTED):
    axis.tick_labels.font.size = Pt(size)
    axis.tick_labels.font.name = FONT
    axis.tick_labels.font.color.rgb = color


def insight_cards(slide, y, h, items):
    """items: list of (kicker, sentence). Three cards across the slide."""
    n = len(items)
    gap = 0.15
    w = (12.73 - gap * (n - 1)) / n
    for i, (kicker, sentence) in enumerate(items):
        x = 0.3 + i * (w + gap)
        rect(slide, x, y, w, h, PANEL)
        circ = rect(slide, x + 0.12, y + 0.14, 0.34, 0.34, NAVY, shape=MSO_SHAPE.OVAL)
        ctf = circ.text_frame
        ctf.margin_left = ctf.margin_right = 0
        cp = ctf.paragraphs[0]
        cp.alignment = PP_ALIGN.CENTER
        cr = cp.add_run()
        cr.text = str(i + 1)
        cr.font.size, cr.font.bold, cr.font.color.rgb, cr.font.name = Pt(12), True, WHITE, FONT
        tb(slide, x + 0.55, y + 0.1, w - 0.65, 0.3, kicker, size=11.5, bold=True, color=NAVY)
        tb(slide, x + 0.55, y + 0.4, w - 0.65, h - 0.45, sentence, size=10.5, color=INK)


def table(slide, x, y, w, h, header, rows, col_w, fills=None, font=9.5, head_font=9.5,
          first_bold=True, align_first=PP_ALIGN.LEFT):
    nr, nc = len(rows) + 1, len(header)
    gt = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(h))
    t = gt.table
    t.first_row = True
    for j, cw in enumerate(col_w):
        t.columns[j].width = Inches(cw)
    rh = Inches(h / nr)
    for i in range(nr):
        t.rows[i].height = rh
    for i in range(nr):
        for j in range(nc):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.04)
            c.margin_top = c.margin_bottom = Inches(0.01)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            txt = header[j] if i == 0 else rows[i - 1][j]
            tf = c.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = align_first if j == 0 else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = txt
            r.font.name = FONT
            if i == 0:
                r.font.size, r.font.bold, r.font.color.rgb = Pt(head_font), True, WHITE
                c.fill.solid()
                c.fill.fore_color.rgb = NAVY
            else:
                r.font.size = Pt(font)
                r.font.bold = first_bold and j == 0
                r.font.color.rgb = INK
                c.fill.solid()
                if fills and fills[i - 1][j] is not None:
                    c.fill.fore_color.rgb = fills[i - 1][j]
                else:
                    c.fill.fore_color.rgb = WHITE if i % 2 else PANEL
    return t


def skip(chart):
    """Force every category label to print (no auto-skipping)."""
    from lxml import etree
    ns = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
    cax = chart._chartSpace.find('.//' + ns + 'catAx')
    for tag in ('tickLblSkip', 'noMultiLvlLbl'):
        old = cax.find(ns + tag)
        if old is not None:
            cax.remove(old)
    el = etree.SubElement(cax, ns + 'tickLblSkip')
    el.set('val', '1')
    nm = etree.SubElement(cax, ns + 'noMultiLvlLbl')
    nm.set('val', '0')


def index_fill(idx):
    d = idx - 100
    if d >= 2.5:
        return ABOVE[1]
    if d >= 1.0:
        return ABOVE[0]
    if d <= -2.5:
        return BELOW[1]
    if d <= -1.0:
        return BELOW[0]
    return PARITY


# =================================================================
# SLIDE 1: SINGLE SEGMENT
# =================================================================
def slide_single(prs, v, q, segment, label):
    P = segment_prices(v, q, REGION, segment, "Sub-segment")
    avg = region_total(v, q, REGION, segment)
    a26, a36 = avg[BASE], avg[END]
    p26 = P[BASE].sort_values(ascending=False)
    idx = (P[BASE] / a26 * 100)
    g = cagr(P[BASE], P[END], END - BASE)
    hi, lo = p26.index[0], p26.index[-1]
    spread = p26.iloc[0] / p26.iloc[-1] - 1
    fast, slow = g.idxmax(), g.idxmin()
    disp21 = P[START].max() / P[START].min() - 1
    disp36 = P[END].max() / P[END].min() - 1

    headline = (f"{hi} command the highest {BASE} price point at US$ {fmt(p26.iloc[0])}/MT, "
                f"a {pct(spread, sign=False)} premium over {lo}.")
    s = frame(prs,
              f"{MARKET} Market {REGION} Pricing Assessment: By {label}",
              f"Average price (US$/MT) by {label.lower()}, {START} to {END}, with {BASE} price index and growth outlook",
              headline,
              f"Source: FMI analysis. Average price = market value (US$) / market volume (MT). "
              f"{START} to {BASE - 1} historical, {BASE} to {END} forecast.")

    # ---- left: trend line chart
    rect(s, 0.3, 1.65, 7.05, 3.85, PANEL)
    panel_title(s, 0.42, 1.7, 6.8, f"Price Trend by {label}, {START} to {END} (US$/MT)")
    cd = CategoryChartData()
    cd.categories = [str(y) for y in YEARS]
    order = list(p26.index)
    for name in order:
        cd.add_series(name, [round(x, 0) for x in P.loc[name, YEARS]])
    gf = s.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(0.38), Inches(1.98), Inches(6.9), Inches(3.48), cd)
    ch = gf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(9)
    ch.legend.font.name = FONT
    va = ch.value_axis
    lo_v = P[YEARS].min().min()
    hi_v = P[YEARS].max().max()
    va.minimum_scale = int(lo_v // 500 * 500)
    va.maximum_scale = int((hi_v // 500 + 1) * 500)
    va.major_unit = 500
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GRID
    va.tick_labels.number_format = '#,##0' if not MASK else '"XX"'
    va.tick_labels.number_format_is_linked = False
    va.format.line.fill.background()
    style_axis_text(va)
    ca = ch.category_axis
    style_axis_text(ca)
    ca.format.line.color.rgb = GRID
    ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    for i, ser in enumerate(ch.plots[0].series):
        ser.smooth = False
        ser.format.line.width = Pt(2)
        ser.format.line.color.rgb = hexrgb(SERIES[i % len(SERIES)])
        ser.marker.style = None
        from pptx.enum.chart import XL_MARKER_STYLE
        ser.marker.style = XL_MARKER_STYLE.NONE
    # forecast marker label
    tb(s, 5.0, 1.72, 2.3, 0.25, f"Forecast period: {BASE} to {END}", size=8.5, italic=True,
       color=MUTED, align=PP_ALIGN.RIGHT)

    # ---- right top: price index bar chart
    rect(s, 7.5, 1.65, 5.53, 1.75, PANEL)
    panel_title(s, 7.62, 1.7, 5.3, f"{BASE} Price Index vs {REGION} Average (= 100)")
    idx_sorted = idx.sort_values()
    cd2 = CategoryChartData()
    cd2.categories = list(idx_sorted.index)
    cd2.add_series("Index", [round(x, 1) for x in idx_sorted.values])
    gf2 = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(7.55), Inches(1.93),
                             Inches(5.43), Inches(1.45), cd2)
    c2 = gf2.chart
    c2.has_title = False
    c2.has_legend = False
    skip(c2)
    pl = c2.plots[0]
    pl.gap_width = 45
    pl.has_data_labels = True
    pl.data_labels.number_format = '0.0'
    pl.data_labels.number_format_is_linked = False
    pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    pl.data_labels.font.size = Pt(9)
    pl.data_labels.font.bold = True
    pl.data_labels.font.color.rgb = INK
    for i, val in enumerate(idx_sorted.values):
        pt = pl.series[0].points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = hexrgb("C55A11") if val >= 100 else hexrgb("5B7DB1")
    v2 = c2.value_axis
    v2.minimum_scale = round(min(idx_sorted.min(), 100) - 3)
    v2.maximum_scale = round(max(idx_sorted.max(), 100) + 2)
    v2.visible = False
    v2.has_major_gridlines = False
    style_axis_text(c2.category_axis, size=8.5, color=INK)
    c2.category_axis.format.line.color.rgb = GRID

    # ---- right bottom: KPI table
    header = [label, f"{START}", f"{BASE}", f"{END}", f"CAGR {BASE} to {END}", "Index"]
    rows, fills = [], []
    for name in order:
        rows.append([name, fmt(P.loc[name, START]), fmt(P.loc[name, BASE]), fmt(P.loc[name, END]),
                     pct(g[name], 2, sign=False), f"{idx[name]:.1f}"])
        fills.append([None] * 5 + [index_fill(idx[name])])
    rows.append([f"{REGION} Average", fmt(avg[START]), fmt(a26), fmt(a36),
                 pct(cagr(a26, a36, END - BASE), 2, sign=False), "100.0"])
    fills.append([PANEL] * 6)
    table(s, 7.5, 3.48, 5.53, 2.02, header, rows, [2.33, 0.6, 0.6, 0.6, 0.85, 0.55],
          fills=fills, font=8.5, head_font=8.5)

    # ---- insights
    insight_cards(s, 5.6, 1.35, [
        ("Price Hierarchy",
         f"{hi} sit at the top of the {BASE} price ladder at an index of {idx[hi]:.1f}, while {lo} "
         f"trade at {idx[lo]:.1f}, which keeps the full {label.lower()} spread within {pct(spread, sign=False)}."),
        ("Price Momentum",
         f"{fast} post the fastest price growth at a {pct(g[fast], 2, sign=False)} CAGR between {BASE} and {END}, "
         f"compared with {pct(g[slow], 2, sign=False)} for {slow}."),
        ("Outlook",
         f"The {REGION} average price rises from US$ {fmt(a26)}/MT in {BASE} to US$ {fmt(a36)}/MT in {END}, "
         f"and the gap between the highest and lowest priced {label.lower()} moves from "
         f"{pct(disp21, sign=False)} in {START} to {pct(disp36, sign=False)} in {END}."),
    ])
    return s


# =================================================================
# SLIDE 2: CROSS-SECTIONAL MATRIX
# =================================================================
def slide_cross(prs, v, q, cross_seg, row_label, col_seg, col_label):
    # map column child -> column parent from the single segment table
    cmap = (v.xs((REGION, col_seg), level=("Region", "Segment"))
             .reset_index()[["Sub-segment", "Sub-segment 1"]]
             .drop_duplicates().set_index("Sub-segment 1")["Sub-segment"])
    col_order = list(dict.fromkeys(cmap.values))

    vx = v.xs((REGION, cross_seg), level=("Region", "Segment"))[BASE].reset_index()
    qx = q.xs((REGION, cross_seg), level=("Region", "Segment"))[BASE].reset_index()
    vx["col"] = vx["Sub-segment 2"].map(cmap)
    qx["col"] = qx["Sub-segment 2"].map(cmap)
    mv = vx.groupby(["Sub-segment", "col"])[BASE].sum().unstack()
    mq = qx.groupby(["Sub-segment", "col"])[BASE].sum().unstack()
    M = (mv * 1e6 / mq)[col_order]
    rowavg = mv.sum(axis=1) * 1e6 / mq.sum(axis=1)
    tot = mv.values.sum() * 1e6 / mq.values.sum()
    row_order = list(rowavg.sort_values(ascending=False).index)
    M = M.loc[row_order]
    I = M / tot * 100

    # finest-grain range per row (child x child)
    fine = vx.assign(p=vx[BASE] * 1e6 / qx[BASE])
    rng = fine.groupby("Sub-segment")["p"].agg(["min", "max"]).loc[row_order]
    fine_hi = fine.loc[fine.p.idxmax()]
    fine_lo = fine.loc[fine.p.idxmin()]

    best = I.stack().idxmax()
    worst = I.stack().idxmin()
    colavg = mv.sum(axis=0) * 1e6 / mq.sum(axis=0)
    widest = (rng["max"] / rng["min"] - 1).idxmax()
    widest_pct = (rng.loc[widest, "max"] / rng.loc[widest, "min"] - 1)

    headline = (f"{best[0]} for {best[1]} is the highest priced combination in {BASE} at "
                f"US$ {fmt(M.loc[best])}/MT, {pct(I.loc[best]/100 - 1)} above the {REGION} average.")
    s = frame(prs,
              f"{MARKET} Market {REGION} Price Matrix: {row_label} × {col_label}",
              f"Average price (US$/MT) for every {row_label.lower()} and {col_label.lower()} combination, {BASE}",
              headline,
              f"Source: FMI analysis. Cell price = combined value (US$) / combined volume (MT); "
              f"shading shows the index vs the {REGION} average of US$ {fmt(tot)}/MT.")

    # ---- left: heatmap
    rect(s, 0.3, 1.65, 8.25, 3.85, PANEL)
    panel_title(s, 0.42, 1.7, 8.0, f"{BASE} Price Matrix (US$/MT), Shaded by Index vs {REGION} Average")
    header = [row_label] + col_order + ["Row Avg."]
    rows, fills = [], []
    for r in row_order:
        rows.append([r] + [fmt(M.loc[r, c]) for c in col_order] + [fmt(rowavg[r])])
        fills.append([WHITE] + [index_fill(I.loc[r, c]) for c in col_order] + [PANEL])
    rows.append(["Column Avg."] + [fmt(colavg[c]) for c in col_order] + [fmt(tot)])
    fills.append([PANEL] * (len(col_order) + 2))
    ncol = len(col_order)
    first_w, last_w = 1.7, 0.66
    cw = (8.05 - first_w - last_w) / ncol
    table(s, 0.4, 2.02, 8.05, 3.1, header, rows, [first_w] + [cw] * ncol + [last_w],
          fills=fills, font=9, head_font=8)
    # legend chips
    lx = 0.42
    for col, txt in [(BELOW[1], "Index below 97.5"), (BELOW[0], "97.5 to 99"), (PARITY, "Parity: 99 to 101"),
                     (ABOVE[0], "101 to 102.5"), (ABOVE[1], "Index above 102.5")]:
        rect(s, lx, 5.22, 0.2, 0.16, col, line=GRID)
        tb(s, lx + 0.23, 5.16, 1.4, 0.26, txt, size=8.5, color=MUTED)
        lx += 1.6

    # ---- right: range chart
    rect(s, 8.7, 1.65, 4.33, 3.85, PANEL)
    panel_title(s, 8.82, 1.7, 4.1, f"Price Range Across All {col_label} Sub-Types, {BASE}")
    tb(s, 8.82, 1.96, 4.1, 0.24, "Bar spans lowest to highest price (US$/MT) at the finest data level.",
       size=8.5, italic=True, color=MUTED)
    rr = rng.iloc[::-1]
    cd = CategoryChartData()
    cd.categories = list(rr.index)
    floor = int(rng["min"].min() // 2000 * 2000)
    cd.add_series("Base", [round(x) for x in rr["min"]])
    cd.add_series("Range", [round(b - a) for a, b in zip(rr["min"], rr["max"])])
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(8.75), Inches(2.2), Inches(4.23), Inches(3.25), cd)
    c = gf.chart
    c.has_legend = False
    skip(c)
    pl = c.plots[0]
    pl.gap_width = 55
    pl.overlap = 100
    base_s, rng_s = pl.series[0], pl.series[1]
    base_s.format.fill.background()
    rng_s.format.fill.solid()
    rng_s.format.fill.fore_color.rgb = TEAL
    for i, (a, b) in enumerate(zip(rr["min"], rr["max"])):
        dl = rng_s.points[i].data_label
        dl.has_text_frame = True
        dl.text_frame.text = f"{fmt(a)} to {fmt(b)}"
        dl.position = XL_LABEL_POSITION.INSIDE_BASE
        for p in dl.text_frame.paragraphs:
            for r in p.runs:
                r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(8.5), True, WHITE, FONT
    va = c.value_axis
    va.minimum_scale = floor
    va.maximum_scale = int((rng["max"].max() // 2000 + 1) * 2000)
    va.major_unit = 2000
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GRID
    va.format.line.fill.background()
    va.tick_labels.number_format = '#,##0' if not MASK else '"XX"'
    va.tick_labels.number_format_is_linked = False
    style_axis_text(va, size=8)
    style_axis_text(c.category_axis, size=8.5, color=INK)
    c.category_axis.format.line.color.rgb = GRID

    insight_cards(s, 5.6, 1.35, [
        ("Premium Pockets",
         f"{best[0]} positioned on {best[1]} carry the strongest premium at an index of {I.loc[best]:.1f}, "
         f"which signals room for value-added claims in this combination."),
        ("Value Pockets",
         f"{worst[0]} for {worst[1]} is the lowest priced combination at an index of {I.loc[worst]:.1f}, "
         f"sitting {pct(1 - I.loc[worst]/100, sign=False)} below the {REGION} average of US$ {fmt(tot)}/MT."),
        ("Within-Type Spread",
         f"{widest} show the widest internal price spread at {pct(widest_pct, sign=False)}, ranging from "
         f"US$ {fmt(rng.loc[widest, 'min'])}/MT to US$ {fmt(rng.loc[widest, 'max'])}/MT across sub-types."),
    ])
    return s


# =================================================================
# SLIDE 3: COUNTRY BENCHMARKING
# =================================================================
def slide_country(prs, v, q, segment, label, plural):
    regions = [r for r in v.index.get_level_values("Region").unique() if r != REGION]
    recs = []
    for c in regions:
        try:
            P = segment_prices(v, q, c, segment, "Sub-segment")
        except KeyError:
            continue
        a = region_total(v, q, c, segment)
        recs.append(dict(country=c, mn=P[BASE].min(), mx=P[BASE].max(), avg=a[BASE],
                         g=cagr(a[BASE], a[END], END - BASE),
                         top=P[BASE].idxmax()))
    D = pd.DataFrame(recs).set_index("country")
    D = D[D.index != "Other European Countries"].sort_values("avg", ascending=False)
    eut = region_total(v, q, REGION, segment)
    eu = eut[BASE]
    eu_g = cagr(eut[BASE], eut[END], END - BASE)
    D["idx"] = D["avg"] / eu * 100

    top, bot = D.index[0], D.index[-1]
    headline = (f"{top} is the most premium {REGION} market at US$ {fmt(D.loc[top, 'avg'])}/MT in {BASE}, "
                f"{pct(D.loc[top, 'idx']/100 - 1)} above the regional average, while {bot} is the lowest priced.")
    s = frame(prs,
              f"{MARKET} Market {REGION} Pricing Assessment: Country Benchmarking",
              f"Country average price and price range across {plural.lower()}, {BASE} (US$/MT)",
              headline,
              f"Source: FMI analysis. Range = lowest to highest {label.lower()} price in each country; "
              f"countries ranked by average price. Other European Countries excluded from ranking.")

    rect(s, 0.3, 1.65, 8.25, 3.85, PANEL)
    panel_title(s, 0.42, 1.7, 8.0, f"{BASE} Price Range Across {plural} by Country (US$/MT), Ranked by Average")
    cd = CategoryChartData()
    cd.categories = list(D.index)
    floor = int(D["mn"].min() // 1000 * 1000)
    cd.add_series("Base", [round(x) for x in D["mn"]])
    cd.add_series("Range", [round(b - a) for a, b in zip(D["mn"], D["mx"])])
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(0.38), Inches(2.18), Inches(8.1), Inches(3.3), cd)
    c = gf.chart
    c.has_legend = False
    pl = c.plots[0]
    pl.gap_width = 40
    pl.overlap = 100
    pl.series[0].format.fill.background()
    rs = pl.series[1]
    rs.format.fill.solid()
    rs.format.fill.fore_color.rgb = TEAL
    for i, cn in enumerate(D.index):
        if D.loc[cn, "idx"] < 100:
            pt = rs.points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = hexrgb("5B7DB1")
    va = c.value_axis
    va.minimum_scale = floor
    va.maximum_scale = int((D["mx"].max() // 1000 + 1) * 1000)
    va.major_unit = 1000
    va.tick_labels.number_format = '#,##0' if not MASK else '"XX"'
    va.tick_labels.number_format_is_linked = False
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GRID
    va.format.line.fill.background()
    # show real price on axis by offsetting the number format is not possible; label via custom text
    style_axis_text(va, size=8.5)
    ca = c.category_axis
    style_axis_text(ca, size=8.5, color=INK)
    ca.format.line.color.rgb = GRID
    ca.tick_labels.font.size = Pt(8.5)
    # axis: write custom price ticks on the left as text so the scale reads in real US$/MT
    tb(s, 0.42, 1.95, 7.9, 0.24,
       f"Bar spans the lowest to highest {label.lower()} price in each country; label = country price index "
       f"(Europe = 100). Teal = above {REGION} average, blue = below.", size=8.5, italic=True, color=MUTED)
    for i, cn in enumerate(D.index):
        dl = rs.points[i].data_label
        dl.has_text_frame = True
        dl.text_frame.text = f"{D.loc[cn, 'idx']:.0f}"
        dl.position = XL_LABEL_POSITION.INSIDE_END
        for p in dl.text_frame.paragraphs:
            for r in p.runs:
                r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(8), True, WHITE, FONT

    # right: premium / discount table
    rect(s, 8.7, 1.65, 4.33, 3.85, PANEL)
    panel_title(s, 8.82, 1.7, 4.1, "Premium and Discount Markets")
    sel = list(D.index[:5]) + list(D.index[-5:])
    rows, fills = [], []
    for cn in sel:
        rows.append([cn, fmt(D.loc[cn, "avg"]), f"{D.loc[cn, 'idx']:.1f}", pct(D.loc[cn, "g"], 2, sign=False)])
        fills.append([None, None, index_fill(D.loc[cn, "idx"]), None])
    table(s, 8.8, 2.02, 4.13, 3.38, ["Country", f"Avg. {BASE}", "Index", f"CAGR {BASE} to {END}"],
          rows, [1.55, 0.86, 0.72, 1.0], fills=fills, font=9, head_font=9)

    spread_c = D["avg"].max() / D["avg"].min() - 1
    fastest = D["g"].idxmax()
    within = (D["mx"] / D["mn"] - 1)
    wid = within.idxmax()
    insight_cards(s, 5.6, 1.35, [
        ("Country Spread",
         f"Average prices across the ranked countries differ by {pct(spread_c, sign=False)} in {BASE}, "
         f"from US$ {fmt(D['avg'].min())}/MT in {bot} to US$ {fmt(D['avg'].max())}/MT in {top}."),
        ("Growth Leader",
         f"{fastest} records the fastest average price growth at a {pct(D.loc[fastest, 'g'], 2, sign=False)} CAGR "
         f"between {BASE} and {END}, against {pct(eu_g, 2, sign=False)} for {REGION} as a whole."),
        ("Mix Exposure",
         f"{wid} shows the widest gap between its cheapest and most expensive {label.lower()} at "
         f"{pct(within[wid], sign=False)}, so product mix shifts move its average price the most."),
    ])
    return s


if __name__ == "__main__":
    v, q = load(DATA)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide_single(prs, v, q, "Product Type", "Product Type")
    slide_cross(prs, v, q, "Product Type × Primary Product Function / Benefit", "Product Type",
                "Primary Product Function / Benefit", "Primary Function")
    slide_country(prs, v, q, "Product Type", "Product Type", "Product Types")
    prs.save(OUT)
    print("saved", OUT)
