"""
Advanced / composite visual concepts  -  standalone, touches nothing in the project.

Everything here is a COMBINATION visual (dashboard, flow, radial, matrix with
in-cell charts ...) rather than a single bar/line, built from the formatted
Performance Knitted Fabrics dataset.

Run:  python chart_options/build_advanced_visuals.py
Out:  chart_options/Advanced_Visuals_Performance_Knitted_Fabrics.pptx
Set DARK = False below for a white-background version.
"""
from __future__ import annotations

import math
from pathlib import Path

import pandas as pd
from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

DARK = True
ROOT = Path(__file__).resolve().parent.parent
DATA = (ROOT / "Input" / "Performance_Knitted_Fabrics_Market" /
        "Global_Performance_Knitted_Fabrics_Market_Formatted_Dataset.xlsx")
OUT = Path(__file__).resolve().parent / "Advanced_Visuals_Performance_Knitted_Fabrics.pptx"

if DARK:
    BG, PANEL, GRID, TXT, SUB = "07182B", "0F2A44", "24435F", "F2F6FA", "8FA8C0"
    ACC = ["2EC4F1", "1FD1B0", "FFB703", "FF6B57", "A78BFA", "9BDE5C", "F472B6", "5B8DEF"]
else:
    BG, PANEL, GRID, TXT, SUB = "FFFFFF", "F5F8FA", "E1E7ED", "0B2E4A", "6C7A89"
    ACC = ["0E72AE", "008189", "A86F00", "E44D3A", "6F3CC2", "5E9B2E", "C2337F", "3D85C6"]
YEARS = [str(y) for y in range(2021, 2037)]


def rgb(h): return RGBColor.from_string(h)


# ---------------------------------------------------------------- data ----
D = pd.read_excel(DATA, sheet_name="Master Sheet", header=5)
D.columns = [str(c) for c in D.columns]
D = D.dropna(subset=["Region"])


def seg(region, segment, unit="Value", by="Sub-segment"):
    s = D[(D.Region == region) & (D.Segment == segment) & (D.Unit == unit)]
    return s.groupby(by, sort=False)[YEARS].sum()


def cagr(a, b, n=10): return (b / a) ** (1 / n) - 1 if a > 0 else 0.0


REG = seg("Global", "Region")
REGV = seg("Global", "Region", "Volume")
MAT = seg("Global", "Material")                    # parents
MAT_LEAF = seg("Global", "Material", by="Sub-segment 1")
MAT_CHILD = D[(D.Region == "Global") & (D.Segment == "Material") & (D.Unit == "Value")][
    ["Sub-segment", "Sub-segment 1", "2026", "2036"]]
END = seg("Global", "End-Use Application")
FUNC = seg("Global", "Primary Performance Function")
X = D[(D.Region == "Global") & (D.Segment == "End-Use Application Vs Material") & (D.Unit == "Value")]
FLOW = X.groupby(["Sub-segment", "Sub-segment 1"], sort=False)["2036"].sum().unstack().fillna(0)
GT = REG.sum()
GV = REGV.sum()
REGION_OF = {}
CTRY = {}
for r in REG.index:
    c = D[(D.Region == r) & (D.Segment == "Country") & (D.Unit == "Value")]
    for _, row in c.iterrows():
        CTRY[row["Sub-segment"]] = row[YEARS].astype(float)
        REGION_OF[row["Sub-segment"]] = r
REG_COL = {r: ACC[i % len(ACC)] for i, r in enumerate(REG.index)}


def short(n, k=24): return n if len(n) <= k else n[: k - 1].rstrip() + "…"


# ------------------------------------------------------------- drawing ----
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
IN = Inches


def tb(s, x, y, w, h, text, size=12, bold=False, color=None, align=PP_ALIGN.LEFT,
       anchor=MSO_ANCHOR.TOP, italic=False):
    t = s.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = t.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = IN(0.03)
    tf.margin_top = tf.margin_bottom = IN(0.01)
    for i, ln in enumerate(text if isinstance(text, list) else [text]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.size, r.font.bold, r.font.italic = Pt(size), bold, italic
        r.font.color.rgb = rgb(color or TXT)
        r.font.name = "Calibri"
    return t


def rect(s, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, lw=0.75, alpha=None):
    r = s.shapes.add_shape(shape, int(x), int(y), int(w), int(h))
    if fill:
        r.fill.solid()
        r.fill.fore_color.rgb = rgb(fill)
        if alpha is not None:
            set_alpha(r, alpha)
    else:
        r.fill.background()
    if line:
        r.line.color.rgb = rgb(line)
        r.line.width = Pt(lw)
    else:
        r.line.fill.background()
    r.shadow.inherit = False
    return r


def set_alpha(shape, pct):
    clr = shape._element.spPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
    a = etree.SubElement(clr, qn("a:alpha"))
    a.set("val", str(int(pct * 1000)))


def put(shape, text, size=10, color=None, bold=False, align=PP_ALIGN.CENTER):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = IN(0.02)
    for i, ln in enumerate(text if isinstance(text, list) else [text]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.size, r.font.bold = Pt(size), bold
        r.font.color.rgb = rgb(color or TXT)
        r.font.name = "Calibri"


def ln(s, x1, y1, x2, y2, color=None, w=1.0, dash=False):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, int(x1), int(y1), int(x2), int(y2))
    c.line.color.rgb = rgb(color or GRID)
    c.line.width = Pt(w)
    if dash:
        c.line.dash_style = 4
    return c


def poly(s, pts, fill, line=None, alpha=None, lw=0.75):
    fb = s.shapes.build_freeform(int(pts[0][0]), int(pts[0][1]), scale=1.0)
    fb.add_line_segments([(int(x), int(y)) for x, y in pts[1:]], close=True)
    sh = fb.convert_to_shape()
    sh.fill.solid()
    sh.fill.fore_color.rgb = rgb(fill)
    if alpha is not None:
        set_alpha(sh, alpha)
    if line:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def pol(cx, cy, r, a):
    t = math.radians(a)
    return cx + r * math.sin(t), cy - r * math.cos(t)


def sector(s, cx, cy, r0, r1, a0, a1, fill, line=None, alpha=None, lw=1.0):
    n = max(2, int(abs(a1 - a0) / 2.5))
    outer = [pol(cx, cy, r1, a0 + (a1 - a0) * i / n) for i in range(n + 1)]
    inner = [pol(cx, cy, r0, a1 - (a1 - a0) * i / n) for i in range(n + 1)]
    return poly(s, outer + inner, fill, line, alpha, lw)


def slide(title, sub, footer):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, prs.slide_width, prs.slide_height, BG)
    tb(s, IN(0.5), IN(0.25), IN(12.3), IN(0.6), title, 26, True)
    tb(s, IN(0.5), IN(0.86), IN(12.3), IN(0.4), sub, 13, color=SUB)
    rect(s, IN(0.5), IN(1.3), IN(0.9), IN(0.05), ACC[0])
    tb(s, IN(0.5), IN(7.08), IN(12.3), IN(0.3), footer, 9.5, color=SUB, italic=True)
    return s


def style(ch, legend=False, pos=XL_LEGEND_POSITION.BOTTOM, size=10):
    ch.font.size, ch.font.name = Pt(size), "Calibri"
    ch.font.color.rgb = rgb(TXT)
    ch.has_legend = legend
    if legend:
        ch.legend.position = pos
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(size)
        ch.legend.font.color.rgb = rgb(TXT)


def axes(ch, val=True, grid=True, fmt="#,##0"):
    ca, va = ch.category_axis, ch.value_axis
    ca.format.line.color.rgb = rgb(GRID)
    ca.has_major_gridlines = False
    ca.tick_labels.font.size = Pt(9)
    ca.tick_labels.font.color.rgb = rgb(SUB)
    va.has_major_gridlines = grid
    if grid:
        va.major_gridlines.format.line.color.rgb = rgb(GRID)
    va.format.line.fill.background()
    va.tick_labels.number_format = fmt
    va.tick_labels.number_format_is_linked = False
    va.tick_labels.font.size = Pt(9)
    va.tick_labels.font.color.rgb = rgb(SUB)
    va.visible = val


def chart(s, kind, cd, x, y, w, h):
    return s.shapes.add_chart(kind, int(x), int(y), int(w), int(h), cd).chart


def title_of(ch, text, size=11):
    ch.has_title = True
    ch.chart_title.text_frame.text = text
    r = ch.chart_title.text_frame.paragraphs[0].runs[0]
    r.font.size, r.font.bold, r.font.color.rgb = Pt(size), True, rgb(TXT)


def fill_s(sr, col):
    sr.format.fill.solid()
    sr.format.fill.fore_color.rgb = rgb(col)


def smooth(t): return t * t * (3 - 2 * t)


# ============================================================== COVER =====
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, prs.slide_width, prs.slide_height, BG)
rect(s, IN(0.8), IN(2.2), IN(1.3), IN(0.08), ACC[0])
tb(s, IN(0.8), IN(2.4), IN(11.5), IN(1.2), "Beyond bar charts: combination visuals", 40, True)
tb(s, IN(0.8), IN(3.65), IN(11.5), IN(1.2),
   ["Eleven composite concepts built on the Global Performance Knitted Fabrics dataset",
    "Dashboard · flow · radial · matrix with in-cell charts · packing · bump · stream"], 16, color=SUB)
for i, c in enumerate(ACC):
    rect(s, IN(0.8) + IN(0.55) * i, IN(5.6), IN(0.45), IN(0.45), c)

# ======================================================== 1. DASHBOARD ====
s = slide("Executive dashboard: size, momentum, mix and growth drivers in one frame",
          f"Global performance knitted fabrics - US$ {GT['2026']:,.0f} Mn in 2026 heading to US$ {GT['2036']:,.0f} Mn by 2036",
          "COMBINES: KPI cards + column/line combo + doughnut + ranked bar  |  REPLACES: four separate summary slides")
cards = [("2026 value", f"US$ {GT['2026']:,.0f} Mn", ACC[0]),
         ("2036 value", f"US$ {GT['2036']:,.0f} Mn", ACC[1]),
         ("CAGR 26-36", f"{cagr(GT['2026'], GT['2036'])*100:.1f}%", ACC[2]),
         ("Volume 2036", f"{GV['2036']:,.0f} Mn Sq. Yd", ACC[4])]
cw = (IN(12.33) - IN(0.45)) / 4
for i, (t, v, c) in enumerate(cards):
    x = IN(0.5) + (cw + IN(0.15)) * i
    rect(s, x, IN(1.55), cw, IN(0.95), PANEL)
    rect(s, x, IN(1.55), IN(0.07), IN(0.95), c)
    tb(s, x + IN(0.2), IN(1.6), cw - IN(0.3), IN(0.3), t.upper(), 10, True, SUB)
    tb(s, x + IN(0.2), IN(1.88), cw - IN(0.3), IN(0.55), v, 24, True, c)
# combo (column + YoY line)
yoy = [None] + [round((GT[YEARS[i]] / GT[YEARS[i - 1]] - 1) * 100, 1) for i in range(1, 16)]
cd = CategoryChartData()
cd.categories = YEARS
cd.add_series("Value (US$ Mn)", [round(v) for v in GT[YEARS]])
cd.add_series("YoY growth (%)", yoy)
ch = chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cd, IN(0.4), IN(2.65), IN(6.6), IN(4.35))
style(ch, True)
axes(ch)
fill_s(ch.plots[0].series[0], ACC[0])
ch.plots[0].gap_width = 45
title_of(ch, "Value and YoY growth")
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
ns = {"c": C}
plot = ch._chartSpace.find(".//c:plotArea", ns)
bar = plot.find("c:barChart", ns)
ser2 = bar.findall("c:ser", ns)[1]
bar.remove(ser2)
for el in ser2.findall("c:spPr", ns) + ser2.findall("c:invertIfNegative", ns):
    ser2.remove(el)
sp = etree.fromstring(f'<c:spPr xmlns:c="{C}" xmlns:a="{A}"><a:ln w="31750"><a:solidFill>'
                      f'<a:srgbClr val="{ACC[2]}"/></a:solidFill></a:ln></c:spPr>')
ser2.insert(list(ser2).index(ser2.find("c:tx", ns)) + 1, sp)
mk = etree.fromstring(f'<c:marker xmlns:c="{C}" xmlns:a="{A}"><c:symbol val="circle"/><c:size val="6"/>'
                      f'<c:spPr><a:solidFill><a:srgbClr val="{ACC[2]}"/></a:solidFill></c:spPr></c:marker>')
ser2.insert(list(ser2).index(sp) + 1, mk)
lc = etree.Element(f"{{{C}}}lineChart")
etree.SubElement(lc, f"{{{C}}}grouping").set("val", "standard")
etree.SubElement(lc, f"{{{C}}}varyColors").set("val", "0")
lc.append(ser2)
etree.SubElement(lc, f"{{{C}}}marker").set("val", "1")
etree.SubElement(lc, f"{{{C}}}axId").set("val", "50010")
etree.SubElement(lc, f"{{{C}}}axId").set("val", "50020")
plot.insert(list(plot).index(bar) + 1, lc)
cat2 = etree.fromstring(
    f'<c:catAx xmlns:c="{C}"><c:axId val="50010"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
    f'<c:delete val="1"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
    f'<c:tickLblPos val="nextTo"/><c:crossAx val="50020"/><c:crosses val="autoZero"/><c:auto val="1"/>'
    f'<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>')
val2 = etree.fromstring(
    f'<c:valAx xmlns:c="{C}" xmlns:a="{A}"><c:axId val="50020"/><c:scaling><c:orientation val="minMax"/>'
    f'<c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="r"/>'
    f'<c:numFmt formatCode="0&quot;%&quot;" sourceLinked="0"/><c:majorTickMark val="none"/>'
    f'<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
    f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="{SUB}"/>'
    f'</a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
    f'<c:crossAx val="50010"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>')
last_ax = [e for e in plot if e.tag.endswith("Ax")][-1]
i_ = list(plot).index(last_ax) + 1
plot.insert(i_, cat2)
plot.insert(i_ + 1, val2)
# doughnut: regional share 2036
cd = CategoryChartData()
cd.categories = [short(r, 18) for r in REG.index]
cd.add_series("2036", [round(v) for v in REG["2036"]])
ch = chart(s, XL_CHART_TYPE.DOUGHNUT, cd, IN(7.05), IN(2.65), IN(2.95), IN(4.35))
style(ch, True, size=8)
for i in range(len(REG)):
    p = ch.plots[0].series[0].points[i]
    p.format.fill.solid()
    p.format.fill.fore_color.rgb = rgb(ACC[i % len(ACC)])
    p.format.line.color.rgb = rgb(BG)
ch.plots[0].has_data_labels = True
dl = ch.plots[0].data_labels
dl.show_percentage, dl.show_value = True, False
dl.number_format, dl.number_format_is_linked = "0%", False
dl.font.size, dl.font.bold, dl.font.color.rgb = Pt(9), True, rgb(BG if DARK else "FFFFFF")
title_of(ch, "2036 regional share")
# ranked bar: end-use CAGR
ec = {n: cagr(END.loc[n, "2026"], END.loc[n, "2036"]) * 100 for n in END.index}
o = sorted(ec, key=ec.get)
cd = CategoryChartData()
cd.categories = [short(n, 20) for n in o]
cd.add_series("CAGR", [round(ec[n], 1) for n in o])
ch = chart(s, XL_CHART_TYPE.BAR_CLUSTERED, cd, IN(10.05), IN(2.65), IN(2.9), IN(4.35))
style(ch)
axes(ch, val=False, grid=False)
fill_s(ch.plots[0].series[0], ACC[1])
ch.plots[0].gap_width = 40
ch.plots[0].has_data_labels = True
d = ch.plots[0].data_labels
d.number_format, d.number_format_is_linked = '0.0"%"', False
d.font.size, d.font.color.rgb, d.font.bold = Pt(9), rgb(TXT), True
d.position = XL_LABEL_POSITION.OUTSIDE_END
title_of(ch, "Growth by application")

# ============================================================ 2. SANKEY ===
mats = FLOW.sum(axis=0).sort_values(ascending=False).index.tolist()
ends = FLOW.sum(axis=1).sort_values(ascending=False).index.tolist()
tot = FLOW.values.sum()
s = slide("Where each fibre ends up: material-to-application flow, 2036",
          f"US$ {tot:,.0f} Mn of 2036 demand traced from six materials into seven end-use applications",
          "COMBINES: stacked bar + flow ribbons  |  REPLACES: the End-Use × Material table and side-by-side bars")
top, H, gap = IN(1.6), IN(5.25), IN(0.1)
scale = (H - gap * 7) / tot
XL_, XR_, NW = IN(3.0), IN(9.75), IN(0.2)
y = top + (H - (scale * tot + gap * (len(mats) - 1))) / 2
node_l = {}
for i, m in enumerate(mats):
    hh = FLOW[m].sum() * scale
    node_l[m] = [y, y]
    rect(s, XL_, y, NW, hh, ACC[i % len(ACC)])
    tb(s, IN(0.5), y + hh / 2 - IN(0.2), IN(2.45), IN(0.4),
       [f"{short(m, 24)}", f"US$ {FLOW[m].sum():,.0f} Mn"], 10.5, True, TXT, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
    y += hh + gap
y = top + (H - (scale * tot + gap * (len(ends) - 1))) / 2
node_r = {}
for e in ends:
    hh = FLOW.loc[e].sum() * scale
    node_r[e] = [y, y]
    rect(s, XR_, y, NW, hh, SUB)
    tb(s, XR_ + NW + IN(0.1), y + hh / 2 - IN(0.2), IN(2.7), IN(0.4),
       [short(e, 30), f"US$ {FLOW.loc[e].sum():,.0f} Mn"], 10.5, True, TXT, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    y += hh + gap
for i, m in enumerate(mats):
    for e in ends:
        v = FLOW.loc[e, m]
        if v <= 0:
            continue
        hh = v * scale
        y0, y1 = node_l[m][1], node_r[e][1]
        node_l[m][1] += hh
        node_r[e][1] += hh
        n = 24
        xs = [XL_ + NW + (XR_ - XL_ - NW) * k / n for k in range(n + 1)]
        top_e = [(x, y0 + (y1 - y0) * smooth(k / n)) for k, x in enumerate(xs)]
        bot_e = [(x, yy + hh) for x, yy in reversed(top_e)]
        poly(s, top_e + bot_e, ACC[i % len(ACC)], alpha=45 if DARK else 50)

# ======================================================= 3. STREAMGRAPH ===
s = slide("The market stream: every application's contribution, 2021-2036",
          "Centred stream of global value by end-use application; thickness = US$ Mn",
          "COMBINES: stacked area + centred baseline + direct labels  |  REPLACES: stacked column / stacked area")
names = END.sort_values("2036", ascending=False).index.tolist()
order = []
for i, n in enumerate(names):
    (order.insert(0, n) if i % 2 == 0 else order.append(n))
x0, x1 = IN(0.9), IN(10.2)
cy, mxh = IN(4.2), IN(4.4)
mx = max(END[y].sum() for y in YEARS)
sc = mxh / mx
step = 8
xsn = [x0 + (x1 - x0) * k / ((len(YEARS) - 1) * step) for k in range((len(YEARS) - 1) * step + 1)]


def interp(vals):
    out = []
    for yi in range(len(vals) - 1):
        for k in range(step):
            out.append(vals[yi] + (vals[yi + 1] - vals[yi]) * smooth(k / step))
    out.append(vals[-1])
    return out


curves = {n: interp([float(END.loc[n, y]) for y in YEARS]) for n in order}
totals = [sum(curves[n][i] for n in order) for i in range(len(xsn))]
base = [cy - t * sc / 2 for t in totals]
for j, n in enumerate(order):
    topl = [(xsn[i], base[i]) for i in range(len(xsn))]
    nb = [base[i] + curves[n][i] * sc for i in range(len(xsn))]
    bot = [(xsn[i], nb[i]) for i in reversed(range(len(xsn)))]
    poly(s, topl + bot, ACC[j % len(ACC)], line=BG, alpha=88, lw=0.5)
    mid_s = (base[0] + nb[0]) / 2
    mid_e = (base[-1] + nb[-1]) / 2
    tb(s, x1 + IN(0.1), mid_e - IN(0.18), IN(2.8), IN(0.36),
       f"{short(n, 28)}  {END.loc[n, '2036']:,.0f}", 10.5, True, ACC[j % len(ACC)], anchor=MSO_ANCHOR.MIDDLE)
    base = nb
for yi in (0, 5, 10, 15):
    xx = x0 + (x1 - x0) * yi / 15
    ln(s, xx, IN(1.75), xx, IN(6.75), GRID, 0.75, True)
    tb(s, xx - IN(0.6), IN(6.78), IN(1.2), IN(0.28), YEARS[yi], 11, True, SUB, PP_ALIGN.CENTER)
    tb(s, xx - IN(0.9), IN(1.45), IN(1.8), IN(0.3), f"US$ {GT[YEARS[yi]]:,.0f} Mn", 10.5, True, TXT, PP_ALIGN.CENTER)

# ========================================================= 4. SUNBURST ===
s = slide("Inside the material mix: two-ring sunburst, 2036",
          "Inner ring = material family, outer ring = specific fibre; angle = share of US$ value",
          "COMBINES: doughnut + drill-down ring + legend table  |  REPLACES: one pie plus a separate sub-segment table")
cx, cy_ = IN(4.1), IN(4.3)
tot36 = MAT["2036"].sum()
fam = MAT["2036"].sort_values(ascending=False)
a = 0.0
for i, (m, v) in enumerate(fam.items()):
    span = 360 * v / tot36
    col = ACC[i % len(ACC)]
    sector(s, cx, cy_, IN(1.0), IN(1.95), a, a + span, col, BG, lw=1.5)
    kids = MAT_CHILD[MAT_CHILD["Sub-segment"] == m]
    ka = a
    for k, (_, kr) in enumerate(kids.iterrows()):
        ksp = span * float(kr["2036"]) / v
        sector(s, cx, cy_, IN(2.0), IN(2.55), ka, ka + ksp, col, BG, alpha=100 - 18 * (k % 3) - (18 if len(kids) > 1 else 0), lw=1.5)
        ka += ksp
    if span > 22:
        mx_, my_ = pol(cx, cy_, IN(1.48), a + span / 2)
        tb(s, mx_ - IN(0.4), my_ - IN(0.15), IN(0.8), IN(0.3), f"{v/tot36*100:.0f}%", 11, True,
           BG if DARK else "FFFFFF", PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    a += span
tb(s, cx - IN(0.9), cy_ - IN(0.4), IN(1.8), IN(0.8), ["2036", f"US$ {tot36:,.0f} Mn"], 14, True, TXT,
   PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
yy = IN(1.65)
tb(s, IN(8.0), yy, IN(5), IN(0.3), "MATERIAL FAMILY / FIBRE   ·   US$ Mn 2036   ·   CAGR", 10, True, SUB)
yy += IN(0.38)
for i, (m, v) in enumerate(fam.items()):
    col = ACC[i % len(ACC)]
    rect(s, IN(8.0), yy + IN(0.05), IN(0.16), IN(0.16), col)
    c_ = cagr(MAT.loc[m, "2026"], MAT.loc[m, "2036"]) * 100
    tb(s, IN(8.25), yy, IN(3.2), IN(0.28), short(m, 34), 12, True, TXT)
    tb(s, IN(11.0), yy, IN(1.9), IN(0.28), f"{v:,.0f}  ·  {c_:.1f}%", 12, True, col, PP_ALIGN.RIGHT)
    yy += IN(0.3)
    kids = MAT_CHILD[MAT_CHILD["Sub-segment"] == m]
    if len(kids) > 1:
        for _, kr in kids.iterrows():
            tb(s, IN(8.55), yy, IN(2.6), IN(0.24), short(kr["Sub-segment 1"], 30), 10, color=SUB)
            tb(s, IN(11.0), yy, IN(1.9), IN(0.24), f"{kr['2036']:,.0f}", 10, color=SUB, align=PP_ALIGN.RIGHT)
            yy += IN(0.24)
    yy += IN(0.08)

# ======================================================= 5. RADIAL BAR ===
s = slide("Radial ranking: application size with growth tags, 2036",
          "Each ring = one end-use application; arc length = 2036 value (full 270° = largest)",
          "COMBINES: bar ranking + circular layout + CAGR chips  |  REPLACES: horizontal bar of end-use")
cx, cy_ = IN(4.3), IN(4.25)
names = END["2036"].sort_values(ascending=False).index.tolist()
vmax = END["2036"].max()
r0, stp, th = IN(0.75), IN(0.3), IN(0.23)
for i, n in enumerate(names):
    rr = r0 + stp * (len(names) - 1 - i)
    sector(s, cx, cy_, rr, rr + th, 0, 270, PANEL)
    ang = 270 * END.loc[n, "2036"] / vmax
    sector(s, cx, cy_, rr, rr + th, 0, ang, ACC[i % len(ACC)])
    tb(s, cx - IN(2.85), cy_ - rr - th + IN(0.0), IN(2.78), th, short(n, 30), 10, True, TXT,
       PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
tb(s, cx - IN(0.75), cy_ - IN(0.35), IN(1.5), IN(0.7), ["US$ Mn", f"{END['2036'].sum():,.0f}"], 13, True,
   TXT, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
yy = IN(1.8)
tb(s, IN(8.4), yy, IN(4.5), IN(0.3), "APPLICATION  ·  2036 value  ·  share  ·  CAGR", 10, True, SUB)
yy += IN(0.4)
for i, n in enumerate(names):
    col = ACC[i % len(ACC)]
    c_ = cagr(END.loc[n, "2026"], END.loc[n, "2036"]) * 100
    rect(s, IN(8.4), yy + IN(0.06), IN(0.16), IN(0.16), col)
    tb(s, IN(8.65), yy, IN(2.5), IN(0.3), short(n, 26), 11.5, True)
    tb(s, IN(10.9), yy, IN(0.9), IN(0.3), f"{END.loc[n,'2036']:,.0f}", 11.5, False, TXT, PP_ALIGN.RIGHT)
    tb(s, IN(11.8), yy, IN(0.55), IN(0.3), f"{END.loc[n,'2036']/END['2036'].sum()*100:.0f}%", 11.5, False, SUB, PP_ALIGN.RIGHT)
    chip = rect(s, IN(12.4), yy + IN(0.03), IN(0.55), IN(0.25), col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    put(chip, f"{c_:.1f}%", 9, BG if DARK else "FFFFFF", True)
    yy += IN(0.5)

# =================================================== 6. OPPORTUNITY MAP ===
rows = []
for c, ser in CTRY.items():
    if c.startswith("Rest of") or c.startswith("Other"):
        continue
    rows.append((c, REGION_OF[c], float(ser["2036"]), cagr(float(ser["2026"]), float(ser["2036"])) * 100,
                 float(ser["2036"] - ser["2026"])))
cdf = pd.DataFrame(rows, columns=["c", "r", "v36", "g", "abs"])
s = slide("Opportunity map: where size and speed meet - 25 countries",
          "x = 2036 value (log scale), y = CAGR 2026-36, bubble = absolute US$ gain, colour = region",
          "COMBINES: scatter + quadrant logic + bubble size + region colour  |  REPLACES: country CAGR table")
PX0, PX1, PY0, PY1 = IN(1.0), IN(10.2), IN(1.75), IN(6.7)
lx0, lx1 = math.log10(cdf.v36.min() * 0.8), math.log10(cdf.v36.max() * 1.25)
gy0, gy1 = cdf.g.min() - 0.6, cdf.g.max() + 0.6
mxv = math.log10(cdf.v36.median())
myg = cdf.g.median()
X_ = lambda v: PX0 + (PX1 - PX0) * (math.log10(v) - lx0) / (lx1 - lx0)
Y_ = lambda g: PY1 - (PY1 - PY0) * (g - gy0) / (gy1 - gy0)
xm, ym = PX0 + (PX1 - PX0) * (mxv - lx0) / (lx1 - lx0), Y_(myg)
rect(s, xm, PY0, PX1 - xm, ym - PY0, PANEL, alpha=100)
ln(s, xm, PY0, xm, PY1, SUB, 1, True)
ln(s, PX0, ym, PX1, ym, SUB, 1, True)
ln(s, PX0, PY1, PX1, PY1, GRID, 1)
for txt, xx, yy_, al in (("PRIORITISE  · large & fast", PX1 - IN(0.1), PY0 + IN(0.05), PP_ALIGN.RIGHT),
                         ("EMERGING  · small & fast", PX0 + IN(0.1), PY0 + IN(0.05), PP_ALIGN.LEFT),
                         ("DEFEND  · large & slower", PX1 - IN(0.1), PY1 - IN(0.32), PP_ALIGN.RIGHT),
                         ("MONITOR  · small & slower", PX0 + IN(0.1), PY1 - IN(0.32), PP_ALIGN.LEFT)):
    tb(s, xx - IN(3) if al == PP_ALIGN.RIGHT else xx, yy_, IN(3), IN(0.3), txt, 10, True, SUB, al)
amax = cdf["abs"].max()
for _, r in cdf.sort_values("abs", ascending=False).iterrows():
    d = IN(0.2) + IN(0.75) * math.sqrt(r["abs"] / amax)
    cxx, cyy = X_(r.v36), Y_(r.g)
    rect(s, cxx - d / 2, cyy - d / 2, d, d, REG_COL[r.r], BG, MSO_SHAPE.OVAL, 1, alpha=80)
for _, r in cdf.sort_values("abs", ascending=False).head(14).iterrows():
    d = IN(0.2) + IN(0.75) * math.sqrt(r["abs"] / amax)
    tb(s, X_(r.v36) + d / 2, Y_(r.g) - IN(0.11), IN(1.4), IN(0.22), short(r.c, 16), 9, True, TXT)
tb(s, PX0, PY1 + IN(0.05), PX1 - PX0, IN(0.28), "2036 value (US$ Mn, log scale)  →", 10, False, SUB, PP_ALIGN.CENTER)
tb(s, IN(0.1), PY0 - IN(0.3), IN(0.9), IN(0.3), "CAGR %", 10, False, SUB)
yy = IN(1.8)
tb(s, IN(10.6), yy, IN(2.4), IN(0.3), "REGION", 10, True, SUB)
for i, (r_, col) in enumerate(REG_COL.items()):
    rect(s, IN(10.6), yy + IN(0.42) + IN(0.36) * i, IN(0.17), IN(0.17), col, shape=MSO_SHAPE.OVAL)
    tb(s, IN(10.85), yy + IN(0.36) + IN(0.36) * i, IN(2.1), IN(0.3), short(r_, 22), 11, True)

# ============================================================ 7. BUMP ====
yrs_b = ["2021", "2024", "2027", "2030", "2033", "2036"]
allv = pd.DataFrame({c: s_ for c, s_ in CTRY.items() if not (c.startswith("Rest of") or c.startswith("Other"))}).T
rk = allv[yrs_b].rank(ascending=False, method="first")
sel = set(rk[rk["2021"] <= 10].index) | set(rk[rk["2036"] <= 10].index)
s = slide("Ranking race: who climbs the country league table, 2021 → 2036",
          "Rank of each country by market value; lines coloured by region, thick = moved up",
          "COMBINES: rank lines + numbered nodes + movement highlight  |  REPLACES: year-by-year country tables")
BX0, BX1, BY0, BY1 = IN(2.3), IN(10.3), IN(1.75), IN(6.85)
maxr = int(max(rk.loc[list(sel), yrs_b].max()))
Yr = lambda r: BY0 + (BY1 - BY0) * (r - 1) / max(1, maxr - 1)
Xt = lambda i: BX0 + (BX1 - BX0) * i / (len(yrs_b) - 1)
for i, y_ in enumerate(yrs_b):
    tb(s, Xt(i) - IN(0.5), IN(1.42), IN(1.0), IN(0.3), y_, 12, True, SUB, PP_ALIGN.CENTER)
for c in sorted(sel, key=lambda k: rk.loc[k, "2036"], reverse=True):
    col = REG_COL[REGION_OF[c]]
    up = rk.loc[c, "2021"] - rk.loc[c, "2036"]
    w = 4.5 if up >= 2 else 2.0
    for i in range(len(yrs_b) - 1):
        ln(s, Xt(i), Yr(rk.loc[c, yrs_b[i]]), Xt(i + 1), Yr(rk.loc[c, yrs_b[i + 1]]), col, w)
    for i, y_ in enumerate(yrs_b):
        dd = IN(0.3)
        n_ = rect(s, Xt(i) - dd / 2, Yr(rk.loc[c, y_]) - dd / 2, dd, dd, col, BG, MSO_SHAPE.OVAL, 1)
        put(n_, str(int(rk.loc[c, y_])), 8, BG if DARK else "FFFFFF", True)
    tb(s, BX0 - IN(1.95), Yr(rk.loc[c, "2021"]) - IN(0.12), IN(1.6), IN(0.24), short(c, 18), 10, True, TXT, PP_ALIGN.RIGHT)
    tb(s, BX1 + IN(0.28), Yr(rk.loc[c, "2036"]) - IN(0.12), IN(1.7), IN(0.24),
       f"{short(c, 14)}  {'▲' if up > 0 else ('▼' if up < 0 else '–')}{abs(int(up)) or ''}", 10, True, col)
yy = IN(1.8)
for i, (r_, col) in enumerate(REG_COL.items()):
    rect(s, IN(11.0), yy + IN(0.36) * i + IN(0.05), IN(0.15), IN(0.15), col, shape=MSO_SHAPE.OVAL)
    tb(s, IN(11.2), yy + IN(0.36) * i, IN(1.8), IN(0.26), short(r_, 18), 10, True, col)

# ========================================================= 8. SCORECARD ===
s = slide("Regional scorecard: bars, trend and growth inside one table",
          "Every row carries its own mini-charts so the table reads like a dashboard",
          "COMBINES: table + in-cell bars + sparkline + CAGR heat chip  |  REPLACES: the regional value table")
cols = [("REGION", 2.7), ("SHARE OF 2036 MARKET", 2.7), ("2026 → 2036 (US$ Mn)", 2.6),
        ("TREND 2021-36", 1.8), ("CAGR", 0.95), ("US$/SQ YD", 1.3)]
xs = [IN(0.5)]
for _, w in cols:
    xs.append(xs[-1] + IN(w))
for (t, _), x in zip(cols, xs):
    tb(s, x, IN(1.5), IN(2.5), IN(0.3), t, 9.5, True, SUB)
ln(s, IN(0.5), IN(1.82), xs[-1], IN(1.82), GRID, 1)
order = REG["2036"].sort_values(ascending=False).index
rowh = IN(0.66)
gmax = max(cagr(REG.loc[r, "2026"], REG.loc[r, "2036"]) for r in order)
gmin = min(cagr(REG.loc[r, "2026"], REG.loc[r, "2036"]) for r in order)
vmx = REG["2036"].max()
for i, r in enumerate(order):
    y0 = IN(1.92) + rowh * i
    if i % 2 == 0:
        rect(s, IN(0.5), y0, xs[-1] - IN(0.5), rowh - IN(0.04), PANEL)
    col = REG_COL[r]
    rect(s, xs[0] + IN(0.1), y0 + IN(0.22), IN(0.18), IN(0.18), col, shape=MSO_SHAPE.OVAL)
    tb(s, xs[0] + IN(0.38), y0 + IN(0.14), IN(2.3), IN(0.34), short(r, 24), 13, True, anchor=MSO_ANCHOR.MIDDLE)
    shr = REG.loc[r, "2036"] / GT["2036"]
    rect(s, xs[1], y0 + IN(0.22), IN(1.9), IN(0.2), GRID)
    rect(s, xs[1], y0 + IN(0.22), IN(1.9) * shr / 0.3, IN(0.2), col)
    tb(s, xs[1] + IN(1.95), y0 + IN(0.16), IN(0.7), IN(0.3), f"{shr*100:.1f}%", 11, True)
    for k, yk in enumerate(("2026", "2036")):
        w_ = IN(1.45) * REG.loc[r, yk] / vmx
        rect(s, xs[2], y0 + IN(0.09) + IN(0.23) * k, max(w_, IN(0.03)), IN(0.18), col if k else SUB)
        tb(s, xs[2] + w_ + IN(0.05), y0 + IN(0.04) + IN(0.23) * k, IN(0.9), IN(0.26),
           f"{REG.loc[r, yk]:,.0f}", 9.5, k == 1, TXT if k else SUB)
    sp = [float(REG.loc[r, y]) for y in YEARS]
    lo, hi = min(sp), max(sp)
    bw = IN(1.6) / 16
    for k, v in enumerate(sp):
        hh = IN(0.4) * (0.12 + 0.88 * (v - lo) / (hi - lo))
        rect(s, xs[3] + bw * k, y0 + IN(0.55) - hh, bw - IN(0.025), hh, col if k >= 5 else SUB)
    g = cagr(REG.loc[r, "2026"], REG.loc[r, "2036"])
    t = (g - gmin) / (gmax - gmin) if gmax > gmin else 0.5
    chip = rect(s, xs[4], y0 + IN(0.16), IN(0.8), IN(0.32), ACC[1] if t > 0.5 else ACC[0],
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, alpha=40 + 60 * t)
    put(chip, f"{g*100:.1f}%", 11, TXT if DARK else "0B2E4A", True)
    pr = REG.loc[r, "2036"] / REGV.loc[r, "2036"]
    tb(s, xs[5], y0 + IN(0.14), IN(1.2), IN(0.34), f"{pr:.2f}", 12, True, anchor=MSO_ANCHOR.MIDDLE)

# ====================================================== 9. CIRCLE PACK ====
s = slide("Country landscape: 2036 market packed by size and region",
          "Circle area = 2036 value; colour = region; all 31 countries / blocks in one picture",
          "COMBINES: treemap logic + region colour + labels  |  REPLACES: country bar list or country table")
items = sorted(((c, REGION_OF[c], float(CTRY[c]["2036"])) for c in CTRY), key=lambda t: -t[2])
PXa, PXb, PYa, PYb = IN(0.6), IN(9.2), IN(1.55), IN(6.95)
cxm, cym = (PXa + PXb) / 2, (PYa + PYb) / 2
kk = IN(0.02)
placed = []
for c, r, v in items:
    rad = kk * math.sqrt(v) * 1.0
    done = False
    for step_ in range(0, 4000):
        ang = step_ * 0.35
        dist = IN(0.012) * ang
        x = cxm + dist * math.cos(ang) * 1.45
        y = cym + dist * math.sin(ang)
        if x - rad < PXa or x + rad > PXb or y - rad < PYa or y + rad > PYb:
            continue
        if all((x - px) ** 2 + (y - py) ** 2 >= (rad + pr + IN(0.03)) ** 2 for _, px, py, pr, _ in placed):
            placed.append((c, x, y, rad, r))
            done = True
            break
    if not done:
        placed.append((c, cxm, cym, rad * 0.4, r))
for c, x, y, rad, r in placed:
    o = rect(s, x - rad, y - rad, rad * 2, rad * 2, REG_COL[r], BG, MSO_SHAPE.OVAL, 1.5, alpha=88)
    if rad > IN(0.33):
        put(o, [short(c, 14), f"{float(CTRY[c]['2036']):,.0f}"], 9 if rad < IN(0.5) else 11, BG if DARK else "FFFFFF", True)
yy = IN(1.7)
tb(s, IN(9.6), yy, IN(3.3), IN(0.3), "REGION  ·  2036 US$ Mn", 10, True, SUB)
for i, r_ in enumerate(REG["2036"].sort_values(ascending=False).index):
    rect(s, IN(9.6), yy + IN(0.45) + IN(0.42) * i, IN(0.2), IN(0.2), REG_COL[r_], shape=MSO_SHAPE.OVAL)
    tb(s, IN(9.9), yy + IN(0.4) + IN(0.42) * i, IN(2.1), IN(0.3), short(r_, 20), 11.5, True)
    tb(s, IN(11.9), yy + IN(0.4) + IN(0.42) * i, IN(1.0), IN(0.3), f"{REG.loc[r_, '2036']:,.0f}", 11.5, True,
       REG_COL[r_], PP_ALIGN.RIGHT)

# ============================================================ 10. WAFFLE ==
s = slide("100 squares, one market: who owns each point of 2036 demand",
          "1 square = 1% of 2036 global value, grouped by material",
          "COMBINES: waffle (pictogram) + share list + CAGR chips  |  REPLACES: material pie / doughnut")
shares = (MAT["2036"] / MAT["2036"].sum() * 100)
shares = shares.sort_values(ascending=False)
cnt = shares.apply(math.floor)
rem = (shares - cnt).sort_values(ascending=False)
for k in rem.index[: int(100 - cnt.sum())]:
    cnt[k] += 1
cells = []
for i, (m, n) in enumerate(cnt.items()):
    cells += [i] * int(n)
sq, g_ = IN(0.46), IN(0.06)
x0w, y0w = IN(0.8), IN(1.65)
for idx, mi in enumerate(cells):
    r_, c_ = divmod(idx, 10)
    rect(s, x0w + (sq + g_) * c_, y0w + (sq + g_) * r_, sq, sq, ACC[mi % len(ACC)], shape=MSO_SHAPE.ROUNDED_RECTANGLE)
yy = IN(1.75)
for i, (m, n) in enumerate(cnt.items()):
    col = ACC[i % len(ACC)]
    rect(s, IN(6.7), yy + IN(0.02), IN(0.5), IN(0.5), col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tb(s, IN(7.35), yy - IN(0.02), IN(3.6), IN(0.32), short(m, 34), 14, True)
    tb(s, IN(7.35), yy + IN(0.28), IN(3.6), IN(0.28),
       f"US$ {MAT.loc[m, '2036']:,.0f} Mn  ·  CAGR {cagr(MAT.loc[m,'2026'], MAT.loc[m,'2036'])*100:.1f}%", 11, color=SUB)
    tb(s, IN(11.2), yy, IN(1.7), IN(0.55), f"{int(n)}%", 28, True, col, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
    yy += IN(0.85)

# ======================================================== 11. TRIPTYCH ====
s = slide("Volume, value and price on one aligned timeline",
          "Global performance knitted fabrics: Mn Sq. Yards, US$ Mn and implied US$ per sq. yard, 2021-2036",
          "COMBINES: three aligned charts sharing one x-axis  |  REPLACES: separate volume and value slides")
price = [GT[y] / GV[y] for y in YEARS]
specs = [("Volume (Mn Sq. Yards)", [round(float(GV[y])) for y in YEARS], XL_CHART_TYPE.COLUMN_CLUSTERED, ACC[4], "#,##0"),
         ("Value (US$ Mn)", [round(float(GT[y])) for y in YEARS], XL_CHART_TYPE.COLUMN_CLUSTERED, ACC[0], "#,##0"),
         ("Implied price (US$ / sq. yard)", [round(p, 2) for p in price], XL_CHART_TYPE.LINE_MARKERS, ACC[2], "0.00")]
for i, (t, vals, kind, col, fmt) in enumerate(specs):
    cd = CategoryChartData()
    cd.categories = YEARS
    cd.add_series(t, vals)
    ch = chart(s, kind, cd, IN(0.4), IN(1.5) + IN(1.82) * i, IN(8.6), IN(1.8))
    style(ch)
    axes(ch, fmt=fmt)
    ch.value_axis.minimum_scale = 0 if i < 2 else min(vals) * 0.9
    title_of(ch, t, 11)
    sr = ch.plots[0].series[0]
    if kind == XL_CHART_TYPE.LINE_MARKERS:
        sr.format.line.color.rgb = rgb(col)
        sr.format.line.width = Pt(3)
        sr.smooth = False
        sr.marker.format.fill.solid()
        sr.marker.format.fill.fore_color.rgb = rgb(col)
    else:
        fill_s(sr, col)
        ch.plots[0].gap_width = 40
    if i < 2:
        ch.category_axis.tick_labels.font.size = Pt(1)
facts = [("VOLUME CAGR 26-36", f"{cagr(GV['2026'], GV['2036'])*100:.1f}%", ACC[4]),
         ("VALUE CAGR 26-36", f"{cagr(GT['2026'], GT['2036'])*100:.1f}%", ACC[0]),
         ("PRICE CHANGE 26-36", f"{(price[15]/price[5]-1)*100:+.1f}%", ACC[2])]
for i, (t, v, c) in enumerate(facts):
    y = IN(1.6) + IN(1.82) * i
    rect(s, IN(9.4), y, IN(3.45), IN(1.55), PANEL)
    rect(s, IN(9.4), y, IN(0.07), IN(1.55), c)
    tb(s, IN(9.65), y + IN(0.18), IN(3.1), IN(0.3), t, 10.5, True, SUB)
    tb(s, IN(9.65), y + IN(0.55), IN(3.1), IN(0.8), v, 36, True, c)

prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
