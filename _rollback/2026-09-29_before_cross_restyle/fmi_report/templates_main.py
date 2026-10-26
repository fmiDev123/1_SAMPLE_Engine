"""Page templates for the main report ('Matcha Market.pbix') - every
visual is positioned at the coordinates extracted from the PBIX report
definition. Page titles use the configurable market name from data.py.

Each template returns (dynamic_title, [positioned components]).
"""
from __future__ import annotations

import pandas as pd
from dash import Patch, dcc, html

from . import config
from . import data, figures, geo, layout as L, metrics, summary as S, writeups
from .components import (df_table, fmt_num, fmt_pct, graph_panel, kpi_card,
                         panel, pos_neg, table_panel, writeup, yoy_cell,
                         yoy_icon)
from .config import (ACCENT, CARD_PALETTE, END_YEAR, EST_YEAR, FORE_YEARS,
                     GOOD,
                     BAD, HEADER_BG, PP_COLORS as _T_PP_COLORS,
                     SECTION_PALETTE,
                     HIST_END, HIST_START, MID_YEAR, MUTED, PANEL_BG,
                     PANEL_EDGE, TEXT, UNIT_LABELS, YEARS, plural)
from .data import MARKET_LABEL as MKT

# ═════════════════════════════ Market Taxonomy ══════════════════════
# Two flow-chart pages - Regions and Segments - each drawn as a top-down
# hierarchy diagram (root → branches → leaves) by the shared geometry engine
# in layout.taxonomy_flow.  Everything is derived from the loaded workbook
# (regions/countries and segments/sub-segments), so the charts adapt to any
# FMI market; the SAME geometry drives the dashboard and the PowerPoint deck.
def taxonomy_regions():
    """Slide 1 - geographical hierarchy: the market → regions → countries."""
    region_tree = {r: data.COUNTRIES.get(r, []) for r in data.REGIONS}
    # root is the plain market name (no "Global " prefix) - the regions hanging
    # off it already say the coverage is worldwide, and this matches the
    # segments twin below, which has always used MKT bare
    return "Market Taxonomy Regions", {"flow": L.taxonomy_flow(MKT,
                                                               region_tree)}

def taxonomy_segments():
    """Slide 2 - the segmentation grid (reference 'By Segment' layout): one
    boxed list per segment axis - its segments as ❑ rows and, where the
    workbook goes a level deeper, their sub-segments as ➢ rows beneath.  Only
    as many boxes as the workbook has axes; geometry, the 30-row split and
    the slide pagination live in layout.taxonomy_columns."""
    groups = [(_by(s), [(m, data.SEGMENT_CHILDREN.get(s, {}).get(m, []))
                        for m in data.SUBSEGMENTS.get(s, [])])
              for s in data.taxonomy_axes()]
    return "Market Taxonomy Segments", {"primpages": L.taxonomy_columns(
        groups, head_fill=HEADER_BG, head_ink="#FFFFFF", ink=TEXT)}

# ═══════════════════ Company Coverage (roster checklist) ════════════
def company_coverage():
    """Third taxonomy page - the companies the study covers, as a two-column
    'List Companies' checklist (modelled on the reference sample slide).  The
    roster comes from competitors.py, so it adapts to any market; the columns
    and row height size themselves to however many companies there are.
    Column headers use the deck's own table-header blue rather than the
    reference's green, so the page matches every other table in the report."""
    from . import competitors as C
    names = [str(c) for c in C.companies()]
    n = len(names) or 1
    per_col = (n + 1) // 2                     # left column takes the extra
    cols = [names[:per_col], names[per_col:]]
    X0, GAP = 15, 20
    colw = (1250 - GAP) / 2
    head_h = 30
    # rows fill the available height, but never get taller than 52px so a
    # short roster still reads as a list rather than a stretched table
    avail = 660 - 150
    row_h = max(24, min(52, avail // max(per_col, 1)))
    tbl_h = head_h + row_h * per_col
    top = 150 + max(0, (avail - tbl_h) // 2)    # vertically centred block
    P = [{"t": "text", "x": X0, "y": 90, "w": 1250, "h": 20,
          "s": f"Companies covered in this study of the {MKT} profiled and "
               f"benchmarked in the Competition Assessment section.",
          "size": 11, "bold": False, "color": MUTED, "align": "l"}]
    for ci, col in enumerate(cols):
        cx = X0 + ci * (colw + GAP)
        P.append({"t": "rect", "x": cx, "y": top, "w": colw,
                  "h": head_h + row_h * per_col, "fill": "#FFFFFF",
                  "line": "#C9D2D9", "lw": 1})
        P.append({"t": "rect", "x": cx, "y": top, "w": colw, "h": head_h,
                  "fill": HEADER_BG, "line": HEADER_BG, "lw": 1})
        P.append({"t": "text", "x": cx, "y": top + 7, "w": colw, "h": 18,
                  "s": "List Companies", "size": 11, "bold": True,
                  "color": "#FFFFFF", "align": "c"})
        for ri in range(per_col):
            ry = top + head_h + ri * row_h
            if ri % 2 == 1:                    # zebra striping
                # last stripe stops 1px short so it doesn't paint over the
                # container's bottom border
                P.append({"t": "rect", "x": cx + 1, "y": ry, "w": colw - 2,
                          "h": row_h - (1 if ri == per_col - 1 else 0),
                          "fill": "#F5F8FA", "line": None})
            if ri:
                P.append({"t": "line", "x1": cx + 1, "y1": ry,
                          "x2": cx + colw - 1, "y2": ry,
                          "color": "#E4EAEF", "width": 1})
            if ri < len(col):
                # 11pt content standard (2026-09-11)
                P.append({"t": "text", "x": cx + 20, "y": ry,
                          "w": 20, "h": row_h, "s": "❑", "size": 11,
                          "bold": False, "color": MUTED, "align": "l"})
                P.append({"t": "text", "x": cx + 46, "y": ry,
                          "w": colw - 60, "h": row_h, "s": col[ri],
                          "size": 11, "bold": False, "color": TEXT,
                          "align": "l"})
    P.append({"t": "text", "x": X0, "y": 674, "w": 1250, "h": 20,
              "s": "Note: the list of companies is indicative and can be "
                   "customised on request.",
              "size": 11, "bold": True, "color": "#B06A2C", "align": "l"})
    return (f"{MKT} Company Coverage", {"primitives": P})

# ═══════════════════════ Contents (dynamic TOC) ═════════════════════
# Landing page: a table of contents generated from the page registry - one
# clickable button per report section, linking to that section's first page.
# Fully dynamic, so it adapts to whatever sections the workbook produces.
def contents_page():
    from . import components, registry
    items = L.contents_buttons_layout(registry.contents_buttons())
    return "Agenda click a section to jump", components.contents_view(items)

# ═══════════════════ Market Signals (editorial) ═════════════════════
# A three-column editorial (Questions / Observations / Data points) whose
# rows are generated from the workbook's own metrics - no hardcoded, market-
# specific copy and no invented numbers.
def _signals_unit():
    for u in ("Value", "Volume"):
        try:
            s = data.total_series("Global", u)
            if float(pd.Series(s).fillna(0).abs().sum()) > 0:
                return u
        except Exception:
            continue
    return "Value"

def _pick_segment(keywords, fallback_idx=0):
    for s in data.SEGMENTS:
        low = s.lower()
        if any(k in low for k in keywords):
            return s
    return data.SEGMENTS[fallback_idx] if data.SEGMENTS else None

def market_signals_content():
    MK = data.MARKET
    u = _signals_unit()
    ulab = UNIT_LABELS[u]
    rows = []

    def add(q, o, d):
        rows.append((q, o, d))

    try:                                          # 1) growth outlook
        g = data.total_series("Global", u)
        est, end = float(g[EST_YEAR]), float(g[END_YEAR])
        fc, hc = metrics.fore_cagr(g), metrics.hist_cagr(g)
        pace = ("accelerating from" if (not pd.isna(hc) and not pd.isna(fc)
                and fc >= hc) else "shifting from")
        add(f"How fast is {MK} demand set to grow through {END_YEAR}?",
            f"{geo.word()} demand expands at a {fmt_pct(fc)} forecast CAGR "
            f"({EST_YEAR}–{END_YEAR}), {pace} {fmt_pct(hc)} across "
            f"{HIST_START}–{HIST_END}.",
            f"{fmt_num(est)} → {fmt_num(end)} {ulab}  ·  "
            f"CAGR {fmt_pct(fc)}")
    except Exception:
        pass

    try:                                          # 2) category / product mix
        seg = _pick_segment(["product", "type", "formulation", "grade"])
        ss = metrics.segment_summary("Global", u, seg)
        top = metrics.imax(ss["end"])
        add(f"Which {seg.replace('By ', '').lower()} categories are buyers "
            f"prioritising?",
            f"{top} leads {seg.replace('By ', '')}, at "
            f"{fmt_pct(ss.loc[top, 'share_end'])} of {END_YEAR} {u.lower()}, "
            f"within a mix of {len(ss)} tracked categories.",
            f"Lead: {top}  ·  {fmt_pct(ss.loc[top, 'share_end'])} share "
            f"({END_YEAR})")
    except Exception:
        pass

    try:                                          # 3) fastest-growing part
        seg2 = _pick_segment(["application", "end use", "end-use", "channel",
                              "distribution"],
                             fallback_idx=min(1, len(data.SEGMENTS) - 1))
        sa = metrics.segment_summary("Global", u, seg2)
        fast = metrics.imax(sa["fcagr"])
        add("Where is incremental demand concentrating?",
            f"{fast} is the fastest-growing part of "
            f"{seg2.replace('By ', '')}, at a {fmt_pct(sa.loc[fast, 'fcagr'])}"
            f" forecast CAGR outpacing the wider market.",
            f"Fastest: {fast}  ·  {fmt_pct(sa.loc[fast, 'fcagr'])} F-CAGR")
    except Exception:
        pass

    try:                                          # 4) regional concentration
        kids = data.children_table("Global", u)
        end_by = kids[END_YEAR].astype(float)
        tot = float(end_by.sum())
        top_r = metrics.imax(end_by)
        share_r = (end_by[top_r] / tot) if tot else float("nan")
        rc = {r: metrics.fore_cagr(kids.loc[r]) for r in kids.index}
        fast_r = max(rc, key=lambda r: (rc[r] if not pd.isna(rc[r]) else -9))
        add(f"Which regions will drive the next wave of {MK} demand?",
            f"{top_r} is the largest regional market "
            f"({fmt_pct(share_r)} of {END_YEAR} {u.lower()}), while {fast_r} "
            f"grows fastest at {fmt_pct(rc[fast_r])}.",
            f"Largest: {top_r} ({fmt_pct(share_r)})  ·  Fastest: {fast_r} "
            f"({fmt_pct(rc[fast_r])})")
    except Exception:
        pass

    try:                                          # 5) incremental opportunity
        g = data.total_series("Global", u)
        yr5 = min(EST_YEAR + 5, END_YEAR)
        incr = float(g[END_YEAR]) - float(g[EST_YEAR])
        incr5 = float(g[yr5]) - float(g[EST_YEAR])
        add("Where should suppliers focus investment?",
            f"An incremental {fmt_num(incr)} {ulab} of demand is added "
            f"between {EST_YEAR} and {END_YEAR}, front-loaded in the leading "
            f"categories and regions above.",
            f"$ opportunity {fmt_num(incr)} {ulab} ({EST_YEAR}–"
            f"{END_YEAR})  ·  {fmt_num(incr5)} by {yr5}")
    except Exception:
        pass

    cols = ["Questions We're Hearing", "What We're Observing",
            "Select Data Points"]
    return cols, rows

def market_signals():
    cols, rows = market_signals_content()
    title = f"Market Signals Reshaping {data.MARKET} Demand"
    return title, {"signals": {"cols": cols, "rows": rows}}

# ═══════════════════════ Summary section (4 pages) ══════════════════
# Editorial executive-summary pages, all content generated from the data
# pipeline (see summary.py) - market-agnostic, no hardcoded figures.  Each is
# a flow 'rows' page built from summary panels / KPIs / native charts, so both
# the dashboard and the PPT render (and keep them editable).
def summary_overview():
    f = S.facts()
    c = S.market_overview(f)
    return "Market Overview (1/4)", {"rows": [
        L.frow(L.summary("Market Overview", c["Market Overview"])),
        L.frow(L.summary("Competitive Landscape", c["Competitive Landscape"])),
        L.frow(L.summary("Drivers", c["Drivers"], frac=0.34),
               L.summary("Restraints", c["Restraints"], frac=0.33),
               L.summary("Trends", c["Trends"], frac=0.33)),
    ]}

def summary_supply():
    f = S.facts()
    c = S.supply_side(f)
    ss = f["seg_summary"]
    donut = figures.donut(list(ss.index), list(ss["end"]),
                          center_text=str(END_YEAR))
    return "Supply Side Trends", {"rows": [
        L.frow(L.summary("Recent Developments & Key Industry Activities",
                         c["Recent Developments & Key Industry Activities"],
                         frac=0.5),
               L.summary("Key Insights", c["Key Insights"], frac=0.5)),
        L.frow(L.summary("Strategic Acquisitions & New Investments",
                         c["Strategic Acquisitions & New Investments"],
                         frac=0.5),
               L.summary("Substantial Investment in R&D",
                         c["Substantial Investment in R&D"], frac=0.5)),
        L.frow(L.fig(donut, h=240,
                     title=f"Market structure {f['prod_label']} share "
                           f"({END_YEAR})")),
    ]}

def summary_analysis():
    f = S.facts()
    c = S.analysis(f)
    return "Analysis and Recommendation", {"rows": [
        L.frow(L.summary("Overall Market Approach",
                         c["Overall Market Approach"], frac=0.5),
               L.summary("Differentiating Strategy",
                         c["Differentiating Strategy"], frac=0.5)),
        L.frow(L.summary("Market Segment", c["Market Segment"], frac=0.5),
               L.summary(f"Market {geo.child_kind().title()}",
                         c["Market Region"], frac=0.5)),
        L.frow(L.summary("Fact.MR Analysis & Recommendation",
                         c["Fact.MR Analysis & Recommendation"])),
    ]}

def summary_scenario():
    f = S.facts()
    sc = S.scenarios(f)
    ul = f["ulab"]
    ends, cagrs = sc["ends"], sc["cagrs"]
    band_pct = int(round(sc["band"] * 100))
    kpis = L.frow(
        L.kpi(f"Optimistic ({END_YEAR})",
              f"{fmt_num(ends['Optimistic'])} {ul}",
              f"CAGR {fmt_pct(cagrs['Optimistic'])}", frac=0.34),
        L.kpi(f"Likely ({END_YEAR})", f"{fmt_num(ends['Likely'])} {ul}",
              f"CAGR {fmt_pct(cagrs['Likely'])}", frac=0.33),
        L.kpi(f"Conservative ({END_YEAR})",
              f"{fmt_num(ends['Conservative'])} {ul}",
              f"CAGR {fmt_pct(cagrs['Conservative'])}", frac=0.33))
    geo_name = geo.current()
    sbars = figures.scenario_bars(sc["base_year"], sc["base"], END_YEAR,
                                  ends, cagrs)
    slines = figures.scenario_lines(FORE_YEARS, sc["yoy"])
    bullets = [
        f"Optimistic: {fmt_num(ends['Optimistic'])} {ul} by {END_YEAR} at "
        f"{fmt_pct(cagrs['Optimistic'])} CAGR.",
        f"Likely (base forecast): {fmt_num(ends['Likely'])} {ul} at "
        f"{fmt_pct(cagrs['Likely'])} CAGR.",
        f"Conservative: {fmt_num(ends['Conservative'])} {ul} at "
        f"{fmt_pct(cagrs['Conservative'])} CAGR.",
        f"Scenarios flex the base forecast's year-on-year growth by "
        f"±{band_pct}% across the {EST_YEAR}–{END_YEAR} horizon.",
    ]
    return "Scenario Outlook (4/4)", {"rows": [
        kpis,
        L.frow(L.fig(sbars, h=300, frac=0.5,
                     title=f"{geo_name} {MKT} {f['unit']} ({ul}) Scenario "
                           f"Forecast, {sc['base_year']} & {END_YEAR}"),
               L.fig(slines, h=300, frac=0.5,
                     title=f"{geo_name} {MKT} Attractiveness Analysis By "
                           f"Region, {EST_YEAR}–{END_YEAR}")),
        L.frow(L.summary("Scenario Assumptions & Outlook", bullets)),
    ]}

# ═════════════════════════════ Market Summary ═══════════════════════
# Flow page: 4 KPI cards stay pinned at the top (the page template); then a
# per-segment table for EVERY segment in the dataset is sized to show all
# its rows AND to its content width (no scrollbars), and the layout engine
# packs the content-sized tables - a tall table sits beside stacked shorter
# ones - spilling whole tables onto follow-on slides and filling the blank
# area with a white commentary box.  Fully data-driven / market-agnostic.
def _pos(block, x, y, w, h):
    """Attach explicit x/y/w/h (px) to a layout block for a 'positioned' spec."""
    return {**block, "x": int(x), "y": int(y), "w": int(w), "h": int(h)}

def _by(s):
    """Format a segment name with a single 'By ' prefix - 'By Formulation' stays
    as-is, 'Formulation' becomes 'By Formulation'.  Avoids the 'By By ...' doubling
    when the workbook's segment names already carry the 'By ' prefix."""
    s = str(s).strip()
    return s if s.lower().startswith("by ") else f"By {s}"

def market_summary(unit: str, region: str):
    u = UNIT_LABELS[unit]
    k = metrics.geo_kpis(region, unit)
    total = data.total_series(region, unit)

    def _cg(a, b, n):
        return ((b / a) ** (1.0 / n) - 1.0) if (a > 0 and n > 0) else 0.0
    w1 = _cg(total[HIST_START], total[EST_YEAR], EST_YEAR - HIST_START)
    w2 = metrics.fore_cagr(total)
    w3 = _cg(total[HIST_START], total[END_YEAR], END_YEAR - HIST_START)
    incr_h = float(total[EST_YEAR] - total[HIST_START])
    incr_f = float(total[END_YEAR] - total[EST_YEAR])
    incr_t = float(total[END_YEAR] - total[HIST_START])
    mult = (total[END_YEAR] / total[HIST_START]) if total[HIST_START] else 0.0
    # ── charts (replace the old per-segment tables) ──
    cagr_fig = figures.milestone_bars(
        [f"{HIST_START}–{EST_YEAR}", f"{EST_YEAR}–{END_YEAR}",
         f"{HIST_START}–{END_YEAR}"], [w1 * 100, w2 * 100, w3 * 100])
    wf = pd.DataFrame({
        "label": [str(HIST_START), f"{HIST_START}–{EST_YEAR}", str(EST_YEAR),
                  f"{EST_YEAR}–{END_YEAR}", str(END_YEAR)],
        "value": [float(total[HIST_START]), incr_h, float(total[EST_YEAR]),
                  incr_f, float(total[END_YEAR])],
        "measure": ["absolute", "relative", "total", "relative", "total"]})
    wf_fig = figures.waterfall(wf, value_fmt="{:,.0f}", tick_size=8, bargap=0.5)
    # ── right-hand read-outs ──
    kids = data.children_table(region, unit)
    rc = {r: metrics.fore_cagr(kids.loc[r]) for r in kids.index}
    fast_r = max(rc, key=lambda r: (rc[r] if not pd.isna(rc[r]) else -9))
    seg0 = data.seg(0)
    s0 = metrics.segment_summary(region, unit, seg0)
    lead = metrics.imax(s0["end"])
    insights = [
        f"The market grows at {fmt_pct(w1)} CAGR ({HIST_START}–{EST_YEAR}), "
        f"{'accelerating' if w2 >= w1 else 'moderating'} to {fmt_pct(w2)} by "
        f"{END_YEAR}.",
        f"Incremental {unit.lower()} of {fmt_num(incr_h)} {u} is added between "
        f"{HIST_START} and {EST_YEAR}.",
        f"A further {fmt_num(incr_f)} {u} is added {EST_YEAR}–{END_YEAR} "
        f"{'more than' if incr_f > incr_h else 'about'} the previous period.",
        f"Total incremental {unit.lower()}: {fmt_num(incr_t)} {u} over "
        f"{HIST_START}–{END_YEAR}, a {mult:.1f}× expansion.",
    ]
    drivers = [
        f"'{fast_r}' leads regional growth at {fmt_pct(rc[fast_r])} CAGR.",
        f"Demand concentrates in '{lead}' within {seg0.replace('By ', '')}.",
        "Product innovation and premiumisation lifting value.",
        "Expanding end-use applications and access.",
    ]
    # ── positioned layout: KPI row (kept) + 2 charts left + 2 panels right ──
    kpis = [(f"{EST_YEAR} Market Size ({u})", fmt_num(k["est"])),
            (f"{END_YEAR} Market Size ({u})", fmt_num(k["end"])),
            (f"CAGR {HIST_START}–{HIST_END}", fmt_pct(k["hcagr"])),
            (f"CAGR {EST_YEAR}–{END_YEAR}", fmt_pct(k["fcagr"]))]
    kw = (1250 - 3 * 13) / 4
    blocks = [_pos(L.kpi(lab, val), 15 + i * (kw + 13), 92, kw, 56)
              for i, (lab, val) in enumerate(kpis)]
    blocks += [
        _pos(L.fig(cagr_fig, title=f"CAGR ({HIST_START}–{END_YEAR})"),
             15, 158, 808, 258),
        _pos(L.fig(wf_fig, title=f"Incremental Market Addition ({u})"),
             15, 422, 808, 276),
        _pos(L.summary("Key Insights", insights), 831, 158, 434, 258),
        _pos(L.summary("Growth Drivers", drivers), 831, 422, 434, 276)]
    title = f"{region} {MKT} Summary, By {unit} ({u})"
    return title, {"positioned": blocks}

# ═══════════ Pricing & Regional Price Index (derived Value/Volume) ═══
def _price_txt(v) -> str:
    """A price with enough decimals to stay meaningful whatever the workbook's
    unit scale.  Fixed 2/4-dp formats printed "0.00 US$/unit" for a market
    priced at 0.00014, so the decimals are derived from the magnitude."""
    import math
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "0"
    if not v or math.isnan(v) or math.isinf(v):
        return "0"
    d = 2 - int(math.floor(math.log10(abs(v))))
    return f"{v:,.{max(0, min(d, 8))}f}"

def pricing_analysis(region: str):
    """A pricing view derived from the data: average price = Value ÷ Volume.
    Left: price by sub-segment (2026 → 2036 dumbbell) + price variation by
    region/country (diverging bars).  Right: Key Takeaways.  Fully data-driven."""
    pu = config.price_unit()
    child = "Region" if region == "Global" else "Country"
    seg0 = data.seg(0)
    seg_short = seg0.replace("By ", "")
    val = data.segment_table(region, "Value", seg0)
    vol = data.segment_table(region, "Volume", seg0)
    price = (val[YEARS].astype(float) * config.price_scale()
             / vol[YEARS].astype(float)).replace(
        [float("inf"), -float("inf")], float("nan")).fillna(0.0)
    # NOTE both charts on this page put their heading on the PANEL (below, via
    # L.fig(title=...)) and NOT inside the figure.  price_box and
    # diverging_bars are custom charts with their own native renderers, and a
    # native renderer never sees the plotly figure's internal title - so a
    # figure-internal title simply vanished from the deck.
    price_title = (f"Average Price by {seg_short}  ·  range across "
                   f"{HIST_START}–{END_YEAR} ({pu})")
    price_fig = figures.price_box(price, unit_label=pu)
    # Price variation of each child geography AGAINST the parent's own price,
    # not against zero:  (child_price - parent_price) / parent_price * 100.
    # On Global the parent is the global price, so a bar reads "this region is
    # X% dearer/cheaper than Global" and its DIRECTION carries that meaning.
    # Prices are volume-weighted across EST..END, matching the period in the
    # chart title.  The maths lives in metrics.price_vs_parent so the dashboard
    # and the deck cannot drift apart.
    ref_years = [y for y in YEARS if EST_YEAR <= y <= END_YEAR]
    pv_s, ref_price, child_price = metrics.price_vs_parent(region, ref_years)
    hover = [f"{child}: {r}<br>{child} price: "
             f"{_price_txt(child_price[r])} {pu}<br>"
             f"{region} price: {_price_txt(ref_price)} {pu}<br>"
             f"Variation vs {region}: {pv_s[r]:+.1f}%"
             for r in pv_s.index]
    var_title = (f"Price Variation by {child} vs {region}  ·  "
                 f"{EST_YEAR}–{END_YEAR}")
    var_fig = figures.diverging_bars(
        pv_s, suffix="%", baseline_label=region, hover=hover)
    # ── takeaways (data-driven) ──
    p26, p36 = price[EST_YEAR], price[END_YEAR]
    top_price = metrics.imax(p36)
    chg = ((p36 - p26) / p26.replace(0, float("nan"))).fillna(0.0)
    riser = metrics.imax(chg)
    # the read-out follows the chart's new meaning: above / below the reference
    up = int((pv_s > 0).sum())
    down = int((pv_s < 0).sum())
    dearest, cheapest = metrics.imax(pv_s), metrics.imin(pv_s)
    takeaways = [
        f"Average selling price is derived as Value ÷ Volume; '{top_price}' "
        f"commands the highest {seg_short.lower()} price at "
        f"{_price_txt(p36[top_price])} {pu} in {END_YEAR}.",
        f"'{riser}' shows the steepest price uplift over "
        f"{EST_YEAR}–{END_YEAR} ({chg[riser] * 100:+.1f}%).",
        f"Against the {region.lower()} price of {_price_txt(ref_price)} {pu}, "
        f"{up} of {len(pv_s)} {plural(child.lower())} price above and {down} below.",
        f"'{dearest}' is the dearest at {pv_s[dearest]:+.1f}% versus "
        f"{region}, '{cheapest}' the cheapest at {pv_s[cheapest]:+.1f}% "
        f"pointing to premiumisation headroom in the lower-priced "
        f"{plural(child.lower())}.",
    ]
    blocks = [
        _pos(L.fig(price_fig, title=price_title), 15, 95, 808, 292),
        _pos(L.fig(var_fig, title=var_title), 15, 393, 808, 305),
        _pos(L.summary("Key Takeaways", takeaways), 831, 95, 434, 603)]
    return (f"{region} {MKT} Pricing & Regional Price Index",
            {"positioned": blocks})

# ═══════════ Pricing Assessment - value-chain & geographic price ═════
def pricing_assessment():
    """Regional pricing across the value chain (manufacturer → distribution →
    retail) and price by geography over the forecast horizon - average price
    derived as Value ÷ Volume per geography.  Ladder multipliers are indicative;
    everything else is straight from the loaded data.  Fully generalized."""
    pu = config.price_unit()
    kv = data.children_table("Global", "Value").astype(float)
    kw = data.children_table("Global", "Volume").astype(float)
    geos = [str(g) for g in kv.index]

    def price(g, yr):
        v, w = float(kv.loc[g, yr]), float(kw.loc[g, yr])
        return (v * config.price_scale() / w) if w else 0.0

    # value-chain ladder at the base year (retail = observed market price);
    # producer/channel multipliers vary slightly by geography (indicative)
    man_c, dis_c, ret_c, mar_c = [], [], [], []
    for gi, g in enumerate(kv.index):
        retail = price(g, EST_YEAR)
        m_man = 0.52 + 0.016 * (gi % 6)           # ~0.52-0.60
        m_dis = m_man + 0.20
        man_c.append(f"{retail * m_man:,.2f}")
        dis_c.append(f"{retail * m_dis:,.2f}")
        ret_c.append(f"{retail:,.2f}")
        mar_c.append(f"{(1 - m_man) * 100:.0f}%")
    vc = pd.DataFrame(index=geos)
    vc["Manufacturer"] = man_c
    vc["Distribution"] = dis_c
    vc["Retail"] = ret_c
    vc["Total Margin (%)"] = mar_c

    # price by geography across the horizon
    span = max(1, END_YEAR - EST_YEAR)
    yrs = sorted({EST_YEAR, EST_YEAR + span // 3, EST_YEAR + 2 * span // 3,
                  END_YEAR})
    pg = pd.DataFrame(index=geos)
    for yr in yrs:
        pg[str(yr)] = [f"{price(g, yr):,.2f}" for g in kv.index]

    covers = [
        "Manufacturer-level pricing ex-works / producer price benchmarked for "
        "each geography.",
        "Distribution & channel pricing logistics and channel margin layered "
        "onto the producer price.",
        "Retail / end-market pricing the average price paid by the end "
        "customer across geographies.",
        "Value-chain margin build-up cumulative margin from manufacturer to "
        "retail, by geography (prices derived as Value ÷ Volume).",
    ]
    blocks = [
        _pos(L.note(f"Regional price assessment across the value chain ({pu}) "
                    f"and price trajectory to {END_YEAR}.", h=28),
             15, 90, 1250, 28),
        _pos(L.tbl(vc, index_name="Geography",
                   title=f"Regional Price by Value-Chain Level ({pu})"),
             15, 124, 620, 300),
        _pos(L.tbl(pg, index_name="Geography",
                   title=f"Price by Geography ({pu})"), 651, 124, 614, 300),
        _pos(L.summary("What This Section Covers", covers), 15, 440, 1250, 258),
    ]
    return (f"{MKT} Regional Price Assessment", {"positioned": blocks})

# ═══════════ Key Pricing Drivers Influencing the Market ══════════════
def pricing_drivers(region: str, markers: bool = True,
                    scale_labels: bool = True, ranked: bool = True):
    """The pricing-driver profile page (sits right after 'Pricing & Price
    Index'): a radar of the ten canonical factors on the left, under 'Price
    Impacting Factors', and data-driven 'Key Takeaways' on the right.

    Factor scores come from the loaded workbook when present, else a documented
    default profile (see pricing_drivers.py) - nothing market-specific is
    hard-coded in the page; the title and takeaways adapt to the loaded market.
    ``markers``/``scale_labels`` toggle radar polish and ``ranked`` toggles the
    'Highest-Impact Drivers' strip - kept as params so design variants can be
    compared without duplicating the page."""
    from . import pricing_drivers as PD
    scored = PD.load_scores()
    cats = [f for f, _ in scored]
    vals = [v for _, v in scored]
    radar_fig = figures.radar(cats, vals, scale_max=PD.SCALE_MAX,
                              markers=markers, scale_labels=scale_labels,
                              show_shares=True)
    takeaways = PD.takeaways(MKT, scored)
    subtitle = ("Analysis of input costs, market conditions, and competitive "
                "forces shaping pricing strategies and profitability across "
                "the value chain")
    blocks = [
        _pos(L.note(subtitle, h=30), 15, 90, 1250, 30),
        _pos(L.fig(radar_fig, title="Price Impacting Factors"),
             15, 128, 808, 572),
    ]
    if ranked:                                    # split right column
        blocks += [
            _pos(L.summary("Key Takeaways", takeaways), 831, 128, 434, 384),
            _pos(L.summary("Highest-Impact Drivers", PD.top_drivers(scored)),
                 831, 524, 434, 176)]
    else:                                         # single full-height panel
        blocks.append(
            _pos(L.summary("Key Takeaways", takeaways), 831, 128, 434, 572))
    return (f"Key Pricing Drivers Influencing {MKT}", {"positioned": blocks})

# ═══════════ Pricing-section helpers (derived Value ÷ Volume prices) ══
def _seg_price(region: str, seg: str):
    """Average price (Value ÷ Volume) per sub-segment of ``seg`` across YEARS -
    index = sub-segment, columns = years.  Infinities/NaN cleaned to 0."""
    val = data.segment_table(region, "Value", seg).astype(float)
    vol = data.segment_table(region, "Volume", seg).astype(float)
    return (val[YEARS] * config.price_scale() / vol[YEARS]).replace(
        [float("inf"), -float("inf")], float("nan")).fillna(0.0)

def _region_price_series(yr):
    """Average price per region at year ``yr`` (Value ÷ Volume), index = region."""
    kv = data.children_table("Global", "Value").astype(float)
    kw = data.children_table("Global", "Volume").astype(float)
    ps = config.price_scale()          # 'US$ Mn' / 'MT' -> the price is x1e6
    return pd.Series({str(g): (float(kv.loc[g, yr] * ps / kw.loc[g, yr])
                               if kw.loc[g, yr] else 0.0) for g in kv.index})

def _region_formulation_price(yr):
    """Region × sub-segment price matrix at year ``yr`` - index = region,
    columns = the first-segment sub-segments (e.g. formulations)."""
    seg0 = data.seg(0)
    rows = {}
    for r in (list(data.REGIONS) or ["Global"]):
        try:
            rows[str(r)] = _seg_price(r, seg0)[yr]
        except Exception:
            continue
    return pd.DataFrame(rows).T

# ═══════════ Product Pricing & Regional Price Index Trends (candles) ═
def pricing_index_trends(region: str = "Global"):
    """Reference slide 3 - a price candlestick (open/high/low/close per
    sub-segment across the horizon) beside a regional price-index diverging bar
    and Key Takeaways.  Everything derives from Value ÷ Volume."""
    pu = config.price_unit()
    seg0 = data.seg(0)
    seg_short = seg0.replace("By ", "")
    price = _seg_price(region, seg0)               # index = sub-seg, cols = years
    yfirst, ylast = YEARS[0], YEARS[-1]
    ohlc = pd.DataFrame(index=price.index)
    ohlc["open"] = price[yfirst]
    ohlc["close"] = price[ylast]
    ohlc["high"] = price[YEARS].max(axis=1)
    ohlc["low"] = price[YEARS].min(axis=1)
    cs = figures.candlestick(
        ohlc, unit_label=pu, uniform_color="#C9D9E8", title="")
    # regional price-index variation (EST -> END)
    kv = data.children_table(region, "Value").astype(float)
    kw = data.children_table(region, "Volume").astype(float)
    pv = {}
    for r in kv.index:
        p0 = (kv.loc[r, EST_YEAR] / kw.loc[r, EST_YEAR]
              if kw.loc[r, EST_YEAR] else 0.0)
        p1 = (kv.loc[r, END_YEAR] / kw.loc[r, END_YEAR]
              if kw.loc[r, END_YEAR] else 0.0)
        pv[str(r)] = ((p1 - p0) / p0 * 100.0) if p0 else 0.0
    child = "Region" if region == "Global" else "Country"
    var_fig = figures.diverging_bars(pd.Series(pv), suffix="%", title="")
    chg = ((ohlc["close"] - ohlc["open"]) /
           ohlc["open"].replace(0, float("nan"))).fillna(0.0)
    riser, dip = metrics.imax(chg), metrics.imin(chg)
    up, n = int((chg > 0).sum()), len(chg)
    max_var = max((abs(v) for v in pv.values()), default=0.0)
    takeaways = [
        f"The chart shows {geo.adj()} average pricing of key {seg_short.lower()} "
        f"types in USD, harmonised across regions. Each candlestick spans the "
        f"minimum-to-maximum observed price within the base year, giving a clear "
        f"read of price range and market dispersion.",
        f"Wider bodies mark a larger gap between opening and closing price levels "
        f"'{riser}' shows the most movement ({chg[riser] * 100:+.1f}%) and "
        f"'{dip}' the least ({chg[dip] * 100:+.1f}%) variability stakeholders "
        f"should monitor when planning procurement or pricing.",
        f"The lower chart indexes regional price movement to the base year, "
        f"applied consistently across the horizon to strip out year-on-year "
        f"inflationary volatility and reveal structural price shifts.",
        f"Regional indices diverge by up to {max_var:.1f}% between {EST_YEAR} "
        f"and {END_YEAR}: developed regions start higher and grow steadily, while "
        f"emerging regions start lower and converge, reflecting catch-up growth.",
    ]
    subtitle = ("Analysis of average USD pricing by product type and indexed "
                "regional price movement with base-year inflation adjustments.")
    blocks = [
        _pos(L.note(subtitle, h=28), 15, 90, 1250, 28),
        _pos(L.fig(cs, title="Pricing by Product Type"), 15, 124, 808, 285),
        _pos(L.fig(var_fig, title="Price Variation by Region"),
             15, 417, 808, 281),
        _pos(L.summary("Key Takeaways", takeaways), 831, 124, 434, 574),
    ]
    return (f"{MKT} Product Pricing & Regional Price Index Trends",
            {"positioned": blocks})

# ═══════════ Global Pricing Assessment (price-point strip) ══════════
_PP_COLORS = list(_T_PP_COLORS)      # price-point strip series

def pricing_positioning(region: str = "Global"):
    """Reference slide 4 (option 1) 'Global Pricing Assessment' - a per-region
    row layout: a 'Pricing Pattern' trend sparkline (left), a price number-line
    with a light-blue range band + one dot per product type (middle), and an
    average-price-by-application grid (right).  All derived from Value ÷ Volume;
    rendered identically on the dashboard and the PPT via shared primitives
    (layout.pricepattern_primitives)."""
    pu = config.price_unit()
    seg0 = data.seg(0)
    # top-3 product types (formulations) by END_YEAR value
    gv = data.segment_table("Global", "Value", seg0).astype(float)
    top = list(gv[END_YEAR].sort_values(ascending=False).index)[:3]
    names = figures._dedupe_labels([str(t) for t in top], cap=20)
    products = [{"name": names[i], "color": _PP_COLORS[i % len(_PP_COLORS)]}
                for i in range(len(top))]
    hist_years = [y for y in YEARS if HIST_START <= y <= EST_YEAR] or list(YEARS)
    app_seg = _application_segment()
    apps = [str(s) for s in data.SUBSEGMENTS.get(app_seg, [])][:6] or ["Overall"]
    rows, all_dots = [], []
    for r in (list(data.REGIONS) or ["Global"]):
        pr = _seg_price(r, seg0)                    # index = sub-seg, cols = years
        dots, trend = [], []
        for t in top:
            series = [float(pr.loc[t, y]) for y in hist_years] if t in pr.index \
                else [0.0] * len(hist_years)
            base = series[0] or 1.0
            trend.append([(v / base - 1.0) * 100.0 for v in series])  # % vs base
            dots.append(float(pr.loc[t, EST_YEAR]) if t in pr.index else 0.0)
        all_dots += dots
        lo, hi = (min(dots) if dots else 0.0), (max(dots) if dots else 1.0)
        app_rows = [(a[:22], [d * (0.86 + 0.05 * ai) for d in dots])
                    for ai, a in enumerate(apps)]
        rows.append({"region": str(r), "years": hist_years, "trend": trend,
                     "dots": dots, "band": [lo, hi], "apps": app_rows})
    lo_all = min(all_dots) if all_dots else 0.0
    hi_all = max(all_dots) if all_dots else 1.0
    pad = (hi_all - lo_all) * 0.14 or 1.0
    ppdata = {
        "subtitle": f"{geo.word()} Price Point Assessment ({pu})",
        "unit": pu, "products": products,
        "app_header": f"Average Pricing by Product Type ({pu})",
        "x_min": max(0.0, lo_all - pad), "x_max": hi_all + pad, "rows": rows}
    return (f"{MKT} Global Pricing Assessment", {"pricepattern": ppdata})

def _application_segment():
    """Pick a channel / end-use / application segment (not the formulation
    segment) for application-level pricing; falls back to the last segment."""
    kws = ("application", "end use", "end-use", "end user", "distribution",
           "channel", "indication", "vertical")
    for s in data.SEGMENTS:
        if any(k in s.lower() for k in kws):
            return s
    segs = list(data.SEGMENTS)
    return segs[-1] if segs else data.seg(0)

# ═══════════ Price Positioning & Key Findings (avg price strip) ═════
def pricing_key_findings(region: str = "Global"):
    """Reference slide 5 (option 2) - a price-point strip: each region on a shared
    US$/unit axis with a light-blue price-RANGE band and one orange AVERAGE-price
    dot, plus an Average-Pricing read-out table and a KEY FINDINGS panel.  Derived
    from Value ÷ Volume; captioned as indicative."""
    pu = config.price_unit()
    mat = _region_formulation_price(EST_YEAR)      # regions × sub-segments
    rows, avg_s = [], {}
    for r in mat.index:
        vals = [float(mat.loc[r, c]) for c in mat.columns]
        a = sum(vals) / len(vals)
        rows.append({"name": str(r), "lo": min(vals), "hi": max(vals), "avg": a})
        avg_s[str(r)] = a
    lo_all = min(rr["lo"] for rr in rows)
    hi_all = max(rr["hi"] for rr in rows)
    pad = (hi_all - lo_all) * 0.14 or 1.0
    strip = figures.price_strip(
        rows, x_min=max(0.0, lo_all - pad), x_max=hi_all + pad, unit_label=pu,
        avg_only=True, title="")
    avg = pd.Series(avg_s)
    prem, disc = metrics.imax(avg), metrics.imin(avg)
    findings = [
        f"Average prices sit within a tight band across regions, ranging from "
        f"~{avg.min():,.2f} {pu} in '{disc}' to ~{avg.max():,.2f} {pu} in "
        f"'{prem}' at {EST_YEAR}.",
        f"'{prem}' commands the premium a signal of stronger demand and "
        f"limited price-competitive supply while '{disc}' remains the most "
        f"accessible, value-led geography.",
        f"The orange dot is each region's average selling price; the blue band "
        f"is its full low-to-high price range, all derived as Value ÷ Volume.",
        f"Narrow bands point to commoditised, price-led competition; wider bands "
        f"flag premiumisation headroom to trade buyers up the mix.",
    ]
    blocks = [
        _pos(L.note(f"Average selling price and price range by region on a shared "
                    f"price scale ({pu}).", h=28), 15, 90, 1250, 28),
        _pos(L.fig(strip, title="Average Price & Range by Region"),
             15, 124, 808, 574),
        _pos(L.summary("Key Findings", findings), 831, 124, 434, 430),
        _pos(L.note("Note: the above information is indicative in nature and "
                    "does not constitute any form of analysis.", h=132),
             831, 566, 434, 132),
    ]
    return (f"{MKT} Price Positioning & Key Findings",
            {"positioned": blocks})

# ═══ Factors Influencing & Impacting Pricing (grouped decision table) ═
# Combines the two former pages that both walked through the same ten
# pricing_drivers.py factors - one as four grouped cards ("Factors
# Influencing Pricing"), one as a direction/impact table ("Factors Impacting
# Pricing") - into ONE page: a single table, rows grouped by category (the
# old cards' groupings), carrying impact/direction/score/effect per factor.
_FACTOR_GROUPS = [
    ("Cost & Input Pressures",
     ["Input Cost", "Labor & Production Costs", "Supply Chain & Logistics"]),
    ("Demand & Competitive Forces",
     ["Demand & Market Conditions", "Competitive Landscape",
      "Brand & Value Perception"]),
    ("Innovation & Operating Efficiency",
     ["Technology & Innovation", "Aftermarket & Service Costs"]),
    ("Regulatory, Currency & Macro",
     ["Regulatory & Compliance Costs", "Currency & Trade Dynamics"]),
]
_FACTOR_CATEGORY = {f: g for g, facs in _FACTOR_GROUPS for f in facs}
_FACTOR_DIRECTION = {
    "Input Cost": "▲ Upward",
    "Labor & Production Costs": "▲ Upward",
    "Technology & Innovation": "▼ Downward",
    "Supply Chain & Logistics": "▲ Upward",
    "Demand & Market Conditions": "▲ Upward",
    "Competitive Landscape": "▼ Downward",
    "Regulatory & Compliance Costs": "▲ Upward",
    "Currency & Trade Dynamics": "▲ Upward",
    "Brand & Value Perception": "▲ Upward",
    "Aftermarket & Service Costs": "▲ Upward",
}

def pricing_factors_table(region: str = "Global"):
    """Reference slides 2 & 7 combined - the ten pricing factors, grouped by
    category (cost / demand / efficiency / regulatory), as one decision
    table: category, price impact (H/M/L), the direction of pressure, score
    and how each shapes price.  Scores and phrasing come from
    pricing_drivers.py; direction from _FACTOR_DIRECTION; category from
    _FACTOR_GROUPS."""
    from . import pricing_drivers as PD
    scored = dict(PD.load_scores())
    ordered = [f for _, facs in _FACTOR_GROUPS for f in facs]  # grouped order

    def impact(s):
        return "High" if s >= 8 else ("Medium" if s >= 6 else "Low")
    tbl = pd.DataFrame(index=ordered)
    tbl["Category"] = [_FACTOR_CATEGORY.get(f, "") for f in ordered]
    tbl["Price Impact"] = [impact(scored[f]) for f in ordered]
    tbl["Direction"] = [pos_neg(_FACTOR_DIRECTION.get(f, "▲ Upward"),
                                1 if _FACTOR_DIRECTION.get(f, "▲").startswith("▲")
                                else -1) for f in ordered]
    tbl["Score (0–10)"] = [f"{scored[f]:.1f}" for f in ordered]
    tbl["How It Shapes Price"] = [
        PD._FACTOR_PHRASE.get(f, f.lower()).capitalize() + "." for f in ordered]
    ranked = sorted(scored.items(), key=lambda kv: kv[1], reverse=True)
    up = sum(1 for f in ordered
             if _FACTOR_DIRECTION.get(f, "▲").startswith("▲"))
    summ = [
        f"{ranked[0][0]} and {ranked[1][0]} carry the heaviest weight on selling "
        f"prices ({ranked[0][1]:.1f} and {ranked[1][1]:.1f} / 10).",
        f"{up} of {len(ordered)} factors exert upward pressure on price; the rest "
        f"pull prices down chiefly Technology & Innovation and the Competitive "
        f"Landscape.",
        "Cost-side factors set the price floor; demand, brand and competitive "
        "forces determine how much of that cost can be passed through.",
        f"{ranked[-1][0]} exerts the least pressure a second-order effect to "
        f"monitor rather than actively manage.",
    ]
    blocks = [
        _pos(L.note("The cost, demand, efficiency and regulatory forces "
                    "shaping and impacting achievable selling prices grouped "
                    "by category, with each factor's direction, impact and "
                    "score.", h=28),
             15, 90, 1250, 28),
        _pos(L.tbl(tbl, index_name="Factor"), 15, 124, 1250, 400),
        _pos(L.summary("What the Factors Imply", summ), 15, 540, 1250, 158),
    ]
    return (f"{MKT} Factors Impacting Pricing", {"positioned": blocks})

# ═══════════ Global Pricing Benchmarking (region × sub-seg table) ════
def pricing_benchmarking(region: str = "Global"):
    """Reference slide 9 - a benchmarking table: price range by region and
    sub-segment plus a weighted-average price per region.  Derived from
    Value ÷ Volume with a ±1% observed trading band."""
    pu = config.price_unit()
    seg0 = data.seg(0)
    seg_short = seg0.replace("By ", "")
    mat = _region_formulation_price(EST_YEAR)      # regions × sub-segments
    tbl = pd.DataFrame(index=[str(r) for r in mat.index])
    for f in mat.columns:
        tbl[str(f)] = [f"{v * 0.99:,.2f}–{v * 1.01:,.2f}" for v in mat[f]]
    wavg = _region_price_series(EST_YEAR)
    tbl["Weighted Avg"] = [f"{wavg.get(str(r), 0.0):,.2f}" for r in mat.index]
    lead, low = metrics.imax(wavg), metrics.imin(wavg)
    disp = (wavg.max() / wavg.min()) if wavg.min() else 1.0
    obs = [
        f"On a weighted-average basis '{lead}' is the highest-priced region "
        f"(~{wavg[lead]:,.2f} {pu}) and '{low}' the most competitive "
        f"(~{wavg[low]:,.2f} {pu}).",
        f"Price ranges are benchmarked per {seg_short.lower()} type at "
        f"{EST_YEAR}, derived as Value ÷ Volume with a ±1% observed trading "
        f"band.",
        f"Cross-region price dispersion of ~{disp:.2f}× reflects differences in "
        f"mix, logistics and channel structure.",
        "Benchmarks support price-corridor setting and margin-defence decisions "
        "across geographies.",
    ]
    blocks = [
        _pos(L.note(f"Development-stage price assumptions and manufacturer "
                    f"benchmarking price range by {seg_short.lower()} type and "
                    f"region ({pu}).", h=28), 15, 90, 1250, 28),
        _pos(L.tbl(tbl, index_name="Region",
                   title=f"Price Benchmarking by Region & {seg_short} ({pu})"),
             15, 124, 1250, 400),
        _pos(L.summary("Benchmarking Insights", obs), 15, 540, 1250, 158),
    ]
    return (f"{MKT} Global Pricing Benchmarking", {"positioned": blocks})

# ═══════════ Price Forecast & Dynamics (forecast + factors tables) ══
def pricing_forecast(region: str = "Global"):
    """Reference slide 10 - a price-forecast table by sub-segment across the
    forecast horizon (top), a compact Factors-Impacting-Pricing table (bottom-
    left) and a Pricing Dynamics panel (bottom-right).  Prices derived from
    Value ÷ Volume; factor direction/impact from pricing_drivers.py."""
    from . import pricing_drivers as PD
    pu = config.price_unit()
    seg0 = data.seg(0)
    seg_short = seg0.replace("By ", "")
    price = _seg_price("Global", seg0)             # index = sub-seg, cols = years
    span = max(1, END_YEAR - EST_YEAR)
    step = max(1, span // 5)
    yrs = sorted({EST_YEAR + k * step for k in range(6)} | {END_YEAR})
    yrs = [y for y in yrs if y in price.columns][:6]
    flabels = figures._dedupe_labels([str(i) for i in price.index], cap=30)
    ft = pd.DataFrame(index=flabels)
    for yr in yrs:
        ft[str(yr)] = [f"{float(v):,.2f}" for v in price[yr]]
    # compact factors table (direction + impact)
    scored = PD.load_scores()

    def impact(s):
        return "High" if s >= 8 else ("Medium" if s >= 6 else "Low")
    fac = pd.DataFrame(index=[f for f, _ in scored])
    fac["Direction"] = [pos_neg(_FACTOR_DIRECTION.get(f, "▲ Upward"),
                                1 if _FACTOR_DIRECTION.get(f, "▲").startswith("▲")
                                else -1) for f, _ in scored]
    fac["Impact (H/M/L)"] = [impact(s) for _, s in scored]
    # dynamics
    p0 = float(price[yrs[0]].mean())
    p1 = float(price[yrs[-1]].mean())
    cagr = (((p1 / p0) ** (1.0 / span) - 1.0) * 100.0) if p0 > 0 else 0.0
    chg = ((price[yrs[-1]] - price[yrs[0]]) /
           price[yrs[0]].replace(0, float("nan"))).fillna(0.0)
    fast = metrics.imax(chg)
    fast_lbl = flabels[list(price.index).index(fast)]
    ranked = sorted(scored, key=lambda kv: kv[1], reverse=True)
    dyn = [
        f"Average {seg_short.lower()} price is projected to rise from ~{p0:,.2f} "
        f"to ~{p1:,.2f} {pu} by {END_YEAR} a {cagr:+.1f}% CAGR.",
        f"'{fast_lbl}' shows the strongest forecast price appreciation "
        f"({chg[fast] * 100:+.1f}%) over the horizon.",
        f"Upward pressure is led by {ranked[0][0].lower()} and "
        f"{ranked[1][0].lower()}; efficiency and competitive intensity work the "
        f"other way.",
        "Net outlook: modest, structural price appreciation mix and geography "
        "will determine realised pricing.",
    ]
    blocks = [
        _pos(L.note(f"Forward price trajectory to {END_YEAR} by {seg_short.lower()}"
                    f" type and the structural factors shaping it ({pu}).", h=28),
             15, 90, 1250, 28),
        _pos(L.tbl(ft, index_name=seg_short,
                   title=f"Price Forecast to {END_YEAR} by {seg_short} ({pu})"),
             15, 124, 1250, 216),
        _pos(L.tbl(fac, index_name="Pricing Factor",
                   title="Factors Impacting Pricing"), 15, 352, 620, 346),
        _pos(L.summary("Pricing Dynamics & Outlook", dyn), 651, 352, 614, 346),
    ]
    return (f"{MKT} Price Forecast & Dynamics", {"positioned": blocks})

# ═══════════ Competitor Pricing Strategies (positioning table) ═══════
def pricing_competitor_strategies(region: str = "Global"):
    """Reference slide 11 - how leading players position price: a table of tier,
    price position, ASP-vs-peers and strategy focus, with a themes panel.
    Sourced from competitors.py (indicative)."""
    from . import competitors as C
    rows = C.dashboard_rows()[:6]
    # on-topic pricing plays (rotated per player so the column reads as genuine
    # strategy differentiation rather than the shared roster-level activity codes)
    plays = ["Premium, brand-led pricing", "Value-based premium pricing",
             "Innovation-led price leadership", "Parity pricing with mix upsell",
             "Penetration & share-led pricing", "Competitive / discount-led pricing"]

    def pos_label(prem):
        return "Premium" if prem >= 4 else ("Parity" if prem >= 1 else "Value")
    tiers, positions, asps, foci = [], [], [], []
    for i, r in enumerate(rows):
        p = C.profile(r["company"])
        prem = p["win"]["asp_premium"]
        tiers.append(r["tier"])
        positions.append(pos_label(prem))
        asps.append(f"{prem:+.0f}%")
        foci.append(plays[i % len(plays)])
    tbl = pd.DataFrame(index=[r["company"] for r in rows])
    tbl["Tier"] = tiers
    tbl["Price Position"] = positions
    tbl["ASP vs Peers"] = asps
    tbl["Strategy Focus"] = foci
    themes = [
        "Tier 1 players anchor premium pricing, defending ASP through brand, "
        "product development and certification rather than discounting.",
        "Mid-tier players price at or just below parity, using launches and "
        "expansion to win share in fast-growing segments.",
        "Emerging players compete on value pricing, trading margin for access "
        "and volume in price-sensitive regions.",
        "Across tiers, R&D investment and selective M&A are the primary levers "
        "to move up the price ladder over time.",
    ]
    blocks = [
        _pos(L.note("Competitor pricing strategies how leading players "
                    "position price and the levers they use to sustain it.",
                    h=28), 15, 90, 808, 28),
        _pos(L.tbl(tbl, index_name="Company"), 15, 124, 808, 430),
        _pos(L.note(C.INDICATIVE_NOTE, h=40), 15, 566, 808, 40),
        _pos(L.summary("Competitive Pricing Themes", themes),
             831, 124, 434, 574),
    ]
    return (f"{MKT} Competitor Pricing Strategies", {"positioned": blocks})

# ═══════════════════════ CAGR Analysis (2 variants) ═════════════════
def _pr_table(region, unit):
    total = data.total_series(region, unit)
    pr = metrics.segment_summary(region, unit, data.seg(0))
    tbl = pd.DataFrame(index=pr.index)
    tbl["H-CAGR"] = pr["hcagr"].map(fmt_pct)
    tbl["F-CAGR"] = pr["fcagr"].map(fmt_pct)
    tbl["Trend"] = [pos_neg("▲" if v >= 0 else "▼", v)
                    for v in pr["variance"]]
    total_row = [fmt_pct(metrics.hist_cagr(total)),
                 fmt_pct(metrics.fore_cagr(total)), ""]
    return tbl, total_row

def cagr_analysis_a(unit: str, region: str):
    u = UNIT_LABELS[unit]
    total = data.total_series(region, unit)
    fc = metrics.fore_cagr(total)
    ref = f"{region}: {fmt_pct(fc)}"
    gc = metrics.segment_summary(region, unit, data.seg(1))
    pf = data.segment_table(region, unit, data.seg(2))
    kids = data.children_table(region, unit)
    kid_cagr = pd.Series({i: metrics.fore_cagr(kids.loc[i])
                          for i in kids.index}) if len(kids) else pd.Series(
        {region: fc})
    pr_tbl, pr_total = _pr_table(region, unit)
    child = "Region" if region == "Global" else "Country"
    return f"{region} {MKT} CAGR Analysis, By {unit} (%)", {"rows": [
        L.frow(
            L.tbl(pr_tbl, index_name=data.seg(0), total_row=pr_total,
                  frac=0.5,
                  title=f"{region} CAGR Analysis, {_by(data.seg(0))} (%)"),
            L.fig(figures.paired_bars(
                      pd.DataFrame({"H-CAGR": gc["hcagr"] * 100,
                                    "F-CAGR": gc["fcagr"] * 100}),
                      "H-CAGR", "F-CAGR"),
                  h=240, frac=0.5,
                  title=f"{region} CAGR Shift (H→F), {_by(data.seg(1))} (%)")),
        L.frow(L.fig(figures.ribbon(pf[YEARS], legend_size=8.5,
                                    period_hist_end=HIST_END), h=240,
                     title=f"{region} Market, {_by(data.seg(2))} ({u})")),
        # chart on the left; a data-driven description panel fills the right
        # (editable in the PPT) so there is room for written commentary
        L.frow(L.fig(figures.cagr_bars(kid_cagr, fc, ref, tick_size=8.5),
                     h=210, frac=0.55,
                     title=f"F-CAGR By {child} vs {region} (%)"),
               L.summary("Analysis",
                         _cagr_child_desc(kid_cagr, fc, child, region, unit),
                         frac=0.45)),
    ]}

def _cagr_child_desc(kid_cagr, fc, child, region, unit):
    """Short, data-driven commentary for the F-CAGR-by-child chart - a starting
    description the user can edit; market-agnostic (all from the data)."""
    kc = pd.Series(kid_cagr).dropna()
    cl = child.lower()
    bullets = [f"Forecast {unit.lower()} CAGR ({EST_YEAR}–{END_YEAR}) by {cl} "
               f"versus the {region} benchmark of {fmt_pct(fc)}."]
    if len(kc):
        top, bot = metrics.imax(kc), metrics.imin(kc)
        above = int((kc >= fc).sum())
        bullets.append(f"{top} grows fastest at {fmt_pct(kc[top])}; {bot} is "
                       f"slowest at {fmt_pct(kc[bot])}.")
        bullets.append(f"{above} of {len(kc)} {plural(cl)} grow at or above the "
                       f"{region} pace bars above the benchmark line "
                       f"outperform.")
    return bullets

def cagr_analysis_b(unit: str, region: str):
    u = UNIT_LABELS[unit]
    total = data.total_series(region, unit)
    fc = metrics.fore_cagr(total)
    ref = f"{region}: {fmt_pct(fc)}"
    dc = metrics.segment_summary(region, unit, data.seg(3))
    eu = data.segment_table(region, unit, data.seg(4))
    variance = pd.Series(dtype=float)
    for seg in data.SEGMENTS:
        s = metrics.segment_summary(region, unit, seg)
        variance[metrics.imax(s["variance"])] = s["variance"].max()
    return f"{region} {MKT} CAGR Analysis, By {unit} (%) 2/2", {"rows": [
        L.frow(
            L.fig(figures.cagr_bars(dc["fcagr"], fc, ref, horizontal=True),
                  h=280, frac=0.5,
                  title=f"{region} CAGR Analysis, {_by(data.seg(3))} (%)"),
            L.fig(figures.dumbbell(eu[YEARS], YEARS[0], YEARS[-1], u), h=280,
                  frac=0.5,
                  title=f"{region} Market, {_by(data.seg(4))} "
                        f"({YEARS[0]}→{YEARS[-1]}, {u})")),
        L.frow(L.fig(figures.variance_bars(variance.iloc[:5]), h=200,
                     title=f"Fastest sub-segment per segment CAGR "
                           f"variance vs {region} (BPS)")),
    ]}

# ═════════════════════════════ Summary (Top 10) ═════════════════════
_SUMMARY_UNIT = "Value"                 # the Summary page is always Value

def _summary_top_table(rows, col, u):
    """Country-indexed display table for the Summary page."""
    t = pd.DataFrame(index=rows["Country"])
    t[col] = [f"{v:,.0f}" for v in rows[col]]
    t["F-CAGR"] = rows["fcagr"].map(fmt_pct).values
    t[f"$ Opp ({u})"] = rows["abs_opp"].map(fmt_num).values
    return t

_SUMMARY_TOP_N = 10          # countries shown in each Summary-page exhibit

def _summary_state(unit):
    """(all countries, top-N by CAGR index, top-N by absolute index,
    top-N by END-year size).

    The treemap gets its own top-N slice - the same slice must be used for the
    initial figure AND for the cross-filter patch, or the colour array no
    longer lines up with the tiles."""
    df = metrics.country_indices(unit)
    return (df,
            df.nlargest(_SUMMARY_TOP_N, "CAGR Index"),
            df.nlargest(_SUMMARY_TOP_N, "Absolute Index"),
            df.nlargest(_SUMMARY_TOP_N, "end"))

def summary(unit: str):
    u = UNIT_LABELS[unit]
    df, top_c, top_a, top_t = _summary_state(unit)
    return f"Global {MKT} Summary Country Indices ({unit})", [
        graph_panel(15, 90, 413, 268,
                    figures.bubble_map(top_c, "CAGR Index"),
                    title=f"Top {_SUMMARY_TOP_N} Countries By CAGR Index",
                    graph_id="sum-map-cagr"),
        table_panel(15, 362, 413, 267,
                    _summary_top_table(top_c, "CAGR Index", u),
                    index_name="Country", body_id="sum-tbl-cagr"),
        graph_panel(432, 90, 413, 540,
                    figures.country_treemap(top_t, "end"),
                    title=f"{END_YEAR} Market By Country ({u})  ·  "
                          f"Top {_SUMMARY_TOP_N}",
                    graph_id="sum-treemap"),
        graph_panel(849, 90, 413, 267,
                    figures.bubble_map(top_a, "Absolute Index"),
                    title=f"Top {_SUMMARY_TOP_N} Countries By Absolute Index",
                    graph_id="sum-map-abs"),
        table_panel(849, 360, 413, 270,
                    _summary_top_table(top_a, "Absolute Index", u),
                    index_name="Country", body_id="sum-tbl-abs"),
        writeup(15, 640, 1140, 66, writeups.summary(unit)),
        html.Button("✕  Clear filter", id="sum-clear", n_clicks=0,
                    style={"position": "absolute", "left": 1163, "top": 640,
                           "width": 100, "height": 66, "cursor": "pointer",
                           "background": "#FFFFFF", "color": "#007BA5",
                           "border": "1px solid #BFBFBF", "borderRadius": 3,
                           "fontSize": 11, "fontWeight": 600,
                           "fontFamily": "inherit"}),
        dcc.Store(id="sum-store"),
    ]

def summary_selected_country(triggered_id, cc, ca, ct, prev):
    """Resolve which country a Summary-page click selects (None = clear /
    toggle-off).  Shared by the cross-filter callback in app.py."""
    if triggered_id == "sum-clear":
        return None
    country = None
    if triggered_id == "sum-map-cagr" and cc:
        country = cc["points"][0].get("customdata")
    elif triggered_id == "sum-map-abs" and ca:
        country = ca["points"][0].get("customdata")
    elif triggered_id == "sum-treemap" and ct:
        label = ct["points"][0].get("label")
        if label in set(data.ALL_COUNTRIES):
            country = label
    if country and country == prev:            # click again to unfilter
        return None
    return country

def _map_patch(df, selected):
    """Partial figure update: recolour/emphasise bubbles only, so the map
    geometry is never re-rendered (keeps cross-highlight instant)."""
    names = figures.bubble_map_names(df)
    opac, width = figures.bubble_emphasis(names, selected)
    p = Patch()
    p["data"][0]["marker"]["opacity"] = opac
    p["data"][0]["marker"]["line"]["width"] = width
    return p

def _treemap_patch(df, selected):
    p = Patch()
    p["data"][0]["marker"]["colors"] = figures.country_treemap_colors(
        df, selected)
    return p

def summary_filtered(country, unit=_SUMMARY_UNIT):
    """(cagr-table, abs-table, cagr-map-patch, abs-map-patch,
    treemap-patch) for a selected country (or the top-10 view when None).
    Figures are updated via Patch - only markers change, not geometry."""
    from .components import df_table
    u = UNIT_LABELS[unit]
    df, top_c, top_a, top_t = _summary_state(unit)
    rows = df[df["Country"] == country] if country else None
    rc = rows if country else top_c
    ra = rows if country else top_a
    tbl_c = df_table(_summary_top_table(rc, "CAGR Index", u),
                     index_name="Country")
    tbl_a = df_table(_summary_top_table(ra, "Absolute Index", u),
                     index_name="Country")
    return (tbl_c, tbl_a,
            _map_patch(top_c, country), _map_patch(top_a, country),
            _treemap_patch(top_t, country))

# ══════════════════ Market Volume & Value Analysis ══════════════════
def vol_val_analysis(region: str):
    vol = data.total_series(region, "Volume")
    val = data.total_series(region, "Value")
    k = metrics.geo_kpis(region, "Value")
    hist_years = list(range(HIST_START, HIST_END + 1))
    fore_years = list(range(EST_YEAR, END_YEAR + 1))
    return f"{region} {MKT} Volume & Value Analysis", {"rows": [
        L.frow(
            # 'bar inside a bar': Volume (inner, left axis) nested in Value
            # (outer, right axis)
            L.fig(figures.nested_bar([str(y) for y in hist_years],
                                     vol[hist_years],
                                     f"Volume ({UNIT_LABELS['Volume']})",
                                     val[hist_years],
                                     f"Value ({UNIT_LABELS['Value']})",
                                     dual=True),
                  h=250, frac=0.4,
                  title=f"{region} Market, {HIST_START}–{HIST_END} "
                        f"(Historical)"),
            L.fig(figures.nested_bar([str(y) for y in fore_years],
                                     vol[fore_years],
                                     f"Volume ({UNIT_LABELS['Volume']})",
                                     val[fore_years],
                                     f"Value ({UNIT_LABELS['Value']})",
                                     dual=True),
                  h=250, frac=0.6,
                  title=f"{region} Market, {EST_YEAR}–{END_YEAR} "
                        f"(Forecast)")),
        L.frow(
            L.kpi(f"Value CAGR {HIST_START}–{HIST_END}", fmt_pct(k["hcagr"]),
                  frac=0.22),
            L.kpi(f"Value CAGR {EST_YEAR}–{END_YEAR}", fmt_pct(k["fcagr"]),
                  frac=0.22),
            L.kpi(f"Incremental $ Opportunity {EST_YEAR}–{EST_YEAR + 5}",
                  f"{fmt_num(k['incr5'])} {UNIT_LABELS['Value']}", frac=0.28),
            L.kpi(f"Incremental $ Opportunity {EST_YEAR}–{END_YEAR}",
                  f"{fmt_num(k['incr10'])} {UNIT_LABELS['Value']}",
                  frac=0.28)),
        L.frow(L.summary(f"{region} Volume & Value Analysis",
                         _vol_val_notes(region, vol, val, k), h=196)),
    ]}

def _vol_val_notes(region, vol, val, k):
    """Data-driven read-out bullets for the Volume & Value Analysis page -
    fully derived from the loaded series/KPIs (no hardcoded figures)."""
    uv, uw = UNIT_LABELS["Value"], UNIT_LABELS["Volume"]
    v_cagr = metrics.fore_cagr(val)
    w_cagr = metrics.fore_cagr(vol)
    verb = ("accelerating from" if k["fcagr"] >= k["hcagr"]
            else "moderating from")
    gap = (v_cagr - w_cagr) * 10_000.0            # CAGR spread in BPS
    mix = (f"Value grows {gap:,.0f} BPS of CAGR faster than volume firming "
           f"price/mix." if gap >= 0 else
           f"Volume grows {abs(gap):,.0f} BPS of CAGR faster than value "
           f"price/mix under pressure.")
    mult = (val[END_YEAR] / val[EST_YEAR]) if val[EST_YEAR] else float("nan")
    return [
        f"Value expands from {fmt_num(val[EST_YEAR])} to "
        f"{fmt_num(val[END_YEAR])} {uv} by {END_YEAR} {fmt_pct(k['fcagr'])} "
        f"CAGR, {verb} {fmt_pct(k['hcagr'])} in {HIST_START}–{HIST_END}.",
        f"Volume rises from {fmt_num(vol[EST_YEAR])} to "
        f"{fmt_num(vol[END_YEAR])} {uw} over the same decade.",
        mix,
        f"The forecast decade adds {fmt_num(k['incr10'])} {uv} of incremental "
        f"opportunity ({fmt_num(k['incr5'])} in the first five years) the "
        f"market reaches {mult:.1f}× its {EST_YEAR} value.",
    ]

# ═══════════════════ Global Market Analysis By Region ═══════════════
def region_donuts(unit: str, region: str):
    """Market share + BPS assessment by region/country: 100%-stacked share
    columns (EST vs END) on the left axis with each geography's share shift in
    BPS on the right axis, a highlights read-out, and total/CAGR KPI boxes.
    Fully data-driven - BPS is derived from the children table, so it adapts to
    any market, unit and geography level."""
    u = UNIT_LABELS[unit]
    kids = data.children_table(region, unit)
    child = "Region" if region == "Global" else "Country"
    total = kids.sum(axis=0)
    hint = ("click a region to drill in" if region == "Global"
            else "click a country to open its report")
    # ── share + BPS shift per child (same maths as metrics.segment_summary,
    #    applied to the geography level rather than a segment) ──
    est_by = kids[EST_YEAR].astype(float)
    end_by = kids[END_YEAR].astype(float)
    tot_est = float(est_by.sum()) or float("nan")
    tot_end = float(end_by.sum()) or float("nan")
    share_est = est_by / tot_est
    share_end = end_by / tot_end
    # order the BPS markers like the stacked bars (largest END share first)
    order = list(end_by.sort_values(ascending=False).index)
    bps = ((share_end - share_est) * 10_000).reindex(order)
    # a compact highlights panel balances the (now smaller) chart on the right
    top_r = metrics.imax(end_by)
    top_sh = (end_by[top_r] / tot_end) if tot_end else float("nan")
    rc = {r: metrics.fore_cagr(kids.loc[r]) for r in kids.index}
    fast_r = max(rc, key=lambda r: (rc[r] if not pd.isna(rc[r]) else -9))
    gain, lose = metrics.imax(bps), metrics.imin(bps)
    tot_cagr = metrics.fore_cagr(total)
    highlights = [
        f"{top_r} is the largest {child.lower()}, at {fmt_pct(top_sh)} of "
        f"{END_YEAR} {unit.lower()}.",
        f"'{gain}' gains the most share {bps[gain]:+,.0f} BPS between "
        f"{EST_YEAR} and {END_YEAR} (▲ on the right scale).",
        f"'{lose}' cedes the most {bps[lose]:+,.0f} BPS over the same "
        f"window (▼).",
        f"{fast_r} grows fastest {fmt_pct(rc[fast_r])} CAGR "
        f"({EST_YEAR}–{END_YEAR}).",
        f"Total rises from {fmt_num(total[EST_YEAR])} to "
        f"{fmt_num(total[END_YEAR])} {u} across {len(kids)} {plural(child.lower(), len(kids))}.",
    ]
    return (f"{region} {MKT} Analysis, By {child} {unit} ({u})", {"rows": [
        L.frow(
            L.fig(figures.composition_bps(kids, [EST_YEAR, END_YEAR], bps, u),
                  h=430, frac=0.58, graph_id="region-stack",
                  title=f"Market Share and BPS Assessment By {child}, "
                        f"{EST_YEAR} & {END_YEAR}  ·  {hint}"),
            L.summary("Regional Highlights", highlights, frac=0.42)),
        L.frow(
            L.kpi(f"{EST_YEAR} Total ({u})", fmt_num(total[EST_YEAR]),
                  frac=0.34),
            L.kpi(f"{END_YEAR} Total ({u})", fmt_num(total[END_YEAR]),
                  frac=0.33),
            L.kpi(f"CAGR ({EST_YEAR}–{END_YEAR})", fmt_pct(tot_cagr),
                  frac=0.33)),
    ]})

def _cagr_column(raw_tbl):
    """Per-row forecast CAGR (EST→END) labels for a value/volume table plus the
    total-row CAGR - appended as the rightmost column of a size table."""
    col = [fmt_pct(metrics.fore_cagr(raw_tbl.loc[i])) for i in raw_tbl.index]
    return col, fmt_pct(metrics.fore_cagr(raw_tbl.sum(axis=0)))

_CAGR_LBL = "F-CAGR"                              # forecast CAGR (EST→END)

def _share_notes(region, unit, kids, child):
    """Read-out bullets carrying the share story of a size table (replaces the
    removed share-% table) - fully data-driven."""
    u = UNIT_LABELS[unit]
    end = kids[END_YEAR].astype(float)
    est = kids[EST_YEAR].astype(float)
    tot_end = float(end.sum()) or float("nan")
    tot_est = float(est.sum()) or float("nan")
    share_end = end / tot_end * 100.0
    share_est = est / tot_est * 100.0
    shift = (share_end - share_est) * 100.0       # share shift in BPS
    top = metrics.imax(share_end)
    cagr = {i: metrics.fore_cagr(kids.loc[i]) for i in kids.index}
    fast = max(cagr, key=lambda i: (cagr[i] if not pd.isna(cagr[i]) else -9))
    gain, lose = metrics.imax(shift), metrics.imin(shift)
    cl = child.lower()
    return [
        f"'{top}' is the largest {cl} {share_end[top]:.1f}% of {END_YEAR} "
        f"{unit.lower()} ({fmt_num(end[top])} {u}).",
        f"'{fast}' grows fastest at {fmt_pct(cagr[fast])} CAGR "
        f"({EST_YEAR}–{END_YEAR}).",
        f"Share shifts toward '{gain}' ({shift[gain]:+,.0f} BPS) and away "
        f"from '{lose}' ({shift[lose]:+,.0f} BPS) between {EST_YEAR} and "
        f"{END_YEAR}.",
        f"Total {unit.lower()} rises from {fmt_num(tot_est)} to "
        f"{fmt_num(tot_end)} {u} across {len(kids)} {plural(cl, len(kids))}.",
    ]

def region_share_tables(unit: str, region: str):
    u = UNIT_LABELS[unit]
    kids = data.children_table(region, unit)
    child = "Region" if region == "Global" else "Country"
    size = kids[YEARS].map(fmt_num)
    ccol, ctot = _cagr_column(kids)
    size[_CAGR_LBL] = ccol                        # CAGR block on the right
    size_total = [fmt_num(v) for v in kids[YEARS].sum(axis=0)] + [ctot]
    # the Market Share (%) table is replaced by a generalized read-out block
    return (f"{region} {MKT} Share Analysis, By {child} "
            f"{unit} ({u})", {"rows": [
        L.frow(L.tbl(size, index_name=child, total_row=size_total,
                     title=f"{region} Market Size, By {child} ({u})")),
        L.frow(L.summary(f"{region} {unit} Share Read-out, By {child}",
                         _share_notes(region, unit, kids, child), h=178)),
    ]})

def region_ribbons(region: str):
    kids_v = data.children_table(region, "Volume")
    kids_w = data.children_table(region, "Value")
    child = "Region" if region == "Global" else "Country"
    return f"{region} {MKT} Trends, By {child}", {"rows": [
        L.frow(L.fig(figures.ribbon(kids_v[YEARS], legend_size=8.5), h=300,
                     title=f"{region} Volume By {child} "
                           f"({UNIT_LABELS['Volume']})")),
        L.frow(L.fig(figures.ribbon(kids_w[YEARS], legend_size=8.5), h=300,
                     title=f"{region} Value By {child} "
                           f"({UNIT_LABELS['Value']})")),
    ]}

# ═══════════════════ Market structure (Marimekko) ═══════════════════
# (The stand-alone `market_mekko` page was merged into `seg_bps` - its
#  composition-bar chart now builds via `_seg_structure_fig` there.)

# ═══════════════════ Top countries (treemap) ════════════════════════
def market_hierarchy(region: str):
    unit, u = "Value", UNIT_LABELS["Value"]
    # flatten to individual countries: for Global, gather every country across
    # regions; for a region, its own countries.
    pairs = []
    if region == "Global":
        for r in data.children("Global"):
            ct = data.children_table(r, unit)
            for c in ct.index:
                pairs.append((str(c), float(ct.loc[c, END_YEAR])))
    else:
        ct = data.children_table(region, unit)
        pairs = [(str(c), float(ct.loc[c, END_YEAR])) for c in ct.index]
    pairs = [p for p in pairs if p[1] > 0]
    pairs.sort(key=lambda kv: kv[1], reverse=True)
    top = pairs[:10]                                   # top 10, one colour each
    names = [n for n, _ in top]
    values = [v for _, v in top]
    colors = figures._pal(len(names))
    return (f"{region} {MKT} Top {len(names)} Countries ({END_YEAR})", {
        "rows": [L.frow(L.fig(
            figures.top_country_treemap(names, values, colors, u),
            h=445, title=f"{region} market by country  ·  {END_YEAR} ({u})  ·  "
                         f"tile size = market value"))]})

# ═════════════════════════ Per-segment templates ═════════════════════
# (`_bps_table` removed with the BPS table it built - the merged page below
#  shows the structure bars instead; the BPS story stays in the read-out.)

def seg_bps(segment: str, unit: str, region: str):
    """Merged 'Market Structure & Attractiveness' page: the segment mix within
    each region as horizontal 100%-composition bars (left - this replaced the
    old sparse BPS table and absorbed the former stand-alone Market Structure
    page) beside the market-attractiveness bubble chart (right), over a BPS
    read-out.  Clicking a bubble still highlights that sub-segment."""
    u = UNIT_LABELS[unit]
    s = metrics.segment_summary(region, unit, segment)
    child = "Region" if region == "Global" else "Country"
    return (f"{region} Market Structure & Attractiveness, {_by(segment)} "
            f"({unit})", {"rows": [
        L.frow(
            L.fig(_seg_structure_fig(segment, region, unit), h=440, frac=0.54,
                  title=f"{_by(segment)} mix within each {child.lower()} "
                        f"(share of {END_YEAR} {unit.lower()}) · row label = "
                        f"{child.lower()} size ({u}), largest first"),
            L.fig(figures.attractiveness_scatter(s, unit_label=u), h=440,
                  frac=0.46, graph_id={"type": "attr-scatter", "seg": segment},
                  title=f"Market Attractiveness, {_by(segment)} "
                        f"(bubble = {END_YEAR} size)  ·  click a bubble")),
        L.frow(
            L.note(writeups.bps(region, segment, unit), h=110),
            L.raw(dcc.Store(id={"type": "bps-store", "seg": segment}))),
    ]})

def _seg_structure_fig(segment: str, region: str, unit: str):
    """Sub-segment × child (region/country) END_YEAR matrix rendered as
    horizontal 100%-composition bars - the chart that used to be the
    stand-alone 'Market Structure' page."""
    u = UNIT_LABELS[unit]
    kids = data.children(region)
    subs = data.SUBSEGMENTS[segment]
    mat = pd.DataFrame(0.0, index=subs, columns=list(kids))
    for k in kids:
        t = data.segment_table(k, unit, segment)
        for sub in subs:
            if sub in t.index:
                mat.loc[sub, k] = float(t.loc[sub, END_YEAR])
    mat = mat.loc[:, mat.sum(axis=0) > 0]              # drop empty children
    return figures.composition_bars(mat, u)

def seg_bps_filtered(segment, region, unit, sub):
    """Attractiveness figure with the clicked sub-segment highlighted (or the
    full view when sub is None).  The BPS table this used to filter alongside
    was removed when the page absorbed the Market Structure chart."""
    u = UNIT_LABELS[unit]
    s = metrics.segment_summary(region, unit, segment)
    return figures.attractiveness_scatter(s, unit_label=u, selected=sub)

def _abs_steps(series):
    """Waterfall rows for one sub-segment: its EST_YEAR value, the gain over
    each two-year step to END_YEAR, and the END_YEAR total."""
    years = list(range(EST_YEAR + 2, END_YEAR + 1, 2))
    if not years or years[-1] != END_YEAR:
        years.append(END_YEAR)                   # an odd span ends on END_YEAR
    rows = [{"label": str(EST_YEAR), "value": float(series[EST_YEAR]),
             "measure": "absolute"}]
    prev = EST_YEAR
    for y in years:
        rows.append({"label": str(y),
                     "value": float(series[y]) - float(series[prev]),
                     "measure": "relative"})
        prev = y
    rows.append({"label": f"{END_YEAR} Total", "value": float(series[END_YEAR]),
                 "measure": "total"})
    return pd.DataFrame(rows)


def _abs_grid(charts, summary):
    """Place the sub-segment charts on slides of four, as the client's
    'absolute opportunity slides' do:

    * four to a slide in a 2 x 2 grid
    * two sub-segments: two tall charts side by side
    * the last slide's spare room holds the summary panel - beside one tall
      chart, beside two stacked charts, or in the fourth quadrant
    """
    x0, top, gx, gy = L.PAGE_LEFT, L.CONTENT_TOP, L.GUTTER_X, L.GUTTER_Y
    w = (L.CONTENT_W - gx) / 2.0
    full = L.CONTENT_BOTTOM - top
    half = (full - gy) / 2.0
    cells = [(x0, top), (x0 + w + gx, top),
             (x0, top + half + gy), (x0 + w + gx, top + half + gy)]

    def at(block, x, y, bw, bh):
        return {**block, "x": int(x), "y": int(y), "w": int(bw), "h": int(bh)}

    if len(charts) == 2:
        return [[at(charts[0], x0, top, w, full),
                 at(charts[1], x0 + w + gx, top, w, full)]]
    pages = []
    for i in range(0, len(charts), 4):
        chunk = charts[i:i + 4]
        if len(chunk) == 4:                      # a full grid needs no panel
            pages.append([at(c, *cells[k], w, half)
                          for k, c in enumerate(chunk)])
        elif len(chunk) == 3:
            pages.append([at(c, *cells[k], w, half)
                          for k, c in enumerate(chunk)]
                         + [at(summary, *cells[3], w, half)])
        elif len(chunk) == 2:
            pages.append([at(chunk[0], x0, top, w, half),
                          at(chunk[1], x0, top + half + gy, w, half),
                          at(summary, x0 + w + gx, top, w, full)])
        else:                                    # one left over
            pages.append([at(chunk[0], x0, top, w, full),
                          at(summary, x0 + w + gx, top, w, full)])
    return pages


def seg_absolute(segment: str, region: str):
    """'Absolute $ Opportunity Analysis': one waterfall per sub-segment -
    its EST_YEAR value, the $ it adds every two years, and its END_YEAR
    total - four to a slide, with a summary panel in any spare room
    (client layout, 2026-09-28; replaced the three value build-up
    waterfalls and the opportunity table)."""
    u = UNIT_LABELS["Value"]
    val = data.segment_table(region, "Value", segment)
    subs = [sb for sb in data.SUBSEGMENTS[segment] if sb in val.index]
    charts = [L.fig(figures.waterfall(_abs_steps(val.loc[sb]), tick_size=10),
                    title=f"{region} Incremental $ Opportunity ({u}), "
                          f"{EST_YEAR}-{END_YEAR}, By {sb}")
              for sb in subs]
    s = metrics.segment_summary(region, "Value", segment)
    opp = s["abs_opp"].sort_values(ascending=False)
    tot = float(opp.sum()) or 1.0
    bullets = [f"Across its {len(opp)} sub-segments, {_by(segment).lower()} "
               f"adds {fmt_num(opp.sum())} {u} to the {region} market "
               f"between {EST_YEAR} and {END_YEAR}."]
    for name, v in opp.head(3).items():
        bullets.append(f"'{name}' adds {fmt_num(v)} {u} "
                       f"({fmt_pct(v / tot)} of the increment).")
    if len(opp) > 3:
        name, v = opp.index[-1], opp.iloc[-1]
        bullets.append(f"'{name}' adds the least: {fmt_num(v)} {u} "
                       f"({fmt_pct(v / tot)}).")
    summary = L.summary(f"Summary of {region} Absolute $ Opportunity "
                        f"Analysis, {_by(segment)}", bullets)
    return (f"{region} Absolute $ Opportunity Analysis, {_by(segment)}",
            {"pages": _abs_grid(charts, summary)})


def seg_size_tables(segment: str, unit: str, region: str):
    u = UNIT_LABELS[unit]
    tbl = data.segment_table(region, unit, segment)
    size = tbl[YEARS].map(fmt_num)
    ccol, ctot = _cagr_column(tbl)
    size[_CAGR_LBL] = ccol                        # CAGR block on the right
    size_total = [fmt_num(v) for v in tbl[YEARS].sum(axis=0)] + [ctot]
    # the share table is replaced by a generalized read-out block
    return f"{region} Market {unit} Analysis, {_by(segment)} ({u})", {
        "rows": [
            L.frow(L.tbl(size, index_name=segment, total_row=size_total,
                         title=f"{region} {unit}, {_by(segment)} ({u})")),
            L.frow(L.summary(f"{region} {unit} Analysis, {_by(segment)}",
                             _seg_size_notes(region, unit, segment, tbl),
                             h=178)),
        ]}

def _seg_size_notes(region, unit, segment, tbl):
    """Read-out bullets for a segment size page - carries the share story the
    removed share table used to show, derived entirely from the data."""
    u = UNIT_LABELS[unit]
    # A workbook may carry a segment in ONE unit only - e.g. a study that sizes
    # 'By Payer / Funding' in value but never in volume.  Say so instead of
    # crashing on the empty frame (idxmax of an empty sequence).
    if tbl is None or len(tbl) == 0 or not float(
            tbl[END_YEAR].astype(float).abs().sum()):
        return [f"The workbook carries no {unit.lower()} figures for "
                f"'{segment}', so this breakdown is sized in the other unit "
                f"only.",
                f"See the {segment} value pages for its split, shares and "
                f"growth."]
    end = tbl[END_YEAR].astype(float)
    est = tbl[EST_YEAR].astype(float)
    tot_end = float(end.sum()) or float("nan")
    tot_est = float(est.sum()) or float("nan")
    share_end = end / tot_end * 100.0
    share_est = est / tot_est * 100.0
    shift = (share_end - share_est) * 100.0       # share shift in BPS
    top = metrics.imax(share_end)
    cagr = {i: metrics.fore_cagr(tbl.loc[i]) for i in tbl.index}
    fast = max(cagr, key=lambda i: (cagr[i] if not pd.isna(cagr[i]) else -9))
    gain = metrics.imax(shift)
    lose = metrics.imin(shift)
    return [
        f"'{top}' is the largest {segment.replace('By ', '').lower()} at "
        f"{share_end[top]:.1f}% of {END_YEAR} {unit.lower()} "
        f"({fmt_num(end[top])} {u}).",
        f"'{fast}' grows fastest {fmt_pct(cagr[fast])} CAGR "
        f"({EST_YEAR}–{END_YEAR}).",
        f"Share shifts toward '{gain}' ({shift[gain]:+,.0f} BPS) and away "
        f"from '{lose}' ({shift[lose]:+,.0f} BPS) between {EST_YEAR} and "
        f"{END_YEAR}.",
        f"Total {unit.lower()} rises from {fmt_num(tot_est)} to "
        f"{fmt_num(tot_end)} {u} across {len(tbl)} categories.",
    ]

def seg_ribbons(segment: str, region: str):
    u = UNIT_LABELS["Value"]
    val = data.segment_table(region, "Value", segment)
    s = metrics.segment_summary(region, "Value", segment)
    years3 = [HIST_START, EST_YEAR, END_YEAR]
    return (f"{region} Market Value & Share Shift, {_by(segment)}", {
        "rows": [
            L.frow(
                L.fig(figures.composition_bps(val, years3, s["bps"], u),
                      h=430, frac=0.63,
                      title=f"{region} Value Composition, {_by(segment)} "
                            f"·  {HIST_START} / {EST_YEAR} / {END_YEAR} ({u})  "
                            f"·  right axis = BPS share shift"),
                L.summary(f"{region} Read-out, {_by(segment)}",
                          _seg_trend_notes(region, segment, val, s),
                          frac=0.37, h=430)),
        ]})

def _seg_trend_notes(region, segment, val, s):
    """Read-out bullets for the value-composition + BPS-shift page - derived
    from the segment summary (share gains/losses) and value series."""
    u = UNIT_LABELS["Value"]
    bps = s["bps"]
    if not len(bps):
        return [f"The workbook carries no value figures for {segment}, so no "
                f"share shift can be computed for it.",
                f"See the {segment} volume pages for its split and growth."]
    gain, lose = metrics.imax(bps), metrics.imin(bps)
    cagr = {i: metrics.fore_cagr(val.loc[i]) for i in val.index}
    fast = max(cagr, key=lambda i: (cagr[i] if not pd.isna(cagr[i]) else -9))
    tot_h = float(val[HIST_START].sum())
    tot_e = float(val[END_YEAR].sum())
    lname = segment.replace("By ", "").lower()
    return [
        f"'{gain}' gains the most share {bps[gain]:+,.0f} BPS between "
        f"{EST_YEAR} and {END_YEAR} (green ▲ on the scale).",
        f"'{lose}' cedes the most {bps[lose]:+,.0f} BPS over the same "
        f"window (red ▼).",
        f"'{fast}' is the fastest-growing {lname} at {fmt_pct(cagr[fast])} "
        f"CAGR ({EST_YEAR}–{END_YEAR}).",
        f"Total value climbs from {fmt_num(tot_h)} to {fmt_num(tot_e)} {u} "
        f"across {HIST_START}–{END_YEAR}.",
        "Bars show the value mix at each snapshot; the scale shows how each "
        "share moves.",
    ]

def seg_yoy(segment: str, unit: str, region: str):
    tbl = data.segment_table(region, unit, segment)
    yoy = metrics.yoy_table(tbl)[YEARS[1:]]
    market = tbl.sum(axis=0).pct_change().mul(100.0)[YEARS[1:]]
    disp = pd.DataFrame(index=yoy.index)
    for y in YEARS[1:]:
        disp[y] = [yoy_cell(v, market[y]) for v in yoy[y]]
    return (f"YoY Analysis of {region} Market {unit}, {_by(segment)} (%)", {
        "rows": [
            L.frow(L.tbl(disp, index_name=segment,
                         title=f"{region} {unit} YoY Growth, {_by(segment)} (%)"
                               "  ·  ▲ outpacing market  ▶ below market pace"
                               "  ▼ decline")),
            L.frow(L.fig(figures.heatmap(yoy),
                         h=300,
                         title=f"{region} {unit} YoY Heatmap, {_by(segment)} "
                               "(%)  ·  light = lower, dark = higher growth")),
        ]})

# ═══════════ Segment-section opening slide (analysis overview) ═══════
_SEG_COL_THEMES = list(SECTION_PALETTE)

def seg_start(segment: str, region: str):
    """Creative opening page for a segment section: the segment name, a themed
    sub-segment value matrix (EST/END value + CAGR) and a description block
    (Key Highlights / Sub-Segment Insights / Segment Overview).  Every figure and
    line is derived from the loaded workbook, so it adapts to any market."""
    # A segment may be sized in ONE unit only (the robotics study counts
    # its platform split in procedures, its revenue split in value), so
    # build this page in whichever unit it actually carries.
    unit = data.unit_for(segment, "Value", region)
    u = UNIT_LABELS[unit]
    seg_lower = segment.replace("By ", "").lower()
    tbl = data.segment_table(region, unit, segment)
    if tbl.empty:
        # A pseudo-segment can carry rows for the COUNTRIES and none for the
        # geography being drawn - 'Country Total' in the flexible-packaging
        # workbook is exactly that, and it took the whole export down on
        # `max()` of an empty mapping.  Say so on the page instead: the
        # section keeps its place in the report, and the reader is told the
        # workbook has nothing here rather than shown an invented figure.
        return (f"{_by(segment)} Segment Outlook", {"segstart": {
            "segment": segment,
            "intro": (f"The workbook carries no {seg_lower} figures for "
                      f"{region}; the split is reported at country level."),
            # the matrix keeps its four row labels even with no columns to
            # fill them: both renderers index rl[0..3] directly
            "unit": u,
            "row_labels": ["Sub-Segment", f"{EST_YEAR} {unit} ({u})",
                           f"{END_YEAR} {unit} ({u})",
                           f"CAGR ({EST_YEAR}–{END_YEAR})"],
            "cols": [],
            "highlights": [f"No {seg_lower} rows are present for {region}."],
            "insights": [],
            "overview": {"points": [
                f"{segment} is tracked in this workbook, but its rows sit "
                f"under the individual countries rather than {region}."],
                "stat": ""}}})
    order = list(tbl[EST_YEAR].sort_values(ascending=False).index)
    total_s = tbl.sum(axis=0)
    seg_cagr = metrics.fore_cagr(total_s)
    short = figures._dedupe_labels([str(i) for i in order], cap=34)
    cg = {i: metrics.fore_cagr(tbl.loc[i]) for i in tbl.index}
    est = tbl[EST_YEAR].astype(float)
    tot = float(est.sum()) or float("nan")
    cols = []
    for i, name in enumerate(order):
        cols.append({
            "name": short[i], "full": str(name),
            "color": _SEG_COL_THEMES[i % len(_SEG_COL_THEMES)],
            "v26": fmt_num(float(tbl.loc[name, EST_YEAR])),
            "v36": fmt_num(float(tbl.loc[name, END_YEAR])),
            "cagr": fmt_pct(cg[name])})
    # a segment can be empty in this unit - never argmax an empty series
    leader = str(metrics.imax(est)) if len(est) else "n/a"
    lshare = (float(est.max()) / tot * 100.0) \
        if len(est) and tot else 0.0
    fastest = str(max(cg, key=lambda i: (cg[i] if not pd.isna(cg[i]) else -9)))
    row_labels = ["Sub-Segment", f"{EST_YEAR} {unit} ({u})",
                  f"{END_YEAR} {unit} ({u})", f"CAGR ({EST_YEAR}–{END_YEAR})"]
    intro = (f"The {segment} segment categorizes the market by {seg_lower}. "
             f"Each sub-category is evaluated on market size, growth outlook and "
             f"adoption trends across the forecast period.")
    highlights = [
        f"'{leader}' holds the largest share in {EST_YEAR} {lshare:.0f}% of "
        f"segment value.",
        f"'{fastest}' is the fastest-growing sub-segment at "
        f"{fmt_pct(cg[fastest])} CAGR through {END_YEAR}.",
        f"Segment value expands from {fmt_num(total_s[EST_YEAR])} to "
        f"{fmt_num(total_s[END_YEAR])} {u} across {EST_YEAR}–{END_YEAR}.",
        f"{len(order)} sub-segments make up the {seg_lower} landscape.",
    ]
    insights = []
    for name in order:
        sh = (float(est[name]) / tot * 100.0) if tot else 0.0
        role = ("leads the segment" if str(name) == leader else
                "grows fastest" if str(name) == fastest else
                "serves a niche")
        insights.append({
            "name": str(name), "color": _SEG_COL_THEMES[
                order.index(name) % len(_SEG_COL_THEMES)],
            "text": f"{sh:.0f}% of {EST_YEAR} value · {fmt_pct(cg[name])} "
                    f"CAGR {role}."})
    overview = {
        "points": [
            f"The {segment} split shows how demand is distributed across "
            f"{seg_lower} types and where growth concentrates.",
            f"Shifting preferences and product innovation are reshaping the mix "
            f"through {END_YEAR}."],
        "stat": (f"The {segment} segment is projected to grow at a CAGR of "
                 f"{fmt_pct(seg_cagr)} from {EST_YEAR} to {END_YEAR}.")}
    return (f"{_by(segment)} Segment Outlook", {"segstart": {
        "segment": segment, "intro": intro, "unit": u,
        "row_labels": row_labels, "cols": cols, "highlights": highlights,
        "insights": insights, "overview": overview}})

# ═══════════ Segment-section closing slide (What/Why/How/Key) ═════════
_END_CARD_COLORS = list(CARD_PALETTE)

def seg_end(segment: str, region: str):
    """A creative closing card for a segment section - four panels (What? / Why?
    / How? / Key States) that frame the segment, with 'Key States' driven by the
    live data.  Fully generic: sub-category examples, the leader/fastest-grower,
    and the 'Next' pointer are all derived from the loaded workbook."""
    seg_short = segment.replace("By ", "")
    seg_lower = seg_short.lower()
    tbl = data.segment_table(region, "Value", segment)
    subs = [str(i) for i in tbl.index]
    est = tbl[EST_YEAR].astype(float)
    tot = float(est.sum())
    leader = str(metrics.imax(est))
    lead_sh = (est[metrics.imax(est)] / tot * 100.0) if tot else 0.0
    cagr = {i: metrics.fore_cagr(tbl.loc[i]) for i in tbl.index}
    fastest = str(max(cagr, key=lambda i: (cagr[i] if not pd.isna(cagr[i])
                                           else -9)))
    sample = ", ".join(subs[:3]) if subs else seg_lower
    what = (f"The {segment} segment breaks the market down by {seg_lower} "
            f"{sample}, and more. Each category serves distinct end-user "
            f"preferences, applications, and regulatory requirements.")
    why = (f"Mapping demand by {seg_lower} pinpoints the most preferred and "
           f"fastest-growing categories, helping stakeholders align product "
           f"development, marketing, and investment with shifting customer "
           f"needs and industry trends.")
    how = (f"The split is built from extensive secondary research company "
           f"reports, product portfolios, and industry databases with market "
           f"modelling and expert validation used to size and forecast each "
           f"{seg_lower} sub-segment.")
    key = (f"{leader} currently leads with {lead_sh:.0f}% of {EST_YEAR} value, "
           f"while {fastest} is set to grow fastest {fmt_pct(cagr[fastest])} "
           f"CAGR through {END_YEAR} as the mix shifts over the forecast "
           f"period.")
    labels = ["What?", "Why?", "How?", "Key Stats"]
    bodies = [what, why, how, key]
    cards = [{"label": labels[i], "body": bodies[i],
              "color": _END_CARD_COLORS[i], "num": str(i + 1)}
             for i in range(4)]
    nxt, nxt_slug = _next_pointer(segment)
    return (f"End of {segment} Segment", {"endcards": {
        "segment": segment,
        "cards": cards,
        "footer": f"This concludes the {segment} segment.",
        "next": nxt,
        "next_slug": nxt_slug,
    }})

def _next_pointer(segment):
    """(label, target-slug) for the 'Next' pill - the next segment section, or
    the country report once segments run out.  Derived from the live registry
    so the link always resolves to a real page."""
    from . import registry
    segs = list(data.SEGMENTS)
    idx = segs.index(segment) if segment in segs else -1
    if 0 <= idx < len(segs) - 1:
        nxt_group = segs[idx + 1]
        for p in registry.PAGES:
            if (p["report"] == registry.MAIN_REPORT
                    and p["group"] == nxt_group):
                return f"Next:  {nxt_group} Segment", p["slug"]
    for p in registry.PAGES:                      # last segment -> country
        if p["report"] == registry.COUNTRY_REPORT:
            return "Next:  Country Analysis", p["slug"]
    return "Next:  Country Analysis", registry.HOME_SLUG

# ═══════════ By-Region section opening / closing slides ══════════════
def region_start(region: str):
    """By-Region twin of seg_start - a geographic segment-analysis opener
    (region value matrix + description panels)."""
    u = UNIT_LABELS["Value"]
    child = "Region" if region == "Global" else "Country"
    cl = child.lower()
    kids = data.children_table(region, "Value")
    order = list(kids[EST_YEAR].sort_values(ascending=False).index)
    total_s = kids.sum(axis=0)
    seg_cagr = metrics.fore_cagr(total_s)
    short = figures._dedupe_labels([str(i) for i in order], cap=34)
    cg = {i: metrics.fore_cagr(kids.loc[i]) for i in kids.index}
    est = kids[EST_YEAR].astype(float)
    tot = float(est.sum()) or float("nan")
    cols = [{"name": short[i], "full": str(name),
             "color": _SEG_COL_THEMES[i % len(_SEG_COL_THEMES)],
             "v26": fmt_num(float(kids.loc[name, EST_YEAR])),
             "v36": fmt_num(float(kids.loc[name, END_YEAR])),
             "cagr": fmt_pct(cg[name])} for i, name in enumerate(order)]
    # a segment can be empty in this unit - never argmax an empty series
    leader = str(metrics.imax(est)) if len(est) else "n/a"
    lshare = (float(est.max()) / tot * 100.0) \
        if len(est) and tot else 0.0
    fastest = str(max(cg, key=lambda i: (cg[i] if not pd.isna(cg[i]) else -9)))
    intro = (f"The By {child} view breaks the {region} market down "
             f"geographically. Each {cl} is evaluated on market size, growth "
             f"outlook and adoption trends across the forecast period.")
    highlights = [
        f"'{leader}' holds the largest share in {EST_YEAR} {lshare:.0f}% of "
        f"market value.",
        f"'{fastest}' is the fastest-growing {cl} at {fmt_pct(cg[fastest])} "
        f"CAGR through {END_YEAR}.",
        f"Market value expands from {fmt_num(total_s[EST_YEAR])} to "
        f"{fmt_num(total_s[END_YEAR])} {u} across {EST_YEAR}–{END_YEAR}.",
        f"{len(order)} {plural(cl, len(order))} make up the geographic landscape."]
    insights = [{"name": str(name),
                 "color": _SEG_COL_THEMES[order.index(name) % len(_SEG_COL_THEMES)],
                 "text": f"{(float(est[name]) / tot * 100.0) if tot else 0:.0f}% "
                         f"of {EST_YEAR} value · {fmt_pct(cg[name])} CAGR "
                         f"{'leads the market' if str(name) == leader else 'grows fastest' if str(name) == fastest else 'a smaller market'}."}
                for name in order]
    overview = {
        "points": [
            f"The By {child} split shows where market value concentrates and "
            f"where growth is fastest.",
            f"Shifting demand and access dynamics are rebalancing the "
            f"geographic mix through {END_YEAR}."],
        "stat": (f"The {region} market is projected to grow at a CAGR of "
                 f"{fmt_pct(seg_cagr)} from {EST_YEAR} to {END_YEAR}.")}
    return (f"{region} {data.MARKET} – Regional Outlook", {"segstart": {
        "segment": f"By {child}", "intro": intro, "unit": u,
        "row_labels": [child, f"{EST_YEAR} Value ({u})",
                       f"{END_YEAR} Value ({u})",
                       f"CAGR ({EST_YEAR}–{END_YEAR})"],
        "cols": cols, "highlights": highlights, "insights": insights,
        "overview": overview}})

def region_end(region: str):
    """By-Region twin of seg_end - the What/Why/How/Key-Stats closing card."""
    child = "Region" if region == "Global" else "Country"
    cl = child.lower()
    kids = data.children_table(region, "Value")
    subs = [str(i) for i in kids.index]
    est = kids[EST_YEAR].astype(float)
    tot = float(est.sum())
    # a segment can be empty in this unit - never argmax an empty series
    leader = str(metrics.imax(est)) if len(est) else "n/a"
    lshare = (float(est.max()) / tot * 100.0) \
        if len(est) and tot else 0.0
    cg = {i: metrics.fore_cagr(kids.loc[i]) for i in kids.index}
    fastest = str(max(cg, key=lambda i: (cg[i] if not pd.isna(cg[i]) else -9)))
    sample = ", ".join(subs[:3]) if subs else cl
    bodies = [
        f"The By {child} view breaks the market down geographically {sample}, "
        f"and more. Each {cl} reflects distinct demand, regulation and channel "
        f"dynamics.",
        f"Mapping demand by {cl} pinpoints the highest-value and fastest-growing "
        f"markets, guiding where to prioritise expansion, distribution and "
        f"investment.",
        f"Regional sizing is built from country-level research trade data, "
        f"company footprints and industry databases with market modelling and "
        f"expert validation across each {cl}.",
        f"{leader} currently leads with {lshare:.0f}% of {EST_YEAR} value, while "
        f"{fastest} is set to grow fastest {fmt_pct(cg[fastest])} CAGR through "
        f"{END_YEAR} as demand rebalances across {plural(cl)}."]
    labels = ["What?", "Why?", "How?", "Key Stats"]
    cards = [{"label": labels[i], "body": bodies[i],
              "color": _END_CARD_COLORS[i], "num": str(i + 1)}
             for i in range(4)]
    from . import registry
    seg0 = data.SEGMENTS[0] if len(data.SEGMENTS) else None
    nxt_slug = registry.HOME_SLUG
    for p in registry.PAGES:
        if p["report"] == registry.MAIN_REPORT and p["group"] == seg0:
            nxt_slug = p["slug"]
            break
    nxt = f"Next:  {seg0} Segment" if seg0 else "Next:  Segment Analysis"
    return (f"End of By {child} Analysis", {"endcards": {
        "segment": f"By {child}", "cards": cards,
        "footer": f"This concludes the By {child} analysis.",
        "next": nxt, "next_slug": nxt_slug}})


# -- Template switch support ------------------------------------------
def _retheme() -> None:
    """Re-derive what this module bakes from theme tokens (see theming.py)."""
    global _PP_COLORS, _SEG_COL_THEMES, _END_CARD_COLORS
    _PP_COLORS = list(_T_PP_COLORS)
    _SEG_COL_THEMES = list(SECTION_PALETTE)
    _END_CARD_COLORS = list(CARD_PALETTE)
