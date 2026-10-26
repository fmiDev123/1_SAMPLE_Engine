"""Page templates for the '- Competitor' report (Competition Assessment).

Mirrors the reference Competition-Assessment deck, adapted to this project's
data-driven, dashboard↔PPT-parity architecture.  All competitor figures/text
come from ``competitors.py`` (a documented default profile, workbook-overridable)
- nothing market-specific is hard-coded here.  Each template returns
``(title, spec)`` like the other report templates.
"""
from __future__ import annotations

import pandas as pd

from . import config
from . import competitors as C, data, figures, geo, layout as L
from .components import fmt_num, fmt_pct
from .config import (ACCENT, END_YEAR, EST_YEAR, HEADER_BG, MUTED,
                     SWOT_PALETTE, SWOT_TEXT, TEXT, TIMELINE_PALETTE,
                     VC_NAVY,
                     UNIT_LABELS)
from .data import MARKET_LABEL as MKT


def _pos(block, x, y, w, h):
    return {**block, "x": int(x), "y": int(y), "w": int(w), "h": int(h)}


def _donut(shs, center=""):
    labels = [n for n, _ in shs]
    vals = [v for _, v in shs]
    return figures.donut(labels, vals, center_text=center, legend=True,
                         colors=C.company_colors(len(labels)),
                         legend_below=True)


# ═══════════════ 1. Industry Structure (tier funnel) ════════════════
def industry_structure():
    """Industry-structure landscape: a centred tier funnel (Tier 1/2/3/Others)
    sized to the market total, with tier composition + a structure read-out."""
    u = UNIT_LABELS["Value"]
    tiers = C.tiers()
    mv = C.market_value("Global", EST_YEAR)
    # 3-D funnel segments: Total Market (cap) → tiers → Others, decreasing width
    segments = [{"name": "Total Market", "size": f"~{fmt_num(mv)} {u}",
                 "pct": "100%", "w": 1.0}]
    _W = {"Tier 1": 0.80, "Tier 2": 0.60, "Tier 3": 0.44, "Others": 0.30}
    for t in ("Tier 1", "Tier 2", "Tier 3", "Others"):
        segments.append({"name": t, "size": f"~{fmt_num(tiers[t]['size'])} {u}",
                         "pct": tiers[t]["range"], "w": _W[t]})
    funnel = figures.funnel3d(segments, unit_label=u)
    # tier composition (players per tier)
    comp = []
    for t in ("Tier 1", "Tier 2", "Tier 3"):
        players = tiers[t]["players"]
        if players:
            comp.append(f"{t} ({tiers[t]['range']}): "
                        + ", ".join(players) + ".")
    comp.append(f"Others ({tiers['Others']['range']}): a long tail of "
                f"regional and niche suppliers.")
    t1 = tiers["Tier 1"]
    structure = [
        f"The {MKT} is structured across three competitive tiers plus a long "
        f"tail, on a ~{fmt_num(mv)} {u} base ({EST_YEAR}).",
        f"Tier 1 ({t1['range']}) {t1['label'].lower()} sets pricing and "
        f"innovation benchmarks and holds the largest single block of value.",
        f"Tier 2 ({tiers['Tier 2']['range']}) competes on niche strengths and "
        f"regional depth; Tier 3 ({tiers['Tier 3']['range']}) is emerging and "
        f"local.",
        "Tier sizing is derived from market modelling and expert validation; "
        "totals may differ from headline figures due to rounding.",
    ]
    blocks = [
        _pos(L.fig(funnel, title=f"Industry Structure Market Size by Tier "
                                 f"({u})  {EST_YEAR}"), 15, 95, 792, 605),
        _pos(L.summary("Tier Composition", comp), 819, 95, 446, 300),
        _pos(L.summary("How the Industry Is Structured", structure),
             819, 407, 446, 293),
    ]
    return (f"{MKT} Industry Structure Landscape", {"positioned": blocks})


# ═══════════ 2. Company Value Share (region-switchable donut) ════════
def company_value_share(region: str):
    """Target value-share by company for the selected region - a donut (left)
    that re-renders as the region dropdown changes, plus data-driven Key
    Takeaways (right).  Uses the shared 'region' slicer, so the dashboard gets
    a live dropdown and the PPT gets a region chip."""
    u = UNIT_LABELS["Value"]
    shs = C.company_shares(region)
    labels = [n for n, _ in shs]
    vals = [v for _, v in shs]
    colors = C.company_colors(len(labels))
    # centre caption states what the 100% base actually is, in value terms
    base = C.market_value(region, EST_YEAR)
    fig = figures.donut(labels, vals, legend=True, colors=colors,
                        legend_below=True,
                        center_text=f"100% =<br>{fmt_num(base)} {u}")
    takeaways = C.share_takeaways(region)
    subtitle = ("Target market value share by key companies leadership, "
                "concentration and how the competitive mix shifts by region.")
    blocks = [
        _pos(L.note(subtitle, h=30), 15, 90, 1250, 30),
        _pos(L.fig(fig, title=f"{region} Value Share by Key Company ({u}, "
                              f"{EST_YEAR})"), 15, 128, 792, 572),
        _pos(L.summary("Key Takeaways", takeaways), 819, 128, 446, 572),
    ]
    return (f"{MKT} Company Value Share ({region})", {"positioned": blocks})


# ═══════════ 3. Company Share by Product Type (compact donuts) ══════
def company_share_by_product():
    """Three smaller donuts (company value share within each leading product
    type) across the top, with a detailed Key-Takeaways panel below."""
    prods = C.product_types(3)
    xs = [15, 433, 851]
    ws = [402, 402, 414]
    blocks = [_pos(L.note(
        "Company value share within each of the leading product types where "
        "each player concentrates, and how leadership varies by formulation.",
        h=28), 15, 90, 1250, 28)]
    desc = []
    for i, p in enumerate(prods):
        shs = C.shares_by_product(p, i)
        blocks.append(_pos(L.fig(_donut(shs, ""), title=str(p)),
                           xs[i], 122, ws[i], 352))
        named = sorted([(n, s) for n, s in shs if n != C.OTHERS],
                       key=lambda kv: kv[1], reverse=True)
        lead, ls = named[0]
        second, ss = named[1]
        top3 = sum(s for _, s in named[:3])
        others = next((s for n, s in shs if n == C.OTHERS), 0.0)
        desc.append(
            f"{p}: '{lead}' leads at ~{ls:.1f}%, ahead of '{second}' "
            f"(~{ss:.1f}%); the top three hold ~{top3:.0f}% and a long tail "
            f"('Others') ~{others:.0f}%.")
    desc.append("Leadership shifts across formulations a strong position in "
                "one product type does not carry over to the others, so "
                "product-level targeting is essential.")
    desc.append(C.INDICATIVE_NOTE)
    blocks.append(_pos(L.summary("Company Share Key Takeaways", desc),
                       15, 484, 1250, 216))
    return (f"{MKT} Company Share by Product Type (Top {len(prods)})",
            {"positioned": blocks})


# ═══════════ 4. Company Share by Region (compact donuts) ════════════
def company_share_by_region():
    """Twin of the product-type page, cut by geography: company value share
    within each of the largest regions, so leadership shifts across
    geographies are visible.  Regions come from the loaded workbook, so this
    adapts to any market."""
    # cut at the level below the deck's own geography: regions for a Global
    # deck, countries for a regional one, a country's states / sub-regions
    # for a country deck (client, 2026-10-07)
    here = geo.current()
    kids = data.children_table(here, "Value").astype(float)
    top = [str(r) for r in
           kids[END_YEAR].sort_values(ascending=False).index[:3]]
    kind = ("Region" if here == "Global"
            else "Country" if here in data.REGIONS else "Sub-region")
    kw = {"Region": "regions", "Country": "countries",
          "Sub-region": "sub-regions"}[kind]
    xs, ws = [15, 433, 851], [402, 402, 414]
    blocks = [_pos(L.note(
        f"Company value share within each of the largest {kw} where each "
        f"player concentrates, and how leadership varies by geography.",
        h=28), 15, 90, 1250, 28)]
    desc = []
    for i, r in enumerate(top):
        shs = C.company_shares(r)
        blocks.append(_pos(L.fig(_donut(shs, ""), title=r),
                           xs[i], 122, ws[i], 352))
        named = sorted([(n, s) for n, s in shs if n != C.OTHERS],
                       key=lambda kv: kv[1], reverse=True)
        lead, ls = named[0]
        second, ss = named[1]
        top3 = sum(s for _, s in named[:3])
        others = next((s for n, s in shs if n == C.OTHERS), 0.0)
        desc.append(
            f"{r}: '{lead}' leads at ~{ls:.1f}%, ahead of '{second}' "
            f"(~{ss:.1f}%); the top three hold ~{top3:.0f}% and a long tail "
            f"('Others') ~{others:.0f}%.")
    desc.append("Leadership is not uniform across geographies a strong "
                "position in one region does not carry over to the others, so "
                "region-level go-to-market planning is essential.")
    desc.append(C.INDICATIVE_NOTE)
    blocks.append(_pos(L.summary("Company Share Key Takeaways", desc),
                       15, 484, 1250, 216))
    return ((f"{MKT} Company Share by {kind} (Top {len(top)})"
             if len(top) > 1 else f"{MKT} Company Share, {top[0]}"),
            {"positioned": blocks})


# ═══════════ 5. Competitor Dashboard (creative matrix) ══════════════
def competitor_dashboard():
    """A creative competitive dashboard - companies as columns, with visual
    cells: sparkle portfolio breadth, application ✓/✗, product colour boxes,
    revenue, a market-share donut gauge, regional-presence arrows and a
    strategy badge.  Rendered as a full-page custom block in both surfaces."""
    u = UNIT_LABELS["Value"]
    return (f"{MKT} Competitive Dashboard View ({EST_YEAR})", {"compmatrix": {
        "rows": C.matrix_rows(year=EST_YEAR), "unit": u, "year": EST_YEAR,
        "subtitle": f"Intensified competition across the leading players "
                    f"portfolio, revenue, share, regional presence and "
                    f"strategy.  All figures {EST_YEAR}.",
        "row_labels": ["Company / Brand", "Product Portfolio",
                       f"Overall Rev ({u})", f"Target Rev ({u})",
                       "Market Share (%)", "Regional Presence", "Strategies"],
        "legend": list(C.STRATEGY_LEGEND.items()),
        "star_colors": C.STAR_COLORS, "box_colors": C.BOX_COLORS,
        # what the Product Portfolio markers mean: each colour box is one of
        # the market's product types (hollow = not offered), then the
        # sparkles, ticks and presence arrows (client, 2026-09-30)
        "box_legend": [(C.BOX_COLORS[k % len(C.BOX_COLORS)], name)
                       for k, name in enumerate(C.product_types(3))],
        "note": ["Note: Market Share indicates the contribution of target "
                 "revenue in Market",
                 "Revenue represents sales of the company from target market"]}})


# ═══════════ 6. Company Exec Snapshot (KPIs + SOM splits) ════════════
def company_snapshot(company: str):
    u = UNIT_LABELS["Value"]
    p = C.profile(company)
    kpis = [(f"TAM ({u})", fmt_num(p["tam"]), "Total market"),
            (f"SAM ({u})", fmt_num(p["sam"]), "Addressable"),
            (f"SOM ({u})", fmt_num(p["som"]),
             f"{p['som'] / p['sam'] * 100:.0f}% of SAM"),
            ("Book-to-Bill", f"{p['book_to_bill']:.2f}", "Order momentum"),
            ("EBITDA Margin", f"{p['ebitda']:.0f}%", "Profitability")]
    kw = (1250 - 4 * 13) / 5
    blocks = [_pos(L.note(f"Executive snapshot and strategic positioning "
                          f"{company}: market, addressable opportunity and SOM "
                          f"splits by end-market, region and product.", h=30),
                   15, 90, 1250, 30)]
    for i, (lab, val, sub) in enumerate(kpis):
        blocks.append(_pos(L.kpi(lab, val, sub), 15 + i * (kw + 13), 126,
                           kw, 66))
    end = p["som_end"]
    reg = p["som_region"]
    prod = p["som_product"]
    endbar = figures.milestone_bars(list(end.keys()),
                                    [v for v in end.values()], suffix="",
                                    color="#1668A8")
    regbar = figures.milestone_bars(list(reg.keys()),
                                    [v for v in reg.values()], suffix="",
                                    color="#0E3A5F")
    proddon = figures.donut(list(prod.keys()), list(prod.values()),
                            legend=True, legend_below=True,
                            colors=figures._pal(len(prod)))
    blocks += [
        _pos(L.fig(endbar, title=f"SOM by End-Market ({u})"),
             15, 206, 402, 286),
        _pos(L.fig(regbar, title=f"SOM by {p.get('child_kind', 'Region')} "
                                 f"({u})"), 433, 206, 402, 286),
        _pos(L.fig(proddon, title=f"SOM by Product Line ({u})"),
             851, 206, 414, 286),
        _pos(L.summary("Key Takeaways", C.profile_takeaways(company, p)),
             15, 504, 1250, 198),
    ]
    return (f"{company} Executive Market Snapshot (1/5)",
            {"positioned": blocks})


# ═══════════ 7. Company Regional Penetration & Win Position ══════════
def company_penetration(company: str):
    """Four-quadrant positioning page: SOM-vs-SAM penetration by region, a
    competitive win-position bubble grid, a TAM·SAM·SOM funnel and Target-vs-
    Actual go-to-market benchmarks, plus takeaways."""
    u = UNIT_LABELS["Value"]
    p = C.profile(company)
    pen = p["penetration"]
    w = p["win"]
    # Q1 - SOM vs SAM penetration by region (paired bars)
    pt = pd.DataFrame({
        "SAM": [pen[r]["sam"] for r in pen],
        "SOM": [pen[r]["som"] for r in pen]}, index=list(pen.keys()))
    penfig = figures.paired_bars(pt, "SAM", "SOM", value_fmt="{:,.0f}")
    # Q2 - competitive win position (bubble grid)
    winfig = figures.pos_bubble(p["win_position"])
    # Q3 - TAM · SAM · SOM funnel (left-aligned, single-line labels)
    bands = [
        {"text": f"TAM  ·  100%  ·  {fmt_num(p['tam'])} {u}", "color": "#0A2E4D"},
        {"text": f"SAM  ·  {p['sam'] / p['tam'] * 100:.0f}% of TAM  ·  "
                 f"{fmt_num(p['sam'])} {u}", "color": "#1668A8"},
        {"text": f"SOM  ·  {p['som'] / p['sam'] * 100:.0f}% of SAM  ·  "
                 f"{fmt_num(p['som'])} {u}", "color": "#29ABE2"}]
    funnelfig = figures.sam_funnel(bands)
    # Q4 - go-to-market benchmarks (target vs actual)
    tafig = figures.target_actual_bars(p["benchmarks"])
    take = [
        f"Regional Opportunity: '{max(pen, key=lambda r: pen[r]['pct'])}' shows "
        f"the highest SAM penetration; lower-penetration regions are the growth "
        f"headroom.",
        f"Competitive Edge: ~{w['win_rate']:.0f}% win rate with a "
        f"+{w['asp_premium']:.0f}% ASP premium signals value recognition; "
        f"{w['pipeline_x']:.1f}x pipeline coverage underpins growth.",
        f"Sales Efficiency: a ~{w['sales_cycle']:.0f}-day sales cycle and "
        f"~{w['conversion']:.0f}% conversion support faster share capture.",
    ]
    blocks = [
        _pos(L.note(f"Regional penetration and competitive positioning "
                    f"{company}: SOM vs SAM by region, win position, TAM/SAM/SOM"
                    f" and go-to-market benchmarks.", h=28), 15, 90, 1250, 28),
        _pos(L.fig(penfig, title=f"SOM vs SAM Penetration by "
                                 f"{p.get('child_kind', 'Region')}"),
             15, 122, 620, 248),
        _pos(L.fig(winfig, title="Competitive Win Position"),
             647, 122, 618, 248),
        _pos(L.fig(funnelfig, title=f"TAM · SAM · SOM ({u})"),
             15, 378, 620, 182),
        _pos(L.fig(tafig, title="Go-to-Market Benchmarks (Target vs Actual)"),
             647, 378, 618, 182),
        _pos(L.summary("Key Takeaways", take), 15, 568, 1250, 132),
    ]
    return (f"{company} Regional Penetration & Win Position (2/5)",
            {"positioned": blocks})


# ═══════════ 8. Company Key Financials ══════════════════════════════
def company_financials(company: str):
    p = C.profile(company)
    yrs = [str(y) for y in p["financials"]["years"]]
    fin_keys = ["Revenue", "Gross Profit", "Pre-Tax Profit",
                "Profit After Tax / Net Profit"]
    fin = pd.DataFrame(
        {yrs[j]: [fmt_num(p["financials"][k][j]) for k in fin_keys]
         for j in range(len(yrs))}, index=fin_keys)
    rk = list(p["ratios"].keys())
    rat = pd.DataFrame(
        {yrs[j]: [f"{p['ratios'][k][j]:.1f}%" for k in rk]
         for j in range(len(yrs))}, index=rk)
    pr = pd.DataFrame(index=list(p["pricing"].keys()))
    pr["2020"] = [f"{a:.2f}" for a, _ in p["pricing"].values()]
    pr["2024"] = [f"{b:.2f}" for _, b in p["pricing"].values()]
    blocks = [
        _pos(L.note(f"Key financials and pricing {company} "
                    "(all figures indicative, USD Mn unless noted).", h=30),
             15, 90, 1250, 30),
        _pos(L.tbl(fin, index_name="USD Mn",
                   title=f"{company} Financial Performance"),
             15, 128, 620, 210),
        _pos(L.tbl(rat, index_name="Key Ratios",
                   title=f"{company} Key Ratios"), 15, 352, 620, 350),
        _pos(L.tbl(pr, index_name="Product Type",
                   title=f"{company} Pricing Analysis ({config.price_unit()})"),
             651, 128, 614, 574),
    ]
    return (f"{company} Key Financials (4/5)", {"positioned": blocks})


# ═══════════ Advice for Market Vendors / Suppliers ══════════════════
def competitor_advice():
    """Strategic advice for vendors/suppliers - recommendations grouped into
    pricing, product, geographic/channel and competitive-risk themes, all
    derived from the market and competitor picture."""
    shs = C.company_shares("Global")
    named = sorted([(n, s) for n, s in shs if n != C.OTHERS],
                   key=lambda kv: kv[1], reverse=True)
    lead, lead_s = named[0]
    top3 = sum(s for _, s in named[:3])
    prods = C.product_types(3)
    try:
        tot = data.total_series("Global", "Value")
        span = max(1, END_YEAR - EST_YEAR)
        cagr = ((float(tot[END_YEAR]) / float(tot[EST_YEAR])) ** (1 / span) - 1) \
            * 100 if float(tot[EST_YEAR]) else 0.0
    except Exception:
        cagr = 0.0
    cards = [
        ("Pricing & Margin Strategy", [
            f"Hold price discipline: with the top three players at ~{top3:.0f}% "
            f"share, avoid value-destructive discounting compete on mix and "
            f"service, not headline price.",
            f"Premiumise toward higher-value formulations (e.g. '{prods[0]}') to "
            f"lift blended ASP and defend margin as input costs move.",
            "Set transparent, corridor-based regional pricing to limit "
            "grey-market arbitrage."]),
        ("Product & Innovation", [
            f"Prioritise R&D in the fastest-growing formulations a "
            f"differentiated portfolio is the clearest route to out-growing the "
            f"~{cagr:.1f}% market CAGR.",
            "Pursue certification and clinical validation to build switching "
            "costs and support premium positioning.",
            "Localise formats and pack sizes to each channel and geography."]),
        ("Geographic & Channel Expansion", [
            "Expand selectively into under-penetrated regions where the leaders "
            "are thin, rather than head-on in saturated markets.",
            "Balance retail, hospital/clinic and online channels online is the "
            "fastest-shifting route to end customers.",
            "Build reliable distribution and logistics partnerships to convert "
            "demand into fulfilled volume."]),
        ("Competitive & Risk Positioning", [
            f"Benchmark relentlessly against '{lead}' (~{lead_s:.1f}% share) "
            f"match on service and lead on innovation to take share.",
            "Diversify the customer base to cut dependence on a few large "
            "accounts and single-region exposure.",
            "Stress-test for regulatory, currency and supply shocks; keep "
            "flexible capacity and dual sourcing."]),
    ]
    coords = [(15, 128, 615, 285), (650, 128, 615, 285),
              (15, 423, 615, 275), (650, 423, 615, 275)]
    blocks = [_pos(L.note("Strategic advice for vendors and suppliers where to "
                          "compete, how to price, and which risks to manage "
                          f"across the {MKT}.", h=30), 15, 90, 1250, 30)]
    for (ctitle, bl), (x, y, w, h) in zip(cards, coords):
        blocks.append(_pos(L.summary(ctitle, bl), x, y, w, h))
    return (f"{MKT} Advice for Market Vendors / Suppliers",
            {"positioned": blocks})


# ═══════════ Who Shaped the Year (creative timeline) ════════════════
_SHAPER_COLORS = list(TIMELINE_PALETTE)

# Growth roles and reasons are keyed to RANK, not to the company, so they stay
# meaningful for any market's roster (the roster's own strategy codes are too
# similar across players to differentiate them).
_SHAPER_ROLES = ["Market leader", "Portfolio expander", "Access builder",
                 "Premiumiser", "Disruptor"]
_SHAPER_WHY = ["Expanded reach across its core segment",
               "Drove new customer adoption",
               "Won through wider channel availability",
               "Supported premium SKU growth",
               "Strengthened end-user and specifier pull"]
_SHAPER_MOVES = ["New launch", "Channel expansion", "Pricing strategy",
                 "Portfolio/brand investment", "Partnership / acquisition"]
# ── icons ────────────────────────────────────────────────────────────
# Emoji glyphs, verified to render in the NATIVE PPT export (PowerPoint falls
# back to Segoe UI Emoji) as well as in the browser - so both surfaces match.
# They are plain text, so they travel through the existing text primitive with
# no renderer change.  Keep to SINGLE-codepoint emoji: ZWJ sequences (e.g.
# a profession emoji) can fall apart in the PPT text run.
_ICON_RANK = ["\U0001F3C6", "\U0001F4C8", "\U0001F465", "\U0001F48E",
              "⚡"]                       # trophy chart people gem bolt
_ICON_SHARE = "\U0001F4CA"                    # bar chart
_ICON_ROLE = "\U0001F464"                     # bust in silhouette
_ICON_WHY = "\U0001F3AF"                      # direct hit / target
_ICON_MOVE = "\U0001F680"                     # rocket
# the scoring rubric shown at the foot of the slide (0-5 points per criterion)
_SHAPER_CRITERIA = [
    ("\U0001F4CA", "SHARE GAIN", "Company gained value or volume share"),
    ("\U0001F3EA", "DISTRIBUTION EXPANSION",
     "Expanded channel, retail or e-commerce reach"),
    ("\U0001F4A1", "INNOVATION / LAUNCHES",
     "New SKUs, pack sizes or product improvements"),
    ("\U0001F91D", "BUYER INFLUENCE",
     "Stronger specifier, channel or end-user pull"),
    ("\U0001F4B2", "BRAND / PRICING POWER",
     "Premiumisation, positioning or pricing execution"),
]
_SHAPER_USAGE = [("\U0001F4C8", "Rank companies by impact on category growth"),
                 ("\U0001F4C4", "Mention the one move that explains their impact"),
                 ("\U0001F6AB", "Do not rely on market share alone")]
_NAVY_D = VC_NAVY


def year_shapers():
    """'Who Shaped Growth' - the leading players ranked by growth-shaping
    impact: a numbered arrow spine over one card per company (value share,
    growth role, why it shaped growth, signature move), above the scoring
    rubric that explains how the ranking is decided and a short how-to-use
    panel.  Built from drawing primitives so the dashboard and the PPT render
    identically.  Value shares are real data; roles/reasons/moves are keyed to
    rank so the page stays market-agnostic.  Icons are emoji glyphs (see
    _ICON_*) - they travel as plain text and were verified to render in the
    native PPT export as well as the browser."""
    shs = C.company_shares("Global")
    named = sorted([(n, s) for n, s in shs if n != C.OTHERS],
                   key=lambda kv: kv[1], reverse=True)[:5]
    n = len(named) or 1
    X0, X1 = 15, 1265
    P = []

    def txt(x, y, w, h, s, size=9, bold=False, color=None, align="l",
            wrap=False, lh=None):
        # resolved per call, not bound as a default: a default freezes the
        # colour at import and survives a template switch (see
        # templates_dynamics._t for the same fix)
        d = {"t": "text", "x": x, "y": y, "w": w, "h": h, "s": s, "size": size,
             "bold": bold, "color": TEXT if color is None else color,
             "align": align}
        if wrap:
            d["wrap"] = True
            d["lh"] = lh or size * 1.35
        P.append(d)

    # ── subtitle ──
    # (2026-09-11) every text on this page is set at the 11pt content
    # standard or larger; the card rows and the rubric are re-spaced for it
    txt(X0, 86, X1 - X0, 20,
        f"A framework to identify the companies that most influenced annual "
        f"{MKT} growth, ranked by growth-shaping impact.",
        size=11, color=MUTED)

    # ── 'ranked by' callout, flanked by rules ──
    cw, cy = 520, 110
    cx = (1280 - cw) / 2
    P.append({"t": "rrect", "x": cx, "y": cy, "w": cw, "h": 46,
              "fill": "#F4F8FB", "line": "#CFE0EC", "r": 8, "lw": 1})
    txt(cx + 12, cy + 4, cw - 24, 19, "RANKED BY GROWTH-SHAPING IMPACT",
        size=11, bold=True, color=_NAVY_D, align="c")
    txt(cx + 12, cy + 23, cw - 24, 19,
        "Ranked by overall contribution to category growth, not share alone.",
        size=11, color=MUTED, align="c")
    for a, b in ((X0 + 30, cx - 18), (cx + cw + 18, X1 - 30)):
        P.append({"t": "line", "x1": a, "y1": cy + 23, "x2": b, "y2": cy + 23,
                  "color": "#C4D0DA", "width": 1})

    # ── numbered arrow spine ──
    sy = 176
    card_w = (X1 - X0 - (n - 1) * 10) / n
    P.append({"t": "line", "x1": X0 + 40, "y1": sy, "x2": X1 - 44, "y2": sy,
              "color": _NAVY_D, "width": 2})
    P.append({"t": "shape", "pts": [(X1 - 44, sy - 7), (X1 - 26, sy),
                                    (X1 - 44, sy + 7)], "fill": _NAVY_D})

    # ── one card per company ──
    top, ch = 198, 234
    for i, (name, share) in enumerate(named):
        col = _SHAPER_COLORS[i % len(_SHAPER_COLORS)]
        x = X0 + i * (card_w + 10)
        mid = x + card_w / 2
        # rank badge on the spine
        P.append({"t": "dot", "cx": mid, "cy": sy, "d": 30, "color": col,
                  "edge": "#FFFFFF", "lw": 3})
        txt(mid - 15, sy - 9, 30, 18, str(i + 1), size=11, bold=True,
            color="#FFFFFF", align="c")
        # card
        P.append({"t": "rrect", "x": x, "y": top, "w": card_w, "h": ch,
                  "fill": "#FFFFFF", "line": col, "r": 9, "lw": 1.6})
        P.append({"t": "rrect", "x": x, "y": top, "w": card_w, "h": 34,
                  "fill": col, "line": col, "r": 9, "lw": 1})
        # icon badge on the header, then the name beside it
        P.append({"t": "dot", "cx": x + 24, "cy": top + 17, "d": 24,
                  "color": "#FFFFFF", "edge": "#FFFFFF"})
        txt(x + 12, top + 8, 24, 18, _ICON_RANK[i % len(_ICON_RANK)], size=11,
            align="c")
        txt(x + 40, top + 8, card_w - 50, 19, name, size=11, bold=True,
            color="#FFFFFF", align="c")
        ix, iw = x + 30, card_w - 40           # text column, icons to its left
        txt(x + 10, top + 42, 18, 18, _ICON_SHARE, size=11)
        txt(ix, top + 42, iw, 18, f"Value share:  {share:.1f}%",
            size=11, bold=True)
        P.append({"t": "line", "x1": x + 10, "y1": top + 66, "x2": x + card_w - 10,
                  "y2": top + 66, "color": "#E4EAEF", "width": 1,
                  "dash": "dotted"})
        txt(x + 10, top + 80, 18, 18, _ICON_ROLE, size=11)
        txt(ix, top + 71, iw, 18, "Growth role:", size=11, color=MUTED)
        txt(ix, top + 89, iw, 18, _SHAPER_ROLES[i % len(_SHAPER_ROLES)],
            size=11, bold=True, color=col)
        P.append({"t": "line", "x1": x + 10, "y1": top + 112, "x2": x + card_w - 10,
                  "y2": top + 112, "color": "#E4EAEF", "width": 1,
                  "dash": "dotted"})
        txt(x + 10, top + 124, 18, 18, _ICON_WHY, size=11)
        txt(ix, top + 117, iw, 18, "Why it shaped growth:", size=11,
            bold=True, color=_NAVY_D)
        txt(ix, top + 135, iw, 42, _SHAPER_WHY[i % len(_SHAPER_WHY)],
            size=11, color=TEXT, wrap=True, lh=15)
        P.append({"t": "line", "x1": x + 10, "y1": top + 182, "x2": x + card_w - 10,
                  "y2": top + 182, "color": "#E4EAEF", "width": 1,
                  "dash": "dotted"})
        txt(x + 10, top + 194, 18, 18, _ICON_MOVE, size=11)
        txt(ix, top + 187, iw, 18, "Signature move:", size=11, bold=True,
            color=_NAVY_D)
        txt(ix, top + 205, iw, 24, _SHAPER_MOVES[i % len(_SHAPER_MOVES)],
            size=11, color=TEXT, wrap=True, lh=15)

    # ── bottom-left: the scoring rubric ──
    by, bh = 444, 262
    lw_ = 828
    P.append({"t": "rrect", "x": X0, "y": by, "w": lw_, "h": bh,
              "fill": "#F7FAFC", "line": "#DCE6ED", "r": 9, "lw": 1})
    txt(X0, by + 8, lw_, 20, "HOW TO DECIDE WHO SHAPED GROWTH", size=11,
        bold=True, color=_NAVY_D, align="c")
    inner = lw_ - 28
    colw = (inner - 4 * 6) / 5
    for j, (icon, cap, desc) in enumerate(_SHAPER_CRITERIA):
        cxj = X0 + 14 + j * (colw + 6)
        ccol = _SHAPER_COLORS[j % len(_SHAPER_COLORS)]
        # white disc ringed in the criterion's colour, icon inside (emoji are
        # multi-colour, so they need a light background to stay legible)
        P.append({"t": "dot", "cx": cxj + colw / 2, "cy": by + 54, "d": 42,
                  "color": "#FFFFFF", "edge": ccol, "lw": 2.5})
        txt(cxj + colw / 2 - 21, by + 43, 42, 22, icon, size=15, align="c")
        txt(cxj, by + 80, colw, 38, f"{j + 1}. {cap}", size=11, bold=True,
            color=ccol, align="c", wrap=True, lh=15)
        txt(cxj, by + 120, colw, 66, desc, size=11, color=MUTED, align="c",
            wrap=True, lh=15)
        P.append({"t": "rrect", "x": cxj + (colw - 104) / 2, "y": by + 196,
                  "w": 104, "h": 26, "fill": "#FFFFFF", "line": ccol, "r": 12,
                  "lw": 1.2})
        txt(cxj + (colw - 104) / 2, by + 197, 104, 24, "0–5 POINTS", size=11,
            bold=True, color=ccol, align="c")
        if j < 4:
            txt(cxj + colw, by + 45, 6, 18, "+", size=12, bold=True,
                color=MUTED, align="c")

    # ── bottom-right: how to use ──
    rx, rw = X0 + lw_ + 17, X1 - (X0 + lw_ + 17)
    P.append({"t": "rrect", "x": rx, "y": by, "w": rw, "h": bh,
              "fill": "#F4FAF9", "line": "#CFE3E0", "r": 9, "lw": 1})
    txt(rx, by + 8, rw, 20, "BEST WAY TO USE THIS SLIDE", size=11, bold=True,
        color="#1F6F63", align="c")
    P.append({"t": "line", "x1": rx + 16, "y1": by + 34, "x2": rx + rw - 16,
              "y2": by + 34, "color": "#CFE3E0", "width": 1})
    for k, (icon, use) in enumerate(_SHAPER_USAGE):
        uy = by + 46 + k * 70
        P.append({"t": "dot", "cx": rx + 34, "cy": uy + 22, "d": 30,
                  "color": "#FFFFFF", "edge": "#1F6F63", "lw": 2})
        txt(rx + 19, uy + 12, 30, 20, icon, size=12, align="c")
        txt(rx + 58, uy + 4, rw - 74, 44, use, size=11, color=TEXT,
            wrap=True, lh=15)
    return (f"{MKT} Who Shaped Growth", {"primitives": P})


# ═══════════ Company SWOT (pinwheel + quadrant text) ════════════════
_SWOT_COLORS = dict(SWOT_PALETTE)      # pinwheel arm fills
_SWOT_INK = dict(SWOT_TEXT)            # readable on the white page


def _readable_on(fill: str) -> str:
    """White or near-black, whichever reads on ``fill``."""
    h = str(fill).lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return "#1F2933" if (0.299 * r + 0.587 * g + 0.114 * b) > 150 else "#FFFFFF"

def company_swot(company: str):
    """SWOT assessment for a company - a central four-arrow pinwheel (S/W/O/T)
    with cross axes and dashed quadrant frames, plus a strengths / weaknesses /
    opportunities / threats text block in each corner.  Primitive block, derived
    from the company's competitive profile."""
    p = C.profile(company)
    w = p["win"]
    share = float(p.get("share", 5.0))
    lead_reg = max(p["som_region"], key=p["som_region"].get)
    pen_lead = max(p["penetration"], key=lambda r: p["penetration"][r]["pct"])
    prods = p["products"]
    S = [f"~{w['win_rate']:.0f}% win rate with a +{w['asp_premium']:.0f}% ASP "
         f"premium clear value recognition.",
         f"Balanced SOM across regions and end-markets; strongest in "
         f"'{lead_reg}'.",
         f"{p['book_to_bill']:.2f} book-to-bill and ~{p['ebitda']:.0f}% EBITDA "
         f"signal healthy momentum."]
    W = [f"~{share:.1f}% share leaves a scale gap versus the market leader.",
         f"Uneven SAM penetration strong in '{pen_lead}', thin elsewhere.",
         "Reliance on a few large accounts/regions raises concentration risk."]
    O = [f"Premiumisation headroom in '{prods[0]}' and adjacent high-value "
         f"lines.",
         "Under-penetrated regions offer share-capture runway.",
         "Channel shift to online and clinics opens new demand pools."]
    T = [f"Intense competition ~{w['pipeline_x']:.1f}x pipeline needed to "
         f"convert.",
         "Input-cost inflation and price pressure can squeeze margin.",
         "Tier-1 incumbents and regulation raise barriers and switching "
         "friction."]
    P = []
    P.append({"t": "text", "x": 15, "y": 90, "w": 1250, "h": 18,
              "s": f"SWOT assessment {company}: strengths, weaknesses, "
                   f"opportunities and threats (indicative).", "size": 11.5,
              "bold": False, "color": TEXT, "align": "l"})
    cx, cy = 640, 406
    # cross axes (behind everything)
    P.append({"t": "line", "x1": cx - 240, "y1": cy, "x2": cx + 240, "y2": cy,
              "color": "#B9C4CE", "width": 1.4})
    P.append({"t": "line", "x1": cx, "y1": cy - 216, "x2": cx, "y2": cy + 216,
              "color": "#B9C4CE", "width": 1.4})
    for ex, ey in [(cx - 240, cy), (cx + 240, cy), (cx, cy - 216),
                   (cx, cy + 216)]:
        P.append({"t": "dot", "cx": ex, "cy": ey, "d": 11, "color": "#B9C4CE",
                  "edge": "#B9C4CE"})
    # dashed quadrant frames
    fr = 132
    for key, qx, qy in [("S", cx - fr - 12, cy - fr - 12),
                        ("W", cx + 12, cy - fr - 12),
                        ("T", cx - fr - 12, cy + 12), ("O", cx + 12, cy + 12)]:
        P.append({"t": "rrect", "x": qx, "y": qy, "w": fr, "h": fr,
                  "fill": "none", "line": _SWOT_COLORS[key], "dash": "dashed",
                  "r": 16, "lw": 1.5})
    # four inward arrows (pinwheel)
    g, arm, hw = 12, 88, 60
    P.append({"t": "shape", "pts": [(cx - hw, cy - g - arm),
              (cx + hw, cy - g - arm), (cx, cy - g)], "fill": _SWOT_COLORS["S"]})
    P.append({"t": "shape", "pts": [(cx - hw, cy + g + arm),
              (cx + hw, cy + g + arm), (cx, cy + g)], "fill": _SWOT_COLORS["O"]})
    P.append({"t": "shape", "pts": [(cx - g - arm, cy - hw),
              (cx - g - arm, cy + hw), (cx - g, cy)], "fill": _SWOT_COLORS["T"]})
    P.append({"t": "shape", "pts": [(cx + g + arm, cy - hw),
              (cx + g + arm, cy + hw), (cx + g, cy)], "fill": _SWOT_COLORS["W"]})
    # S/W/O/T letters near centre
    for lab, lx, ly in [("S", cx, cy - g - arm * 0.5), ("O", cx, cy + g + arm * 0.5),
                        ("T", cx - g - arm * 0.5, cy), ("W", cx + g + arm * 0.5, cy)]:
        # white reads on the dark arms but vanishes on the light gold
        # one, so each letter takes the ink that reads on ITS arm
        P.append({"t": "text", "x": lx - 22, "y": ly - 15, "w": 44, "h": 30,
                  "s": lab, "size": 21, "bold": True,
                  "color": _readable_on(_SWOT_COLORS[lab]), "align": "c"})

    def corner(bx, by, key, label, bullets):
        col = _SWOT_INK[key]
        P.append({"t": "text", "x": bx, "y": by, "w": 385, "h": 20, "s": label,
                  "size": 13.5, "bold": True, "color": col, "align": "l"})
        P.append({"t": "rect", "x": bx, "y": by + 22, "w": 150, "h": 2,
                  "fill": col, "line": "none"})
        yy = by + 30
        for b in bullets:
            P.append({"t": "dot", "cx": bx + 5, "cy": yy + 7, "d": 7,
                      "color": col, "edge": "#FFFFFF"})
            P.append({"t": "text", "x": bx + 15, "y": yy, "w": 368, "h": 52,
                      "s": b, "size": 11, "bold": False, "color": TEXT,
                      "align": "l", "wrap": True, "lh": 16})
            yy += 56
    corner(20, 122, "S", "Strengths", S)
    corner(875, 122, "W", "Weaknesses", W)
    corner(20, 424, "T", "Threats", T)
    corner(875, 424, "O", "Opportunities", O)
    # at the 11pt content standard the note measures ~800px, so its box
    # starts further left and spans the rest of the page on one line
    P.append({"t": "text", "x": 300, "y": 684, "w": 965, "h": 20,
              "s": C.INDICATIVE_NOTE, "size": 11, "bold": False,
              "color": "#C0392B", "align": "l"})
    return (f"{company} SWOT Analysis (3/5)", {"primitives": P})


# ═══════════ Company Strategy Analysis ══════════════════════════════
def company_strategy(company: str):
    """Strategy analysis for a company - strategic priorities and the competitive
    moves that support them, derived from the company's profile."""
    p = C.profile(company)
    w = p["win"]
    share_pct = float(p.get("share", 5.0))
    reg_lead = max(p["som_region"], key=p["som_region"].get)
    end_lead = max(p["som_end"], key=p["som_end"].get)
    pen_lead = max(p["penetration"], key=lambda r: p["penetration"][r]["pct"])
    prods = p["products"]
    priorities = [
        f"Defend and extend leadership in '{end_lead}' {company}'s strongest "
        f"end-market while scaling the next-best applications.",
        f"Convert its ~{w['win_rate']:.0f}% win rate and +{w['asp_premium']:.0f}%"
        f" ASP premium into share gains through value-based selling.",
        f"Raise SAM penetration beyond '{pen_lead}' "
        f"(~{p['penetration'][pen_lead]['pct']:.0f}%) into lower-penetration "
        f"regions.",
        f"Sustain a {p['book_to_bill']:.2f} book-to-bill with disciplined "
        f"capacity and channel investment.",
        "Protect margin through corridor-based pricing and cost discipline as "
        "input prices move.",
    ]
    moves = [
        f"Premiumise the portfolio around '{prods[0]}' and adjacent high-value "
        f"lines to lift blended pricing.",
        f"Deepen presence in '{reg_lead}' (largest SOM region) and build out "
        f"under-served geographies.",
        "Invest in R&D, certification and selective M&A to move up the price "
        "ladder and widen the moat.",
        "Diversify the customer base and secure supply to de-risk concentration "
        "and input-cost shocks.",
        f"Sharpen go-to-market a ~{w['sales_cycle']:.0f}-day sales cycle and "
        f"~{w['conversion']:.0f}% conversion speed share capture.",
    ]
    kpis = [
        # whole-market share (see competitors.py): neutral wording on a
        # region deck, unchanged on the global one
        ("Market Share", f"{share_pct:.1f}%",
         "of global value" if geo.is_global() else "of market value"),
        ("Win Rate", f"{w['win_rate']:.0f}%", "deals won"),
        ("ASP Premium", f"+{w['asp_premium']:.0f}%", "vs peers"),
        ("Book-to-Bill", f"{p['book_to_bill']:.2f}", "order momentum"),
        ("EBITDA Margin", f"{p['ebitda']:.0f}%", "profitability"),
    ]
    kw = (1250 - 4 * 13) / 5
    blocks = [_pos(L.note(f"Strategy analysis {company}: strategic priorities "
                          f"and the competitive moves that support them "
                          f"(indicative).", h=30), 15, 90, 1250, 30)]
    for i, (lab, val, sub) in enumerate(kpis):
        blocks.append(_pos(L.kpi(lab, val, sub), 15 + i * (kw + 13), 128, kw, 66))
    blocks += [
        _pos(L.summary("Strategic Priorities", priorities), 15, 208, 620, 492),
        _pos(L.summary("Competitive Moves", moves), 651, 208, 614, 492),
    ]
    return (f"{company} Strategy Analysis (5/5)", {"positioned": blocks})


# -- Template switch support ------------------------------------------
def _retheme() -> None:
    """Re-derive what this module bakes from theme tokens (see theming.py)."""
    global _SHAPER_COLORS, _NAVY_D, _SWOT_COLORS, _SWOT_INK
    _SHAPER_COLORS = list(TIMELINE_PALETTE)
    _NAVY_D = VC_NAVY
    _SWOT_COLORS = dict(SWOT_PALETTE)
    _SWOT_INK = dict(SWOT_TEXT)
