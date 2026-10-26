"""One-line sub headers for the exported deck.

Every content slide carries a short descriptive line under its title that
says what the exhibit shows, over which period and in which unit.  The line
is derived from the page's TEMPLATE (which decides what is drawn) and the
loaded workbook's own facts (geographies, segments, years, unit labels), so
it adapts to any market the engine is pointed at.

The wording follows the FMI reference decks: an analytical framing line
("Analysis of X, Y and Z shaping ...") that says what the page evaluates and
why it matters, carrying the workbook's own specifics (segment names,
geography counts, years) so every slide's line is its own.  Kept to roughly
120 characters so it sits on one line beside the slicer chips.
"""
from __future__ import annotations

import functools

from . import competitors, config, data, geo

_MAX = 124          # last-resort cap; the exporter measures the real width


def _years():
    c = config
    return c.HIST_START, c.HIST_END, c.EST_YEAR, c.MID_YEAR, c.END_YEAR


def _unit(unit: str) -> str:
    return config.UNIT_LABELS.get(unit, unit)


def _by(segment: str) -> str:
    """'Product Type' -> 'product type' for use inside a sentence."""
    return str(segment).strip()


def _kids(region: str):
    """(kind word, plural phrase) for the geography's children, e.g.
    ('region', '6 regions') under Global or ('country', '12 countries')."""
    kind = geo.child_kind(region)
    try:
        n = len(data.children(region))
    except Exception:
        n = 0
    return kind, (f"{n} {config.plural(kind, n)}" if n else
                  config.plural(kind, 2))


def _subsegs(segment: str) -> str:
    try:
        n = len(data.SUBSEGMENTS.get(segment) or [])
    except Exception:
        n = 0
    return f"{n} sub-segments" if n else "sub-segments"


def _template(page):
    """(function name, positional args) of the page's render callable."""
    r = page["render"]
    if isinstance(r, functools.partial):
        return r.func.__name__, tuple(r.args)
    return r.__name__, ()


def for_page(page: dict, region: str | None, country: str | None) -> str:
    """The sub header for one registry page rendered for (region, country)."""
    name, args = _template(page)
    region = region or geo.current()
    H0, H1, E, M, Y = _years()
    UV, UW = _unit("Volume"), _unit("Value")
    G, g = geo.word(region), geo.adj(region)
    kind, kids = _kids(region)
    c = country or ""
    a0 = args[0] if args else ""
    a1 = args[1] if len(args) > 1 else ""

    def unit_lbl(u):
        return _unit(u)

    text = {
        # ── Market Assessment ──
        "contents_page": lambda: (
            f"Storyline of the report by section · click a section card to "
            f"jump to it"),
        "market_signals": lambda: (
            f"Signals reshaping {data.MARKET} demand, the evidence behind "
            f"each and what it implies for the {g} market"),
        "taxonomy_regions": lambda: (
            f"Geographic scope of the study: the {g} market broken into "
            f"{kids} and the countries assessed under each"),
        "taxonomy_segments": lambda: (
            f"Segmentation framework of the study: {len(data.taxonomy_axes())} "
            f"axes and the sub-segments sized under each"),
        "company_coverage": lambda: (
            f"Companies profiled and benchmarked across the competition "
            f"assessment ({len(competitors.companies())} players)"),
        # ── Global overview ──
        "market_summary": lambda: (
            f"Analysis of {g} market {a0.lower()} growth, incremental "
            f"opportunity and the geographies driving it, {H0}–{Y}"),
        "pricing_analysis": lambda: (
            f"Analysis of average price by sub-segment and price variation "
            f"by {kind}, highlighting where pricing power sits"),
        "vol_val_analysis": lambda: (
            f"Analysis of volume and value growth of the {g} market and the "
            f"incremental $ opportunity it creates, {H0}–{Y}"),
        "summary": lambda: (
            f"Country outlook ranking {Y} market size, growth and "
            f"absolute-growth indices to highlight the strongest markets"),
        "summary_overview": lambda: (
            f"Executive summary of market size, competitive landscape and "
            f"the drivers, restraints and trends shaping demand"),
        "summary_supply": lambda: (
            f"Executive summary of recent developments, investments and R&D "
            f"activity shaping the supply side of the market"),
        "summary_analysis": lambda: (
            f"Executive summary of the market approach, differentiating "
            f"strategy and FMI's analysis and recommendation"),
        "summary_scenario": lambda: (
            f"Scenario outlook comparing optimistic, likely and conservative "
            f"outcomes for {Y} and the CAGR behind each"),
        # ── By region ──
        "region_start": lambda: (
            f"Regional analysis of size, growth and share across {kids}, "
            f"highlighting where value concentrates, {E}–{Y}"),
        "region_donuts": lambda: (
            f"Analysis of {a0.lower()} share by {kind} in {E} and {Y}, "
            f"highlighting share gains and losses in basis points"),
        "region_share_tables": lambda: (
            f"Year-by-year {a0.lower()} and share by {kind}, {H0}–{Y}, "
            f"highlighting the geographies gaining and ceding ground"),
        "region_ribbons": lambda: (
            f"Analysis of how {geo.child_adj(region)} volume and value "
            f"shares shift over {H0}–{Y}, highlighting who gains ground"),
        "market_hierarchy": lambda: (
            f"Country hierarchy of the {g} market in {Y}, highlighting the "
            f"leading countries within each region by value"),
        # ── By segment ──
        "seg_start": lambda: (
            f"Segment analysis of market size, growth and share across "
            f"{_subsegs(a0)} of {_by(a0)}, {E}–{Y}"),
        "seg_bps": lambda: (
            f"{_by(a0)} mix within each {kind} and the attractiveness of "
            f"each sub-segment on growth and size"),
        "seg_absolute": lambda: (
            f"Value build-up by {_by(a0)} and the absolute $ opportunity "
            f"each sub-segment adds, {H0}–{Y}"),
        "seg_size_tables": lambda: (
            f"Year-by-year {a1.lower()} by {_by(a0)}, {H0}–{Y}, "
            f"highlighting the sub-segments driving growth"),
        "seg_ribbons": lambda: (
            f"Value composition by {_by(a0)} in {H0}, {E} and {Y}, "
            f"highlighting share shifts in basis points"),
        "seg_yoy": lambda: (
            f"Year-on-year {a1.lower()} growth by {_by(a0)}, {H0 + 1}–{Y}, "
            f"highlighting acceleration and slowdown"),
        # ── Country assessment ──
        "overview": lambda: (
            f"Overview of the {c} market in {a0.lower()}: size, growth, "
            f"YoY momentum and gap to the {G} benchmark"),
        "seg_overview_value": lambda: (
            f"Overview of {c}'s key segments by value, highlighting the "
            f"leading sub-segment of every segmentation"),
        "segmental": lambda: (
            f"{c} segmental analysis by {_by(a0)}: growth contribution, "
            f"share variance and absolute $ opportunity"),
        "tables": lambda: (
            f"Year-by-year volume and value of the {c} market by {_by(a0)}, "
            f"{H0}–{Y}"),
        "yoy": lambda: (
            f"{c} year-on-year volume and value growth by {_by(a0)}, "
            f"highlighting acceleration and slowdown"),
        # ── Pricing ──
        "pricing_index_trends": lambda: (
            f"Price range by sub-segment and indexed regional price "
            f"movement, highlighting dispersion and structural shifts"),
        "pricing_positioning": lambda: (
            f"Assessment of pricing pattern, price range and average price "
            f"by product for each {kind} of the {g} market"),
        "pricing_key_findings": lambda: (
            f"Price positioning of each {kind} on a common axis with FMI's "
            f"key findings on where premiums and discounts sit"),
        "pricing_drivers": lambda: (
            f"Analysis of input costs, market conditions and competitive "
            f"forces shaping pricing strategies and profitability"),
        "pricing_factors_table": lambda: (
            f"Factors affecting the pricing decision of the product: "
            f"category, price impact, direction of pressure and rationale"),
        "pricing_assessment": lambda: (
            f"Regional price assessment across the value chain, from "
            f"manufacturer to retail, and average price by geography, "
            f"{E}–{Y}"),
        "pricing_benchmarking": lambda: (
            f"Benchmarking of price range by {kind} and sub-segment against "
            f"the weighted average price of each {kind}"),
        "pricing_forecast": lambda: (
            f"Price forecast by sub-segment, {E}–{Y}, with the factors and "
            f"dynamics shaping the trajectory"),
        "pricing_competitor_strategies": lambda: (
            f"Analysis of how leading players position price: tier, price "
            f"position versus peers and strategy focus"),
        # ── Market dynamics ──
        "dynamics_geography": lambda: (
            f"Analysis of the growth driver, restraint and opportunity "
            f"shaping each {kind} of the {g} market"),
        "driving_factors_supply": lambda: (
            f"Analysis of economic and supply-side factors scored by their "
            f"impact on the {g} market"),
        "driving_factors_demand": lambda: (
            f"Analysis of demand-side factors scored by their impact on the "
            f"{g} market"),
        "market_drivers": lambda: (
            f"Analysis of the key growth drivers of the {g} market and the "
            f"likelihood of their impact across {H0}–{Y}"),
        "market_restraints": lambda: (
            f"Analysis of the key restraints on the {g} market and the "
            f"likelihood of their impact across {H0}–{Y}"),
        "market_opportunities": lambda: (
            f"Analysis of the key opportunities in the {g} market and the "
            f"likelihood of their impact across {H0}–{Y}"),
        "recent_developments": lambda: (
            f"Recent developments and milestones shaping the {g} market, "
            f"by year"),
        "five_forces": lambda: (
            f"Analysis of supplier power, buyer power, competition, "
            f"substitutes and entry barriers shaping the market's "
            f"profitability"),
        # ── Supply chain ──
        "value_chain_map": lambda: (
            f"Mapping of the value chain stages, the stakeholders at each "
            f"stage and the value added along the chain"),
        "operating_margin": lambda: (
            f"Operating margin by region and value chain node, "
            f"highlighting where value concentrates along the chain"),
        "buyers_assessment": lambda: (
            f"Assessment of buyers and end users by volume, average price "
            f"and revenue, highlighting purchasing power"),
        # ── Competition ──
        "industry_structure": lambda: (
            f"Industry structure sizing Tier 1, 2 and 3 players against the "
            f"{g} market total, highlighting concentration"),
        "competitor_advice": lambda: (
            f"Strategic advice for vendors and suppliers on pricing, "
            f"product, geography, channel and competitive risk"),
        "company_value_share": lambda: (
            f"Analysis of value share by company in the {g} market, {H1}, "
            f"highlighting leadership and the long tail"),
        "company_share_by_product": lambda: (
            f"Analysis of company value share within each leading "
            f"{_by(data.SEGMENTS[0]).lower() if data.SEGMENTS else 'segment'}"
            f", highlighting where leadership differs"),
        "company_share_by_region": lambda: (
            f"Company value share within the largest regions, highlighting "
            f"where leadership shifts across geographies"),
        "year_shapers": lambda: (
            f"Ranking of the players that shaped growth, by value share, "
            f"growth role and signature move"),
        "competitor_dashboard": lambda: (
            f"Benchmarking of companies on portfolio, coverage, revenue, "
            f"share, regional presence and strategy"),
        "company_snapshot": lambda: (
            f"Overview of total market, addressable opportunity and share "
            f"of market by end-market, region and product"),
        "company_penetration": lambda: (
            f"Regional penetration and competitive positioning highlighting "
            f"growth potential, win rates and sales efficiency"),
        "company_swot": lambda: (
            f"SWOT assessment of {a0}: strengths, weaknesses, opportunities "
            f"and threats shaping its competitive position"),
        "company_financials": lambda: (
            f"Analysis of {a0}'s financial performance, key ratios and "
            f"pricing, highlighting margin and momentum"),
        "company_strategy": lambda: (
            f"Analysis of {a0}'s strategic priorities and the competitive "
            f"moves that support them"),
    }.get(name)

    try:
        s = text() if text else ""
    except Exception:
        s = ""
    if not s:
        s = _fallback(page, region, country)
    return _tidy(s)


def _fallback(page, region, country) -> str:
    parts = [geo.label(page.get("section") or ""),
             geo.label(page.get("group") or "")]
    scope = country if "country" in page.get("slicers", ()) else region
    if scope:
        parts.append(str(scope))
    return "  ·  ".join(p for p in parts if p)


def _tidy(s: str) -> str:
    s = " ".join(str(s).split())
    if len(s) > _MAX:
        cut = s[:_MAX - 1].rsplit(" ", 1)[0]
        s = cut.rstrip(" ,;:·") + "…"
    return s
