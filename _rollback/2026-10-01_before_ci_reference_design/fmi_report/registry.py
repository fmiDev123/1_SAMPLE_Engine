"""Build the full page list of both PBIX reports.

Each entry: slug, name, report, group, slicers (shared slicer roles) and
a render function whose positional args are the slicer values in order.
"""
from __future__ import annotations

from functools import partial

from . import (data, geo, templates_competitor as tcomp,
               templates_country as tc, templates_dynamics as td,
               templates_main as tm, templates_value_chain as tvc)
from . import templates_cross as tx
from . import templates_industry as tind
from . import templates_ci as tci

MAIN_REPORT = data.MARKET_LABEL                     # e.g. "Matcha Market"
COUNTRY_REPORT = f"{data.MARKET_LABEL} Country"
DYNAMICS_REPORT = f"{data.MARKET_LABEL} Market Dynamics"
COMPETITOR_REPORT = f"{data.MARKET_LABEL} Competitor"

# stable URL prefixes, independent of the configurable market name
_REPORT_KEY = {MAIN_REPORT: "main", COUNTRY_REPORT: "country",
               DYNAMICS_REPORT: "dynamics", COMPETITOR_REPORT: "competitor"}

def _slug(*parts) -> str:
    out = "-".join(str(p) for p in parts).lower()
    out = out.replace("—", " ").replace("–", " ")   # keep slugs ASCII
    for ch in "&/()$,:;":
        out = out.replace(ch, "")
    return "-".join(out.split()).replace("--", "-")

PAGES: list[dict] = []

# ── Storyline sections ───────────────────────────────────────────────
# The left menu (and the Agenda page) is organised into these top-level
# SECTIONS rather than into the four underlying PBIX "reports".  Each is
# prefixed with the market name, so the whole storyline adapts to any
# workbook.  `PAGES` is registered in this same order, so the deck flows
# exactly as the menu reads.
_M = data.MARKET
SEC_ASSESS = f"{_M} Market Assessment"
SEC_GLOBAL = f"{_M} Global Market Assessment"
SEC_COUNTRY = f"{_M} Country Assessment"
SEC_PRICING = f"{_M} Pricing Analysis"
SEC_DYNAMICS = f"{_M} Market Dynamics and Trends"
SEC_SUPPLY = f"{_M} Supply Chain Analysis"
SEC_COMPETE = f"{_M} Competition Assessment"

# sub-sections nest one level deeper (section ▸ subsection ▸ group ▸ page)
SUB_BY_REGION = "Market Size-By Regions"
SUB_BY_SEGMENT = "Market Size-By Segment"

# every page of these reports belongs to one section, so it need not be
# repeated on each _add(); the main report spans several and passes its own
_DEFAULT_SECTION = {
    MAIN_REPORT: SEC_GLOBAL,
    COUNTRY_REPORT: SEC_COUNTRY,
    DYNAMICS_REPORT: SEC_DYNAMICS,
    COMPETITOR_REPORT: SEC_COMPETE,
}

# menu / deck order of the sections
SECTION_ORDER = [SEC_ASSESS, SEC_GLOBAL, SEC_COUNTRY, SEC_PRICING,
                 SEC_DYNAMICS, SEC_SUPPLY, SEC_COMPETE]


_UNIT_WORDS = ("Volume", "Value")


def _page_has_its_data(name, render, section) -> bool:
    """False for a page this workbook cannot fill.

    A study is normally sized in BOTH Volume and Value.  A value-only workbook
    carries one of them, and then every Volume page has nothing to draw and
    every price nothing to divide by - price is Value / Volume, derived, never
    read - so `metrics.region_price_variation` raised a KeyError and took the
    whole export down.  The page is left out instead of emitted empty.

    The unit a page needs is not annotated anywhere: it is passed to the
    template, so it is read back off the `partial` the registry was given,
    and off the page's own NAME for the pages that carry it there
    ("Volume Market Summary", "Market Volume & Value Analysis").  A workbook
    with both units keeps every page, so nothing moves for the usual case.
    """
    wanted = {u for u in _UNIT_WORDS if u in str(name)}
    for a in getattr(render, "args", ()) or ():
        if a in _UNIT_WORDS:
            wanted.add(a)
    for a in (getattr(render, "keywords", None) or {}).values():
        if a in _UNIT_WORDS:
            wanted.add(a)
    if any(not data.has_unit(u) for u in wanted):
        return False
    # Anything about PRICE needs both units, wherever it sits: the pricing
    # storyline is a whole section, but "Pricing & Price Index" sits in the
    # Global Overview group, so the section alone does not catch it.
    if not data.has_price() and (section == SEC_PRICING
                                 or "Pric" in str(name)):
        return False
    return True


def _add(report, group, name, render, slicers=(), flow=False,
         section=None, subsection=None):
    if not _page_has_its_data(name, render, section or
                              _DEFAULT_SECTION[report]):
        return
    # ``flow=True`` marks a page whose render() returns a layout spec
    # ({"fixed": [...], "tables": [...]}, see layout.py) instead of
    # pre-positioned parts - the engine content-sizes and paginates it.
    # ``section``/``subsection`` place the page in the storyline menu tree.
    PAGES.append({
        "slug": _slug(_REPORT_KEY[report], group, name),
        "name": name, "report": report, "group": group,
        "render": render, "slicers": list(slicers), "flow": flow,
        "section": section or _DEFAULT_SECTION[report],
        "subsection": subsection,
    })

# ═══════════════════════ Matcha Market.pbix ══════════════════════════
# Report opens on a dynamic Contents page, then the Market Signals editorial,
# then the two Market Taxonomy flow charts - all ahead of the analytical pages.
_add(MAIN_REPORT, "Agenda", "Agenda", tm.contents_page,
     section=SEC_ASSESS)
_add(MAIN_REPORT, "Market Signals", "Market Signals",
     tm.market_signals, flow=True, section=SEC_ASSESS)
_add(MAIN_REPORT, "Market Taxonomy", "Geographic Coverage",
     tm.taxonomy_regions, flow=True, section=SEC_ASSESS)
_add(MAIN_REPORT, "Market Taxonomy", "Market Segments",
     tm.taxonomy_segments, flow=True, section=SEC_ASSESS)
_add(MAIN_REPORT, "Market Taxonomy", "Company Coverage",
     tm.company_coverage, flow=True, section=SEC_ASSESS)
# "Overview" group renamed to a market-specific assessment title (data-driven
# via data.MARKET, so it adapts to any workbook - e.g. "Matcha- Global Market
# Assessment"). Order: Volume Summary → Pricing & Price Index → Value Summary →
# Market Volume & Value Analysis → Country Outlook → executive-summary series
# the section heading already carries the market name + "Global Market
# Assessment", so this group just needs to distinguish itself from its two
# siblings ("Market Size-By Regions" / "Market Size-By Segment")
_OVERVIEW_GRP = "Global Overview"
_add(MAIN_REPORT, _OVERVIEW_GRP, "Volume Market Summary",
     partial(tm.market_summary, "Volume"), ["region"], flow=True)
_add(MAIN_REPORT, _OVERVIEW_GRP, "Pricing & Price Index",
     tm.pricing_analysis, ["region"], flow=True)
_add(MAIN_REPORT, _OVERVIEW_GRP, "Value Market Summary",
     partial(tm.market_summary, "Value"), ["region"], flow=True)
_add(MAIN_REPORT, _OVERVIEW_GRP, "Market Volume & Value Analysis",
     tm.vol_val_analysis, ["region"], flow=True)
# (Volume/Value CAGR Analysis 1/2 & 2/2 pages removed per request)
_add(MAIN_REPORT, _OVERVIEW_GRP, "Country Outlook",
     partial(tm.summary, "Value"))
# Executive-summary series lives under the Overview section, right after the
# Country Outlook page - shown as 1/4 ... 4/4 in each page/slide heading.
_add(MAIN_REPORT, _OVERVIEW_GRP, "Market Overview (1/4)",
     tm.summary_overview, flow=True)
_add(MAIN_REPORT, _OVERVIEW_GRP, "Supply Side Trends",
     tm.summary_supply, flow=True)
_add(MAIN_REPORT, _OVERVIEW_GRP,
     "Analysis and Recommendation",
     tm.summary_analysis, flow=True)
_add(MAIN_REPORT, _OVERVIEW_GRP, "Scenario Outlook (4/4)",
     tm.summary_scenario, flow=True)

# ── Global Market Assessment ▸ "Market Size-By Regions" ──────────────
# (Pricing / Supply Chain used to sit here; they now follow Country and
#  Dynamics respectively, so the storyline reads in one pass.)
# menu label spelled out in full so it reads identically to the page heading
# (tm.region_start renders "{region} {MARKET} - Regional Outlook"; the menu is
# static, so it carries the default root geography, "Global")
_add(MAIN_REPORT, "By Region", f"Global {data.MARKET} – Regional Outlook",
     tm.region_start, ["region"], flow=True, subsection=SUB_BY_REGION)
_add(MAIN_REPORT, "By Region", "Market Analysis By Region (1/2)",
     partial(tm.region_donuts, "Volume"), ["region"], flow=True,
     subsection=SUB_BY_REGION)
_add(MAIN_REPORT, "By Region", "Market Analysis By Region (2/2)",
     partial(tm.region_donuts, "Value"), ["region"], flow=True,
     subsection=SUB_BY_REGION)
_add(MAIN_REPORT, "By Region", "Volume Share Analysis By Region",
     partial(tm.region_share_tables, "Volume"), ["region"], flow=True,
     subsection=SUB_BY_REGION)
_add(MAIN_REPORT, "By Region", "Value Share Analysis By Region",
     partial(tm.region_share_tables, "Value"), ["region"], flow=True,
     subsection=SUB_BY_REGION)
_add(MAIN_REPORT, "By Region", "Market Trends By Region",
     tm.region_ribbons, ["region"], flow=True, subsection=SUB_BY_REGION)
_add(MAIN_REPORT, "By Region", "Top Countries",
     tm.market_hierarchy, ["region"], flow=True, subsection=SUB_BY_REGION)
# ("End of Segment" closing card removed per request - see the segment loop
#  below, where the per-segment twin was dropped too.  tm.region_end /
#  tm.seg_end and the "endcards" block renderers are kept but unused.)

for seg in data.SEGMENTS:
    # each segment section opens with a creative segment-outlook front page,
    # then the merged Market Structure & Attractiveness page.  (The former
    # stand-alone "Market Structure" mekko page was merged INTO the old
    # "BPS & Market Attractiveness" page: its horizontal composition-bar chart
    # took the place of that page's sparse BPS table - see tm.seg_bps.)
    # menu label spelled out in full so it matches the page heading exactly
    # (tm.seg_start renders "{By Seg} Segment Outlook")
    _seg_sub = dict(subsection=SUB_BY_SEGMENT)   # Global MA ▸ By Segment
    _add(MAIN_REPORT, seg, f"{tm._by(seg)} Segment Outlook",
         partial(tm.seg_start, seg), ["region"], flow=True, **_seg_sub)
    _add(MAIN_REPORT, seg, "Market Structure & Attractiveness",
         partial(tm.seg_bps, seg, "Value"), ["region"], flow=True, **_seg_sub)
    _add(MAIN_REPORT, seg, "Absolute $ Analysis",
         partial(tm.seg_absolute, seg), ["region"], flow=True, **_seg_sub)
    _add(MAIN_REPORT, seg, "Market Volume Analysis",
         partial(tm.seg_size_tables, seg, "Volume"), ["region"], flow=True,
         **_seg_sub)
    _add(MAIN_REPORT, seg, "Market Value Analysis",
         partial(tm.seg_size_tables, seg, "Value"), ["region"], flow=True,
         **_seg_sub)
    _add(MAIN_REPORT, seg, "Value & Share Shift",
         partial(tm.seg_ribbons, seg), ["region"], flow=True, **_seg_sub)
    _add(MAIN_REPORT, seg, "YoY Analysis Volume",
         partial(tm.seg_yoy, seg, "Volume"), ["region"], flow=True, **_seg_sub)
    _add(MAIN_REPORT, seg, "YoY Analysis Value",
         partial(tm.seg_yoy, seg, "Value"), ["region"], flow=True, **_seg_sub)
    # ("End of Segment" closing card removed per request)

# ═════════════ Cross-sectional Assessment (2026-09-29) ═══════════════
# Three pages for each cross-tab segment ('A x B'), after every segment's
# own pages.  The deck shows these for a cross-tab in place of the
# per-segment pages above (see export_native); the dashboard keeps both.
for seg in [s for s in data.SEGMENTS if data.is_cross(s)]:
    _seg_sub = dict(subsection=SUB_BY_SEGMENT)
    _add(MAIN_REPORT, seg, tx.PAGE_NAMES[0],
         partial(tx.cross_snapshot, seg), ["region"], flow=True, **_seg_sub)
    _add(MAIN_REPORT, seg, tx.PAGE_NAMES[1],
         partial(tx.cross_share, seg), ["region"], flow=True, **_seg_sub)
    _add(MAIN_REPORT, seg, tx.PAGE_NAMES[2],
         partial(tx.cross_opportunity, seg), ["region"], flow=True,
         **_seg_sub)

# ═══════════════════ Matcha Market- Country.pbix ═════════════════════
_add(COUNTRY_REPORT, "Overview", "Volume Overview",
     partial(tc.overview, "Volume"), ["country"], flow=True)
_add(COUNTRY_REPORT, "Overview", "Value Overview",
     partial(tc.overview, "Value"), ["country"], flow=True)
_add(COUNTRY_REPORT, "Overview", "Key Segments (Value)",
     tc.seg_overview_value, ["country"], flow=True)

for seg in data.SEGMENTS:
    _add(COUNTRY_REPORT, seg, "Segmental",
         partial(tc.segmental, seg), ["country"], flow=True)
    _add(COUNTRY_REPORT, seg, "Table (Volume & Value)",
         partial(tc.tables, seg), ["country"], flow=True)
    _add(COUNTRY_REPORT, seg, "YoY",
         partial(tc.yoy, seg), ["country"], flow=True)

# ═════════════════ Pricing Analysis (after Country) ══════════════════
# The (now 10-topic) Pricing Assessment sequence, ordered per the reference
# deck (price trends → positioning → key findings → drivers → factor table →
# assessment → benchmarking → forecast → competitor strategies).  The former
# "Factors Influencing Pricing" opener was merged into "Factors Impacting
# Pricing" (same ten pricing_drivers.py factors, was shown twice) - see
# tm.pricing_factors_table.
_add(MAIN_REPORT, "Pricing", "Product Pricing & Regional Price Index Trends",
     tm.pricing_index_trends, flow=True, section=SEC_PRICING)
_add(MAIN_REPORT, "Pricing", "Global Pricing Assessment",
     tm.pricing_positioning, flow=True, section=SEC_PRICING)
_add(MAIN_REPORT, "Pricing", "Price Positioning & Key Findings",
     tm.pricing_key_findings, flow=True, section=SEC_PRICING)
_add(MAIN_REPORT, "Pricing", "Key Pricing Drivers",
     tm.pricing_drivers, ["region"], flow=True, section=SEC_PRICING)
_add(MAIN_REPORT, "Pricing", "Factors Impacting Pricing",
     tm.pricing_factors_table, flow=True, section=SEC_PRICING)
_add(MAIN_REPORT, "Pricing", "Regional Price Assessment",
     tm.pricing_assessment, flow=True, section=SEC_PRICING)
_add(MAIN_REPORT, "Pricing", "Global Pricing Benchmarking",
     tm.pricing_benchmarking, flow=True, section=SEC_PRICING)
_add(MAIN_REPORT, "Pricing", "Price Forecast & Dynamics",
     tm.pricing_forecast, flow=True, section=SEC_PRICING)
_add(MAIN_REPORT, "Pricing", "Competitor Pricing Strategies",
     tm.pricing_competitor_strategies, flow=True, section=SEC_PRICING)

# ═════════════ Market Dynamics (immediately after Pricing) ═══════════
# A self-contained dynamics section modelled on the reference deck: the
# geography dynamics matrix, the two driving-factor radars, the Drivers /
# Restraints / Opportunities write-ups, the milestone table and Porter's Five
# Forces.  Every page is region-sliced and fully derived from the workbook.
_DYN_FACTORS = "Global Driving Factors"
_DYN_FORCES = "Market Dynamics & Forces"

_add(DYNAMICS_REPORT, _DYN_FACTORS, "Market Dynamics by Geography",
     td.dynamics_geography, ["region"], flow=True)
_add(DYNAMICS_REPORT, _DYN_FACTORS, "Driving Factors Supply Side (1/2)",
     td.driving_factors_supply, ["region"], flow=True)
_add(DYNAMICS_REPORT, _DYN_FACTORS, "Driving Factors Demand Side (2/2)",
     td.driving_factors_demand, ["region"], flow=True)
_add(DYNAMICS_REPORT, _DYN_FORCES, "Market Drivers",
     td.market_drivers, ["region"], flow=True)
_add(DYNAMICS_REPORT, _DYN_FORCES, "Market Restraints",
     td.market_restraints, ["region"], flow=True)
_add(DYNAMICS_REPORT, _DYN_FORCES, "Market Opportunities",
     td.market_opportunities, ["region"], flow=True)
_add(DYNAMICS_REPORT, _DYN_FORCES, "Recent Developments & Milestones",
     td.recent_developments, ["region"], flow=True)
_add(DYNAMICS_REPORT, _DYN_FORCES, "Five Forces Analysis",
     td.five_forces, ["region"], flow=True)

# ═══════════ Supply Chain Analysis (after Market Dynamics) ═══════════
# (formerly the "Value Chain" group - same three pages, renamed to match the
#  storyline; templates still live in templates_value_chain.py)
_SUPPLY_GRP = "Supply Chain Analysis"
_add(MAIN_REPORT, _SUPPLY_GRP, "Value Chain & Stakeholder Mapping",
     tvc.value_chain_map, flow=True, section=SEC_SUPPLY)
_add(MAIN_REPORT, _SUPPLY_GRP, "Operating Margin Analysis",
     tvc.operating_margin, flow=True, section=SEC_SUPPLY)
_add(MAIN_REPORT, _SUPPLY_GRP, "Buyers Assessment",
     tvc.buyers_assessment, flow=True, section=SEC_SUPPLY)

# ═══════════════════ Competition Assessment (last) ═══════════════════
# A competitor report parallel to the Country report - the last top-level
# section.  Its slides come from templates_competitor / competitors.py.
# Competitive Landscape opens the competition section (2026-09-29), with the
# ten competitive-intelligence pages first (2026-09-30)
for _ci_name, _ci_fn in tci.PAGES:
    _add(COMPETITOR_REPORT, "Competitive Landscape", _ci_name, _ci_fn,
         flow=True)
_add(COMPETITOR_REPORT, "Competitive Landscape",
     "Industry Structure Landscape", tind.industry_landscape, flow=True)
# the Competitive Dashboard follows it, just before Market Share by Leading
# Players (moved from that section, 2026-09-30)
_add(COMPETITOR_REPORT, "Competitive Landscape", "Competitive Dashboard",
     tcomp.competitor_dashboard, flow=True)

_SHARE_GRP = "Market Share by Leading Players"
# ('Industry Structure' was dropped from the deck on 2026-09-29)
_add(COMPETITOR_REPORT, _SHARE_GRP, "Advice for Vendors & Suppliers",
     tcomp.competitor_advice, flow=True)
_add(COMPETITOR_REPORT, _SHARE_GRP, "Company Value Share",
     tcomp.company_value_share, ["region"], flow=True)
_add(COMPETITOR_REPORT, _SHARE_GRP, "Company Share by Product Type",
     tcomp.company_share_by_product, flow=True)
_add(COMPETITOR_REPORT, _SHARE_GRP, "Company Share by Region",
     tcomp.company_share_by_region, flow=True)
_add(COMPETITOR_REPORT, _SHARE_GRP, "Who Shaped Growth",
     tcomp.year_shapers, flow=True)

_PROFILE_GRP = "Company Profiles & Strategic Benchmarking"
for _co in tcomp.C.profile_companies():
    _add(COMPETITOR_REPORT, _PROFILE_GRP, f"{_co} Snapshot",
         partial(tcomp.company_snapshot, _co), flow=True)
    _add(COMPETITOR_REPORT, _PROFILE_GRP, f"{_co} Penetration & Win",
         partial(tcomp.company_penetration, _co), flow=True)
    _add(COMPETITOR_REPORT, _PROFILE_GRP, f"{_co} SWOT Analysis",
         partial(tcomp.company_swot, _co), flow=True)
    _add(COMPETITOR_REPORT, _PROFILE_GRP, f"{_co} Key Financials",
         partial(tcomp.company_financials, _co), flow=True)
    _add(COMPETITOR_REPORT, _PROFILE_GRP, f"{_co} Strategy Analysis",
         partial(tcomp.company_strategy, _co), flow=True)

# sanity: slugs must be unique
_slugs = [p["slug"] for p in PAGES]
assert len(_slugs) == len(set(_slugs)), "duplicate page slugs"

HOME_SLUG = PAGES[0]["slug"]
BY_SLUG = {p["slug"]: p for p in PAGES}


def section_index() -> list[dict]:
    """Ordered list of report sections for the Contents page / PPT menu.

    One entry per (report, group) in registry order - excluding the Contents
    group itself - pointing at the FIRST page of that group.  Fully derived
    from PAGES, so it adapts to whatever sections the loaded workbook yields.
    Each entry: {report, group, slug, name}.
    """
    out, seen = [], set()
    for p in PAGES:
        if p["group"] == "Contents":
            continue
        key = (p["report"], p["group"])
        if key in seen:
            continue
        seen.add(key)
        out.append({"report": p["report"], "group": p["group"],
                    "slug": p["slug"], "name": p["name"]})
    return out


def contents_buttons() -> list[dict]:
    """Ordered, flat list of Contents navigation buttons - one per report
    section, generated from the page registry (the single source of truth).

    A section is a (report, group); the Market Taxonomy group is expanded into
    one button per page (Regions / Segments) since those are distinct sections.
    Order: Overview, Market Signals, Market Taxonomy pages, then the remaining
    main sections in page order, then the country sections (suffixed
    ``(Country)``).  Structural groups only drive ordering - segment/market
    names come straight from the data, so it adapts to any workbook.

    Each button: {label, slug}.
    """
    buttons = []
    for section in SECTION_ORDER:
        spages = [p for p in PAGES if p.get("section") == section
                  and p["group"] != "Agenda"]
        if not spages:
            continue
        if section == SEC_ASSESS:
            # the opening section is short - expand it to its own pages so the
            # Agenda lists Market Signals / Geographic Coverage / ... directly
            buttons += [{"label": geo.label(p["name"]), "slug": p["slug"]}
                        for p in spages]
            continue
        buttons.append({"label": geo.label(section),
                        "slug": spages[0]["slug"]})
    return buttons


def contents_outline() -> list[dict]:
    """Grouped, numbered outline for the formal Table of Contents.

    Framing groups (Overview / Market Signals / Market Taxonomy ...) each become
    their own numbered section; the analytical groups (per-segment + By Region)
    are bundled under a 'Market Analysis' section, and the whole country report
    under 'Country Analysis'.  Sub-sections carry their own page links.  Groups
    are identified structurally (segment groups via data.SEGMENTS), so the
    outline adapts to any workbook - nothing market-specific is hardcoded.

    Each section: {no, title, slug (first page), subs: [{label, slug}]}.
    """
    from . import data
    seg_groups = set(data.SEGMENTS)
    main_groups, cty_groups, comp_groups, pages_by = [], [], [], {}
    dyn_groups = []
    for p in PAGES:
        if p["group"] == "Contents":
            continue
        pages_by.setdefault((p["report"], p["group"]), []).append(p)
        if p["report"] == MAIN_REPORT and p["group"] not in main_groups:
            main_groups.append(p["group"])
        elif p["report"] == COUNTRY_REPORT and p["group"] not in cty_groups:
            cty_groups.append(p["group"])
        elif p["report"] == DYNAMICS_REPORT and p["group"] not in dyn_groups:
            dyn_groups.append(p["group"])
        elif p["report"] == COMPETITOR_REPORT and p["group"] not in comp_groups:
            comp_groups.append(p["group"])

    def short(name, group):                       # 'Market Taxonomy - Regions'
        for sep in (" ", " – ", " - ", ": "):   # -> 'Regions'
            if name.startswith(group + sep):
                return name[len(group + sep):]
        return name

    sections, analytical = [], []
    for g in main_groups:
        if g in seg_groups or g == "By Region":   # analytical -> bundle later
            analytical.append(g)
            continue
        pgs = pages_by[(MAIN_REPORT, g)]
        subs = ([{"label": short(p["name"], g), "slug": p["slug"]}
                 for p in pgs] if 2 <= len(pgs) <= 4 else [])
        sections.append({"title": g, "slug": pgs[0]["slug"], "subs": subs})
    if analytical:
        subs = [{"label": g, "slug": pages_by[(MAIN_REPORT, g)][0]["slug"]}
                for g in analytical]
        sections.append({"title": "Market Analysis", "slug": subs[0]["slug"],
                         "subs": subs})
    if cty_groups:
        subs = [{"label": g, "slug": pages_by[(COUNTRY_REPORT, g)][0]["slug"]}
                for g in cty_groups]
        sections.append({"title": "Country Analysis", "slug": subs[0]["slug"],
                         "subs": subs})
    if dyn_groups:
        subs = [{"label": g, "slug": pages_by[(DYNAMICS_REPORT, g)][0]["slug"]}
                for g in dyn_groups]
        sections.append({"title": "Market Dynamics", "slug": subs[0]["slug"],
                         "subs": subs})
    if comp_groups:
        subs = [{"label": g, "slug": pages_by[(COMPETITOR_REPORT, g)][0]["slug"]}
                for g in comp_groups]
        sections.append({"title": "Competition Assessment",
                         "slug": subs[0]["slug"], "subs": subs})
    # the outline is rebuilt on every call, so re-pointing the labels at the
    # active geography here is safe - and it covers every branch above at once
    for i, s in enumerate(sections, 1):
        s["no"] = f"{i:02d}"
        s["title"] = geo.label(s["title"])
        for sub in s["subs"]:
            sub["label"] = geo.label(sub["label"])
    return sections
