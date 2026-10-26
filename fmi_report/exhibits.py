"""The List of Figures and the List of Tables of the full study.

The two lists describe the report as it is SOLD, not only the deck that
previews it: a fixed number of exhibits (200 figures, 200 tables by default)
spread over the report's analysis blocks by a weightage, and titled from the
workbook's own geographies, segments, units and years.

    Analysis block                           Figures   Tables
    Executive Summary and Market Overview       10        8
    Global Market Size and Forecast             12       15
    Segment Analysis                            35       45
    Cross-Segment Analysis                      15       20
    Regional Analysis                           30       28
    Country-Level Analysis                      40       38
    Pricing, Value Chain and Trade              15       18
    Competitive Landscape                       25       20
    Trends, Opportunities and Scenarios         12        5
    Client-Specific USP Analysis                 6        3
    Total                                      200      200

The weightage is a STARTING point.  `_adapted_weights` bends it to the study
in hand: a seven-segment study needs more segment exhibits than a
three-segment one, a study with one segment has no cross-segment exhibits,
and a single-country study moves its regional exhibits to the country and
segment blocks.  After that the weights are scaled back to the totals, and a
block whose workbook cannot fill its share (no volume, few countries) hands
the rest to the blocks that can.

To change the split for every market, edit `BLOCKS` below.  To change it for
ONE build, point the REPORT_EXHIBIT_WEIGHTS environment variable at a JSON
file such as

    {"totals": {"figures": 180, "tables": 220},
     "weights": {"segment": [40, 50], "country": [30, 30]}}

where each weight is [figures, tables] and the keys are the block keys below.

Title rules.  Every title states its geography, indicator, segmentation
level, unit and period:

    Figure [No.]: [Geography] [Market/Segment] [Metric], [Unit], [Period]
    Table [No.]:  [Geography] [Market/Segment] [Measure] by [Dimension],
                  [Unit], [Period]

Country coverage follows the priority rule: priority countries (the largest
by forecast-year value) get 8-12 exhibits each, secondary countries 2-4, so
minor countries never take 15-20 exhibits apiece.
"""
from __future__ import annotations

import json
import math
import os
import re
from pathlib import Path

from . import config, data, geo

# ── the weightage ────────────────────────────────────────────────────
# (key, block title, figures, tables)
BLOCKS = [
    ("summary",     "Executive Summary and Market Overview", 10, 8),
    ("size",        "Global Market Size and Forecast",       12, 15),
    ("segment",     "Segment Analysis",                      35, 45),
    ("cross",       "Cross-Segment Analysis",                15, 20),
    ("regional",    "Regional Analysis",                     30, 28),
    ("country",     "Country-Level Analysis",                40, 38),
    ("pricing",     "Pricing, Value Chain and Trade",        15, 18),
    ("competitive", "Competitive Landscape",                 25, 20),
    ("trends",      "Trends, Opportunities and Scenarios",   12, 5),
    ("usp",         "Client-Specific USP Analysis",          6,  3),
]
TOTAL_FIGURES = 200
TOTAL_TABLES = 200

# country coverage: exhibits per country, per kind (figures or tables).
# Leaned to the Croissants 200/200 reference (2026-10-23), which gives
# every country the same 2 combined exhibits rather than drilling 8-12
# deep into a few priority ones - so coverage reaches every country
# first, and only a handful of top ones get the extra two.
PRIORITY_MIN, PRIORITY_MAX = 3, 4
SECONDARY_MIN, SECONDARY_MAX = 2, 2

_DASH = "–"                           # en dash, as in "2021–2036"


# ── configuration ────────────────────────────────────────────────────
def _overrides() -> dict:
    path = os.environ.get("REPORT_EXHIBIT_WEIGHTS")
    if not path:
        return {}
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return {}


def _base_weights() -> tuple[dict, dict, int, int]:
    ov = _overrides()
    figs = {k: float(f) for k, _, f, _ in BLOCKS}
    tabs = {k: float(t) for k, _, _, t in BLOCKS}
    for k, pair in (ov.get("weights") or {}).items():
        if k in figs and isinstance(pair, (list, tuple)) and len(pair) == 2:
            figs[k], tabs[k] = float(pair[0]), float(pair[1])
    tot = ov.get("totals") or {}
    return (figs, tabs, int(tot.get("figures", TOTAL_FIGURES)),
            int(tot.get("tables", TOTAL_TABLES)))


def _short_market() -> str:
    """The market as titles name it.  A name that carries its own
    abbreviation - 'Primary Sclerosing Cholangitis (PSC)' - is titled by the
    abbreviation, so the unit and period at the END of a title still fit on
    its one row of the list."""
    name = str(data.MARKET).strip()
    m = re.search(r"\(([A-Za-z0-9&\-]{2,12})\)\s*$", name)
    return f"{m.group(1) if m else name} Market"


# ── the study in hand ────────────────────────────────────────────────
class _Study:
    """Everything the titles are written from, read once from the loaded
    workbook for the geography the deck is built for."""

    def __init__(self, root: str):
        self.root = root
        self.market = _short_market()
        self.axes = [a for a in data.taxonomy_axes()
                     if data.SUBSEGMENTS.get(a)]
        self.subs = {a: list(data.SUBSEGMENTS.get(a, [])) for a in self.axes}
        self.volume = data.has_unit("Volume")
        self.value = data.has_unit("Value") or not self.volume
        self.price = data.has_price()
        self.vu = config.UNIT_LABELS.get("Value", "")
        self.qu = config.UNIT_LABELS.get("Volume", "")
        self.pu = config.price_unit()
        self.cur = config.currency()
        y = config
        self.h0, self.h1 = y.HIST_START, y.HIST_END
        self.e, self.z = y.EST_YEAR, y.END_YEAR
        self.full = f"{self.h0}{_DASH}{self.z}"
        self.hist = f"{self.h0}{_DASH}{self.h1}"
        self.fore = f"{self.e}{_DASH}{self.z}"
        self.ez = f"{self.e} and {self.z}"
        # geography tree below the root
        kids = list(data.children(root))
        if root == "Global":
            self.regions = kids
            self.countries = {r: list(data.children(r)) for r in kids}
        else:
            # a regional deck: its countries are the "regions" level's
            # children; a single-country study has neither
            self.regions = []
            self.countries = {root: kids} if kids else {}
        self.country_rank = self._rank_countries()
        try:
            from . import competitors
            self.companies = [str(c) for c in competitors.companies()]
        except Exception:
            self.companies = []

    def _rank_countries(self) -> list[tuple[str, str]]:
        unit = "Value" if self.value else "Volume"
        rows = []
        for reg, cs in self.countries.items():
            for c in cs:
                if re.match(r"\s*rest of\b", str(c), re.I):
                    continue                 # an aggregate, not a country
                try:
                    v = float(data.total_series(c, unit)[self.z])
                except Exception:
                    v = 0.0
                rows.append((v if v == v else 0.0, reg, c))
        rows.sort(key=lambda r: -r[0])
        return [(reg, c) for _, reg, c in rows]

    # naming helpers
    def m(self, where=None) -> str:
        return f"{where or self.root} {self.market}"

    @staticmethod
    def dim(axis: str) -> str:
        """'By Product Type' -> 'Product Type'."""
        return re.sub(r"^\s*by\s+", "", str(axis), flags=re.I).strip()

    def child_word(self) -> str:
        return "Region" if self.regions else "Country"


# ── title pools, one per block and kind ──────────────────────────────
# Each pool lists its titles MOST important first, so a block given fewer
# exhibits than its pool keeps the headline ones.  Pools that cover many
# members (segments, regions) interleave them, so a short allocation still
# touches every member rather than exhausting the first.

def _rr(*lists):
    """Round-robin merge: a1, b1, c1, a2, b2, ..."""
    out = []
    for i in range(max((len(x) for x in lists), default=0)):
        for x in lists:
            if i < len(x):
                out.append(x[i])
    return out


def _summary(s: _Study, kind):
    M = s.m()
    if kind == "figure":
        t = [f"{M} Value and YoY Growth, {s.vu} and %, {s.full}",
             f"{M} Value Snapshot, {s.vu}, {s.ez}",
             f"{M} Taxonomy by Segment, {s.e}"]
        t += [f"{M} Share by {s.dim(a)}, %, {s.e}" for a in s.axes]
        if s.regions:
            t.append(f"{M} Share by Region, %, {s.e}")
        t += [f"{M} Incremental $ Opportunity, {s.vu}, {s.fore}",
              f"{M} Key Demand Drivers Impact Assessment, {s.fore}",
              f"{M} Key Restraints Impact Assessment, {s.fore}",
              f"{M} Value Forecast by Scenario, {s.vu}, {s.fore}",
              f"{M} Supply-Side Trends Assessment, {s.e}"]
        return t
    t = [f"{M} Snapshot: Value, Volume and CAGR, {s.vu}, {s.h1}{_DASH}{s.z}",
         f"{M} Taxonomy by Segment, {s.e}",
         f"{M} Scope: Inclusions and Exclusions by Segment, {s.e}",
         f"{M} Key Demand Drivers and Impact Rating, {s.fore}",
         f"{M} Key Restraints and Impact Rating, {s.fore}",
         f"{M} Analyst Recommendations by Segment, {s.fore}",
         f"{M} Key Market Signals and Data Points, {s.e}",
         f"{M} Research Methodology and Data Sources, {s.e}"]
    return t


def _size(s: _Study, kind):
    M = s.m()
    if kind == "figure":
        t = [f"{M} Historical Value, {s.vu}, {s.hist}",
             f"{M} Forecast Value, {s.vu}, {s.fore}",
             f"{M} YoY Growth, %, {s.full}"]
        if s.volume:
            t += [f"{M} Volume and YoY Growth, {s.qu} and %, {s.full}",
                  f"{M} Volume vs Value CAGR Comparison, %, {s.fore}"]
        if s.price:
            t.append(f"{M} Average Selling Price, {s.pu}, {s.full}")
        t += [f"{M} Absolute $ Opportunity, {s.vu}, {s.fore}",
              f"{M} Incremental Opportunity by Year, {s.vu}, {s.fore}",
              f"{M} Historical vs Forecast CAGR, %, {s.full}",
              f"{M} Value by Scenario (Optimistic, Likely, Pessimistic), "
              f"{s.vu}, {s.fore}",
              f"{M} Size Evolution Index, {s.h1}=100, {s.full}",
              f"{M} Cumulative Value Addition, {s.vu}, {s.fore}",
              f"{M} Five-Year Absolute $ Opportunity, {s.vu}, "
              f"{s.e}{_DASH}{min(s.e + 5, s.z)}"]
        return t
    t = [f"{M} Value and YoY Growth, {s.vu} and %, {s.full}",
         f"{M} Historical Value Analysis, {s.vu}, {s.hist}",
         f"{M} Forecast Value Analysis, {s.vu}, {s.fore}"]
    if s.volume:
        t += [f"{M} Volume and YoY Growth, {s.qu} and %, {s.full}",
              f"{M} Volume Forecast, {s.qu}, {s.fore}"]
    if s.price:
        t += [f"{M} Average Selling Price, {s.pu}, {s.full}",
              f"{M} Average Selling Price YoY Change, %, {s.fore}"]
    t += [f"{M} CAGR and Absolute $ Opportunity, %, {s.vu}, {s.fore}",
          f"{M} Absolute $ Opportunity by Year, {s.vu}, {s.fore}",
          f"{M} Value by Scenario (Optimistic, Likely, Pessimistic), "
          f"{s.vu}, {s.fore}",
          f"{M} Scenario CAGR Comparison, %, {s.fore}",
          f"{M} Historical vs Forecast Growth Comparison, %, {s.full}",
          f"{M} Five-Year Block Value Summary, {s.vu}, {s.full}",
          f"{M} Size Evolution Index, {s.h1}=100, {s.full}",
          f"{M} Key Forecast Assumptions, {s.fore}"]
    return t


def _segment(s: _Study, kind):
    """A lean, fixed set of metrics per axis (client reference, 2026-10-23:
    the Croissants 200/200 exhibit list) - no per-sub-segment breakout, no
    scenario line, no ASP: just the headline metrics every axis gets."""
    M = s.m()
    per_axis = []
    for a in s.axes:
        d = s.dim(a)
        if kind == "figure":
            t = [f"{M} Share by {d}, %, {s.e}",
                 f"{M} Growth by {d}, CAGR %, {s.fore}",
                 f"{M} Absolute $ Opportunity by {d}, {s.vu}, {s.fore}",
                 f"{M} Attractiveness by {d}, {s.fore}"]
        else:
            t = [f"{M} Value by {d}, {s.vu}, {s.full}"]
            if s.volume:
                t.append(f"{M} Volume by {d}, {s.qu}, {s.full}")
            t += [f"{M} Share by {d}, %, {s.h1}, {s.e} and {s.z}",
                  f"{M} YoY Growth by {d}, %, {s.fore}",
                  f"{M} CAGR and Absolute $ Opportunity by {d}, %, "
                  f"{s.vu}, {s.fore}",
                  f"{M} Attractiveness and BPS Change by {d}, {s.fore}"]
        per_axis.append(t)
    return _rr(*per_axis)


def _cross(s: _Study, kind):
    """Adjacent, non-overlapping axis pairs rather than every combination
    (client reference, 2026-10-23: Croissants pairs its six axes into
    three, not all fifteen combinations)."""
    M = s.m()
    pairs = list(zip(s.axes[0::2], s.axes[1::2]))
    per = []
    for a, b in pairs:
        da, db = s.dim(a), s.dim(b)
        if kind == "figure":
            per.append([f"{da} × {db} Opportunity Matrix, {s.fore}",
                        f"{M} Value by {da} and {db}, {s.vu}, {s.z}",
                        f"{M} {db} Mix within {da}, %, {s.z}"])
        else:
            per.append([f"{M} Value by {da} and {db}, {s.vu}, {s.fore}",
                        f"{M} CAGR by {da} and {db}, %, {s.fore}",
                        f"{M} Share of {db} within {da}, %, {s.ez}"])
    if s.regions:
        for a in s.axes:
            d = s.dim(a)
            if kind == "figure":
                per.append([f"Region × {d} Opportunity Matrix, {s.fore}",
                            f"{M} {d} Mix by Region, %, {s.z}"])
            else:
                per.append([f"{M} Value by Region and {d}, {s.vu}, {s.fore}",
                            f"{M} CAGR by Region and {d}, %, {s.fore}"])
    return _rr(*per)


def _regional(s: _Study, kind):
    """A lean regional block (client reference, 2026-10-23: the Croissants
    200/200 exhibit list) - a few headline metrics per region rather than
    every segment axis, Demand Drivers and Competitive Intensity repeated
    for each one."""
    M = s.m()
    if not s.regions:
        return []
    if kind == "figure":
        head = [f"{M} Value by Region, {s.vu}, {s.full}",
                f"{M} Share by Region, %, {s.ez}",
                f"{M} CAGR by Region, %, {s.fore}",
                f"{M} Attractiveness by Region, {s.fore}"]
    else:
        head = [f"{M} Value by Region, {s.vu}, {s.full}",
                f"{M} CAGR and Absolute $ Opportunity by Region, %, {s.vu}, "
                f"{s.fore}"]
        if s.volume:
            head.append(f"{M} Volume by Region, {s.qu}, {s.full}")
    # two representative axes (first and last) stand in for "every axis" -
    # Croissants names two of its six (Product, Channel) rather than all
    reps = (s.axes[:1] + s.axes[-1:]) if len(s.axes) > 1 else s.axes[:1]
    per = []
    for r in s.regions:
        R = s.m(r)
        if kind == "figure":
            t = [f"{R} Value and YoY Growth, {s.vu} and %, {s.full}"]
            t += [f"{R} Share by {s.dim(a)}, %, {s.z}" for a in reps]
            t.append(f"{R} Attractiveness by Country, {s.fore}")
        else:
            t = [f"{R} Value by Country, {s.vu}, {s.full}"]
            if s.volume:
                t.append(f"{R} Volume by Country, {s.qu}, {s.full}")
            t += [f"{R} Value by {s.dim(a)}, {s.vu}, {s.full}" for a in reps]
            rest = [a for a in s.axes if a not in reps]
            if rest:
                t.append(f"{R} {', '.join(s.dim(a) for a in rest)} "
                         f"Opportunity, {s.fore}")
            t.append(f"{R} Leading Countries by Value, {s.vu}, {s.z}")
        per.append(t)
    return head + _rr(*per)


def _country_titles(s: _Study, reg: str, c: str, kind) -> list[str]:
    """One country's exhibits, most important first - the first two are
    what a secondary country gets (client reference, 2026-10-23: the
    Croissants 200/200 list treats every country alike, one combined
    size/growth exhibit and one combined segment-opportunity exhibit,
    rather than drilling deep into a handful of priority countries)."""
    C = s.m(c)
    main = s.axes[0] if s.axes else None
    if kind == "figure":
        size = (f"{C} Value and Volume, {s.vu} and {s.qu}, {s.full}"
                if s.volume else
                f"{C} Value and YoY Growth, {s.vu} and %, {s.full}")
        t = [size,
             f"{C} Segment Opportunity and Growth Matrix, {s.fore}"
             if main else f"{C} Opportunity Score, {s.fore}",
             f"{C} Share within {reg}, %, {s.ez}",
             f"{C} Leading Suppliers Share, %, {s.e}"]
    else:
        size = (f"{C} Value, Volume, Growth and Average Selling Price, "
                f"{s.full}" if s.volume or s.price else
                f"{C} Value and Historical/Forecast Growth, {s.full}")
        t = [size,
             f"{C} Segment Share, CAGR and Opportunity, {s.fore}"
             if main else f"{C} Historical and Forecast Growth, %, {s.full}",
             f"{C} Share within {reg}, %, {s.h1}, {s.e} and {s.z}",
             f"{C} Leading Suppliers and Brands, {s.e}"]
    return t[:PRIORITY_MAX]


def _country(s: _Study, kind, k: int):
    """The country block, sized to `k`: priority countries 8-12 exhibits,
    secondary ones 2-4, never more."""
    ranked = s.country_rank
    if not ranked:
        if s.regions or s.root == "Global":
            return []
        # a single-country study: the root IS the country - give it the
        # deep treatment (every segment and sub-segment exhibit)
        return _single_country(s, kind)
    pools = [_country_titles(s, reg, c, kind) for reg, c in ranked]
    n = len(ranked)
    p = max(1, min(n, k // 10))
    quota = [PRIORITY_MIN] * p
    rest = k - sum(quota)
    for i in range(p, n):
        if rest < SECONDARY_MIN:
            break
        quota.append(SECONDARY_MIN)
        rest -= SECONDARY_MIN
    # grow secondaries to their cap, then priorities to theirs
    for lo, hi, cap in ((p, len(quota), SECONDARY_MAX),
                        (0, p, PRIORITY_MAX)):
        grew = True
        while rest > 0 and grew:
            grew = False
            for i in range(lo, hi):
                if rest and quota[i] < min(cap, len(pools[i])):
                    quota[i] += 1
                    rest -= 1
                    grew = True
    quota = [min(q, len(pools[i])) for i, q in enumerate(quota)]
    out = []
    for i, q in enumerate(quota):
        out += pools[i][:q]
    return out


def _single_country(s: _Study, kind):
    M = s.m()
    t = []
    for a in s.axes:
        d = s.dim(a)
        t += [f"{M} Value of {sub} ({d}), {s.vu}, {s.full}"
              for sub in s.subs[a]]
        t.append(f"{M} Customer and Channel Mix by {d}, %, {s.z}")
    return t


def _pricing(s: _Study, kind):
    M = s.m()
    d0 = s.dim(s.axes[0]) if s.axes else "Product Type"
    t = []
    if kind == "figure":
        if s.price:
            t += [f"{M} Average Selling Price Trend, {s.pu}, {s.full}",
                  f"{M} Average Selling Price by Region, {s.pu}, {s.ez}"
                  if s.regions else
                  f"{M} Average Selling Price Outlook, {s.pu}, {s.fore}"]
            t += [f"{M} Average Selling Price by {s.dim(a)}, {s.pu}, {s.ez}"
                  for a in s.axes]
        t += [f"{M} Value Chain Structure, {s.e}",
              f"{M} Margin Distribution across the Value Chain, %, {s.e}",
              f"{M} Cost Structure Breakdown, %, {s.e}",
              f"{M} Pricing Pressure Assessment by {d0}, {s.fore}",
              f"{M} Import and Export Share by Region, %, {s.h1}"
              if s.regions else f"{M} Import and Export Share, %, {s.h1}",
              f"{M} Trade Balance by Leading Country, {s.vu}, {s.h1}",
              f"{M} Raw Material Price Index, {s.h1}=100, {s.full}",
              f"{M} Distribution Channel Margin Comparison, %, {s.e}",
              f"{M} Price Band Distribution by {d0}, %, {s.e}",
              f"{M} Pricing Factors Impact Assessment, {s.fore}",
              f"{M} Supply Chain Risk Assessment, {s.e}"]
    else:
        t += [f"{s.market} Product Price Comparison by Company and {d0}, "
              f"{s.pu}, {s.e}"]
        if s.price:
            t += [f"{M} Average Selling Price by Region, {s.pu}, {s.full}"
                  if s.regions else
                  f"{M} Average Selling Price, {s.pu}, {s.full}"]
            t += [f"{M} Average Selling Price by {s.dim(a)}, {s.pu}, "
                  f"{s.full}" for a in s.axes]
        top = s.country_rank[0][1] if s.country_rank else s.root
        t += [f"Supplier-to-Customer Relationship Mapping in {top}, {s.e}",
              f"{M} Value Chain Participants by Stage, {s.e}",
              f"{M} Margin Analysis by Value Chain Stage, %, {s.e}",
              f"{M} Cost Structure by Component, %, {s.e}",
              f"{M} Import Value by Leading Country, {s.vu}, {s.hist}",
              f"{M} Export Value by Leading Country, {s.vu}, {s.hist}",
              f"{M} Trade Flow Matrix by Region, {s.vu}, {s.h1}"
              if s.regions else f"{M} Trade Flow Summary, {s.vu}, {s.h1}",
              f"{M} Pricing Factors and Impact Rating, {s.fore}",
              f"{M} Raw Material Price Trend, {s.cur}, {s.full}",
              f"{M} Distribution Channel Margin Structure, %, {s.e}",
              f"{M} Regulatory and Tariff Overview by Country, {s.e}",
              f"{M} Price Forecast by {d0}, {s.pu}, {s.fore}",
              f"{M} Supply Chain Risk Matrix, {s.e}"]
    return t


def _competitive(s: _Study, kind):
    M = s.m()
    mk = s.market.replace(" Market", "")
    t = []
    if kind == "figure":
        t += [f"Company Market Share in the {M}, %, {s.e}",
              f"Competitive Positioning of Leading {mk} Manufacturers, {s.e}",
              f"Company Revenue from the {M}, {s.vu}, {s.e}",
              f"{M} Concentration by Company Tier, %, {s.e}",
              f"{M} Competitive Intensity Assessment, {s.e}"]
        t += [f"Company Market Share in the {s.m(r)}, %, {s.e}"
              for r in s.regions]
        t += [f"{c}: Revenue from the {M} and Share, {s.vu} and %, {s.e}"
              for c in s.companies]
        t += [f"{M} Strategic Developments by Type, %, {s.h1}{_DASH}{s.e}",
              f"{M} Company Footprint by Region, {s.e}"]
    else:
        t += [f"Leading {mk} Manufacturers: Product Portfolio and "
              f"Geographic Presence, {s.e}",
              f"Company Revenue from the {M}, {s.vu}, {s.h1}{_DASH}{s.e}",
              f"Company Market Share in the {M} by Region, %, {s.e}",
              f"{M} Competitive Benchmarking Matrix, {s.e}",
              f"{M} Mergers, Acquisitions and Partnerships, {s.h0}{_DASH}"
              f"{s.e}"]
        t += [f"{c}: Company Overview, Products and Strategy, {s.e}"
              for c in s.companies]
    return t


def _trends(s: _Study, kind):
    M = s.m()
    d0 = s.dim(s.axes[0]) if s.axes else "Segment"
    if kind == "figure":
        return [f"{M} Value Forecast by Scenario, {s.vu}, {s.fore}",
                f"{M} YoY Growth by Scenario, %, {s.fore}",
                f"{M} Key Trends Impact Assessment, {s.fore}",
                f"{M} Drivers, Restraints and Opportunities Impact, {s.fore}",
                f"{M} Opportunity Map by Region and {d0}, {s.fore}"
                if s.regions else f"{M} Opportunity Map by {d0}, {s.fore}",
                f"{M} Technology Adoption Roadmap, {s.full}",
                f"{M} Emerging Opportunity Pockets by {d0}, {s.vu}, {s.fore}",
                f"{M} Risk and Sensitivity Analysis, %, {s.fore}",
                f"{M} Regulatory Landscape Timeline, {s.full}",
                f"{M} Investment Hotspots by Region, {s.fore}"
                if s.regions else f"{M} Investment Hotspots, {s.fore}",
                f"{M} Innovation Pipeline Assessment, {s.e}",
                f"{M} Demand Shift Analysis by {d0}, %, {s.fore}"]
    return [f"{M} Value by Scenario and Region, {s.vu}, {s.fore}"
            if s.regions else f"{M} Value by Scenario, {s.vu}, {s.fore}",
            f"{M} Key Trends and Impact Rating, {s.fore}",
            f"{M} Opportunity Assessment by {d0}, {s.vu}, {s.fore}",
            f"{M} Scenario Assumptions, {s.fore}",
            f"{M} Risk Register and Mitigation, {s.fore}"]


def _usp(s: _Study, kind):
    M = s.m()
    d0 = s.dim(s.axes[0]) if s.axes else "Segment"
    if kind == "figure":
        return [f"{M} White-Space Opportunity Assessment, {s.fore}",
                f"{M} Customer Purchase Criteria Assessment, {s.e}",
                f"{M} Go-to-Market Priority by Region, {s.fore}"
                if s.regions else f"{M} Go-to-Market Priority, {s.fore}",
                f"{M} Unmet Needs Analysis by {d0}, {s.e}",
                f"{M} Right-to-Win Assessment, {s.e}",
                f"{M} Entry Strategy Scorecard, {s.fore}"]
    return [f"{M} Client Opportunity Prioritisation by {d0}, {s.fore}",
            f"{M} Customer Purchase Criteria by {d0}, {s.e}",
            f"{M} Strategic Recommendations and Action Plan, {s.fore}"]


_POOLS = {"summary": _summary, "size": _size, "segment": _segment,
          "cross": _cross, "regional": _regional, "pricing": _pricing,
          "competitive": _competitive, "trends": _trends, "usp": _usp}


# ── allocation ───────────────────────────────────────────────────────
def _adapted_weights(s: _Study, w: dict) -> dict:
    """Bend the weightage to the study: more segment axes -> more segment
    exhibits, one axis -> no cross-segment, no regions -> the regional share
    goes to the country and segment blocks."""
    w = dict(w)
    n_ax = len(s.axes)
    # four axes is the weightage's reference study
    w["segment"] *= min(2.0, max(0.5, (n_ax / 4.0) ** 0.75)) if n_ax else 0.0
    pairs = n_ax * (n_ax - 1) / 2
    w["cross"] *= (min(1.6, max(0.5, (pairs / 6.0) ** 0.5)) if pairs else 0.0)
    if not s.regions:
        freed = w["regional"]
        w["regional"] = 0.0
        w["country"] += freed * 0.5
        w["segment"] += freed * 0.5
    if not s.country_rank and s.root == "Global":
        w["segment"] += w["country"]
        w["country"] = 0.0
    return w


def _apportion(weights: dict, total: int) -> dict:
    """Largest-remainder rounding of `weights` to integers summing to
    `total`."""
    s = sum(v for v in weights.values() if v > 0)
    if s <= 0:
        return {k: 0 for k in weights}
    raw = {k: max(0.0, v) * total / s for k, v in weights.items()}
    out = {k: int(math.floor(v)) for k, v in raw.items()}
    left = total - sum(out.values())
    for k in sorted(raw, key=lambda k: -(raw[k] - out[k]))[:left]:
        out[k] += 1
    return out


def _pool(s: _Study, key: str, kind: str, k: int) -> list[str]:
    if key == "country":
        return _country(s, kind, k)
    fn = _POOLS[key]
    return _dedupe(fn(s, kind))


def _dedupe(titles):
    seen, out = set(), []
    for t in titles:
        key = t.casefold()
        if t and key not in seen:
            seen.add(key)
            out.append(t)
    return out


def plan(kind: str, root: str | None = None) -> list[tuple[str, str]]:
    """[(block title, exhibit title)] for `kind` ('figure' or 'table'), in
    report order, sized to the configured total."""
    root = root or geo.current()
    s = _Study(root)
    fw, tw, ftot, ttot = _base_weights()
    w = _adapted_weights(s, fw if kind == "figure" else tw)
    total = ftot if kind == "figure" else ttot
    alloc = _apportion(w, total)
    pools = {k: _pool(s, k, kind, alloc[k]) for k, *_ in BLOCKS}
    take = {k: min(alloc[k], len(pools[k])) for k in pools}
    # a block that cannot fill its share hands the remainder to the blocks
    # that can, in proportion to their weight
    short = total - sum(take.values())
    for _ in range(6):
        if short <= 0:
            break
        spare = {k: w[k] for k in pools
                 if w[k] > 0 and len(pools[k]) > take[k]}
        if not spare:
            break
        extra = _apportion(spare, short)
        for k, n in extra.items():
            if k == "country":
                grown = _country(s, kind, take[k] + n)
                pools[k] = grown
                add = min(n, len(grown) - take[k])
            else:
                add = min(n, len(pools[k]) - take[k])
            take[k] += max(0, add)
        short = total - sum(take.values())
    out = []
    for key, title, *_ in BLOCKS:
        out += [(title, t) for t in pools[key][:take[key]]]
    return out


def figure_entries(root: str | None = None) -> list[str]:
    """"Figure 1: ...", the full study's List of Figures."""
    return [f"Figure {i}: {t}"
            for i, (_, t) in enumerate(plan("figure", root), start=1)]


def table_entries(root: str | None = None) -> list[str]:
    """"Table 1: ...", the full study's List of Tables."""
    return [f"Table {i}: {t}"
            for i, (_, t) in enumerate(plan("table", root), start=1)]
