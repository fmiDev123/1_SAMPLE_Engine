"""Auto-generated commentary - the Python equivalent of the PBIX
'Write-Up' DAX text measures shown in the strips under each page."""
from __future__ import annotations

from . import data, metrics
from . import geo as _geo          # `geo` is a parameter name throughout here
from .components import fmt_num, fmt_pct
from .config import END_YEAR, EST_YEAR, HIST_END, HIST_START, UNIT_LABELS
from .data import MARKET

_MKT = MARKET.lower() + " market"       # e.g. "matcha market"


def _remarket() -> None:
    """Re-derive what this module built from the market name.

    `markets._bind()` rebinds MARKET itself; `_MKT` is derived FROM it, so it
    would otherwise keep naming the market this process loaded first.
    """
    global _MKT
    _MKT = MARKET.lower() + " market"

def market_summary(geo: str, unit: str) -> str:
    k = metrics.geo_kpis(geo, unit)
    verb = ("accelerating from" if k["fcagr"] >= k["hcagr"]
            else "moderating from")
    leaders = "; ".join(
        f"'{metrics.segment_summary(geo, unit, s)['fcagr'].idxmax()}' "
        f"leads growth by {s.lower()}"
        for s in data.SEGMENTS[:3]
    )
    return (
        f"The {geo} {_MKT} is projected to grow from "
        f"{fmt_num(k['est'])} {UNIT_LABELS[unit]} in {EST_YEAR} to "
        f"{fmt_num(k['end'])} {UNIT_LABELS[unit]} by {END_YEAR}, at a CAGR "
        f"of {fmt_pct(k['fcagr'])} {verb} {fmt_pct(k['hcagr'])} "
        f"during {HIST_START}–{HIST_END}. {leaders}."
    )

def summary(unit: str) -> str:
    df = metrics.country_indices(unit)
    if df.empty:
        return (f"This {_MKT} workbook has a single geography, so no "
                f"cross-country index comparison is available.")
    top_c = df.nlargest(1, "CAGR Index").iloc[0]
    top_a = df.nlargest(1, "Absolute Index").iloc[0]
    return (
        f"{top_c['Country']} posts the highest CAGR Index "
        f"({top_c['CAGR Index']:.0f}), growing at {fmt_pct(top_c['fcagr'])} "
        f"against the {_geo.adj()} benchmark, while {top_a['Country']} offers "
        f"the largest absolute opportunity "
        f"{fmt_num(top_a['abs_opp'])} {UNIT_LABELS[unit]} of incremental "
        f"market between {EST_YEAR} and {END_YEAR} "
        f"({top_a['Absolute Index']:.0f}% of the {_geo.adj()} increment)."
    )

def bps(geo: str, segment: str, unit: str) -> str:
    s = metrics.segment_summary(geo, unit, segment)
    if not len(s):
        return (f"The workbook carries no {unit.lower()} figures for "
                f"'{segment}' in {geo}, so no share shift can be measured "
                f"for it in this unit.")
    gain, lose = metrics.imax(s["bps"]), metrics.imin(s["bps"])
    return (
        f"Between {EST_YEAR} and {END_YEAR}, '{gain}' gains the most share "
        f"in the {geo} market ({s.loc[gain, 'bps']:+,.0f} BPS) while "
        f"'{lose}' cedes {abs(s.loc[lose, 'bps']):,.0f} BPS. Bubbles right "
        f"of the CAGR-Index line grow faster than the segment average."
    )

def country_kpi(country: str, unit: str) -> str:
    region = data.region_of(country) or _geo.current()
    s = data.total_series(country, unit)
    fc = metrics.fore_cagr(s)
    reg_fc = metrics.fore_cagr(data.total_series(region, unit))
    verb = "outpacing" if fc >= reg_fc else "trailing"
    dev_bps = abs(metrics.growth_deviation_bps(country, unit, region))
    return (
        f"{country}'s {_MKT} expands at {fmt_pct(fc)} CAGR through "
        f"{END_YEAR}, {verb} the {region} benchmark of {fmt_pct(reg_fc)} "
        f"by {dev_bps:,.0f} BPS."
    )

def country_share(country: str, unit: str) -> str:
    region = data.region_of(country) or _geo.GLOBAL
    root = _geo.current()
    s = data.total_series(country, unit)
    r = metrics.share_base(region, unit)          # sum of the region's countries
    if root == region:
        # A deck scoped to this country's OWN region: the wider world is not
        # what the report is about, and quoting both shares just stutters the
        # same number twice ("14.5% of the global and 14.5% of the Europe").
        return (
            f"{country} holds {s[EST_YEAR] / r[EST_YEAR]:.1%} of the "
            f"{region} {_MKT} in {EST_YEAR}; by {END_YEAR} that share moves "
            f"to {s[END_YEAR] / r[END_YEAR]:.1%}."
        )
    g = metrics.share_base(root, unit)
    return (
        f"{country} holds {s[EST_YEAR] / g[EST_YEAR]:.1%} of the "
        f"{_geo.adj()} and {s[EST_YEAR] / r[EST_YEAR]:.1%} of the {region} "
        f"{_MKT} in {EST_YEAR}; by {END_YEAR} those shares move to "
        f"{s[END_YEAR] / g[END_YEAR]:.1%} and "
        f"{s[END_YEAR] / r[END_YEAR]:.1%} respectively."
    )

def country_momentum(country: str, unit: str) -> str:
    s = data.total_series(country, unit)
    incr = s[END_YEAR] - s[EST_YEAR]
    mult = s[END_YEAR] / s[EST_YEAR]
    return (
        f"The forecast build-up adds {fmt_num(incr)} {UNIT_LABELS[unit]} "
        f"between {EST_YEAR} and {END_YEAR} the {country} market grows "
        f"to {mult:.1f}× its {EST_YEAR} size over the decade."
    )
