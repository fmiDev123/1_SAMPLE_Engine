"""Derived measures - the Python equivalents of the report's DAX measures."""
from __future__ import annotations

import numpy as np
import pandas as pd

from .config import EST_YEAR, END_YEAR, HIST_START, HIST_END
from . import config
from . import data
from . import geo as _geo

# ── Primitives ───────────────────────────────────────────────────────

def imax(s, default=None):
    """`idxmax` that survives an empty or all-NaN column.

    Ranking is everywhere in this report â€" leader, fastest, biggest gain â€" and
    pandas raises on both an empty frame ("argmax of an empty sequence") and a
    column with nothing but NaN ("Encountered all NA values").  Both happen for
    real: a segment the workbook sizes in the other unit, a CAGR over a series
    of zeros, a workbook saved before its figures were pasted in.  One bad
    column should not take a 174-slide export down, so the first label is
    returned when there is nothing to rank."""
    try:
        if len(s) == 0:
            return default
        v = s.dropna()
        return s.idxmax() if len(v) else (default if default is not None
                                          else s.index[0])
    except (ValueError, TypeError):
        return default if default is not None else (
            s.index[0] if len(getattr(s, "index", [])) else None)


def imin(s, default=None):
    """`idxmin` with the same guarantees as `imax`."""
    try:
        if len(s) == 0:
            return default
        v = s.dropna()
        return s.idxmin() if len(v) else (default if default is not None
                                          else s.index[0])
    except (ValueError, TypeError):
        return default if default is not None else (
            s.index[0] if len(getattr(s, "index", [])) else None)


def cagr(v0: float, v1: float, n_years: int) -> float:
    if n_years <= 0 or pd.isna(v0) or pd.isna(v1) or v0 <= 0:
        return np.nan
    return (v1 / v0) ** (1.0 / n_years) - 1.0

def series_cagr(s: pd.Series, y0: int, y1: int) -> float:
    return cagr(s.get(y0, np.nan), s.get(y1, np.nan), y1 - y0)

def hist_cagr(s: pd.Series) -> float:
    """H-CAGR: 2021-2025."""
    return series_cagr(s, HIST_START, HIST_END)

def fore_cagr(s: pd.Series) -> float:
    """F-CAGR: 2026-2036."""
    return series_cagr(s, EST_YEAR, END_YEAR)

def yoy(s: pd.Series) -> pd.Series:
    return s.pct_change()

# ── Price benchmarking ───────────────────────────────────────────────
def avg_price(value_row, volume_row, years) -> float:
    """Average selling price over ``years`` = total Value / total Volume.

    This is the report's single definition of price (there is no price column
    in any workbook), and summing before dividing makes a multi-year figure a
    VOLUME-WEIGHTED average rather than a mean of yearly ratios.

    Value is quoted in a magnitude ('US$ Mn'), so the raw quotient is that many
    MILLION dollars per volume unit; `config.price_scale()` puts it back into
    currency units.  Without it every price in the deck was too small by 1e6 -
    silently, because the result still looked like a price: CF-PEEK printed
    '0.73 US$/unit' for a market at US$ 727,000 per ton, and it only became
    visible on a high-tonnage market like coffee, where the quotient fell below
    a rounding threshold and printed '0.00'.  Ratios of two prices are
    unaffected either way, which is why the percentage columns beside those
    zeros always looked right."""
    years = list(years)
    vol = float(pd.Series(volume_row)[years].astype(float).sum())
    if not vol:
        return float("nan")
    val = float(pd.Series(value_row)[years].astype(float).sum())
    return val * config.price_scale() / vol

def price_vs_parent(geo: str, years):
    """How far each child geography's price sits from ``geo``'s OWN price.

        variation % = (child_price - parent_price) / parent_price * 100

    Positive means the child is dearer than the parent, negative cheaper, zero
    exactly on it.  With ``geo="Global"`` the parent price is the global price,
    so this is the "vs Global" benchmark used by the pricing page.

    Everything comes from the existing model - `data.children_table` for the
    children and `data.total_series` for the parent - so it stays correct for
    any workbook and any slicer selection.

    Returns ``(variation_pct, parent_price, child_prices)``; the index keeps
    the workbook's own child order."""
    years = list(years)
    kv = data.children_table(geo, "Value").astype(float)
    kw = data.children_table(geo, "Volume").astype(float)
    child = pd.Series(
        {str(c): avg_price(kv.loc[c], kw.loc[c], years) for c in kv.index})
    parent = avg_price(data.total_series(geo, "Value"),
                       data.total_series(geo, "Volume"), years)
    if not parent or pd.isna(parent):
        return pd.Series(0.0, index=child.index), float("nan"), child
    return ((child - parent) / parent * 100.0).fillna(0.0), parent, child

def share_base(parent: str, unit: str):
    """The denominator a geography's SHARE is taken against: the sum of the
    parent's children, not the parent's own row.

    For a well-formed workbook the two are identical.  Where they are not -
    some sheets state a parent total its children do not add up to - the
    children's sum is the only base on which a page's shares total 100%, and
    it is what every regional exhibit already uses.  Dividing by the parent's
    own row instead let one slide call a country "138.1% of North America"
    while another slide in the same deck called it 61.4%."""
    try:
        kids = data.children_table(parent, unit)
        if kids is not None and len(kids):
            return kids.sum(axis=0)
    except Exception:
        pass
    return data.total_series(parent, unit)


def growth_deviation_bps(geo: str, unit: str, benchmark: str) -> float:
    """How much faster/slower `geo` grows than a `benchmark` geography over the
    forecast, in BASIS POINTS (1 percentage point = 100 bps).
    = (F-CAGR of geo − F-CAGR of benchmark) × 10,000.  One source for the
    dashboard gauge and the PowerPoint deviation card."""
    fc = fore_cagr(data.total_series(geo, unit))
    bc = fore_cagr(data.total_series(benchmark, unit))
    if pd.isna(fc) or pd.isna(bc):
        return np.nan
    return (fc - bc) * 10_000.0

# ── Geography-level measures ─────────────────────────────────────────
def geo_kpis(geo: str, unit: str) -> dict:
    s = data.total_series(geo, unit)
    return {
        "est":    s[EST_YEAR],
        "end":    s[END_YEAR],
        "hcagr":  hist_cagr(s),
        "fcagr":  fore_cagr(s),
        "incr5":  s[EST_YEAR + 5] - s[EST_YEAR],
        "incr10": s[END_YEAR] - s[EST_YEAR],
    }

# ── Segment-level measures ───────────────────────────────────────────
def segment_summary(geo: str, unit: str, segment: str) -> pd.DataFrame:
    """Per sub-segment: est/end size, shares, F-CAGR, BPS, abs opportunity."""
    tbl = data.segment_table(geo, unit, segment)
    total = tbl.sum(axis=0)
    out = pd.DataFrame(index=tbl.index)
    out["est"] = tbl[EST_YEAR]
    out["end"] = tbl[END_YEAR]
    out["share_est"] = tbl[EST_YEAR] / total[EST_YEAR]
    out["share_end"] = tbl[END_YEAR] / total[END_YEAR]
    out["bps"] = (out["share_end"] - out["share_est"]) * 10_000
    out["hcagr"] = [hist_cagr(tbl.loc[i]) for i in tbl.index]
    out["fcagr"] = [fore_cagr(tbl.loc[i]) for i in tbl.index]
    out["abs_opp"] = tbl[END_YEAR] - tbl[EST_YEAR]
    total_fcagr = fore_cagr(total)
    out["variance"] = out["fcagr"] - total_fcagr
    n = max(len(out), 1)
    out["ms_index"] = out["share_est"] / (1.0 / n) * 100.0
    out["cagr_index"] = out["fcagr"] / total_fcagr * 100.0
    return out

def share_table(tbl: pd.DataFrame) -> pd.DataFrame:
    """Column-wise share (%) of a sub-segment/child x year matrix."""
    return tbl.div(tbl.sum(axis=0), axis=1) * 100.0

def yoy_table(tbl: pd.DataFrame) -> pd.DataFrame:
    """Row-wise YoY (%) of a sub-segment/child x year matrix."""
    return tbl.pct_change(axis=1) * 100.0

# ── Waterfall helpers ────────────────────────────────────────────────
def forecast_waterfall(s: pd.Series) -> pd.DataFrame:
    """Base 2026, yearly increments, total 2036 - the report's
    'Absolute $ Opportunity' waterfall."""
    rows = [{"label": str(EST_YEAR), "value": s[EST_YEAR], "measure": "absolute"}]
    for y in range(EST_YEAR + 1, END_YEAR + 1):
        rows.append({"label": str(y), "value": s[y] - s[y - 1],
                     "measure": "relative"})
    rows.append({"label": f"{END_YEAR} Total", "value": s[END_YEAR],
                 "measure": "total"})
    return pd.DataFrame(rows)

def contribution_waterfall(tbl: pd.DataFrame) -> pd.DataFrame:
    """Growth contribution: total est -> each sub-segment's delta -> total end."""
    total = tbl.sum(axis=0)
    rows = [{"label": str(EST_YEAR), "value": total[EST_YEAR],
             "measure": "absolute"}]
    for name in tbl.index:
        rows.append({"label": name,
                     "value": tbl.loc[name, END_YEAR] - tbl.loc[name, EST_YEAR],
                     "measure": "relative"})
    rows.append({"label": f"{END_YEAR} Total", "value": total[END_YEAR],
                 "measure": "total"})
    return pd.DataFrame(rows)

# ── Country indices (Executive Summary page) ─────────────────────────
_CI_COLS = ["Region", "Country", "est", "end", "fcagr", "abs_opp",
            "CAGR Index", "Absolute Index"]

def country_indices(unit: str = "Value") -> pd.DataFrame:
    """Per-geography CAGR / Absolute indices for the Summary page.

    Uses the country slicer options; on datasets with no country level
    (single-country or region-only workbooks) it falls back to the
    regions, then to Global - so the page degrades instead of crashing.

    Scoped to the ACTIVE geography: a Europe deck lists Europe's countries and
    indexes them against Europe's own CAGR and increment, so the page's
    "benchmark" is the one the reader selected.  Under Global it is every
    country against the world, exactly as before.
    """
    root = _geo.current()
    g = data.total_series(root, unit)
    g_fcagr, g_abs = fore_cagr(g), g[END_YEAR] - g[EST_YEAR]
    geos = ((data.ALL_COUNTRIES or data.REGIONS or [root]) if _geo.is_global(root)
            else (data.children(root) or [root]))
    rows = []
    # real countries only: 'Rest of ...' / 'Other ...' are buckets, not places
    geos = [g for g in geos if not data.is_bucket(g)] or list(geos)
    for geo in geos:
        s = data.total_series(geo, unit)
        fc, abs_opp = fore_cagr(s), s[END_YEAR] - s[EST_YEAR]
        rows.append({
            "Region": data.region_of(geo) or "Global",
            "Country": geo, "est": s[EST_YEAR], "end": s[END_YEAR],
            "fcagr": fc, "abs_opp": abs_opp,
            "CAGR Index": (fc / g_fcagr * 100.0) if g_fcagr else 0.0,
            "Absolute Index": (abs_opp / g_abs * 100.0) if g_abs else 0.0,
        })
    return pd.DataFrame(rows, columns=_CI_COLS)
