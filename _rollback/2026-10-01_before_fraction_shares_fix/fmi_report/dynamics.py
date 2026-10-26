"""Market-dynamics derivation for the '- Market Dynamics' report.

Everything on the Dynamics slides - the driving-factor scores, the driver /
restraint / opportunity write-ups, the impact-on-timeline grids, the milestone
table and the Porter's Five Forces ratings - is **computed from the loaded
workbook** (``data`` + ``metrics``).  Nothing is market-specific or hardcoded:
segment names, sub-segment names, geographies, units, years and every number in
the prose come from whatever dataset is currently loaded, so swapping the
``Dataset_*.xlsx`` regenerates the whole section.

Design mirrors ``competitors.py`` / ``pricing_drivers.py`` / ``value_chain.py``:
this module is *configuration + derivation*, never UI code.  The page templates
in ``templates_dynamics.py`` read from these functions and hold no data.

How a qualitative label becomes a number
----------------------------------------
Each factor / force is backed by a real metric (a CAGR, a share, a
concentration index, a price movement...).  ``_score(value, mid, span)`` maps that
metric onto the 0-10 scale used by the radar and the force ratings: ``mid``
scores 5, ``mid ± span`` scores 10 / 0.  The mapping constants are documented at
each call site so they can be tuned in one place.
"""
from __future__ import annotations

import math
from functools import lru_cache

import pandas as pd

from . import config
from . import data, metrics
from .components import fmt_num, fmt_pct
from .config import (END_YEAR, EST_YEAR, HIST_END, HIST_START,
                     UNIT_LABELS, YEARS, plural)
from .data import MARKET_LABEL as MKT

INDICATIVE_NOTE = (
    "Note: the dynamics framing is indicative factors, impact ratings and "
    "narrative are generated from the loaded dataset's own values."
)

SCALE_MAX = 10.0                      # radar / force axis maximum

# Pastel accents for the Five-Forces discs (mirrors the reference deck).
FORCE_COLORS = ["#F2D24B", "#D9D9D9", "#E88A8A", "#C9C0E0", "#BFDCEA"]
FORCE_EDGES = ["#D8B92F", "#BFBFBF", "#D06A6A", "#A99CC9", "#8FBBD0"]
# Pictogram drawn inside each disc, and its ink colour - the reference deck
# shows an icon per force (dollar / cart / target / cubes / people), never a
# score.  Icons are built from drawing primitives in templates_dynamics so the
# dashboard and the deck render the identical shape.
FORCE_ICONS = ["dollar", "cart", "target", "cubes", "people"]
# First entry darkened from #D9A400: gold ink on the gold disc left the
# "$" pictogram barely visible.
FORCE_INK = ["#8A6A00", "#2B2B2B", "#B23A3A", "#4A3F6B", "#5B7C99"]

# Kind -> (accent, tint) for the Drivers / Restraints / Opportunities pages.
KIND_THEME = {
    "drivers":       ("#1668A8", "#DCEAF6"),
    "restraints":    ("#C0603A", "#F7E3DA"),
    "opportunities": ("#5B7C99", "#E8EEF4"),
}
KIND_LABEL = {"drivers": "Market Drivers", "restraints": "Market Restraints",
              "opportunities": "Market Opportunities"}
KIND_ORDER = ["drivers", "restraints", "opportunities"]

# Likelihood-of-impact cells are GRADED (High / Medium / Low) rather than a
# binary on/off.  The old binary test ("does this item out-grow the market in
# this period?") came out True for almost every item in every period, so every
# pointer's grid looked identical and carried no information.
IMPACT_ON = "#F2C230"                  # kept: 'High' fill (was the on colour)
IMPACT_OFF = "#FFFFFF"
IMPACT_LEVELS = ("Low", "Medium", "High")
IMPACT_FILL = {"High": "#F2C230", "Medium": "#F9E4A6", "Low": "#FFFFFF"}
IMPACT_TEXT = {"High": "#4A3A05", "Medium": "#6B5613", "Low": "#9AA0A6"}


# ── small helpers ────────────────────────────────────────────────────
def _clamp(v, lo=0.0, hi=SCALE_MAX):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return (lo + hi) / 2.0
    if v != v:                                     # NaN
        return (lo + hi) / 2.0
    return max(lo, min(hi, v))


def _score(value, mid, span):
    """Map a real metric onto 0-10: ``mid`` -> 5, ``mid ± span`` -> 10 / 0."""
    if not span:
        return 5.0
    return _clamp(5.0 + 5.0 * (float(value) - mid) / float(span))


def _pick_segment(keywords, fallback_idx=0):
    """First segment whose name mentions one of ``keywords`` (else the i-th)."""
    for s in data.SEGMENTS:
        low = s.lower()
        if any(k in low for k in keywords):
            return s
    return data.seg(fallback_idx)


def _short(seg) -> str:
    """'By Distribution Channel' -> 'distribution channel'."""
    return str(seg or "").replace("By ", "").strip().lower() or "segment"


def _cap(text: str, n: int) -> str:
    text = str(text)
    return text if len(text) <= n else text[:n - 1].rstrip(" ,;-") + "…"


def _series(geo: str, unit: str):
    try:
        s = data.total_series(geo, unit).astype(float)
        return s if float(s.fillna(0).abs().sum()) > 0 else None
    except Exception:
        return None


def _price_series(geo: str = "Global"):
    """Average realisation (Value ÷ Volume) per year, or None if unavailable."""
    val, vol = _series(geo, "Value"), _series(geo, "Volume")
    if val is None or vol is None:
        return None
    p = val * config.price_scale() / vol.replace(0.0, float("nan"))
    return p if p.notna().any() else None


def _or(series, fallback):
    """``series if series is not None else fallback`` - pandas Series have no
    truth value, so ``a or b`` cannot be used on them."""
    return fallback if series is None else series


def _cagr(s, y0, y1):
    if s is None:
        return 0.0
    try:
        v = metrics.series_cagr(s, y0, y1)
        return 0.0 if v != v else float(v)
    except Exception:
        return 0.0


def _hhi(shares) -> float:
    """Herfindahl concentration of a share series (fractions), 0..1."""
    s = pd.Series(shares).astype(float).fillna(0.0)
    tot = float(s.sum())
    if tot <= 0:
        return 0.0
    return float(((s / tot) ** 2).sum())


# ── impact periods (derived from the workbook's own year range) ──────
@lru_cache(maxsize=1)
def periods() -> list:
    """Four consecutive year buckets spanning the dataset - two historical,
    two forecast - used by the 'Likelihood of impact on timeline' grids.

    Split points are proportional, so a 2021-2036 workbook yields
    2021-22 / 2023-25 / 2026-28 / 2029-36 and any other range adapts."""
    ys = [int(y) for y in YEARS]
    est_i = ys.index(EST_YEAR) if EST_YEAR in ys else max(1, len(ys) // 2)
    hist, fore = ys[:est_i] or ys[:1], ys[est_i:] or ys[-1:]

    def split(chunk):
        if len(chunk) < 2:
            return [chunk, chunk]
        k = max(1, min(len(chunk) - 1, int(round(len(chunk) * 0.3))))
        return [chunk[:k], chunk[k:]]

    a, b = split(hist)
    c, d = split(fore)
    return [(ch[0], ch[-1]) for ch in (a, b, c, d)]


def period_labels() -> list:
    out = []
    for y0, y1 in periods():
        out.append(str(y0) if y0 == y1 else f"{y0}–{str(y1)[-2:]}")
    return out


def _impact_gap(series, benchmark, favourable=True) -> list:
    """Per-period impact STRENGTH for one item, in percentage points.

    The item's own CAGR over each period is measured against the market's.
    ``favourable=True`` (drivers / opportunities) scores a period higher the
    further the item out-grows the market; ``False`` (restraints) scores it
    higher the further the item LAGS - i.e. where the restraint bites hardest.
    Returned as a signed magnitude so `_grade_impacts` can rank the periods
    against each other; NaN where a period is too short to measure."""
    gaps = []
    for y0, y1 in periods():
        if y0 == y1 or series is None or benchmark is None:
            gaps.append(float("nan"))
            continue
        a, b = _cagr(series, y0, y1), _cagr(benchmark, y0, y1)
        d = (a - b) * 100.0
        gaps.append(d if favourable else -d)
    return gaps


def _grade_impacts(out: list) -> None:
    """Turn each item's raw per-period gaps into High/Medium/Low levels,
    IN PLACE (sets item['impact']).

    Levels are assigned by RANK across every (item, period) cell on the page,
    split into three equal buckets.  Ranking rather than absolute thresholds
    means the grid always shows a real spread - which periods and which
    pointers carry the most impact - instead of the old binary test that came
    out identical for nearly every row.  It is also unit-free, so it adapts to
    any market's growth profile."""
    cells = [(i, j, g) for i, it in enumerate(out)
             for j, g in enumerate(it["gap"])]
    ranked = sorted((c for c in cells if c[2] == c[2]),   # drop NaN
                    key=lambda c: c[2])
    n = len(ranked)
    level_of = {}
    for pos, (i, j, _) in enumerate(ranked):
        # lowest third -> Low, middle -> Medium, top -> High
        level_of[(i, j)] = IMPACT_LEVELS[min(2, pos * 3 // n)] if n else "Medium"
    for i, it in enumerate(out):
        it["impact"] = [level_of.get((i, j), "Low")
                        for j in range(len(it["gap"]))]
        # the period where this item bites hardest - the one cell its grid
        # highlights (client reference 'RO', 2026-10-01)
        valid = [(g, j) for j, g in enumerate(it["gap"]) if g == g]
        it["peak"] = max(valid)[1] if valid else 0
        it.pop("gap", None)


# ── the fact base ────────────────────────────────────────────────────
def _value_segment(region, keywords, fallback_idx=0):
    """A segment matching `keywords` that the workbook actually sizes in VALUE.

    Every fact built from it reads the value summary, so a breakdown carried in
    volume only (a platform split counted in procedures, say) would leave the
    summary empty and take the page down with an argmax on nothing.
    """
    seg = _pick_segment(keywords, fallback_idx)
    if len(metrics.segment_summary(region, "Value", seg)):
        return seg
    return next((sg for sg in data.SEGMENTS
                 if len(metrics.segment_summary(region, "Value", sg))), seg)


@lru_cache(maxsize=32)
def facts(region: str = "Global") -> dict:
    """Every number the Dynamics slides draw on, computed once per geography."""
    f: dict = {"region": region, "market": MKT,
               "vlab": UNIT_LABELS["Value"], "qlab": UNIT_LABELS["Volume"]}
    val = _series(region, "Value")
    if val is None:                                # value-less workbook
        val = _series(region, "Volume")
        f["vlab"] = UNIT_LABELS["Volume"]
    vol = _series(region, "Volume")
    price = _price_series(region)
    f["val"], f["vol"], f["price"] = val, vol, price

    f["est"] = float(val[EST_YEAR]) if val is not None else 0.0
    f["end"] = float(val[END_YEAR]) if val is not None else 0.0
    f["incr"] = f["end"] - f["est"]
    f["mult"] = (f["end"] / f["est"]) if f["est"] else 0.0
    f["fc"] = _cagr(val, EST_YEAR, END_YEAR)
    f["hc"] = _cagr(val, HIST_START, HIST_END)
    f["qfc"] = _cagr(vol, EST_YEAR, END_YEAR)
    f["qhc"] = _cagr(vol, HIST_START, HIST_END)
    f["pc"] = _cagr(price, EST_YEAR, END_YEAR)
    f["p_end"] = float(price[END_YEAR]) if price is not None else 0.0
    f["p_est"] = float(price[EST_YEAR]) if price is not None else 0.0

    # ── child geographies ──
    kids = data.children_table(region, "Value").astype(float)
    f["kids"] = kids
    f["child_kind"] = "region" if region == "Global" else "country"
    tot_end = float(kids[END_YEAR].sum()) or 1.0
    f["kid_share"] = (kids[END_YEAR] / tot_end).sort_values(ascending=False)
    f["kid_cagr"] = pd.Series({str(r): _cagr(kids.loc[r], EST_YEAR, END_YEAR)
                               for r in kids.index})
    f["top_kid"] = str(f["kid_share"].index[0])
    f["top_kid_share"] = float(f["kid_share"].iloc[0])
    f["fast_kid"] = str(metrics.imax(f["kid_cagr"]))
    f["fast_kid_cagr"] = float(f["kid_cagr"].max())
    f["slow_kid"] = str(metrics.imin(f["kid_cagr"]))
    f["slow_kid_cagr"] = float(f["kid_cagr"].min())
    f["n_kids"] = int(len(kids))
    f["above_pace"] = int((f["kid_cagr"] >= f["fc"]).sum())
    f["kid_hhi"] = _hhi(kids[END_YEAR])

    # ── product segment (the market's primary breakdown) ──
    # The primary breakdown has to be one the workbook actually sizes in VALUE:
    # a study can carry a platform split in procedures and a revenue split in
    # value, and every fact below reads the value summary.
    pseg = _value_segment(region,
                          ["product", "type", "formulation", "grade"], 0)
    f["prod_seg"], f["prod_label"] = pseg, _short(pseg)
    ps = metrics.segment_summary(region, "Value", pseg)
    f["prod_tbl"] = data.segment_table(region, "Value", pseg).astype(float)
    f["prod_sum"] = ps
    f["lead_sub"] = str(metrics.imax(ps["end"]))
    f["lead_sub_share"] = float(ps.loc[f["lead_sub"], "share_end"])
    f["fast_sub"] = str(metrics.imax(ps["fcagr"]))
    f["fast_sub_cagr"] = float(ps["fcagr"].max())
    f["slow_sub"] = str(metrics.imin(ps["fcagr"]))
    f["slow_sub_cagr"] = float(ps["fcagr"].min())
    f["drop_sub"] = str(metrics.imin(ps["bps"]))
    f["drop_sub_bps"] = float(ps["bps"].min())
    f["gain_sub"] = str(metrics.imax(ps["bps"]))
    f["gain_sub_bps"] = float(ps["bps"].max())
    f["n_sub"] = int(len(ps))
    f["sub_hhi"] = _hhi(ps["end"])

    # ── route-to-market segment (channel / application) ──
    cseg = _value_segment(region,
                          ["channel", "distribution", "application",
                           "end use", "end-use"],
                          min(1, len(data.SEGMENTS) - 1))
    f["chan_seg"], f["chan_label"] = cseg, _short(cseg)
    cs = metrics.segment_summary(region, "Value", cseg)
    f["chan_sum"] = cs
    f["chan_tbl"] = data.segment_table(region, "Value", cseg).astype(float)
    f["lead_chan"] = str(metrics.imax(cs["end"]))
    f["lead_chan_share"] = float(cs.loc[f["lead_chan"], "share_end"])
    f["fast_chan"] = str(metrics.imax(cs["fcagr"]))
    f["fast_chan_cagr"] = float(cs["fcagr"].max())
    f["n_chan"] = int(len(cs))
    f["chan_hhi"] = _hhi(cs["end"])

    f["n_seg"] = len(data.SEGMENTS)
    f["n_countries"] = len(data.ALL_COUNTRIES)
    f["n_regions"] = len(data.REGIONS)

    # ── biggest absolute opportunity among the country slicer options ──
    try:
        ci = metrics.country_indices("Value")
        ci = ci.sort_values("abs_opp", ascending=False)
        f["opp_geo"] = str(ci.iloc[0]["Country"])
        f["opp_abs"] = float(ci.iloc[0]["abs_opp"])
        ci_fast = ci.sort_values("fcagr", ascending=False)
        f["opp_fast_geo"] = str(ci_fast.iloc[0]["Country"])
        f["opp_fast_cagr"] = float(ci_fast.iloc[0]["fcagr"])
        # under-penetrated: below-median size but above-median growth
        med_e = float(ci["end"].median())
        med_c = float(ci["fcagr"].median())
        under = ci[(ci["end"] < med_e) & (ci["fcagr"] > med_c)]
        under = under.sort_values("fcagr", ascending=False)
        f["under_geo"] = (str(under.iloc[0]["Country"]) if len(under)
                          else f["opp_fast_geo"])
        f["under_cagr"] = (float(under.iloc[0]["fcagr"]) if len(under)
                           else f["opp_fast_cagr"])
        f["n_under"] = int(len(under))
    except Exception:
        f["opp_geo"] = f["opp_fast_geo"] = f["under_geo"] = f["fast_kid"]
        f["opp_abs"] = f["incr"]
        f["opp_fast_cagr"] = f["under_cagr"] = f["fast_kid_cagr"]
        f["n_under"] = 0
    return f


# ══════════════════ 1. Dynamics by geography (matrix) ════════════════
def geography_matrix(region: str = "Global", top: int = 10) -> dict:
    """Per-geography growth-driver / restraint / opportunity phrases.

    Rows are the largest child geographies of ``region`` by END-year value;
    every phrase is assembled from that geography's own numbers (its rank,
    forecast CAGR versus the market's, its leading and fastest sub-segments,
    its price position) so the table restates itself for any workbook."""
    f = facts(region)
    kids, pace = f["kids"], f["fc"]
    order = list(f["kid_share"].index[:max(1, top)])
    pseg = f["prod_seg"]
    rows = []
    for rank, geo in enumerate(order, 1):
        geo = str(geo)
        s = kids.loc[geo]
        fc = _cagr(s, EST_YEAR, END_YEAR)
        hc = _cagr(s, HIST_START, HIST_END)
        share = float(f["kid_share"][geo])
        # this geography's own product mix (falls back to the parent's)
        try:
            gs = metrics.segment_summary(geo, "Value", pseg)
            lead = str(metrics.imax(gs["end"]))
            fast = str(metrics.imax(gs["fcagr"]))
            fast_c = float(gs["fcagr"].max())
            gain = str(metrics.imax(gs["bps"]))
        except Exception:
            gs, lead, fast = None, f["lead_sub"], f["fast_sub"]
            fast_c, gain = f["fast_sub_cagr"], f["gain_sub"]
        pg = _price_series(geo)
        p_move = _cagr(pg, EST_YEAR, END_YEAR) if pg is not None else 0.0

        drivers = []
        if rank <= 3:
            drivers.append(f"Scale base ({share * 100:.1f}% share)")
        if fc >= pace:
            drivers.append(f"Above-market pace ({fmt_pct(fc)})")
        else:
            drivers.append(f"Installed demand ({fmt_pct(fc)} CAGR)")
        drivers.append(f"{lead} demand")
        if fc > hc:
            drivers.append("Accelerating vs history")

        restraints = []
        if fc < pace:
            restraints.append(f"Trails market pace by "
                              f"{(pace - fc) * 10_000:,.0f} BPS")
        if share < 0.06:
            restraints.append("Sub-scale base")
        if p_move > 0.005:
            restraints.append(f"Rising realisation ({fmt_pct(p_move)})")
        elif p_move < -0.005:
            restraints.append(f"Price erosion ({fmt_pct(p_move)})")
        if hc > fc:
            restraints.append("Growth moderating vs history")
        if not restraints:
            restraints.append("Competitive intensity at scale")

        opps = [f"Scale {fast} ({fmt_pct(fast_c)})"]
        if gain != fast:
            opps.append(f"Mix shift to {gain}")
        opps.append(f"{fmt_num(float(s[END_YEAR] - s[EST_YEAR]))} "
                    f"{f['vlab']} headroom")

        rows.append({
            "geo": geo, "rank": rank, "share": share, "fcagr": fc,
            # capped to what one un-wrapped table cell holds at the report's
            # cell font - three columns plus the geography name fit 1250 px
            "drivers": _cap(", ".join(drivers[:3]), 62),
            "restraints": _cap(", ".join(restraints[:2]), 58),
            "opportunities": _cap(", ".join(opps[:2]), 60),
            # the same phrases in full - the region table wraps its rows, so
            # no cell ends on an ellipsis (client, 2026-10-01)
            "drivers_full": ", ".join(drivers) + ".",
            "restraints_full": ", ".join(restraints) + ".",
            "opportunities_full": ", ".join(opps) + ".",
        })

    kind = f["child_kind"]
    lead_names = ", ".join(r["geo"] for r in rows[:3])
    takeaway = (
        f"{lead_names} anchor scale with a combined "
        f"{sum(r['share'] for r in rows[:3]) * 100:.1f}% of {region} value in "
        f"{END_YEAR}, while '{f['fast_kid']}' offers the steepest upside at "
        f"{fmt_pct(f['fast_kid_cagr'])} CAGR against a market pace of "
        f"{fmt_pct(pace)}. {f['above_pace']} of {f['n_kids']} "
        f"{plural(kind, f['n_kids'])} grow at "
        f"or above that pace, so entry logic has to differ by {kind} scale "
        f"defence in the leaders, mix and premiumisation in the mid-tier, and "
        f"availability-led expansion in the fastest movers.")
    return {"rows": rows, "takeaway": takeaway, "child_kind": kind}


# ══════════════════ 2. Driving factors (radar pages) ═════════════════
def driving_factors(side: str = "supply", region: str = "Global") -> list:
    """[(factor, score 0-10)] for the supply- or demand-side radar.

    Each factor is a documented normalisation of a real workbook metric, so
    the radar's shape changes with the dataset."""
    f = facts(region)
    if side == "supply":
        return [
            # volume expansion: 4% CAGR scores 5, 0%/8% score 0/10
            ("Production Scale-up", _score(f["qfc"] * 100, 4.0, 4.0)),
            # realisation drift: flat scores 5, ±3% scores 10/0
            ("Input Cost Pressure", _score(f["pc"] * 100, 0.0, 3.0)),
            # incremental volume relative to the base
            ("Capacity Add", _score(
                ((float(f["vol"][END_YEAR] / f["vol"][EST_YEAR]) - 1.0) * 100
                 if f["vol"] is not None and float(f["vol"][EST_YEAR]) else 40.0),
                40.0, 40.0)),
            # supply concentration: top child share, 25% scores 5
            ("Supply Concentration", _score(f["top_kid_share"] * 100,
                                            25.0, 25.0)),
            # geographic spread: 12 active geographies scores 5
            ("Geographic Spread", _score(f["n_countries"] or f["n_regions"],
                                         12.0, 12.0)),
            # portfolio breadth: inverse concentration of the product segment
            ("Portfolio Breadth", _score((1.0 - f["sub_hhi"]) * 10, 5.0, 4.0)),
            # route-to-market depth: number of channels/applications
            ("Route Depth", _score(f["n_chan"], 5.0, 4.0)),
            # value realisation vs the base year
            ("Value Realisation", _score(
                ((f["p_end"] / f["p_est"] - 1.0) * 100) if f["p_est"] else 0.0,
                10.0, 20.0)),
        ]
    return [
        # value demand growth: 5% CAGR scores 5
        ("Demand Growth", _score(f["fc"] * 100, 5.0, 5.0)),
        # premiumisation = value CAGR running ahead of volume CAGR
        ("Premiumisation", _score((f["fc"] - f["qfc"]) * 100, 0.0, 3.0)),
        # application breadth of the channel/application segment
        ("App. Breadth", _score(f["n_chan"], 5.0, 4.0)),
        # emerging pull: fastest child vs the market pace
        ("Emerging Pull", _score(
            (f["fast_kid_cagr"] - f["fc"]) * 100, 1.5, 3.0)),
        # volume momentum: forecast vs historical volume growth
        ("Volume Momentum", _score((f["qfc"] - f["qhc"]) * 100, 0.0, 3.0)),
        # maturity: historical growth, 4% scores 5
        ("Category Maturity", _score(f["hc"] * 100, 4.0, 4.0)),
        # adoption depth: leading sub-segment share
        ("Adoption Depth", _score(f["lead_sub_share"] * 100, 35.0, 30.0)),
        # channel concentration eases reach when low
        ("Consumer Reach", _score((1.0 - f["chan_hhi"]) * 10, 5.0, 4.0)),
    ]


def factor_narrative(side: str, region: str = "Global") -> list:
    """Bullets explaining the radar - each cites the number behind a factor."""
    f = facts(region)
    scored = sorted(driving_factors(side, region), key=lambda t: -t[1])
    top = [n for n, _ in scored[:3]]
    lead = ", ".join(f"{n} ({v:.1f}/10)" for n, v in scored[:3])
    if side == "supply":
        return [
            f"Supply-side conditions are led by {lead} on a 0-10 scale scored "
            f"from the {region.lower()} dataset itself.",
            f"Volume expands at {fmt_pct(f['qfc'])} CAGR "
            f"({EST_YEAR}–{END_YEAR}) while average realisation moves "
            f"{fmt_pct(f['pc'])} a year the balance of those two sets the "
            f"cost-pass-through picture.",
            f"Supply is {'concentrated' if f['top_kid_share'] >= 0.35 else 'dispersed'}"
            f": '{f['top_kid']}' holds {f['top_kid_share'] * 100:.1f}% of "
            f"{END_YEAR} value across {f['n_kids']} {plural(f['child_kind'], f['n_kids'])}.",
            f"Portfolio breadth spans {f['n_sub']} {f['prod_label']} lines and "
            f"{f['n_chan']} {f['chan_label']} routes, giving producers "
            f"{'wide' if f['n_sub'] + f['n_chan'] >= 8 else 'focused'} "
            f"reformulation and channel latitude.",
            f"Capacity added between {EST_YEAR} and {END_YEAR} is worth "
            f"{fmt_num(f['incr'])} {f['vlab']}, a {f['mult']:.1f}× expansion "
            f"on the {EST_YEAR} base.",
            INDICATIVE_NOTE,
        ]
    return [
        f"Demand-side conditions are led by {lead} on a 0-10 scale scored from "
        f"the {region.lower()} dataset itself.",
        f"Value grows at {fmt_pct(f['fc'])} against volume at "
        f"{fmt_pct(f['qfc'])} a "
        f"{'premiumising' if f['fc'] > f['qfc'] else 'value-seeking'} market "
        f"through {END_YEAR}.",
        f"'{f['lead_sub']}' carries {f['lead_sub_share'] * 100:.1f}% of "
        f"{f['prod_label']} demand while '{f['fast_sub']}' compounds fastest "
        f"at {fmt_pct(f['fast_sub_cagr'])}.",
        f"Pull is strongest in '{f['fast_kid']}' "
        f"({fmt_pct(f['fast_kid_cagr'])} CAGR, "
        f"{(f['fast_kid_cagr'] - f['fc']) * 10_000:+,.0f} BPS versus the "
        f"market), "
        f"with {f['above_pace']} of {f['n_kids']} {plural(f['child_kind'], f['n_kids'])} at or "
        f"above pace.",
        f"Reach runs through {f['n_chan']} {f['chan_label']} routes, led by "
        f"'{f['lead_chan']}' at {f['lead_chan_share'] * 100:.1f}% and with "
        f"'{f['fast_chan']}' growing fastest at "
        f"{fmt_pct(f['fast_chan_cagr'])}.",
        INDICATIVE_NOTE,
    ]


# ═════════════ 3. Drivers / Restraints / Opportunities ═══════════════
def items(kind: str, region: str = "Global") -> list:
    """Three write-ups for ``kind`` ('drivers' | 'restraints' |
    'opportunities'), each with a title, a data-grounded paragraph and the
    per-period impact flags computed from its own backing series."""
    f = facts(region)
    mkt, prod, chan = f["val"], f["prod_tbl"], f["chan_tbl"]
    vlab, kidk = f["vlab"], f["child_kind"]
    out = []

    def add(title, body, series, favourable=True):
        out.append({"title": title, "body": body,
                    "gap": _impact_gap(series, mkt, favourable)})

    if kind == "drivers":
        s = prod.loc[f["fast_sub"]] if f["fast_sub"] in prod.index else mkt
        add(f"{f['fast_sub']} Compounds Fastest Across the {f['prod_label'].title()} Mix",
            f"'{f['fast_sub']}' is the fastest-compounding {f['prod_label']} "
            f"line in the {region.lower()} market, growing at "
            f"{fmt_pct(f['fast_sub_cagr'])} CAGR between {EST_YEAR} and "
            f"{END_YEAR} against a market pace of {fmt_pct(f['fc'])}. It adds "
            f"{fmt_num(float(s[END_YEAR] - s[EST_YEAR]))} {vlab} of incremental "
            f"value over the horizon and moves "
            f"{f['gain_sub_bps']:+,.0f} BPS of share where mix shifts in its "
            f"favour. With {f['n_sub']} lines competing inside "
            f"{f['prod_label']}, that concentration of growth makes it the "
            f"single clearest demand driver for the category producers "
            f"weighting capacity and promotion toward it capture a "
            f"disproportionate share of the {fmt_num(f['incr'])} {vlab} the "
            f"market adds in total.",
            s)

        ks = f["kids"].loc[f["fast_kid"]] if f["fast_kid"] in f["kids"].index else mkt
        add(f"{f['fast_kid']} Emerges as the Growth Engine Among "
            f"{plural(kidk.title())}",
            f"'{f['fast_kid']}' expands at {fmt_pct(f['fast_kid_cagr'])} CAGR "
            f"({EST_YEAR}–{END_YEAR}), running "
            f"{(f['fast_kid_cagr'] - f['fc']) * 10_000:+,.0f} BPS ahead of "
            f"the "
            f"{region.lower()} market and adding "
            f"{fmt_num(float(ks[END_YEAR] - ks[EST_YEAR]))} {vlab}. "
            f"{f['above_pace']} of {f['n_kids']} {plural(kidk, f['n_kids'])} grow at or above the "
            f"market pace, so growth is broad rather than isolated, but the "
            f"leaders compound from a materially larger base "
            f"'{f['top_kid']}' alone holds "
            f"{f['top_kid_share'] * 100:.1f}% of {END_YEAR} value. The "
            f"combination sustains volume through the forecast period and "
            f"pulls investment toward the faster {plural(kidk)}.",
            ks)

        prem = f["fc"] > f["qfc"]
        add("Premiumisation Lifts Average Realisation" if prem
            else "Volume-Led Expansion Broadens the Addressable Base",
            (f"Value compounds at {fmt_pct(f['fc'])} against volume at "
             f"{fmt_pct(f['qfc'])}, a "
             f"{(f['fc'] - f['qfc']) * 10_000:+,.0f} BPS spread that shows "
             f"demand "
             f"trading up rather than simply buying more. Average realisation "
             f"moves from {f['p_est']:,.2f} to {f['p_end']:,.2f} "
             f"{config.price_unit()} "
             f"between {EST_YEAR} and {END_YEAR} ({fmt_pct(f['pc'])} a year). "
             if prem else
             f"Volume compounds at {fmt_pct(f['qfc'])} against value at "
             f"{fmt_pct(f['fc'])}, so growth is being carried by units rather "
             f"than price average realisation moves "
             f"{fmt_pct(f['pc'])} a year to {f['p_end']:,.2f} "
             f"{config.price_unit()} by "
             f"{END_YEAR}. ") +
            f"'{f['lead_chan']}' carries "
            f"{f['lead_chan_share'] * 100:.1f}% of {f['chan_label']} demand "
            f"and '{f['fast_chan']}' grows fastest at "
            f"{fmt_pct(f['fast_chan_cagr'])}, giving suppliers a route to "
            f"convert that shift into margin across "
            f"{f['n_chan']} {f['chan_label']} routes.",
            f["price"] if f["price"] is not None else mkt)

    elif kind == "restraints":
        ss = prod.loc[f["slow_sub"]] if f["slow_sub"] in prod.index else mkt
        add(f"{f['slow_sub']} Lags the Category and Dilutes Blended Growth",
            f"'{f['slow_sub']}' grows at only {fmt_pct(f['slow_sub_cagr'])} "
            f"CAGR against the market's {fmt_pct(f['fc'])}, a "
            f"{(f['slow_sub_cagr'] - f['fc']) * 10_000:+,.0f} BPS gap, and "
            f"'{f['drop_sub']}' surrenders "
            f"{abs(f['drop_sub_bps']):,.0f} BPS of {f['prod_label']} share "
            f"between {EST_YEAR} and {END_YEAR}. Because these lines still "
            f"carry meaningful volume, the drag is absorbed by the whole "
            f"category: blended growth sits below what the leading lines "
            f"achieve on their own. Rationalising or repositioning the "
            f"lagging end of a {f['n_sub']}-line portfolio is therefore a "
            f"precondition for realising the headline forecast.",
            ss, favourable=False)

        sk = f["kids"].loc[f["slow_kid"]] if f["slow_kid"] in f["kids"].index else mkt
        below = f["n_kids"] - f["above_pace"]
        add(f"Uneven {kidk.title()} Performance Caps the Blended Rate",
            f"'{f['slow_kid']}' grows at just "
            f"{fmt_pct(f['slow_kid_cagr'])} CAGR "
            f"{(f['slow_kid_cagr'] - f['fc']) * 10_000:+,.0f} BPS against "
            f"the "
            f"market and {below} of {f['n_kids']} {plural(kidk, f['n_kids'])} sit below the "
            f"market pace altogether. Demand is therefore not uniformly "
            f"available: suppliers carrying a broad geographic footprint fund "
            f"presence in slow {plural(kidk)} out of returns earned in the fast ones. "
            f"With '{f['top_kid']}' at {f['top_kid_share'] * 100:.1f}% of "
            f"value, a shortfall in a single leading {kidk} is difficult to "
            f"offset elsewhere in the forecast window.",
            sk, favourable=False)

        conc = f["top_kid_share"] >= 0.30 or f["sub_hhi"] >= 0.30
        add("Concentration Risk Limits Pricing and Sourcing Flexibility" if conc
            else "Fragmented Demand Raises Cost-to-Serve",
            (f"The market is concentrated: '{f['top_kid']}' accounts for "
             f"{f['top_kid_share'] * 100:.1f}% of {END_YEAR} value and the "
             f"leading {f['prod_label']} line for "
             f"{f['lead_sub_share'] * 100:.1f}% of its segment. "
             if conc else
             f"The market is fragmented: no {kidk} exceeds "
             f"{f['top_kid_share'] * 100:.1f}% of {END_YEAR} value and demand "
             f"spreads across {f['n_sub']} {f['prod_label']} lines and "
             f"{f['n_chan']} {f['chan_label']} routes. ") +
            f"Average realisation moves {fmt_pct(f['pc'])} a year, so cost "
            f"movements cannot be passed on evenly across "
            f"{f['n_kids']} {plural(kidk, f['n_kids'])}. The practical constraint is the same in "
            f"either case commercial terms set for the market as a whole "
            f"underserve part of it, and margin is lost in the mismatch "
            f"rather than in the headline growth rate.",
            f["price"] if f["price"] is not None else mkt, favourable=False)

    else:                                              # opportunities
        add(f"{f['opp_geo']} Carries the Largest Absolute Opportunity",
            f"'{f['opp_geo']}' adds {fmt_num(f['opp_abs'])} {vlab} between "
            f"{EST_YEAR} and {END_YEAR} the single largest absolute "
            f"increment of any geography in the dataset. Capturing it is a "
            f"question of scale execution rather than market creation: demand "
            f"already exists and grows at "
            f"{fmt_pct(f['fc'])} at the market level. Prioritising supply, "
            f"channel coverage and service levels there converts the largest "
            f"share of the {fmt_num(f['incr'])} {vlab} the whole market adds "
            f"over the forecast period.",
            _or(_series(f["opp_geo"], "Value"), mkt))

        add(f"{f['under_geo']} Offers Under-penetrated, Above-pace Headroom",
            f"'{f['under_geo']}' compounds at "
            f"{fmt_pct(f['under_cagr'])} CAGR from a below-median base one "
            f"of {max(f['n_under'], 1)} geographies in the dataset combining "
            f"above-median growth with below-median size. That profile is "
            f"where share is cheapest to buy: competitive intensity is lower "
            f"than in the {f['top_kid_share'] * 100:.1f}%-share leaders, while "
            f"growth runs "
            f"{(f['under_cagr'] - f['fc']) * 10_000:+,.0f} BPS ahead of the "
            f"market. "
            f"Early distribution and reference accounts established here "
            f"compound through {END_YEAR}.",
            _or(_series(f["under_geo"], "Value"), mkt))

        cs = chan.loc[f["fast_chan"]] if f["fast_chan"] in chan.index else mkt
        add(f"{f['fast_chan']} Opens the Fastest Route to Market",
            f"'{f['fast_chan']}' grows at "
            f"{fmt_pct(f['fast_chan_cagr'])} CAGR, the fastest of "
            f"{f['n_chan']} {f['chan_label']} routes, and adds "
            f"{fmt_num(float(cs[END_YEAR] - cs[EST_YEAR]))} {vlab} over the "
            f"forecast period while '{f['lead_chan']}' still carries "
            f"{f['lead_chan_share'] * 100:.1f}% of demand today. Pairing the "
            f"incumbent route for volume with the fastest route for growth "
            f"and leading with '{f['fast_sub']}' at "
            f"{fmt_pct(f['fast_sub_cagr'])} is the highest-yield "
            f"combination the dataset supports.",
            cs)
    _grade_impacts(out)          # raw gaps -> High/Medium/Low per period
    return out


# ══════════════════════ 4. Milestone / development table ═════════════
def milestones(region: str = "Global") -> list:
    """[(year label, development text)] derived from the market's own curve -
    the strongest historical year, the base year, the crossover year, the
    mid-horizon checkpoint and the forecast outcome."""
    f = facts(region)
    val, vlab = f["val"], f["vlab"]
    if val is None:
        return []
    out = []
    hist = [y for y in YEARS if y <= HIST_END]
    yoy = val.pct_change()
    if len(hist) > 1:
        hy = max(hist[1:], key=lambda y: (yoy[y] if yoy[y] == yoy[y] else -9))
        out.append((str(hy),
                    f"Strongest historical expansion on record: {region} value "
                    f"rose {yoy[hy] * 100:+.1f}% year on year to "
                    f"{fmt_num(float(val[hy]))} {vlab}, the steepest single-year "
                    f"move of the {HIST_START}–{HIST_END} period. Demand for "
                    f"'{f['lead_sub']}' {f['lead_sub_share'] * 100:.1f}% of "
                    f"{f['prod_label']} value by {END_YEAR} set the pace, and "
                    f"the {f['n_chan']} {f['chan_label']} routes tracked here "
                    f"all carried volume through the year."))
    out.append((str(EST_YEAR),
                f"Base year for the forecast: the {region.lower()} market is "
                f"estimated at {fmt_num(f['est'])} {vlab}, having compounded at "
                f"{fmt_pct(f['hc'])} across {HIST_START}–{HIST_END}. "
                f"'{f['top_kid']}' holds {f['top_kid_share'] * 100:.1f}% of "
                f"value across {f['n_kids']} {plural(f['child_kind'], f['n_kids'])}, and the "
                f"forecast is built from {f['n_seg']} segment breakdowns "
                f"covering {f['n_sub']} {f['prod_label']} lines."))
    # crossover: first forecast year past 1.5x the base
    tgt = f["est"] * 1.5
    cross = next((y for y in YEARS if y >= EST_YEAR and float(val[y]) >= tgt),
                 None)
    if cross and cross < END_YEAR:
        out.append((str(cross),
                    f"The market passes {fmt_num(tgt)} {vlab} one and a half "
                    f"times the {EST_YEAR} base reaching "
                    f"{fmt_num(float(val[cross]))} {vlab}. Growth by this point "
                    f"is led by '{f['fast_sub']}' at "
                    f"{fmt_pct(f['fast_sub_cagr'])} CAGR and by "
                    f"'{f['fast_kid']}' among {plural(f['child_kind'])} at "
                    f"{fmt_pct(f['fast_kid_cagr'])}, with "
                    f"{f['above_pace']} of {f['n_kids']} {plural(f['child_kind'], f['n_kids'])} "
                    f"running at or above the market rate."))
    mid = YEARS[(YEARS.index(EST_YEAR) + len(YEARS) - 1) // 2] \
        if EST_YEAR in YEARS else YEARS[len(YEARS) // 2]
    if mid not in (EST_YEAR, END_YEAR, cross):
        out.append((str(mid),
                    f"Mid-horizon checkpoint: value reaches "
                    f"{fmt_num(float(val[mid]))} {vlab}, "
                    f"{float(val[mid]) / f['est']:.2f}× the {EST_YEAR} base. "
                    f"Average realisation stands at {'%.2f' % float(f['price'][mid]) if f['price'] is not None else 'n/a'} "
                    f"{config.price_unit()}, and mix has moved "
                    f"{f['gain_sub_bps']:+,.0f} BPS toward '{f['gain_sub']}' "
                    f"while '{f['drop_sub']}' gives up "
                    f"{abs(f['drop_sub_bps']):,.0f} BPS."))
    out.append((str(END_YEAR),
                f"Forecast outcome: the {region.lower()} market reaches "
                f"{fmt_num(f['end'])} {vlab}, a {f['mult']:.1f}× expansion on "
                f"{EST_YEAR} and {fmt_num(f['incr'])} {vlab} of incremental "
                f"value at {fmt_pct(f['fc'])} CAGR. Volume compounds at "
                f"{fmt_pct(f['qfc'])} over the same window, leaving "
                f"realisation {fmt_pct(f['pc'])} a year the gap between the "
                f"two is the premiumisation the forecast assumes."))
    return out[:5]


# ══════════════════════ 5. Porter's Five Forces ══════════════════════
def _rating(score: float) -> str:
    return ("Low" if score < 3.6 else "Moderate" if score < 5.6
            else "Moderate to High" if score < 7.4 else "High")


def five_forces(region: str = "Global") -> list:
    """Five forces scored 0-10 from the workbook, each with its own evidence."""
    f = facts(region)
    kidk = f["child_kind"]
    pc, fc = f["pc"], f["fc"]

    supplier = _score(f["top_kid_share"] * 100, 25.0, 22.0) * 0.6 + \
        _score(pc * 100, 0.0, 3.0) * 0.4
    buyer = _score((1.0 - f["chan_hhi"]) * 10, 5.0, 4.0) * 0.6 + \
        _score(-pc * 100, 0.0, 3.0) * 0.4
    rivalry = _score((1.0 - f["sub_hhi"]) * 10, 5.0, 3.5) * 0.6 + \
        _score(f["n_sub"], 5.0, 4.0) * 0.4
    substitute = _score(abs(f["drop_sub_bps"]) / 100.0, 2.0, 2.0) * 0.6 + \
        _score(-(f["slow_sub_cagr"] * 100), -2.0, 4.0) * 0.4
    entrant = _score(fc * 100, 5.0, 5.0) * 0.5 + \
        _score((1.0 - f["kid_hhi"]) * 10, 5.0, 3.5) * 0.5

    forces = [
        ("Bargaining Power of Suppliers", supplier, [
            f"'{f['top_kid']}' supplies {f['top_kid_share'] * 100:.1f}% of "
            f"{END_YEAR} value across {f['n_kids']} {plural(kidk, f['n_kids'])}, concentrating "
            f"upstream leverage.",
            f"Average realisation moves {fmt_pct(pc)} a year "
            f"({EST_YEAR}–{END_YEAR}), indicating "
            f"{'input costs are being passed through' if pc > 0 else 'limited pass-through'}.",
            f"Buyers can rebalance across {f['n_sub']} {f['prod_label']} "
            f"lines, which caps supplier leverage in commodity grades.",
        ]),
        ("Bargaining Power of Buyers", buyer, [
            f"Demand routes through {f['n_chan']} {f['chan_label']} channels "
            f"led by '{f['lead_chan']}' at "
            f"{f['lead_chan_share'] * 100:.1f}%, giving large buyers scale.",
            f"Price sensitivity tracks realisation moving {fmt_pct(pc)} a year "
            f"against volume growth of {fmt_pct(f['qfc'])}.",
            f"Switching is easiest where {f['n_sub']} comparable "
            f"{f['prod_label']} lines compete for the same demand.",
        ]),
        ("Intensity of Competition", rivalry, [
            f"{f['n_sub']} {f['prod_label']} lines compete inside the segment; "
            f"the leader holds {f['lead_sub_share'] * 100:.1f}%.",
            f"Growth is unevenly distributed '{f['fast_sub']}' at "
            f"{fmt_pct(f['fast_sub_cagr'])} against '{f['slow_sub']}' at "
            f"{fmt_pct(f['slow_sub_cagr'])}.",
            f"{f['above_pace']} of {f['n_kids']} {plural(kidk, f['n_kids'])} grow at or above the "
            f"{fmt_pct(fc)} market pace, keeping share contested.",
        ]),
        ("Threat of Substitutes", substitute, [
            f"'{f['drop_sub']}' loses "
            f"{abs(f['drop_sub_bps']):,.0f} BPS of share to "
            f"'{f['gain_sub']}' ({f['gain_sub_bps']:+,.0f} BPS) by "
            f"{END_YEAR} substitution is visible inside the mix.",
            f"The slowest line grows at {fmt_pct(f['slow_sub_cagr'])}, "
            f"{(f['slow_sub_cagr'] - fc) * 10_000:+,.0f} BPS against the "
            f"market.",
            f"Realisation moving {fmt_pct(pc)} a year "
            f"{'widens' if pc > 0 else 'narrows'} the price gap substitutes "
            f"must close.",
        ]),
        ("Threat of New Entrants", entrant, [
            f"The market grows at {fmt_pct(fc)} CAGR and adds "
            f"{fmt_num(f['incr'])} {f['vlab']} by {END_YEAR}, attracting entry.",
            f"Demand spreads over {f['n_kids']} {plural(kidk, f['n_kids'])} with no single one "
            f"above {f['top_kid_share'] * 100:.1f}%, so regional entry is "
            f"feasible.",
            f"Incumbent scale in '{f['top_kid']}' and "
            f"'{f['lead_sub']}' ({f['lead_sub_share'] * 100:.1f}% of segment) "
            f"remains the principal barrier.",
        ]),
    ]
    out = []
    for i, (n, s, b) in enumerate(forces):
        rating = _rating(_clamp(s))
        # the reference deck leads each column with its assessment inside the
        # first bullet rather than as a separate score, so the derived rating
        # stays visible without putting a number on the slide
        bullets = [f"Assessed {rating.lower()}. {b[0]}"] + list(b[1:])
        out.append({"name": n, "score": round(float(_clamp(s)), 1),
                    "rating": rating, "bullets": bullets,
                    "color": FORCE_COLORS[i % len(FORCE_COLORS)],
                    "edge": FORCE_EDGES[i % len(FORCE_EDGES)],
                    "icon": FORCE_ICONS[i % len(FORCE_ICONS)],
                    "ink": FORCE_INK[i % len(FORCE_INK)]})
    return out


def forces_subtitle(region: str = "Global") -> str:
    f = facts(region)
    return (f"Supplier power, buyer power, competitive intensity, substitution "
            f"and entry barriers scored from the {region.lower()} dataset "
            f"{f['n_sub']} {f['prod_label']} lines, "
            f"{f['n_chan']} {f['chan_label']} routes and "
            f"{f['n_kids']} {plural(f['child_kind'], f['n_kids'])} at {fmt_pct(f['fc'])} CAGR.")
