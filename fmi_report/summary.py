"""Dynamic content for the Summary-section slides.

Every string and number here is derived from the loaded workbook via the
existing data / metrics pipeline - there are NO hardcoded market names,
company names, drivers, values or scenarios.  Drop in a different FMI
workbook and all of it re-derives.  Qualitative sections (drivers, restraints,
supply-side activity, recommendations) are phrased as generic, data-grounded
statements built from the quantitative facts (fastest/largest segment and
region, share shifts, CAGR, incremental opportunity), never invented figures.
"""
from __future__ import annotations

import hashlib
import random

import pandas as pd

from . import data, geo, metrics
from . import trend_patterns as _tp
from .components import fmt_num, fmt_pct
from .config import (END_YEAR, EST_YEAR, FORE_YEARS, HIST_END, HIST_START,
                     UNIT_LABELS, plural)

MK = data.MARKET                                   # e.g. "Matcha"


def _unit() -> str:
    for u in ("Value", "Volume"):
        try:
            s = data.total_series(geo.current(), u)
            if float(pd.Series(s).fillna(0).abs().sum()) > 0:
                return u
        except Exception:
            continue
    return "Value"


def _pick_segment(keywords, fallback_idx=0, unit=None, exclude=()):
    """Pick a segment matching `keywords`.

    With `unit` given, only segments that actually carry rows in that unit are
    eligible: a workbook can size one breakdown in value and another in volume,
    and these summary pages read one unit at a time.

    A segment that also mentions one of ``exclude`` is skipped even when it
    matches ``keywords`` too, so a channel/end-use segment named e.g.
    "Buyer Type / End-User" is not picked for the PRODUCT segment just
    because "type" is one of its words (2026-10-22: see dynamics.py's twin
    of this helper, _pick_segment, for the full story).
    """
    def has_data(seg):
        if unit is None:
            return True
        try:
            t = data.segment_table(geo.current(), unit, seg)
        except Exception:
            return False
        return bool(len(t)) and bool(float(t.abs().to_numpy().sum()))

    usable = [x for x in data.SEGMENTS if has_data(x)] or list(data.SEGMENTS)
    for x in usable:
        low = x.lower()
        if exclude and any(k in low for k in exclude):
            continue
        if any(k in low for k in keywords):
            return x
    for x in usable:
        if any(k in x.lower() for k in keywords):
            return x
    return usable[min(fallback_idx, len(usable) - 1)] if usable else None


def _label(seg):                                   # "By Product Type" -> "product type"
    return (seg or "").replace("By ", "").strip().lower() or "segment"


def facts() -> dict:
    """All the numbers the Summary slides draw on, computed once."""
    u = _unit()
    f = {"unit": u, "ulab": UNIT_LABELS[u]}
    g = data.total_series(geo.current(), u)
    f["est"], f["end"] = float(g[EST_YEAR]), float(g[END_YEAR])
    f["fc"], f["hc"] = metrics.fore_cagr(g), metrics.hist_cagr(g)
    f["incr"] = f["end"] - f["est"]
    f["mult"] = (f["end"] / f["est"]) if f["est"] else float("nan")
    f["series"] = g

    ps = _pick_segment(
        ["product", "offering", "type", "formulation", "grade"], unit=u,
        exclude=["channel", "distribution", "application", "end use",
                 "end-use"])
    f["prod_seg"], f["prod_label"] = ps, _label(ps)
    ss = metrics.segment_summary(geo.current(), u, ps)
    f["top_sub"] = metrics.imax(ss["end"])
    f["top_sub_share"] = ss.loc[f["top_sub"], "share_end"]
    f["fast_sub"] = metrics.imax(ss["fcagr"])
    f["fast_sub_cagr"] = ss.loc[f["fast_sub"], "fcagr"]
    f["slow_sub"] = metrics.imin(ss["fcagr"])
    f["slow_sub_cagr"] = ss.loc[f["slow_sub"], "fcagr"]
    f["n_sub"] = len(ss)
    f["n_seg"] = len(data.SEGMENTS)
    f["seg_summary"] = ss

    ap = _pick_segment(["application", "end use", "end-use", "channel",
                        "distribution"], unit=u,
                       fallback_idx=min(1, len(data.SEGMENTS) - 1))
    f["app_seg"], f["app_label"] = ap, _label(ap)
    sa = metrics.segment_summary(geo.current(), u, ap)
    f["fast_app"] = metrics.imax(sa["fcagr"])
    f["fast_app_cagr"] = sa.loc[f["fast_app"], "fcagr"]

    kids = data.children_table(geo.current(), u)
    end_by = kids[END_YEAR].astype(float)
    tot = float(end_by.sum())
    f["top_region"] = metrics.imax(end_by)
    f["top_region_share"] = (end_by[f["top_region"]] / tot) if tot else float("nan")
    rc = {r: metrics.fore_cagr(kids.loc[r]) for r in kids.index}
    f["fast_region"] = max(rc, key=lambda r: (rc[r] if not pd.isna(rc[r])
                                              else -9))
    f["fast_region_cagr"] = rc[f["fast_region"]]
    f["n_regions"] = len(kids)

    sh = f["top_sub_share"]
    f["concentration"] = ("consolidated" if (not pd.isna(sh) and sh >= 0.40)
                          else "moderately fragmented"
                          if (not pd.isna(sh) and sh >= 0.20) else "fragmented")
    return f


# ── Slide 1: Market Overview ─────────────────────────────────────────
def market_overview(f):
    u, ul = f["unit"], f["ulab"]
    return {
        "Market Overview": [
            f"The {geo.adj()} {MK} market is valued at {fmt_num(f['est'])} {ul} in "
            f"{EST_YEAR} and is forecast to reach {fmt_num(f['end'])} {ul} by "
            f"{END_YEAR}.",
            f"That implies a {fmt_pct(f['fc'])} forecast CAGR "
            f"({EST_YEAR}–{END_YEAR}), versus {fmt_pct(f['hc'])} over "
            f"{HIST_START}–{HIST_END}.",
            f"An incremental {fmt_num(f['incr'])} {ul} of demand is created "
            f"over the forecast a {f['mult']:.1f}× expansion.",
        ],
        "Competitive Landscape": [
            f"The market is {f['concentration']}: the leading {f['prod_label']} "
            f"category '{f['top_sub']}' holds {fmt_pct(f['top_sub_share'])} of "
            f"{END_YEAR} {u.lower()}.",
            f"Demand spans {f['n_regions']} "
            f"{plural(geo.child_kind(), f['n_regions'])} and "
            f"{f['n_seg']} "
            f"segmentation lenses ({f['n_sub']} {f['prod_label']} categories).",
            f"{f['top_region']} is the largest {geo.child_adj()} market at "
            f"{fmt_pct(f['top_region_share'])} of {END_YEAR} {u.lower()}.",
        ],
        "Drivers": [
            f"Accelerating uptake of '{f['fast_sub']}' "
            f"({fmt_pct(f['fast_sub_cagr'])} CAGR).",
            f"Fast-growing demand in {f['fast_region']} "
            f"({fmt_pct(f['fast_region_cagr'])}).",
            f"Momentum in {f['fast_app']} within {f['app_label']} "
            f"({fmt_pct(f['fast_app_cagr'])}).",
        ],
        "Restraints": [
            f"Slower momentum in '{f['slow_sub']}' "
            f"({fmt_pct(f['slow_sub_cagr'])}).",
            f"Maturing demand in established categories caps the blended "
            f"{fmt_pct(f['fc'])} CAGR.",
            f"Value concentration in '{f['top_sub']}' raises barriers for "
            f"new entrants.",
        ],
        "Trends": [
            f"Mix shifting toward '{f['fast_sub']}' and higher-growth tiers.",
            f"{f['fast_region']} outpacing the {geo.adj()} average.",
            f"Demand broadening across {f['n_sub']} {f['prod_label']} "
            f"categories.",
        ],
    }


# ── Slide 2: Executive Summary - Supply Side Trends (2/3) ─────────────
def supply_side(f):
    u = f["unit"]
    return {
        "Recent Developments & Key Industry Activities": [
            f"Capacity and product development are concentrating in "
            f"'{f['fast_sub']}' and {f['fast_region']}.",
            f"Portfolio expansion tracks the shift toward higher-growth "
            f"{f['prod_label']} categories.",
            f"Commercial focus is intensifying across {f['top_region']} and "
            f"{f['fast_region']}.",
        ],
        "Strategic Acquisitions & New Investments": [
            f"Investment is flowing toward the fastest-growing tier "
            f"'{f['fast_sub']}' ({fmt_pct(f['fast_sub_cagr'])}).",
            f"{f['fast_region']} ({fmt_pct(f['fast_region_cagr'])}) is a "
            f"priority {geo.child_kind()} for new capacity.",
            f"Scale players anchor around '{f['top_sub']}' "
            f"({fmt_pct(f['top_sub_share'])} share).",
        ],
        "Substantial Investment in R&D": [
            f"R&D aligns with the migration to '{f['fast_sub']}' and adjacent "
            f"premium tiers.",
            f"Development spend targets {f['prod_label']} categories gaining "
            f"share to {END_YEAR}.",
            f"Innovation underpins the {fmt_pct(f['fc'])} forecast CAGR.",
        ],
        "Key Insights": [
            f"{EST_YEAR} size {fmt_num(f['est'])} {f['ulab']}; {END_YEAR} "
            f"{fmt_num(f['end'])} {f['ulab']}.",
            f"Forecast CAGR {fmt_pct(f['fc'])} ({EST_YEAR}–{END_YEAR}).",
            f"Lead category '{f['top_sub']}' {fmt_pct(f['top_sub_share'])} "
            f"of {END_YEAR}.",
            f"Lead {geo.child_kind()} {f['top_region']} "
            f"{fmt_pct(f['top_region_share'])}.",
            f"Incremental opportunity {fmt_num(f['incr'])} {f['ulab']}.",
        ],
    }


# ── Slide 3: Executive Summary - Analysis & Recommendation (3/3) ──────
def analysis(f):
    return {
        "Overall Market Approach": [
            f"Balance scale in '{f['top_sub']}' with targeted growth in "
            f"'{f['fast_sub']}', pursuing the {fmt_pct(f['fc'])} CAGR "
            f"trajectory to {END_YEAR}.",
        ],
        "Differentiating Strategy": [
            f"Differentiate through the high-growth '{f['fast_sub']}' tier and "
            f"{f['fast_app']} within {f['app_label']}.",
        ],
        "Market Segment": [
            f"Prioritise '{f['fast_sub']}' ({fmt_pct(f['fast_sub_cagr'])}); "
            f"'{f['top_sub']}' anchors {fmt_pct(f['top_sub_share'])} of value.",
        ],
        # NB: a KEY, not a heading - templates_main looks it up by this exact
        # string and supplies its own (geography-aware) title.
        "Market Region": [
            f"Lead with {f['top_region']} ({fmt_pct(f['top_region_share'])}); "
            f"capture upside in {f['fast_region']} "
            f"({fmt_pct(f['fast_region_cagr'])}).",
        ],
        "Fact.MR Analysis & Recommendation": [
            f"{MK} adds {fmt_num(f['incr'])} {f['ulab']} of demand to "
            f"{END_YEAR} a {f['mult']:.1f}× expansion.",
            f"Concentrate resources on the leading "
            f"segment/{geo.child_kind()} while "
            f"seeding the fastest-growing pockets early.",
            f"Defend share in '{f['top_sub']}' as the value anchor of the "
            f"portfolio.",
        ],
    }


# ── Slide 4: Scenario Outlook ────────────────────────────────────────
def _pattern_rates(ends, base, g, band):
    """{scenario: {year: yoy}} with each scenario following its own curve
    from `trend_patterns`, calibrated to compound from ``base`` to the
    scenario's end value; None when the data cannot support it."""
    try:
        if not base or pd.isna(base) or base <= 0:
            return None
        fc = abs(float(g[END_YEAR] / g[EST_YEAR]) ** (1.0 / (END_YEAR - EST_YEAR))
                 - 1.0)
    except Exception:
        return None
    gap = max(0.003, band * fc)            # spacing between scenario lines
    seed = int(hashlib.md5(
        f"{data.MARKET}|{geo.current()}".encode()).hexdigest()[:8], 16)
    names = random.Random(seed).sample(sorted(_tp.PATTERNS), 3)
    out = {}
    for scen, nm in zip(("Optimistic", "Likely", "Conservative"), names):
        curve = _tp.PATTERNS[nm]
        pts = [curve[min(max(5 + (y - EST_YEAR), 0), len(curve) - 1)]
               for y in FORE_YEARS]
        mean = sum(pts) / len(pts)
        dev = [v - mean for v in pts]
        peak = max(abs(d) for d in dev) or 1.0
        k = min(1.0, 0.45 * gap / peak)    # shape kept, swing kept inside gap
        dev = [d * k for d in dev]
        target = ends[scen] / base
        if target <= 0:
            return None

        def grown(shift):
            t = 1.0
            for d in dev:
                t *= max(1e-6, 1.0 + d + shift)
            return t
        lo, hi = -0.95, 2.0
        for _ in range(60):                 # growth is monotone in the shift
            mid = (lo + hi) / 2.0
            if grown(mid) < target:
                lo = mid
            else:
                hi = mid
        shift = (lo + hi) / 2.0
        out[scen] = {y: d + shift for y, d in zip(FORE_YEARS, dev)}
    return out


def scenarios(f):
    """Likely = the base forecast.  Optimistic / Conservative flex the base
    forecast's YEAR-ON-YEAR growth by a fixed ±25% band (a generic,
    market-agnostic spread), every year from EST_YEAR to END_YEAR, all three
    starting from the same last historical year (HIST_END).  The level
    trajectories, the end values and the CAGRs are all derived from those
    yearly rates, so the bar chart, the growth lines and the KPI cards can
    never disagree."""
    band = 0.25
    horizon = END_YEAR - EST_YEAR
    g = f["series"]
    base = float(g[HIST_END]) if HIST_END in g.index else float("nan")
    likely = {y: float(g[y]) for y in FORE_YEARS}

    yoy_l = {}
    prev = base
    for y in FORE_YEARS:
        yoy_l[y] = ((likely[y] / prev - 1.0) if prev and not pd.isna(prev)
                    else 0.0)
        prev = likely[y]
    # the band widens growth either way by the same share of its size, so a
    # declining market's optimistic case declines LESS, not more
    yoy = {"Optimistic": {y: r + band * abs(r) for y, r in yoy_l.items()},
           "Likely": yoy_l,
           "Conservative": {y: r - band * abs(r) for y, r in yoy_l.items()}}

    def project(rates):
        out, lvl = {}, base
        for y in FORE_YEARS:
            lvl = lvl * (1.0 + rates[y])
            out[y] = lvl
        return out

    traj = pd.DataFrame([project(yoy["Optimistic"]), likely,
                         project(yoy["Conservative"])],
                        index=["Optimistic", "Likely", "Conservative"])
    ends = {k: float(traj.loc[k, END_YEAR]) for k in traj.index}
    # The three year-on-year lines follow growth PATTERNS (trend_patterns:
    # the client's sheet of curves) instead of three flat parallel lines.
    # Each scenario takes its own curve, its LEVEL re-calibrated so that
    # compounding it from the base year lands on exactly the end value above
    # - the bars, cards and CAGRs are unchanged, only the path between moves.
    shaped = _pattern_rates(ends, base, g, band)
    if shaped:
        yoy = shaped
        traj = pd.DataFrame([project(yoy["Optimistic"]),
                             project(yoy["Likely"]),
                             project(yoy["Conservative"])],
                            index=["Optimistic", "Likely", "Conservative"])
        ends = {k: float(traj.loc[k, END_YEAR]) for k in traj.index}
    cagrs = {}
    for k in traj.index:
        a, b = float(traj.loc[k, EST_YEAR]), float(traj.loc[k, END_YEAR])
        cagrs[k] = ((b / a) ** (1.0 / horizon) - 1.0
                    if a > 0 and b > 0 and horizon else float("nan"))
    cagrs["Likely"] = f["fc"]          # the report's own forecast CAGR
    return {"traj": traj, "ends": ends, "cagrs": cagrs, "yoy": yoy,
            "base": base, "base_year": HIST_END,
            "band": band, "horizon": horizon}
