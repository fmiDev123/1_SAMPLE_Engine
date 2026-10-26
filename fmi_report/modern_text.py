"""Overview-page write-up from the modern template library (fmi_templates).

Covers the three executive-summary pages that used to carry short, generic,
hard-coded sentences:

    G07  Supply Side Trends            -> developments / insights / invest / rd
    G08  Analysis and Recommendation   -> approach / differentiate / segment / region / reco
    G09  Scenario Outlook              -> readout bullets, chart titles, card captions

The wording lives in `fmi_templates/slide_templates.py`; this module only builds
the template context from the engine's own numbers (`summary.facts()` and
`summary.scenarios()`), renders it with `render_slide(mask=False)` and runs
`qc_slide` on the result.  No number or name is typed here.

Contract with the call sites (same as `classic_text`): every public function
returns the rendered dict or **None** - a missing value, an unknown industry,
a QC finding - and the caller then keeps the text it had before, so a template
can never break an export.  The lock pass withholds the figures afterwards, as
on every other page, so the deck is rendered unmasked.

Industry comes from FMI_INDUSTRY (food | consumer | healthcare | pharma |
industrial); without it the library's lexicon phrases cannot be chosen and the
old text is kept.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

_ROOT = str(Path(__file__).resolve().parents[1])
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from fmi_templates.engine import (BANNED_BY_INDUSTRY, build_ctx,      # noqa: E402
                                  qc_slide, render_slide)

from . import classic_text as _ct                                      # noqa: E402
from . import data, geo, metrics                                       # noqa: E402
from .components import fmt_num, fmt_pct                               # noqa: E402
from .config import END_YEAR, EST_YEAR                                 # noqa: E402

LAST_ISSUES: list = []           # what the last refused slide failed on


def enabled() -> bool:
    return _ct.enabled() and _ct.industry() is not None


def _num(x) -> bool:
    return x is not None and not pd.isna(x)


def _ctx(f: dict) -> dict | None:
    """Template context for the active geography from `summary.facts()`."""
    if f.get("unit") != "Value":
        return None                 # the templates are written for value
    need = ("est", "end", "fc", "incr", "mult", "top_sub", "top_sub_share",
            "fast_sub", "fast_sub_cagr", "fast_app", "top_region",
            "top_region_share", "fast_region", "fast_region_cagr")
    if any(k not in f for k in need):
        return None
    for k in ("est", "end", "fc", "incr", "mult", "top_sub_share",
              "fast_sub_cagr", "top_region_share", "fast_region_cagr"):
        if not _num(f[k]):
            return None

    base = _ct.base_ctx()
    c = {k: v for k, v in base.items()
         if k not in ("drivers", "restraints", "trends")}
    kind = geo.child_kind()
    both = None
    try:                            # lead + fastest child, together
        kids = data.children_table(geo.current(), "Value")
        by = kids[END_YEAR].astype(float)
        tot = float(by.sum())
        names = {f["top_region"], f["fast_region"]}
        if tot > 0:
            both = float(sum(by[n] for n in names if n in by.index)) / tot
    except Exception:
        both = None
    if both is None:
        return None

    c.update(
        source=f"FMI Analysis, {EST_YEAR}",
        fc_period=f"{EST_YEAR}-{END_YEAR}",
        child_label=kind, child_label_pl=kind[:-1] + "ies" if kind.endswith("y")
        else kind + "s", child_label_cap=kind.title(),
        seg1=str(f["prod_label"]).title(), seg1_lc=_ct._lc(f["prod_label"]),
        seg1_lead=f["top_sub"], seg1_fast=f["fast_sub"],
        seg2=f["app_label"], seg2_lc=_ct._lc(f["app_label"]),
        seg2_fast=f["fast_app"],
        child_lead=f["top_region"], child_fast=f["fast_region"],
        v_val_est=fmt_num(f["est"]), v_val_end=fmt_num(f["end"]),
        v_f_cagr=fmt_pct(f["fc"]), v_abs_opp=fmt_num(f["incr"]),
        v_mult=f"{f['mult']:.1f}",
        v_seg1_lead_share_end=fmt_pct(f["top_sub_share"]),
        v_seg1_fast_cagr=fmt_pct(f["fast_sub_cagr"]),
        v_child_lead_share_end=fmt_pct(f["top_region_share"]),
        v_child_fast_cagr=fmt_pct(f["fast_region_cagr"]),
        v_child_lead_fast_share=fmt_pct(both),
        _whole=both >= 0.995,
        r_f_cagr=float(f["fc"]),
        val_unit=f["ulab"], val_word=c.get("val_word", ""))
    return c


def _render(sid: str, c: dict) -> dict | None:
    """Render, QC and return the slide's elements, or None (reason in
    LAST_ISSUES)."""
    LAST_ISSUES.clear()
    try:
        ctx = build_ctx(c)
        out = render_slide(sid, ctx, mask=False)
        banned = BANNED_BY_INDUSTRY.get(c.get("industry"), ())
        issues = qc_slide(out, mask=False, banned_terms=banned)
    except Exception as e:                      # KeyError, MissingKeys ...
        LAST_ISSUES.append((sid, "render", repr(e)[:120]))
        return None
    if issues:
        LAST_ISSUES.extend(issues)
        return None
    return out


def _guard(fn):
    def run(*a, **kw):
        if not enabled():
            return None
        try:
            return fn(*a, **kw)
        except Exception as e:
            LAST_ISSUES[:] = [(fn.__name__, "error", repr(e)[:120])]
            return None
    run.__name__ = fn.__name__
    return run


# ── G07  Supply Side Trends ──────────────────────────────────────────
@_guard
def supply_side(f: dict) -> dict | None:
    """{panel title: [sentences]} for the four Supply Side boxes, plus
    'chart_title'; None to keep the old text."""
    c = _ctx(f)
    if c is None:
        return None
    out = _render("G07", c)
    if out is None:
        return None
    if c.pop("_whole", False):
        # two children make up the whole market: "together account for
        # 100.0% of value" says nothing, so the sentence names the fact
        kind = c["child_label_pl"]
        out["developments"] = [
            (f"Commercial focus is intensifying in {c['child_lead']} and "
             f"{c['child_fast']}, the two {kind} that make up the "
             f"{c['geo']} market.") if "together account for" in t else t
            for t in out["developments"]]
    return {
        "Recent Developments & Key Industry Activities": out["developments"],
        "Key Insights": out["insights"],
        "Strategic Acquisitions & New Investments": out["invest"],
        "Substantial Investment in R&D": out["rd"],
        "_chart_title": out["chart_title"],
    }


# ── G08  Analysis and Recommendation ─────────────────────────────────
@_guard
def analysis(f: dict) -> dict | None:
    """{key: [sentences]} using the page's existing keys (see
    `summary.analysis`) plus 'reco_head', the heading of the last box."""
    c = _ctx(f)
    if c is None:
        return None
    out = _render("G08", c)
    if out is None:
        return None
    return {
        "Overall Market Approach": out["approach"],
        "Differentiating Strategy": out["differentiate"],
        "Market Segment": out["segment"],
        "Market Region": out["region"],
        "Fact.MR Analysis & Recommendation": out["reco"],
        "_reco_head": out["reco_head"],
    }


# ── G09  Scenario Outlook ────────────────────────────────────────────
@_guard
def scenario(f: dict, sc: dict) -> dict | None:
    """Readout bullets and chart titles from the engine's scenario model
    (`summary.scenarios`): {'bullets': [...], 'bars_title': str,
    'lines_title': str}."""
    c = _ctx(f)
    if c is None:
        return None
    ends, cagrs = sc["ends"], sc["cagrs"]
    for k in ("Optimistic", "Likely", "Conservative"):
        if not (_num(ends.get(k)) and _num(cagrs.get(k))):
            return None
    band = sc.get("band")
    if not _num(band):
        return None
    c.update(
        v_scn_opt=fmt_num(ends["Optimistic"]),
        v_scn_base=fmt_num(ends["Likely"]),
        v_scn_con=fmt_num(ends["Conservative"]),
        v_scn_opt_cagr=fmt_pct(cagrs["Optimistic"]),
        v_scn_base_cagr=fmt_pct(cagrs["Likely"]),
        v_scn_con_cagr=fmt_pct(cagrs["Conservative"]),
        v_scn_flex=f"{int(round(band * 100))}%",
        v_scn_spread=fmt_num(ends["Optimistic"] - ends["Conservative"]))
    out = _render("G09", c)
    if out is None:
        return None
    # the bar chart starts at the last historical year, which the template's
    # {est_year} is not: render that one title with the chart's own base year
    c2 = dict(c, est_year=sc["base_year"])
    bars = _render("G09", c2)
    if bars is None:
        return None
    return {
        "bullets": out["readout"],
        "bars_title": bars["chart1_title"],
        "cards": out["cards"],
        "legend": out["legend"],
        "head": out["readout_head"],
    }
