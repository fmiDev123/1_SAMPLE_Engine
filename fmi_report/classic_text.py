"""Classic (2015-style) write-up, driven by the engine's own numbers.

The wording lives in `fmi_templates/classic.py` (workbook: FMI_Classic_Report_
Templates.xlsx).  This module only builds the template context from `data` /
`metrics` and hands back the rendered sentences, so every figure still comes
from the loaded dataset.

Contract with the call sites: every public function returns a list of strings
(or one string) or **None** when it cannot write the passage - a missing
number, a one-row table, a placeholder with no value.  The caller then keeps
the write-up it had before, so a classic template can never break an export.

Switch back to the old wording for a whole build with FMI_WRITEUP=asis.
Industry (for the driver/restraint phrases the classic text borrows) comes
from FMI_INDUSTRY = food | consumer | healthcare | pharma | industrial; when
it is not set those phrases are left out rather than guessed.
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path

import pandas as pd

_ROOT = str(Path(__file__).resolve().parents[1])
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from fmi_templates.classic import (CLASSIC_INDEX, classic_ctx,        # noqa: E402
                                   item_classic, dynamics_ctx)
from fmi_templates.engine import MissingKeys, render_text              # noqa: E402

from . import config, data, geo, metrics                                # noqa: E402
from .components import fmt_num, fmt_pct                                # noqa: E402
from .config import END_YEAR, EST_YEAR, HIST_END, HIST_START, UNIT_LABELS  # noqa: E402

INDUSTRIES = ("food", "consumer", "healthcare", "pharma", "industrial")


def enabled() -> bool:
    return os.environ.get("FMI_WRITEUP", "classic").strip().lower() != "asis"


def industry() -> str | None:
    v = os.environ.get("FMI_INDUSTRY", "").strip().lower()
    return v if v in INDUSTRIES else None


# ── names ────────────────────────────────────────────────────────────
def _lc(name) -> str:
    """Lower-case a segment name for mid-sentence use; keep acronyms, dotted
    codes, anything with a digit and CamelCase brand names as written."""
    out = []
    for w in str(name).split():
        keep = ((w.isupper() and len(w) > 1) or "." in w
                or any(ch.isdigit() for ch in w)
                or any(ch.isupper() for ch in w[1:]))
        out.append(w if keep else w.lower())
    return " ".join(out)


def _bare(segment) -> str:
    return re.sub(r"^by\s+", "", str(segment).strip(), flags=re.I)


def _cap(s) -> str:
    return s[:1].upper() + s[1:] if s else s


def _num(x) -> bool:
    return x is not None and not pd.isna(x)


def _mult(x) -> str:
    return f"{x:.1f}"


# ── context ──────────────────────────────────────────────────────────
def _val_word() -> str:
    lab = UNIT_LABELS.get("Value", "")
    parts = lab.split(None, 1)
    return parts[1] if len(parts) == 2 else ""


def _vol_single() -> str:
    n = config.volume_noun() or UNIT_LABELS.get("Volume", "unit")
    return n[:-1] if n.lower().endswith("s") and len(n) > 3 else n


def base_ctx(region: str | None = None) -> dict:
    g = region or geo.current()
    mk = data.MARKET
    ind = industry()
    c = dict(
        industry=ind or "consumer", market=mk, market_lc=_lc(mk), mkt_short=mk,
        publisher="Future Market Insights", publisher_short="FMI",
        pub_year=EST_YEAR, geo=g, geo_ref=geo.adj(g),
        geo_ref_cap=geo.adj(g)[:1].upper() + geo.adj(g)[1:], geo_top="Global",
        hist_start=HIST_START, base_year=HIST_END, est_year=EST_YEAR,
        mid_year=(EST_YEAR + END_YEAR) // 2, fc_end=END_YEAR,
        val_unit=UNIT_LABELS["Value"], val_word=_val_word(),
        vol_unit=UNIT_LABELS["Volume"], vol_unit_single=_vol_single(),
        price_unit=config.price_unit(), metric="value", metric_cap="Value",
        m_unit=UNIT_LABELS["Value"], seg_tag="Segment")
    if ind:
        c.update(dynamics_ctx(ind))
    else:                       # no industry given: no borrowed driver phrases
        c.update(drivers=[], restraints=[], trends=[])
    return c


def _kids(g: str):
    """Children of a geography without the 'Rest of ...' buckets."""
    try:
        t = data.children_table(g, "Value")
    except Exception:
        return None
    if t is None or not len(t):
        return None
    keep = [i for i in t.index if not data.is_bucket(i)] or list(t.index)
    return t.loc[keep]


def _level(g: str) -> str:
    return "region" if geo.is_global(g) else "country"


# ── rendering helpers ────────────────────────────────────────────────
def _fix_unit(text: str, unit: str) -> str:
    if unit == "Volume":
        text = (text.replace("value share", "volume share")
                    .replace("higher value", "higher volume")
                    .replace("in terms of value", "in terms of volume"))
    return text


def _strip_segment(text: str, names, geo_items: bool) -> str:
    """'North America segment is ...' -> 'North America is ...' for geographies."""
    if not geo_items:
        return text
    for n in names:
        if n:
            text = (text.replace(f"the {n} segment", str(n))
                        .replace(f"The {n} segment", str(n))
                        .replace(f"{n} segment", str(n)))
    return text


def _one(text, ctx):
    try:
        return render_text(text, ctx, mask=False)
    except MissingKeys:
        return None


def _prep(ctx: dict, items: dict | None = None) -> dict:
    full = classic_ctx(ctx)
    if items:
        full.update(items)
    return full


def _lines(cid: str, ctx: dict, key: str = "outlook", unit: str = "Value",
           geo_items=(), skip=()) -> list:
    """Render the Market Outlook lines of one classic template.  A line whose
    placeholders cannot be filled is dropped (never printed half-empty);
    'Figure NN highlights ...' lines belong to the figure numbering, which
    this text-swap does not use, so they are left out too."""
    t = CLASSIC_INDEX[cid]["elements"]
    tmpl = t.get(key) or t.get(f"{key}__{'volume' if unit == 'Volume' else 'value'}") or []
    out = []
    full = _prep(ctx)
    full.update(item_classic(full, full))
    for line in tmpl:
        if "{fig_" in line or any(s in line for s in skip):
            continue
        r = _one(line, full)
        if r:
            out.append(_strip_segment(_fix_unit(r, unit), geo_items,
                                      bool(geo_items)))
    return out


def _head(cid: str, ctx: dict, unit: str = "Value", geo_items=()):
    full = _prep(ctx)
    full.update(item_classic(full, full))
    r = _one(CLASSIC_INDEX[cid]["headline"], full)
    if not r:
        return None
    return _strip_segment(_fix_unit(r, unit), geo_items, bool(geo_items))


def _stop(text: str) -> str:
    """Bullets end with a full stop (house style); headlines arrive without."""
    text = text.strip()
    return text if not text or text[-1] in ".!?" else text + "."


def _guard(fn):
    """Run a builder; any failure means 'no classic text' (caller keeps old)."""
    def run(*a, **kw):
        if not enabled():
            return None
        try:
            res = fn(*a, **kw)
        except Exception as exc:            # noqa: BLE001 - never break an export
            if os.environ.get("FMI_WRITEUP_DEBUG"):
                raise
            print(f"[classic write-up] {fn.__name__}: {exc!r} - keeping as-is text",
                  file=sys.stderr)
            return None
        if isinstance(res, str):
            return _stop(res) or None
        return [_stop(x) for x in res] or None if res else None
    run.__name__ = fn.__name__
    run.__doc__ = fn.__doc__
    return run


# ── geography overview (G03 / G04 -> R14, R26, R14V) ─────────────────
def _geo_value_ctx(g: str) -> dict | None:
    s = data.total_series(g, "Value")
    base, est, end = s[HIST_END], s[EST_YEAR], s[END_YEAR]
    if not all(_num(v) and v > 0 for v in (base, est, end)):
        return None
    c = base_ctx(g)
    f = metrics.fore_cagr(s)
    c.update(r_mult=end / base, v_mult=_mult(end / base),
             v_val_base=fmt_num(base), v_val_est=fmt_num(est),
             v_val_end=fmt_num(end), v_yoy_est=fmt_pct(est / base - 1),
             v_abs_opp=fmt_num(end - base), v_abs_est=fmt_num(est - base),
             v_f_cagr=fmt_pct(f), r_f_cagr=f if _num(f) else 0.0,
             r_h_cagr=metrics.hist_cagr(s) if _num(metrics.hist_cagr(s)) else 0.0)
    k = _kids(g)
    if k is not None and len(k) >= 1:
        end_by = k[END_YEAR].astype(float).sort_values(ascending=False)
        c["geo_level"] = _level(g)
        c["cty_lead"] = str(end_by.index[0])
        if len(end_by) > 1:
            c["cty_second"] = str(end_by.index[1])
    return c


def _geo_volume_ctx(g: str, c: dict) -> dict | None:
    v = data.total_series(g, "Volume")
    base, end = v[HIST_END], v[END_YEAR]
    if not all(_num(x) and x > 0 for x in (base, end)):
        return None
    vc = metrics.fore_cagr(v)
    out = dict(c)
    out.update(metric="volume", metric_cap="Volume", m_unit=UNIT_LABELS["Volume"],
               v_vol_mult=_mult(end / base), v_vol_base=fmt_num(base),
               v_vol_end=fmt_num(end), v_vol_f_cagr=fmt_pct(vc),
               r_vol_f_cagr=vc if _num(vc) else 0.0)
    return out


@_guard
def geo_value_bullets(g: str) -> list | None:
    """Value read-out for a geography: R14 (global) or R26 (region/country)."""
    c = _geo_value_ctx(g)
    if c is None:
        return None
    if geo.is_global(g):
        c.update(cp_geo_country_line="")
        out = [_head("R14", c)] + _lines("R14", c)
    else:
        out = [_head("R26", c)] + _lines("R26", c)
    return [x for x in out if x]


@_guard
def geo_volume_bullets(g: str) -> list | None:
    c = _geo_value_ctx(g)
    if c is None:
        return None
    c = _geo_volume_ctx(g, c)
    if c is None:
        return None
    out = [_head("R14V", c)] + _lines("R14V", c, unit="Volume")
    if not geo.is_global(g):
        out = [x.replace("Global ", f"{g} ", 1) if x.startswith("Global ") else x
               for x in out if x]
    return [x for x in out if x]


@_guard
def vol_val_bullets(g: str) -> list | None:
    """Volume & Value Analysis page: value read-out then volume read-out."""
    val = geo_value_bullets(g) or []
    vol = geo_volume_bullets(g) or []
    # the volume headline repeats the value one's shape; keep the figures lines
    return (val + vol) or None


@_guard
def summary_page_bullets(g: str, unit: str = "Value") -> list | None:
    """Market Summary 'Key Insights' panel (G01 / G03): the outlook lines."""
    c = _geo_value_ctx(g)
    if c is None:
        return None
    if unit == "Volume":
        c = _geo_volume_ctx(g, c)
        if c is None:
            return None
        out = [_head("R14V", c, "Volume")] + _lines("R14V", c, unit="Volume")
        return [x.replace("Global ", f"{g} ", 1) if (x.startswith("Global ")
                and not geo.is_global(g)) else x for x in out] or None
    if geo.is_global(g):
        c.update(cp_geo_country_line="")
    cid = "R14" if geo.is_global(g) else "R26"
    return [_head(cid, c)] + _lines(cid, c) or None


# ── a split of the market into parts (segment / region / country) ────
def _summary_of(tbl: pd.DataFrame) -> pd.DataFrame:
    tot = tbl.sum(axis=0)
    out = pd.DataFrame(index=tbl.index)
    out["est"] = tbl[EST_YEAR].astype(float)
    out["end"] = tbl[END_YEAR].astype(float)
    out["share_est"] = out["est"] / tot[EST_YEAR]
    out["share_end"] = out["end"] / tot[END_YEAR]
    out["bps"] = (out["share_end"] - out["share_est"]) * 10_000
    out["fcagr"] = [metrics.fore_cagr(tbl.loc[i]) for i in tbl.index]
    out["abs_opp"] = out["end"] - out["est"]
    return out


def _split_ctx(g: str, tbl: pd.DataFrame, label: str, unit: str,
               geo_items: bool) -> dict | None:
    """Context for the 'parts of the market' templates (R21 / R22 / R24)."""
    tbl = tbl[tbl[END_YEAR].astype(float) > 0] if len(tbl) else tbl
    if len(tbl) < 2:
        return None
    df = _summary_of(tbl)
    if df[["share_end", "bps"]].isna().all().all():
        return None
    lead = metrics.imax(df["end"])
    fast = metrics.imax(df["fcagr"])
    lose = metrics.imin(df["bps"])
    c = base_ctx(g)
    c.update(segment=label, segment_lc=_lc(label),
             metric="volume" if unit == "Volume" else "value",
             metric_cap=unit, m_unit=UNIT_LABELS[unit],
             sub_lead=str(lead), sub_fast=str(fast),
             r_sub_lead_bps=float(df.loc[lead, "bps"]),
             v_sub_lead_bps=f"{abs(df.loc[lead, 'bps']):,.0f}",
             v_sub_lead_share_end=fmt_pct(df.loc[lead, "share_end"]),
             v_sub_fast_cagr=fmt_pct(df.loc[fast, "fcagr"]))
    if geo_items:                         # proper nouns keep their capitals
        c.update(sub_lead_lc=str(lead), sub_fast_lc=str(fast))
    if df.loc[lose, "bps"] <= -10 and lose != lead:
        c.update(sub_lose=str(lose),
                 v_sub_lose_bps=f"{abs(df.loc[lose, 'bps']):,.0f}",
                 v_sub_lose_share_end=fmt_pct(df.loc[lose, "share_end"]))
        if geo_items:
            c["sub_lose_lc"] = str(lose)
    order = df["end"].sort_values(ascending=False).index
    c.update(sub_second=str(order[1]),
             v_lead_vs_second=_mult(df.loc[order[0], "end"] / df.loc[order[1], "end"])
             if df.loc[order[1], "end"] else "1.0")
    if geo_items:
        c["sub_second_lc"] = str(order[1])
    c["_df"] = df
    c["_names"] = [str(lead), str(fast), str(order[1])] + (
        [str(lose)] if "sub_lose" in c else [])
    return c


def _public(c: dict) -> dict:
    return {k: v for k, v in c.items() if not k.startswith("_")}


def _split_bullets(c: dict, unit: str, geo_items: bool, with_table=False):
    pub = _public(c)
    names = c["_names"]
    out = [_head("R21", pub, unit, names if geo_items else ())]
    out += _lines("R21", pub, unit=unit, geo_items=names if geo_items else ())
    if with_table:
        out.insert(0, _head("R24", pub, unit, names if geo_items else ()))
    return [x for x in out if x]


@_guard
def children_bullets(g: str, unit: str) -> list | None:
    """Share read-out by child geography (B01-B03: R21 + R24 + R28 headline)."""
    kids = data.children_table(g, unit)
    kids = kids.loc[[i for i in kids.index if not data.is_bucket(i)] or list(kids.index)]
    label = "Region" if geo.is_global(g) else "Country"
    c = _split_ctx(g, kids, label, unit, geo_items=True)
    if c is None:
        return None
    out = _split_bullets(c, unit, True)
    p = dict(_public(c))
    top = _guard_head("R28", p, unit)
    return ([top] if top else []) + out


def _guard_head(cid, ctx, unit="Value"):
    try:
        full = dict(ctx)
        full["cty_lead"] = full.get("sub_lead")
        if full.get("sub_second"):
            full["cty_second"] = full["sub_second"]
            full["v_cty_ratio"] = full.get("v_lead_vs_second", "1.0")
        h = _head(cid, full, unit)
        return h
    except Exception:
        return None


@_guard
def segment_size_bullets(g: str, unit: str, segment: str, tbl) -> list | None:
    """Size-table read-out for one segment (D04 / D05: R24 headline + R21)."""
    if tbl is None or not len(tbl):
        return None
    c = _split_ctx(g, tbl, _bare(segment), unit, geo_items=False)
    if c is None:
        return None
    return _split_bullets(c, unit, False, with_table=True)


@_guard
def segment_share_bullets(g: str, segment: str, tbl) -> list | None:
    """Value-composition + BPS page (D06: R21)."""
    if tbl is None or not len(tbl):
        return None
    c = _split_ctx(g, tbl, _bare(segment), "Value", geo_items=False)
    if c is None:
        return None
    return _split_bullets(c, "Value", False)


@_guard
def attractiveness_bullets(g: str, segment: str, unit: str) -> list | None:
    """Market structure & attractiveness page (D02: R21 + R22).

    The chart ranks sub-segments on a CAGR index and a market-share index
    (100 = the segment average on each).  'Most attractive' is read the same
    way the chart is: highest average of the two indices.  No attractiveness
    score is quoted - the template's 'index of 3.7' has no counterpart here."""
    tbl = data.segment_table(g, unit, segment)
    s = metrics.segment_summary(g, unit, segment)
    if len(s) < 2:
        return None
    c = _split_ctx(g, tbl, _bare(segment), unit, geo_items=False)
    if c is None:
        return None
    score = (s["cagr_index"] + s["ms_index"]) / 2.0
    order = list(score.sort_values(ascending=False).index)
    c["sub_attr1"] = str(order[0])
    c["sub_attr2"] = str(order[1])
    if len(order) > 2 and score[order[-1]] < 100:
        c["sub_attr_low"] = str(order[-1])
    pub = _public(c)
    out = [_head("R22", pub, unit)]
    for line in _lines("R22", pub, unit=unit, skip=("attractiveness index of",)):
        out.append(line)
    # the template's 'attractiveness index of N' line, restated with the two
    # indices the chart actually plots (100 = the segment average)
    a, b = order[0], order[1]
    if c["sub_attr2"] != c["sub_attr1"]:
        out.insert(1, f"{_cap(_lc(a))} and {_lc(b)} exhibit a CAGR index of "
                      f"{s.loc[a, 'cagr_index']:.1f} and {s.loc[b, 'cagr_index']:.1f} "
                      f"and a market share index of {s.loc[a, 'ms_index']:.1f} and "
                      f"{s.loc[b, 'ms_index']:.1f} respectively, against the segment "
                      f"average of 100.")
    return [x for x in out if x]


# ── one sub-segment (D01 / D03: R23) ─────────────────────────────────
def _item_ctx(g: str, tbl: pd.DataFrame, name: str, label: str, lead: bool,
              unit: str, base_year: int = HIST_END) -> dict | None:
    row = tbl.loc[name].astype(float)
    base, est, end = row[base_year], row[EST_YEAR], row[END_YEAR]
    if not all(_num(v) and v > 0 for v in (base, end)):
        return None
    tot_base = float(tbl[base_year].astype(float).sum())
    f = metrics.fore_cagr(row)
    c = base_ctx(g)
    c["base_year"] = base_year
    c.update(segment=label, segment_lc=_lc(label), item=str(name),
             is_lead=lead, r_item_mult=end / base, v_item_mult=_mult(end / base),
             v_item_base=fmt_num(base), v_item_est=fmt_num(est),
             v_item_end=fmt_num(end), v_item_cagr=fmt_pct(f),
             v_item_opp=fmt_num(end - base),
             v_item_abs_est=fmt_num(est - base),
             v_item_yoy_est=fmt_pct(est / base - 1),
             v_item_share_base=fmt_pct(base / tot_base) if tot_base else "0%",
             metric="volume" if unit == "Volume" else "value",
             metric_cap=unit, m_unit=UNIT_LABELS[unit])
    return c


@_guard
def segment_opportunity_bullets(g: str, segment: str, tbl, top: int = 3) -> list | None:
    """Absolute $ opportunity summary (D03): the R23 sentences for the
    sub-segments that add the most."""
    if tbl is None or not len(tbl):
        return None
    label = _bare(segment)
    # measured from the first forecast year - the same span the waterfalls on
    # the page draw - so the sentences and the charts quote one number
    opp = (tbl[END_YEAR].astype(float) - tbl[EST_YEAR].astype(float)).sort_values(ascending=False)
    lead = metrics.imax(tbl[EST_YEAR].astype(float))
    out = []
    for name in list(opp.index[:top]):
        c = _item_ctx(g, tbl, name, label, str(name) == str(lead), "Value", EST_YEAR)
        if c is None:
            continue
        lines = [x for x in _lines("R23", c, key="outlook", unit="Value")
                 if "incremental opportunity" in x]
        out += [x.replace("This segment", f"The {_lc(name)} segment", 1) for x in lines[:1]]
    head = None
    if len(opp):
        c = _item_ctx(g, tbl, opp.index[0], label, str(opp.index[0]) == str(lead), "Value", EST_YEAR)
        head = _head("R23", c) if c else None
    return ([head] if head else []) + out


@_guard
def segment_highlights(g: str, segment: str, unit: str, tbl) -> list | None:
    """Segment outlook opener (D01): the classic read-out for the segment."""
    if tbl is None or len(tbl) < 2:
        return None
    c = _split_ctx(g, tbl, _bare(segment), unit, geo_items=False)
    if c is None:
        return None
    pub = _public(c)
    out = [_head("R21", pub, unit)] + _lines("R21", pub, unit=unit)[:1]
    out.append(_head("R24", pub, unit))          # three lines fit the panel
    return [x for x in out if x]


# ── country page (C01: R26) ──────────────────────────────────────────
@_guard
def country_momentum(country: str) -> str | None:
    c = _geo_value_ctx(country)
    if c is None:
        return None
    c["geo_level"] = "country"
    c.pop("cty_lead", None), c.pop("cty_second", None)
    lines = _lines("R26", c)
    keep = [x for x in lines if "incremental opportunity" in x or "at a CAGR of" in x]
    return " ".join(keep) or None


# ── G05 country indices (R28 headline) ───────────────────────────────
@_guard
def country_headline(unit: str = "Value") -> str | None:
    df = metrics.country_indices(unit)
    if df is None or len(df) < 2:
        return None
    df = df.sort_values("end", ascending=False)
    a, b = df.iloc[0], df.iloc[1]
    if not (_num(a["end"]) and _num(b["end"]) and b["end"] > 0):
        return None
    c = base_ctx()
    c.update(cty_lead=str(a["Country"]), cty_second=str(b["Country"]),
             v_cty_ratio=_mult(a["end"] / b["end"]))
    return _head("R28", c)


# ── executive overview (G06: R07 overview + R06 foreword) ────────────
@_guard
def overview_bullets(g: str) -> list | None:
    c = _geo_value_ctx(g)
    if c is None:
        return None
    k = _kids(g)
    if k is not None and len(k) >= 2:
        cg = {i: metrics.fore_cagr(k.loc[i]) for i in k.index}
        fast = max(cg, key=lambda i: cg[i] if _num(cg[i]) else -9)
        lead = metrics.imax(k[END_YEAR].astype(float))
        c.update(child_lead=str(lead), child_fast=str(fast),
                 v_child_fast_cagr=fmt_pct(cg[fast]))
    c.update(v_val_base=c["v_val_base"], v_val_end=c["v_val_end"])
    segs = [_bare(s) for s in data.SEGMENTS if not data.is_cross(s)][:4]
    if segs:
        c["segments"] = segs
    out = [_head("R07", {**c, "market_lc_cap": None} if False else c)]
    out += _lines("R07", c, key="overview")
    if "child_lead" in c:
        out += _lines("R06", c, key="foreword", skip=("Global demand for",
                                                      "Globally,"))
    return [x for x in out if x]


# ── average price (G02 / P05: R30) ───────────────────────────────────
def _price_fmt(v: float) -> str:
    return f"{v:,.2f}" if abs(v) < 100 else f"{v:,.0f}"


@_guard
def price_bullets(g: str, segment: str | None = None) -> list | None:
    """Average-price read-out: headline, market price base -> end, and the
    dearest / cheapest child geography.  Price = Value / Volume, as on the
    pricing slides themselves."""
    val, vol = data.total_series(g, "Value"), data.total_series(g, "Volume")
    p_base = metrics.avg_price(val, vol, [HIST_END])
    p_end = metrics.avg_price(val, vol, [END_YEAR])
    if not all(_num(x) and x > 0 for x in (p_base, p_end)):
        return None
    c = base_ctx(g)
    cg = (p_end / p_base) ** (1.0 / (END_YEAR - HIST_END)) - 1.0
    c.update(r_price_avg_cagr=cg, v_price_avg_base=_price_fmt(p_base),
             v_price_avg_end=_price_fmt(p_end),
             seg1=_bare(segment) if segment else "segment",
             seg1_lc=_lc(_bare(segment)) if segment else "segment")
    if segment:
        sv = data.segment_table(g, "Value", segment).astype(float)
        sw = data.segment_table(g, "Volume", segment).astype(float)
        ps = config.price_scale()
        both = [i for i in sv.index if i in sw.index and sw.loc[i, END_YEAR] > 0]
        if both:
            pe = pd.Series({i: sv.loc[i, END_YEAR] * ps / sw.loc[i, END_YEAR] for i in both})
            top = metrics.imax(pe)
            c.update(price_seg_top=str(top), v_price_seg_top_end=_price_fmt(pe[top]))
    kv, kw = _kids(g), data.children_table(g, "Volume")
    if kv is not None and len(kv) >= 2:
        ps = config.price_scale()
        cp = pd.Series({str(i): kv.loc[i, EST_YEAR] * ps / kw.loc[i, EST_YEAR]
                        for i in kv.index if i in kw.index and kw.loc[i, EST_YEAR] > 0})
        if len(cp) >= 2:
            c.update(child_price_high=str(metrics.imax(cp)),
                     child_price_low=str(metrics.imin(cp)))
    out = []
    if "price_seg_top" in c:
        out.append(_head("R30", c))
    out += _lines("R30", c, skip=("derived as market value",))
    if not geo.is_global(g):
        out = [x.replace("-priced region", "-priced country") for x in out]
    return out


# ── cross-section (X01: R31 headline) ────────────────────────────────
@_guard
def cross_headline(region: str, a_name: str, b_name: str, m1) -> str | None:
    """'X segment is projected to account for the largest share ... led by Y
    at US$ N in <end year>' - read off the end-year matrix on the page."""
    if m1 is None or m1.empty:
        return None
    a_tot = m1.sum(axis=1)
    top_a = metrics.imax(a_tot)
    top_b = metrics.imax(m1.loc[top_a])
    c = base_ctx(region)
    c.update(xa_lead=str(top_a), xb_lead_in_a=str(top_b),
             v_xb_lead_in_a_end=fmt_num(m1.loc[top_a, top_b]),
             seg_a=a_name, seg_b=b_name)
    return _head("R31", c)


# ── key players (L02: R32) ───────────────────────────────────────────
@_guard
def company_share_bullets(region: str, shares) -> list | None:
    """Market-share read-out: leader, concentration band (top-three share)
    and the runner-up, from the same [(company, share%)] the donut draws."""
    named = sorted([(n, s) for n, s in shares if n.lower() != "others"],
                   key=lambda kv: kv[1], reverse=True)
    if len(named) < 3:
        return None
    (c1, s1), (c2, s2) = named[0], named[1]
    cr3 = sum(s for _, s in named[:3]) / 100.0
    c = base_ctx(region)
    c.update(base_year=EST_YEAR,           # the donut is drawn at the first forecast year
             comp1=c1, comp2=c2, v_comp1_share=f"{s1:.1f}%",
             v_comp2_share=f"{s2:.1f}%", r_cr3=cr3, v_cr3=f"{cr3 * 100:.0f}%")
    out = [_head("R32", c)] + _lines("R32", c)
    word = geo.adj(region)
    return [x.replace("the global ", f"the {word} ").replace("Global ", f"{word} ")
            if word != "global" else x for x in out]


# ═════════════════════════════════════════════════════════════════════
# Paragraphs and AI-written reasons
# ═════════════════════════════════════════════════════════════════════
import atexit
import json

_COLLECT = os.environ.get("FMI_REASONS_COLLECT", "").strip()
_REASON_FILE = os.environ.get("FMI_REASONS", "").strip()
_FACTS: dict = {}
_REASONS: dict | None = None


def _reasons() -> dict:
    global _REASONS
    if _REASONS is None:
        _REASONS = {}
        if _REASON_FILE and Path(_REASON_FILE).exists():
            _REASONS = json.loads(Path(_REASON_FILE).read_text(encoding="utf-8"))
    return _REASONS


def _dump_facts() -> None:
    if _COLLECT and _FACTS:
        Path(_COLLECT).write_text(json.dumps(_FACTS, indent=1, ensure_ascii=False),
                                  encoding="utf-8")


atexit.register(_dump_facts)


def _clean(v):
    if isinstance(v, float):
        return round(v, 4)
    if isinstance(v, dict):
        return {str(k): _clean(x) for k, x in v.items()}
    if isinstance(v, (list, tuple)):
        return [_clean(x) for x in v]
    return v


def reason(kind: str, g: str, subject: str, unit: str, facts: dict) -> str | None:
    """The AI-written 'why' for one slide.  In collect mode (FMI_REASONS_COLLECT)
    the slide's facts are recorded under its key; in use mode (FMI_REASONS) the
    written reason is returned, or None when there is none for that key."""
    if not enabled() or not facts:
        return None
    key = f"{kind}|{g}|{subject}|{unit}"
    if _COLLECT:
        _FACTS[key] = _clean(dict(facts, kind=kind, geography=g, subject=subject,
                                  unit=unit, market=data.MARKET))
    return _reasons().get(key) or None


def paras(sentences, why: str | None = None) -> list:
    """Sentences -> two or three paragraphs ('¶' marks a paragraph, so the
    panel draws it without a bullet).  With a reason it is the last paragraph
    and the facts above it take two."""
    s = [x.strip() for x in sentences if x and x.strip()]
    if not s:
        return [("¶ " + why)] if why else []
    n = 2 if (why or len(s) <= 4) else 3
    n = min(n, len(s))
    k, r = divmod(len(s), n)
    out, i = [], 0
    for j in range(n):
        step = k + (1 if j < r else 0)
        out.append("¶ " + " ".join(_stop(x) for x in s[i:i + step]))
        i += step
    if why:
        out.append("¶ " + why.strip())
    return out


def split_facts(g: str, tbl: pd.DataFrame, label: str, unit: str) -> dict | None:
    """Facts behind a 'parts of the market' slide, for the reason writer."""
    if tbl is None or len(tbl) < 2:
        return None
    t = tbl[tbl[END_YEAR].astype(float) > 0]
    if len(t) < 2:
        return None
    df = _summary_of(t)
    lead, fast = metrics.imax(df["end"]), metrics.imax(df["fcagr"])
    gain, lose = metrics.imax(df["bps"]), metrics.imin(df["bps"])
    tot = metrics.fore_cagr(t.sum(axis=0))
    rows = {str(i): dict(share_2026=round(df.loc[i, "share_est"] * 100, 1),
                         share_2036=round(df.loc[i, "share_end"] * 100, 1),
                         bps_shift=round(df.loc[i, "bps"]),
                         cagr_pct=round(df.loc[i, "fcagr"] * 100, 1))
            for i in df.index[:12]}
    return dict(split_by=label, total_cagr_pct=round(tot * 100, 1),
                largest=str(lead), fastest_growing=str(fast),
                biggest_share_gain=str(gain), biggest_share_loss=str(lose),
                parts=rows)


def geo_facts(g: str) -> dict | None:
    v, w = data.total_series(g, "Value"), data.total_series(g, "Volume")
    if not (_num(v[EST_YEAR]) and _num(w[EST_YEAR])):
        return None
    if not float(w[EST_YEAR]) > 0:          # a value-only market has no volume
        return None
    p0 = metrics.avg_price(v, w, [EST_YEAR])
    p1 = metrics.avg_price(v, w, [END_YEAR])
    k = _kids(g)
    kids = {}
    if k is not None:
        for i in k.index[:12]:
            kids[str(i)] = round(metrics.fore_cagr(k.loc[i]) * 100, 1)
    return dict(value_cagr_2026_36_pct=round(metrics.fore_cagr(v) * 100, 1),
                value_cagr_2021_25_pct=round(metrics.hist_cagr(v) * 100, 1),
                volume_cagr_2026_36_pct=round(metrics.fore_cagr(w) * 100, 1),
                volume_cagr_2021_25_pct=round(metrics.hist_cagr(w) * 100, 1),
                avg_price_2026=round(p0), avg_price_2036=round(p1),
                price_unit=config.price_unit(), value_unit=UNIT_LABELS["Value"],
                volume_unit=UNIT_LABELS["Volume"],
                child_cagr_pct=kids)
