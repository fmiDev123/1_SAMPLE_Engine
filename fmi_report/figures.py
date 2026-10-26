"""Plotly figure builders themed to the PBIX report (CY23SU08 theme).

All figures autosize to fill their visual container, exactly like Power
BI visuals resize inside their frames.
"""
from __future__ import annotations

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from . import config
from .config import (ACCENT, BAD, DATA_COLORS, DECLINE_COLOR, DIV_NEG, DIV_POS,
                     DN_FILL, DN_LINE, F3_BODY, F3_CONE, F3_LBL, F3_RIM,
                     F3_TOP, FAINT, FONT_STACK, FONT_UI, GEO_ISO3, GOOD, GRID,
                     HEADER_BG, HEAT_STOPS, MUTED, NEST_INNER, NEST_OUTER,
                     NEUTRAL_LINE, ON_FILL_TEXT, ORANGE, OUTPACE_COLOR,
                     PANEL_BG, RADAR_FILL, RADAR_LINE, REF_LINE, SEQ_HIGH,
                     SEQ_LOW, SEQ_STOPS, TA_ACTUAL, TA_TARGET, TEXT,
                     TREND_DOWN, TREND_UP, UP_FILL, UP_LINE)

FONT = dict(family=FONT_STACK, size=11, color=TEXT)

# Standard bar geometry: the gap between bars is 50% of the bar width.
# bar : gap = 2 : 1  ->  bar = 2/3 of a slot, gap = 1/3 of a slot.
BAR_GAP = 1.0 / 3.0        # Plotly layout.bargap (single-series / stacked)
BAR_WIDTH = 2.0 / 3.0      # Plotly go.Bar width (when set explicitly per trace)

# Plotly's built-in template supplies its OWN palette whenever a trace is not
# given an explicit colour - and its 3rd default colour is a green (#00CC96),
# its 8th a lime (#B6E880), and its default diverging scale is PiYG (more
# greens).  Overriding both here means no chart can ever fall back to a
# non-brand colour: every builder goes through _layout(), so this is the one
# place that has to be right.
_COLORWAY = list(DATA_COLORS)
_SEQ_SCALE = [[0.0, SEQ_LOW], [1.0, SEQ_HIGH]]
_DIV_SCALE = [[0.0, BAD], [0.5, "#F4F7FA"], [1.0, GOOD]]   # decline → growth


# A bar is never wider than this fraction of the plot area.  With one category
# Plotly hands the bar BAR_WIDTH of the whole axis (two thirds), which reads as
# a slab; this keeps a 1-2 category chart looking like a bar chart.
MAX_BAR_FRAC = 0.18

# Every trace type drawn as vertical/horizontal bars.  'waterfall' matters as
# much as 'bar' here: the Growth Contribution and Absolute $ Opportunity charts
# are waterfalls, and with a single-sub-segment segment they are the widest
# slabs in the report.
_BAR_LIKE = ("bar", "waterfall", "funnel", "histogram")


def bar_slot_width(n_cats: int) -> float:
    """The bar width (in category units) to use for `n_cats` categories."""
    if n_cats <= 0:
        return BAR_WIDTH
    return min(BAR_WIDTH, MAX_BAR_FRAC * n_cats)


def cap_bar_widths(figure):
    """Narrow the bars of a sparse bar chart, in place.

    Called from layout.fig(), so it applies to every figure the flow engine
    places - no builder has to remember it.  Traces that already carry an
    explicit width are SCALED by the same factor, which preserves the
    inner/outer ratio of the nested 'bar inside a bar' charts.
    """
    if getattr(figure, "_bars_capped", False):
        return figure                      # idempotent: called from two places
    try:
        bars = [t for t in figure.data
                if getattr(t, "type", "") in _BAR_LIKE]
    except Exception:
        return figure
    try:
        figure._bars_capped = True
    except Exception:
        pass
    if not bars:
        return figure
    n = 0
    for t in bars:
        cats = t.y if getattr(t, "orientation", None) == "h" else t.x
        if cats is not None and len(cats):
            n = max(n, len({str(c) for c in cats}))
    cap = bar_slot_width(n)
    if n == 0 or cap >= BAR_WIDTH:            # nothing to narrow
        return figure
    k = cap / BAR_WIDTH
    for t in bars:
        if getattr(t, "width", None) is None:
            t.width = cap
        else:
            try:
                t.width = float(t.width) * k
            except (TypeError, ValueError):
                t.width = cap
    return figure


def _layout(title=""):
    return dict(
        autosize=True,
        plot_bgcolor=PANEL_BG, paper_bgcolor=PANEL_BG, font=FONT,
        colorway=_COLORWAY,
        colorscale=dict(sequential=_SEQ_SCALE, sequentialminus=_SEQ_SCALE,
                        diverging=_DIV_SCALE),
        margin=dict(l=8, r=8, t=30 if title else 10, b=8),
        title=dict(text=title, font=dict(size=11.5, color=TEXT),
                   x=0.01) if title else None,
        hoverlabel=dict(font=dict(family=FONT_UI, size=11)),
    )

def _pal(n):
    return [DATA_COLORS[i % len(DATA_COLORS)] for i in range(n)]

def _abbrev_floor() -> float:
    """Magnitude below which numbers are printed in the workbook's own unit -
    see `config.abbrev_floor`.  Read per call so it follows the loaded units."""
    try:
        return float(config.abbrev_floor())
    except Exception:
        return 1e3

def _abbr_lbl(v):
    """Short numeric label (211K / 1.2M) for on-chart data labels, abbreviated
    only as far as the workbook's unit labels allow: figures captioned
    'US$ Mn' keep their own unit (50,531) instead of restating a magnitude."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ""
    if v != v:                                  # NaN
        return ""
    a = abs(v)
    floor = _abbrev_floor()
    if a >= 1e9 and floor <= 1e9:
        return f"{v / 1e9:.1f}B"
    if a >= 1e6 and floor <= 1e6:
        return f"{v / 1e6:.1f}M"
    if a >= 1e4 and floor <= 1e3:
        return f"{v / 1e3:.0f}K"
    if a >= 100:
        return f"{v:,.0f}"
    return f"{v:.1f}"

# ── 'Bar inside a bar' (nested overlapping bars) ─────────────────────
def nested_bar(x, inner, inner_name, outer, outer_name, dual=False):
    """A wide OUTER bar (secondary metric) with a narrower INNER bar centred
    inside it.  ``dual=True`` gives the inner its own LEFT axis and the outer a
    RIGHT axis (e.g. Volume left / Value right); ``dual=False`` keeps both on
    one axis (e.g. current vs previous-year value)."""
    xs = [str(v) for v in x]
    inner, outer = list(inner), list(outer)

    def _mx(v):
        m = [abs(float(t)) for t in v if t == t and t is not None]
        return max(m) if m else 1.0

    fig = make_subplots(specs=[[{"secondary_y": dual}]])
    # outer (wide, light) on the PRIMARY axis, drawn first (behind).  width=0.5
    # leaves a 50% gap between adjacent bars (the inner is scaled to keep the
    # 'bar inside a bar' nesting proportion)
    fig.add_trace(go.Bar(
        x=xs, y=outer, name=outer_name, width=0.5,
        marker=dict(color=NEST_OUTER, line=dict(color="#7FBEDD", width=1),
                    cornerradius=2),
        text=[_abbr_lbl(v) for v in outer], textposition="outside",
        textfont=dict(size=8, color=MUTED), cliponaxis=False,
    ), secondary_y=False)
    # inner (narrow) on the SECONDARY axis when dual, so it draws ON TOP of the
    # outer bar; drawn second either way
    fig.add_trace(go.Bar(
        x=xs, y=inner, name=inner_name, width=0.235,
        marker=dict(color=NEST_INNER, cornerradius=2),
        text=[_abbr_lbl(v) for v in inner], textposition="inside",
        insidetextanchor="middle", textfont=dict(size=7.5, color="#FFFFFF"),
        cliponaxis=False,
    ), secondary_y=dual)
    fig.update_layout(
        **_layout(), barmode="overlay", bargap=0.28,
        legend=dict(orientation="h", y=1.16, x=0, font=dict(size=9.5)),
        xaxis=dict(tickfont=dict(size=9), type="category"),
    )
    if dual:
        # outer (e.g. Value) reads on the RIGHT, inner (e.g. Volume) on the
        # LEFT; each axis scaled to its own metric so both bars nest cleanly
        fig.update_yaxes(secondary_y=False, side="right", showgrid=True,
                         gridcolor=GRID, tickfont=dict(size=9), tickformat="~s",
                         range=[0, _mx(outer) * 1.18])
        fig.update_yaxes(secondary_y=True, side="left", showgrid=False,
                         tickfont=dict(size=9), tickformat="~s",
                         range=[0, _mx(inner) * 1.18])
    else:
        fig.update_yaxes(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                         rangemode="tozero", tickformat="~s")
    return fig

# ── Combo: bars + line (PBIX lineStackedColumnComboChart) ────────────
def combo_chart(x, bars, bar_name, line, line_name,
                bar_fmt="{:,.0f}", line_is_pct=True, show_bar_labels=True):
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=list(x), y=list(bars), name=bar_name,
        marker=dict(color=ACCENT, cornerradius=3),
        offsetgroup="a", alignmentgroup="g",
        text=[bar_fmt.format(v) for v in bars] if show_bar_labels else None,
        textposition="outside", textfont=dict(size=8.5, color=MUTED),
    ), secondary_y=False)
    if line_is_pct:
        # YoY / percentage metric -> line on the secondary axis
        fig.add_trace(go.Scatter(
            x=list(x), y=list(line), name=line_name, mode="lines+markers",
            line=dict(color=ORANGE, width=2.5), marker=dict(size=5),
        ), secondary_y=True)
    else:
        # a second absolute metric (e.g. Value vs Volume) -> grouped double
        # bar on its own axis.  Line charts are reserved for YoY only.
        fig.add_trace(go.Bar(
            x=list(x), y=list(line), name=line_name,
            marker=dict(color=ORANGE, cornerradius=3),
            offsetgroup="b", alignmentgroup="g",
            text=[bar_fmt.format(v) for v in line] if show_bar_labels else None,
            textposition="outside", textfont=dict(size=8.5, color=MUTED),
        ), secondary_y=True)
    fig.update_layout(
        **_layout(),
        barmode="group",
        legend=dict(orientation="h", y=1.14, x=0, font=dict(size=9.5)),
        xaxis=dict(tickfont=dict(size=9), type="category"),
    )
    fig.update_yaxes(showgrid=True, gridcolor=GRID, secondary_y=False,
                     tickfont=dict(size=9), rangemode="tozero",
                     tickformat="~s")            # 30000 -> "30k", 2e6 -> "2M"
    fig.update_yaxes(showgrid=False, secondary_y=True,
                     tickformat=".0%" if line_is_pct else "~s",
                     tickfont=dict(size=9),
                     rangemode="tozero" if not line_is_pct else "normal")
    return fig

# ── Donut (PBIX donutChart) ──────────────────────────────────────────
def donut(labels, values, title="", center_text=None, legend=True,
          colors=None, legend_below=False):
    fig = go.Figure(go.Pie(
        labels=list(labels), values=list(values), hole=0.55,
        marker=dict(colors=colors or _pal(len(labels)),
                    line=dict(color=PANEL_BG, width=2)),
        textinfo="percent", textfont=dict(size=10),
        sort=False,
    ))
    leg = (dict(orientation="h", x=0.5, xanchor="center", y=-0.06,
                yanchor="top", font=dict(size=9)) if legend_below
           else dict(orientation="v", x=1.0, y=0.5, font=dict(size=9)))
    fig.update_layout(**_layout(title), showlegend=legend, legend=leg)
    if center_text:
        fig.add_annotation(text=str(center_text), showarrow=False,
                           font=dict(size=13, color=TEXT))
    fig._chart_meta = {"chart": "donut", "center": str(center_text or "")}
    return fig

def region_stack(kids, years, unit_label=""):
    """100%-stacked columns - one full-height bar per year in ``years``, split by
    child (region/country) as a share of that year's total, largest at the
    bottom.  Each segment shows its share % and carries the child name as
    ``customdata`` so a click still drives the region drill-down / drill-through."""
    xs = [str(y) for y in years]
    order = list(kids[years[-1]].sort_values(ascending=False).index)
    colors = _pal(len(order))
    totals = {y: (sum(float(kids.loc[nm, y]) for nm in order) or 1.0)
              for y in years}
    fig = go.Figure()
    meta_series = []
    for i, name in enumerate(order):
        vals = [float(kids.loc[name, y]) for y in years]
        shares = [v / totals[y] * 100.0 for v, y in zip(vals, years)]
        fig.add_trace(go.Bar(
            x=xs, y=vals, name=str(name),
            marker=dict(color=colors[i], line=dict(color=PANEL_BG, width=1)),
            customdata=[str(name)] * len(xs),
            text=[f"{s:.0f}%" if s >= 4 else "" for s in shares],
            textposition="inside", insidetextanchor="middle",
            textfont=dict(size=9, color="#FFFFFF"),
        ))
        meta_series.append({"name": str(name), "color": colors[i],
                            "shares": shares})
    fig.update_layout(
        **_layout(), barmode="stack", barnorm="percent", bargap=BAR_GAP,
        legend=dict(orientation="h", y=-0.12, x=0, font=dict(size=8.5)),
        xaxis=dict(tickfont=dict(size=12), type="category"),
        yaxis=dict(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                   tickformat="~s"),
    )
    fig._chart_meta = {"chart": "pct_stack", "years": xs, "series": meta_series}
    return fig

def strip_brackets(names):
    """Names with their trailing '( ... )' removed - 'Grupo Bimbo (Bimbo,
    Marinela, ...)' reads 'Grupo Bimbo' on a chart or table axis.  Left as
    they are wherever stripping would leave a name empty or make two names
    identical."""
    import re
    names = [str(n) for n in names]
    cut = [re.sub(r"\s*\([^()]*\)?\s*$", "", n).strip() or n for n in names]
    return cut if len(set(cut)) == len(cut) else names


def _dedupe_labels(names, cap=120):
    """Shorten long category labels by stripping the longest word-prefix shared
    by ≥2 of them (e.g. '1.1% Sodium Fluoride with '), keeping each label unique
    and short.  Names that don't carry the shared prefix are left intact.  Fully
    generic - nothing market-specific is assumed."""
    from collections import Counter
    names = strip_brackets(names)
    toks = [n.split() for n in names]
    prefix = []
    while True:
        idx = len(prefix)
        cand = [t[idx] for t in toks if len(t) > idx and t[:idx] == prefix]
        if len(cand) < 2:
            break
        word, cnt = Counter(cand).most_common(1)[0]
        if cnt < 2:
            break
        prefix.append(word)
    plen = len(prefix)
    out = []
    for t in toks:
        rest = " ".join(t[plen:]) if (plen and t[:plen] == prefix) else " ".join(t)
        rest = rest or "(base)"
        if len(rest) > cap:
            rest = rest[: cap - 1].rstrip() + "…"
        out.append(rest)
    # guard: if stripping collapsed any duplicates, fall back to originals
    if len(set(out)) < len(out):
        return [n if len(n) <= cap else n[: cap - 1].rstrip() + "…" for n in names]
    return out

def composition_bps(val_tbl, years, bps, unit_label="", title=""):
    """100%-stacked value-composition columns (one full-height bar per year) on
    the LEFT axis, with a second BPS scale on the RIGHT y-axis.  Each category's
    2026→2036 share change is plotted as a green ▲ / red ▼ marker read against
    that right BPS axis - so 'share mix' and 'share momentum' share one chart."""
    order = list(val_tbl[years[-1]].sort_values(ascending=False).index)
    colors = _pal(len(order))
    xpos = list(range(len(years)))
    totals = {y: (sum(float(val_tbl.loc[nm, y]) for nm in order) or 1.0)
              for y in years}
    fig = go.Figure()
    seg_shares = {}
    for i, name in enumerate(order):
        yvals = [float(val_tbl.loc[name, y]) for y in years]
        shares = [v / totals[y] * 100.0 for v, y in zip(yvals, years)]
        seg_shares[name] = shares
        fig.add_trace(go.Bar(
            x=xpos, y=yvals, name=str(name), yaxis="y",
            marker=dict(color=colors[i], line=dict(color=PANEL_BG, width=1)),
            # carry the category name so a click can drive the By-Region
            # drill-down / drill-through (harmless where nothing listens)
            customdata=[str(name)] * len(xpos),
            text=[f"{s:.0f}%" if s >= 5 else "" for s in shares],
            textposition="inside",
            insidetextanchor="middle", textfont=dict(size=9, color="#FFFFFF"),
        ))
    # BPS markers on the right axis, in the gap to the right of the last column
    bvals = [float(v) for v in bps.values]
    # the ▲/▼ BPS markers are the ones the read-out bullets beside this chart
    # call "green ▲" and "red ▼", so they take the CLASSIFICATION colours, not
    # the TREND_UP/TREND_DOWN line accents (blue/red - the ▲ read blue)
    UP, DOWN = OUTPACE_COLOR, DECLINE_COLOR
    bcolors = [UP if v >= 0 else DOWN for v in bvals]
    # the MARKER is the arrow - a green ▲ for a share gain, a red ▼ for a
    # loss - so the label beside it carries the number only (it used to repeat
    # the glyph, which read as two arrows once the marker became one)
    bsymbols = ["triangle-up" if v >= 0 else "triangle-down" for v in bvals]
    btext = [f"{v:+,.0f}" for v in bvals]
    mx = len(years) - 1 + 0.55
    fig.add_trace(go.Scatter(
        x=[mx] * len(bvals), y=bvals, yaxis="y2", mode="markers+text",
        marker=dict(color=bcolors, size=11, symbol=bsymbols,
                    line=dict(color=PANEL_BG, width=1)),
        text=btext, textposition="middle right", cliponaxis=False,
        textfont=dict(size=8.5), showlegend=False, hoverinfo="skip",
    ))
    blim = (max((abs(v) for v in bvals), default=1.0) or 1.0) * 1.3
    lay = _layout(title)
    fig.update_layout(
        **lay, barmode="stack", barnorm="percent", bargap=BAR_GAP,
        showlegend=True,
        legend=dict(orientation="h", y=-0.14, x=0, font=dict(size=8)),
        xaxis=dict(tickmode="array", tickvals=xpos,
                   ticktext=[str(y) for y in years],
                   range=[-0.6, len(years) - 1 + 1.25], tickfont=dict(size=12)),
        yaxis=dict(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                   ticksuffix="%",
                   title=dict(text="Share of value (%)",
                              font=dict(size=9, color=MUTED))),
        yaxis2=dict(overlaying="y", side="right", range=[-blim, blim],
                    zeroline=True, zerolinecolor=MUTED, zerolinewidth=1,
                    showgrid=False, tickfont=dict(size=8),
                    title=dict(text="BPS  (2026→2036)",
                               font=dict(size=13, color=ACCENT))),
    )
    fig._chart_meta = dict(
        chart="comp_bps", years=[str(y) for y in years], pct=True,
        series=[dict(name=str(n), color=colors[i],
                     vals=[float(val_tbl.loc[n, y]) for y in years],
                     shares=seg_shares[n])
                for i, n in enumerate(order)],
        vmax=100.0,
        snames=_dedupe_labels([str(n) for n in order]),
        bnames=_dedupe_labels([str(i) for i in bps.index]),
        bvals=bvals, bcolors=bcolors, btext=btext, bsymbols=bsymbols,
        blim=blim,
        unit=unit_label)
    return fig

def share_donut(share, label):
    share = max(0.0, min(float(share), 1.0))
    fig = go.Figure(go.Pie(
        values=[share, 1 - share], labels=[label, "Rest"],
        hole=0.72, sort=False, direction="clockwise",
        marker=dict(colors=[ACCENT, "#E2F1F9"]),
        textinfo="none", hoverinfo="label+percent",
    ))
    fig.update_layout(**_layout(), showlegend=False)
    fig.add_annotation(text=f"{share:.1%}", showarrow=False,
                       font=dict(size=17, color=ACCENT, family="Segoe UI"),
                       yshift=6)
    fig.add_annotation(text=label, showarrow=False, yshift=-14,
                       font=dict(size=9, color=MUTED))
    return fig

def _period_bracket(fig, columns, hist_end):
    """Add a Historical / Forecast period bracket ABOVE the plot: two labelled
    horizontal rules (split at ``hist_end``) with down-ticks, plus a faint
    dotted divider through the chart.  Positioned in x-DOMAIN units so it lines
    up with the year categories regardless of the data."""
    yrs = [int(c) for c in columns]
    n = len(yrs)
    hc = sum(1 for y in yrs if y <= hist_end)
    if hc <= 0 or hc >= n:
        return
    bx = hc / n                                    # boundary in x-domain [0,1]
    yL, yTick, yLab = 1.05, 1.02, 1.15
    D = dict(xref="x domain", yref="paper")

    def rule(x0, x1, col):
        fig.add_shape(type="line", x0=x0, x1=x1, y0=yL, y1=yL,
                      line=dict(color=col, width=2.5), **D)
        for xc in (x0, x1):                        # down-ticks -> bracket look
            fig.add_shape(type="line", x0=xc, x1=xc, y0=yTick, y1=yL,
                          line=dict(color=col, width=2.5), **D)

    rule(0.004, bx - 0.012, MUTED)
    rule(bx + 0.012, 0.996, ACCENT)
    fig.add_shape(type="line", x0=bx, x1=bx, y0=0, y1=1,
                  line=dict(color="#D0D7DE", width=1, dash="dot"), **D)
    fig.add_annotation(x=bx / 2, y=yLab, showarrow=False,
                       text=f"Historical  {yrs[0]}–{hist_end}",
                       font=dict(size=9, color=MUTED), **D)
    fig.add_annotation(x=(bx + 1) / 2, y=yLab, showarrow=False,
                       text=f"Forecast  {hist_end + 1}–{yrs[-1]}",
                       font=dict(size=9, color=ACCENT), **D)

# ── Stacked "ribbon" (PBIX ribbonChart) ──────────────────────────────
def ribbon(tbl: pd.DataFrame, title="", pct=False, legend_size=9,
           period_hist_end=None):
    order = tbl[tbl.columns[0]].sort_values(ascending=False).index
    fig = go.Figure()
    colors = _pal(len(order))
    for i, name in enumerate(order):
        fig.add_trace(go.Bar(
            x=[str(c) for c in tbl.columns], y=tbl.loc[name].values,
            name=str(name),
            marker=dict(color=colors[i],
                        line=dict(color=PANEL_BG, width=1)),
        ))
    fig.update_layout(
        **_layout(title), barmode="stack", bargap=0.35,
        legend=dict(orientation="h", y=-0.14, x=0,
                    font=dict(size=legend_size)),
        xaxis=dict(tickfont=dict(size=9)),
        yaxis=dict(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                   tickformat=".0%" if pct else "~s"),
    )
    if period_hist_end is not None:
        fig.update_layout(margin=dict(l=8, r=8, t=44, b=8))  # room for bracket
        _period_bracket(fig, tbl.columns, period_hist_end)
        yrs = [int(c) for c in tbl.columns]
        hc = sum(1 for y in yrs if y <= period_hist_end)
        fig._chart_meta = {"period_band": {
            "hist_end": period_hist_end, "first": yrs[0], "last": yrs[-1],
            "hc": hc, "n": len(yrs)}}
    return fig

# ── YoY lines ────────────────────────────────────────────────────────
def yoy_lines(tbl: pd.DataFrame, title=""):
    fig = go.Figure()
    colors = _pal(len(tbl))
    for i, name in enumerate(tbl.index):
        fig.add_trace(go.Scatter(
            x=[str(c) for c in tbl.columns], y=tbl.loc[name].values,
            name=str(name), mode="lines+markers",
            line=dict(color=colors[i], width=2), marker=dict(size=4),
        ))
    fig.update_layout(
        **_layout(title),
        legend=dict(orientation="h", y=-0.15, x=0, font=dict(size=9)),
        xaxis=dict(tickfont=dict(size=9)),
        yaxis=dict(showgrid=True, gridcolor=GRID, ticksuffix="%",
                   tickfont=dict(size=9)),
    )
    return fig

# ── Trend lines (absolute values, non-stacked) ───────────────────────
def trend_lines(tbl: pd.DataFrame, title="", legend_size=9):
    """One line per row over the year columns - absolute values, NOT stacked
    (used for the scenario-trajectory comparison)."""
    fig = go.Figure()
    colors = _pal(len(tbl))
    for i, name in enumerate(tbl.index):
        fig.add_trace(go.Scatter(
            x=[str(c) for c in tbl.columns], y=tbl.loc[name].values,
            name=str(name), mode="lines+markers",
            line=dict(color=colors[i], width=2.5), marker=dict(size=4),
        ))
    fig.update_layout(
        **_layout(title),
        legend=dict(orientation="h", y=-0.15, x=0,
                    font=dict(size=legend_size)),
        xaxis=dict(tickfont=dict(size=9)),
        yaxis=dict(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                   tickformat="~s"),
    )
    return fig

# ── Trend bars (absolute values, grouped-bar twin of trend_lines) ────
def trend_bars(tbl: pd.DataFrame, title="", legend_size=9):
    """One bar-group per year column, one bar per row - absolute values,
    grouped (NOT stacked).  Bar-chart twin of trend_lines() for callers that
    want the same trajectory comparison rendered as bars."""
    fig = go.Figure()
    colors = _pal(len(tbl))
    for i, name in enumerate(tbl.index):
        fig.add_trace(go.Bar(
            x=[str(c) for c in tbl.columns], y=tbl.loc[name].values,
            name=str(name), marker=dict(color=colors[i]),
            # A value on every bar.  Years x scenarios leaves each bar far too
            # narrow for a horizontal number (they used to collide into
            # "143.0K43.0K43.0K"), so the labels stand on end, bottom-to-top:
            # rotated, a value costs only its glyph height in width and
            # nothing overlaps.  export_native._add_cat_chart rotates the PPT
            # twin the same way, and `~s` matches the axis tick format so the
            # two surfaces read identically.
            # '.3~s': 3 significant digits, SI-abbreviated, trailing zeros
            # trimmed -> "110", "220", "143k".  Plain '~s' would print the raw
            # "109.66"/"219.909" and not match the deck, whose labels are
            # rounded by export_native._num_format.
            texttemplate="%{y:.3~s}", textposition="outside", textangle=-90,
            textfont=dict(size=7.5), cliponaxis=False,
        ))
    top = float(pd.DataFrame(tbl).max().max() or 0)
    fig.update_layout(
        **_layout(title), barmode="group", bargap=BAR_GAP,
        uniformtext=dict(mode="show", minsize=6),
        legend=dict(orientation="h", y=-0.15, x=0,
                    font=dict(size=legend_size)),
        xaxis=dict(tickfont=dict(size=9)),
        # headroom so the upright labels clear the top of the plot area
        yaxis=dict(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                   tickformat="~s",
                   range=[0, top * 1.30] if top > 0 else None),
    )
    return fig

# ── Waterfall (PBIX waterfallChart) ──────────────────────────────────
def waterfall(wf: pd.DataFrame, title="", value_fmt="{:,.0f}",
              tick_size=8.5, bargap=None):
    fig = go.Figure(go.Waterfall(
        x=wf["label"], y=wf["value"], measure=wf["measure"],
        text=[value_fmt.format(v) for v in wf["value"]],
        textposition="outside", textfont=dict(size=8, color=MUTED),
        connector=dict(line=dict(color=FAINT, width=1)),
        increasing=dict(marker=dict(color=GOOD)),
        decreasing=dict(marker=dict(color=BAD)),
        totals=dict(marker=dict(color=ACCENT)),
    ))
    fig.update_layout(
        **_layout(title), showlegend=False,
        xaxis=dict(tickfont=dict(size=tick_size), type="category"),
        yaxis=dict(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                   tickformat="~s"),
    )
    if bargap is not None:
        fig.update_layout(bargap=bargap)
    return fig

# ── Variance bars around zero ────────────────────────────────────────
def variance_bars(s: pd.Series, suffix="BPS"):
    names, vals = list(s.index), [v * 10_000 for v in s.values]
    colors = [GOOD if v >= 0 else BAD for v in vals]
    fig = go.Figure(go.Bar(
        y=names, x=vals, orientation="h",
        marker=dict(color=colors, cornerradius=3),
        text=[f"{v:+,.0f} {suffix}" for v in vals],
        textposition="outside", textfont=dict(size=9),
    ))
    fig.add_vline(x=0, line_width=1, line_color=FAINT)
    m = max(abs(v) for v in vals) * 1.55 + 1
    fig.update_layout(
        **_layout(), bargap=BAR_GAP,
        xaxis=dict(visible=False, range=[-m, m]),
        yaxis=dict(autorange="reversed", tickfont=dict(size=9.5)),
    )
    return fig

_DIV_POS, _DIV_NEG = DIV_POS, DIV_NEG               # growth / decline

# ── Diverging vertical bars (green up / cyan down, centred zero) ─────
def diverging_bars(series, suffix="%", title="", baseline_label=None,
                   hover=None):
    """Vertical bars diverging from a centred, dashed baseline - green above /
    cyan below - for a signed metric per category.  Clean (no gridlines);
    carries chart-meta for native export.

    The values are already SIGNED DEVIATIONS, so the baseline is that metric's
    zero.  ``baseline_label`` names what the line represents (the pricing page
    passes "Global"): the middle tick then reads that word instead of "0",
    because on a benchmark chart the line means "same as Global", not "zero".
    ``hover`` is an optional per-category tooltip string for the dashboard.

    The range is symmetric about the baseline and derived from the data
    (max |value| x 1.35), so it adapts to whatever spread a workbook produces
    and always leaves room both above and below."""
    names = [str(i) for i in series.index]
    vals = [float(v) for v in series.values]
    colors = [_DIV_POS if v >= 0 else _DIV_NEG for v in vals]
    m = (max((abs(v) for v in vals), default=1.0) or 1.0) * 1.35
    bar = dict(
        x=names, y=vals, width=BAR_WIDTH, marker=dict(color=colors),
        text=[f"{v:+.1f}{suffix}" for v in vals], textposition="outside",
        textfont=dict(size=8.5, color=TEXT), cliponaxis=False)
    if hover:
        bar["customdata"] = list(hover)
        bar["hovertemplate"] = "%{customdata}<extra></extra>"
    fig = go.Figure(go.Bar(**bar))
    fig.add_hline(y=0, line=dict(color=MUTED, width=1.2, dash="dash"))
    hi = round(m * 0.85, 1)
    ticks = {"tickvals": [-hi, 0, hi]}
    if baseline_label:                      # name the line instead of "0"
        ticks["ticktext"] = [f"{-hi:+.1f}{suffix}", baseline_label,
                             f"{hi:+.1f}{suffix}"]
    fig.update_layout(
        **_layout(title), showlegend=False, bargap=BAR_GAP,
        xaxis=dict(tickfont=dict(size=8.5), type="category", showgrid=False,
                   showline=False, ticks=""),
        yaxis=dict(range=[-m, m], zeroline=False, showgrid=False,
                   tickfont=dict(size=8.5),
                   ticksuffix="" if baseline_label else suffix, **ticks))
    fig._chart_meta = {"chart": "diverging", "names": names, "vals": vals,
                       "colors": colors, "m": m, "suffix": suffix,
                       "blabel": baseline_label or ""}
    return fig

# ── Radar / spider (pricing-driver impact profile) ──────────────────
_RADAR_LINE = RADAR_LINE                            # brand outline
_RADAR_FILL = RADAR_FILL                            # translucent area

def radar(cats, vals, scale_max=10.0, title="", color=_RADAR_LINE,
          fillcolor=_RADAR_FILL, markers=True, scale_labels=True,
          gridshape="linear", show_shares=False):
    """Filled radar/spider chart - one spoke per category, the polygon reaching
    ``val`` on a 0..``scale_max`` axis.  Clean FMI styling: blue-toned filled
    area, light-grey rings, category labels, no legend.

    ``markers`` puts a small dot at each vertex; ``scale_labels`` shows a faint
    0..max scale along the top spoke; ``gridshape`` ('linear' = straight-edged
    decagon rings, like the reference; 'circular' = round rings).  Carries
    ``_chart_meta`` so the PPT export rebuilds it as an editable native polygon.

    ``show_shares`` appends each spoke's SHARE of the total score to its
    category label ("Input Cost  13%").  The share is baked into the label
    text, so the native PPT twin picks it up from ``_chart_meta['cats']``
    with no renderer change and the two surfaces stay identical.
    """
    cats = [str(c) for c in cats]
    vals = [float(v) for v in vals]
    if show_shares:
        tot = sum(abs(v) for v in vals) or 1.0
        cats = [f"{c}  {abs(v) / tot * 100:.1f}%" for c, v in zip(cats, vals)]
    r = vals + vals[:1]                            # close the loop for plotly
    theta = cats + cats[:1]
    mode = "lines+markers" if markers else "lines"
    fig = go.Figure(go.Scatterpolar(
        r=r, theta=theta, fill="toself", fillcolor=fillcolor,
        line=dict(color=color, width=2.2),
        marker=(dict(size=6, color=color,
                     line=dict(color=PANEL_BG, width=1)) if markers else None),
        mode=mode, hovertemplate="%{theta}: %{r:.1f}<extra></extra>"))
    step = 2.0 if scale_max <= 10 else scale_max / 5.0
    tickvals = [round(step * i, 2) for i in range(1, int(scale_max / step) + 1)]
    fig.update_layout(
        autosize=True, paper_bgcolor=PANEL_BG, plot_bgcolor=PANEL_BG,
        font=FONT, showlegend=False,
        margin=dict(l=66, r=66, t=30 if title else 18, b=26),
        title=dict(text=title, font=dict(size=11.5, color=TEXT), x=0.01)
        if title else None,
        polar=dict(
            bgcolor=PANEL_BG, gridshape=gridshape,
            radialaxis=dict(range=[0, scale_max], showticklabels=scale_labels,
                            tickvals=tickvals, angle=90, tickangle=0,
                            tickfont=dict(size=7.5, color=MUTED),
                            ticks="", gridcolor=GRID, gridwidth=1,
                            linecolor=GRID),
            angularaxis=dict(rotation=90, direction="clockwise",
                             gridcolor=GRID, linecolor=GRID,
                             tickfont=dict(size=9.5, color=TEXT))))
    fig._chart_meta = {"chart": "radar", "cats": cats, "vals": vals,
                       "scale_max": float(scale_max), "color": color,
                       "markers": bool(markers),
                       "scale_labels": bool(scale_labels),
                       "tickvals": tickvals}
    return fig

# ── 3-D industry-structure funnel (cone + cylinder discs) ───────────
_F3_CONE, _F3_BODY = F3_CONE, F3_BODY
_F3_TOP, _F3_RIM, _F3_LBL = F3_TOP, F3_RIM, F3_LBL

def funnel3d(segments, unit_label="", title=""):
    """A 3-D funnel: a light-green narrowing cone with a blue cylinder 'disc'
    per segment (lighter lid / body / darker bottom rim), tier + size above each
    disc and the share % on it.  ``segments`` = [{name, size, pct, w}] top→bottom
    (``w`` = relative width 0..1).  Carries chart-meta for the native twin."""
    n = len(segments) or 1
    top_y, bot_y = 96.0, 5.0
    slot = (top_y - bot_y) / n
    maxhw = 44.0
    shapes, annos = [], []

    def hw(i):
        return segments[i]["w"] * maxhw

    ry = min(2.8, slot * 0.16)
    bh0 = slot * 0.24
    # green cone - hugs each disc (constant width across the disc, tapering
    # only in the gaps between discs), drawn below everything
    right = [(50 + hw(0), top_y)]
    for i in range(n):
        cyi = top_y - (i + 0.5) * slot
        right.append((50 + hw(i), cyi + bh0 + ry))        # disc top (lid)
        right.append((50 + hw(i), cyi - bh0 - ry))        # disc bottom (rim)
    right.append((50 + hw(n - 1), bot_y))
    left = [(50 - x, y) for x, y in reversed(right)]
    path = "M " + " L ".join(f"{x:.2f} {y:.2f}" for x, y in right + left) + " Z"
    shapes.append(dict(type="path", path=path, fillcolor=_F3_CONE,
                       line=dict(width=0), layer="below"))
    for i, s in enumerate(segments):
        cy = top_y - (i + 0.5) * slot
        bh = slot * 0.24                              # half body height
        x0, x1 = 50 - hw(i), 50 + hw(i)
        shapes.append(dict(type="rect", x0=x0, x1=x1, y0=cy - bh, y1=cy + bh,
                           fillcolor=_F3_BODY, line=dict(width=0)))
        shapes.append(dict(type="circle", x0=x0, x1=x1, y0=cy - bh - ry,
                           y1=cy - bh + ry, fillcolor=_F3_RIM,
                           line=dict(width=0)))
        shapes.append(dict(type="circle", x0=x0, x1=x1, y0=cy + bh - ry,
                           y1=cy + bh + ry, fillcolor=_F3_TOP,
                           line=dict(color="#FFFFFF", width=1)))
        annos.append(dict(x=50, y=cy + bh + ry + slot * 0.18,
                          text=f"<b>{s['name']}</b>  ·  {s['size']}",
                          showarrow=False, xanchor="center", yanchor="middle",
                          font=dict(size=10, color=_F3_LBL)))
        annos.append(dict(x=50, y=cy, text=f"<b>{s['pct']}</b>",
                          showarrow=False, xanchor="center", yanchor="middle",
                          font=dict(size=9.5, color="#FFFFFF")))
    lay = _layout(title)
    lay["shapes"] = shapes
    lay["annotations"] = annos
    fig = go.Figure()
    fig.update_layout(
        **lay, xaxis=dict(range=[0, 100], visible=False, fixedrange=True),
        yaxis=dict(range=[0, 100], visible=False, fixedrange=True))
    fig._chart_meta = {"chart": "funnel3d", "segments": segments,
                       "unit": unit_label}
    return fig


# ── Industry-structure tier funnel (monotonic silhouette) ───────────
def tier_funnel(bands, unit_label="", title=""):
    """A centred tier funnel - an ordered list of ``bands`` drawn as stacked,
    monotonically NARROWING centred bars (a clean funnel silhouette), each
    labelled with its name, share and size.  ``bands`` = [{name, sharetext,
    size, color, label}] top→bottom (e.g. Total Market → Tier 1 ... → Others).
    Widths are by rank (funnel look); the true share/size live in the labels.
    Carries ``_chart_meta`` for a native editable twin."""
    n = len(bands)
    if n == 0:
        n = 1
    widths = [1.0 - i * (0.66 / max(1, n - 1)) for i in range(len(bands))]
    names = [b["name"] for b in bands]
    colors = [b.get("color", "#1668A8") for b in bands]
    text = [f"{b['name']}  ·  {b.get('sharetext', '')}  ·  "
            f"{_abbr_lbl(b.get('size', 0))} {unit_label}".rstrip(" ·").rstrip()
            for b in bands]
    fig = go.Figure(go.Bar(
        x=widths, y=list(range(len(bands))), base=[-w / 2 for w in widths],
        orientation="h", width=0.72,
        marker=dict(color=colors, line=dict(color=PANEL_BG, width=1.5)),
        text=text, textposition="inside", insidetextanchor="middle",
        textfont=dict(size=11, color="white"), cliponaxis=False,
        hoverinfo="text"))
    fig.update_layout(
        **_layout(title), showlegend=False, bargap=0.34,
        xaxis=dict(range=[-0.56, 0.56], visible=False),
        yaxis=dict(autorange="reversed", showgrid=False, showticklabels=False,
                   ticks=""))
    fig._chart_meta = {"chart": "tier_funnel", "unit": unit_label,
                       "bands": [{"name": b["name"],
                                  "sharetext": b.get("sharetext", ""),
                                  "size": b.get("size", 0),
                                  "label": b.get("label", ""),
                                  "color": b.get("color", "#1668A8")}
                                 for b in bands]}
    return fig

# ── Positioning bubble scatter (e.g. competitive win position) ──────
def pos_bubble(points, title="", xlab="Win Rate %", ylab="ASP vs Peer",
               xmax=3.5, ymax=4.0):
    """A few labelled bubbles on an x/y positioning grid (win rate vs ASP).
    ``points`` = [{name, x, y, color, size}].  Carries chart-meta for native."""
    fig = go.Figure()
    for p in points:
        fig.add_trace(go.Scatter(
            x=[p["x"]], y=[p["y"]], mode="markers+text",
            marker=dict(size=p.get("size", 26), color=p["color"],
                        line=dict(color=PANEL_BG, width=1.5)),
            text=["  " + p["name"]], textposition="middle right",
            textfont=dict(size=10, color=TEXT), cliponaxis=False,
            hoverinfo="text", hovertext=p["name"]))
    lay = _layout(title)
    lay["showlegend"] = False
    fig.update_layout(
        **lay,
        xaxis=dict(range=[0, xmax], title=dict(text=xlab, font=dict(size=9)),
                   tickfont=dict(size=8), showgrid=True, gridcolor=GRID,
                   zeroline=False),
        yaxis=dict(range=[0, ymax], title=dict(text=ylab, font=dict(size=9)),
                   tickfont=dict(size=8), showgrid=True, gridcolor=GRID,
                   zeroline=False))
    fig._chart_meta = {"chart": "posbubble", "points": points, "xmax": xmax,
                       "ymax": ymax, "xlab": xlab, "ylab": ylab}
    return fig


# ── Left-aligned TAM/SAM/SOM funnel (single-line labels) ────────────
def sam_funnel(rows, title=""):
    """A left-aligned funnel - descending-width bars (widest on top), each with
    a single-line white label inside.  ``rows`` = [{text, color}]."""
    n = len(rows) or 1
    widths = [100 - i * (46.0 / max(1, len(rows) - 1)) for i in range(len(rows))]
    fig = go.Figure(go.Bar(
        y=list(range(len(rows))), x=widths, orientation="h", base=0,
        width=0.68, marker=dict(color=[r["color"] for r in rows],
                                line=dict(color=PANEL_BG, width=1.5)),
        text=[r["text"] for r in rows], textposition="inside",
        insidetextanchor="start", textfont=dict(size=11.5, color="white"),
        cliponaxis=False, hoverinfo="text"))
    fig.update_layout(
        **_layout(title), showlegend=False, bargap=0.3,
        xaxis=dict(range=[0, 100], visible=False),
        yaxis=dict(range=[-0.5, len(rows) - 0.5], autorange="reversed",
                   visible=False))
    fig._chart_meta = {"chart": "samfunnel", "rows": rows}
    return fig


# ── Target vs Actual horizontal grouped bars (GTM benchmarks) ───────
_TA_TARGET, _TA_ACTUAL = TA_TARGET, TA_ACTUAL

def target_actual_bars(rows, title="", vmax=5.0):
    """Horizontal grouped bars - Target (dark) vs Actual (light) per metric on a
    shared index scale.  ``rows`` = [{name, target, actual}]."""
    names = [r["name"] for r in rows]
    fig = go.Figure()
    fig.add_trace(go.Bar(y=names, x=[r["target"] for r in rows],
                         orientation="h", name="Target",
                         marker=dict(color=_TA_TARGET)))
    fig.add_trace(go.Bar(y=names, x=[r["actual"] for r in rows],
                         orientation="h", name="Actual",
                         marker=dict(color=_TA_ACTUAL)))
    lay = _layout(title)
    fig.update_layout(
        **lay, barmode="group", bargap=BAR_GAP, bargroupgap=0.12,
        legend=dict(orientation="h", x=1, xanchor="right", y=1.16,
                    font=dict(size=9)),
        xaxis=dict(range=[0, vmax], tickfont=dict(size=8), showgrid=True,
                   gridcolor=GRID, zeroline=False),
        yaxis=dict(tickfont=dict(size=9), autorange="reversed"))
    fig._chart_meta = {"chart": "target_actual", "rows": rows, "vmax": vmax}
    return fig


# ── Uniform-colour milestone bars (e.g. CAGR at key windows) ─────────
def milestone_bars(labels, vals, suffix="%", title="", color="#003E58"):
    """Value-labelled vertical bars in one brand colour - for milestone
    metrics such as the CAGR over successive windows."""
    fig = go.Figure(go.Bar(
        x=[str(l) for l in labels], y=list(vals), width=BAR_WIDTH,
        marker=dict(color=color, cornerradius=2),
        text=[f"{v:.1f}{suffix}" for v in vals], textposition="outside",
        textfont=dict(size=9.5, color=TEXT), cliponaxis=False))
    m = (max(vals) if vals else 1.0) * 1.32
    fig.update_layout(
        **_layout(title), showlegend=False, bargap=BAR_GAP,
        xaxis=dict(tickfont=dict(size=10.5), type="category"),
        yaxis=dict(range=[0, m], showgrid=True, gridcolor=GRID,
                   tickfont=dict(size=9), ticksuffix=suffix))
    return fig

# ── CAGR bars with reference line ────────────────────────────────────
def cagr_bars(s: pd.Series, ref: float, ref_label: str, horizontal=False,
              tick_size=9):
    names = list(s.index)
    vals = [v * 100 for v in s.values]
    colors = _pal(len(names))
    if horizontal:
        fig = go.Figure(go.Bar(
            y=names, x=vals, orientation="h",
            marker=dict(color=colors, cornerradius=3), width=BAR_WIDTH,
            text=[f"{v:.1f}%" for v in vals], textposition="outside",
            textfont=dict(size=9),
        ))
        fig.add_vline(x=ref * 100, line_dash="dash", line_color=REF_LINE,
                      line_width=2, annotation_text=ref_label,
                      annotation_font=dict(color=REF_LINE, size=9))
        fig.update_layout(**_layout(),
                          xaxis=dict(visible=False,
                                     range=[0, max(vals + [ref * 100]) * 1.3]),
                          yaxis=dict(autorange="reversed",
                                     tickfont=dict(size=tick_size)))
    else:
        fig = go.Figure(go.Bar(
            x=names, y=vals,
            marker=dict(color=colors, cornerradius=3), width=BAR_WIDTH,
            text=[f"{v:.1f}%" for v in vals], textposition="outside",
            textfont=dict(size=9),
        ))
        fig.add_hline(y=ref * 100, line_dash="dash", line_color=REF_LINE,
                      line_width=2, annotation_text=ref_label,
                      annotation_position="top right",
                      annotation_font=dict(color=REF_LINE, size=9))
        fig.update_layout(**_layout(),
                          xaxis=dict(tickfont=dict(size=tick_size),
                                     tickangle=-12),
                          yaxis=dict(visible=False,
                                     range=[0, max(vals + [ref * 100]) * 1.3]))
    return fig

# ── Country overview: year columns, CAGR + deviation, share split ─────
def year_columns(years, values, name, color=None):
    """One plain column per year - no previous-year companion, no figure on
    every bar (the first and last year carry theirs).  Native twin:
    export_native._render_year_columns."""
    xs = [str(y) for y in years]
    vals = [float(v) for v in values]
    fig = go.Figure(go.Bar(x=xs, y=vals, name=name,
                           marker=dict(color=color or _pal(1)[0])))
    fig.update_layout(**_layout(), showlegend=False,
                      xaxis=dict(type="category", tickfont=dict(size=9)),
                      yaxis=dict(showgrid=True, gridcolor=GRID,
                                 tickfont=dict(size=9), tickformat="~s"))
    fig._chart_meta = dict(chart="year_columns", years=xs, vals=vals,
                           name=name, color=color or _pal(1)[0])
    return fig


def cagr_dev(s: pd.Series, dev_bps: float, region: str, analysis: str):
    """CAGR comparison as HORIZONTAL bars merged with the deviation from the
    regional benchmark and the analysis that goes with it, one panel.

    A geography whose base-year value is 0 (no footprint yet in a market
    this granular) has an UNDEFINED CAGR - metrics.fore_cagr/
    growth_deviation_bps already return NaN for it rather than raising, so
    this is read as 0% growth here instead of crashing the chart the first
    time a country-level series is genuinely all zero (2026-10-22, Digital
    Identity Wallet & Personal Data Vault Services)."""
    names = [str(i) for i in s.index]
    vals = [0.0 if pd.isna(v) else float(v) * 100 for v in s.values]
    colors = _pal(len(names))
    fig = go.Figure(go.Bar(y=names, x=vals, orientation="h",
                           marker=dict(color=colors)))
    fig._chart_meta = dict(chart="cagr_dev", names=names, vals=vals,
                           colors=[str(c) for c in colors],
                           dev=0.0 if pd.isna(dev_bps) else float(dev_bps),
                           region=region, analysis=analysis)
    return fig


def share_split(rows, country, rest_name):
    """100% horizontal bars: ``country`` against ``rest_name`` for each
    (label, share-fraction) in ``rows`` - two colours, named, summing to the
    whole."""
    labels = [r[0] for r in rows]
    mine = [max(0.0, min(1.0, float(r[1]))) * 100 for r in rows]
    rest = [100.0 - v for v in mine]
    c1, c2 = _pal(2)[:2]
    fig = go.Figure()
    fig.add_trace(go.Bar(y=labels, x=mine, name=country, orientation="h",
                         marker=dict(color=c1),
                         text=[f"{v:.1f}%" for v in mine],
                         textposition="inside"))
    fig.add_trace(go.Bar(y=labels, x=rest, name=rest_name, orientation="h",
                         marker=dict(color=c2),
                         text=[f"{v:.1f}%" for v in rest],
                         textposition="inside"))
    fig.update_layout(**_layout(), barmode="stack", bargap=0.45,
                      xaxis=dict(range=[0, 100], ticksuffix="%",
                                 tickfont=dict(size=9), showgrid=False),
                      yaxis=dict(autorange="reversed",
                                 tickfont=dict(size=10)),
                      legend=dict(orientation="h", y=-0.2))
    return fig


# ── CAGR line (vertical, line-chart twin of cagr_bars) ───────────────
def cagr_line(s: pd.Series, ref: float, ref_label: str, tick_size=9):
    """One connected line across scalar per-category values (e.g. scenario
    CAGRs).  Line-chart twin of cagr_bars() (vertical variant only) for
    callers that want the same comparison rendered as a line instead of
    bars."""
    names = list(s.index)
    vals = [v * 100 for v in s.values]
    colors = _pal(len(names))
    fig = go.Figure(go.Scatter(
        x=names, y=vals, mode="lines+markers+text",
        line=dict(color=ACCENT, width=2.5),
        marker=dict(size=9, color=colors),
        text=[f"{v:.1f}%" for v in vals], textposition="top center",
        textfont=dict(size=9),
    ))
    fig.add_hline(y=ref * 100, line_dash="dash", line_color=REF_LINE,
                  line_width=2, annotation_text=ref_label,
                  annotation_position="top right",
                  annotation_font=dict(color=REF_LINE, size=9))
    fig.update_layout(**_layout(),
                      xaxis=dict(tickfont=dict(size=tick_size),
                                 tickangle=-12),
                      yaxis=dict(visible=False,
                                 range=[0, max(vals + [ref * 100]) * 1.3]))
    return fig

# ── Scenario forecast: base-year bar + three end-year scenario bars ──
# Colours follow the reference slide ("Scenario Outlook" charts): a deep
# base-year bar, then Conservative / Likely / Optimistic side by side at the
# horizon, each carrying a "CAGR: x%" callout on an arrow from the base bar.
SCENARIO_BAR_COLORS = {"Base": "#403152", "Conservative": "#92D050",
                       "Likely": "#0078D4", "Optimistic": "#808080"}
SCENARIO_LINE_COLORS = {"Optimistic": "#2E5F8A", "Likely": "#8B1A1A",
                        "Conservative": "#A08000"}


def scenario_bars(base_year, base_val, end_year, ends, cagrs):
    """Base-year value beside the three end-year scenario values.  The deck
    draws it natively from ``_chart_meta`` (see
    export_native._render_scenario_bars); the plotly twin is a plain grouped
    bar for the dashboard."""
    order = ["Conservative", "Likely", "Optimistic"]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=[str(base_year)], y=[base_val], name=str(base_year),
                         marker=dict(color=SCENARIO_BAR_COLORS["Base"]),
                         showlegend=False))
    for k in order:
        fig.add_trace(go.Bar(x=[str(end_year)], y=[ends[k]],
                             name=f"{k} Scenario",
                             marker=dict(color=SCENARIO_BAR_COLORS[k])))
    fig.update_layout(**_layout(), barmode="group",
                      legend=dict(orientation="h", y=-0.12, x=0),
                      yaxis=dict(visible=False))
    fig._chart_meta = {
        "chart": "scenario_bars", "base_year": str(base_year),
        "base": float(base_val), "end_year": str(end_year),
        "bars": [(k, float(ends[k]), float(cagrs[k])) for k in order],
        "colors": dict(SCENARIO_BAR_COLORS)}
    return fig


def scenario_lines(years, yoy):
    """Year-on-year growth of each scenario across the forecast years, one
    marker line per scenario on a percentage axis (no point labels)."""
    xs = [str(y) for y in years]
    order = ["Optimistic", "Likely", "Conservative"]
    fig = go.Figure()
    for k in order:
        fig.add_trace(go.Scatter(
            x=xs, y=[yoy[k][y] for y in years], mode="lines+markers",
            name=f"{k} Scenario",
            line=dict(color=SCENARIO_LINE_COLORS[k], width=2.5),
            marker=dict(size=7, color=SCENARIO_LINE_COLORS[k])))
    fig.update_layout(**_layout(),
                      legend=dict(orientation="h", y=-0.14, x=0),
                      yaxis=dict(tickformat=".2%", rangemode="tozero"))
    fig._chart_meta = {
        "chart": "scenario_lines", "years": xs,
        "series": [(f"{k} Scenario", [float(yoy[k][y]) for y in years],
                    SCENARIO_LINE_COLORS[k]) for k in order]}
    return fig

# ── Marimekko / Mekko (region width × segment height) ───────────────
def mekko(mat, unit_label="", title=""):
    """Marimekko: columns = children (width ∝ their market size), each split by
    segment (height ∝ within-column share).  Shows size AND mix at once."""
    cols = [str(c) for c in mat.columns]
    rows = [str(r) for r in mat.index]
    col_tot = [float(mat[c].sum()) for c in mat.columns]
    total = sum(col_tot) or 1.0
    widths = [ct / total * 100 for ct in col_tot]
    centers, acc = [], 0.0
    for w in widths:
        centers.append(acc + w / 2)
        acc += w
    colors = _pal(len(rows))
    heights = [[(float(mat.iloc[ri, ci]) / col_tot[ci] * 100 if col_tot[ci]
                 else 0.0) for ci in range(len(cols))] for ri in range(len(rows))]
    fig = go.Figure()
    for ri, r in enumerate(rows):
        fig.add_trace(go.Bar(
            x=centers, y=heights[ri], width=widths, name=r,
            marker=dict(color=colors[ri], line=dict(color="#FFFFFF", width=1)),
            text=[f"{v:.0f}%" if v >= 9 else "" for v in heights[ri]],
            textposition="inside", insidetextanchor="middle",
            textfont=dict(size=8, color="#FFFFFF"),
        ))
    fig.update_layout(
        **_layout(title), barmode="stack", bargap=0,
        legend=dict(orientation="h", y=-0.14, x=0, font=dict(size=8.5)),
        xaxis=dict(tickmode="array", tickvals=centers,
                   ticktext=[f"{c}<br><b>{ct:,.0f}</b>" for c, ct
                             in zip(cols, col_tot)],
                   tickfont=dict(size=8.5), range=[0, 100], showgrid=False),
        yaxis=dict(range=[0, 100], ticksuffix="%", tickfont=dict(size=9),
                   showgrid=False),
    )
    fig._chart_meta = dict(chart="mekko", cols=cols, rows=rows, widths=widths,
                           col_tot=col_tot, heights=heights,
                           colors=[str(c) for c in colors])
    return fig

def composition_bars(mat, unit_label="", title=""):
    """100%-stacked horizontal bars - a robust, no-overflow alternative to the
    Marimekko.  One row per column of ``mat`` (e.g. a region, largest first with
    its absolute size shown in the row label), split into its rows' shares (e.g.
    the formulation mix).  Because every group gets a full bar-height row, the
    labels can't be squeezed the way a Marimekko's thin columns are."""
    groups = [str(c) for c in mat.columns]
    parts = [str(r) for r in mat.index]
    totals = [float(mat[c].sum()) for c in mat.columns]
    order = sorted(range(len(groups)), key=lambda i: totals[i], reverse=True)

    def _short(v):
        floor = _abbrev_floor()
        for d, s in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
            if abs(v) >= d and floor <= d:
                return f"{v / d:.1f}{s}"
        return f"{v:,.0f}"

    cats = [f"{groups[i]}  ·  {_short(totals[i])}" for i in order]
    colors = _pal(len(parts))
    shares = [[(float(mat.iloc[pi, i]) / totals[i] * 100 if totals[i] else 0.0)
               for i in order] for pi in range(len(parts))]
    fig = go.Figure()
    for pi, p in enumerate(parts):
        fig.add_trace(go.Bar(
            y=cats, x=shares[pi], name=p, orientation="h",
            marker=dict(color=colors[pi], line=dict(color="#FFFFFF", width=1)),
            text=[f"{v:.0f}%" if v >= 7 else "" for v in shares[pi]],
            textposition="inside", insidetextanchor="middle",
            textfont=dict(size=8, color="#FFFFFF"),
        ))
    fig.update_layout(
        **_layout(title), barmode="stack", bargap=BAR_GAP,
        legend=dict(orientation="h", y=-0.16, x=0, font=dict(size=8.5)),
        xaxis=dict(range=[0, 100], ticksuffix="%", tickfont=dict(size=9),
                   showgrid=False),
        yaxis=dict(autorange="reversed", tickfont=dict(size=9), showgrid=False),
    )
    fig._chart_meta = dict(chart="composition", cats=cats, parts=parts,
                           shares=shares, colors=[str(c) for c in colors])
    return fig

def _lighten(hexcol, f=0.55):
    h = hexcol.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return "#%02X%02X%02X" % (int(r + (255 - r) * f), int(g + (255 - g) * f),
                              int(b + (255 - b) * f))

# ── Sunburst (hierarchy: geography → sub-geography) ──────────────────
def sunburst(labels, parents, values, colors=None, title=""):
    fig = go.Figure(go.Sunburst(
        labels=labels, parents=parents, values=values, branchvalues="total",
        marker=dict(colors=colors, line=dict(color="#FFFFFF", width=1.5))
        if colors else dict(line=dict(color="#FFFFFF", width=1.5)),
        textinfo="label+percent parent", insidetextorientation="radial",
        textfont=dict(size=9),
    ))
    lay = _layout(title)
    lay["margin"] = dict(l=4, r=4, t=26, b=4)
    fig.update_layout(**lay)
    return fig

# ── Country treemap (top-N countries, one colour each) ───────────────
def top_country_treemap(names, values, colors=None, unit_label="", title=""):
    """Flat squarified treemap - one tile per country, area ∝ market value,
    a distinct colour per country.  A rectangular, formal replacement for the
    round hierarchy sunburst; each tile shows the country, its value and its
    share of the shown total."""
    colors = list(colors) if colors is not None else _pal(len(names))
    total = float(sum(values)) or 1.0
    text = [f"{n}<br>{_abbr_lbl(v)}<br>{v / total * 100:.1f}%"
            for n, v in zip(names, values)]
    fig = go.Figure(go.Treemap(
        labels=list(names), parents=[""] * len(names), values=list(values),
        marker=dict(colors=colors, line=dict(color="#FFFFFF", width=2)),
        text=text, textinfo="text", texttemplate="%{text}",
        textposition="middle center",
        textfont=dict(size=11, color="#FFFFFF", family="Segoe UI"),
        tiling=dict(pad=2), sort=True, branchvalues="total",
        hovertemplate="<b>%{label}</b><br>%{value:,.0f} " + unit_label
                      + "<extra></extra>",
    ))
    lay = _layout(title)
    lay["margin"] = dict(l=4, r=4, t=26, b=4)
    fig.update_layout(**lay)
    return fig

# ── Heatmap (segment × year matrix, e.g. YoY %) ─────────────────────
# Single-hue sequential 'meter' gradient (ColorBrewer 'Blues' family, anchored
# to the FMI deep teal): white = lowest value → deep = highest.  Perceptually
# ordered and colour-blind-safe - the researched choice for a magnitude scale.
HEAT_SCALE = [list(stop) for stop in HEAT_STOPS]

def heatmap(tbl, title="", suffix="%"):
    """Colour-coded intensity matrix: rows = sub-segments, columns = years.
    A single-hue sequential 'meter' gradient (``HEAT_SCALE``) encodes magnitude -
    white = lowest, deep = highest - and a colour-bar legend on the right maps
    shade→value.  No numbers sit on the tiles (colour is the read-out); exact
    figures remain on hover.  A faint grid keeps the near-white tiles visible."""
    ys = [str(c) for c in tbl.columns]
    rows = [str(i) for i in tbl.index]
    z = tbl.astype(float).values
    fig = go.Figure(go.Heatmap(
        z=z, x=ys, y=rows, xgap=1, ygap=1, showscale=True,
        colorscale=HEAT_SCALE,
        colorbar=dict(thickness=10, len=0.86, x=1.012, xpad=0,
                      outlinewidth=0, ticks="outside", ticklen=3,
                      tickfont=dict(size=7.5, color=MUTED), nticks=5,
                      title=dict(text=suffix, side="top",
                                 font=dict(size=8, color=MUTED))),
        hovertemplate="%{y} · %{x}: %{z:.1f}" + suffix + "<extra></extra>",
    ))
    lay = _layout(title)
    lay["plot_bgcolor"] = "#D9DEE3"                 # gaps → subtle grid lines
    fig.update_layout(
        **lay,
        xaxis=dict(side="top", tickfont=dict(size=8.5), showgrid=False),
        yaxis=dict(autorange="reversed", tickfont=dict(size=9),
                   showgrid=False),
    )
    return fig

# ── Flow (Sankey) between two axes of a cross-section ────────────────
def sankey(matrix, unit_label="", title=""):
    """Flow diagram of an A x B matrix (``matrix`` rows = A, columns = B):
    every A sub-segment on the left, every B group on the right, a ribbon per
    cell as wide as its value.  Both sides are ordered largest first; a ribbon
    takes its A node's colour from the deck's palette.  The native deck draws
    it from the same trace (export_native._render_sankey)."""
    m = matrix.astype(float).clip(lower=0.0)
    a_order = list(m.sum(axis=1).sort_values(ascending=False).index)
    b_order = list(m.sum(axis=0).sort_values(ascending=False).index)
    cols = _pal(len(a_order))
    nodes = [str(a) for a in a_order] + [str(b) for b in b_order]
    ncol = list(cols) + ["#8CA3B8"] * len(b_order)
    src, dst, val, lcol = [], [], [], []
    for i, a in enumerate(a_order):
        for j, b in enumerate(b_order):
            v = float(m.loc[a, b])
            if v > 0:
                src.append(i)
                dst.append(len(a_order) + j)
                val.append(v)
                lcol.append(cols[i])
    fig = go.Figure(go.Sankey(
        arrangement="fixed",
        node=dict(label=nodes, color=ncol, pad=10, thickness=14),
        link=dict(source=src, target=dst, value=val, color=lcol),
        valueformat=",.0f", valuesuffix=f" {unit_label}".rstrip()))
    fig.update_layout(**_layout(title))
    return fig


# ── Slope chart (change between two points, per category) ────────────
def slope(tbl, x0_label, x1_label, title="", value_fmt="{:.1f}%"):
    """A slope graph: one line per row from ``x0_label`` to ``x1_label`` (two
    columns of ``tbl``) - shows which categories accelerate or decelerate."""
    colors = _pal(len(tbl))
    fig = go.Figure()
    for i, name in enumerate(tbl.index):
        y0, y1 = float(tbl.iloc[i, 0]), float(tbl.iloc[i, 1])
        fig.add_trace(go.Scatter(
            x=[x0_label, x1_label], y=[y0, y1], name=str(name),
            mode="lines+markers+text",
            text=[value_fmt.format(y0), value_fmt.format(y1)],
            textposition=["middle left", "middle right"],
            textfont=dict(size=9, color=colors[i]),
            line=dict(color=colors[i], width=2.5), marker=dict(size=7),
        ))
    fig.update_layout(
        **_layout(title), showlegend=True,
        legend=dict(orientation="h", y=-0.16, x=0, font=dict(size=8.5)),
        xaxis=dict(tickfont=dict(size=10), type="category",
                   range=[-0.35, 1.35]),
        yaxis=dict(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                   ticksuffix="%"),
    )
    return fig

# ── Paired grouped bars (e.g. H-CAGR vs F-CAGR shift) ────────────────
def paired_bars(tbl, name0, name1, title="", value_fmt="{:.1f}%"):
    """Two clustered HORIZONTAL bars per category (columns 0 and 1 of ``tbl``)
    - e.g. historical vs forecast CAGR per sub-segment.  Reads cleanly with
    long category labels and exports as a native clustered bar chart."""
    cats = [str(i) for i in tbl.index]
    s0 = [float(tbl.iloc[i, 0]) for i in range(len(tbl))]
    s1 = [float(tbl.iloc[i, 1]) for i in range(len(tbl))]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=cats, x=s0, orientation="h", name=name0,
        marker=dict(color=NEST_OUTER, cornerradius=3),
        text=[value_fmt.format(v) for v in s0], textposition="outside",
        textfont=dict(size=8, color=MUTED), cliponaxis=False))
    fig.add_trace(go.Bar(
        y=cats, x=s1, orientation="h", name=name1,
        marker=dict(color=ACCENT, cornerradius=3),
        text=[value_fmt.format(v) for v in s1], textposition="outside",
        textfont=dict(size=8, color=MUTED), cliponaxis=False))
    hi = max(s0 + s1 + [0]) * 1.28 or 1
    fig.update_layout(
        **_layout(title), barmode="group", bargap=BAR_GAP, bargroupgap=0.12,
        legend=dict(orientation="h", y=-0.1, x=0, font=dict(size=8.5)),
        xaxis=dict(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                   ticksuffix="%", range=[0, hi]),
        yaxis=dict(autorange="reversed", tickfont=dict(size=8.5)))
    return fig

# ── Dumbbell / connected-dot (change over a period) ─────────────────
def dumbbell(tbl, y0, y1, unit_label="", title="", x_floor=None):
    """A connected-dot ('dumbbell') chart - one row per category, a light dot at
    the ``y0`` value joined by a bar to a bold dot at the ``y1`` value: a
    consultant-style read of change over the period.  Sorted so the biggest
    category sits on top; carries chart-meta so the PPT export rebuilds it from
    native shapes rather than a default chart type.  ``x_floor`` starts the axis
    above 0 (zoom to the data band) when values are clustered far from zero."""
    order = list(tbl[y1].sort_values(ascending=True).index)
    cats = [str(i) for i in order]
    start = [float(tbl.loc[i, y0]) for i in order]
    end = [float(tbl.loc[i, y1]) for i in order]
    fig = go.Figure()
    for i in range(len(cats)):
        fig.add_trace(go.Scatter(
            x=[start[i], end[i]], y=[cats[i], cats[i]], mode="lines",
            line=dict(color="#CBD8E1", width=4), showlegend=False,
            hoverinfo="skip"))
    fig.add_trace(go.Scatter(
        x=start, y=cats, mode="markers", name=str(y0),
        marker=dict(color=NEST_OUTER, size=12,
                    line=dict(color="#7FBEDD", width=1.5))))
    fig.add_trace(go.Scatter(
        x=end, y=cats, mode="markers+text", name=str(y1),
        marker=dict(color=ACCENT, size=15),
        text=[_abbr_lbl(v) for v in end], textposition="middle right",
        textfont=dict(size=9, color=HEADER_BG), cliponaxis=False))
    peak = max(end + start) if end else 1.0
    lo = 0.0
    hi = peak * 1.28
    if x_floor is not None:                       # zoom to the data band
        lo = float(x_floor)
        hi = lo + (peak - lo) * 1.18
    fig.update_layout(
        **_layout(title), showlegend=True,
        legend=dict(orientation="h", y=-0.14, x=0, font=dict(size=9)),
        xaxis=dict(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                   tickformat="~s", range=[lo, hi], zeroline=False),
        yaxis=dict(tickfont=dict(size=8.5)))
    fig._chart_meta = {"chart": "dumbbell", "cats": cats, "start": start,
                       "end": end, "y0": str(y0), "y1": str(y1),
                       "unit": unit_label, "x0": lo, "xmax": hi}
    return fig

# ── Horizontal box plot (distribution per category) ─────────────────
def price_box(price_tbl, unit_label="", title=""):
    """A VERTICAL box plot - one box per row of ``price_tbl`` (index = category,
    columns = the periods) standing on a shared value axis: min / Q1 / median /
    Q3 / max of that category's value across the periods.  Carries chart-meta so
    the PPT export rebuilds it from native shapes."""
    order = list(price_tbl.index)                   # keep the sub-segment order
    short = _dedupe_labels([str(i) for i in order], cap=22)  # tidy x labels
    cats = list(short)
    fill, line = "#CFE7F5", HEADER_BG
    fig = go.Figure()
    stats = []
    lo_all, hi_all = [], []
    for idx, name in enumerate(order):
        vals = [float(v) for v in price_tbl.loc[name].values
                if v == v]                          # drop NaN
        if not vals:
            vals = [0.0]
        fig.add_trace(go.Box(
            y=vals, name=short[idx], width=0.5, whiskerwidth=0.5,
            marker=dict(color=ACCENT, size=4), line=dict(color=line, width=1.5),
            fillcolor=fill, boxpoints=False, showlegend=False,
            hovertemplate="%{y:.2f}<extra>" + short[idx] + "</extra>"))
        s = pd.Series(vals)
        stats.append({"name": short[idx], "low": float(s.min()),
                      "q1": float(s.quantile(0.25)), "med": float(s.median()),
                      "q3": float(s.quantile(0.75)), "high": float(s.max())})
        lo_all.append(float(s.min()))
        hi_all.append(float(s.max()))
    lo, hi = (min(lo_all) if lo_all else 0), (max(hi_all) if hi_all else 1)
    pad = (hi - lo) * 0.15 or 1.0
    y0, y1 = max(0.0, lo - pad), hi + pad
    fig.update_layout(
        **_layout(title), showlegend=False,
        xaxis=dict(tickfont=dict(size=8.5), type="category"),
        yaxis=dict(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                   range=[y0, y1], zeroline=False,
                   title=dict(text=unit_label, font=dict(size=9, color=MUTED))))
    fig._chart_meta = {"chart": "boxplot", "cats": cats, "stats": stats,
                       "y0": y0, "y1": y1, "unit": unit_label}
    return fig

# ── Price OHLC candlestick (open/high/low/close per category) ───────
def candlestick(ohlc, unit_label="", title="", uniform_color=None):
    """A price candlestick - one candle per row of ``ohlc`` (index = category,
    columns open/high/low/close), thin wick low→high + a body open→close.  By
    default green body when close≥open / red otherwise; pass ``uniform_color``
    (fill hex) to render every candle in ONE steel-blue tone (FMI reference look).
    Carries chart-meta so the PPT export rebuilds it from native shapes."""
    order = list(ohlc.index)
    cats = _dedupe_labels([str(i) for i in order], cap=22)
    if uniform_color:
        up_f = dn_f = uniform_color
        up_l = dn_l = NEUTRAL_LINE
    else:
        up_f, up_l = UP_FILL, UP_LINE
        dn_f, dn_l = DN_FILL, DN_LINE
    fig = go.Figure()
    candles, lo_all, hi_all = [], [], []
    for idx, name in enumerate(order):
        o = float(ohlc.loc[name, "open"])
        c = float(ohlc.loc[name, "close"])
        hi = float(ohlc.loc[name, "high"])
        lo = float(ohlc.loc[name, "low"])
        up = c >= o
        fig.add_trace(go.Candlestick(
            x=[cats[idx]], open=[o], high=[hi], low=[lo], close=[c],
            increasing=dict(line=dict(color=up_l), fillcolor=up_f),
            decreasing=dict(line=dict(color=dn_l), fillcolor=dn_f),
            whiskerwidth=0.4, showlegend=False,
            hovertext=cats[idx], hoverinfo="y"))
        candles.append({"name": cats[idx], "open": o, "high": hi, "low": lo,
                        "close": c, "up": up})
        lo_all.append(lo)
        hi_all.append(hi)
    lo, hi = (min(lo_all) if lo_all else 0.0), (max(hi_all) if hi_all else 1.0)
    pad = (hi - lo) * 0.15 or 1.0
    y0, y1 = max(0.0, lo - pad), hi + pad
    fig.update_layout(
        **_layout(title), showlegend=False,
        xaxis=dict(tickfont=dict(size=8.5), type="category",
                   rangeslider=dict(visible=False)),
        yaxis=dict(showgrid=True, gridcolor=GRID, tickfont=dict(size=9),
                   range=[y0, y1], zeroline=False,
                   title=dict(text=unit_label, font=dict(size=9, color=MUTED))))
    fig._chart_meta = {"chart": "candlestick", "cats": cats, "candles": candles,
                       "y0": y0, "y1": y1, "unit": unit_label,
                       "uniform": uniform_color or ""}
    return fig

# ── Price-point strip / range-band dot plot (FMI "pricing pattern") ──
def price_strip(rows, x_min, x_max, unit_label="", series=None,
                series_colors=None, avg_only=False, title=""):
    """Horizontal price-point strip - one row per category (e.g. region) on a
    SHARED price axis.  Each row shows a light-blue price-RANGE band (lo→hi) plus
    either an orange AVERAGE-price dot (``avg_only=True``) or one coloured dot per
    ``series`` (product type).  Rebuilt natively for the PPT via chart-meta.

    ``rows`` = [{"name", "lo", "hi", "avg", "dots":[v0,v1,...]}] (dots align to
    ``series``)."""
    names = [str(r["name"]) for r in rows]
    n = len(rows)
    yv = list(range(n))
    band_fill, band_line = "#CFE7F5", "#7FBEDD"
    avg_col, avg_edge = "#F2994A", "#B9701F"
    series = list(series or [])
    cols = list(series_colors or (_pal(len(series)) if series else []))
    fig = go.Figure()
    shapes = [dict(type="rect", xref="x", yref="y", layer="below",
                   x0=r["lo"], x1=r["hi"], y0=i - 0.30, y1=i + 0.30,
                   fillcolor=band_fill, line=dict(color=band_line, width=1))
              for i, r in enumerate(rows)]
    # "Price Range" legend swatch (dummy trace)
    fig.add_trace(go.Scatter(
        x=[None], y=[None], mode="markers", name="Price Range",
        marker=dict(size=13, color=band_fill, symbol="square",
                    line=dict(color=band_line, width=1))))
    if avg_only:
        fig.add_trace(go.Scatter(
            x=[r["avg"] for r in rows], y=yv, mode="markers",
            name="Average Price", marker=dict(size=15, color=avg_col,
            line=dict(color=avg_edge, width=1.2))))
    else:
        for si, sname in enumerate(series):
            fig.add_trace(go.Scatter(
                x=[r["dots"][si] for r in rows], y=yv, mode="markers",
                name=sname, marker=dict(size=13, color=cols[si],
                line=dict(color="#FFFFFF", width=1))))
    fig.update_layout(
        **_layout(title), shapes=shapes,
        legend=dict(orientation="h", y=1.03, x=0, yanchor="bottom",
                    font=dict(size=9)),
        xaxis=dict(range=[x_min, x_max], showgrid=True, gridcolor=GRID,
                   tickfont=dict(size=9), zeroline=False,
                   title=dict(text=unit_label, font=dict(size=9, color=MUTED))),
        yaxis=dict(tickmode="array", tickvals=yv, ticktext=names,
                   autorange="reversed", tickfont=dict(size=9.5),
                   showgrid=False, zeroline=False))
    fig._chart_meta = {
        "chart": "price_strip", "x_min": float(x_min), "x_max": float(x_max),
        "unit": unit_label, "series": series, "colors": cols,
        "avg_only": bool(avg_only), "avg_color": avg_col, "avg_edge": avg_edge,
        "band_fill": band_fill, "band_line": band_line,
        "rows": [{"name": str(r["name"]), "lo": float(r["lo"]),
                  "hi": float(r["hi"]),
                  "avg": float(r.get("avg", (r["lo"] + r["hi"]) / 2)),
                  "dots": [float(v) for v in r.get("dots", [])]} for r in rows]}
    return fig

# ── Comparison bars (replaces the share donuts) ─────────────────────
def segment_bar(series, theme="#1F6FB0", unit_label="", title="", slots=None):
    """Compact horizontal ranking bar for one segment's sub-categories at a
    single year - largest at top, a single-hue gradient (darkest = largest),
    value labels outside.  ``slots`` pads every panel to the same number of rows
    (the max sub-segment count) with blank rows so the BAR THICKNESS is uniform
    across all segments.  Long names are shortened; exports natively."""
    # ascending so the largest ends up at the TOP on BOTH surfaces - Plotly
    # (default y order) and the native Excel BAR chart both plot the first
    # category at the bottom, so no axis reversal is needed (keeps them aligned)
    s = series.sort_values(ascending=True)
    names = _dedupe_labels([str(i) for i in s.index], cap=22)
    vals = [float(v) for v in s.values]
    n = len(vals)
    # largest (last) = darkest theme colour; smallest (first) = lightest tint
    colors = [_lighten(theme, 0.60 * ((n - 1 - i) / (n - 1) if n > 1 else 0.0))
              for i in range(n)]
    text = [f"{v:,.0f}" for v in vals]
    real_max = max(vals) if vals else 1.0
    pad = max(0, (slots or n) - n)               # blank rows at the bottom
    if pad:
        names = [" " * (k + 1) for k in range(pad)] + names
        vals = [None] * pad + vals               # None -> no bar, no label
        colors = [PANEL_BG] * pad + colors
        text = [""] * pad + text
    fig = go.Figure(go.Bar(
        y=names, x=vals, orientation="h", width=BAR_WIDTH,
        marker=dict(color=colors, cornerradius=2),
        text=text, textposition="outside",
        textfont=dict(size=8, color=TEXT), cliponaxis=False,
    ))
    fig.update_layout(
        **_layout(title), bargap=BAR_GAP,
        xaxis=dict(visible=False, range=[0, real_max * 1.34]),
        yaxis=dict(tickfont=dict(size=8), automargin=True),
    )
    # a ranking: no category-axis line (it ran on below the bars into the
    # blank padding rows of a segment with fewer sub-segments)
    fig._chart_meta = dict(ranking=True)
    return fig

def compare_bars(series, value_fmt="{:.1f}%", title=""):
    """Horizontal 'versus' bars - e.g. the selected geography's share of its
    region vs of the world, or size compared against region/Global."""
    names = [str(i) for i in series.index]
    vals = [float(v) for v in series.values]
    colors = _pal(len(names))
    fig = go.Figure(go.Bar(
        y=names, x=vals, orientation="h", width=BAR_WIDTH,
        marker=dict(color=colors, cornerradius=3),
        text=[value_fmt.format(v) for v in vals],
        textposition="outside", textfont=dict(size=10.5, color=TEXT),
    ))
    fig.update_layout(
        **_layout(title), bargap=0.4,
        xaxis=dict(visible=False, range=[0, (max(vals) if vals else 1) * 1.28]),
        yaxis=dict(autorange="reversed", tickfont=dict(size=10.5)),
    )
    return fig

# ── Gauge (PBIX gauge) ───────────────────────────────────────────────
def gauge(value_bps: float, title: str, suffix="BPS", span=None):
    """Deviation gauge in basis points (1 pp = 100 bps).  The PPT deviation
    card reads the same value + suffix off this trace."""
    span = max(abs(value_bps) * 1.3, 25.0) if span is None else span
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=round(float(value_bps), 0),
        number=dict(suffix=f" {suffix}", font=dict(size=17, color=TEXT),
                    valueformat="+,.0f"),
        gauge=dict(
            axis=dict(range=[-span, span], nticks=5,
                      tickfont=dict(size=8)),
            bar=dict(color=OUTPACE_COLOR if value_bps >= 0 else DECLINE_COLOR),
            bordercolor=GRID,
            steps=[dict(range=[-span, 0], color="#FBEAF2"),
                   dict(range=[0, span], color="#EDF2E7")],
            threshold=dict(line=dict(color=TEXT, width=2), value=0),
        ),
    ))
    fig.update_layout(**_layout(title))
    return fig

# ── Attractiveness bubble scatter (PBIX scatterChart) ────────────────
def attractiveness_scatter(df: pd.DataFrame, unit_label="", title="",
                           selected=None):
    fig = go.Figure()
    sizes = df["end"].abs()
    sizeref = 2.0 * sizes.max() / (50.0 ** 2) if sizes.max() else 1
    colors = _pal(len(df))
    for i, name in enumerate(df.index):
        r = df.loc[name]
        dim = selected is not None and str(name) != str(selected)
        fig.add_trace(go.Scatter(
            x=[r["cagr_index"]], y=[r["ms_index"]], name=str(name),
            mode="markers+text", text=[str(name)], customdata=[str(name)],
            textposition="top center",
            textfont=dict(size=9, color=MUTED),
            marker=dict(size=[max(float(r["end"]), 1e-9)], sizemode="area",
                        sizeref=sizeref, sizemin=9,
                        color=colors[i], opacity=0.18 if dim else 0.85,
                        line=dict(color=HEADER_BG if str(name) == str(selected)
                                  else PANEL_BG,
                                  width=2.5 if str(name) == str(selected)
                                  else 1.5)),
            hovertemplate=(f"<b>{name}</b><br>CAGR Index: %{{x:.0f}}"
                           f"<br>MS Index: %{{y:.0f}}"
                           f"<br>2036 size: {r['end']:,.0f} {unit_label}"
                           "<extra></extra>"),
        ))
    fig.add_vline(x=100, line_dash="dash", line_color=FAINT, line_width=1)
    fig.add_hline(y=100, line_dash="dash", line_color=FAINT, line_width=1)
    fig.update_layout(
        **_layout(title), showlegend=False,
        # axis titles spell out the 100 baseline (client reference,
        # 2026-09-29); the native PPT bubble chart carries the same titles
        xaxis=dict(title=dict(text="CAGR Index (100 = segment average "
                                   "growth)", font=dict(size=10)),
                   showgrid=True, gridcolor=GRID, tickfont=dict(size=9)),
        yaxis=dict(title=dict(text="Market Share Index (100 = equal share)",
                              font=dict(size=10)),
                   showgrid=True, gridcolor=GRID, tickfont=dict(size=9)),
    )
    return fig

# ── Choropleth map (PBIX map) ────────────────────────────────────────
def index_map(df: pd.DataFrame, value_col: str, title=""):
    iso, vals, names = [], [], []
    for _, row in df.iterrows():
        for code in GEO_ISO3.get(row["Country"], []):
            iso.append(code)
            vals.append(row[value_col])
            names.append(row["Country"])
    fig = go.Figure(go.Choropleth(
        locations=iso, z=vals, text=names,
        colorscale=[[0, SEQ_LOW], [1, SEQ_HIGH]],
        marker_line_color=PANEL_BG, marker_line_width=0.4,
        colorbar=dict(thickness=7, tickfont=dict(size=8), title=""),
        hovertemplate="<b>%{text}</b><br>" + value_col +
                      ": %{z:,.0f}<extra></extra>",
    ))
    fig.update_layout(
        **_layout(title),
        geo=dict(showframe=False, showcoastlines=False,
                 projection_type="natural earth", bgcolor=PANEL_BG,
                 landcolor="#F3F2F1"),
    )
    return fig

# ── Bubble map (PBIX Bing map with sized bubbles) ────────────────────
def bubble_map_names(df: pd.DataFrame) -> list:
    """Mappable country names in trace order (matches bubble_map) - used
    to build partial marker updates for fast cross-highlighting."""
    return [str(r["Country"]) for _, r in df.iterrows()
            if GEO_ISO3.get(r["Country"])]

def bubble_emphasis(names, selected):
    """(opacity, line-width) arrays highlighting the selected bubble."""
    opac = [0.92 if (selected is None or n == selected) else 0.18
            for n in names]
    width = [2.4 if n == selected else 1.1 for n in names]
    return opac, width

def bubble_map(df: pd.DataFrame, value_col: str, title="", selected=None):
    """Vector world map with one sized/colored bubble per country - the
    Python twin of the PBIX Bing bubble maps.  Uses Scattergeo (built-in
    vector geography, no tiles/token) and auto-zooms to the plotted
    countries.  ``selected`` (a country name) highlights one bubble and
    dims the rest so it reads as cross-filtered."""
    iso, vals, names = [], [], []
    for _, row in df.iterrows():
        codes = GEO_ISO3.get(row["Country"], [])
        if codes:
            iso.append(codes[0])
            vals.append(float(row[value_col]))
            names.append(str(row["Country"]))
    if not iso:                                   # nothing mappable
        fig = go.Figure()
        fig.update_layout(**_layout(title))
        return fig
    vmax = max((abs(v) for v in vals), default=1) or 1
    sizes = [13 + 33 * (abs(v) / vmax) for v in vals]
    opac, lines = bubble_emphasis(names, selected)
    line_col = ["#FFFFFF" for _ in names]
    fig = go.Figure(go.Scattergeo(
        locationmode="ISO-3", locations=iso, text=names, customdata=names,
        marker=dict(size=sizes, sizemode="diameter",
                    color=vals,
                    colorscale=[[0.0, SEQ_STOPS[0]], [0.5, SEQ_STOPS[1]],
                                [1.0, SEQ_STOPS[2]]],
                    opacity=opac, cmin=min(vals), cmax=max(vals),
                    line=dict(color=line_col, width=lines),
                    colorbar=dict(thickness=8, tickfont=dict(size=8),
                                  outlinewidth=0, title="", len=0.8,
                                  x=0.99)),
        hovertemplate="<b>%{text}</b><br>" + value_col +
                      ": %{marker.color:,.1f}<extra></extra>",
    ))
    fig.update_geos(
        resolution=110, showframe=False,
        showcoastlines=True, coastlinecolor="#C3D4E6", coastlinewidth=0.5,
        showcountries=True, countrycolor="#FFFFFF", countrywidth=0.9,
        landcolor="#E4EDF6", showland=True,
        showocean=True, oceancolor="#F3F8FD", bgcolor=PANEL_BG,
        showlakes=True, lakecolor="#F3F8FD",
        projection_type="natural earth", fitbounds="locations",
    )
    fig.update_layout(uirevision="summary-map")   # keep zoom on updates
    lay = _layout(title)
    lay["margin"] = dict(l=4, r=4, t=26 if title else 6, b=4)
    fig.update_layout(**lay)
    return fig

# ── Treemap (PBIX treemap) ───────────────────────────────────────────
def country_treemap_labels(df: pd.DataFrame) -> list:
    return df["Country"].tolist() + df["Region"].drop_duplicates().tolist()

def country_treemap_colors(df: pd.DataFrame, selected=None) -> list:
    labels = country_treemap_labels(df)
    base = _pal(len(labels))
    row = df[df["Country"] == selected] if selected else None
    # nothing selected, or the selection is outside this (top-N) slice -
    # leave every tile at full colour rather than dimming the whole treemap
    if selected is None or row is None or not len(row):
        return base
    sel_region = row.iloc[0]["Region"]             # dim all but the selection
    return [base[i] if (lab == selected or lab == sel_region)
            else "#E3E9F0" for i, lab in enumerate(labels)]

def country_treemap(df: pd.DataFrame, value_col: str, title="",
                    selected=None):
    regions = df["Region"].drop_duplicates().tolist()
    labels = df["Country"].tolist() + regions
    colors = country_treemap_colors(df, selected)
    fig = go.Figure(go.Treemap(
        labels=labels,
        parents=df["Region"].tolist() + [""] * len(regions),
        values=df[value_col].tolist() + [0.0] * len(regions),
        branchvalues="remainder",
        marker=dict(colors=colors, line=dict(color=PANEL_BG, width=1.5)),
        textinfo="label+value",
        texttemplate="<b>%{label}</b><br>%{value:,.0f}",
        textfont=dict(size=10),
    ))
    fig.update_layout(**_layout(title))
    return fig


# ── Template switch support ──────────────────────────────────────────
def _retheme() -> None:
    """Re-derive every module-level constant that bakes a theme token.

    `theming.apply()` rebinds the tokens imported at the top of this module,
    then calls this hook.  Anything computed FROM a token at import time has to
    be recomputed here or the module keeps showing the previous template's
    colours.  Keep this list in step with the constants above.
    """
    global FONT, _COLORWAY, _SEQ_SCALE, _DIV_SCALE
    global _DIV_POS, _DIV_NEG, _RADAR_LINE, _RADAR_FILL
    global _F3_CONE, _F3_BODY, _F3_TOP, _F3_RIM, _F3_LBL
    global _TA_TARGET, _TA_ACTUAL, HEAT_SCALE

    FONT = dict(family=FONT_STACK, size=11, color=TEXT)
    _COLORWAY = list(DATA_COLORS)
    _SEQ_SCALE = [[0.0, SEQ_LOW], [1.0, SEQ_HIGH]]
    _DIV_SCALE = [[0.0, BAD], [0.5, "#F4F7FA"], [1.0, GOOD]]
    _DIV_POS, _DIV_NEG = DIV_POS, DIV_NEG
    _RADAR_LINE, _RADAR_FILL = RADAR_LINE, RADAR_FILL
    _F3_CONE, _F3_BODY = F3_CONE, F3_BODY
    _F3_TOP, _F3_RIM, _F3_LBL = F3_TOP, F3_RIM, F3_LBL
    _TA_TARGET, _TA_ACTUAL = TA_TARGET, TA_ACTUAL
    HEAT_SCALE = [list(stop) for stop in HEAT_STOPS]
