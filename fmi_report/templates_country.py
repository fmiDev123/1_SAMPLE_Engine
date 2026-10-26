"""Page templates for 'Matcha Market- Country.pbix' - visuals positioned
at the coordinates extracted from the PBIX report definition."""
from __future__ import annotations

import pandas as pd
from dash import html

from . import classic_text, data, figures, geo, layout as L, metrics, writeups
from .components import (df_table, fmt_num, fmt_pct, graph_panel, kpi_card,
                         panel, pos_neg, table_panel, writeup, yoy_cell,
                         yoy_icon)
from .config import (BAD, COUNTRY_PALETTE, END_YEAR, EST_YEAR, GOOD,
                     HIST_END, HIST_START,
                     UNIT_LABELS, YEARS)
from .data import MARKET_LABEL as MKT
from .templates_main import _by

def _country_note(country, unit):
    cm = classic_text.country_momentum(country) if unit == "Value" else None
    if not cm:
        return writeups.country_momentum(country, unit)
    why = classic_text.reason("cty_growth", country, "market", "Value and Volume",
                              classic_text.geo_facts(country))
    return cm + (" " + why if why else "")

# ═════════════════════ Volume / Value Overview ═══════════════════════
def overview(unit: str, country: str):
    u = UNIT_LABELS[unit]
    root = geo.current()                     # the geography this deck is for
    region = data.region_of(country) or root
    # On a region-scoped deck the root IS this country's region, so the third
    # comparator would just repeat the second ("14.5% of Europe, 14.5% of
    # Global") - it is dropped rather than printed twice.
    same = root == region
    s = data.total_series(country, unit)
    reg = data.total_series(region, unit)
    glob = data.total_series(root, unit)
    # shares divide by the SUM OF THE SIBLINGS (see metrics.share_base), so the
    # chart and the note beside it cannot disagree with the regional pages
    reg_base = metrics.share_base(region, unit)
    glob_base = metrics.share_base(root, unit)
    hc, fc = metrics.hist_cagr(s), metrics.fore_cagr(s)
    reg_fc, glob_fc = metrics.fore_cagr(reg), metrics.fore_cagr(glob)
    prev_s = s.shift(1)                           # previous-year value/volume

    return f"{country} {MKT} {unit} Overview ({u})", {"rows": [
        L.frow(
            L.kpi(f"{HIST_END} Market Size ({u})", fmt_num(s[HIST_END]),
                  "estimate year", frac=0.25),
            L.kpi(f"{END_YEAR} Market Size ({u})", fmt_num(s[END_YEAR]),
                  "forecast horizon", frac=0.25),
            L.kpi(f"H-CAGR ({HIST_START}–{HIST_END})", fmt_pct(hc),
                  frac=0.25),
            L.kpi(f"F-CAGR ({EST_YEAR}–{END_YEAR})", fmt_pct(fc), frac=0.25)),
        L.frow(
            # one column per year - the previous-year companion bar and the
            # figure on every column are gone (client, 2026-10-06)
            L.fig(figures.year_columns([str(y) for y in YEARS], s[YEARS],
                                       f"{unit} ({u})"),
                  h=170, frac=0.4,
                  title=f"{country} {unit} ({u}), {YEARS[0]}–{YEARS[-1]}"),
            # CAGR comparison as horizontal bars, merged with the deviation
            # from the regional benchmark and its analysis
            L.fig(figures.cagr_dev(
                      pd.Series(
                          {country: fc, region: reg_fc} if same
                          else {country: fc, region: reg_fc, root: glob_fc}),
                      metrics.growth_deviation_bps(country, unit, region),
                      region, writeups.country_kpi(country, unit)),
                  h=170, frac=0.6,
                  title="CAGR Comparison & Growth Deviation from Regional "
                        "Benchmark")),
        L.frow(
            L.fig(figures.waterfall(metrics.forecast_waterfall(s),
                                    tick_size=7.5),
                  h=170, frac=0.4,
                  title=f"Forecast Build-up {EST_YEAR}–{END_YEAR} ({u})"),
            # 'versus' comparison - the country's footprint within its region and
            # the world (replaces the two share pies)
            # the country against the rest of its region, two named colours
            L.fig(figures.share_split(
                      [(f"Share of {region}",
                        s[EST_YEAR] / reg_base[EST_YEAR])] if same else
                      [(f"Share of {region}",
                        s[EST_YEAR] / reg_base[EST_YEAR]),
                       (f"Share of {root}",
                        s[EST_YEAR] / glob_base[EST_YEAR])],
                      country, f"Rest of {region}" if same
                      else "Rest of the market"),
                  h=170, frac=0.6,
                  title=f"{country} Market Share, {EST_YEAR}")),
        L.frow(
            L.note(writeups.country_share(country, unit), h=85, frac=0.5),
            L.note(_country_note(country, unit), h=85,
                   frac=0.5)),
    ]}

# ═══════════════ Key-Segments overview (Value only) ══════════════════
# blue / navy / purple / teal / amber / red / deep teal - the report's own
# corporate set, no green, no two neighbours in the same hue family
_SEG_THEME = list(COUNTRY_PALETTE)

def seg_overview_value(country: str):
    """A single-glance overview of the country market at the estimate year, one
    ranking bar per segment (Value only) plus data-driven Key Takeaways.  Fully
    dynamic - adapts to whatever segments / sub-segments the workbook holds."""
    u = UNIT_LABELS["Value"]
    total = float(data.total_series(country, "Value")[EST_YEAR])
    segs = [sg for sg in data.SEGMENTS if not data.is_cross(sg)]   # no cross-tabs here
    # uniform bar thickness: pad every panel to the busiest segment's row count
    slots = max((len(data.segment_table(country, "Value", sg).index)
                 for sg in segs), default=1)
    # Panels wrap into rows of at most PER_ROW.  A single row of 12 segments
    # squeezed each panel to ~100px and pushed the last ones off the page.
    PER_ROW = 4
    n_rows = max(1, -(-len(segs) // PER_ROW))          # ceil
    per_row = min(PER_ROW, len(segs))
    frac = 1.0 / per_row
    row_h = 418 if n_rows == 1 else max(170, int(430 / n_rows))
    # ⚠ Padding buys uniform bar thickness with the CATEGORY LABELS, and the
    # labels are worth more.  A panel is padded with blank-named rows; a native
    # PowerPoint category axis then has that many labels to draw, and when they
    # do not fit it silently draws every Nth - so on a market whose busiest
    # segment is a 17-row cross-segment, a four-row panel was padded to 17 and
    # PowerPoint showed two names out of four, some of them the blank padding.
    # Never pad past what this panel's height can label; a segment that really
    # has that many rows still gets them, because `pad` only ever ADDS.
    slots = min(slots, max(3, int(row_h / 18)))
    charts, takeaways = [], []
    for i, seg in enumerate(segs):
        s = data.segment_table(country, "Value", seg)[EST_YEAR].astype(float)
        theme = _SEG_THEME[i % len(_SEG_THEME)]
        short = seg.replace("By ", "")
        charts.append(L.fig(figures.segment_bar(s, theme, u, slots=slots),
                            h=row_h, frac=frac, title=f"By {short}  ({u})"))
        srt = s.sort_values(ascending=False)
        if not len(srt):
            # the workbook sizes this breakdown in the other unit only
            takeaways.append(
                f"By {short}, the workbook carries no value figures for "
                f"{country}; see its volume pages.")
            continue
        lead, lval = str(srt.index[0]), float(srt.iloc[0])
        share = (lval / total * 100.0) if total else 0.0
        if len(srt) > 1:
            sec, sval = str(srt.index[1]), float(srt.iloc[1])
            takeaways.append(
                f"By {short}, '{lead}' leads at {fmt_num(lval)} {u} "
                f"({share:.0f}% share), ahead of '{sec}' ({fmt_num(sval)} {u}).")
        else:
            takeaways.append(
                f"By {short}, '{lead}' accounts for the whole "
                f"{fmt_num(lval)} {u}.")
    takeaways.append(
        "Use the Country slicer in the top band to switch country the "
        "selection stays synced across every page of the report.")
    title = f"{country} {MKT} Key Segments Overview, By Value ({EST_YEAR})"
    chart_rows = [L.frow(*charts[i:i + PER_ROW])
                  for i in range(0, len(charts), PER_ROW)]
    # the takeaways card carries one bullet per segment, so it has to grow
    take_h = min(300, max(152, 26 + 22 * len(takeaways)))
    return title, {"rows": [
        *chart_rows,
        L.frow(L.summary(
            f"Key Takeaways  ·  {EST_YEAR} Total: {fmt_num(total)} {u}",
            takeaways, h=take_h)),
    ]}

# ═════════════════════════ Segmental page ════════════════════════════
def segmental(segment: str, country: str):
    val = data.segment_table(country, "Value", segment)
    sv = metrics.segment_summary(country, "Value", segment)
    so = metrics.segment_summary(country, "Volume", segment)

    def share_pivot(su):
        tbl = pd.DataFrame(index=su.index)
        tbl[f"Share {EST_YEAR}"] = su["share_est"].map(fmt_pct)
        tbl[f"Share {END_YEAR}"] = su["share_end"].map(fmt_pct)
        tbl["Variance (BPS)"] = [pos_neg(f"{v:+,.0f}", v)
                                 for v in su["bps"]]
        tbl["F-CAGR"] = su["fcagr"].map(fmt_pct)
        return tbl

    abs_wf = pd.DataFrame({
        "label": list(sv.index) + ["Total Absolute $ Opportunity"],
        "value": list(sv["abs_opp"]) + [sv["abs_opp"].sum()],
        "measure": ["relative"] * len(sv) + ["total"],
    })
    return f"{country} {MKT} Segmental Analysis {segment}", {"rows": [
        L.frow(
            L.fig(figures.waterfall(metrics.contribution_waterfall(val),
                                    tick_size=8), h=290, frac=0.5,
                  title=f"Growth Contribution {EST_YEAR} → {END_YEAR}, "
                        f"{_by(segment)} ({UNIT_LABELS['Value']})"),
            L.tbl(share_pivot(sv), index_name=segment, frac=0.5,
                  title=f"Value Shares & Variance, {_by(segment)}")),
        L.frow(
            L.fig(figures.waterfall(abs_wf, tick_size=8), h=290, frac=0.5,
                  title=f"Absolute $ Opportunity, {_by(segment)} "
                        f"({UNIT_LABELS['Value']}), between "
                        f"{EST_YEAR}–{END_YEAR}"),
            L.tbl(share_pivot(so), index_name=segment, frac=0.5,
                  title=f"Volume Shares & Variance, {_by(segment)}")),
    ]}

# ══════════════════════ Table (Volume & Value) ═══════════════════════
def tables(segment: str, country: str):
    vol_t, vol_m = data.deep_table(country, "Volume", segment)
    val_t, val_m = data.deep_table(country, "Value", segment)
    vol = vol_t[YEARS].map(fmt_num)
    val = val_t[YEARS].map(fmt_num)
    clbl = "F-CAGR"                                # forecast-CAGR block (right)
    vol[clbl] = [fmt_pct(metrics.fore_cagr(vol_t.loc[i])) for i in vol_t.index]
    val[clbl] = [fmt_pct(metrics.fore_cagr(val_t.loc[i])) for i in val_t.index]
    vol_tm, val_tm = vol_t.loc[vol_m], val_t.loc[val_m]   # totals: parents only
    vol_total = [fmt_num(v) for v in vol_tm[YEARS].sum(axis=0)] + \
        [fmt_pct(metrics.fore_cagr(vol_tm.sum(axis=0)))]
    val_total = [fmt_num(v) for v in val_tm[YEARS].sum(axis=0)] + \
        [fmt_pct(metrics.fore_cagr(val_tm.sum(axis=0)))]
    return f"{country} {MKT} Tables {segment}", {"tables": [
        L.tbl(vol, index_name=segment, total_row=vol_total,
              title=f"{country} Volume, {_by(segment)} "
                    f"({UNIT_LABELS['Volume']})"),
        L.tbl(val, index_name=segment, total_row=val_total,
              title=f"{country} Value, {_by(segment)} "
                    f"({UNIT_LABELS['Value']})"),
    ]}

# ═════════════════════════════ YoY page ══════════════════════════════
def yoy(segment: str, country: str):
    def icon_pivot(unit):
        tbl, mains = data.deep_table(country, unit, segment)
        y = metrics.yoy_table(tbl)[YEARS[1:]]
        market = tbl.loc[mains].sum(axis=0).pct_change().mul(100.0)[YEARS[1:]]
        disp = pd.DataFrame(index=y.index)
        for c in YEARS[1:]:
            disp[c] = [yoy_cell(v, market[c]) for v in y[c]]
        return disp

    legend = "  ·  ▲ outpacing market  ▶ below market pace  ▼ decline"
    return f"{country} YoY Analysis {segment}", {"tables": [
        L.tbl(icon_pivot("Volume"), index_name=segment,
              title=f"{country} Volume YoY, {_by(segment)} (%)" + legend),
        L.tbl(icon_pivot("Value"), index_name=segment,
              title=f"{country} Value YoY, {_by(segment)} (%)" + legend),
    ]}


# -- Template switch support ------------------------------------------
def _retheme() -> None:
    """Re-derive what this module bakes from theme tokens (see theming.py)."""
    global _SEG_THEME
    _SEG_THEME = list(COUNTRY_PALETTE)
