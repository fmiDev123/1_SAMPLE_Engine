"""Cross-sectional Assessment pages (2026-09-29).

Three pages for every cross-tab segment the workbook carries ('Material Type
x End Use Application'), following the client's reference deck
'Cross_Sectional_Assessment_with_2D_Frameworks' for WHAT each page shows, and
built from this deck's own blocks for HOW it looks - the same tables, panels,
charts, summary cards and locking as every other page:

* Market Snapshot - the A x B value matrix, EST_YEAR and END_YEAR for each B
  sub-segment, with key findings
* Market Share Analysis - each A sub-segment's split across B (rows sum to
  100%), a 100% stacked mix chart (the deck's own `figures.region_stack`) and
  share insights
* Absolute $ Opportunity & Attractiveness - each cell's incremental value and
  its CAGR, the top five cells with a 1-5 attractiveness index, and three
  read-out cards

Every figure comes from `data.cross_matrix`, i.e. the workbook's own cells.
"""
from __future__ import annotations

import re

import pandas as pd

from . import classic_text, config, data, figures, layout as L
from .components import fmt_num, fmt_pct
from .config import END_YEAR, EST_YEAR

# the section writes its years as the reference does: 2026E, 2036F
Y0, Y1 = f"{EST_YEAR}E", f"{END_YEAR}F"

# names the registry gives these pages - the export keeps only these for a
# cross-tab segment's section
PAGE_NAMES = ("Market Snapshot", "Market Share Analysis",
              "Absolute $ Opportunity & Attractiveness")


# ── small helpers ──────────────────────────────────────────────────────────
def _short(name):
    """A column header's name without its bracketed detail - 'Beverages
    (Water, Juices, ...)' heads its column as 'Beverages'.  The full name
    stays in every sentence on the page."""
    s = re.sub(r"\s*\([^)]*\)", "", str(name)).strip()
    return s or str(name)


def _headers(names):
    """Short column headers, falling back to the full name wherever two
    would otherwise read the same."""
    short = [_short(n) for n in names]
    return [s if short.count(s) == 1 else str(n) for s, n in zip(short, names)]


def _cagr(a, b, years=END_YEAR - EST_YEAR):
    try:
        return (b / a) ** (1.0 / years) - 1.0 if a > 0 and b > 0 \
            else float("nan")
    except Exception:
        return float("nan")


def _matrices(segment, region):
    """(m0, m1): the EST_YEAR and END_YEAR value matrices, A sub-segments
    down and the B axis's MAIN sub-segments across - the 2D framework the
    snapshot and the opportunity pages read.  A B axis with no level above
    its sub-segments is its own main level, so those matrices are unchanged."""
    return (data.cross_matrix_main(region, "Value", segment, EST_YEAR),
            data.cross_matrix_main(region, "Value", segment, END_YEAR))


# A share table gets a column per B sub-segment, headed by its main group,
# when there are no more than this many - beyond it the columns are too
# narrow to read and the table stays at the main level.
_MAX_SUB_COLS = 12


def _share_detail(segment, region):
    """(m0, m1, groups) for the share table: the B axis at SUB-segment level
    with ``groups`` = [(main, [sub, ...])] heading it - or None for groups
    when the B axis has no second level (or too many columns), in which case
    the matrices are the main-level ones."""
    groups = data.cross_b_groups(segment)
    grouped = any(len(subs) > 1 or main != subs[0] for main, subs in groups)
    n_sub = sum(len(subs) for _m, subs in groups)
    if grouped and n_sub <= _MAX_SUB_COLS:
        subs = [s for _m, ss in groups for s in ss]
        m0 = data.cross_matrix(region, "Value", segment, EST_YEAR)
        m1 = data.cross_matrix(region, "Value", segment, END_YEAR)
        if m0.empty:                     # this geography carries no cells
            return m0, m1, None
        subs = [x for x in subs if x in m0.columns]
        return m0[subs], (m1[subs] if not m1.empty else m1), groups
    m0, m1 = _matrices(segment, region)
    return m0, m1, None


def _empty(title, region, segment):
    return title, {"rows": [L.frow(L.note(
        f"The workbook carries no {segment} cells for {region}, so this "
        f"cross-section cannot be drawn for this geography.", h=120))]}


def _cagr_frame(m0, m1):
    return pd.DataFrame({g: [_cagr(m0.loc[r, g], m1.loc[r, g])
                             for r in m0.index] for g in m0.columns},
                        index=m0.index)


# ═════════════════════ 1. Market Snapshot ═══════════════════════════════
def cross_snapshot(segment: str, region: str):
    a_name, b_name = data.cross_axes(segment)
    u = config.UNIT_LABELS["Value"]
    title = (f"{region} Market Snapshot: {a_name} × {b_name}, "
             f"{Y0} vs {Y1}")
    m0, m1 = _matrices(segment, region)
    if m0.empty or m1.empty:
        return _empty(title, region, segment)
    groups = list(m0.columns)
    heads = _headers(groups)
    # two header rows: each sub-segment over its Y0 | Y1 columns
    tbl = pd.DataFrame(index=m0.index)
    for g, h in zip(groups, heads):
        tbl[f"{h}|{Y0}"] = m0[g].map(fmt_num)
        tbl[f"{h}|{Y1}"] = m1[g].map(fmt_num)
    tbl[f"Total|{Y0}"] = m0.sum(axis=1).map(fmt_num)
    tbl[f"Total|{Y1}"] = m1.sum(axis=1).map(fmt_num)
    tbl.attrs["col_groups"] = [(h, 2) for h in heads] + [("Total", 2)]
    total = []
    for g in groups:
        total += [fmt_num(m0[g].sum()), fmt_num(m1[g].sum())]
    total += [fmt_num(m0.values.sum()), fmt_num(m1.values.sum())]

    a_tot = m0.sum(axis=1)
    top_a = a_tot.idxmax()
    top_b_in_a = m0.loc[top_a].idxmax()
    big = m0.stack().idxmax()
    cg = _cagr_frame(m0, m1)
    fast = cg.stack().idxmax()
    inc = (m1 - m0).stack().sort_values(ascending=False)
    finds = [
        f"'{top_a}' accounts for the largest share of the {region} market by "
        f"{a_name.lower()} ({fmt_num(a_tot[top_a])} {u}, "
        f"{fmt_pct(a_tot[top_a] / a_tot.sum())} in {Y0}), with "
        f"'{top_b_in_a}' its leading {b_name.lower()} sub-segment at "
        f"{fmt_num(m0.loc[top_a, top_b_in_a])} {u}.",
        f"'{big[0]} × {big[1]}' is the single largest cross-sectional cell "
        f"({fmt_num(m0.loc[big])} {u}, {Y0}), while '{fast[0]} × "
        f"{fast[1]}' records the fastest growth at a CAGR of "
        f"{fmt_pct(cg.loc[fast])} over {Y0}–{Y1}.",
        f"Combined, the top three {a_name.lower()} × {b_name.lower()} cells "
        f"generate an absolute $ opportunity of {fmt_num(inc.head(3).sum())} "
        f"{u} by {Y1}.",
    ]
    lead = classic_text.cross_headline(region, a_name, b_name, m1)
    if lead:
        finds[0] = lead
    finds = classic_text.paras(finds)
    # the flow of the END_YEAR value from every A sub-segment into the B
    # groups, beside the key findings (the table above stays as it is)
    flow = figures.sankey(m1, unit_label=u)
    return title, {"rows": [
        L.frow(L.tbl(tbl, index_name=a_name, total_row=total,
                     title=f"{region} Market Value, {a_name} × {b_name}, "
                           f"{Y0} and {Y1} ({u})")),
        L.frow(L.fig(flow, h=330, frac=0.62,
                     title=f"{region} Value Flow, {a_name} → {b_name}, "
                           f"{Y1} ({u})"),
               L.summary("Key Findings", finds, frac=0.38, h=330)),
    ]}


_DETAIL_COLS = 10        # most sub-segment columns on one detail table


def _share_detail_tables(segment, region, a_name, b_name):
    """When the B axis has too many sub-segments for one share table, the
    sub-segment detail follows as extra tables - the B main groups taken in
    runs of at most `_DETAIL_COLS` columns - each row still a share of its
    WHOLE row, so every figure reads the same as on the main table.  The
    layout flows them onto further slides; nothing is squeezed or
    overlapped.  [] when the single table already shows the sub-segments."""
    groups = data.cross_b_groups(segment)
    grouped = any(len(ss) > 1 or m != ss[0] for m, ss in groups)
    n_sub = sum(len(ss) for _m, ss in groups)
    if not grouped or n_sub <= _MAX_SUB_COLS:
        return []
    full = data.cross_matrix(region, "Value", segment, EST_YEAR)
    if full.empty:
        return []
    row_sum = full.sum(axis=1).replace(0, float("nan"))
    col_share = full.sum(axis=0) / (full.values.sum() or 1)
    chunks, cur, cnt = [], [], 0
    for main, subs in groups:
        if cur and cnt + len(subs) > _DETAIL_COLS:
            chunks.append(cur)
            cur, cnt = [], 0
        cur.append((main, subs))
        cnt += len(subs)
    if cur:
        chunks.append(cur)
    out = []
    for k, chunk in enumerate(chunks, 1):
        subs_all = [x for _m, ss in chunk for x in ss if x in full.columns]
        heads = dict(zip(subs_all, _headers(subs_all)))
        tbl = pd.DataFrame(index=full.index)
        for main, subs in chunk:
            for sub in subs:
                if sub in full.columns:
                    tbl[f"{_short(main)}|{heads[sub]}"] = (
                        full[sub].div(row_sum).fillna(0).map(fmt_pct))
        tbl.attrs["col_groups"] = [(_short(m), len([x for x in ss
                                                    if x in full.columns]))
                                   for m, ss in chunk]
        total = [fmt_pct(col_share[x]) for x in subs_all]
        out.append(L.frow(L.tbl(
            tbl, index_name=a_name, total_row=total,
            title=f"{region} Value Share by {b_name} Sub-segment, {a_name} × "
                  f"{b_name}, {Y0} (% of each row) – part {k} of "
                  f"{len(chunks)}")))
    return out


# ═════════════════════ 2. Market Share Analysis ═════════════════════════
def cross_share(segment: str, region: str):
    a_name, b_name = data.cross_axes(segment)
    title = f"{region} Market Share Analysis: {a_name} × {b_name}, {Y0}"
    m0, m1, bgroups = _share_detail(segment, region)
    if m0.empty:
        return _empty(title, region, segment)
    groups = list(m0.columns)
    heads = _headers(groups)
    share = m0.div(m0.sum(axis=1).replace(0, float("nan")), axis=0).fillna(0)
    b_share0 = m0.sum(axis=0) / (m0.values.sum() or 1)
    b_share1 = (m1.sum(axis=0) / (m1.values.sum() or 1)) if not m1.empty         else b_share0
    tbl = pd.DataFrame(index=m0.index)
    if bgroups:
        # two header rows: each B main group over its sub-segments' shares
        col = {g: h for g, h in zip(groups, heads)}
        for main, subs in bgroups:
            for sub in subs:
                tbl[f"{_short(main)}|{col[sub]}"] = share[sub].map(fmt_pct)
        tbl["Total|"] = "100.0%"
        tbl.attrs["col_groups"] = ([(_short(m), len(ss)) for m, ss in bgroups]
                                   + [("Total", 1)])
    else:
        for g, h in zip(groups, heads):
            tbl[h] = share[g].map(fmt_pct)
        tbl["Total"] = "100.0%"
    total = [fmt_pct(b_share0[g]) for g in groups] + ["100.0%"]

    # the mix chart: one 100%-stacked column per A sub-segment, split by the
    # B axis's main groups - the deck's own stacked-share figure (native
    # stacked chart in the PPT)
    c0, _c1 = _matrices(segment, region)
    kids = c0.T.copy()
    kids.index = _headers(list(c0.columns))
    kids.columns = [_short(c) for c in kids.columns]
    fig = figures.region_stack(kids, list(kids.columns))

    top = share.stack().idxmax()
    conc = share.max(axis=1).idxmax()
    conc_b = share.loc[conc].idxmax()
    bps = ((b_share1 - b_share0) * 10000).sort_values()
    gain = bps.index[-1]
    ins = [
        f"'{top[1]}' holds the highest share of '{top[0]}' at "
        f"{fmt_pct(share.loc[top])} ({Y0}).",
        f"'{conc}' is the most concentrated {a_name.lower()} sub-segment, "
        f"with {fmt_pct(share.loc[conc, conc_b])} of its value via "
        f"'{conc_b}'.",
        f"'{gain}' gains {bps[gain]:+,.0f} BPS of share of the {region} "
        f"cross-section by {Y1}.",
    ]
    return title, {"rows": [
        L.frow(L.tbl(tbl, index_name=a_name, total_row=total,
                     title=f"{region} Value Share, {a_name} × {b_name}, "
                           f"{Y0} (% of each row)")),
        L.frow(L.fig(fig, h=300, frac=0.62,
                     title=f"{b_name} Mix by {a_name}, {Y0} "
                           f"(% share)"),
               L.summary("Share Insights", ins, frac=0.38, h=300)),
        *_share_detail_tables(segment, region, a_name, b_name),
    ]}


# ═════════════ 3. Absolute $ Opportunity & Attractiveness ═══════════════
def cross_opportunity(segment: str, region: str):
    a_name, b_name = data.cross_axes(segment)
    u = config.UNIT_LABELS["Value"]
    title = (f"{region} Absolute $ Opportunity & Attractiveness: {a_name} × "
             f"{b_name}, {Y0}–{Y1}")
    m0, m1 = _matrices(segment, region)
    if m0.empty or m1.empty:
        return _empty(title, region, segment)
    groups = list(m0.columns)
    heads = _headers(groups)
    inc = m1 - m0
    cg = _cagr_frame(m0, m1)

    # one full-width table, as the snapshot: each B sub-segment's
    # incremental value and CAGR side by side, every row one line tall
    opp_tbl = pd.DataFrame(index=m0.index)
    for g, h in zip(groups, heads):
        opp_tbl[f"{h}|{u}"] = inc[g].map(fmt_num)
        opp_tbl[f"{h}|CAGR"] = cg[g].map(fmt_pct)
    opp_tbl[f"Total|{u}"] = inc.sum(axis=1).map(fmt_num)
    opp_tbl["Total|CAGR"] = [fmt_pct(_cagr(m0.loc[r].sum(), m1.loc[r].sum()))
                             for r in m0.index]
    opp_tbl.attrs["col_groups"] = [(h, 2) for h in heads] + [("Total", 2)]
    opp_total = []
    for g in groups:
        opp_total += [fmt_num(inc[g].sum()),
                      fmt_pct(_cagr(m0[g].sum(), m1[g].sum()))]
    opp_total += [fmt_num(inc.values.sum()),
                  fmt_pct(_cagr(m0.values.sum(), m1.values.sum()))]

    # attractiveness 1-5: size of the prize (60%) and pace (40%), each as a
    # percentile rank across every cell of the matrix
    flat_inc, flat_cg = inc.stack(), cg.stack()
    score = (0.6 * flat_inc.rank(pct=True)
             + 0.4 * flat_cg.rank(pct=True).fillna(0))
    index = (1 + (score - score.min()) / ((score.max() - score.min()) or 1)
             * 4).round().astype(int)
    top5 = flat_inc.sort_values(ascending=False).head(5)
    # one line each: the units sit in the panel title
    top_lines = [f"{j}. {_short(a)} × {_short(b)}  –  {fmt_num(v)}  ·  "
                 f"{fmt_pct(cg.loc[(a, b)])}  ·  {index[(a, b)]}/5"
                 for j, ((a, b), v) in enumerate(top5.items(), start=1)]

    big = flat_inc.idxmax()
    fast = flat_cg.idxmax()
    mkt_cagr = _cagr(m0.values.sum(), m1.values.sum())
    hi, lo = index.idxmax(), index.idxmin()
    def cell(ab):
        return f"'{_short(ab[0])} × {_short(ab[1])}'"
    cards = [
        ("Largest Pool",
         f"{cell(big)} adds {fmt_num(flat_inc[big])} {u} by {Y1} – "
         f"{fmt_pct(flat_inc[big] / (flat_inc.sum() or 1))} of the total "
         f"opportunity."),
        ("Fastest Growth",
         f"{cell(fast)} grows at a {fmt_pct(flat_cg[fast])} CAGR, "
         f"{(flat_cg[fast] - mkt_cagr) * 10000:,.0f} BPS above the "
         f"cross-section."),
        ("Where to Play",
         f"{cell(hi)} ranks highest ({index[hi]}/5); {cell(lo)} lowest "
         f"({index[lo]}/5), a deprioritised cell."),
    ]
    return title, {"rows": [
        L.frow(L.tbl(opp_tbl, index_name=a_name, total_row=opp_total,
                     title=f"Incremental $ Opportunity ({u}) and Value CAGR, "
                           f"{a_name} × {b_name}, {Y0}–{Y1}")),
        L.frow(L.summary(f"Top 5 Opportunity Cells ({u} · CAGR · Index)",
                         top_lines, frac=0.49, h=190),
               *[L.summary(h_, [t_], frac=0.17, h=190) for h_, t_ in cards]),
    ]}
