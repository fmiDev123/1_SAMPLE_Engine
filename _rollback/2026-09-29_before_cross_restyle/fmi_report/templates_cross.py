"""Cross-sectional Assessment pages (2026-09-29).

Three pages for every cross-tab segment the workbook carries ('Material Type
x End Use Application'), modelled on the client's reference deck
'Cross_Sectional_Assessment_with_2D_Frameworks' and drawn in this deck's own
template and palette:

* Market Snapshot - the A x B value matrix, EST_YEAR and END_YEAR side by
  side under each B sub-segment, with key findings
* Market Share Analysis - each A sub-segment's split across B (rows sum to
  100%), a 100% stacked mix chart and share insights
* Absolute $ Opportunity & Attractiveness - each cell's incremental value and
  CAGR, the top five cells with a 1-5 attractiveness index, and three cards

Every figure comes from `data.cross_matrix`, i.e. the workbook's own cells.
The pages are primitive-drawn ({"primpages": ...}) so the two-level column
headers of the reference can be reproduced exactly, in the deck and on the
dashboard alike.
"""
from __future__ import annotations

import math
import re

import pandas as pd

from . import config, data, figures
from .components import fmt_num, fmt_pct
from .config import END_YEAR, EST_YEAR

# names the registry gives these pages - the export keeps only these for a
# cross-tab segment's section
PAGE_NAMES = ("Market Snapshot", "Market Share Analysis",
              "Absolute $ Opportunity & Attractiveness")

X0, W = 15, 1250                       # content band, as every other page
TOP = 112                              # just under the header rule
EDGE = "#C9D2D9"                       # table hairlines (as Company Coverage)
WHITE = "#FFFFFF"


# ── small helpers ──────────────────────────────────────────────────────────
def _hex(c):
    c = str(c).lstrip("#")
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def _mix(a, b, t):
    """Colour `t` of the way from `a` to `b`."""
    ra, rb = _hex(a), _hex(b)
    return "#%02X%02X%02X" % tuple(round(x + (y - x) * t) for x, y in zip(ra, rb))


def _theme():
    head = config.HEADER_BG
    acc = config.ACCENT
    return {
        "head": head,                              # row-label header
        "total": _mix(head, "#000000", 0.25),      # the Total column header
        "label": _mix(acc, WHITE, 0.84),           # row labels / total row
        "text": config.TEXT,
        "muted": config.MUTED,
        "accent": acc,
        "card": _mix(acc, WHITE, 0.90),
    }


def _group_colours(n):
    """One header blue per B sub-segment, stepping from the template's header
    colour towards its accent - the reference's run of blues."""
    th = _theme()
    if n <= 1:
        return [th["head"]]
    return [_mix(th["head"], th["accent"], 0.15 + 0.7 * i / (n - 1))
            for i in range(n)]


def _t(x, y, w, h, s, size=11, bold=False, color=None, align="c",
       anchor=None, wrap=False):
    p = {"t": "text", "x": x, "y": y, "w": w, "h": h, "s": str(s),
         "size": size, "bold": bold, "color": color or config.TEXT,
         "align": align}
    if anchor:
        p["anchor"] = anchor
    if wrap:
        p["wrap"] = True
    return p


def _r(x, y, w, h, fill, line=EDGE, lw=0.75, rounded=False):
    return {"t": "rrect" if rounded else "rect", "x": x, "y": y, "w": w,
            "h": h, "fill": fill, "line": line, "lw": lw}


def _bullets(x, y, w, head, lines):
    """A small titled bullet block; each bullet as tall as its wrapped
    text (11pt Calibri, ~6.2px a character)."""
    th = _theme()
    out = [_t(x, y, w, 22, head, 12, True, th["head"], "l")]
    yy = y + 28
    for s in lines:
        n = max(1, math.ceil(len(s) * 6.6 / max(w - 24, 60)))
        h = n * 18 + 4
        out.append(_t(x, yy, 14, 18, "•", 11, True, th["accent"], "l",
                      anchor="t"))
        out.append(_t(x + 16, yy, w - 16, h, s, 11, False, th["text"],
                      "l", anchor="t", wrap=True))
        yy += h + 8
    return out


def _short(name):
    """A column header's name without its bracketed detail - 'Beverages
    (Water, Juices, ...)' heads its column as 'Beverages'.  The full name
    stays in every sentence on the page."""
    s = re.sub(r"\s*\([^)]*\)", "", str(name)).strip()
    return s or str(name)


def _cagr(a, b, years=END_YEAR - EST_YEAR):
    try:
        return (b / a) ** (1.0 / years) - 1.0 if a > 0 and b > 0 else float("nan")
    except Exception:
        return float("nan")


def _matrices(segment, region):
    """(m0, m1): the EST_YEAR and END_YEAR value matrices."""
    return (data.cross_matrix(region, "Value", segment, EST_YEAR),
            data.cross_matrix(region, "Value", segment, END_YEAR))


def _empty(title, region, segment):
    th = _theme()
    return title, {"primpages": [[
        _t(X0, 330, W, 40,
           f"The workbook carries no {segment} cells for {region}, so this "
           f"cross-section cannot be drawn for this geography.",
           13, False, th["muted"], "c")]]}


def _grid_header(x, y, label_w, col_w, groups, sub, a_name, th,
                 gh=58, sh=26, total_label="Total"):
    """The two-level header: a row-label cell spanning both rows, one blue
    group cell per B sub-segment over its `sub` columns, and a Total group."""
    P = []
    n_sub = len(sub)
    full_h = gh + sh
    P.append(_r(x, y, label_w, full_h, th["head"], th["head"]))
    P.append(_t(x + 8, y, label_w - 12, full_h, a_name, 11, True, WHITE, "l"))
    cx = x + label_w
    cols = _group_colours(len(groups))
    for g, colour in zip(groups + [total_label], cols + [th["total"]]):
        gw = col_w * n_sub
        P.append(_r(cx, y, gw, gh, colour, WHITE, 0.75))
        P.append(_t(cx + 3, y, gw - 6, gh, _short(g), 11, True, WHITE, "c",
                    wrap=True))
        for k, s in enumerate(sub):
            P.append(_r(cx + k * col_w, y + gh, col_w, sh, colour, WHITE,
                        0.75))
            P.append(_t(cx + k * col_w, y + gh, col_w, sh, s, 11, True,
                        WHITE, "c"))
        cx += gw
    return P, full_h


def _body_rows(x, y, label_w, col_w, row_h, labels, cells, th,
               bold_last=True, colours=None):
    """Body rows: a light-blue label cell then one cell per value; the last
    row (the Total) is set in bold on the label tint."""
    P = []
    for i, (lab, vals) in enumerate(zip(labels, cells)):
        yy = y + i * row_h
        last = bold_last and i == len(labels) - 1
        P.append(_r(x, yy, label_w, row_h, th["label"]))
        P.append(_t(x + 8, yy, label_w - 12, row_h, lab, 11, True,
                    th["text"], "l", wrap=True))
        for k, v in enumerate(vals):
            cx = x + label_w + k * col_w
            fill = th["label"] if last else WHITE
            P.append(_r(cx, yy, col_w, row_h, fill))
            col = (colours[k] if colours else None) or th["text"]
            P.append(_t(cx, yy, col_w, row_h, v, 11, last, col, "c"))
    return P


# ═════════════════════ 1. Market Snapshot ═══════════════════════════════
def cross_snapshot(segment: str, region: str):
    a_name, b_name = data.cross_axes(segment)
    u = config.UNIT_LABELS["Value"]
    title = (f"{region} Market Snapshot: {a_name} × {b_name}, "
             f"{EST_YEAR} vs {END_YEAR}")
    m0, m1 = _matrices(segment, region)
    if m0.empty or m1.empty:
        return _empty(title, region, segment)
    th = _theme()
    rows = list(m0.index) + ["Total"]
    groups = list(m0.columns)
    years = [str(EST_YEAR), str(END_YEAR)]
    label_w = 210
    col_w = (W - label_w) / (2 * (len(groups) + 1))
    P, hh = _grid_header(X0, TOP, label_w, col_w, groups, years, a_name, th)
    cells = []
    for r in m0.index:
        vals = []
        for g in groups:
            vals += [fmt_num(m0.loc[r, g]), fmt_num(m1.loc[r, g])]
        vals += [fmt_num(m0.loc[r].sum()), fmt_num(m1.loc[r].sum())]
        cells.append(vals)
    tot = []
    for g in groups:
        tot += [fmt_num(m0[g].sum()), fmt_num(m1[g].sum())]
    tot += [fmt_num(m0.values.sum()), fmt_num(m1.values.sum())]
    cells.append(tot)
    avail = 440 - hh
    row_h = max(26, min(42, avail / len(rows)))
    P += _body_rows(X0, TOP + hh, label_w, col_w, row_h, rows, cells, th)
    # key findings
    a_tot = m0.sum(axis=1)
    top_a = a_tot.idxmax()
    top_b_in_a = m0.loc[top_a].idxmax()
    big = m0.stack().idxmax()
    cg = pd.DataFrame({g: [_cagr(m0.loc[r, g], m1.loc[r, g])
                           for r in m0.index] for g in groups},
                      index=m0.index)
    fast = cg.stack().idxmax()
    inc = (m1 - m0).stack().sort_values(ascending=False)
    finds = [
        f"'{top_a}' accounts for the largest share of the {region} market by "
        f"{a_name.lower()} ({fmt_num(a_tot[top_a])} {u}, "
        f"{fmt_pct(a_tot[top_a] / a_tot.sum())} in {EST_YEAR}), with "
        f"'{top_b_in_a}' its leading {b_name.lower()} sub-segment at "
        f"{fmt_num(m0.loc[top_a, top_b_in_a])} {u}.",
        f"'{big[0]} × {big[1]}' is the single largest cross-sectional cell "
        f"({fmt_num(m0.loc[big])} {u}, {EST_YEAR}), while '{fast[0]} × "
        f"{fast[1]}' records the fastest growth at a CAGR of "
        f"{fmt_pct(cg.loc[fast])} over {EST_YEAR}–{END_YEAR}.",
        f"Combined, the top three {a_name.lower()} × {b_name.lower()} cells "
        f"generate an absolute $ opportunity of {fmt_num(inc.head(3).sum())} "
        f"{u} by {END_YEAR}.",
    ]
    fy = TOP + hh + row_h * len(rows) + 22
    P += _bullets(X0 + 10, fy, W - 20, "Key findings", finds)
    return title, {"primpages": [P]}


# ═════════════════════ 2. Market Share Analysis ═════════════════════════
def cross_share(segment: str, region: str):
    a_name, b_name = data.cross_axes(segment)
    title = f"{region} Market Share Analysis: {a_name} × {b_name}, {EST_YEAR}"
    m0, m1 = _matrices(segment, region)
    if m0.empty:
        return _empty(title, region, segment)
    th = _theme()
    groups = list(m0.columns)
    share = m0.div(m0.sum(axis=1).replace(0, float("nan")), axis=0).fillna(0)
    b_share0 = m0.sum(axis=0) / (m0.values.sum() or 1)
    b_share1 = m1.sum(axis=0) / (m1.values.sum() or 1) if not m1.empty \
        else b_share0
    label_w = 270
    col_w = (W - label_w) / (len(groups) + 1)
    # one-level header here (a B sub-segment per column), tall enough for
    # its name to wrap
    hh = 58
    P = [_r(X0, TOP, label_w, hh, th["head"], th["head"]),
         _t(X0 + 8, TOP, label_w - 12, hh, a_name, 11, True, WHITE, "l")]
    cols = _group_colours(len(groups))
    for k, (g, colour) in enumerate(zip(groups + ["Total"],
                                        cols + [th["total"]])):
        cx = X0 + label_w + k * col_w
        P.append(_r(cx, TOP, col_w, hh, colour, WHITE))
        P.append(_t(cx + 3, TOP, col_w - 6, hh, _short(g), 11, True, WHITE,
                    "c", wrap=True))
    rows = list(m0.index) + ["Share of total market"]
    cells = [[fmt_pct(share.loc[r, g]) for g in groups] + ["100.0%"]
             for r in m0.index]
    cells.append([fmt_pct(b_share0[g]) for g in groups] + ["100.0%"])
    row_h = max(26, min(36, (300 - hh) / len(rows)))
    P += _body_rows(X0, TOP + hh, label_w, col_w, row_h, rows, cells, th)
    # 100% stacked mix, one column per A sub-segment, drawn in the palette
    cy = TOP + hh + row_h * len(rows) + 16
    P.append(_t(X0, cy, 800, 22, f"{b_name} mix by {a_name.lower()}, "
                f"{EST_YEAR} (% share)", 12, True, th["head"], "l"))
    pal = figures._pal(len(groups))
    ch_x, ch_w = X0 + 10, 800
    ch_top, ch_bot = cy + 30, 700 - 92
    n = len(m0.index)
    slot = ch_w / max(n, 1)
    bw = min(120, slot * 0.6)
    for i, r in enumerate(m0.index):
        bx = ch_x + i * slot + (slot - bw) / 2
        yy = ch_bot
        for g, colour in zip(groups, pal):
            h = (ch_bot - ch_top) * float(share.loc[r, g])
            if h > 0.3:
                P.append(_r(bx, yy - h, bw, h, colour, WHITE, 0.5))
            yy -= h
        P.append(_t(ch_x + i * slot + 4, ch_bot + 2, slot - 8, 34,
                    _short(r), 11, False, th["text"], "c", anchor="t",
                    wrap=True))
    P.append({"t": "line", "x1": ch_x, "y1": ch_bot, "x2": ch_x + ch_w,
              "y2": ch_bot, "color": EDGE, "width": 1})
    # legend
    per_row = 3                   # legend: rows of three, the full width
    step = (W - 20) / per_row     # (the insights column ends above it)
    for k, (g, colour) in enumerate(zip(groups, pal)):
        lx = ch_x + (k % per_row) * step
        ly = ch_bot + 40 + (k // per_row) * 20
        P.append(_r(lx, ly + 5, 10, 10, colour, colour))
        P.append(_t(lx + 14, ly, step - 16, 20, _short(g), 11, False,
                    th["muted"], "l"))
    # share insights
    top = share.stack().idxmax()
    conc = share.max(axis=1).idxmax()
    conc_b = share.loc[conc].idxmax()
    bps = ((b_share1 - b_share0) * 10000).sort_values()
    gain = bps.index[-1]
    ins = [
        f"'{top[1]}' holds the highest share of '{top[0]}' at "
        f"{fmt_pct(share.loc[top])} ({EST_YEAR}).",
        f"'{conc}' is the most concentrated {a_name.lower()} sub-segment, "
        f"with {fmt_pct(share.loc[conc, conc_b])} of its value via "
        f"'{conc_b}'.",
        f"'{gain}' gains {bps[gain]:+,.0f} bps of share of the {region} "
        f"cross-section by {END_YEAR}.",
    ]
    P += _bullets(X0 + ch_w + 40, cy, W - ch_w - 40, "Share insights", ins)
    return title, {"primpages": [P]}


# ═════════════ 3. Absolute $ Opportunity & Attractiveness ═══════════════
def cross_opportunity(segment: str, region: str):
    a_name, b_name = data.cross_axes(segment)
    u = config.UNIT_LABELS["Value"]
    title = (f"{region} Absolute $ Opportunity & Attractiveness: {a_name} × "
             f"{b_name}, {EST_YEAR}–{END_YEAR}")
    m0, m1 = _matrices(segment, region)
    if m0.empty or m1.empty:
        return _empty(title, region, segment)
    th = _theme()
    groups = list(m0.columns)
    inc = m1 - m0
    cg = pd.DataFrame({g: [_cagr(m0.loc[r, g], m1.loc[r, g])
                           for r in m0.index] for g in groups},
                      index=m0.index)
    # attractiveness 1-5: size of the prize (60%) and pace (40%), each as a
    # percentile rank across every cell of the matrix
    flat_inc, flat_cg = inc.stack(), cg.stack()
    score = (0.6 * flat_inc.rank(pct=True)
             + 0.4 * flat_cg.rank(pct=True).fillna(0))
    index = (1 + (score - score.min()) / ((score.max() - score.min()) or 1)
             * 4).round().astype(int)
    MW = 860                                         # matrix width
    label_w = 170
    col_w = (MW - label_w) / (len(groups) + 1)
    hh = 72
    P = [_r(X0, TOP, label_w, hh, th["head"], th["head"]),
         _t(X0 + 8, TOP, label_w - 12, hh, f"{a_name} ({u})", 11, True,
            WHITE, "l", wrap=True)]
    cols = _group_colours(len(groups))
    for k, (g, colour) in enumerate(zip(groups + ["Total"],
                                        cols + [th["total"]])):
        cx = X0 + label_w + k * col_w
        P.append(_r(cx, TOP, col_w, hh, colour, WHITE))
        # regular weight: a narrow column still takes a long word whole
        P.append(_t(cx + 2, TOP, col_w - 4, hh, _short(g), 11, False, WHITE,
                    "c", wrap=True))
    rows = list(m0.index) + ["Total"]
    row_h = max(40, min(56, (400 - hh) / len(rows)))

    def cell(a, b):
        return fmt_num(b - a), fmt_pct(_cagr(a, b))
    body = []
    for r in m0.index:
        body.append([cell(m0.loc[r, g], m1.loc[r, g]) for g in groups]
                    + [cell(m0.loc[r].sum(), m1.loc[r].sum())])
    body.append([cell(m0[g].sum(), m1[g].sum()) for g in groups]
                + [cell(m0.values.sum(), m1.values.sum())])
    for i, (lab, vals) in enumerate(zip(rows, body)):
        yy = TOP + hh + i * row_h
        last = i == len(rows) - 1
        P.append(_r(X0, yy, label_w, row_h, th["label"]))
        P.append(_t(X0 + 8, yy, label_w - 12, row_h, lab, 11, True,
                    th["text"], "l", wrap=True))
        for k, (v, c) in enumerate(vals):
            cx = X0 + label_w + k * col_w
            fill = th["label"] if last or k == len(vals) - 1 else WHITE
            P.append(_r(cx, yy, col_w, row_h, fill))
            P.append(_t(cx, yy + row_h / 2 - 19, col_w, 20, v, 11, True,
                        th["text"], "c"))
            P.append(_t(cx, yy + row_h / 2 - 1, col_w, 20, f"CAGR {c}", 11,
                        False, th["accent"], "c"))
    # top five cells
    rx = X0 + MW + 30
    rw = W - MW - 30
    P.append(_t(rx, TOP, rw, 26, f"Top 5 opportunity cells, "
                f"{EST_YEAR}–{END_YEAR}", 12, True, th["head"], "l"))
    top5 = flat_inc.sort_values(ascending=False).head(5)
    shades = [_mix(th["head"], th["accent"], t)
              for t in (0.0, 0.3, 0.5, 0.7, 0.85)]
    for j, ((a, b), v) in enumerate(top5.items()):
        yy = TOP + 40 + j * 62
        P.append({"t": "dot", "cx": rx + 18, "cy": yy + 22, "d": 34,
                  "color": shades[j], "edge": shades[j]})
        P.append(_t(rx, yy + 11, 36, 22, str(j + 1), 12, True, WHITE, "c"))
        P.append(_t(rx + 46, yy, rw - 46, 22, f"{a} × {b}", 11, True,
                    th["text"], "l"))
        P.append(_t(rx + 46, yy + 22, rw - 46, 22,
                    f"{fmt_num(v)} {u}  ·  CAGR {fmt_pct(cg.loc[(a, b)])}  ·  "
                    f"Index {index[(a, b)]}/5", 11, False, th["muted"], "l"))
    # three cards
    cy = TOP + hh + row_h * len(rows) + 18
    ch = 700 - cy
    cw = (MW - 2 * 16) / 3
    big = flat_inc.idxmax()
    fast = flat_cg.idxmax()
    mkt_cagr = _cagr(m0.values.sum(), m1.values.sum())
    hi, lo = index.idxmax(), index.idxmin()
    cards = [
        ("Largest pool",
         f"'{big[0]} × {big[1]}' adds {fmt_num(flat_inc[big])} {u} of "
         f"incremental value by {END_YEAR} – "
         f"{fmt_pct(flat_inc[big] / (flat_inc.sum() or 1))} of the total "
         f"{a_name.lower()} × {b_name.lower()} opportunity."),
        ("Fastest growth",
         f"'{fast[0]} × {fast[1]}' expands at a CAGR of "
         f"{fmt_pct(flat_cg[fast])}, "
         f"{(flat_cg[fast] - mkt_cagr) * 10000:,.0f} bps above the "
         f"cross-section as a whole."),
        ("Where to play",
         f"The attractiveness index ranks '{hi[0]} × {hi[1]}' highest "
         f"({index[hi]}/5); '{lo[0]} × {lo[1]}' scores lowest "
         f"({index[lo]}/5) and is a deprioritised cell."),
    ]
    for k, (head, body_txt) in enumerate(cards):
        cx = X0 + k * (cw + 16)
        P.append(_r(cx, cy, cw, ch, th["card"], th["card"], rounded=True))
        P.append(_t(cx + 14, cy + 8, cw - 28, 22, head, 12, True,
                    th["head"], "l"))
        P.append(_t(cx + 14, cy + 34, cw - 28, ch - 40, body_txt, 11, False,
                    th["text"], "l", anchor="t", wrap=True))
    return title, {"primpages": [P]}
