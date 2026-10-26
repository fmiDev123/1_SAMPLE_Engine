"""Competitive Landscape: 'Industry Structure Landscape' (2026-09-29).

Two pages modelled on the client's reference 'Industry_Structure_Landscape
(2)' - its content and its tier colours - set inside this deck's own chrome:
every block sits in the deck's white soft-cornered panel with its title.

* page 1 - tier ladder (share of market value), concentration read-out,
  tier profile matrix and key takeaways
* page 2 - structural shift END_YEAR vs EST_YEAR (tier migration, entry and
  exit), four headline cards, the structural forces table and key takeaways

Tier sizes and players come from `competitors.tiers()` (the roster, or the
workbook's 'Top 10 Companies' sheet), concentration from
`competitors.company_shares()`, and the forces from the workbook-driven
Five Forces scores in `dynamics.five_forces()`.  Competitor figures are
indicative, as on every other competition page.
"""
from __future__ import annotations

from . import competitors as C
from . import config, dynamics as D
from .components import fmt_num, fmt_pct
from .config import END_YEAR, EST_YEAR
from .data import MARKET_LABEL as MKT

# the reference deck's colours
T_COL = {"Tier 1": "#1F3A6E", "Tier 2": "#2E6DB4", "Tier 3": "#6FA8DC",
         "Others": "#9FC5E8"}
TOTAL_HEAD = "#102447"
TINT = "#EAF1FA"                  # label cells, cards
TEAL_TINT = "#E0F3F3"             # score / posture cells
TEAL, ORANGE = "#0E8C8C", "#D08C2A"
INK, GREY = "#14213D", "#6B7280"
BAND = ("#C9DAF0", "#6FA8DC", "#1F3A6E")      # HHI read-out bands
# the standard HHI classes - fixed thresholds, not market figures, so the
# locked deck keeps them readable (see locking.default_policy)
HHI_BANDS = ("Fragmented HHI < 1,500", "Moderate 1,500–2,500",
             "Concentrated > 2,500")
WHITE = "#FFFFFF"
GRID = "#C9D2D9"
TIERS = ("Tier 1", "Tier 2", "Tier 3", "Others")

# indicative growth of each tier against the market (percentage points of
# CAGR): the long tail consolidates, mid-tier specialists scale fastest
_TIER_DELTA = {"Tier 1": 0.003, "Tier 2": 0.008, "Tier 3": 0.002,
               "Others": -0.015}
# tier qualities - generic to any market
_QUAL = {
    "Price positioning": ("Premium", "Mid / value", "Value / entry",
                          "Mixed"),
    "Route-to-market": ("Direct / key accounts", "Distributors",
                        "Regional partners", "Local / online"),
    "Portfolio breadth": ("Full portfolio", "2–3 core lines", "Single line",
                          "Niche"),
    "Strategic posture": ("Defend & bundle", "Niche deepening",
                          "Scale-up or exit", "Long-tail"),
}


# ── primitives ──────────────────────────────────────────────────────────────
def _t(x, y, w, h, s, size=11, bold=False, color=INK, align="l",
       anchor=None, wrap=False):
    p = {"t": "text", "x": x, "y": y, "w": w, "h": h, "s": str(s),
         "size": size, "bold": bold, "color": color, "align": align}
    if anchor:
        p["anchor"] = anchor
    if wrap:
        p["wrap"] = True
    return p


def _r(x, y, w, h, fill, line=None, lw=0.75, r=None):
    p = {"t": "rrect" if r else "rect", "x": x, "y": y, "w": w, "h": h,
         "fill": fill, "line": line, "lw": lw}
    if r:
        p["r"] = r
    return p


def _panel(x, y, w, h, title, sub=None):
    """The deck's panel: white, soft-cornered, hairline edge, bold title."""
    r = config.PANEL_RADIUS if getattr(config, "PANEL_ROUNDED", True) else None
    P = [_r(x, y, w, h, WHITE, config.PANEL_EDGE, 0.75, r=r),
         _t(x + 10, y + 6, w - 20, 22, title, 12, True, config.TEXT)]
    if sub:
        P.append(_t(x + 10, y + 27, w - 20, 26, sub, 11, False, GREY))
    return P


def _arrow(x1, x2, y, colour, left):
    """A horizontal arrow from x1 to x2, head at the left or right end."""
    head = x1 if left else x2
    d = 1 if left else -1
    return [{"t": "line", "x1": x1 + (8 * d if left else 0), "y1": y,
             "x2": x2 - (0 if left else 8), "y2": y, "color": colour,
             "width": 2},
            {"t": "shape", "pts": [(head, y), (head + 10 * d, y - 5),
                                   (head + 10 * d, y + 5)],
             "fill": colour, "line": colour}]


def _bullets(x, y, w, lines, gap=6):
    P, yy = [], y
    for s in lines:
        n = max(1, -(-len(s) * 7.0 // max(w - 24, 60)))
        h = int(n * 17 + 4)
        P.append(_t(x, yy, 12, 17, "•", 11, True, config.ACCENT, anchor="t"))
        P.append(_t(x + 14, yy, w - 14, h + 18, s, 11, False, INK,
                    anchor="t", wrap=True))
        yy += h + gap
    return P


# ── the numbers ─────────────────────────────────────────────────────────────
def _pct(v):
    """A share as a bare figure, '44.5' - the column says it is a %."""
    return fmt_pct(v).rstrip("%")


def _players(n):
    return "–" if not n else f"{n} player" + ("" if n == 1 else "s")


def _cagr(a, b, n=END_YEAR - EST_YEAR):
    return (b / a) ** (1.0 / n) - 1.0 if a > 0 and b > 0 else float("nan")


def _model():
    u = config.UNIT_LABELS["Value"]
    tiers = C.tiers()
    mv0 = C.market_value("Global", EST_YEAR)
    mv1 = C.market_value("Global", END_YEAR) or mv0
    mkt_cagr = _cagr(mv0, mv1)
    # each tier grows at the market pace plus its own (indicative) delta;
    # the END_YEAR values are then scaled back onto the market total
    raw1 = {t: tiers[t]["size"] * (1 + mkt_cagr + _TIER_DELTA[t])
            ** (END_YEAR - EST_YEAR) for t in TIERS}
    k = mv1 / (sum(raw1.values()) or 1)
    v1 = {t: raw1[t] * k for t in TIERS}
    s0 = s0_ = {t: tiers[t]["share"] for t in TIERS}
    s1 = {t: v1[t] / (mv1 or 1) for t in TIERS}
    n0 = {t: len(tiers[t]["players"]) for t in TIERS}
    # tier migration (indicative): one mid-tier player moves up via M&A, one
    # emerging player scales into Tier 2, entrants scale up from the tail,
    # and one Tier 3 player exits or is acquired
    up = {("Tier 2", "Tier 1"): 1, ("Tier 3", "Tier 2"): 1,
          ("Others", "Tier 3"): 2}
    down = {("Tier 1", "Tier 2"): 0, ("Tier 2", "Tier 3"): 0,
            ("Tier 3", "Others"): 1}
    n1 = dict(n0)
    for (a, b), m in up.items():
        n1[a] -= m if a != "Others" else 0
        n1[b] += m
    for (a, b), m in down.items():
        n1[a] -= m
        n1[b] += m if b != "Others" else 0
    n1 = {t: max(0, v) for t, v in n1.items()}
    entrants = int(round(2 + max(mkt_cagr, 0) * 20))
    exits = max(1, int(round(entrants * 0.5)))
    # concentration from the company shares
    shares = [s for name, s in C.company_shares("Global") if name != C.OTHERS]
    shares.sort(reverse=True)
    hhi = sum(s * s for s in shares)
    cr4 = sum(shares[:4])
    cls = ("Fragmented" if hhi < 1500 else
           "Moderately concentrated" if hhi <= 2500 else "Concentrated")
    # structural forces, 1-5, from the workbook-driven Five Forces (0-10)
    ff = {f["name"]: f["score"] for f in D.five_forces("Global")}

    def s5(v):
        return int(max(1, min(5, round(1 + float(v) / 10 * 4))))
    forces = [
        ("Entry barriers",
         "Certification, qualification, scale economics",
         s5(10 - ff.get("Threat of New Entrants", 5)),
         "Protects Tier 1–2; slows new entry"),
        ("Buyer power",
         "Key-account purchasing, multi-year supply contracts",
         s5(ff.get("Bargaining Power of Buyers", 5)),
         "Compresses Tier 3 pricing"),
        ("Switching costs",
         "Requalification, integration and retraining",
         s5(10 - ff.get("Intensity of Competition", 5)),
         "Locks in the installed base"),
        ("Technology substitution",
         "Alternative materials and next-generation formats",
         s5(ff.get("Threat of Substitutes", 5)),
         "Opens demand to Tier 2–3 innovators"),
        ("Supplier power",
         "Few qualified input and component suppliers",
         s5(ff.get("Bargaining Power of Suppliers", 5)),
         "Raises input cost for sub-scale players"),
    ]
    overall = round(sum(f[2] for f in forces) / len(forces), 1)
    # installed base (units), indicative: value share over each tier's
    # relative price point - premium tiers ship fewer units per $
    price = {"Tier 1": 1.3, "Tier 2": 1.0, "Tier 3": 0.85, "Others": 0.7}
    raw_u = {t: s0_[t] / price[t] for t in TIERS}
    units = {t: raw_u[t] / (sum(raw_u.values()) or 1) for t in TIERS}
    return {"u": u, "tiers": tiers, "mv0": mv0, "mv1": mv1, "units": units,
            "mkt_cagr": mkt_cagr, "v0": {t: tiers[t]["size"] for t in TIERS},
            "v1": v1, "s0": s0, "s1": s1, "n0": n0, "n1": n1, "up": up,
            "down": down, "entrants": entrants, "exits": exits, "hhi": hhi,
            "cr4": cr4, "top": shares[0] if shares else 0, "cls": cls,
            "forces": forces, "overall": overall}


# ═══════════════════════════ page 1 ══════════════════════════════════════
def _page1(m):
    u, tiers = m["u"], m["tiers"]
    s0, s1, v0, v1, n0 = m["s0"], m["s1"], m["v0"], m["v1"], m["n0"]
    P = []
    # tier ladder
    LX, LY, LW, LH = 15, 110, 610, 420
    n_players = sum(n0.values())
    P += _panel(LX, LY, LW, LH, f"Tier Ladder – Share of Market Value, "
                f"{EST_YEAR}",
                f"Market {fmt_num(m['mv0'])} {u}  ·  Tier 1 share "
                f"{fmt_pct(s0['Tier 1'])} → {fmt_pct(s1['Tier 1'])} by "
                f"{END_YEAR}")
    top_share = max(s0.values()) or 1
    bar_x, bar_max = LX + 118, LW - 118 - 16
    row_h = (LH - 60) / 4
    for i, t in enumerate(TIERS):
        y = LY + 56 + i * row_h
        P.append(_t(LX + 12, y + 2, 104, 22, t, 11, True, INK))
        P.append(_t(LX + 12, y + 24, 104, 34, tiers[t]["label"], 11, False,
                    GREY, anchor="t", wrap=True))
        # wide enough for its read-out; longer bars still read as bigger
        bw = 340 + (bar_max - 340) * s0[t] / top_share
        P.append(_r(bar_x, y, bw, 50, T_COL[t]))
        ink = WHITE if t != "Others" else INK
        # a box just the figure's width: a locked deck puts its padlock
        # exactly where the figure is written, at the bar's left
        P.append(_t(bar_x + 10, y + 3, 70, 22, fmt_pct(s0[t]), 11, True,
                    ink))
        avg = (f"avg {fmt_num(v0[t] / n0[t])} {u} / player" if n0[t]
               else "long tail")
        P.append(_t(bar_x + 10, y + 24, bw - 16, 26,
                    f"{fmt_num(v0[t])} {u}  ·  {_players(n0[t])}  ·  "
                    f"{avg}", 11, False, ink))
        names = (" · ".join(tiers[t]["players"]) if tiers[t]["players"]
                 else "Regional & niche suppliers")
        P.append(_t(bar_x, y + 51, bar_max, 36, names, 11, False, GREY,
                    anchor="t", wrap=True))
    # concentration read-out
    CY, CH = LY + LH + 12, 700 - (LY + LH + 12)
    P += _panel(LX, CY, LW, CH, "Concentration Read-out")
    bx, bw_ = LX + 14, (LW - 28) / 3
    labels = HHI_BANDS
    for k, (lab, col) in enumerate(zip(labels, BAND)):
        P.append(_r(bx + k * bw_, CY + 36, bw_, 26, col, WHITE))
        P.append(_t(bx + k * bw_, CY + 36, bw_, 26, lab, 11, True,
                    INK if k == 0 else WHITE, "c"))
    pos = min(max(m["hhi"], 0), 3000) / 3000.0
    mx = bx + pos * (3 * bw_)
    P.append(_t(mx - 90, CY + 63, 180, 26, f"▲ HHI = {m['hhi']:,.0f}",
                11, True, INK, "c"))
    ratio = s0["Tier 1"] / (s0["Others"] or 1)
    P.append(_t(LX + 10, CY + 86, LW - 20, 44,
                f"CR4 = {fmt_pct(m['cr4'] / 100)}  ·  Tier 1 : Others = "
                f"{ratio:.1f} : 1  ·  Top-player share = "
                f"{fmt_pct(m['top'] / 100)}  ·  Class: {m['cls']}",
                11, True, INK, "c"))
    # tier profile matrix
    RX, RW = LX + LW + 15, 1265 - (LX + LW + 15)
    P += _panel(RX, LY, RW, LH, "Tier Profile Matrix",
                "Size, growth and competitive characteristics of each tier")
    rows = [
        ("Players", [str(n0[t]) if n0[t] else "Long tail"
                     for t in TIERS], str(n_players) + "+"),
        (f"Value {EST_YEAR} ({u})", [fmt_num(v0[t]) for t in TIERS],
         fmt_num(m["mv0"])),
        # shares carry their % (a locked deck keeps it beside the padlock);
        # a Total of 100% is a definition and stays readable
        (f"Share of value {EST_YEAR}", [fmt_pct(s0[t]) for t in TIERS],
         "100%"),
        (f"Share of value {END_YEAR}", [fmt_pct(s1[t]) for t in TIERS],
         "100%"),
        ("Share change (pp)", [(s1[t] - s0[t]) * 100 for t in TIERS], "–"),
        (f"CAGR {EST_YEAR}–{END_YEAR}",
         [fmt_pct(_cagr(v0[t], v1[t])) for t in TIERS],
         fmt_pct(m["mkt_cagr"])),
        (f"Avg / player ({u})",
         [fmt_num(v0[t] / n0[t]) if n0[t] else "–" for t in TIERS], "–"),
        ("Installed base share (units)",
         [fmt_pct(m["units"][t]) for t in TIERS], "100%"),
    ] + [(k, list(v), "–") for k, v in _QUAL.items()]
    tx, ty = RX + 10, LY + 54
    lw = 170
    cw = (RW - 20 - lw) / 5
    hh = 26
    P.append(_r(tx, ty, lw, hh, T_COL["Tier 1"], WHITE))
    P.append(_t(tx + 6, ty, lw - 8, hh, "Metric", 11, True, WHITE))
    for k, (h, col) in enumerate(zip(list(TIERS) + ["Total"],
                                     [T_COL[t] for t in TIERS]
                                     + [TOTAL_HEAD])):
        P.append(_r(tx + lw + k * cw, ty, cw, hh, col, WHITE))
        P.append(_t(tx + lw + k * cw, ty, cw, hh, h, 11, True,
                    INK if h == "Others" else WHITE, "c"))
    yy = ty + hh
    for name, vals, tot in rows:
        qual = name in _QUAL
        # a label too long for one line gets a two-line row
        rh = 34 if qual or len(name) > 22 else 24
        post = name == "Strategic posture"
        P.append(_r(tx, yy, lw, rh, TINT, GRID, 0.5))
        P.append(_t(tx + 6, yy, lw - 8, rh, name, 11, True, INK, wrap=True))
        for k, v in enumerate(vals + [tot]):
            fill = TEAL_TINT if post and k < 4 else WHITE
            cx0 = tx + lw + k * cw
            P.append(_r(cx0, yy, cw, rh, fill, GRID, 0.5))
            if isinstance(v, float):
                # a signed change: the sign in a box of its own, just left
                # of the figure, so a locked deck keeps '+' / '-' beside
                # the padlock that replaces the number
                num = f"{abs(v):.1f}"
                half = len(num) * 3.4
                P.append(_t(cx0 + cw / 2 - half - 16, yy, 14, rh,
                            "+" if v >= 0 else "−", 11, False, INK, "r"))
                P.append(_t(cx0 + cw / 2 - half - 2, yy, 2 * half + 4, rh,
                            num, 11, False, INK, "c"))
                continue
            if (isinstance(v, str) and v.endswith("%") and v != "100%"
                    and v[:-1].replace(".", "").replace("-", "").isdigit()):
                # a share: the figure centred in the cell, its '%' just to
                # the right in a box of its own - so a locked deck puts the
                # padlock on the column's centre line, the '%' beside it
                num = v[:-1]
                half = len(num) * 3.4
                P.append(_t(cx0 + cw / 2 - half - 2, yy, 2 * half + 4, rh,
                            num, 11, k == 4, INK, "c"))
                P.append(_t(cx0 + cw / 2 + half + 1, yy, 14, rh, "%", 11,
                            k == 4, INK, "l"))
                continue
            P.append(_t(cx0 + 2, yy, cw - 4, rh, v, 11,
                        post or k == 4, INK, "c", wrap=qual))
        yy += rh
    # key takeaways
    fast = max(("Tier 1", "Tier 2", "Tier 3"),
               key=lambda t: _cagr(v0[t], v1[t]))
    tail = s0["Tier 3"] + s0["Others"]
    avg1 = v0["Tier 1"] / (n0["Tier 1"] or 1)
    avg3 = v0["Tier 3"] / (n0["Tier 3"] or 1)
    take = [
        f"Tier 1 ({fmt_pct(s0['Tier 1'])}) holds the largest single block of "
        f"value on a {fmt_num(m['mv0'])} {u} base; CR4 of "
        f"{fmt_pct(m['cr4'] / 100)} and HHI of {m['hhi']:,.0f} classify "
        f"the market as {m['cls'].lower()}.",
        f"{fast} grows fastest at {fmt_pct(_cagr(v0[fast], v1[fast]))} CAGR, "
        f"moving {(s1[fast] - s0[fast]) * 100:+.1f} pp of share by "
        f"{END_YEAR}.",
        f"Tier 3 and the long tail together hold {fmt_pct(tail)} of value; "
        f"the average Tier 3 player is {avg1 / (avg3 or 1):.1f}x smaller than "
        f"a Tier 1 supplier.",
    ]
    P += _panel(RX, CY, RW, CH, "Key Takeaways")
    P += _bullets(RX + 12, CY + 32, RW - 24, take, gap=2)
    return P


# ═══════════════════════════ page 2 ══════════════════════════════════════
def _page2(m):
    u = m["u"]
    s0, s1, n0, n1 = m["s0"], m["s1"], m["n0"], m["n1"]
    tiers = m["tiers"]
    P = []
    X, Y, W, H = 15, 110, 1250, 250
    P += _panel(X, Y, W, H, f"Structural Shift, {EST_YEAR} → {END_YEAR} – "
                f"Tier Migration and Entry / Exit")
    bw, gap = 200, (W - 24 - 4 * 200) / 3
    by, bh = Y + 36, 92
    for i, t in enumerate(TIERS):
        bx = X + 12 + i * (bw + gap)
        ink = WHITE if t != "Others" else INK
        P.append(_r(bx, by, bw, bh, T_COL[t], None, r=8))
        P.append(_t(bx, by + 6, bw, 22, t, 11, True, ink, "c"))
        P.append(_t(bx, by + 30, bw, 18, tiers[t]["label"], 11, False, ink,
                    "c"))
        players = (f"{n0[t]} → {n1[t]} players" if t != "Others"
                   else "Long tail")
        P.append(_t(bx + 4, by + 46, bw - 8, 44,
                    f"{players}  ·  {fmt_pct(s0[t])} → {fmt_pct(s1[t])} share",
                    11, True, ink, "c", wrap=True))
        if i < 3:
            nxt = TIERS[i + 1]
            ax1, ax2 = bx + bw + 8, bx + bw + gap - 8
            upn = m["up"].get((nxt, t), 0)
            dn = m["down"].get((t, nxt), 0)
            P.append(_t(ax1, by + 8, ax2 - ax1, 18,
                        f"{upn} up-tier (M&A)", 11, True, TEAL, "c"))
            P += _arrow(ax1, ax2, by + 32, TEAL, left=True)
            P += _arrow(ax1, ax2, by + 60, ORANGE, left=False)
            P.append(_t(ax1, by + 66, ax2 - ax1, 18,
                        f"{dn} down / exit", 11, True, ORANGE, "c"))
    moved = sum(max(0.0, s1[t] - s0[t]) for t in TIERS) * 100
    gain_t = max(TIERS, key=lambda t: s1[t] - s0[t])
    lose_t = min(TIERS, key=lambda t: s1[t] - s0[t])
    cards = [
        (str(m["entrants"]), f"New entrants {EST_YEAR}–{END_YEAR}",
         "scaling up from the long tail"),
        (str(m["exits"]), "Exits / acquisitions",
         "absorbed mainly by Tier 1–2 acquirers"),
        (f"{m['entrants'] - m['exits']:+d}", "Net change in player count",
         f"Between {EST_YEAR} and {END_YEAR}"),
        (f"{moved:.1f} pp", "Share moving between tiers",
         f"largest transfer: {lose_t} → {gain_t}"),
    ]
    cw_, cg = (W - 24 - 3 * 14) / 4, 14
    cy = by + bh + 14
    for i, (big, lab, sub) in enumerate(cards):
        cx = X + 12 + i * (cw_ + cg)
        P.append(_r(cx, cy, cw_, H - (cy - Y) - 12, TINT, None, r=8))
        # the headline figure centred in its own box, where its padlock goes
        fig_ = _t(cx + 10, cy + 6, 80, 36, big, 11, True, T_COL["Tier 1"],
                  "c")
        fig_["name"] = "FMI KPI Value"       # see the lock rulebook
        P.append(fig_)
        P.append(_t(cx + 98, cy + 6, cw_ - 106, 36, lab, 11, True, INK,
                    anchor="t", wrap=True))
        P.append(_t(cx + 98, cy + 42, cw_ - 106, 40, sub, 11, False, GREY,
                    anchor="t", wrap=True))
    # structural forces table
    FY = Y + H + 12
    FW, FH = 745, 700 - FY
    P += _panel(X, FY, FW, FH, "Structural Forces Shaping the Tier Structure "
                "(1 = weak, 5 = strong)")
    cols = [("Force", 150, T_COL["Tier 1"]),
            ("Drivers in the market", 245, T_COL["Tier 1"]),
            ("Score", 70, T_COL["Tier 1"]),
            ("Effect on tier structure", FW - 20 - 150 - 245 - 70, TEAL)]
    tx, ty, hh = X + 10, FY + 34, 26
    cx = tx
    for name, w, col in cols:
        P.append(_r(cx, ty, w, hh, col, WHITE))
        P.append(_t(cx + 6, ty, w - 8, hh, name, 11, True, WHITE))
        cx += w
    rows = [(f, d, f"{s} / 5", e) for f, d, s, e in m["forces"]]
    rows.append(("Overall structural pressure", "Average of the five forces",
                 f"{m['overall']:.1f} / 5",
                 "Consolidation towards Tier 1–2"
                 if m["overall"] >= 3 else "Structure broadly stable"))
    rh = (FH - 34 - hh - 10) / len(rows)
    yy = ty + hh
    for i, r in enumerate(rows):
        last = i == len(rows) - 1
        cx = tx
        for k, (v, (name, w, col)) in enumerate(zip(r, cols)):
            fill = TINT if k == 0 else TEAL_TINT if k == 2 else WHITE
            P.append(_r(cx, yy, w, rh, fill, GRID, 0.5))
            P.append(_t(cx + 6, yy, w - 10, rh, v, 11, k in (0, 2) or last,
                        INK, "c" if k == 2 else "l", wrap=True))
            cx += w
        yy += rh
    # key takeaways
    ranked = sorted(m["forces"], key=lambda f: f[2], reverse=True)
    take = [
        f"The market is {m['cls'].lower()} (HHI {m['hhi']:,.0f}); Tier 1 "
        f"sets pricing and innovation benchmarks and is expected to hold "
        f"{fmt_pct(s1['Tier 1'])} of value in {END_YEAR}.",
        f"{m['up'][('Tier 2', 'Tier 1')]} Tier 2 player is positioned to move "
        f"up-tier by {END_YEAR} through M&A and key-account wins – the main "
        f"structural change of the decade.",
        f"Tier 3 and the long tail face the highest structural pressure "
        f"({m['overall']:.1f} / 5); {m['exits']} players are expected to "
        f"exit or be acquired.",
        f"{ranked[0][0]} ({ranked[0][2]} / 5) and {ranked[1][0].lower()} "
        f"({ranked[1][2]} / 5) are the dominant forces shaping the tier "
        f"structure.",
    ]
    KX = X + FW + 15
    P += _panel(KX, FY, 1265 - KX, FH, "Key Takeaways")
    P += _bullets(KX + 12, FY + 32, 1265 - KX - 24, take)
    return P


def industry_landscape():
    """The two 'Industry Structure Landscape' pages (see module doc)."""
    m = _model()
    return (f"Industry Structure Landscape – {MKT}, Global {EST_YEAR}",
            {"primpages": [_page1(m), _page2(m)]})
