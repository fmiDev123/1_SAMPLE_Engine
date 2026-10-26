"""Competitive Landscape: ten competitive-intelligence pages (2026-10-01).

Laid out as the client's standard reference, the PDF
'Aerial_Imaging_Competitive_Deck_10_pages': the same page designs, table
colours, rating circles (Harvey balls) and heatmaps, scaled 1.6x onto this
deck's 1280x720 canvas under its own header.  Drawn with the shared
primitives ({"primpages": ...}) so the deck, the dashboard and the lock all
read them.  All text 11pt (the deck's content font).

Every figure and rating comes from `ci_data.build()` (see there for the
sources); a field only research can supply reads 'n/a', as the reference
does.

    1  Competitive Landscape at a Glance        (one-page summary)
    2  Segment Leadership Map                   (who holds / who leads)
    3  Value Chain Positioning                  (stages, leads, watch)
    4  Player Landscape                         (type x reach)
    5  Corporate Scorecard   (Level 1, who they are)
    6  Offering Scorecard    (Level 2, what they sell)
    7  Execution Scorecard   (Level 3, how they win)
    8  Success Factors and Composite Ranking
    9  Strategic Activity Tracker
    10 Strategic Activity by Function            (heatmaps + departments)
"""
from __future__ import annotations

import math
import re

from . import ci_data as X
from .components import fmt_num, fmt_pct
from .config import END_YEAR, EST_YEAR, UNIT_LABELS
from .data import MARKET_LABEL as MKT

# ── the reference deck's colours ────────────────────────────────────────────
NAVY = "#0F2D52"
ORANGE = "#E87722"
INK = "#1F2933"
GREY = "#5F6B7A"
RULE = "#D5DDE6"
ROW_RULE = "#E3E8EE"
DARK_RED = "#8B1E2D"
TYPE_COLS = ("#1F6FA8", "#E07B24", "#2E8B57", "#8E3A8E", "#6B7280")
GROUP_COLS = ("#0F2D52", "#1F6FA8", "#6FA0CF", "#A9C6E3")
TIER_COLS = {"Tier 1": "#0F2D52", "Tier 2": "#1F6FA8", "Tier 3": "#A9C6E3",
             "Others": "#C9D2D9"}
HEAT = ("#CFE0F1", "#8FB4DC", "#3D73AE")
BADGE = {"MA": "#8E3A8E", "PR": "#1F6FA8", "EX": "#6B7280", "RD": "#0F2D52",
         "RG": "#C0392B", "CP": "#17768A"}
WHITE = "#FFFFFF"

X0, X1 = 20, 1260                 # content band
TOP = 110                         # first line under the deck's header


# ── primitives ──────────────────────────────────────────────────────────────
def _t(x, y, w, h, s, bold=False, color=INK, align="l", anchor="t",
       wrap=True):
    p = {"t": "text", "x": x, "y": y, "w": w, "h": h, "s": str(s),
         "size": 11, "bold": bold, "color": color, "align": align}
    p["anchor"] = anchor or "t"
    if wrap:
        p["wrap"] = True
    return p


def _r(x, y, w, h, fill, line=None, lw=0.75):
    return {"t": "rect", "x": x, "y": y, "w": w, "h": h, "fill": fill,
            "line": line, "lw": lw}


def _hl(x1, x2, y, color=RULE, width=0.75):
    return {"t": "line", "x1": x1, "y1": y, "x2": x2, "y2": y,
            "color": color, "width": width}


def _vl(x, y1, y2, color=RULE, width=0.75):
    return {"t": "line", "x1": x, "y1": y1, "x2": x, "y2": y2,
            "color": color, "width": width}


def _harvey(cx, cy, level, d=16):
    """A rating circle: 3 full, 2 half, 1 quarter, 0 empty, None 'n/a'."""
    if level is None:
        return [_t(cx - 20, cy - 9, 40, 18, "n/a", color=GREY, align="c",
                   anchor=None)]
    out = [{"t": "dot", "cx": cx, "cy": cy, "d": d,
            "color": NAVY if level == 3 else WHITE, "edge": NAVY}]
    if level in (1, 2):
        r = d / 2.0
        sweep = 180 if level == 2 else 90
        pts = [(cx, cy)]
        for k in range(0, sweep + 1, 10):
            a = math.radians(-90 + k)           # clockwise from 12 o'clock
            pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
        out.append({"t": "shape", "pts": pts, "fill": NAVY, "line": NAVY,
                    "lw": 0.5})
    return out


def _legend(x, y):
    """The reference's rating key."""
    P, xx = [], x
    for lvl, word in ((3, "Strong, direct evidence"),
                      (2, "Moderate or partial"), (1, "Weak or indirect"),
                      (0, "None found")):
        P += _harvey(xx + 8, y + 9, lvl, d=14)
        P.append(_t(xx + 20, y, 170, 18, word, anchor="m", wrap=False))
        xx += 44 + len(word) * 6.6
    return P


def _reading(y, text, h=24):
    return [_t(X0, y, 300, 16, "READING", bold=True, color=NAVY,
               anchor=None, wrap=False),
            _t(X0, y + 16, X1 - X0, h, text)]


def _short(name):
    return re.sub(r"\s*\([^)]*\)", "", str(name)).strip() or str(name)


def _cagr(a, b, n=END_YEAR - EST_YEAR):
    return (b / a) ** (1.0 / n) - 1.0 if a > 0 and b > 0 else float("nan")


def _ci():
    return X.build()


def _gk(c):
    """The group a player is shown under: its business type - or, when the
    roster names more than four distinct types (free-text descriptions, one
    per company), its tier, so every page still reads in a few groups
    (2026-10-01, the Dialysis roster)."""
    types = {x["type"] for x in _ci()["companies"]}
    if None not in types and len(types) <= 4:
        return c["type"]
    return c["tier"]


def _country(hq):
    """'Bad Homburg, Germany' -> 'Germany' - the HQ line stays one line."""
    return str(hq).split(",")[-1].strip() if hq else ""


def _type_colour(ci):
    types = list(dict.fromkeys(_gk(c) for c in ci["companies"]))
    return {tp: TYPE_COLS[i % len(TYPE_COLS)] for i, tp in enumerate(types)}


def _name_px(s, size=11.0, bold=True) -> float:
    """Width the native text box needs for ``s`` at ``size`` pt - the layout
    width model at the deck's px-per-pt factor, with a little safety."""
    from . import layout
    return layout.text_width(s, size * 1.33 * (1.05 if bold else 1.0)) + 6


def _company(x, y, h, c, colour, w=150):
    """A company cell: the type-coloured bar, the name, tier and HQ.

    The name stays on ONE line: a name wider than the cell is set a little
    smaller (never below 9pt), so it can no longer wrap onto the tier line
    underneath.  Callers widen the cell first (see ``_scorecard``)."""
    sub = c["tier"] + (f"  {_country(c['hq'])}" if c.get("hq") else "")
    size = 11.0
    while size > 9.0 and _name_px(c["name"], size) > w - 12:
        size -= 0.5
    nm = _t(x + 10, y + 2, w - 12, 18, c["name"], bold=True,
            anchor=None, wrap=False)
    nm["size"] = size
    return [_r(x, y + 3, 4, h - 6, colour), nm,
            _t(x + 10, y + 19, w - 12, 16, sub, color=GREY, anchor=None,
               wrap=False)]


def _top(ci, k):
    return sum(sorted((c["share"] for c in ci["companies"]),
                      reverse=True)[:k])


# ═══════════════════════ 1. At a glance ════════════════════════════════
def ci_glance():
    ci = _ci()
    cos, prod, end = ci["companies"], ci["prod_subs"], ci["end_subs"]
    mv0, mv1 = ci["mv0"] or 1, ci["mv1"]
    etot = sum(r[1] for r in end) or 1
    top3, top5 = _top(ci, 3), _top(ci, 5)
    p3 = " / ".join(f"{v / mv0 * 100:.0f}" for _n, v, _w in prod[:3]) + "%"
    e0 = end[0] if end else ("n/a", 0, 0)
    lead = max(cos, key=lambda c: c["share"])
    kpis = [
        (f"Top 3 share, {EST_YEAR}", f"{top3:.0f}%", ORANGE),
        ("Top 5 share", f"{top5:.0f}%", NAVY),
        ("Top 3 product types", p3, NAVY),
        ("Largest end use", f"{_short(e0[0])}, {e0[1] / etot * 100:.0f}%",
         NAVY),
        ("Named competitors", str(len(cos)), NAVY),
        ("Moves tracked", str(len(ci["activity"])), NAVY),
        ("Largest player", f"{lead['name']}", NAVY),
        (f"CAGR {EST_YEAR}–{END_YEAR}", fmt_pct(_cagr(mv0, mv1)),
         NAVY),
    ]
    P = [_hl(X0, X1, TOP, NAVY, 1.5)]
    cw = (X1 - X0) / len(kpis)
    for i, (lab, val, col) in enumerate(kpis):
        x = X0 + i * cw
        if i:
            P.append(_vl(x, TOP + 8, TOP + 58))
        P.append(_t(x + 6, TOP + 6, cw - 10, 18, lab, color=GREY,
                    anchor=None, wrap=False))
        P.append(_t(x + 6, TOP + 26, cw - 10, 42, val, bold=True,
                    color=col, anchor=None))
    P.append(_hl(X0, X1, TOP + 72, NAVY, 1.5))

    cols = [c for _k, c in ci["lead_cols"]]
    n_lead = len({c["name"] for c in cos if c["led"]})
    max_led = max((c["led"] for c in cos), default=0)
    t = ci["tiers"]
    fast = max(prod, key=lambda r: _cagr(r[1], r[2])) if prod else None
    rank = sorted(cos, key=lambda c: c["composite"], reverse=True)
    reach = {r: sum(1 for c in cos if c["reach"] == r)
             for r in ("Global", "Regional", "Local or niche")}
    n_acq = sum(1 for c in cos if "A" in c["active"])
    blocks = [
        ("Leadership is concentrated in value, spread in capability",
         (f"The top three hold {top3:.0f}% of {EST_YEAR} value, but {n_lead} "
          f"companies lead the {len(cols)} sub-segments tracked; no one "
          f"leads more than {max_led}.") if n_lead > 1 else
         (f"The top three hold {top3:.0f}% of {EST_YEAR} value, and one "
          f"company leads {max_led} of the {len(cols)} sub-segments "
          f"tracked.")),
        ("The tier structure is steep",
         f"Tier 1 ({len(t['Tier 1']['players'])} players) holds "
         f"{fmt_pct(t['Tier 1']['share'])}; Tier 2 and Tier 3 together hold "
         f"{fmt_pct(t['Tier 2']['share'] + t['Tier 3']['share'])}."),
        ("Growth is shifting inside the product mix",
         (f"'{_short(prod[0][0])}' is the largest line; "
          f"'{_short(fast[0])}' compounds fastest at "
          f"{fmt_pct(_cagr(fast[1], fast[2]))} CAGR.") if fast else "n/a"),
        ("Reach divides the field",
         f"{reach['Global']} players compete globally, {reach['Regional']} "
         f"regionally and {reach['Local or niche']} in one market or niche."),
        ("Consolidation is an active lever",
         f"Acquisition is an active strategy at {n_acq} of {len(cos)} named "
         f"players; the independent target list is shrinking."),
        ("Success factors, in order",
         "1 segment breadth, 2 scale and share, 3 geographic reach, "
         "4 innovation pipeline, 5 strategic momentum.  Ranking leads: "
         + ", ".join(c["name"] for c in rank[:3]) + "."),
    ]
    bw = (X1 - X0 - 40) / 2
    for i, (head, body) in enumerate(blocks):
        col, row = i // 3, i % 3
        x = X0 + col * (bw + 40)
        y = TOP + 90 + row * 116
        P.append(_r(x, y + 2, 22, 22, NAVY))
        P.append(_t(x, y + 2, 22, 22, str(i + 1), bold=True, color=WHITE,
                    align="c", anchor="m", wrap=False))
        P.append(_t(x + 32, y, bw - 32, 20, head, bold=True, color=INK,
                    anchor=None, wrap=False))
        P.append(_t(x + 32, y + 22, bw - 32, 80, body))
    yb = TOP + 90 + 3 * 116
    P.append(_hl(X0, X1, yb, RULE))
    P.append(_t(X0, yb + 6, 400, 16, "HOW THIS SECTION IS BUILT",
                bold=True, color=NAVY, anchor=None, wrap=False))
    P.append(_t(X0, yb + 24, X1 - X0, 60,
                "Pages 2 to 4 show where the market is and who plays.  "
                "Pages 5 to 7 rate every named company on twenty Ps (who "
                "they are, what they sell, how they win); a cell reads n/a "
                "where no evidence exists.  Page 8 turns the pattern into "
                "five success factors and a ranking.  Pages 9 and 10 list "
                "what each player is doing and what each of your "
                "departments should do about it."))
    return (f"{MKT} Competitive Landscape at a Glance",
            {"primpages": [P]})


# ═══════════════════════ 2. Segment leadership ═════════════════════════
def ci_leadership():
    ci = _ci()
    cos = ci["companies"]
    tcol = _type_colour(ci)
    srt = sorted(cos, key=lambda c: c["share"], reverse=True)
    groups = [("Top 3", srt[:3]), ("Rest of top 5", srt[3:5]),
              ("Emerging", srt[5:]), ("Others", [])]
    shares = [sum(c["share"] for c in g) for _n, g in groups[:3]]
    shares.append(max(0.0, 100.0 - sum(shares)))
    P = [_t(X0, TOP, 600, 18, f"Who holds the market, {EST_YEAR}",
            bold=True, color=NAVY, anchor=None, wrap=False)]
    x, y, h = X0, TOP + 22, 28
    for (name, g), sh, col in zip(groups, shares, GROUP_COLS):
        w = (X1 - X0) * sh / 100.0
        if w <= 1:
            continue
        P.append(_r(x, y, w, h, col, WHITE))
        if w < 70:                 # too narrow to name: the legend does
            x += w
            continue
        P.append(_t(x + 6, y, w - 8, h, f"{name}  {sh:.0f}%", bold=True,
                    color=WHITE if col != GROUP_COLS[3] else NAVY,
                    anchor="m", wrap=False))
        x += w
    lw = (X1 - X0) / 4
    for k, ((name, g), col) in enumerate(zip(groups, GROUP_COLS)):
        names = ", ".join(c["name"] for c in g) or "Long tail of regional " \
                                                    "and niche suppliers"
        lx = X0 + k * lw
        P.append(_r(lx, y + h + 8, 9, 9, col))
        P.append(_t(lx + 13, y + h + 2, lw - 20, 36, names))
    yb = y + h + 42
    P.append(_hl(X0, X1, yb, NAVY, 1.5))
    cols = [c for _k, c in ci["lead_cols"]]
    P.append(_t(X0, yb + 6, 900, 18,
                f"Who leads which sub-segment, {EST_YEAR}  (full = leader, "
                f"half = active)", bold=True, color=NAVY, anchor=None,
                wrap=False))
    co_w = min(300.0, max(200.0, max((_name_px(c["name"]) + 16
                                      for c in cos), default=200)))
    cx0, cwid = X0 + co_w, (X1 - X0 - co_w) / len(cols)
    hy = yb + 28
    for j, col in enumerate(cols):
        P.append(_t(cx0 + j * cwid + 2, hy - 8, cwid - 4, 50, _short(col),
                    bold=True, color=NAVY, align="c", anchor="b"))
    ry, rh = hy + 44, 36
    lvl = {3: 3, 2: 2, 0: 0}
    for i, c in enumerate(cos):
        y0 = ry + i * rh
        P.append(_hl(X0, X1, y0, ROW_RULE))
        P += _company(X0, y0, rh, c, tcol[_gk(c)], w=co_w)
        for j, col in enumerate(cols):
            P += _harvey(cx0 + j * cwid + cwid / 2, y0 + rh / 2,
                         lvl.get(ci["leadership"][c["name"]][col], 0))
    ye = ry + len(cos) * rh
    P.append(_hl(X0, X1, ye, ROW_RULE))
    lx = X0
    for tp, col in tcol.items():
        P.append(_r(lx, ye + 9, 9, 9, col))
        P.append(_t(lx + 13, ye + 4, 220, 18, tp, anchor=None, wrap=False))
        lx += 30 + len(tp) * 6.5
    counts = {_short(col): sum(1 for c in cos
                               if ci["leadership"][c["name"]][col] >= 2)
              for col in cols}
    most, least = max(counts, key=counts.get), min(counts, key=counts.get)
    P += _reading(ye + 26,
                  f"Read across a row to see breadth; down a column to see "
                  f"contestability.  '{most}' is the most contested column "
                  f"({counts[most]} players active); '{least}' is the most "
                  f"open ({counts[least]}).")
    return (f"{MKT} Segment Leadership Map", {"primpages": [P]})


# ═══════════════════════ 3. Value chain positioning ════════════════════
def ci_value_chain():
    ci = _ci()
    cos, prod, end = ci["companies"], ci["prod_subs"], ci["end_subs"]
    mv0 = ci["mv0"] or 1
    lead = ci["leadership"]
    kpis = [(f"Top 3 share, {EST_YEAR}", f"{_top(ci, 3):.0f}%", ORANGE),
            ("Top 5 share", f"{_top(ci, 5):.0f}%", NAVY)] + \
        [(f"{_short(n)}, share", f"{v / mv0 * 100:.0f}%", NAVY)
         for n, v, _w in prod[:3]]
    P = []
    cw = (X1 - X0) / len(kpis)
    for i, (lab, val, col) in enumerate(kpis):
        x = X0 + i * cw
        if i:
            P.append(_vl(x, TOP + 4, TOP + 46))
        P.append(_t(x + 6, TOP, cw - 10, 18, lab, color=GREY, anchor=None,
                    wrap=False))
        P.append(_t(x + 6, TOP + 20, cw - 10, 24, val, bold=True,
                    color=col, anchor=None, wrap=False))
    P.append(_hl(X0, X1, TOP + 52, RULE))
    by, bh = TOP + 62, 390
    gap = 26
    stages = []
    for i, (name, v0, v1) in enumerate(prod[:3], start=1):
        leads = [c["name"] for c in cos if lead[c["name"]].get(name) == 3]
        also = [c["name"] for c in cos if lead[c["name"]].get(name) == 2][:4]
        stages.append((f"{i}. {_short(name)}",
                       f"{v0 / mv0 * 100:.0f}%",
                       fmt_pct(_cagr(v0, v1)),
                       leads, also,
                       f"{len(leads) + len(also)} named players compete "
                       f"here; '{_short(name)}' adds "
                       f"{fmt_num(v1 - v0)} {UNIT_LABELS['Value']} by "
                       f"{END_YEAR}."))
    # Box count is len(stages) product boxes + 1 application box - some
    # markets only break "Product Type" into 2 sub-segments (not 3), and
    # sizing the row for a fixed 4 boxes left the 3rd slot blank in that
    # case (2026-10-22, Bentonite Market).
    n_boxes = len(stages) + 1
    bw = (X1 - X0 - (n_boxes - 1) * gap) / n_boxes
    for i, st in enumerate(stages):
        x = X0 + i * (bw + gap)
        P.append(_r(x, by, bw, bh, WHITE, RULE))
        P.append(_r(x, by, bw, 5, NAVY))
        P.append(_t(x + 10, by + 10, bw - 20, 18, st[0], bold=True,
                    color=NAVY, anchor=None, wrap=False))
        P.append(_t(x + 10, by + 30, 70, 20, st[1], bold=True,
                    wrap=False))
        P.append(_t(x + 82, by + 30, bw - 92, 20, "of market", bold=True,
                    wrap=False))
        P.append(_t(x + 10, by + 52, 70, 20, st[2], color=GREY,
                    wrap=False))
        P.append(_t(x + 82, by + 52, bw - 92, 20,
                    f"CAGR {EST_YEAR}–{END_YEAR}", color=GREY, wrap=False))
        P.append(_hl(x + 10, x + bw - 10, by + 76))
        yy = by + 82
        P.append(_t(x + 10, yy, bw - 20, 16, "LEADS", bold=True,
                    color=NAVY, anchor=None, wrap=False))
        yy += 18
        for n in st[3] or ["n/a"]:
            P.append(_t(x + 10, yy, bw - 20, 18, f"●  {n}", bold=True,
                        anchor=None, wrap=False))
            yy += 18
        yy += 6
        P.append(_t(x + 10, yy, bw - 20, 16, "ALSO ACTIVE", bold=True,
                    color=NAVY, anchor=None, wrap=False))
        yy += 18
        for n in st[4] or ["n/a"]:
            P.append(_t(x + 10, yy, bw - 20, 18, f"●  {n}", anchor=None,
                        wrap=False, color=INK))
            yy += 18
        P.append(_hl(x + 10, x + bw - 10, by + bh - 96))
        P.append(_t(x + 10, by + bh - 90, bw - 20, 16, "WHAT TO WATCH",
                    bold=True, color=NAVY, anchor=None, wrap=False))
        P.append(_t(x + 10, by + bh - 72, bw - 20, 70, st[5]))
        ax = x + bw + 4
        P.append(_hl(ax, ax + gap - 10, by + bh / 2, NAVY, 2))
        P.append({"t": "shape", "pts": [(ax + gap - 8, by + bh / 2),
                                        (ax + gap - 15, by + bh / 2 - 5),
                                        (ax + gap - 15, by + bh / 2 + 5)],
                  "fill": NAVY, "line": NAVY})
    # end use - always the box right after the last product stage
    x = X0 + len(stages) * (bw + gap)
    P.append(_r(x, by, bw, bh, WHITE, RULE))
    P.append(_r(x, by, bw, 5, NAVY))
    P.append(_t(x + 10, by + 10, bw - 20, 18, f"{n_boxes}. {ci['end_seg']}",
                bold=True, color=NAVY, anchor=None, wrap=False))
    P.append(_t(x + 10, by + 30, bw - 20, 20, "By application", bold=True,
                anchor=None, wrap=False))
    P.append(_t(x + 10, by + 52, bw - 20, 20, "Who buys the output.",
                color=GREY, anchor=None, wrap=False))
    P.append(_hl(x + 10, x + bw - 10, by + 76))
    etot = sum(r[1] for r in end) or 1
    mx = max((r[1] for r in end), default=1) / etot
    for k, (n, v, _w) in enumerate(end[:6]):
        yy = by + 82 + k * 34
        sh = v / etot
        P.append(_t(x + 10, yy, 140, 34, _short(n), anchor="m"))
        P.append(_r(x + 152, yy + 11, (bw - 210) * sh / (mx or 1), 12,
                    "#1F6FA8"))
        P.append(_t(x + 156 + (bw - 210) * sh / (mx or 1), yy + 7, 50, 20,
                    f"{sh * 100:.0f}%", bold=True, anchor="m", wrap=False))
    top2 = sum(r[1] for r in end[:2]) / etot
    P.append(_hl(x + 10, x + bw - 10, by + bh - 96))
    P.append(_t(x + 10, by + bh - 90, bw - 20, 16, "WHAT TO WATCH",
                bold=True, color=NAVY, anchor=None, wrap=False))
    P.append(_t(x + 10, by + bh - 72, bw - 20, 70,
                f"The top two end uses pay for {top2 * 100:.0f}% of the "
                f"market; vendors package for one vertical at a time."))
    try:
        from . import dynamics as D
        enab = "  |  ".join(d["title"] for d in D.items("drivers")[:3])
        cons = "  |  ".join(d["title"] for d in D.items("restraints")[:3])
    except Exception:
        enab = cons = "n/a"
    for k, (lab, txt, col) in enumerate((("ENABLERS", enab, NAVY),
                                         ("CONSTRAINTS", cons, DARK_RED))):
        y = by + bh + 12 + k * 48
        P.append(_r(X0, y, X1 - X0, 42, WHITE, RULE))
        P.append(_r(X0, y, 5, 42, col))
        P.append(_t(X0 + 12, y, 110, 42, lab, bold=True, color=col,
                    anchor="m", wrap=False))
        P.append(_t(X0 + 130, y, X1 - X0 - 140, 42, txt, anchor="m"))
    return (f"{MKT} Value Chain Positioning of Leading Players",
            {"primpages": [P]})


# ═══════════════════════ 4. Player landscape ═══════════════════════════
def ci_players():
    ci = _ci()
    cos = ci["companies"]
    tcol = _type_colour(ci)
    types = list(tcol)[:4]
    reaches = (("Global", "Sells in most major regions"),
               ("Regional", "Strong in one or two regions"),
               ("Local or niche", "One market or one specialist use"))
    lx, cx0 = X0, X0 + 150
    cw = (X1 - cx0) / len(types)
    P = []
    for j, tp in enumerate(types):
        n = sum(1 for c in cos if _gk(c) == tp)
        P.append(_r(cx0 + j * cw + 2, TOP, cw - 4, 28, tcol[tp]))
        P.append(_t(cx0 + j * cw + 10, TOP, cw - 16, 28, f"{tp}  ({n})",
                    bold=True, color=WHITE, anchor="m", wrap=False))
    rh = 108
    for i, (rname, rdesc) in enumerate(reaches):
        y = TOP + 34 + i * rh
        n = sum(1 for c in cos if c["reach"] == rname)
        P.append(_t(lx, y + 4, 145, 18, f"{rname} ({n})", bold=True,
                    color=NAVY, anchor=None, wrap=False))
        P.append(_t(lx, y + 22, 140, 40, rdesc, color=GREY, anchor=None))
        if i:
            P.append(_hl(lx, X1, y, ROW_RULE))
        for j, tp in enumerate(types):
            x = cx0 + j * cw
            P.append(_vl(x + 3, y + 2, y + rh - 4, tcol[tp], 2.5))
            mem = [c for c in cos if _gk(c) == tp
                   and c["reach"] == rname]
            for k, c in enumerate(mem[:5]):
                yy = y + 6 + k * 20
                tnum = c["tier"].replace("Tier ", "") \
                    if c["tier"] != "Others" else "–"
                P.append(_r(x + 12, yy + 2, 15, 15, TIER_COLS[c["tier"]]))
                P.append(_t(x + 12, yy + 1, 15, 17, tnum, bold=True,
                            color=WHITE if c["tier"] in ("Tier 1", "Tier 2")
                            else NAVY, align="c", anchor="m", wrap=False))
                P.append(_t(x + 32, yy, cw - 40, 18,
                            c["name"] + (f" ({_country(c['hq'])})" if c["hq"] else ""),
                            bold=c["tier"] == "Tier 1", anchor=None,
                            wrap=False))
    yl = TOP + 34 + 3 * rh + 4
    xx = X0
    for tr in ("Tier 1", "Tier 2", "Tier 3"):
        sh = ci["tiers"][tr]["share"]
        P.append(_r(xx, yl + 3, 13, 13, TIER_COLS[tr]))
        P.append(_t(xx + 18, yl - 2, 50, 22, tr, anchor="m", wrap=False))
        P.append(_t(xx + 66, yl - 2, 64, 22, fmt_pct(sh), bold=True,
                    anchor="m", wrap=False))
        P.append(_t(xx + 132, yl - 2, 130, 22, "combined share",
                    anchor="m", wrap=False))
        xx += 290
    by_type = "   |   ".join(
        f"{tp} {sum(1 for c in cos if _gk(c) == tp)}"
        for tp in types)
    by_reach = "   |   ".join(
        f"{r} {sum(1 for c in cos if c['reach'] == r)}" for r, _d in reaches)
    P.append(_t(X0, yl + 22, X1 - X0, 18, f"BY TYPE   {by_type}      BY "
                f"REACH   {by_reach}", anchor=None, wrap=False))
    yb = yl + 50
    bw = (X1 - X0 - 3 * 20) / 4
    for j, tp in enumerate(types):
        mem = [c for c in cos if _gk(c) == tp]
        dom = max(set(c["purpose"] for c in mem),
                  key=lambda p: sum(1 for c in mem if c["purpose"] == p))
        x = X0 + j * (bw + 20)
        P.append(_r(x, yb, bw, 4, tcol[tp]))
        P.append(_t(x, yb + 8, bw, 18, tp, bold=True, anchor=None,
                    wrap=False))
        P.append(_t(x, yb + 26, bw, 16, "HOW IT EARNS", bold=True,
                    color=NAVY, anchor=None, wrap=False))
        P.append(_t(x, yb + 42, bw, 36,
                    f"{mem[0]['route']}; {mem[0]['price'].lower()} "
                    f"positioning.", anchor=None))
        P.append(_t(x, yb + 78, bw, 16, "DOMINANT STRATEGY", bold=True,
                    color=NAVY, anchor=None, wrap=False))
        P.append(_t(x, yb + 94, bw, 36,
                    f"{dom}; {fmt_pct(sum(c['share'] for c in mem) / 100)}"
                    f" combined share.", anchor=None))
    return (f"{MKT} Player Landscape by Business Type and Reach",
            {"primpages": [P]})


# ═══════════════════════ 5-7. Scorecards ═══════════════════════════════
def _scorecard(title, band, cols, reading):
    """The reference's Level 1/2/3 page: a navy band, criteria under each
    column heading, one row per company with ratings or short text."""
    ci = _ci()
    cos = ci["companies"]
    tcol = _type_colour(ci)
    P = [_r(X0, TOP, X1 - X0, 22, NAVY),
         _t(X0 + 8, TOP, 700, 22, band, bold=True, color=WHITE, anchor="m",
            wrap=False)]
    # The company column grows to the RIGHT when a name is longer than the
    # default 190px, so the name never wraps onto its tier line; the other
    # columns give up the width (capped, past which the name is set smaller).
    name_w = max((_name_px(c["name"]) + 16 for c in cos), default=190)
    co_w = min(300.0, max(190.0, name_w))
    cx0 = X0 + co_w
    # text columns get more room than rating columns, the ownership column
    # the most (a roster may describe each company in a phrase)
    wts = [2.25 if name.startswith("Parent")
           else 1.0 if name in ("People", "Partnerships")    # n/a for now
           else 1.2 if kind == "text" else 1.0
           for name, _c, kind, _f in cols]
    unit = (X1 - cx0) / sum(wts)
    xs = [cx0 + unit * sum(wts[:j]) for j in range(len(cols))]
    hy = TOP + 26
    for j, (name, crit, _kind, _f) in enumerate(cols):
        x, cw = xs[j], unit * wts[j]
        P.append(_t(x + 4, hy, cw - 8, 36, name, bold=True, color=NAVY,
                    anchor="b"))
        P.append(_t(x + 4, hy + 38, cw - 8, 54, crit, color=GREY))
    P.append(_hl(X0, X1, hy + 96, NAVY, 1))
    ry, rh = hy + 98, 36
    for i, c in enumerate(cos):
        y0 = ry + i * rh
        if i:
            P.append(_hl(X0, X1, y0, ROW_RULE))
        P += _company(X0, y0, rh, c, tcol[_gk(c)], w=co_w)
        for j, (_n, _c, kind, f) in enumerate(cols):
            x, cw = xs[j], unit * wts[j]
            v = f(c)
            if kind == "harvey":
                P += _harvey(x + cw / 2, y0 + rh / 2, v)
            elif v in (None, "n/a"):
                P.append(_t(x, y0, cw, rh, "n/a", color=GREY, align="c",
                            anchor="m", wrap=False))
            else:
                P.append(_t(x + 4, y0, cw - 8, rh, v, anchor="m"))
    ye = ry + len(cos) * rh
    P.append(_hl(X0, X1, ye, ROW_RULE))
    P += _legend(X0, ye + 6)
    P += _reading(ye + 30, reading)
    return (f"{MKT} {title}", {"primpages": [P]})


def ci_corporate():
    ci = _ci()
    cos = ci["companies"]
    n_glob = sum(1 for c in cos if c["reach"] == "Global")
    return _scorecard(
        "Corporate Scorecard: Who They Are",
        "LEVEL 1, CORPORATE: WHO THEY ARE", [
            ("Parent and ownership", "From the company roster", "text", lambda c: c["type"]),
            ("Purpose (strategy)", "Dominant active move",
             "text", lambda c: c["purpose"]),
            ("Portfolio breadth", "Sub-segments: 6+ 3, "
             "4+ 2, 2+ 1", "harvey", lambda c: c["s_portfolio"]),
            ("Presence", "Global 3, regional 2, local 1",
             "harvey", lambda c: c["s_presence"]),
            ("Power (scale)", "Tier 1 = 3, 2 = 2, 3 = 1",
             "harvey", lambda c: c["s_power"]),
            ("Performance", "Share 10%+ = 3, "
             "5-10% = 2, else 1", "harvey", lambda c: c["s_performance"]),
            ("People", "Talent; not assessed",
             "text", lambda c: None),
            ("Partnerships", "Named alliances",
             "text", lambda c: None),
        ],
        f"{n_glob} of {len(cos)} named players compete globally; scale and "
        f"share follow the tier structure.")


def ci_offering():
    ci = _ci()
    patents = "In full study" if X.covered(["patent"], ci) else None
    return _scorecard(
        "Offering Scorecard: What They Sell",
        "LEVEL 2, OFFERING: WHAT THEY SELL", [
            ("Product", "Of the top three types offered",
             "harvey", lambda c: c["s_product"]),
            ("Platform (technology)", "R&D and development: "
             "both 3, one 2", "harvey", lambda c: c["s_platform"]),
            ("Price and model", "Positioning by tier", "text",
             lambda c: c["price"]),
            ("Positioning (verticals)", "End uses served", "text",
             lambda c: ", ".join(_short(v) for v in c["verticals"][:2])
             or None),
            ("Proof", "Named customer references", "text",
             lambda c: None),
            ("Patents and IP", "Where the study covers patents",
             "text", lambda c: patents),
        ],
        "Product breadth and technology strength do not always go "
        "together; the end uses each player serves show where it competes "
        "head to head.  Proof needs customer research.")


def ci_execution():
    return _scorecard(
        "Execution Scorecard: How They Win",
        "LEVEL 3, EXECUTION: HOW THEY WIN", [
            ("Place (channel)", "Route to market by tier", "text",
             lambda c: c["route"]),
            ("Promotion", "Go-to-market; not assessed",
             "text", lambda c: None),
            ("Pace (momentum)", "Active moves: 5+ = 3, 3-4 = 2",
             "harvey", lambda c: c["s_pace"]),
            ("Pipeline", "Launch or development in motion",
             "text", lambda c: c["pipeline"]),
            ("Policy exposure", "Regulation in its favour or against",
             "text", lambda c: None),
            ("Prospects", "Gaining 3, holding 2, at risk 1",
             "harvey", lambda c: c["s_prospects"]),
        ],
        "Momentum sits with the players running the most strategic moves; "
        "prospects combine breadth, reach, scale and pace.  Promotion and "
        "policy exposure need market and regulatory research.")


# ═══════════════════════ 8. Success factors ════════════════════════════
def ci_success():
    ci = _ci()
    cos = sorted(ci["companies"], key=lambda c: c["composite"], reverse=True)
    tcol = _type_colour(ci)
    f = ci["factors"]
    P = []
    cx0, comp_w = X0 + 200, 170
    cw = (X1 - comp_w - cx0) / len(f)
    for j, (name, w, _k, desc) in enumerate(f):
        x = cx0 + j * cw
        P.append(_t(x + 4, TOP, cw - 8, 18, f"{j + 1}. {name}", bold=True,
                    color=NAVY, anchor=None, wrap=False))
        P.append(_t(x + 4, TOP + 17, cw - 8, 20, f"Weight {int(w * 100)}%",
                    bold=True, wrap=False))
        P.append(_t(x + 4, TOP + 34, cw - 8, 36, desc, color=GREY,
                    anchor=None))
    xc = X1 - comp_w
    P.append(_t(xc + 4, TOP, comp_w, 18, "Composite", bold=True, color=NAVY,
                anchor=None, wrap=False))
    P.append(_t(xc + 4, TOP + 18, comp_w, 16, "weighted 0 to 3",
                anchor=None, wrap=False))
    P.append(_hl(X0, X1, TOP + 74, NAVY, 1))
    ry, rh = TOP + 80, 32
    for i, c in enumerate(cos):
        y0 = ry + i * rh
        if i:
            P.append(_hl(X0, X1, y0, ROW_RULE))
        col = tcol[_gk(c)]
        P.append(_r(X0, y0 + 4, 4, rh - 8, col))
        P.append(_t(X0 + 10, y0, 190, rh, c["name"], bold=True, anchor="m",
                    wrap=False))
        for j, (_n, _w, k, _d) in enumerate(f):
            P += _harvey(cx0 + j * cw + cw / 2, y0 + rh / 2, c[k])
        bw = (comp_w - 50) * c["composite"] / 3.0
        P.append(_r(xc + 4, y0 + 9, bw, rh - 18, NAVY))
        P.append(_t(xc + 8 + bw, y0, 44, rh, f"{c['composite']:.1f}",
                    bold=True, anchor="m", wrap=False))
    ye = ry + len(cos) * rh
    P.append(_hl(X0, X1, ye, ROW_RULE))
    P += _legend(X0, ye + 6)
    P.append(_t(X0, ye + 30, 300, 16, "STRATEGIC GROUPS", bold=True,
                color=NAVY, anchor=None, wrap=False))
    groups = {}
    for c in cos:
        groups.setdefault(_gk(c), []).append(c)
    gl = list(groups.items())[:5]
    gw = (X1 - X0 - (len(gl) - 1) * 20) / max(1, len(gl))
    for j, (tp, mem) in enumerate(gl):
        x = X0 + j * (gw + 20)
        best = max(f, key=lambda q: sum(c[q[2]] for c in mem))
        worst = min(f, key=lambda q: sum(c[q[2]] for c in mem))
        P.append(_r(x, ye + 50, gw, 3, NAVY))
        P.append(_t(x, ye + 56, gw, 18, tp, bold=True, anchor=None,
                    wrap=False))
        P.append(_t(x, ye + 74, gw, 34, ", ".join(c["name"] for c in mem),
                    color=GREY, anchor=None))
        P.append(_t(x, ye + 112, gw, 36,
                    f"Wins on {best[0].lower()}; weakest on "
                    f"{worst[0].lower()}.", anchor=None))
    return (f"{MKT} Success Factors and Composite Ranking",
            {"primpages": [P]})


# ═══════════════════════ 9. Activity tracker ═══════════════════════════
def _tracked(ci):
    """One move per player, the most telling one."""
    order = {"MA": 0, "PR": 1, "RD": 2, "EX": 3, "RG": 4, "CP": 5}
    best = {}
    for a in sorted(ci["activity"], key=lambda a: order.get(a["code"], 9)):
        best.setdefault(a["company"], a)
    return [best[c["name"]] for c in ci["companies"] if c["name"] in best]


def ci_activity():
    ci = _ci()
    tcol = _type_colour(ci)
    ctype = {c["name"]: _gk(c) for c in ci["companies"]}
    rows = _tracked(ci)
    cols = [("COMPANY", 140), ("DEVELOPMENT (WHAT IS HAPPENING)", 250),
            ("DATE", 60), ("TYPE", 50), ("SOURCE DEPT", 130),
            ("EVIDENCE", 100), ("SIGNAL (WHAT IT INDICATES)", 230),
            ("WHO SHOULD ACT", 0)]
    used = sum(w for _n, w in cols)
    cols[-1] = (cols[-1][0], X1 - X0 - used)
    P = []
    x = X0
    xs = []
    for name, w in cols:
        xs.append(x)
        P.append(_t(x + 2, TOP, w - 4, 18, name, bold=True, color=NAVY,
                    anchor=None, wrap=False))
        x += w
    P.append(_hl(X0, X1, TOP + 20, NAVY, 1.5))
    rh = 52
    for i, a in enumerate(rows):
        y0 = TOP + 22 + i * rh
        if i:
            P.append(_hl(X0, X1, y0, ROW_RULE))
        P.append(_r(X0, y0 + 4, 4, rh - 8, tcol[ctype[a["company"]]]))
        P.append(_t(X0 + 8, y0 + 3, 130, rh - 6, a["company"], bold=True,
                    anchor=None))
        P.append(_t(xs[1] + 2, y0 + 3, cols[1][1] - 6, rh - 6, a["what"],
                    anchor=None))
        P.append(_t(xs[2] + 2, y0 + 3, cols[2][1] - 4, 18, "n/a",
                    color=GREY, anchor=None, wrap=False))
        P.append(_r(xs[3] + 4, y0 + 6, 30, 18, BADGE.get(a["code"], NAVY)))
        P.append(_t(xs[3] + 4, y0 + 6, 30, 18, a["code"], bold=True,
                    color=WHITE, align="c", anchor="m", wrap=False))
        P.append(_t(xs[4] + 2, y0 + 3, cols[4][1] - 6, rh - 6, a["dept"],
                    anchor=None))
        P.append(_t(xs[5] + 2, y0 + 3, cols[5][1] - 6, rh - 6,
                    "FMI company profiling", color=GREY, anchor=None))
        P.append(_t(xs[6] + 2, y0 + 3, cols[6][1] - 6, rh - 6, a["signal"],
                    anchor=None))
        P.append(_t(xs[7] + 2, y0 + 3, cols[7][1] - 4, rh - 6, a["act"],
                    anchor=None))
    ye = TOP + 22 + len(rows) * rh
    P.append(_hl(X0, X1, ye, ROW_RULE))
    xx = X0
    for code, name in X.TYPE_NAMES.items():
        P.append(_r(xx, ye + 7, 30, 16, BADGE.get(code, NAVY)))
        P.append(_t(xx, ye + 7, 30, 16, code, bold=True, color=WHITE,
                    align="c", anchor="m", wrap=False))
        P.append(_t(xx + 34, ye + 5, 170, 20, name, anchor=None,
                    wrap=False))
        xx += 56 + len(name) * 7.0
    P.append(_t(X0, ye + 30, X1 - X0, 18, "Dates, deal values and evidence "
                "sources populate from the research feed; until then each "
                "row reads the player's most telling active strategy.",
                color=GREY, anchor=None, wrap=False))
    return (f"{MKT} Strategic Activity Tracker", {"primpages": [P]})


# ═══════════════════════ 10. Activity by function (heatmaps) ═══════════
def _heat(x, y, title, cos, cols, count, w):
    """A count heatmap: company rows, short column heads, shaded cells."""
    P = [_t(x, y, w, 18, title, bold=True, color=NAVY, wrap=False)]
    nw = 0                      # the company names stand once, at the left
    cw = (w - nw) / max(1, len(cols))
    for j, (head, _key) in enumerate(cols):
        P.append(_t(x + nw + j * cw, y + 18, cw, 36, head, bold=True,
                    color=INK, align="c", anchor="b"))
    for i, co in enumerate(cos):
        yy = y + 58 + i * 26
        for j, (_h, key) in enumerate(cols):
            n = count(co, key)
            fill = HEAT[min(n, 3) - 1] if n else WHITE
            P.append(_r(x + nw + j * cw + 1, yy + 1, cw - 2, 24, fill,
                        RULE, 0.5))
            if n:
                P.append(_t(x + nw + j * cw, yy + 1, cw, 24, str(n),
                            bold=True, color=NAVY if n < 3 else WHITE,
                            align="c", anchor="m", wrap=False))
    return P


def _abbr(region):
    """A short region head for a narrow heatmap column."""
    known = {"north america": "N. America", "latin america": "LATAM",
             "europe": "Europe", "east asia": "E. Asia",
             "south asia & pacific": "S. Asia & Pac.",
             "south asia and pacific": "S. Asia & Pac.",
             "middle east & africa": "MEA", "middle east and africa": "MEA",
             "eastern europe": "E. Europe", "western europe": "W. Europe"}
    return known.get(region.lower(), region)


def ci_functions():
    ci = _ci()
    acts = ci["activity"]
    cos = [c["name"] for c in ci["companies"]]
    pres = {c["name"]: c["presence"] for c in ci["companies"]}
    codes = [k for k in X.TYPE_NAMES if any(a["code"] == k for a in acts)]
    depts = [d for d in X.DEPARTMENTS if any(a["dept"] == d for a in acts)]
    regs = list(dict.fromkeys(r for c in ci["companies"]
                              for r in c["presence"]))[:6]
    nw = 160
    w = (X1 - X0 - nw - 2 * 20) / 3
    P = []
    tcol = _type_colour(ci)
    for i, c in enumerate(ci["companies"]):
        P.append(_r(X0, TOP + 62 + i * 26, 4, 18,
                    tcol[_gk(c)]))
        P.append(_t(X0 + 10, TOP + 58 + i * 26, nw - 12, 26, c["name"],
                    bold=True, anchor="m", wrap=False))
    x0 = X0 + nw
    P += _heat(x0, TOP, "Moves by type (count)", cos,
               [(k, k) for k in codes],
               lambda co, k: sum(1 for a in acts if a["company"] == co
                                 and a["code"] == k), w)
    short = {"Corporate development": "Corp. dev.",
             "Legal and regulatory": "Legal & reg.",
             "Product and R&D": "Product & R&D"}
    P += _heat(x0 + w + 20, TOP, "By department", cos,
               [(short.get(d, d), d) for d in depts],
               lambda co, d: sum(1 for a in acts if a["company"] == co
                                 and a["dept"] == d), w)
    P += _heat(x0 + 2 * (w + 20), TOP, "By region of activity", cos,
               [(_abbr(r), r) for r in regs],
               lambda co, r: (1 + (pres[co][0] == r)) if r in pres[co]
               else 0, w)
    ye = TOP + 58 + len(cos) * 26 + 6
    xx = X0
    for k, lab in enumerate(("1 move", "2 moves", "3+ moves")):
        P.append(_r(xx, ye + 4, 22, 13, HEAT[k]))
        P.append(_t(xx + 26, ye, 110, 20, lab, anchor="m", wrap=False))
        xx += 110
    P.append(_t(xx + 20, ye, X1 - xx - 20, 20, "Type codes: " + ", ".join(
        f"{k} {v}" for k, v in X.TYPE_NAMES.items()), color=GREY,
        anchor="m", wrap=False))
    yd = ye + 28
    P.append(_hl(X0, X1, yd, NAVY, 1.5))
    P.append(_t(X0, yd + 6, 600, 18, "WHAT IT MEANS FOR EACH OF YOUR "
                "DEPARTMENTS", bold=True, color=NAVY, anchor=None,
                wrap=False))

    def n_of(code):
        return sum(1 for a in acts if a["code"] == code)
    n_acq = len({a["company"] for a in acts if a["code"] == "MA"})
    meaning = [
        ("Sales", f"{n_of('EX')} expansion moves target new regions and "
                  f"volume tiers: defend key accounts where rivals expand."),
        ("Marketing", "Certification and compliance are being used as trust "
                      "signals; make your own credentials visible."),
        ("Product and R&D", f"{n_of('PR') + n_of('RD')} product and R&D moves: "
                            f"benchmark launches and the technology roadmap."),
        ("Corporate development", f"{n_acq} players are acquiring; the "
                                  f"independent target list is shrinking."),
        ("Operations", "Expansion moves add capacity in rivals' home "
                       "regions; review supply and lead times."),
        ("Legal and regulatory", f"{n_of('RG')} players lean on certification; "
                                 f"audit your own compliance position."),
    ]
    bw = (X1 - X0 - 5 * 12) / 6
    for j, (head, body) in enumerate(meaning):
        x = X0 + j * (bw + 12)
        P.append(_r(x, yd + 30, bw, 24, NAVY))
        P.append(_t(x + 8, yd + 30, bw - 12, 24, head, bold=True,
                    color=WHITE, anchor="m", wrap=False))
        P.append(_t(x, yd + 58, bw, 90, body))
    return (f"{MKT} Strategic Activity by Function", {"primpages": [P]})


PAGES = (
    ("Competitive Landscape at a Glance", ci_glance),
    ("Segment Leadership Map", ci_leadership),
    ("Value Chain Positioning", ci_value_chain),
    ("Player Landscape", ci_players),
    ("Corporate Scorecard", ci_corporate),
    ("Offering Scorecard", ci_offering),
    ("Execution Scorecard", ci_execution),
    ("Success Factors", ci_success),
    ("Strategic Activity Tracker", ci_activity),
    ("Strategic Activity by Function", ci_functions),
)
