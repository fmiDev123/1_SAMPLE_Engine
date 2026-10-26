"""Page templates for the '- Market Dynamics' report.

Modelled on the reference deck ``Dynamics.pptx``: a geography dynamics matrix,
two driving-factor radar pages (supply / demand side), the Market Drivers,
Restraints and Opportunities write-up pages with their impact-on-timeline
grids, a recent-developments table and a Porter's Five Forces page.

Every value, label and sentence comes from :mod:`dynamics`, which derives them
from the loaded workbook - these templates hold geometry only.  The bespoke
pages return a ``{"primitives": [...]}`` spec so the Dash canvas
(``components.draw_primitives``) and the native PPT slide
(``export_native._draw_primitives``) draw from one geometry source and match
exactly.
"""
from __future__ import annotations

import pandas as pd

from . import dynamics as D
from . import figures, layout as L
from .config import (DYN_HEAD_BG, END_YEAR, EST_YEAR, MUTED, NEUTRAL,
                     PANEL_EDGE, RULE_TINT, TEXT, VC_NAVY, ZEBRA)
from .data import MARKET_LABEL as MKT

# ── canvas constants (1280x720; header band ends at 84, source line at 708) ──
_X0, _X1 = 15, 1265
_TOP, _BOT = 92, 700
_NAVY = VC_NAVY
_HEAD_BG = DYN_HEAD_BG
_ZEBRA = ZEBRA
_RULE = RULE_TINT


def _pos(block, x, y, w, h):
    """Attach explicit x/y/w/h (px) to a layout block for a 'positioned' spec."""
    return {**block, "x": int(x), "y": int(y), "w": int(w), "h": int(h)}


# ── primitive helpers ───────────────────────────────────────────────
def _t(x, y, w, h, s, size=10, bold=False, color=None, align="l",
       wrap=False, lh=None):
    # `color=TEXT` as a DEFAULT bound the colour at import time, which
    # `theming.apply()` cannot rebind - so every dynamics page kept the
    # Fact.MR ink even under the FMI template.  Resolved per call instead.
    p = {"t": "text", "x": x, "y": y, "w": w, "h": h, "s": str(s),
         "size": size, "bold": bold,
         "color": TEXT if color is None else color, "align": align}
    if wrap:
        p["wrap"] = True
        p["lh"] = lh or (size * 1.42)
    elif lh:
        p["lh"] = lh
    return p


def _r(x, y, w, h, fill="#FFFFFF", line=None, lw=1, dash=None, r=None):
    p = {"t": "rrect" if r else "rect", "x": x, "y": y, "w": w, "h": h,
         "fill": fill, "line": line, "lw": lw}
    if dash:
        p["dash"] = dash
    if r:
        p["r"] = r
    return p


def _ln(x1, y1, x2, y2, color=_RULE, width=1, dash=None):
    p = {"t": "line", "x1": x1, "y1": y1, "x2": x2, "y2": y2,
         "color": color, "width": width}
    if dash:
        p["dash"] = dash
    return p


def _circle(cx, cy, d, color, edge):
    return {"t": "dot", "cx": cx, "cy": cy, "d": d, "color": color,
            "edge": edge, "lw": 1.5}


def _poly(pts, color, width=3):
    return {"t": "poly", "pts": pts, "color": color, "width": width}


def _shape(pts, fill):
    return {"t": "shape", "pts": pts, "fill": fill}


def _icon(kind, cx, cy, ink, bg):
    """A pictogram centred on (cx, cy), built only from drawing primitives so
    the Dash canvas and the PPT slide produce the same mark (no images, no
    per-surface icon code).  ``bg`` is the disc colour, used to knock out the
    rings of the target.  Mirrors the reference deck's five icons."""
    def P(*pts):                                   # local -> absolute coords
        return [(cx + dx, cy + dy) for dx, dy in pts]

    if kind == "dollar":                           # currency glyph
        return [_t(cx - 30, cy - 22, 60, 44, "$", size=34, bold=True,
                   color=ink, align="c")]
    if kind == "cart":                             # handle + basket + wheels
        return [
            _poly(P((-25, -17), (-17, -17), (-9, 6)), ink, 3.4),
            _shape(P((-11, -9), (25, -9), (20, 8), (-6, 8)), ink),
            {"t": "dot", "cx": cx - 3, "cy": cy + 17, "d": 8, "color": ink,
             "edge": ink, "lw": 1},
            {"t": "dot", "cx": cx + 16, "cy": cy + 17, "d": 8, "color": ink,
             "edge": ink, "lw": 1},
        ]
    if kind == "target":                           # bullseye + arrow
        rings = []
        for d, col in ((46, ink), (32, bg), (18, ink), (7, bg)):
            rings.append({"t": "dot", "cx": cx, "cy": cy, "d": d,
                          "color": col, "edge": col, "lw": 1})
        rings.append(_poly(P((4, -4), (26, -26)), ink, 3))
        rings.append(_shape(P((18, -30), (30, -30), (30, -18)), ink))
        return rings
    if kind == "cubes":                            # stacked boxes
        return [_r(cx - 10, cy - 25, 20, 20, fill=bg, line=ink, lw=2.2, r=3),
                _r(cx - 24, cy + 1, 20, 20, fill=bg, line=ink, lw=2.2, r=3),
                _r(cx + 4, cy + 1, 20, 20, fill=bg, line=ink, lw=2.2, r=3)]
    if kind == "people":                           # a group of three
        out = []
        for dx, hd, top, bot in ((-20, 12, -6, 17), (20, 12, -6, 17),
                                 (0, 16, -12, 19)):
            out.append({"t": "dot", "cx": cx + dx, "cy": cy + top - hd / 2,
                        "d": hd, "color": ink, "edge": ink, "lw": 1})
            w = hd * 0.95
            out.append(_shape(P((dx - w, bot), (dx - w * 0.72, top + 3),
                                (dx + w * 0.72, top + 3), (dx + w, bot)), ink))
        return out
    return []


# A primitive's ``size`` is CSS px on the dashboard but POINTS in PowerPoint,
# and 1 slide px = 0.75 pt - so the same primitive renders ~1.35x wider in the
# deck.  Block heights are budgeted against the PPT rendering (the larger of
# the two); the dashboard simply gets a little extra breathing room.
# A wrapped primitive is drawn with 1.15 line spacing on PowerPoint's natural
# 1.2em line, i.e. size x 1.2 x 1.15 x 4/3 = ~1.84 slide px per point.  The
# old 2.25 over-budgeted every block by ~20%, which at the 11pt content
# standard (2026-09-11) pushed bullets off the page for no reason.
_PPT_FONT_K = 1.35                     # wrap width factor
_PPT_LINE_K = 1.86                     # line height as a multiple of `size`


def _text_h(s, w, size, lh):
    """Height ``s`` needs when wrapped into ``w`` px - budgeted for whichever
    surface renders it taller (PowerPoint)."""
    lines = L._wrap_count(str(s), w, size * _PPT_FONT_K)
    return int(lines * max(lh, size * _PPT_LINE_K) + 4)


def _tblock(x, y, w, avail_h, s, size=9.6, lh=14.6, color=None, bold=False,
            align="l", anchor="c"):
    """A wrapped text primitive sized to its own content and placed inside
    ``avail_h`` (``anchor`` 't' = top, 'c' = centred).

    Sizing the box to the text - rather than to the slot - is what keeps the
    Dash div (top-aligned) and the PPT text frame (middle-anchored) in the
    same place on both surfaces."""
    h = min(int(avail_h), _text_h(s, w, size, lh))
    dy = 0 if anchor == "t" else max(0, (avail_h - h) / 2)
    return _t(x, y + dy, w, h, s, size=size, bold=bold, color=color,
              align=align, wrap=True, lh=lh)


def _distribute(needs, avail, gap, minimum):
    """Share ``avail`` px between blocks in proportion to what each needs, so a
    short write-up does not leave a tall empty card behind it."""
    n = len(needs) or 1
    room = avail - (n - 1) * gap
    total = float(sum(needs)) or 1.0
    out = [max(minimum, int(room * v / total)) for v in needs]
    over = sum(out) - room
    while over > 0:                       # trim the tallest until it fits
        i = out.index(max(out))
        take = min(over, out[i] - minimum)
        if take <= 0:
            break
        out[i] -= take
        over -= take
    return out


def _mtitle(region, rest):
    """'<region> <market> <rest>' without the market word twice: the market
    label already ends in 'Market', so 'Market Restraints' reads
    'Global Paper Bottle Market Restraints' (client, 2026-10-01)."""
    if MKT.rstrip().lower().endswith("market") and \
            rest.lower().startswith("market "):
        rest = rest[len("market "):]
    return f"{region} {MKT} {rest}"


# ═══════════════ 1. Dynamics by geography (matrix table) ═════════════
_GEO_HEAD = "#DCE6F0"


def dynamics_geography(region: str):
    """Reference slide 2 - a per-geography growth-driver / restraint /
    opportunity matrix with a KEY TAKEAWAY call-out underneath.

    Drawn from primitives so every cell WRAPS and carries its phrase in full
    - a flow table keeps a cell on one line and cut it off with '...'
    (client, 2026-10-01).  Rows that do not fit run on to a continuation
    page; the takeaway closes the last one."""
    m = D.given_geography(region)
    if m is None:                         # nothing supplied: leave it blank
        kind = "Region" if region == "Global" else "Country"
        return (_mtitle(region, f"Market Dynamics by {kind}"),
                {"primitives": []})
    kind = m["child_kind"].title()
    heads = (kind, "Growth Drivers", "Restraints", "Priority Opportunities")
    widths = (160, 400, 300, _X1 - _X0 - 860)
    xs = [_X0]
    for w in widths[:-1]:
        xs.append(xs[-1] + w)
    keys = ("drivers_full", "restraints_full", "opportunities_full")
    pt, lh = 11, 16

    take = m["takeaway"]
    take_h = 30 + _text_h(take, _X1 - _X0 - 40, pt, lh) + 14

    def row_h(r):
        return max(40, max(_text_h(r[k], widths[j + 1] - 14, pt, lh)
                           for j, k in enumerate(keys)) + 12)

    pages, cur, y = [], [], 0
    head_top = _TOP + 26 + 28
    for r in m["rows"]:
        h = row_h(r)
        if cur and head_top + y + h > _BOT - 24:
            pages.append(cur)
            cur, y = [], 0
        cur.append((r, h))
        y += h
    pages.append(cur)
    used = sum(h for _r, h in pages[-1])
    if head_top + used + 30 + take_h > _BOT:
        pages.append([])                  # the takeaway on a page of its own

    out = []
    for pi, rows in enumerate(pages):
        P = []
        if rows:
            P.append(_t(_X0 + 4, _TOP, 900, 24,
                        f"Demand Dynamics by {kind}  ·  ranked by "
                        f"{END_YEAR} value", size=11.5, bold=True))
            hy = _TOP + 26
            P.append(_r(_X0, hy, _X1 - _X0, 28, fill=_GEO_HEAD, line=None))
            for x, w, hd in zip(xs, widths, heads):
                P.append(_t(x + 6, hy, w - 10, 28, hd, size=11, bold=True,
                            color=_NAVY))
            y = hy + 28
            for i, (r, h) in enumerate(rows):
                P.append(_r(_X0, y, _X1 - _X0, h,
                            fill=_ZEBRA if i % 2 else "#FFFFFF", line=None))
                P.append(_t(xs[0] + 6, y, widths[0] - 10, h, r["geo"],
                            size=11, bold=True))
                for j, k in enumerate(keys):
                    # the cell's full row height, so a padlock in a
                    # one-line cell never forces the text smaller
                    P.append(_t(xs[j + 1] + 6, y + 3, widths[j + 1] - 14,
                                h - 6, r[k], size=pt, wrap=True, lh=lh))
                y += h
            P.append(_ln(_X0, y, _X1, y, color=_RULE, width=1))
            P.append(_t(_X0 + 6, y + 2, 400, 20,
                        f"Source: FMI Analysis, {EST_YEAR}", size=11,
                        color=MUTED))
            y += 30
        else:
            y = _TOP
        if pi == len(pages) - 1:
            P.append(_r(_X0, y, _X1 - _X0, take_h, fill="#FFFFFF",
                        line=PANEL_EDGE, lw=1, r=6))
            P.append(_t(_X0 + 14, y + 4, 400, 24, "KEY TAKEAWAY", size=12,
                        bold=True, color=_NAVY))
            P.append(_ln(_X0 + 14, y + 30, _X1 - 14, y + 30, color=_NAVY,
                         width=1.25))
            P.append(_tblock(_X0 + 14, y + 36, _X1 - _X0 - 40, take_h - 40,
                             "▪  " + take, size=pt, lh=lh, anchor="t"))
        out.append(P)
    title = _mtitle(region, f"Market Dynamics by {kind}")
    if len(out) == 1:
        return title, {"primitives": out[0]}
    return title, {"primpages": out}


# ═══════════════ 2-3. Global driving factors (radar) ═════════════════
# The client's reference ('Economic and Supply-side Factors' / 'Demand-side
# Factors', 2026-10-01): a round web on shaded rings with a 0-100% ruler up
# the top spoke, a dot on the rim of every spoke, the profile as an open
# teal outline, and each spoke named with its share.  The findings sit
# beside it under a ruled heading, split from the web by a thin rule;
# supply-side has the web on the left, demand-side on the right.
_RADAR_STYLE = {
    "supply": {"fill": "#DCE7F3", "ring": "#B9CDE3", "dot": "#2F6DA8"},
    "demand": {"fill": "#E2F0CF", "ring": "#C5DFA4", "dot": "#8CC152"},
}
_RADAR_LINE = "#0F5F63"
_HEAD_RULE = "#1F6FB4"
_RADAR_R = 250


def _radar_prims(cx, cy, R, scored, style, side):
    """A HALF web (client reference, 2026-10-01): every spoke fans over one
    half-circle, top to bottom, on the side facing away from the findings
    (``side`` -1 = left, +1 = right); the centre sits on the dividing rule
    and the findings panel is laid over the other half of the rings."""
    import math
    n = len(scored)
    tot = sum(abs(v) for _n, v in scored) or 1.0
    P = []
    for k in range(10, 0, -1):            # shaded disc, one ring per 10%
        P.append({"t": "dot", "cx": cx, "cy": cy, "d": 2 * R * k / 10,
                  "color": style["fill"], "edge": style["ring"], "lw": 0.75})

    def at(i, rad):
        a = math.radians(-90.0 + i * 180.0 / max(1, n - 1))
        return (cx + side * rad * math.cos(a), cy + rad * math.sin(a),
                math.cos(a), math.sin(a))

    for i in range(n):                    # spokes
        ex, ey, _c, _s = at(i, R)
        P.append(_ln(cx, cy, ex, ey, color=style["ring"], width=0.75))
    pts = []
    for i, (_name, v) in enumerate(scored):
        x, y, _c, _s = at(i, R * max(0.0, min(1.0, v / D.SCALE_MAX)))
        pts.append((x, y))
    P.append(_poly(pts, _RADAR_LINE, 2.5))
    for i, (name, v) in enumerate(scored):
        ex, ey, _c, _s = at(i, R)
        P.append({"t": "dot", "cx": ex, "cy": ey, "d": 10,
                  "color": style["dot"], "edge": style["dot"], "lw": 1})
        lx, ly, c, sn = at(i, R + 14)
        s = f"{name}  {abs(v) / tot * 100:.1f}%"
        # roomy boxes: a padlock in place of the share must not force the
        # label smaller in the locked deck
        h = 26
        if c < 0.25:                      # top / bottom spoke: beside the
            lx = cx + side * 8            # rule, on the open side
        if side < 0:
            x0 = max(2.0, lx - 240)
            w, al = lx - x0, "r"
        else:
            x0 = lx
            w, al = min(240.0, _X1 - lx), "l"
        y0 = ly - h + 4 if sn < -0.8 else ly - 4 if sn > 0.8 else ly - h / 2
        P.append(_t(x0, y0, w, h, s, size=11, align=al))
    return P


def _scale_prims(cx, cy, R, side):
    """The 0-100% ruler up the top spoke, on the open side of the rule."""
    P = []
    for k in range(0, 11):
        x0 = cx - 50 if side < 0 else cx + 6
        p = _t(x0, cy - R * k / 10 - 8, 44, 16, f"{k * 10}%", size=11,
               color=MUTED, align="r" if side < 0 else "l")
        p["name"] = f"FMI Scale {k}"
        P.append(p)
    return P


def _factor_panel(x, w, label, bullets):
    P = [_t(x, _TOP + 6, w, 24, label, size=12, bold=True, color=_NAVY),
         _ln(x, _TOP + 34, x + w, _TOP + 34, color=_HEAD_RULE, width=2.5)]
    y = _TOP + 44
    for b in bullets:
        s = f"▪  {b}"
        h = _text_h(s, w, 11, 17)
        if y + h > _BOT:
            break
        # a little spare room, so a padlock never forces a line smaller
        P.append(_t(x, y, w, h + 8, s, size=11, wrap=True, lh=17))
        y += h + 6
    return P


def _driving_factors(side: str, region: str, part: str):
    given = D.given_factors(side, region)
    if given is None:                     # nothing supplied: leave it blank
        return (f"{region} {MKT} Global Driving Factors ({part})",
                {"primitives": []})
    scored, bullets = given
    label = "Economic and Supply-side Factors" if side == "supply" \
        else "Demand-side Factors"
    cy, R = (_TOP + _BOT) / 2 + 10, _RADAR_R
    st = _RADAR_STYLE[side]
    if side == "supply":                  # half web left, findings over the
        cx = 500                          # right half
        P = _radar_prims(cx, cy, R, scored, st, -1)
        P.append(_r(cx, _TOP - 6, 1280 - cx, _BOT - _TOP + 10,
                    fill="#FFFFFF", line=None))
        P += _scale_prims(cx, cy, R, -1)
        P.append(_ln(cx, _TOP, cx, _BOT, color="#D5DBE1", width=1))
        P += _factor_panel(cx + 22, _X1 - cx - 22, label, bullets)
    else:                                 # findings over the left half, half
        cx = 780                          # web right
        P = _radar_prims(cx, cy, R, scored, st, +1)
        P.append(_r(0, _TOP - 6, cx, _BOT - _TOP + 10, fill="#FFFFFF",
                    line=None))
        P += _scale_prims(cx, cy, R, +1)
        P.append(_ln(cx, _TOP, cx, _BOT, color="#D5DBE1", width=1))
        P += _factor_panel(_X0 + 4, cx - _X0 - 26, label, bullets)
    return (f"{region} {MKT} Global Driving Factors ({part})",
            {"primitives": P})


def driving_factors_supply(region: str):
    return _driving_factors("supply", region, "1/2")


def driving_factors_demand(region: str):
    return _driving_factors("demand", region, "2/2")


# ═════════ 4-6. Drivers / Restraints / Opportunities write-ups ═══════
_TAB_W, _TAB_H, _TAB_GAP = 214, 28, 9
_ITEM_GAP = 10
_IMP_X, _IMP_W = 1030, 235
_CELL_W, _CELL_H, _CELL_GAP = 104, 34, 10
# 11pt content standard (2026-09-11): the grid's caption takes two lines at
# 11pt, so the two cell rows start lower; the cells themselves are a little
# shorter, so a block's floor height barely moves.
_BODY_PT, _BODY_LH = 11, 17
_GRID_CAP_H = 40                       # "Likelihood of impact ..." caption
_GRID_LAB_H = 17                       # a period label over its cell
_GRID_R1 = _GRID_CAP_H + 2             # first cell row, from the block top
_GRID_R2 = _GRID_R1 + _GRID_LAB_H + _CELL_H + 6


def _kind_page(kind: str, region: str):
    """Reference slides 5-7 - a Drivers / Restraints / Opportunities tab strip
    over three write-up blocks, each with its own likelihood-of-impact grid."""
    accent, tint = D.KIND_THEME[kind]
    items = D.given_items(kind, region)
    if items is None:                     # nothing supplied: leave it blank
        return _mtitle(region, D.KIND_LABEL[kind]), {"primitives": []}
    labels = D.period_labels()
    prims = []

    # ── the page's own tab only (client reference 'RO', 2026-10-01) ──
    prims.append(_r(_X0, _TOP, _TAB_W, _TAB_H, fill=accent, line=accent,
                    lw=1, r=4))
    prims.append(_t(_X0, _TOP, _TAB_W, _TAB_H, D.KIND_LABEL[kind], size=11,
                    bold=True, color="#FFFFFF", align="c"))

    top = _TOP + _TAB_H + 12
    body_w = _IMP_X - _X0 - 18
    # the impact grid sets the floor height for a block (2 rows of cells)
    grid_h = _GRID_R2 + _GRID_LAB_H + _CELL_H + 6
    needs = [max(grid_h, 34 + _text_h(it["body"], body_w - 6, _BODY_PT,
                                      _BODY_LH) + 10)
             for it in items]
    heights = _distribute(needs, _BOT - top, _ITEM_GAP, grid_h)

    y0 = top
    for i, (it, ih) in enumerate(zip(items, heights)):
        # title bar (the 5px leading accent tab was removed 2026-09-07 - the
        # heading now sits on the tinted bar alone)
        prims.append(_r(_X0, y0, body_w, 28, fill=tint, line=None))
        prims.append(_t(_X0 + 14, y0, body_w - 24, 28, it["title"], size=11,
                        bold=True, color=_NAVY))
        # body copy
        prims.append(_tblock(_X0 + 2, y0 + 34, body_w - 6, ih - 40,
                             it["body"], size=_BODY_PT, lh=_BODY_LH,
                             anchor="t"))
        # ── likelihood-of-impact grid ──
        if it["impact"]:
            prims.append(_t(_IMP_X, y0, _IMP_W, _GRID_CAP_H,
                            "Likelihood of impact on timeline", size=11,
                            bold=True, color=MUTED, wrap=True, lh=15))
        for j, (lab, lvl) in enumerate(zip(labels, it["impact"])):
            cx = _IMP_X + (j % 2) * (_CELL_W + _CELL_GAP)
            cy = y0 + (_GRID_R1 if j < 2 else _GRID_R2)
            prims.append(_t(cx, cy, _CELL_W, _GRID_LAB_H, lab, size=11,
                            bold=True, color=MUTED, align="c"))
            # only the period where the item bites hardest is shaded, in its
            # level's colour; the others are outlined (reference 'RO')
            prims.append(_r(cx, cy + _GRID_LAB_H + 1, _CELL_W, _CELL_H,
                            fill=(_PEAK_FILL.get(lvl, D.IMPACT_OFF)
                                  if j == it.get("peak", -1) else "#FFFFFF"),
                            line=_CELL_EDGE, lw=1.25))
            # name the level inside the cell so the grid reads without a key
            prims.append(_t(cx, cy + _GRID_LAB_H + 1, _CELL_W, _CELL_H, lvl,
                            size=11, bold=True,
                            color=D.IMPACT_TEXT.get(lvl, MUTED), align="c"))
        if i < len(items) - 1:
            prims.append(_ln(_X0, y0 + ih + 4, _X1, y0 + ih + 4,
                             color="#E4EAEF", width=1))
        y0 += ih + _ITEM_GAP

    title = _mtitle(region, D.KIND_LABEL[kind])
    return title, {"primitives": prims}


# the highlighted cell's fill by level, and the cell outline (reference 'RO')
_PEAK_FILL = {"Low": "#FBF0C4", "Medium": "#F2C94C", "High": "#C99A0A"}
_CELL_EDGE = "#A8740F"


def market_drivers(region: str):
    return _kind_page("drivers", region)


def market_restraints(region: str):
    return _kind_page("restraints", region)


def market_opportunities(region: str):
    return _kind_page("opportunities", region)


# ═══════════════ 7. Recent developments / milestones ═════════════════
def recent_developments(region: str):
    """Reference slide 8 - a Year | Recent Development table, built from
    primitives so the long development copy wraps identically on both
    surfaces (the flow table renderer keeps cells on one line)."""
    rows = D.given_milestones(region)
    if rows is None:                      # nothing supplied: leave it blank
        return (f"{region} {MKT} Recent Developments & Milestones",
                {"primitives": []})
    prims = []
    hy, hh = _TOP, 32
    prims.append(_r(_X0, hy, _X1 - _X0, hh, fill=_HEAD_BG, line=None))
    prims.append(_t(_X0, hy, 150, hh, "Year", size=11.5, bold=True,
                    color="#FFFFFF", align="c"))
    prims.append(_t(_X0 + 165, hy, _X1 - _X0 - 175, hh, "Recent Development",
                    size=11.5, bold=True, color="#FFFFFF", align="c"))

    top = hy + hh
    tw = _X1 - _X0 - 178
    needs = [_text_h(t, tw, 11, 16) + 22 for _, t in rows]
    heights = _distribute(needs, _BOT - top, 0, 56)

    y = top
    for i, ((year, text), rh) in enumerate(zip(rows, heights)):
        prims.append(_r(_X0, y, _X1 - _X0, rh,
                        fill=_ZEBRA if i % 2 else "#FFFFFF",
                        line=_RULE, lw=1))
        prims.append(_t(_X0, y, 150, rh, year, size=12.5, bold=True,
                        color=_HEAD_BG, align="c"))
        prims.append(_ln(_X0 + 150, y, _X0 + 150, y + rh, color=_RULE,
                         width=1))
        prims.append(_tblock(_X0 + 165, y + 8, tw, rh - 16, text, size=11,
                             lh=16))
        y += rh

    title = f"{region} {MKT} Recent Developments & Milestones"
    return title, {"primitives": prims}


# ═══════════════════ 8. Porter's Five Forces ═════════════════════════
def five_forces(region: str):
    """Reference slide 9 - five coloured force discs (each showing the score
    derived from the workbook) with the rating and supporting evidence."""
    forces = D.given_forces(region)
    if forces is None:                    # nothing supplied: leave it blank
        return (f"Porter's Five Forces Analysis {region} {MKT}",
                {"primitives": []})
    n = max(1, len(forces))
    colw = (_X1 - _X0) / n
    prims = []

    # 11pt bullets (2026-09-11) need ~30% more height than the old 8.6pt, so
    # the discs are a little smaller and sit higher to leave them the room
    cy, dia = 184, 96
    for i, f in enumerate(forces):
        cx = _X0 + colw * i + colw / 2
        # disc + pictogram (the reference shows an icon here, never a score)
        prims.append(_circle(cx, cy, dia, f["color"], f["color"]))
        prims.extend(_icon(f["icon"], cx, cy, f["ink"], f["color"]))
        prims.append(_t(cx - colw / 2 + 8, cy + 56, colw - 16, 40, f["name"],
                        size=11.5, bold=True, color=TEXT, align="c",
                        wrap=True, lh=15))
        # bullets stack at their natural height (no equal-slot padding)
        by, bw = cy + 102, colw - 22
        for b in f["bullets"]:
            s = f"▸  {b}"
            bh = _text_h(s, bw, 11, 15.5)
            if by + bh > _BOT:
                break
            prims.append(_t(cx - colw / 2 + 11, by, bw, bh, s, size=11,
                            color=TEXT, wrap=True, lh=15.5))
            by += bh + 10

    title = f"Porter's Five Forces Analysis {region} {MKT}"
    return title, {"primitives": prims}


# -- Template switch support ------------------------------------------
def _retheme() -> None:
    """Re-derive what this module bakes from theme tokens (see theming.py)."""
    global _NAVY, _HEAD_BG, _ZEBRA, _RULE
    _NAVY, _HEAD_BG = VC_NAVY, DYN_HEAD_BG
    _ZEBRA, _RULE = ZEBRA, RULE_TINT
