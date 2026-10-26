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


# ═══════════════ 1. Dynamics by geography (matrix table) ═════════════
def dynamics_geography(region: str):
    """Reference slide 2 - a per-geography growth-driver / restraint /
    opportunity matrix with a KEY TAKEAWAY call-out underneath."""
    m = D.geography_matrix(region)
    kind = m["child_kind"].title()
    df = pd.DataFrame(
        [[r["drivers"], r["restraints"], r["opportunities"]]
         for r in m["rows"]],
        index=[r["geo"] for r in m["rows"]],
        columns=["Growth Drivers", "Restraints", "Priority Opportunities"])
    title = (f"{region} {MKT} Market Dynamics by {kind}")
    return title, {"rows": [
        L.frow(L.tbl(df, index_name=kind,
                     title=f"Demand Dynamics by {kind}  ·  ranked by "
                           f"{END_YEAR} value")),
        L.frow(L.summary("KEY TAKEAWAY", [m["takeaway"]])),
    ]}


# ═══════════════ 2-3. Global driving factors (radar) ═════════════════
def _driving_factors(side: str, region: str, part: str):
    scored = D.driving_factors(side, region)
    fig = figures.radar([n for n, _ in scored], [v for _, v in scored],
                        scale_max=D.SCALE_MAX, markers=True,
                        scale_labels=True, gridshape="linear",
                        show_shares=True)
    label = "Economic and Supply-side Factors" if side == "supply" \
        else "Demand-side Factors"
    blocks = [
        _pos(L.fig(fig, title=f"{label}  ·  impact score (0–{int(D.SCALE_MAX)})"),
             15, 92, 672, 606),
        _pos(L.summary(label, D.factor_narrative(side, region)),
             699, 92, 566, 606),
    ]
    return (f"{region} {MKT} Global Driving Factors ({part})",
            {"positioned": blocks})


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
    items = D.items(kind, region)
    labels = D.period_labels()
    prims = []

    # ── tab strip (the reference's vertical rail, laid out horizontally) ──
    for i, k in enumerate(D.KIND_ORDER):
        x = _X0 + i * (_TAB_W + _TAB_GAP)
        on = k == kind
        a, _tint = D.KIND_THEME[k]
        prims.append(_r(x, _TOP, _TAB_W, _TAB_H,
                        fill=a if on else "#ECEFF1",
                        line=a if on else "#D5DADE", lw=1, r=4))
        prims.append(_t(x, _TOP, _TAB_W, _TAB_H, D.KIND_LABEL[k], size=11,
                        bold=True, color="#FFFFFF" if on else "#8A9096",
                        align="c"))

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
        prims.append(_t(_IMP_X, y0, _IMP_W, _GRID_CAP_H,
                        "Likelihood of impact on timeline", size=11,
                        bold=True, color=MUTED, wrap=True, lh=15))
        for j, (lab, lvl) in enumerate(zip(labels, it["impact"])):
            cx = _IMP_X + (j % 2) * (_CELL_W + _CELL_GAP)
            cy = y0 + (_GRID_R1 if j < 2 else _GRID_R2)
            prims.append(_t(cx, cy, _CELL_W, _GRID_LAB_H, lab, size=11,
                            bold=True, color=MUTED, align="c"))
            prims.append(_r(cx, cy + _GRID_LAB_H + 1, _CELL_W, _CELL_H,
                            fill=D.IMPACT_FILL.get(lvl, D.IMPACT_OFF),
                            line=NEUTRAL, lw=1.25))
            # name the level inside the cell so the grid reads without a key
            prims.append(_t(cx, cy + _GRID_LAB_H + 1, _CELL_W, _CELL_H, lvl,
                            size=11, bold=True,
                            color=D.IMPACT_TEXT.get(lvl, MUTED), align="c"))
        if i < len(items) - 1:
            prims.append(_ln(_X0, y0 + ih + 4, _X1, y0 + ih + 4,
                             color="#E4EAEF", width=1))
        y0 += ih + _ITEM_GAP

    title = f"{region} {MKT} {D.KIND_LABEL[kind]}"
    return title, {"primitives": prims}


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
    rows = D.milestones(region)
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
    forces = D.five_forces(region)
    n = max(1, len(forces))
    colw = (_X1 - _X0) / n
    prims = [_t(_X0, 88, _X1 - _X0, 36, D.forces_subtitle(region), size=11,
                color=MUTED, wrap=True, lh=16)]

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
