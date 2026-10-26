"""Power BI-style building blocks: the 1280x720 canvas, the dark-blue
header band, absolutely positioned visual containers, KPI cards and
pivot-style tables - geometry and colours match the PBIX files."""
from __future__ import annotations

import re as _re

import pandas as pd
from dash import dcc, html

from .config import (ACCENT, ARROW_HIGH, ARROW_LOW, ARROW_MED, BAD,
                     BELOW_COLOR, BOX_COLORS, CANVAS_H, CANVAS_W, COPYRIGHT,
                     DECLINE_COLOR, FAINT, FLOW_LEAF_EDGE, FLOW_LEAF_TEXT,
                     FLOW_LINE, FONT_STACK, FOOTER_BAND, FOOTER_BAND_H,
                     FOOTER_MARGIN, FOOTER_RULE_Y, FOOTER_TEXT_COLOR,
                     FS_SOURCE, GAUGE_TRACK, GOOD, GRID, KPI_TINT,
                     HEADER_BG, HEADER_FILL, HEADER_LABEL, HEADER_RULE,
                     HEADER_RULE_SEGMENTS, HEADER_SHADOW, HEADER_STYLE,
                     MUTED, NOT_OK_COLOR,
                     OK_COLOR, OUTPACE_COLOR, PAGE_BG, PANEL_BG, PANEL_EDGE,
                     NAVY_DEEP, PANEL_RADIUS, PANEL_SHADOW,
                     RULE_H, RULE_Y, SIG_EDGE, SIG_HEAD_BG,
                     SIG_ZEBRA,
                     SOURCE_TEXT, STAR_COLORS, STAR_EMPTY, TABLE_HEAD_BG,
                     TABLE_HEAD_TEXT, TABLE_TOTAL_BG,
                     TEXT, TITLE_COLOR, VC_NAVY, VC_PURPLE, WARN_TEXT)

# ── Formatting helpers ───────────────────────────────────────────────
def fmt_num(x) -> str:
    if pd.isna(x):
        return "–"
    x = float(x)
    if abs(x) >= 1000:
        return f"{x:,.0f}"
    if abs(x) >= 10:
        return f"{x:,.1f}"
    return f"{x:,.2f}"

def fmt_pct(x, decimals=1) -> str:
    return "–" if pd.isna(x) else f"{x * 100:.{decimals}f}%"

def yoy_icon(v) -> str:
    if pd.isna(v):
        return ""
    return "▲" if v >= 0 else "▼"

# darkened palette gold - the raw NEUTRAL is too light for table text
_AMBER = BELOW_COLOR

def yoy_cell(v, benchmark=None):
    """YoY table cell: (text, colour) graded against the market's pace.

    ▼ red    - outright decline (YoY < 0)
    ▶ amber  - growing, but slower than the market that year
    ▲ green  - at or above the market's growth
    With no benchmark the rule degrades to the plain sign test.

    The colours are the CLASSIFICATION tokens, not GOOD/BAD.  The legend that
    sits in the panel title above these cells is split by `title_runs`, which
    reads the classification tokens - so drawing the cells from GOOD/BAD (blue
    and pink on Fact.MR) put the key and the cells it describes in different
    colours on the same page."""
    if pd.isna(v):
        return ("–", MUTED)
    if v < 0:
        return (f"▼ {v:.1f}%", DECLINE_COLOR)
    if benchmark is not None and not pd.isna(benchmark) and v < benchmark:
        return (f"▶ {v:.1f}%", _AMBER)
    return (f"▲ {v:.1f}%", OUTPACE_COLOR)

def pos_neg(text: str, value: float):
    """(text, colour) for a signed figure.  If the text carries a ▲/▼ mark it
    is a performance CLASSIFICATION and takes the classification colours -
    green up, red down - so a 'Direction' or 'Trend' cell matches the ▲/▼ in
    the YoY tables and the read-out bullets beside it.  A bare signed number
    (e.g. a BPS variance) keeps the GOOD/BAD chart tokens it always used."""
    if any(m in str(text) for m in _CLASS_MARK):
        return (text, OUTPACE_COLOR if value >= 0 else DECLINE_COLOR)
    return (text, GOOD if value >= 0 else BAD)

# arrow -> classification colour, shared by cells and legend keywords
_CLASS_MARK = {"▲": OUTPACE_COLOR, "▶": BELOW_COLOR, "▼": DECLINE_COLOR}

def title_runs(title):
    """Split a panel title into (text, colour) runs so its embedded YoY legend
    keywords carry the SAME colours as the table's cells - ▲ outpacing = green,
    ▶ below market pace = orange, ▼ decline = red - while the rest stays black.
    Titles without those markers return a single black run.  One helper, used by
    both the dashboard and the PowerPoint export."""
    if not title:
        return [("", TEXT)]
    positions = sorted((title.find(m), m) for m in _CLASS_MARK if m in title)
    if not positions:
        return [(title, TEXT)]
    runs = []
    if positions[0][0] > 0:
        runs.append((title[:positions[0][0]], TEXT))
    for k, (pos, m) in enumerate(positions):
        end = positions[k + 1][0] if k + 1 < len(positions) else len(title)
        runs.append((title[pos:end], _CLASS_MARK[m]))
    return runs

# ── Performance marks inside RUNNING TEXT ────────────────────────────
# `title_runs` colours a legend: everything from one mark to the next belongs
# to that mark ("▲ outpacing market  ▶ below market pace  ▼ decline").  That
# rule is wrong for a sentence, where the mark is one token inside a clause -
# "'PE-Based' gains the most share +236 BPS between 2026 and 2036 (green ▲ on
# the scale)." would colour " on the scale)." too.  So prose gets a TIGHT span
# instead: the mark, the colour word the sentence uses for it if there is one,
# and the figure it labels if one follows.
#
# A sentence that says "green ▲" or "red ▼" out loud is the reason this exists:
# the words have to be set in the colour they name, or the page contradicts
# itself.  The word wins when present, since it is what the author wrote; a
# bare mark falls back to its classification colour.
_WORD_MARK = {"green": "▲", "red": "▼", "amber": "▶", "orange": "▶"}

_MARK_SPAN = _re.compile(
    r"(?:(?P<word>green|red|amber|orange)\s+)?"      # optional colour word
    r"(?P<mark>[▲▶▼])"                              # the mark itself
    r"(?P<val>\s*[+\-−]?[\d,]+(?:\.\d+)?\s*%?)?",   # the figure it labels
    _re.IGNORECASE)

def mark_runs(text, default=None):
    """Split running text into (text, colour) runs, colouring every ▲/▶/▼ -
    and the colour word and figure attached to it - in its classification
    colour, leaving the rest in ``default``.  Text with no mark returns a
    single run, so callers can use it unconditionally.

    One helper, used by both surfaces: the dashboard's summary/commentary
    blocks and the PowerPoint export's native text boxes.  Marks reach a deck
    through prose as often as through a legend, and a black ▼ in a sentence
    about decline is the same defect as a black ▼ in a table."""
    default = TEXT if default is None else default
    s = "" if text is None else str(text)
    if not any(m in s for m in _CLASS_MARK):
        return [(s, default)]
    runs, i = [], 0
    for m in _MARK_SPAN.finditer(s):
        if m.start() > i:
            runs.append((s[i:m.start()], default))
        word = (m.group("word") or "").lower()
        colour = (_CLASS_MARK.get(_WORD_MARK.get(word))
                  or _CLASS_MARK.get(m.group("mark"), default))
        runs.append((m.group(0), colour))
        i = m.end()
    if i < len(s):
        runs.append((s[i:], default))
    return [r for r in runs if r[0]] or [(s, default)]

def mark_spans(text, default=None, style=None):
    """`mark_runs` as Dash spans - the dashboard half of the same rule."""
    base = dict(style or {})
    return [html.Span(t, style={**base, "color": c})
            for t, c in mark_runs(text, default)]

# ── Export spec collector ────────────────────────────────────────────
# While a collector list is set, every panel-level builder also records
# a structured spec (geometry + payload). The native PowerPoint export
# uses these specs to build editable shapes/charts instead of images.
_COLLECTOR: list | None = None

def set_collector(lst):
    global _COLLECTOR
    _COLLECTOR = lst

def _collect(spec: dict):
    if _COLLECTOR is not None:
        _COLLECTOR.append(spec)

# ── Canvas & header band ─────────────────────────────────────────────
def _footer_band():
    """The template's closing bar, where it has one.

    The deck has always drawn a copyright line; the dashboard drew nothing,
    so the two surfaces disagreed about the foot of the page.  Now both read
    the same tokens: a solid bar with white text (FMI), or nothing at all
    (Fact.MR, whose deck keeps its muted line).
    """
    if not FOOTER_BAND:
        return []
    return [html.Div(COPYRIGHT, style={
        "position": "absolute", "left": 0, "top": FOOTER_RULE_Y,
        "width": CANVAS_W, "height": FOOTER_BAND_H,
        "background": FOOTER_BAND, "color": FOOTER_TEXT_COLOR,
        "fontSize": FS_SOURCE, "display": "flex", "alignItems": "center",
        "padding": f"0 {FOOTER_MARGIN}px", "boxSizing": "border-box"})]


def canvas(children):
    """Fixed 1280x720 report page, light-blue like the PBIX background."""
    kids = list(children) if isinstance(children, (list, tuple)) else [children]
    return html.Div(kids + _footer_band(), style={
        "position": "relative", "width": CANVAS_W, "height": CANVAS_H,
        "background": PAGE_BG, "flexShrink": 0,
        "boxShadow": "0 2px 10px rgba(0,40,80,0.18)",
        "fontFamily": FONT_STACK,
    })

# ── Page header ──────────────────────────────────────────────────────
# THE ONE PLACE the two templates' page chrome differs on the dashboard:
#   Fact.MR (HEADER_STYLE "band") fills (13,9) 1254x75 with the brand colour
#   and sets the title in white, as the PBIX report did.
#   FMI      (HEADER_STYLE "rule") leaves the zone white, sets the title in
#   dark ink and draws the template's light-blue rule underneath - the slide
#   master's own arrangement.
# Everything else (geometry, the title id, the slicers) is shared, so pages
# paginate identically under either template.

def _title_style():
    """Title text style for the active template's header."""
    st = {"color": TITLE_COLOR, "fontSize": 17, "fontWeight": 700,
          "letterSpacing": 0.2, "lineHeight": 1.2}
    if HEADER_STYLE == "band":
        st["textShadow"] = "0 1px 2px rgba(0,0,0,0.25)"
    return st


def _band_style():
    """Container style for the header zone."""
    # A SEGMENTED rule (the FMI template's green/orange/purple bar, one
    # segment per content column) cannot be a CSS border - it is drawn as
    # sibling divs by `_rule_segments` - so the border is dropped when the
    # active theme asks for segments.
    segmented = bool(HEADER_RULE_SEGMENTS) and RULE_H
    return {"position": "absolute", "left": 13, "top": 9,
            "width": 1254, "height": 75, "background": HEADER_FILL,
            "display": "flex", "alignItems": "center",
            "padding": "0 18px", "borderRadius": 2,
            "boxShadow": HEADER_SHADOW,
            "borderBottom": (f"{RULE_H}px solid {HEADER_RULE}"
                             if HEADER_RULE and RULE_H and not segmented
                             else None)}


def _rule_segments():
    """The header rule as coloured segments, or nothing when undivided."""
    if not (HEADER_RULE_SEGMENTS and RULE_H):
        return []
    return [html.Div(style={
        "position": "absolute", "left": 13 + round(start * 1254),
        "top": RULE_Y, "width": round(width * 1254), "height": RULE_H,
        "background": colour}) for start, width, colour in HEADER_RULE_SEGMENTS]


def header_band(title_id: str, slicers):
    """Title zone at (13,9) 1254x75 with dynamic title + slicers."""
    return html.Div([
        html.Div(
            [
                html.Div(id=title_id, style=_title_style()),
                html.Div(slicers, style={"display": "flex", "gap": 10,
                                         "alignItems": "center"}),
            ],
            style={**_band_style(), "justifyContent": "space-between"},
        ),
        *_rule_segments(),
    ])

def static_band(title_text: str):
    """A plain title zone (no slicers) - used for the follow-on canvases of a
    flow page that spilled onto extra slides."""
    return html.Div([
        html.Div(html.Div(title_text, style=_title_style()),
                 style=_band_style()),
        *_rule_segments(),
    ])

def slicer(dd_id: str, label: str, options, value, width=185,
           persistence=True):
    """Region/Country/Segment dropdown, styled like the PBIX slicer tile.
    Selections persist across pages (synced slicers).  ``persistence`` can
    be given a distinct key so a drill-through selection wins over a stale
    persisted value."""
    return html.Div(
        [
            html.Div(label, style={"fontSize": 9.5, "fontWeight": 600,
                                   "color": HEADER_LABEL,
                                   "marginBottom": 1,
                                   "textTransform": "uppercase",
                                   "letterSpacing": 0.6}),
            dcc.Dropdown(id=dd_id,
                         options=[{"label": o, "value": o} for o in options],
                         value=value, clearable=False,
                         persistence=persistence,
                         persistence_type="session",
                         style={"width": width, "fontSize": 12,
                                "minHeight": 30, "height": 30}),
        ],
    )

# ── Visual containers ────────────────────────────────────────────────
def panel(x, y, w, h, children, title: str | None = None, pad=8,
          export_lines=None, collect=True):
    """White bordered visual container at exact PBIX coordinates.

    ``export_lines`` (list of (text, level) tuples) makes free-text
    panels exportable to PowerPoint as editable text boxes.
    """
    if collect and export_lines is not None:
        _collect({"kind": "text", "x": x, "y": y, "w": w, "h": h,
                  "title": title, "lines": export_lines})
    kids = []
    if title:
        kids.append(html.Div([
            # full title shown - wraps onto a second line instead of
            # truncating with an ellipsis; YoY legend keywords are coloured
            html.Div([html.Span(t, style={"color": c})
                      for t, c in title_runs(title)], style={
                "fontSize": 11, "fontWeight": 700,
                "lineHeight": 1.2, "flex": 1, "wordBreak": "break-word"}),
            # dashboard-only panel chrome; never reaches the deck.  The mask
            # icon matches the rest of the shell - see app._ICON_PATHS.
            html.Div([html.Span(html.Span(className="ico ico-expand"),
                                className="vh-icon", title="Focus")],
                     className="vh-icons"),
        ], style={"display": "flex", "alignItems": "flex-start", "gap": 4,
                  "padding": "6px 8px 3px"}))
    kids.append(html.Div(children, style={
        "flex": 1, "minHeight": 0, "padding": pad,
        "paddingTop": 2 if title else pad,
        "display": "flex", "flexDirection": "column"}))
    return html.Div(kids, className="viz-panel", style={
        "position": "absolute", "left": x, "top": y, "width": w,
        "height": h, "background": PANEL_BG,
        "border": f"1px solid {PANEL_EDGE}", "borderRadius": PANEL_RADIUS,
        "boxShadow": PANEL_SHADOW,
        "display": "flex", "flexDirection": "column", "overflow": "hidden",
    })

# ── Taxonomy flow chart (org-chart nodes + connectors) ───────────────
# Renders the pure geometry produced by layout.taxonomy_flow into
# absolutely-positioned boxes and thin connector segments.  The PowerPoint
# export draws the SAME geometry with native shapes, so the two match.
_FLOW_LINE = FLOW_LINE          # connector colour
_FLOW_LEAF_EDGE = FLOW_LEAF_EDGE   # leaf box border
_FLOW_LEAF_TEXT = FLOW_LEAF_TEXT   # leaf box text

def _flow_seg(x1, y1, x2, y2, color=_FLOW_LINE, thick=2):
    """One axis-aligned connector segment as a thin div."""
    if y1 == y2:                                    # horizontal
        left, top, w, h = min(x1, x2), y1 - thick / 2, abs(x2 - x1), thick
    else:                                           # vertical
        left, top, w, h = x1 - thick / 2, min(y1, y2), thick, abs(y2 - y1)
    return html.Div(style={
        "position": "absolute", "left": left, "top": top,
        "width": max(w, thick), "height": max(h, thick),
        "background": color})

def _flow_node(n):
    """One taxonomy box (root / branch / leaf), themed per level."""
    lvl = n["level"]
    if lvl == 0:
        bg, fg, bd, weight = HEADER_BG, "#FFFFFF", HEADER_BG, 700
    elif lvl == 1:
        bg, fg, bd, weight = ACCENT, "#FFFFFF", ACCENT, 700
    else:
        bg, fg, bd, weight = PANEL_BG, _FLOW_LEAF_TEXT, _FLOW_LEAF_EDGE, 500
    align = n.get("align", "center")
    return html.Div(n["text"], style={
        "position": "absolute", "left": n["x"], "top": n["y"],
        "width": n["w"], "height": n["h"], "background": bg, "color": fg,
        "border": f"1px solid {bd}", "borderRadius": 6, "fontSize": n["fs"],
        "fontWeight": weight, "lineHeight": 1.2, "display": "flex",
        "alignItems": "center",
        "justifyContent": "flex-start" if align == "left" else "center",
        "textAlign": align, "boxSizing": "border-box", "overflow": "hidden",
        "wordBreak": "break-word",
        "padding": "3px 9px" if align == "left" else "3px 6px",
        "boxShadow": ("0 1px 3px rgba(0,40,80,0.18)" if lvl < 2
                      else "0 1px 2px rgba(0,0,0,0.08)")})

def flow_chart(nodes, edges):
    """A list of absolutely-positioned components (connectors first, then
    boxes) drawing a taxonomy org chart onto the report canvas."""
    kids = []
    for e in edges:
        pts = e["pts"]
        for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
            kids.append(_flow_seg(x1, y1, x2, y2))
    kids.extend(_flow_node(n) for n in nodes)
    return kids

# ── Market Signals editorial (3-column table) ────────────────────────
_SIG_HEAD_BG = SIG_HEAD_BG      # header cells
_SIG_EDGE = SIG_EDGE            # cell borders
_SIG_ZEBRA = SIG_ZEBRA          # alternate row tint

def signals_page(sig):
    """A 3-column editorial table (Questions / Observations / Data points)
    on the report canvas - mirrors the native PPT table."""
    cols, rows = sig["cols"], sig["rows"]
    fr = [0.30, 0.42, 0.28]                     # column width fractions
    fr = fr[:len(cols)] + [1.0 / len(cols)] * max(0, len(cols) - len(fr))
    # fill the page: the table spans the whole content band (96 -> the bottom
    # margin) and the body rows share the leftover height evenly, exactly as
    # export_native._render_signals_native sizes the native PPT table - so the
    # two surfaces stay in step whatever the row count.
    SIG_TOP, SIG_H, SIG_HEAD_H = 96, 610, 38
    body_h = max(1, (SIG_H - SIG_HEAD_H) // max(1, len(rows)))

    def cell(text, w, *, head=False, first=False, zebra=False):
        style = {"width": f"{w*100:.4f}%", "border": f"1px solid {_SIG_EDGE}",
                 "padding": "9px 12px", "verticalAlign": "top",
                 "lineHeight": 1.32, "wordBreak": "break-word",
                 "whiteSpace": "normal",
                 "height": f"{SIG_HEAD_H if head else body_h}px"}
        if head:
            style.update({"background": _SIG_HEAD_BG, "color": HEADER_BG,
                          "fontWeight": 700, "fontSize": 12,
                          "textTransform": "uppercase", "letterSpacing": 0.4,
                          "textAlign": "left"})
        else:
            style.update({"background": _SIG_ZEBRA if zebra else PANEL_BG,
                          "fontSize": 11.5,
                          "color": HEADER_BG if first else TEXT,
                          "fontWeight": 600 if first else 400})
        return html.Td(text, style=style)

    header = html.Tr([cell(c, fr[i], head=True) for i, c in enumerate(cols)])
    body = []
    for r, row in enumerate(rows):
        body.append(html.Tr([
            cell(v, fr[i], first=(i == 0), zebra=(r % 2 == 1))
            for i, v in enumerate(row)]))
    table = html.Table([html.Thead(header), html.Tbody(body)], style={
        "width": "100%", "height": "100%", "borderCollapse": "collapse",
        "tableLayout": "fixed",
        "fontFamily": FONT_STACK})
    return [html.Div(table, style={
        "position": "absolute", "left": 15, "top": SIG_TOP, "width": 1250,
        "height": SIG_H,
        "background": PANEL_BG, "border": f"1px solid {PANEL_EDGE}",
        "borderRadius": 4, "overflow": "hidden",
        "boxShadow": "0 1px 4px rgba(0,40,80,0.12)"})]

# ── Shared primitive renderer (canvas drawing list -> Dash divs) ────
# A flat primitive list (text/rect/rrect/line/poly/shape/dot) is drawn here for
# the dashboard and by export_native._draw_primitives for the PPT, so any custom
# slide built on primitives matches exactly on both surfaces.
_PRIM_ALIGN = {"l": "left", "c": "center", "r": "right"}

def draw_primitives(prims):
    """Render a shared primitive list as absolutely-positioned divs.  Diagonal
    lines become rotated 1-px bars and filled polygons use CSS clip-path, so no
    SVG is required."""
    import math
    out = []

    def _kw(p, base):                              # attach hover tooltip + class
        tip = p.get("tip")
        if tip:
            base["title"] = tip
            base["className"] = "pp-hot"
        return base
    for p in prims:
        t = p["t"]
        if t == "text":
            style = {"position": "absolute", "left": p["x"], "top": p["y"],
                     "width": p["w"], "height": p["h"], "fontSize": p["size"],
                     "fontWeight": 700 if p["bold"] else 400, "color": p["color"],
                     "textAlign": _PRIM_ALIGN.get(p["align"], "left"),
                     "lineHeight": f"{p.get('lh', p['h'])}px",
                     "overflow": "hidden",
                     "whiteSpace": "normal" if p.get("wrap") else "nowrap",
                     "fontFamily": FONT_STACK}
            if p.get("underline"):
                style["textDecoration"] = "underline"
            out.append(html.Div(p["s"], **_kw(p, {"style": style})))
        elif t in ("rect", "rrect"):
            style = {"position": "absolute", "left": p["x"], "top": p["y"],
                     "width": p["w"], "height": p["h"], "background": p["fill"],
                     "boxSizing": "border-box"}
            lw = p.get("lw", 1)
            style["border"] = ("none" if p.get("line") in (None, "none")
                               else f"{lw}px {p.get('dash', 'solid')} {p['line']}")
            if t == "rrect":
                style["borderRadius"] = p.get("r", 8)
            out.append(html.Div(**_kw(p, {"style": style})))
        elif t in ("line", "poly"):
            segs = ([(p["x1"], p["y1"], p["x2"], p["y2"])] if t == "line"
                    else [(a[0], a[1], b[0], b[1])
                          for a, b in zip(p["pts"], p["pts"][1:])])
            th = max(0.8, p["width"])
            dash = p.get("dash")
            for x1, y1, x2, y2 in segs:
                dx, dy = x2 - x1, y2 - y1
                length = math.hypot(dx, dy)
                ang = math.degrees(math.atan2(dy, dx))
                st = {"position": "absolute", "left": x1, "top": y1 - th / 2,
                      "width": length, "height": th, "transformOrigin": "0 50%",
                      "transform": f"rotate({ang}deg)"}
                if dash:                            # dashed rule via border
                    st["borderTop"] = f"{th}px dashed {p['color']}"
                    st["height"] = 0
                else:
                    st["background"] = p["color"]
                    st["borderRadius"] = th
                out.append(html.Div(style=st))
        elif t == "shape":                          # filled polygon (clip-path)
            xs = [px for px, _ in p["pts"]]
            ys = [py for _, py in p["pts"]]
            bx, by = min(xs), min(ys)
            bw, bh = (max(xs) - bx) or 1, (max(ys) - by) or 1
            poly = ", ".join(f"{(px-bx)/bw*100:.2f}% {(py-by)/bh*100:.2f}%"
                             for px, py in p["pts"])
            style = {"position": "absolute", "left": bx, "top": by,
                     "width": bw, "height": bh, "background": p["fill"],
                     "clipPath": f"polygon({poly})"}
            out.append(html.Div(**_kw(p, {"style": style})))
        elif t == "dot":
            d = p["d"]
            web = p.get("web_only")
            style = {"position": "absolute", "left": p["cx"] - d / 2,
                     "top": p["cy"] - d / 2, "width": d, "height": d,
                     "borderRadius": "50%", "boxSizing": "border-box"}
            if web:                                # invisible hover target
                style["background"] = "transparent"
            else:
                style["background"] = p["color"]
                style["border"] = f"{p.get('lw', 1)}px solid {p['edge']}"
            kw = _kw(p, {"style": style})
            if "className" in kw:
                kw["className"] += " pp-dot"
            out.append(html.Div(**kw))
    return out

def primitives_view(prims):
    """Dashboard view for a generic 'primitives' custom block."""
    out = draw_primitives(prims)
    out.append(html.Div(SOURCE_TEXT, style={
        "position": "absolute", "left": 15, "top": 708, "fontSize": 9,
        "fontStyle": "italic", "color": MUTED}))
    return out

# ── Pricing-pattern page (reference "Global Pricing Assessment") ────
def price_pattern_view(data):
    """Per-region pricing-pattern page on the report canvas - mirrors the native
    PPT slide (export_native._render_pricepattern_native)."""
    from . import layout
    out = draw_primitives(layout.pricepattern_primitives(data))
    out.append(html.Div(SOURCE_TEXT, style={
        "position": "absolute", "left": 15, "top": 708, "fontSize": 9,
        "fontStyle": "italic", "color": MUTED}))
    return out

# ── Segment-section opening page (analysis overview) ────────────────
def seg_start_view(spec):
    """Creative segment opener on the report canvas: intro line, a themed
    sub-segment value matrix (EST/END value + CAGR) and a three-column
    description band.  Mirrors the native PPT slide (_render_segstart)."""
    cols = spec["cols"]
    rl = spec["row_labels"]
    n = len(cols) or 1
    MX, NAVY = 34, VC_NAVY
    out = []
    # intro line
    out.append(html.Div(spec["intro"], style={
        "position": "absolute", "left": MX, "top": 90, "width": CANVAS_W - 2 * MX,
        "fontSize": 11.5, "lineHeight": 1.4, "color": TEXT}))
    # ── value matrix ──
    lab_w = 150
    x0 = MX + lab_w + 12
    col_w = (CANVAS_W - MX - x0 - (n - 1) * 12) / n
    badge_d, badge_y, head_y, head_h, r_h = 34, 138, 174, 48, 52
    rows_y = [head_y + head_h + 2 + k * (r_h + 2) for k in range(3)]
    # left row-label cells
    lab_specs = [(rl[0], head_y, head_h, True),
                 (rl[1], rows_y[0], r_h, False),
                 (rl[2], rows_y[1], r_h, False),
                 (rl[3], rows_y[2], r_h, False)]
    for txt, ly, lh, is_head in lab_specs:
        out.append(html.Div(txt, style={
            "position": "absolute", "left": MX, "top": ly, "width": lab_w,
            "height": lh, "display": "flex", "alignItems": "center",
            "padding": "0 10px", "boxSizing": "border-box",
            "background": "#EEF3F7" if is_head else "#F7FAFC",
            "border": "1px solid #E1E8EE", "fontSize": 10.5,
            "fontWeight": 700 if is_head else 600,
            "color": NAVY if is_head else MUTED, "lineHeight": 1.15}))
    for i, c in enumerate(cols):
        cx = x0 + i * (col_w + 12)
        color = c["color"]
        out.append(html.Div(str(i + 1), style={
            "position": "absolute", "left": cx + col_w / 2 - badge_d / 2,
            "top": badge_y, "width": badge_d, "height": badge_d,
            "borderRadius": "50%", "background": color, "color": "#FFFFFF",
            "display": "flex", "alignItems": "center",
            # a two-digit badge must stay on ONE line, whatever its width
            "whiteSpace": "nowrap", "lineHeight": 1,
            "justifyContent": "center", "fontSize": 15, "fontWeight": 700,
            "border": "3px solid #EAF2F8", "boxSizing": "border-box",
            "zIndex": 3}))
        out.append(html.Div(c["name"], style={
            "position": "absolute", "left": cx, "top": head_y, "width": col_w,
            "height": head_h, "background": color, "color": "#FFFFFF",
            "display": "flex", "alignItems": "center",
            "justifyContent": "center", "textAlign": "center",
            "padding": "0 6px", "boxSizing": "border-box", "fontSize": 11,
            "fontWeight": 700, "lineHeight": 1.12,
            "borderTopLeftRadius": 7, "borderTopRightRadius": 7}))
        for k, val in enumerate((c["v26"], c["v36"], c["cagr"])):
            out.append(html.Div(val, style={
                "position": "absolute", "left": cx, "top": rows_y[k],
                "width": col_w, "height": r_h, "background":
                "#FFFFFF" if k % 2 == 0 else "#F5F8FB",
                "border": f"1px solid {PANEL_EDGE}", "display": "flex",
                "alignItems": "center", "justifyContent": "center",
                "boxSizing": "border-box", "fontSize": 15, "fontWeight": 700,
                "color": color}))
    # ── three description columns (equal-height panel row) ──
    _seg_desc_columns(out, spec, 402, NAVY)
    out.append(html.Div(SOURCE_TEXT, style={
        "position": "absolute", "left": MX, "top": 694, "fontSize": 9,
        "fontStyle": "italic", "color": MUTED}))
    return out

def fit_bullet_fs(chunks, text_w, avail_h,
                  sizes=(11, 10.5, 10, 9.5, 9, 8.5, 8, 7.5)):
    """Largest font size at which `chunks` (paragraph strings) fit `avail_h`
    inside a `text_w` column, with the gap that goes under each bullet.

    The Dash twin of export_native._fit_text_fs, and deliberately measured with
    the SAME layout._wrap_count model, so a bullet list that fits the deck also
    fits the dashboard. Without this the dashboard silently clipped: panels set
    `overflow: hidden`, so a 13-geography insight list lost its tail.

    `layout` is imported lazily because layout imports THIS module.
    """
    from . import layout
    for fs in sizes:
        lh = 1.34 if fs >= 10 else 1.28
        gap = 9 if fs >= 10 else 6
        total = 0.0
        for c in chunks:
            lines = max(1, layout._wrap_count(str(c), text_w, fs))
            total += lines * fs * lh + gap
        if total <= avail_h - 2:
            return fs, gap
    return sizes[-1], 5


def _seg_desc_columns(out, spec, dy, NAVY):
    ph = 684 - dy                                 # equal panel height
    ptitle = lambda t: html.Div(t, style={
        "fontSize": 13, "fontWeight": 700, "color": NAVY, "paddingBottom": 5,
        "marginBottom": 9, "borderBottom": f"2px solid {ACCENT}"})

    def panel(x, w, children, bg="#FFFFFF", edge="#DCE6EF"):
        return html.Div(children, style={
            "position": "absolute", "left": x, "top": dy, "width": w,
            "height": ph, "background": bg, "border": f"1px solid {edge}",
            "borderRadius": 9, "padding": "13px 15px", "boxSizing": "border-box",
            "boxShadow": "0 2px 7px rgba(0,40,80,0.09)", "overflow": "hidden"})

    # Panel geometry.  The widest box goes to the panel with the MOST text:
    # 'Sub-Segment Insights' carries one bullet per sub-segment/geography, while
    # 'Segment Overview' carries two bullets and a stat.  Mirrored in
    # export_native._render_segstart_native.
    A_X, A_W = 34, 340
    B_X, B_W = 386, 470
    C_X, C_W = 868, CANVAS_W - 34 - 868
    INNER = 30                                   # 15px padding each side
    BODY_H = ph - 26 - 32                        # padding + heading strip

    # A · Key Highlights
    a_fs, a_gap = fit_bullet_fs([str(b) for b in spec["highlights"]],
                                A_W - INNER - 18, BODY_H)
    a_items = [html.Div([
        html.Span("▪", style={"color": ACCENT, "marginRight": 8,
                              "fontSize": a_fs - 1}),
        html.Span(b, style={"color": TEXT})],
        style={"fontSize": a_fs, "lineHeight": 1.4, "marginBottom": a_gap + 2,
               "display": "flex", "alignItems": "flex-start"})
        for b in spec["highlights"]]
    out.append(panel(A_X, A_W, [ptitle("Key Highlights"), *a_items]))
    # B · Sub-Segment Insights - one bullet per sub-segment, so this is the
    # panel that overflows on a market with many geographies.  Auto-fitted.
    b_chunks = [f"{it['name']}: {it['text']}" for it in spec["insights"]]
    b_fs, b_gap = fit_bullet_fs(b_chunks, B_W - INNER - 19, BODY_H)
    dot = max(6, round(b_fs * 0.9))
    b_rows = []
    for it in spec["insights"]:
        b_rows.append(html.Div([
            html.Div(style={"width": dot, "height": dot,
                            "borderRadius": "50%",
                            "background": it["color"], "marginRight": 9,
                            "marginTop": 4, "flexShrink": 0}),
            html.Div([
                html.Span(it["name"] + ": ", style={
                    "fontWeight": 700, "color": it["color"]}),
                html.Span(it["text"], style={"color": TEXT})],
                style={"fontSize": b_fs, "lineHeight": 1.34})],
            style={"display": "flex", "alignItems": "flex-start",
                   "marginBottom": b_gap}))
    out.append(panel(B_X, B_W, [ptitle("Sub-Segment Insights"), *b_rows]))
    # C · Segment Overview (accent-tinted panel)
    cx, cw = C_X, C_W
    c_fs, c_gap = fit_bullet_fs([str(p) for p in spec["overview"]["points"]],
                                cw - INNER - 18, BODY_H - 74)
    c_children = [ptitle("Segment Overview")]
    c_children += [html.Div([
        html.Span("✓", style={"color": GOOD, "fontWeight": 700,
                              "marginRight": 8}),
        html.Span(p, style={"color": TEXT})],
        style={"fontSize": c_fs, "lineHeight": 1.4, "marginBottom": c_gap + 2,
               "display": "flex", "alignItems": "flex-start"})
        for p in spec["overview"]["points"]]
    c_children.append(html.Div(spec["overview"]["stat"], style={
        "marginTop": 6, "padding": "11px 13px", "background": "#FFFFFF",
        "borderLeft": f"4px solid {ACCENT}", "borderRadius": 5,
        "fontSize": 11.5, "fontWeight": 600, "color": NAVY, "lineHeight": 1.4}))
    out.append(panel(cx, cw, c_children, bg="#EAF3FB", edge="#CFE2F2"))

# ── Segment-section closing card (What / Why / How / Key States) ─────
def end_cards(spec):
    """A creative section-divider on the report canvas: four colour-headed
    panels with numbered badges and chevron connectors, a 'This concludes...'
    line and a 'Next ...' pill.  Mirrors the native PPT slide (_render_endcards)."""
    cards = spec["cards"]
    n = len(cards) or 1
    MX, GAP = 34, 30
    cw = (CANVAS_W - 2 * MX - (n - 1) * GAP) / n
    badge_d, card_top, head_h, card_bot = 54, 172, 50, 502
    NAVY = VC_NAVY
    out = []
    for i, c in enumerate(cards):
        cx = MX + i * (cw + GAP)
        color = c["color"]
        out.append(html.Div([
            html.Div(c["label"], style={
                "height": head_h, "background": color, "color": "#FFFFFF",
                "display": "flex", "alignItems": "center",
                "justifyContent": "center", "fontSize": 16, "fontWeight": 700,
                "borderTopLeftRadius": 11, "borderTopRightRadius": 11}),
            html.Div(c["body"], style={
                "height": card_bot - card_top - head_h, "padding": "0 16px",
                "display": "flex", "alignItems": "center", "fontSize": 12.5,
                "lineHeight": 1.55, "color": TEXT, "boxSizing": "border-box"})],
            style={
                "position": "absolute", "left": cx, "top": card_top,
                "width": cw, "height": card_bot - card_top,
                "background": PANEL_BG, "borderRadius": 11,
                "boxShadow": "0 3px 9px rgba(0,40,80,0.14)",
                "overflow": "hidden", "boxSizing": "border-box"}))
        out.append(html.Div(c["num"], style={
            "position": "absolute", "left": cx + cw / 2 - badge_d / 2,
            "top": card_top - badge_d + 14, "width": badge_d, "height": badge_d,
            "borderRadius": "50%", "background": color, "color": "#FFFFFF",
            "display": "flex", "alignItems": "center",
            "whiteSpace": "nowrap", "lineHeight": 1,
            "justifyContent": "center", "fontSize": 22, "fontWeight": 700,
            "border": "4px solid #EAF2F8", "boxSizing": "border-box",
            "zIndex": 3}))
        if i < n - 1:
            out.append(html.Div("›", style={
                "position": "absolute", "left": cx + cw + GAP / 2 - 11,
                "top": card_top + head_h / 2 - 20, "width": 22, "fontSize": 34,
                "fontWeight": 700, "color": color, "textAlign": "center",
                "zIndex": 2}))
    # concluding line with decorative rules on each side
    out.append(html.Div(spec["footer"], style={
        "position": "absolute", "left": 0, "top": 546, "width": CANVAS_W,
        "textAlign": "center", "fontSize": 16.5, "fontWeight": 700,
        "color": NAVY, "letterSpacing": 0.2}))
    for side in ("left", "right"):
        out.append(html.Div(style={
            "position": "absolute", "top": 556, side: 175, "width": 110,
            "height": 2, "background": "#9FB6CC"}))
    pill_w, pill_h = 372, 44
    out.append(dcc.Link("➜   " + spec["next"],
                        href=f"/{spec.get('next_slug', '')}",
                        className="endcard-next", style={
        "position": "absolute", "left": CANVAS_W / 2 - pill_w / 2, "top": 590,
        "width": pill_w, "height": pill_h, "background": NAVY,
        "color": "#FFFFFF", "borderRadius": 22, "display": "flex",
        "alignItems": "center", "justifyContent": "center", "fontSize": 13.5,
        "fontWeight": 600, "textDecoration": "none", "cursor": "pointer",
        "boxShadow": "0 3px 8px rgba(0,40,80,0.30)"}))
    out.append(html.Div(SOURCE_TEXT, style={
        "position": "absolute", "left": 30, "top": 694, "fontSize": 9,
        "fontStyle": "italic", "color": MUTED}))
    return out

# ── Contents navigation (button grid) ────────────────────────────────
def contents_view(items):
    """Render the shared button-grid items as clickable teal navigation cards
    (dcc.Link) on the light-blue report canvas - matches the PPT slide."""
    from dash import dcc
    out = []
    for it in items:
        out.append(dcc.Link(it["label"], href=f"/{it['slug']}",
                            className="toc-btn", style={
            "position": "absolute", "left": it["x"], "top": it["y"],
            "width": it["w"], "height": it["h"], "display": "flex",
            "alignItems": "center", "justifyContent": "center",
            "textAlign": "center", "padding": "6px 18px",
            "boxSizing": "border-box", "lineHeight": 1.25}))
    return out

# ── Editorial summary panel (Summary-section pages) ──────────────────
def summary_panel(x, y, w, h, title, bullets):
    """A titled white card with an accent rule and bullet lines - editorial
    layout for the Summary section.  Collects a spec so the PPT export can
    rebuild it as native, editable shapes/text."""
    _collect({"kind": "sumpanel", "x": x, "y": y, "w": w, "h": h,
              "title": title, "bullets": list(bullets)})
    head = html.Div(title, style={
        "fontSize": 12.5, "fontWeight": 700, "color": HEADER_BG,
        "paddingBottom": 5, "marginBottom": 7,
        "borderBottom": f"2px solid {ACCENT}", "lineHeight": 1.2})
    items = [html.Div([
        html.Span("▪", style={"color": ACCENT, "marginRight": 8,
                              "fontSize": 9}),
        # a bullet that says "green ▲" / "red ▼" is set in those colours -
        # same split the PPT export uses, so both surfaces read alike
        html.Span(mark_spans(b))],
        style={"fontSize": 11.5, "lineHeight": 1.4, "marginBottom": 6,
               "display": "flex", "alignItems": "flex-start"})
        for b in bullets]
    return html.Div([head, *items], style={
        "position": "absolute", "left": x, "top": y, "width": w, "height": h,
        "background": PANEL_BG, "border": f"1px solid {PANEL_EDGE}",
        "borderRadius": 4, "boxShadow": "1px 2px 5px rgba(0,40,80,0.10)",
        "padding": "10px 14px", "boxSizing": "border-box",
        "overflow": "hidden"})

# Local topojson (bundled under assets/) so maps render with no internet.
#
# RELATIVE, not "/assets/". The report is mounted at the site root when it is
# run standalone and at "/r/<market>/" when the portal serves it, so a leading
# slash sends plotly to /assets/world_110m.json, which exists under neither
# mount in the portal - it 404s, and every map on the country pages draws
# empty while the rest of the page looks fine. A relative URL resolves against
# the page it is drawn on: "/r/<mid>/<slug>" gives "/r/<mid>/assets/" and
# "/<slug>" gives "/assets/", which is right in both cases.
#
# It has to stay relative rather than being built from the app's
# `requests_pathname_prefix`, because this is module state shared by every
# market in the process and one prefix baked in here would be wrong for all
# the other markets the same worker serves. This is the same shape of bug as
# the assets_folder note in app.create_app.
_GRAPH_CONFIG = {"displayModeBar": False, "topojsonURL": "assets/"}

def graph_panel(x, y, w, h, figure, title: str | None = None,
                graph_id=None):
    """White panel wrapping a Plotly figure.  ``graph_id`` gives the
    ``dcc.Graph`` an id so its click events can drive cross-filtering."""
    # EVERY figure on the dashboard is drawn through here, and the spec
    # collected just below is what the PPT export reuses - so this is the one
    # place that can guarantee a chart-wide rule reaches both surfaces.
    # cap_bar_widths keeps a one- or two-category bar chart from drawing its
    # bars as full-width slabs; it is idempotent, so pages that also go through
    # layout.fig() are unaffected.
    from . import figures as _f
    _f.cap_bar_widths(figure)
    _collect({"kind": "figure", "x": x, "y": y, "w": w, "h": h,
              "title": title, "figure": figure})
    kw = {"figure": figure, "config": _GRAPH_CONFIG,
          "style": {"height": "100%", "width": "100%"}}
    if graph_id:
        kw["id"] = graph_id
    return panel(x, y, w, h, dcc.Graph(**kw), title=title, pad=2,
                 collect=False)

def table_panel(x, y, w, h, df: pd.DataFrame, index_name: str = "",
                title: str | None = None, total_row=None, body_id=None,
                scroll=True, source=False):
    """Panel wrapping a pivot-style table; exportable as a native
    PowerPoint table.  ``body_id`` ids the inner container so a callback
    can swap the table (used for click-to-cross-filter).  ``scroll=False``
    (layout engine) removes the internal scrollbar; ``source=True`` adds the
    'Source: Future Market Insights' caption under the table."""
    _collect({"kind": "table", "x": x, "y": y, "w": w, "h": h,
              "title": title, "df": df, "index_name": index_name,
              "total_row": total_row})
    content = df_table(df, index_name=index_name, total_row=total_row,
                       scroll=scroll)
    if body_id:
        content = html.Div(content, id=body_id,
                           style={"flex": 1, "minHeight": 0,
                                  "display": "flex",
                                  "flexDirection": "column"})
    children = [content, _source_line()] if source else content
    return panel(x, y, w, h, children, title=title, collect=False)

def kpi_card(x, y, w, h, label: str, value: str, sub: str = ""):
    """PBIX card visual: big number, small label."""
    _collect({"kind": "kpi", "x": x, "y": y, "w": w, "h": h,
              "label": label, "value": value, "sub": sub})
    kids = [
        html.Div(label, style={"fontSize": 10.5, "fontWeight": 600,
                               "color": MUTED, "textTransform": "uppercase",
                               "letterSpacing": 0.5}),
        html.Div(value, style={"fontSize": 24 if h > 80 else 19,
                               "fontWeight": 700, "color": ACCENT,
                               "lineHeight": 1.15}),
    ]
    if sub:
        kids.append(html.Div(sub, style={"fontSize": 9.5, "color": FAINT}))
    return html.Div(kids, style={
        "position": "absolute", "left": x, "top": y, "width": w,
        "height": h, "background": KPI_TINT,
        "border": f"1px solid {PANEL_EDGE}", "borderRadius": PANEL_RADIUS,
        "boxShadow": PANEL_SHADOW,
        "display": "flex", "flexDirection": "column",
        "justifyContent": "center", "alignItems": "center",
        "textAlign": "center", "padding": "4px 8px", "gap": 2,
    })

def writeup(x, y, w, h, text):
    """The PBIX 'Write-Up' commentary strip (dynamic DAX text measure)."""
    _collect({"kind": "writeup", "x": x, "y": y, "w": w, "h": h,
              "text": text})
    return html.Div(
        html.Div(mark_spans(text), style={"fontSize": 11, "lineHeight": 1.45,
                                          "color": TEXT}),
        style={"position": "absolute", "left": x, "top": y, "width": w,
               "height": h, "background": PANEL_BG,
               "border": f"1px solid {PANEL_EDGE}", "borderRadius": PANEL_RADIUS,
               "boxShadow": PANEL_SHADOW,
               "padding": "8px 12px", "overflowY": "auto",
               "display": "flex", "alignItems": "center"},
    )

# ── Pivot-style tables ───────────────────────────────────────────────
# Heights are pinned (box-sizing: border-box) so a table's rendered size is
# deterministic from its row count - layout.py sizes panels to fit exactly
# and thus needs no scrollbar.  Keep in sync with layout.HEAD_H / ROW_H.
_TH = {"padding": "5px 8px", "height": 24, "lineHeight": "13px",
       "fontSize": 10, "color": TABLE_HEAD_TEXT,
       "background": TABLE_HEAD_BG, "textAlign": "left",
       "whiteSpace": "nowrap", "position": "sticky", "top": 0,
       "zIndex": 2}
_TD = {"padding": "4px 8px", "height": 22, "lineHeight": "13px",
       "fontSize": 10.5, "textAlign": "left",
       "borderBottom": f"1px solid {GRID}", "whiteSpace": "nowrap",
       "color": TEXT}

def df_table(df: pd.DataFrame, index_name: str = "", total_row=None,
             scroll: bool = True):
    """PBIX pivot-table look: blue header, zebra rows, sticky header.
    Cells may be strings or (text, colour) tuples.  ``scroll=False`` (used
    by the layout engine, which sizes the panel to the content) removes the
    internal scrollbar entirely."""
    groups = getattr(df, "attrs", {}).get("col_groups")
    if groups:
        # two header rows: group names over their columns, then the labels
        head = [
            html.Tr([html.Th(index_name, rowSpan=2,
                             style={**_TH, "textAlign": "left"})]
                    + [html.Th(str(g), colSpan=int(s),
                               style={**_TH, "textAlign": "center"})
                       for g, s in groups]),
            html.Tr([html.Th(str(c).split("|", 1)[-1],
                             style={**_TH, "textAlign": "center"})
                     for c in df.columns]),
        ]
    else:
        head = html.Tr(
            [html.Th(index_name, style={**_TH, "textAlign": "left"})]
            + [html.Th(str(c), style=_TH) for c in df.columns]
        )
    body = []
    for r, (idx, row) in enumerate(df.iterrows()):
        zebra = "#F2FAFD" if r % 2 else PANEL_BG
        cells = [html.Td(str(idx), style={**_TD, "textAlign": "left",
                                          "fontWeight": 600,
                                          "background": zebra})]
        for v in row:
            if isinstance(v, tuple):
                text, color = v
                cells.append(html.Td(text, style={**_TD, "color": color,
                                                  "fontWeight": 700,
                                                  "background": zebra}))
            else:
                cells.append(html.Td(v, style={**_TD, "background": zebra}))
        body.append(html.Tr(cells))
    if total_row is not None:
        cells = [html.Td("Total", style={**_TD, "textAlign": "left",
                                         "fontWeight": 700,
                                         "background": TABLE_TOTAL_BG})]
        for v in total_row:
            cells.append(html.Td(v, style={**_TD, "fontWeight": 700,
                                           "background": TABLE_TOTAL_BG}))
        body.append(html.Tr(cells))
    table = html.Table([html.Thead(head), html.Tbody(body)],
                       style={"width": "100%", "borderCollapse": "collapse"})
    return html.Div(table, style={
        "overflow": "auto" if scroll else "hidden", "flex": 1,
        "minHeight": 0})

# professional source caption shown under every data table.  The wording is a
# template token (Fact.MR / Future Market Insights) imported above; this name
# is the one the rest of the project refers to, and `theming.apply()` rebinds
# it here when the template changes.

def _source_line():
    return html.Div(SOURCE_TEXT, style={
        "fontSize": 8.5, "fontStyle": "italic", "color": MUTED,
        "paddingTop": 3, "flexShrink": 0, "whiteSpace": "nowrap",
        "letterSpacing": 0.2})


# ── Creative competitor dashboard matrix (mirrors export_native) ─────
_ARROW_HTML = {"high": ARROW_HIGH, "med": ARROW_MED, "low": ARROW_LOW}


def _abs(children, x, y, w, h, extra=None):
    st = {"position": "absolute", "left": x, "top": y, "width": w,
          "height": h, "boxSizing": "border-box"}
    if extra:
        st.update(extra)
    return html.Div(children, style=st)


def competitor_matrix_view(data):
    """Dash twin of export_native._render_competitor_matrix_native - the creative
    competitor dashboard (companies as columns; sparkle/checks/boxes portfolio,
    revenue, a conic-gradient share gauge, presence arrows, strategy badge)."""
    rows = data["rows"]
    n = len(rows) or 1
    star_cols = data.get("star_colors", list(STAR_COLORS))
    box_cols = data.get("box_colors", list(BOX_COLORS))
    labels = data.get("row_labels", [])
    LBLX, LBLW, COLX0, GRIGHT = 14, 120, 136, 1266
    colw = (GRIGHT - COLX0) / n
    bands = {"header": (108, 38), "brand": (146, 30), "portfolio": (176, 96),
             "overall": (272, 34), "target": (306, 34), "share": (340, 92),
             "presence": (432, 100), "strategies": (532, 90)}
    top = 108
    bot = 532 + 120
    out = []
    if data.get("subtitle"):
        out.append(_abs(data["subtitle"], 30, 86, 1100, 18,
                        {"fontSize": 12, "color": MUTED}))
    # white card behind the whole matrix (so it reads clean, not on pale blue),
    # with a soft tint on the header row and the left label column
    out.append(_abs("", LBLX, top - 2, GRIGHT - LBLX, bot - top + 4,
                    {"background": PANEL_BG, "border": f"1px solid {PANEL_EDGE}",
                     "borderRadius": 3,
                     "boxShadow": "1px 2px 5px rgba(0,40,80,0.10)"}))
    out.append(_abs("", LBLX, top, GRIGHT - LBLX, bands["header"][1],
                    {"background": "#F1F7FB"}))
    out.append(_abs("", LBLX, top, LBLW, bot - top,
                    {"background": "#F7FAFC"}))
    # grid lines
    for key in ("brand", "portfolio", "overall", "target", "share",
                "presence", "strategies"):
        out.append(_abs("", LBLX, bands[key][0], GRIGHT - LBLX, 1,
                        {"borderTop": f"1px solid {GRID}"}))
    out.append(_abs("", LBLX, bot, GRIGHT - LBLX, 1,
                    {"borderTop": f"1px solid {GRID}"}))
    for c in range(n + 1):
        vx = COLX0 + c * colw
        out.append(_abs("", vx, top, 1, bot - top,
                        {"borderLeft": f"1px solid {GRID}"}))
    out.append(_abs("", LBLX, top, 1, bot - top,
                    {"borderLeft": f"1px solid {GRID}"}))
    # left labels
    lbl_bands = ["brand", "portfolio", "overall", "target", "share",
                 "presence", "strategies"]
    for lb, key in zip(labels, lbl_bands):
        y, h = bands[key]
        out.append(_abs(lb, LBLX + 4, y, LBLW - 6, h,
                        {"fontSize": 10.5, "fontWeight": 700,
                         "color": HEADER_BG, "display": "flex",
                         "alignItems": "center", "lineHeight": 1.1}))
    maxshare = max((r["share"] for r in rows), default=1.0) or 1.0
    for i, r in enumerate(rows):
        cx0 = COLX0 + i * colw
        # header
        y, h = bands["header"]
        out.append(_abs([
            html.Div(r["company"], style={"fontSize": 11, "fontWeight": 700,
                                          "color": TEXT, "textAlign": "center",
                                          "lineHeight": 1.05}),
            html.Div("", style={"width": 48, "height": 2,
                                "background": ACCENT, "margin": "3px auto 0"}),
        ], cx0 + 3, y + 2, colw - 6, h - 4,
            {"display": "flex", "flexDirection": "column",
             "justifyContent": "center"}))
        # brand
        y, h = bands["brand"]
        out.append(_abs(r["brand"], cx0 + 3, y, colw - 6, h,
                        {"fontSize": 10, "color": "#AFB4B9",
                         "textAlign": "center", "display": "flex",
                         "alignItems": "center",
                         "justifyContent": "center"}))
        # portfolio
        y, h = bands["portfolio"]
        stars = [html.Span("✦", style={
            "color": star_cols[k % len(star_cols)] if k < r["stars"]
            else STAR_EMPTY, "fontSize": 17, "margin": "0 2px"})
            for k in range(4)]
        checks = [html.Span("✔" if ok else "✘", style={
            "color": OK_COLOR if ok else NOT_OK_COLOR, "fontSize": 13,
            "fontWeight": 700, "margin": "0 3px"}) for ok in r["apps"]]
        boxes = [html.Div("", style={
            "width": 12, "height": 12, "margin": "0 3px",
            "background": box_cols[k % len(box_cols)] if on else "#FFFFFF",
            "border": f"1.5px solid {box_cols[k % len(box_cols)]}"})
            for k, on in enumerate(r["boxes"])]
        out.append(_abs([
            html.Div(stars, style={"textAlign": "center", "height": 24}),
            html.Div(checks, style={"textAlign": "center", "height": 22}),
            html.Div(boxes, style={"display": "flex",
                                   "justifyContent": "center"}),
        ], cx0 + 2, y + 4, colw - 4, h - 8,
            {"display": "flex", "flexDirection": "column",
             "alignItems": "center", "gap": 3}))
        # revenue rows
        for key, val in (("overall", r["overall_rev"]),
                         ("target", r["target_rev"])):
            y, h = bands[key]
            out.append(_abs(fmt_num(val), cx0 + 3, y, colw - 6, h,
                            {"fontSize": 11, "color": TEXT,
                             "textAlign": "center", "display": "flex",
                             "alignItems": "center",
                             "justifyContent": "center"}))
        # share gauge (conic-gradient donut)
        y, h = bands["share"]
        frac = max(0.0, min(1.0, r["share"] / maxshare))
        deg = frac * 360.0
        d = min(58, h - 20)
        gauge = html.Div(html.Div(f"{r['share']:.1f}%", style={
            "position": "absolute", "inset": f"{d * 0.22}px",
            "borderRadius": "50%", "background": PANEL_BG,
            "display": "flex", "alignItems": "center",
            "justifyContent": "center", "fontSize": 10,
            "fontWeight": 700, "color": HEADER_BG}), style={
            "position": "relative", "width": d, "height": d,
            "borderRadius": "50%",
            "background": f"conic-gradient({OK_COLOR} 0 {deg}deg, "
                          f"{GAUGE_TRACK} {deg}deg 360deg)"})
        out.append(_abs(gauge, cx0 + 3, y + 6, colw - 6, h - 12,
                        {"display": "flex", "alignItems": "center",
                         "justifyContent": "center"}))
        # presence
        y, h = bands["presence"]
        pres = [html.Div([
            html.Span(nm, style={"color": TEXT}),
            html.Span(" ▲", style={"color": _ARROW_HTML.get(st, ARROW_LOW),
                                   "fontWeight": 700})],
            style={"fontSize": 10, "lineHeight": 1.35}) for nm, st in
            r["presence"]]
        out.append(_abs(pres, cx0 + 6, y + 4, colw - 10, h - 8,
                        {"display": "flex", "flexDirection": "column",
                         "justifyContent": "center"}))
        # strategies badge
        y, h = bands["strategies"]
        def circ(letter, active):
            return html.Div(letter, style={
                "width": 17, "height": 17, "borderRadius": "50%",
                "background": NAVY_DEEP if active else "#C9CDD1",
                "color": "#FFFFFF", "fontSize": 9, "fontWeight": 700,
                "display": "flex", "alignItems": "center",
                "justifyContent": "center", "margin": 2})
        r1 = html.Div([circ(l, a) for l, a in r["strategies"][:4]],
                      style={"display": "flex", "justifyContent": "center"})
        r2 = html.Div([circ(l, a) for l, a in r["strategies"][4:]],
                      style={"display": "flex", "justifyContent": "center"})
        out.append(_abs([r1, r2], cx0 + 3, y, colw - 6, h,
                        {"display": "flex", "flexDirection": "column",
                         "justifyContent": "center", "alignItems": "center"}))
    # legend + note
    leg = []
    for k, v in data.get("legend", []):
        leg.append(html.Div([
            html.Span(k, style={"width": 15, "height": 15,
                                "borderRadius": "50%", "background": NAVY_DEEP,
                                "color": "#FFF", "fontSize": 8.5,
                                "fontWeight": 700, "display": "inline-flex",
                                "alignItems": "center",
                                "justifyContent": "center", "marginRight": 4}),
            html.Span(v, style={"fontSize": 9, "color": TEXT})],
            style={"display": "flex", "alignItems": "center",
                   "marginRight": 14}))
    out.append(_abs(leg, 30, 656, GRIGHT - 30, 16,
                    {"display": "flex", "flexWrap": "wrap"}))
    # the Product Portfolio scale, above the strategy legend
    scale = []
    if data.get("box_legend"):
        def sq(colour, hollow=False):
            return html.Span("", style={
                "width": 10, "height": 10, "display": "inline-block",
                "marginRight": 4,
                "background": "#FFF" if hollow else colour,
                "border": f"1px solid {MUTED if hollow else colour}"})

        def txt(t, colour=TEXT, bold=False):
            return html.Span(t, style={"fontSize": 9, "color": colour,
                                       "fontWeight": 700 if bold else 400,
                                       "marginRight": 8})
        scale.append(txt("Product Portfolio:", bold=True))
        for colour, name in data["box_legend"]:
            scale += [sq(colour), txt(name)]
        scale += [sq(None, True), txt("Not offered")]
        out.append(_abs(scale, 30, 624, GRIGHT - 30, 15,
                        {"display": "flex", "alignItems": "center"}))
        scale = [txt("✔", OK_COLOR, True), txt("served"),
                  txt("✘", NOT_OK_COLOR, True), txt("not served"),
                  txt("✦", STAR_COLORS[0]), txt("portfolio breadth (lit, of 4)"),
                  txt("Regional presence:", bold=True),
                  txt("▲", ARROW_HIGH, True), txt("High"),
                  txt("▲", ARROW_MED, True), txt("Medium"),
                  txt("▲", ARROW_LOW, True), txt("Low")]
        out.append(_abs(scale, 30, 640, GRIGHT - 30, 15,
                        {"display": "flex", "alignItems": "center"}))
    # ``note`` may be a single string or a LIST of lines - stacked, so several
    # footnotes fit in the strip between the legend and the page footer
    note = data.get("note")
    if note:
        lines = [note] if isinstance(note, str) else list(note)
        for li, ln in enumerate(lines):
            out.append(_abs(ln, 30, 674 + li * 11, 1180, 12,
                            {"fontSize": 8.5, "fontWeight": 600,
                             "color": WARN_TEXT}))
    return out


_VC_PURPLE, _VC_NAVY = VC_PURPLE, VC_NAVY


def _vc_icon_svg(kind, col):
    """White pictogram as an SVG data URI (matches the native _vc_icon)."""
    import base64
    body = {
        "sprout": '<rect x="22.5" y="22" width="3" height="17" fill="#fff"/>'
                  '<ellipse cx="17" cy="20" rx="8" ry="4.5" fill="#fff" '
                  'transform="rotate(-35 17 20)"/>'
                  '<ellipse cx="31" cy="20" rx="8" ry="4.5" fill="#fff" '
                  'transform="rotate(35 31 20)"/>',
        "factory": '<rect x="6" y="24" width="36" height="16" fill="#fff"/>'
                   '<rect x="9" y="13" width="7" height="12" fill="#fff"/>'
                   '<rect x="20" y="17" width="7" height="8" fill="#fff"/>'
                   '<rect x="31" y="14" width="7" height="11" fill="#fff"/>',
        "truck": '<rect x="5" y="15" width="22" height="18" fill="#fff"/>'
                 '<rect x="28" y="23" width="13" height="10" fill="#fff"/>'
                 f'<circle cx="13" cy="36" r="4.5" fill="{col}"/>'
                 f'<circle cx="32" cy="36" r="4.5" fill="{col}"/>',
        "shop": '<polygon points="5,26 24,12 43,26" fill="#fff"/>'
                '<rect x="9" y="26" width="30" height="16" fill="#fff"/>'
                f'<rect x="20" y="32" width="8" height="10" fill="{col}"/>',
        "person": '<circle cx="24" cy="15" r="7" fill="#fff"/>'
                  '<path d="M11 40 Q11 26 24 26 Q37 26 37 40 Z" fill="#fff"/>',
    }.get(kind, "")
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">'
           f'{body}</svg>')
    b64 = base64.b64encode(svg.encode()).decode()
    return "data:image/svg+xml;base64," + b64


def value_chain_view(data):
    """Dash twin of export_native._render_value_chain_native - a creative value-
    chain flow: colour-coded icon nodes on a connector, stage name + value-add %
    around each, a stakeholder card below, and a value-added read-out."""
    stages = data["stages"]
    colors = data.get("colors", ["#1668A8"] * len(stages))
    n = len(stages) or 1
    MX = 40
    step = (1280 - 2 * MX) / n
    node_d, node_cy = 84, 210
    name_y, name_h, pill_y = 118, 38, 264
    card_y, card_h = 298, 250
    centers = [MX + i * step + step / 2 for i in range(n)]
    out = []
    if data.get("subtitle"):
        out.append(_abs(data["subtitle"], 30, 86, 1120, 18,
                        {"fontSize": 12, "color": MUTED}))
    # connector line + arrow glyphs behind the nodes
    out.append(_abs("", centers[0], node_cy - 2, centers[-1] - centers[0], 4,
                    {"background": "#CBD5DE", "borderRadius": 2}))
    for i in range(n - 1):
        mx = (centers[i] + centers[i + 1]) / 2
        out.append(_abs("▶", mx - 10, node_cy - 11, 20, 20,
                        {"color": "#B4C1CD", "fontSize": 15,
                         "textAlign": "center"}))
    for i, st in enumerate(stages):
        cx = centers[i]
        col = colors[i % len(colors)]
        out.append(_abs(
            html.Div([html.Div(ln) for ln in str(st["name"]).split("\n")]),
            cx - step / 2 + 4, name_y, step - 8, name_h,
            {"color": _VC_NAVY, "fontWeight": 700, "fontSize": 12,
             "textAlign": "center", "lineHeight": 1.05, "display": "flex",
             "flexDirection": "column", "justifyContent": "flex-end"}))
        # node: light halo + coloured circle + icon
        icon = html.Img(src=_vc_icon_svg(st.get("icon", ""), col),
                        style={"width": node_d * 0.5, "height": node_d * 0.5})
        inner = html.Div(icon, style={
            "width": node_d, "height": node_d, "borderRadius": "50%",
            "background": col, "display": "flex", "alignItems": "center",
            "justifyContent": "center"})
        out.append(_abs(inner, cx - (node_d + 10) / 2, node_cy - (node_d + 10) / 2,
                        node_d + 10, node_d + 10,
                        {"background": "#EEF3F7", "borderRadius": "50%",
                         "display": "flex", "alignItems": "center",
                         "justifyContent": "center"}))
        out.append(_abs("▲ " + st.get("valueadd", ""), cx - 42, pill_y, 84, 22,
                        {"color": _VC_PURPLE, "fontWeight": 700,
                         "fontSize": 11, "textAlign": "center",
                         "background": "#F0E8F6",
                         "border": f"1px solid {_VC_PURPLE}",
                         "borderRadius": 11, "display": "flex",
                         "alignItems": "center", "justifyContent": "center"}))
        cw = step - 16
        bar = html.Div("", style={"height": 5, "background": col})
        bodyc = html.Div([
            html.Div(st["title"], style={"fontWeight": 700, "color": _VC_NAVY,
                                         "fontSize": 12.5, "lineHeight": 1.1,
                                         "marginBottom": 4}),
            html.Div(data.get("base_year", ""), style={
                "fontSize": 11, "color": MUTED, "marginBottom": 6}),
            html.Div(st["note"], style={"fontSize": 11, "color": TEXT,
                                        "lineHeight": 1.32}),
        ], style={"padding": "9px 10px"})
        out.append(_abs([bar, bodyc], cx - cw / 2, card_y, cw, card_h,
                        {"background": PANEL_BG,
                         "border": f"1px solid {PANEL_EDGE}",
                         "borderRadius": 4, "overflow": "hidden",
                         "boxShadow": "1px 2px 5px rgba(0,40,80,0.10)"}))
    ry = card_y + card_h + 14
    out.append(summary_panel(40, ry, 1200, 704 - ry, "Value-Added Read-out",
                             data.get("readout", [])))
    return out


# ── Template switch support ──────────────────────────────────────────
def _retheme() -> None:
    """Re-derive every module-level constant that bakes a theme token.

    `theming.apply()` rebinds the tokens imported at the top of this module and
    then calls this hook; the header/panel helpers read the tokens live, so only
    the precomputed dicts and tints need refreshing here.
    """
    global _AMBER, _CLASS_MARK, _FLOW_LINE, _FLOW_LEAF_EDGE, _FLOW_LEAF_TEXT
    global _SIG_HEAD_BG, _SIG_EDGE, _SIG_ZEBRA, _TH, _TD, _ARROW_HTML
    global _VC_PURPLE, _VC_NAVY

    _AMBER = BELOW_COLOR
    _CLASS_MARK = {"▲": OUTPACE_COLOR, "▶": BELOW_COLOR,
                   "▼": DECLINE_COLOR}
    _FLOW_LINE, _FLOW_LEAF_EDGE = FLOW_LINE, FLOW_LEAF_EDGE
    _FLOW_LEAF_TEXT = FLOW_LEAF_TEXT
    _SIG_HEAD_BG, _SIG_EDGE, _SIG_ZEBRA = SIG_HEAD_BG, SIG_EDGE, SIG_ZEBRA
    # the header row now carries its own two tokens: a pale band with
    # dark type (FMI) or the brand band with white type (Fact.MR)
    _TH = {**_TH, "background": TABLE_HEAD_BG, "color": TABLE_HEAD_TEXT}
    _TD = {**_TD, "borderBottom": f"1px solid {GRID}", "color": TEXT}
    _ARROW_HTML = {"high": ARROW_HIGH, "med": ARROW_MED, "low": ARROW_LOW}
    _VC_PURPLE, _VC_NAVY = VC_PURPLE, VC_NAVY
