"""Content-aware table layout / pagination engine (v2).

Goal: no internal scrollbars anywhere - neither vertical nor horizontal -
by sizing every table to its actual content and packing the content-sized
tables onto the page, spilling whole tables onto follow-on pages when they
do not fit.  Leftover space is filled with a white placeholder box (a spot
for future per-slide commentary).  A "Source: Future Market Insights" line
sits under every table.

Everything is derived from the data (row counts, label/number widths), so
the same logic adapts automatically to any Excel dataset / market - no
per-table or per-market hard-coding.

  * VERTICAL  - table height = title + header + every row + source line
    (no vertical scroll; the panel grows to the row count).
  * HORIZONTAL - table width = index-column width (longest label) + each
    value column's width (widest value/header), and no wider (no h-scroll,
    no wasted width).
  * PACKING   - content-sized tables are placed into columns; a tall table
    naturally sits beside shorter stacked ones.  A table that does not fit
    the page's table region moves WHOLE onto the next page.
  * WHITE BOX - the remaining blank area becomes a white placeholder panel.
"""
from __future__ import annotations

from . import components
from .config import (CANVAS_H, CANVAS_W, FLOW_LINE as _T_FLOW_LINE,
                     FS_PANEL_TITLE, PAGE_BG, PANEL_BG, PANEL_EDGE,
                     PP_BAND_FILL as _T_PP_BAND_FILL,
                     PP_BAND_LINE as _T_PP_BAND_LINE)

# ── text width model (Segoe UI, normalised to font-size 1) ───────────
# Generated once with canvas.measureText (600 weight); server-side width
# estimation so tables can be sized without a browser at runtime.
_CW = {'0': 0.5552, '1': 0.4023, '2': 0.5552, '3': 0.5552, '4': 0.5762, '5': 0.5552, '6': 0.5581, '7': 0.5361, '8': 0.5552, '9': 0.5581, 'a': 0.522, 'b': 0.6025, 'c': 0.4702, 'd': 0.6025, 'e': 0.5312, 'f': 0.3447, 'g': 0.6025, 'h': 0.582, 'i': 0.2612, 'j': 0.2612, 'k': 0.5249, 'l': 0.2612, 'm': 0.8857, 'n': 0.5835, 'o': 0.5972, 'p': 0.6025, 'q': 0.6025, 'r': 0.3701, 's': 0.4312, 't': 0.3613, 'u': 0.5835, 'v': 0.5073, 'w': 0.7563, 'x': 0.501, 'y': 0.5083, 'z': 0.4644, 'A': 0.6714, 'B': 0.604, 'C': 0.6211, 'D': 0.7173, 'E': 0.5176, 'F': 0.5024, 'G': 0.6973, 'H': 0.7354, 'I': 0.2915, 'J': 0.3965, 'K': 0.6108, 'L': 0.4888, 'M': 0.9243, 'N': 0.7671, 'O': 0.7559, 'P': 0.5845, 'Q': 0.7559, 'R': 0.6226, 'S': 0.5444, 'T': 0.5518, 'U': 0.7031, 'V': 0.6416, 'W': 0.9658, 'X': 0.6191, 'Y': 0.5771, 'Z': 0.5869, '!': 0.3037, '"': 0.4375, '#': 0.5913, '$': 0.5552, '%': 0.8403, '&': 0.7153, "'": 0.2583, '(': 0.332, ')': 0.332, '*': 0.4341, '+': 0.6943, ',': 0.2412, '-': 0.4019, '.': 0.2412, '/': 0.4136, ':': 0.2412, ';': 0.2412, '<': 0.6943, '=': 0.6943, '>': 0.6943, '?': 0.4438, '@': 0.9546, '[': 0.332, '\\': 0.4048, ']': 0.332, '^': 0.6943, '_': 0.415, '`': 0.2886, '{': 0.332, '|': 0.2783, '}': 0.332, '~': 0.6943, ' ': 0.2749}
_DEF_CW = 0.56

def text_width(s, font_size: float) -> float:
    return sum(_CW.get(c, _DEF_CW) for c in str(s)) * font_size

# ── canvas geometry (px) ─────────────────────────────────────────────
CONTENT_TOP = 103           # under the 96px header and its rule (was 90)
CONTENT_BOTTOM = 696          # 10px of clear paper above the copyright line (client, 2026-10-06)
PAGE_LEFT = 15
PAGE_RIGHT = 15
GUTTER_X = 16
GUTTER_Y = 14
CONTENT_W = CANVAS_W - PAGE_LEFT - PAGE_RIGHT

# ── table metrics (must track components.df_table's pinned CSS) ───────
CELL_FS = 10.5              # value / index font size
HEAD_FS = 10.5             # header font (bold ~ treat same size, add pad)
CELL_PAD_X = 18            # per-column horizontal padding + slack
TITLE_H = 28               # panel title strip (one line)
TITLE_FS = FS_PANEL_TITLE  # panel title font size (template token)
TITLE_CHROME = 46          # title padding + focus icon + slack
HEAD_H = 24                # blue header row
ROW_H = 22                 # each data row
PAD_V = 14                 # panel body vertical padding
SOURCE_H = 18              # "Source: ..." line under the table
MIN_TABLE_W = 210
COL_W_SLACK = 6

def table_size(df, index_name="", has_total=False, title=""):
    """Content (width, height) in px for a pivot table - sized so every row
    and column is fully visible with no scrollbar and no wasted width.  The
    width is also wide enough to fit the panel TITLE on a single line."""
    idx_w = max([text_width(index_name, HEAD_FS)]
                + [text_width(i, CELL_FS) for i in df.index] + [0])
    total_w = idx_w + CELL_PAD_X
    for col in df.columns:
        cell_ws = [text_width(col, HEAD_FS)]
        for v in df[col]:
            cell_ws.append(text_width(v[0] if isinstance(v, tuple) else v,
                                      CELL_FS))
        total_w += max(cell_ws) + CELL_PAD_X
    data_w = int(total_w) + COL_W_SLACK
    title_w = int(text_width(title or "", TITLE_FS) * 1.05) + TITLE_CHROME
    w = min(max(data_w, title_w, MIN_TABLE_W), CONTENT_W)
    rows = len(df) + (1 if has_total else 0)
    # a grouped table (df.attrs["col_groups"]) heads its columns in two rows
    heads = 2 if getattr(df, "attrs", {}).get("col_groups") else 1
    h = TITLE_H + HEAD_H * heads + rows * ROW_H + PAD_V + SOURCE_H
    return w, int(h)

# ── block constructors ───────────────────────────────────────────────
# ``frac`` = fraction of the content width a block occupies in a "rows"
# (document-flow) layout; ignored by the table-packer / slides modes.
def tbl(df, index_name="", title=None, total_row=None, body_id=None,
        frac=1.0) -> dict:
    w, h = table_size(df, index_name, total_row is not None, title)
    return {"kind": "table", "df": df, "index_name": index_name,
            "title": title, "total_row": total_row, "body_id": body_id,
            "cw": w, "ch": h, "frac": frac}

def kpi(label, value, sub="", w=205, h=57, frac=0.25) -> dict:
    return {"kind": "kpi", "label": label, "value": value, "sub": sub,
            "cw": w, "ch": h, "frac": frac}

def fig(figure, title=None, graph_id=None, h=280, frac=1.0) -> dict:
    # Every figure on every page goes through here, which makes it the one
    # place to enforce chart-wide rules.  cap_bar_widths keeps a one- or
    # two-category bar chart from drawing its bars as full-width slabs.
    from . import figures as _f
    _f.cap_bar_widths(figure)
    return {"kind": "figure", "figure": figure, "title": title,
            "graph_id": graph_id, "ch": int(h), "frac": frac}

def note(text, h=120, frac=1.0) -> dict:
    return {"kind": "writeup", "text": text, "ch": int(h), "frac": frac}

def summary(title, bullets, frac=1.0, h=None) -> dict:
    """An editorial 'summary panel' block: a titled white card with an accent
    rule and bullet lines.  Used by the Summary-section pages; rendered as
    editable HTML (dashboard) and native text boxes (PPT)."""
    bullets = [str(b) for b in bullets]
    if h is None:
        h = 40 + len(bullets) * 26 + 6           # title + bullets + padding
    return {"kind": "sumpanel", "title": title, "bullets": bullets,
            "ch": int(h), "frac": frac}

def raw(component) -> dict:
    """A pass-through Dash component (e.g. a dcc.Store or button) that carries
    no layout of its own - attached to the current page as-is."""
    return {"kind": "raw", "component": component, "ch": 0, "frac": 0}

def frow(*blocks) -> list:
    """One row of a document-flow ('rows') layout - blocks placed side by
    side, their ``frac`` widths summing to <= 1."""
    return [b for b in blocks if b is not None]

# ── packing ──────────────────────────────────────────────────────────
MAX_COLS = 2      # tables sit in at most 2 content-width columns per page
# Fraction of the vertical table-region a column may fill before the next
# table is pushed to a new page - the rest of the page is reserved for the
# white commentary box.  Tunable; higher = denser pages, less blank space.
FILL = 0.5

def _pack_tables(tables, top):
    """Pack content-sized tables into up to ``MAX_COLS`` columns within the
    page, spilling whole tables onto follow-on pages.

    Columns are filled breadth-first (each new column opened beside the last
    while there is horizontal room and we are under MAX_COLS) so a tall table
    sits *beside* stacked shorter ones rather than under them; once the
    columns exist, further tables stack into the shortest column.  A table
    whose target column has already passed the fill limit starts a new page -
    reserving the lower part of the page for the white commentary box.  Never
    splits a table.
    """
    avail = CONTENT_BOTTOM - top
    limit = top + FILL * avail
    pages, cur, cols = [], [], []

    def flush():
        nonlocal cur, cols
        if cur:
            pages.append(cur)
        cur, cols = [], []

    for t in tables:
        w, h = t["cw"], t["ch"]
        used_right = max((c["x"] + c["w"] for c in cols), default=None)
        nx = PAGE_LEFT if used_right is None else used_right + GUTTER_X
        can_open = (len(cols) < MAX_COLS
                    and nx + w <= PAGE_LEFT + CONTENT_W
                    and top + h <= CONTENT_BOTTOM)
        if can_open:                              # breadth-first: new column
            cols.append({"x": nx, "y": top + h + GUTTER_Y, "w": w})
            cur.append({**t, "x": nx, "y": int(top), "w": w, "h": h})
            continue
        # otherwise stack into the shortest column that still has room
        cand = min(cols, key=lambda c: c["y"]) if cols else None
        if cand and cand["y"] <= limit and cand["y"] + h <= CONTENT_BOTTOM:
            cur.append({**t, "x": cand["x"], "y": int(cand["y"]),
                        "w": w, "h": h})
            cand["y"] += h + GUTTER_Y
            cand["w"] = max(cand["w"], w)
            continue
        # no room on this page -> spill whole table to a fresh page
        flush()
        cur.append({**t, "x": PAGE_LEFT, "y": int(top), "w": w, "h": h})
        cols = [{"x": PAGE_LEFT, "y": top + h + GUTTER_Y, "w": w}]
    if cur:
        pages.append(cur)
    return pages

def _equalize_widths(tblocks):
    """Give every table column on a slide the SAME width so the tables line
    up neatly.  The shared width is the widest column's content width (so no
    table is ever narrowed below its content - no horizontal scroll), capped
    to what fits the page.  Columns are then re-laid at that width."""
    if not tblocks:
        return tblocks
    xs = sorted({b["x"] for b in tblocks})
    ncols = len(xs)
    want = max(b["w"] for b in tblocks)
    cap = (CONTENT_W - (ncols - 1) * GUTTER_X) / ncols
    uni = int(min(want, cap))
    # a single column of wide (time-series) tables spans the FULL page width
    # so every column fits and no white box sits alongside / clips it
    if ncols == 1 and uni >= 0.5 * CONTENT_W:
        uni = CONTENT_W
    remap = {x: int(round(PAGE_LEFT + j * (uni + GUTTER_X)))
             for j, x in enumerate(xs)}
    for b in tblocks:
        b["x"] = remap[b["x"]]
        b["w"] = uni
    return tblocks

def _justify_columns(tblocks, top):
    """Stretch the tables in each column so they fill the page height and are
    the SAME size (equal width + height) within the column - only the panel
    (block/border) grows; the font and rows are unchanged, leaving blank
    space below the rows.  Never shrinks a table below its content."""
    cols: dict = {}
    for b in tblocks:
        cols.setdefault(b["x"], []).append(b)
    for group in cols.values():
        group.sort(key=lambda b: b["y"])
        n = len(group)
        colw = max(b["w"] for b in group)
        maxc = max(b["h"] for b in group)
        avail = CONTENT_BOTTOM - top - (n - 1) * GUTTER_Y
        each_h = avail / n
        if each_h < maxc:                # not enough room to stretch equally
            for b in group:              # keep natural heights, equalise width
                b["w"] = colw
            continue
        y = top
        for b in group:
            b["y"] = int(round(y))
            b["h"] = int(round(each_h))
            b["w"] = colw
            y += each_h + GUTTER_Y
    return tblocks

def _white_boxes(blocks):
    """Fill the page's blank area with white placeholder panel(s).

    Simple L-cover: one box to the right of the table cluster and one below
    it - enough to reserve the empty space for future commentary."""
    if not blocks:
        return []
    right = max(b["x"] + b["w"] for b in blocks)
    bottom = max(b["y"] + b["h"] for b in blocks)
    left = min(b["x"] for b in blocks)
    top = min(b["y"] for b in blocks)
    boxes = []
    # right region
    rx = right + GUTTER_X
    if PAGE_LEFT + CONTENT_W - rx >= 120:
        boxes.append({"kind": "whitebox", "x": rx, "y": top,
                      "w": PAGE_LEFT + CONTENT_W - rx,
                      "h": CONTENT_BOTTOM - top})
    # bottom region (under the tables, left of the right box)
    by = bottom + GUTTER_Y
    if CONTENT_BOTTOM - by >= 90:
        boxes.append({"kind": "whitebox", "x": left, "y": by,
                      "w": right - left, "h": CONTENT_BOTTOM - by})
    return boxes

def _block_h(b, w):
    if b["kind"] == "table":
        # re-measure height for the (possibly narrower) flow width in case
        # the title now wraps onto a second line
        _, h = table_size(b["df"], b["index_name"],
                          b.get("total_row") is not None, b.get("title"))
        title_w = int(text_width(b.get("title") or "", TITLE_FS) * 1.05) \
            + TITLE_CHROME
        if title_w > w:                        # wrapped title -> taller strip
            h += TITLE_H
        return h
    return int(b.get("ch", 0))

def paginate_rows(rows: list) -> list[dict]:
    """Document-flow layout for mixed chart/table/writeup pages: each row is a
    band of blocks placed side by side (``frac`` widths), tables sized to
    their rows (no scroll), charts/writeups at their given height.  Rows flow
    top-to-bottom and spill WHOLE rows onto follow-on slides.  ``raw`` blocks
    (stores/buttons) ride along on the current page with no layout."""
    pages, cur, y = [], [], CONTENT_TOP

    def has_real():
        return any(b["kind"] != "raw" for b in cur)

    def flush():
        nonlocal cur, y
        if cur:
            reals = [b for b in cur if b["kind"] != "raw"]
            if reals:
                # equalise heights within each row and stretch rows to fill
                # the page (charts/tables grow, bottoms align) - 100% height
                row_ys = sorted({b["y"] for b in reals})
                groups = [[b for b in reals if b["y"] == ry] for ry in row_ys]
                row_hs = [max(b["h"] for b in g) for g in groups]
                used = sum(row_hs) + (len(groups) - 1) * GUTTER_Y
                add = max(0, (CONTENT_BOTTOM - CONTENT_TOP - used)) \
                    / len(groups)
                ny = CONTENT_TOP
                for g, rh in zip(groups, row_hs):
                    nh = int(rh + add)
                    for b in g:
                        b["y"] = int(ny)
                        b["h"] = nh
                    ny += nh + GUTTER_Y
            bottom = max((b["y"] + b["h"] for b in reals),
                         default=CONTENT_BOTTOM)
            pages.append({"blocks": cur,
                          "height": max(CANVAS_H,
                                        bottom + (CANVAS_H - CONTENT_BOTTOM))})
        cur, y = [], CONTENT_TOP

    for row in rows:
        cur.extend(b for b in row if b["kind"] == "raw")
        real = [b for b in row if b["kind"] != "raw"]
        if not real:
            continue
        n = len(real)
        gutters = (n - 1) * GUTTER_X
        avail_w = CONTENT_W - gutters
        widths = [max(120, int(b.get("frac", 1.0 / n) * avail_w))
                  for b in real]
        heights = [_block_h(b, w) for b, w in zip(real, widths)]
        rh = max(heights)
        if has_real() and y + rh > CONTENT_BOTTOM:
            flush()
        x = PAGE_LEFT
        for b, w, h in zip(real, widths, heights):
            cur.append({**b, "x": int(round(x)), "y": int(round(y)),
                        "w": int(w), "h": int(h)})
            x += w + GUTTER_X
        y += rh + GUTTER_Y
    flush()
    return pages

def paginate(spec: dict) -> list[dict]:
    """spec = {"fixed": [kpi blocks on page 1], "tables": [table blocks]}.

    Returns physical pages: {"blocks": [...placed...], "height": canvas_h}.
    Page 1 keeps the fixed (KPI) band at the top; tables pack below it and
    spill onto follow-on pages; each page gets white placeholder box(es).
    """
    if "pages" in spec:                         # pre-placed blocks, one list
        return [{"blocks": list(p), "height": CANVAS_H}   # per slide
                for p in (spec["pages"] or [[]])]
    if "rows" in spec:
        return paginate_rows(spec["rows"])
    if "positioned" in spec:                    # explicit x/y/w/h blocks, 1 page
        return [{"blocks": list(spec["positioned"]), "height": CANVAS_H}]
    if "slides" in spec:                        # pre-positioned blocks/slide
        out = []
        for s in (spec["slides"] or [[]]):
            bottom = max((b.get("y", 0) + b.get("h", 0) for b in s),
                         default=CONTENT_BOTTOM)
            out.append({"blocks": s,
                        "height": max(CANVAS_H,
                                      bottom + (CANVAS_H - CONTENT_BOTTOM))})
        return out
    if "flow" in spec:                          # taxonomy flow chart / slide
        fslides = spec["flow"] or [{"nodes": [], "edges": []}]
        return [{"blocks": [{"kind": "flowchart", "nodes": f["nodes"],
                             "edges": f["edges"]}], "height": CANVAS_H}
                for f in fslides]
    if "signals" in spec:                       # market-signals editorial
        return [{"blocks": [{"kind": "signals", "signals": spec["signals"]}],
                 "height": CANVAS_H}]
    if "endcards" in spec:                       # segment-section closing card
        return [{"blocks": [{"kind": "endcards", "data": spec["endcards"]}],
                 "height": CANVAS_H}]
    if "segstart" in spec:                        # segment-section opening page
        return [{"blocks": [{"kind": "segstart", "data": spec["segstart"]}],
                 "height": CANVAS_H}]
    if "pricepattern" in spec:                    # per-region pricing-pattern page
        return [{"blocks": [{"kind": "pricepattern", "data": pdata}],
                 "height": CANVAS_H}
                for pdata in pricepattern_pages(spec["pricepattern"])]
    if "primitives" in spec:                      # generic primitive-drawn page
        return [{"blocks": [{"kind": "primitives",
                             "prims": spec["primitives"]}], "height": CANVAS_H}]
    if "primpages" in spec:                       # primitive pages, one per slide
        return [{"blocks": [{"kind": "primitives", "prims": p}],
                 "height": CANVAS_H} for p in (spec["primpages"] or [[]])]
    if "compmatrix" in spec:                      # creative competitor matrix
        return [{"blocks": [{"kind": "compmatrix",
                             "data": spec["compmatrix"]}], "height": CANVAS_H}]
    if "valuechain" in spec:                      # value-chain chevron flow
        return [{"blocks": [{"kind": "valuechain",
                             "data": spec["valuechain"]}], "height": CANVAS_H}]
    fixed = spec.get("fixed", [])
    tables = spec.get("tables", [])
    # place the fixed KPI band as an even row spanning the FULL content
    # width (the cards stretch to cover the whole line)
    fixed_placed = []
    if fixed:
        n = len(fixed)
        fw = (CONTENT_W - (n - 1) * GUTTER_X) / n
        x = PAGE_LEFT
        for f in fixed:
            fixed_placed.append({**f, "x": int(round(x)), "y": CONTENT_TOP,
                                 "w": int(round(fw)), "h": f["ch"]})
            x += fw + GUTTER_X
    top1 = CONTENT_TOP + (max(f["ch"] for f in fixed) + GUTTER_Y
                          if fixed else 0)

    tpages = _pack_tables(tables, top1) or [[]]
    pages = []
    for i, tblocks in enumerate(tpages):
        # equalise column widths, then stretch each column's tables to fill
        # the page height - so every table is a uniform, aligned block
        tblocks = _equalize_widths(tblocks)
        tblocks = _justify_columns(tblocks, top1)
        # the KPI band is the page template - repeat it on EVERY slide
        blocks = list(fixed_placed) + tblocks
        blocks += _white_boxes(tblocks)
        bottom = max((b["y"] + b["h"] for b in blocks), default=CONTENT_BOTTOM)
        pages.append({"blocks": blocks,
                      "height": max(CANVAS_H,
                                    bottom + (CANVAS_H - CONTENT_BOTTOM))})
    return pages

# ── rendering: placed blocks -> Dash components ──────────────────────
def _components(blocks):
    from dash import html
    out = []
    for b in blocks:
        k = b["kind"]
        if k == "raw":                          # pass-through (store/button)
            out.append(b["component"])
            continue
        if k == "flowchart":                    # taxonomy org chart (own coords)
            out.extend(components.flow_chart(b["nodes"], b["edges"]))
            continue
        if k == "signals":                      # market-signals table (own coords)
            out.extend(components.signals_page(b["signals"]))
            continue
        if k == "endcards":                     # segment closing card (own coords)
            out.extend(components.end_cards(b["data"]))
            continue
        if k == "segstart":                     # segment opening page (own coords)
            out.extend(components.seg_start_view(b["data"]))
            continue
        if k == "pricepattern":                 # per-region pricing pattern page
            out.extend(components.price_pattern_view(b["data"]))
            continue
        if k == "primitives":                   # generic primitive-drawn page
            out.extend(components.primitives_view(b["prims"]))
            continue
        if k == "compmatrix":                   # creative competitor matrix
            out.extend(components.competitor_matrix_view(b["data"]))
            continue
        if k == "valuechain":                   # value-chain chevron flow
            out.extend(components.value_chain_view(b["data"]))
            continue
        x, y, w, h = b["x"], b["y"], b["w"], b["h"]
        if k == "table":
            out.append(components.table_panel(
                x, y, w, h, b["df"], index_name=b["index_name"],
                title=b["title"], total_row=b.get("total_row"),
                body_id=b.get("body_id"), scroll=False, source=True))
        elif k == "figure":
            out.append(components.graph_panel(
                x, y, w, h, b["figure"], title=b.get("title"),
                graph_id=b.get("graph_id")))
        elif k == "writeup":
            out.append(components.writeup(x, y, w, h, b["text"]))
        elif k == "kpi":
            out.append(components.kpi_card(x, y, w, h, b["label"],
                                           b["value"], b.get("sub", "")))
        elif k == "sumpanel":                   # editorial summary card
            out.append(components.summary_panel(
                x, y, w, h, b["title"], b["bullets"]))
        elif k == "panel":                      # generic titled panel (trees)
            out.append(components.panel(
                x, y, w, h, b["content"], title=b.get("title"),
                export_lines=b.get("export_lines")))
        elif k == "whitebox":
            out.append(html.Div(style={
                "position": "absolute", "left": x, "top": y, "width": w,
                "height": h, "background": PANEL_BG,
                "border": f"1px solid {PANEL_EDGE}", "borderRadius": 3,
                "boxShadow": "1px 2px 5px rgba(0,0,0,0.12)"}))
    return out

def page_body(page):
    """Blocks for one physical page, placed inside a canvas that already owns
    the header band with the slicers."""
    return _components(page["blocks"])

def render_subpage(spec: dict, p: int):
    """Render ONE physical page (1-indexed ``p``) of a flow spec as a single
    slide body - so each spilled page is a separate navigable slide rather
    than being stacked below the first.  Returns (body, total_pages, p).

    A spec may be a table spec ({"fixed", "tables"}), a document-flow spec
    ({"rows": ...}), or {"slides": [[blocks], ...]} (taxonomy) - all handled by
    paginate()."""
    try:
        p = int(p)
    except (TypeError, ValueError):
        p = 1
    pages = paginate(spec)
    k = len(pages)
    p = max(1, min(p, k))
    return page_body(pages[p - 1]), k, p

def page_count(spec: dict) -> int:
    if "slides" in spec:
        return max(1, len(spec["slides"]))
    if "flow" in spec:
        return max(1, len(spec["flow"]))
    if "signals" in spec:
        return 1
    if "endcards" in spec:
        return 1
    if "segstart" in spec:
        return 1
    if "pricepattern" in spec:
        return len(pricepattern_pages(spec["pricepattern"]))
    if "primitives" in spec:
        return 1
    if "primpages" in spec:
        return max(1, len(spec["primpages"]))
    if "compmatrix" in spec:
        return 1
    if "valuechain" in spec:
        return 1
    if "positioned" in spec:
        return 1
    return len(paginate(spec))

# ── tree splitting (taxonomy trees -> content-fit slices) ────────────
TREE_PARENT_H = 27         # a '▸ parent' row (incl. top margin)
TREE_CHILD_H = 21          # a '☐ child' row
TREE_TITLE_PAD = 42        # panel title strip + body padding

def split_tree(tree: dict, max_content_h: int) -> list[dict]:
    """Split an ordered {parent: [children]} tree into consecutive slices,
    each whose rendered height fits ``max_content_h`` - so a long taxonomy
    tree flows across slides instead of scrolling.  Whole parent-groups are
    kept together."""
    slices, cur, h = [], {}, 0
    for parent, kids in tree.items():
        gh = TREE_PARENT_H + len(kids) * TREE_CHILD_H
        if cur and h + gh > max_content_h:
            slices.append(cur)
            cur, h = {}, 0
        cur[parent] = kids
        h += gh
    if cur:
        slices.append(cur)
    return slices or [{}]

def tree_panel_height(tree_slice: dict) -> int:
    h = sum(TREE_PARENT_H + len(k) * TREE_CHILD_H
            for k in tree_slice.values())
    return int(h + TREE_TITLE_PAD)

# ── Pricing-pattern page (reference "Global Pricing Assessment") ─────
# One SEMANTIC geometry pass turns the page data into a flat list of drawing
# PRIMITIVES (text / rect / line / dot / poly).  Both the dashboard component
# (components.price_pattern_view) and the native PPT renderer
# (export_native._render_pricepattern_native) consume this SAME list, so the
# two surfaces are guaranteed to match - exactly like layout.taxonomy_flow.
_PP_BAND_FILL = _T_PP_BAND_FILL
_PP_BAND_LINE = _T_PP_BAND_LINE
_PP_TOP = 130                 # first region row top (below subtitle/legend)
_PP_BOTTOM = 684             # last usable y for rows (above footer captions)
_PP_ROW_H = 138              # MINIMUM per-region row height
_PP_PER_PAGE = max(1, int((_PP_BOTTOM - _PP_TOP) // _PP_ROW_H))   # ≈ 4 rows/page
# (2026-09-11) every label on the page is set at the 11pt content standard.
# The applications are listed one per line at that size, so a market with
# many of them gets taller rows - and fewer rows per page - rather than type
# too small to read.
_PP_APP_LH = 20              # one application line at 11pt

def _pp_row_h(data: dict) -> int:
    """Per-region row height for this pricing page set: the fixed minimum,
    or taller when a region lists more applications than it holds."""
    m = max((len(r.get("apps", [])) for r in data.get("rows", [])), default=0)
    return max(_PP_ROW_H, 30 + m * _PP_APP_LH + 8)

def _pp_tick(v, span):
    """A tick with enough decimals to say something.

    A fixed 1-dp printed every tick as "0.0" on a market whose derived price
    is in the thousandths (Value in US$ Mn over Volume in MT), so the whole
    number line read as zeros.  The decimals follow the SPAN of the axis, so
    a wide axis stays clean and a narrow one keeps its detail."""
    if span >= 20:
        return f"{v:,.0f}"
    import math
    d = 1 if span >= 2 else max(1, 2 - int(math.floor(math.log10(abs(span)))))
    return f"{v:,.{min(d, 6)}f}"

def pricepattern_pages(data: dict) -> list:
    """Split the region rows into fixed-size pages (``_PP_PER_PAGE`` rows each)
    so a market with more regions than fit on one page SPILLS onto follow-on
    pages with the SAME formatting.  Returns a list of per-page data dicts."""
    rows = data.get("rows", [])
    pp = max(1, int((_PP_BOTTOM - _PP_TOP) // _pp_row_h(data)))
    chunks = [rows[i:i + pp] for i in range(0, len(rows), pp)] or [[]]
    k = len(chunks)
    return [{**data, "rows": ch, "_page": (j + 1, k)}
            for j, ch in enumerate(chunks)]

def pricepattern_primitives(data: dict) -> list:
    """Flatten the pricing-pattern page into drawing primitives (px on the
    1280×720 canvas).  Primitive kinds:
      text {x,y,w,h,s,size,bold,color,align('l'|'c'|'r'),underline?}
      rect {x,y,w,h,fill,line}
      line {x1,y1,x2,y2,color,width}
      dot  {cx,cy,d,color,edge}
      poly {pts:[(x,y)...],color,width}
    """
    from .config import TEXT, MUTED, HEADER_BG
    P = []
    prods = data["products"]                       # [{name,color}] (≤3)
    rows = data["rows"]
    unit = data.get("unit", "")
    n = max(1, len(rows))
    x_min, x_max = float(data["x_min"]), float(data["x_max"])
    span = (x_max - x_min) or 1.0
    # column geometry
    TR_L, TR_R = 15, 300                            # trend column
    NL_L, NL_R = 348, 690                           # number-line column
    APP_X, APP_R = 705, 1265                        # application column

    def sx(v):
        return NL_L + (v - x_min) / span * (NL_R - NL_L)

    # ── subtitle (left) + application header (right), legend under them ──
    # (2026-09-11) at the 11pt content standard the legend would run into
    # the application header, so it now has a line of its own
    P.append({"t": "text", "x": TR_L, "y": 84, "w": NL_R - TR_L, "h": 20,
              "s": data["subtitle"], "size": 11.5, "bold": True,
              "color": HEADER_BG, "align": "l"})
    P.append({"t": "text", "x": APP_X, "y": 84, "w": APP_R - APP_X, "h": 20,
              "s": data["app_header"], "size": 11, "bold": True,
              "color": HEADER_BG, "align": "l"})
    lx, ly = TR_L, 106
    for p in prods:
        tw = text_width(p["name"], 11 * 1.35)      # 11pt, in canvas px
        P.append({"t": "dot", "cx": lx + 6, "cy": ly + 10, "d": 11,
                  "color": p["color"], "edge": "#FFFFFF"})
        P.append({"t": "text", "x": lx + 15, "y": ly, "w": tw + 8, "h": 20,
                  "s": p["name"], "size": 11, "bold": False, "color": TEXT,
                  "align": "l"})
        lx += 15 + tw + 26
    P.append({"t": "rect", "x": lx, "y": ly + 4, "w": 22, "h": 12,
              "fill": _PP_BAND_FILL, "line": _PP_BAND_LINE})
    P.append({"t": "text", "x": lx + 26, "y": ly, "w": 120, "h": 20,
              "s": "Price Range", "size": 11, "bold": False, "color": TEXT,
              "align": "l"})

    # ── per-region rows (one height per page set, so its pages match) ──
    # A page set spread over several slides is numbered in the slide heading
    # ("(1/2)", "(2/2)" - see export_native._pricepattern_slides).
    top0 = _PP_TOP
    rh = _pp_row_h(data)
    for i, row in enumerate(rows):
        ry = top0 + i * rh
        # zebra row background
        P.append({"t": "rect", "x": TR_L, "y": ry + 2, "w": APP_R - TR_L,
                  "h": rh - 6, "fill": "#F5F7F9" if i % 2 == 0 else "#FFFFFF",
                  "line": "#EDF1F4"})
        P.append({"t": "text", "x": TR_L + 6, "y": ry + 4, "w": 330, "h": 20,
                  "s": f"{row['region']} Pricing Pattern", "size": 11,
                  "bold": True, "color": HEADER_BG, "align": "l"})
        body_top = ry + 26
        body_h = rh - 32
        # ── trend sparkline ──
        years = row["years"]
        series = row["trend"]
        tx0, tx1 = TR_L + 46, TR_R - 10
        ty0, ty1 = body_top + 10, body_top + body_h - 20
        allv = [v for s in series for v in s] or [0.0, 1.0]
        # The trend is a % change against the first year, so it goes NEGATIVE
        # whenever a price falls.  The old scale pinned zero to the bottom of
        # the band and divided by ymax alone, which sent a falling series
        # BELOW the axis and out of its row - and inverted the whole plot when
        # every value was negative (ymax itself < 0).  The band is now sized
        # from the real range, with zero kept inside it: a rising series still
        # sits on a zero floor, a falling one hangs from a zero ceiling, and a
        # mixed one gets the baseline wherever it belongs.
        ylo, yhi = min(allv), max(allv)
        ylo, yhi = min(ylo, 0.0), max(yhi, 0.0)
        if yhi - ylo < 1e-9:                      # a perfectly flat series
            ylo, yhi = -1.0, 1.0
        pad = (yhi - ylo) * 0.08
        ylo, yhi = ylo - pad, yhi + pad

        def _ty(v, ty0=ty0, ty1=ty1, ylo=ylo, yhi=yhi):
            return ty1 - (v - ylo) / (yhi - ylo) * (ty1 - ty0)

        def _tx(k, tx0=tx0, tx1=tx1, m=len(years)):
            return tx0 + (k / (m - 1)) * (tx1 - tx0) if m > 1 else tx0
        # the ZERO line, drawn where zero actually falls in the band - this is
        # the axis the lines are read against, so it moves up when a series
        # spends its time below zero
        zero_y = _ty(0.0)
        P.append({"t": "line", "x1": tx0, "y1": zero_y, "x2": tx1, "y2": zero_y,
                  "color": "#C9D2DA", "width": 0.8})
        for val in (yhi, ylo):
            P.append({"t": "text", "x": TR_L - 1, "y": _ty(val) - 9, "w": 44,
                      "h": 18, "s": f"{val:+.0f}%", "size": 11, "bold": False,
                      "color": MUTED, "align": "r"})
        for pi, s in enumerate(series):
            pts = [(_tx(k), _ty(v)) for k, v in enumerate(s)]
            P.append({"t": "poly", "pts": pts, "color": prods[pi]["color"],
                      "width": 1.6})
            # invisible, dashboard-only hover targets at each trend vertex
            for k, v in enumerate(s):
                P.append({"t": "dot", "cx": _tx(k), "cy": _ty(v), "d": 12,
                          "color": "transparent", "edge": "transparent",
                          "web_only": True,
                          "tip": f"{prods[pi]['name']} · {years[k]}: "
                                 f"{v:+.1f}% vs {years[0]}"})
        P.append({"t": "text", "x": tx0 - 22, "y": ty1 + 2, "w": 48, "h": 18,
                  "s": str(years[0]), "size": 11, "bold": False,
                  "color": MUTED, "align": "c"})
        P.append({"t": "text", "x": tx1 - 26, "y": ty1 + 2, "w": 48, "h": 18,
                  "s": str(years[-1]), "size": 11, "bold": False,
                  "color": MUTED, "align": "c"})
        # ── price number-line ──
        mid_y = body_top + body_h * 0.56
        P.append({"t": "line", "x1": NL_L, "y1": mid_y, "x2": NL_R, "y2": mid_y,
                  "color": MUTED, "width": 1.0})
        P.append({"t": "text", "x": NL_L - 18, "y": mid_y - 9, "w": 14, "h": 18,
                  "s": "A", "size": 11, "bold": True, "color": TEXT,
                  "align": "c"})
        for k in range(6):
            v = x_min + k / 5.0 * span
            gx = sx(v)
            P.append({"t": "line", "x1": gx, "y1": mid_y - 4, "x2": gx,
                      "y2": mid_y + 4, "color": "#6B6B6B", "width": 0.8})
            P.append({"t": "text", "x": gx - 32, "y": mid_y - 30, "w": 64,
                      "h": 18, "s": _pp_tick(v, span), "size": 11,
                      "bold": False, "color": MUTED, "align": "c"})
        lo, hi = row["band"]
        P.append({"t": "rect", "x": sx(lo), "y": mid_y - 8,
                  "w": max(3.0, sx(hi) - sx(lo)), "h": 16,
                  "fill": _PP_BAND_FILL, "line": _PP_BAND_LINE,
                  "tip": f"{row['region']} price range: "
                         f"{lo:,.2f}–{hi:,.2f} {unit}"})
        for pi, v in enumerate(row["dots"]):
            P.append({"t": "dot", "cx": sx(v), "cy": mid_y, "d": 12,
                      "color": prods[pi]["color"], "edge": "#FFFFFF",
                      "tip": f"{prods[pi]['name']}: {v:,.2f} {unit}"})
        # ── applications: one per line - the name, then each product's value
        # beside that product's colour dot (2026-09-11: the old two-column
        # grid only worked at 7pt; at the 11pt content standard each value
        # needs its own column) ──
        apps = row["apps"]
        m = max(1, len(apps))
        name_w = 236.0
        vx0 = APP_X + name_w
        vw = (APP_R - vx0) / max(1, len(prods))
        lh = min(36.0, body_h / m)
        ay0 = body_top + max(0.0, (body_h - lh * m) / 2)
        for ai, (aname, vals) in enumerate(apps):
            ay = ay0 + ai * lh
            atip = f"{aname} " + " · ".join(
                f"{prods[pi]['name']} {vals[pi]:,.2f} {unit}"
                for pi in range(min(len(vals), len(prods))))
            P.append({"t": "text", "x": APP_X, "y": ay, "w": name_w - 8,
                      "h": lh, "s": aname, "size": 11, "bold": True,
                      "color": HEADER_BG, "align": "l", "tip": atip})
            for pi, v in enumerate(vals[:len(prods)]):
                vx = vx0 + pi * vw
                P.append({"t": "dot", "cx": vx + 5, "cy": ay + lh / 2, "d": 8,
                          "color": prods[pi]["color"], "edge": "#FFFFFF"})
                P.append({"t": "text", "x": vx + 13, "y": ay, "w": vw - 14,
                          "h": lh, "s": f"{v:,.2f}", "size": 11,
                          "bold": False, "color": TEXT, "align": "l",
                          "tip": atip})
    # ── footer captions ──
    P.append({"t": "text", "x": NL_L, "y": 688, "w": 200, "h": 18,
              "s": "Price By Product", "size": 11, "bold": True,
              "color": HEADER_BG, "align": "l", "underline": True})
    P.append({"t": "text", "x": APP_X, "y": 688, "w": 220, "h": 18,
              "s": "Price By Application", "size": 11, "bold": True,
              "color": HEADER_BG, "align": "l", "underline": True})
    return P

# ── Taxonomy flow chart (org-chart geometry) ─────────────────────────
# A pure-geometry builder shared by the dashboard AND the PowerPoint export
# so both render byte-for-byte the same hierarchy diagram.  It lays a
# three-level tree - root → branches → leaves - as a top-down org chart:
# the root centred on top, each branch a column beneath it, and every leaf
# stacked under its branch with an elbow connector.  Everything (column
# widths, box heights from wrapped text, the leaf font that makes the
# tallest column fit) is derived from the data, so it adapts to any market.
FLOW_X0 = PAGE_LEFT                       # 15
FLOW_W = CONTENT_W                        # 1250
FLOW_Y0 = 98                              # below the header/title band
FLOW_YB = 706                             # bottom margin of the canvas
FLOW_CG = 12                              # gutter between branch columns
FLOW_IND = 14                             # leaf indent inside its column
# The gutter and indent are deliberately tight: every px they give back goes
# into the leaf boxes, which is what lets a 10-column page still hold a word
# like "Netherlands" on one line at a readable size instead of hard-wrapping
# it to "Netherland/s" in the deck.
FLOW_LG = 9                               # vertical gap between leaves
FLOW_GAP1 = 42                            # root row -> branch row gap
FLOW_GAP2 = 20                            # branch box -> first leaf gap
FLOW_MINCOL = 100                         # min branch-column width
# 100px still holds a 9pt branch header and a ~78px leaf box, and it is what
# lets a 10-region workbook keep every region on ONE page: at the old 150 the
# page capped out at 7 columns, so 10 regions spilled 7 + 3 and the second
# slide held three lonely branches.  Splitting is a last resort, not a default.
FLOW_MINLEAF = 9.0                        # smallest READABLE leaf font size
FLOW_LINE = _T_FLOW_LINE                  # connector colour
# ── How WIDE a leaf box is allowed to get ────────────────────────────
# A branch column is the canvas split by the branch count, so a page with ONE
# branch (a region-scoped deck: "Europe" over its countries) handed every box
# the full 1250px.  A 1236px-wide box holding the word "Germany" reads as a
# banner, not as a node in a taxonomy, and stacking 17 of them one per row also
# forced a needless second slide.
#
# So a leaf never grows past FLOW_LEAF_PREF: past that the leaves wrap into
# SUB-COLUMNS inside their branch, and the resulting block is centred under the
# branch.  Two knock-on effects, both wanted: the boxes look like boxes, and a
# branch holds ``rows x sub-columns`` leaves instead of ``rows``, so trees that
# used to spill onto a continuation slide now fit on one.
#
# Nothing changes on a page whose columns are already narrower than this - a
# 6-or-more-branch workbook lays out exactly as it did before.
FLOW_LEAF_PREF = 210                      # widest a leaf box may be
FLOW_MINBRANCH = 260                      # min width of a centred branch header
# ── How far APART the sub-columns sit ────────────────────────────────
# A fixed 12px gutter left two 210px columns almost touching: the tree read as
# one dense slab rather than as separate limbs, and the block used a third of
# the canvas while the rest sat empty.  The gutter is now derived from the BOX
# width, so it scales with whatever size the auto-fit picked, and is then capped
# by the room the branch column actually has.  That makes it self-correcting at
# any column count: 2 or 3 columns get the full preferred gutter, and once the
# columns start crowding the page the gutter closes toward the minimum instead
# of overflowing.
FLOW_SUBGAP = 18                          # MINIMUM gutter between sub-columns
FLOW_SUBGAP_K = 0.55                      # preferred gutter, as a fraction of
                                          # the leaf width (210px -> 115px)
# The gutter also has to HOLD the connectors: with inward-facing spines (see
# flow_edges) two adjacent columns put a spine each into the gutter between
# them, so a tight gutter would crowd the lines as well as the boxes.
# Leaf font sizes tried by the auto-fit, largest first.  Anything below
# FLOW_MINLEAF is a last resort used only when a single leaf cannot be made to
# fit a page on its own - normally the column spills to a continuation slide
# instead of shrinking past legibility (see taxonomy_flow).
FLOW_LEAF_SIZES = (11.0, 10.5, 10.0, 9.5, 9.0, 8.5, 8.0, 7.5, 7.0)

def _wrap_count(text: str, max_w: float, fs: float) -> int:
    """Estimate how many lines ``text`` wraps to inside ``max_w`` px at font
    size ``fs`` - greedy word wrap using the same width model as the tables,
    with a char-level fallback for a single word wider than the box."""
    words = str(text).split()
    if not words:
        return 1
    space = text_width(" ", fs)
    lines, cur = 1, 0.0

    def add_long(ww):                     # a word wider than the box hard-wraps
        nonlocal lines
        extra = int(ww // max_w) if max_w > 0 else 0
        lines += extra
        return ww - extra * max_w

    for wd in words:
        ww = text_width(wd, fs)
        if cur <= 0:
            cur = ww if ww <= max_w else add_long(ww)
        elif cur + space + ww <= max_w:
            cur += space + ww
        else:
            lines += 1
            cur = ww if ww <= max_w else add_long(ww)
    return max(1, lines)

# A flow box's font size is CSS px in the browser but POINTS in PowerPoint, so
# the SAME label renders ~1.35x wider in the deck.  Wrapping is measured against
# that larger figure, so a size that "fits" fits BOTH surfaces: the deck stops
# hard-wrapping words mid-way ("Netherland/s") and the dashboard simply gets a
# little extra breathing room.  Same convention templates_dynamics budgets with.
FLOW_PPT_K = 1.35

def _flow_box_h(text: str, w: float, fs: float, pad: int = 14) -> int:
    n = _wrap_count(text, w - 12, fs * FLOW_PPT_K)
    return int(n * (fs * 1.34 + 2) + pad)

def _flow_geom(root_label: str, chunk: list) -> dict:
    """The parts of a flow slide's geometry that depend only on the root label
    and which BRANCHES share the page - not on the leaves.  Shared by
    ``_flow_slide`` and ``_leaf_capacity`` so the capacity check and the slide
    it sizes can never disagree."""
    c = max(1, len(chunk))
    col_w = (FLOW_W - (c - 1) * FLOW_CG) / c
    branch_fs = 11.0 if col_w >= 170 else (10.0 if col_w >= 140 else 9.0)

    # root box - centred, sized to its (wrapped) label
    root_fs = 15.0 if len(str(root_label)) <= 44 else 13.0
    rw = int(min(660, max(300, text_width(root_label, root_fs) + 46)))
    root_h = _flow_box_h(root_label, rw, root_fs, pad=16)

    # branch row - uniform height so the columns line up
    branch_h = max(_flow_box_h(b, col_w, branch_fs) for b in chunk) if chunk \
        else _flow_box_h("", col_w, branch_fs)
    branch_top = FLOW_Y0 + root_h + FLOW_GAP1
    leaf_top0 = branch_top + branch_h + FLOW_GAP2
    return {"col_w": col_w, "branch_fs": branch_fs, "leaf_w": col_w - FLOW_IND,
            "root_fs": root_fs, "rw": rw, "root_h": root_h,
            "root_x": FLOW_X0 + (FLOW_W - rw) / 2, "branch_h": branch_h,
            "branch_top": branch_top, "leaf_top0": leaf_top0,
            "avail_leaf_h": FLOW_YB - leaf_top0}

def _fits_one_line(text: str, w: float, fs: float) -> bool:
    """Does ``text`` sit on ONE line in the deck at ``fs``?  Used to pick a leaf
    font that never hard-wraps a single long word."""
    return _wrap_count(text, w - 12, fs * FLOW_PPT_K) <= 1

def _column_h(leaves, leaf_w: float, fs: float) -> float:
    """Stacked height of one branch's leaf boxes at font ``fs`` (minimum gaps)."""
    hs = [_flow_box_h(l, leaf_w, fs) for l in leaves]
    return sum(hs) + max(0, len(hs) - 1) * FLOW_LG

FLOW_SPINE_OFF = 10       # gap between a leaf box's edge and its spine

def flow_edges(branch: dict, columns: list) -> list:
    """Connectors from ONE branch box down to its leaf sub-columns.

    The tree is drawn as a CENTRED (double-sided) org chart rather than as an
    indented file tree:

    * a column on the LEFT of the block carries its spine on its **right**, and
      its boxes hang off it leftwards;
    * a column on the RIGHT carries its spine on its **left**;
    * an odd column out sits in the MIDDLE and is a **chain** - the trunk drops
      into the top of its first box and each box links to the one below.

    So every connector lives in an interior gutter and the outside edges of the
    diagram are clean box edges, which is what makes it read as designed rather
    than as a bracket with everything hanging off it.  A single column is the
    middle case, so it chains.

    ``columns`` is a list of lists of leaf NODES (dicts with x/y/w/h), ordered
    left to right; each column is sorted top-down here.  Shared by the dashboard
    geometry and the PowerPoint re-measure so the two surfaces cannot drift.
    """
    cols = [sorted(c, key=lambda n: n["y"]) for c in columns if c]
    if not cols:
        return []
    bcx = branch["x"] + branch["w"] / 2
    bbot = branch["y"] + branch["h"]
    k = len(cols)
    mid = k // 2 if k % 2 else None          # even counts have no middle
    edges = []
    for j, col in enumerate(cols):
        x, w = col[0]["x"], col[0]["w"]
        tops = [n["y"] for n in col]
        bots = [n["y"] + n["h"] for n in col]
        centers = [n["y"] + n["h"] / 2.0 for n in col]
        if j == mid:
            cx = x + w / 2.0
            # Within a box-indent of the branch centre, drop STRAIGHT down: the
            # elbow would otherwise put a visible kink under every branch of a
            # classic multi-region chart, where the leaves are indented inside
            # their column and so sit a few px off centre.
            if abs(cx - bcx) <= FLOW_IND:
                edges.append({"pts": [(cx, bbot), (cx, tops[0])]})
            else:
                edges.append({"pts": [(bcx, bbot), (bcx, bbot + 11),
                                      (cx, bbot + 11), (cx, tops[0])]})
            for a, b in zip(bots, tops[1:]):
                edges.append({"pts": [(cx, a), (cx, b)]})
            continue
        if j < (mid if mid is not None else k / 2.0):
            spine_x, tick_x = x + w + FLOW_SPINE_OFF, x + w     # boxes to LEFT
        else:
            spine_x, tick_x = x - FLOW_SPINE_OFF, x            # boxes to RIGHT
        edges.append({"pts": [(bcx, bbot), (bcx, bbot + 11),
                              (spine_x, bbot + 11), (spine_x, centers[-1])]})
        for cy in centers:
            edges.append({"pts": [(spine_x, cy), (tick_x, cy)]})
    return edges

def _sub_gutter(k: int, leaf_w: float, inner: float) -> float:
    """Horizontal gap between a branch's leaf sub-columns.

    Proportional to the BOX width, so the tree keeps the same visual rhythm
    whatever size the auto-fit landed on, then capped by the room the branch
    column can actually spare.  With two 210px columns in a 1236px branch that
    is the full 84px; by the time five columns are competing for the same width
    it closes to whatever is left, never below FLOW_SUBGAP."""
    if k <= 1:
        return 0.0
    want = leaf_w * FLOW_SUBGAP_K
    room = (inner - k * leaf_w) / (k - 1)
    return max(FLOW_SUBGAP, min(want, max(room, FLOW_SUBGAP)))

def _leaf_grid(g: dict, lists: list, fs: float) -> tuple:
    """How one branch's leaves are arranged inside its column.

    Returns ``(k, leaf_w, rows_max, gutter)`` - ``k`` sub-columns of ``leaf_w``
    px separated by ``gutter`` px, and the number of leaf ROWS that fit the
    page's leaf band at font ``fs``.

    On a page whose natural leaf width is already sane this is ``(1, leaf_w,
    rows, 0)`` and every caller behaves exactly as it did before sub-columns
    existed.  It only splits when a column is wide enough to make its boxes
    read as banners (see FLOW_LEAF_PREF)."""
    natural, avail = g["leaf_w"], g["avail_leaf_h"]
    leaves = [l for v in lists for l in v]
    n = max((len(v) for v in lists), default=0)

    def rows_for(w):
        hs = [_flow_box_h(l, w, fs) for l in leaves]
        box = max(hs) if hs else _flow_box_h("", w, fs)
        return max(1, int((avail + FLOW_LG) // (box + FLOW_LG)))

    if natural <= FLOW_LEAF_PREF or not leaves:
        return 1, natural, rows_for(natural), 0.0

    # Boxes are capped at the preferred width, so the block is NARROWER than
    # the column and gets centred.  Pick the FEWEST sub-columns that hold the
    # leaves - fewer, taller columns fill the page instead of leaving a band of
    # empty slide under a wide, shallow grid.
    leaf_w = FLOW_LEAF_PREF
    rows = rows_for(leaf_w)
    k_fit = max(1, -(-n // rows))                      # ceil: columns needed
    # the ceiling uses the MINIMUM gutter - how many columns could ever be made
    # to fit, not how far apart they will end up sitting
    k_cap = max(1, int((natural + FLOW_SUBGAP) // (leaf_w + FLOW_SUBGAP)))
    k = min(k_fit, k_cap)
    gut = _sub_gutter(k, leaf_w, natural)
    # if the block would still overrun the column, share the width instead
    if k * leaf_w + (k - 1) * gut > natural:
        leaf_w = (natural - (k - 1) * FLOW_SUBGAP) / k
        gut = _sub_gutter(k, leaf_w, natural)
        rows = rows_for(leaf_w)
    return k, leaf_w, rows, gut

def _leaf_capacity(root_label: str, groups: dict, chunk: list,
                   fs: float = FLOW_MINLEAF) -> int:
    """How many leaves a branch column on this page can hold at font ``fs``.

    Returns the largest ``n`` for which EVERY consecutive run of ``n`` leaves,
    in every branch on the page, still fits the available height - so slicing
    each branch into runs of ``n`` is guaranteed to fit.  Always ≥ 1.

    Where the leaves lay out in sub-columns the ceiling is ``rows x columns``,
    which is why a region-scoped tree that used to spill onto a continuation
    slide now lands on one."""
    g = _flow_geom(root_label, chunk)
    lists = [list(groups.get(b, [])) for b in chunk]
    k, leaf_w, rows_max, _gut = _leaf_grid(g, lists, fs)
    if k > 1:
        return max(1, rows_max * k)
    avail = g["avail_leaf_h"]
    mx = max((len(v) for v in lists), default=0)
    for n in range(mx, 1, -1):
        if all(_column_h(v[i:i + n], leaf_w, fs) <= avail
               for v in lists for i in range(0, len(v), n)):
            return n
    return 1

def _flow_slide(root_label: str, groups: dict, chunk: list) -> dict:
    """Geometry for ONE slide: the root plus the branch columns in ``chunk``.

    Returns {"nodes": [...], "edges": [...]} where a node is
    {x, y, w, h, text, level (0 root/1 branch/2 leaf), fs, align} and an edge
    is {"pts": [(x, y), ...]} - an axis-aligned polyline."""
    g = _flow_geom(root_label, chunk)
    col_w, branch_fs, leaf_w = g["col_w"], g["branch_fs"], g["leaf_w"]
    root_fs, rw, root_h, root_x = (g["root_fs"], g["rw"], g["root_h"],
                                   g["root_x"])
    root_cx = root_x + rw / 2
    branch_h, branch_top = g["branch_h"], g["branch_top"]
    leaf_top0, avail_leaf_h = g["leaf_top0"], g["avail_leaf_h"]

    lists = [list(groups.get(b, [])) for b in chunk]

    # How the leaves sit inside a branch: normally one stack per branch, but on
    # a page with few branches the boxes would be banner-wide, so they wrap into
    # sub-columns instead (see FLOW_LEAF_PREF).
    def split_cols(v, k):
        """One branch's leaves dealt into ``k`` sub-columns, filled top-down
        then left-to-right so the reading order stays alphabetical."""
        if k <= 1:
            return [list(v)]
        per = max(1, -(-len(v) // k))
        return [v[i:i + per] for i in range(0, len(v), per)] or [[]]

    # auto-fit the leaf font so the tallest sub-column fits the page, then
    # remember each sub-column's stacked leaf-box heights (no gaps yet)
    def boxes_h(fs, k, lw):
        return {b: [[_flow_box_h(l, lw, fs) for l in col]
                    for col in split_cols(v, k)]
                for b, v in zip(chunk, lists)}

    def stacks(bh):                       # every sub-column's stacked height
        return [sum(c) + (len(c) - 1) * FLOW_LG
                for v in bh.values() for c in v if c]

    # a size is only acceptable if the column also fits the page height AND no
    # single word has to be broken across lines -- PowerPoint hard-wraps a word
    # wider than its box ("Netherland/s"), which no amount of height fixes
    all_leaves = [l for b in chunk for l in groups.get(b, [])]

    def no_word_breaks(fs, lw):
        return all(_fits_one_line(max(str(l).split() or [""], key=len),
                                  lw, fs)
                   for l in all_leaves)

    def fitted(fs):
        k, lw, _rows, gut = _leaf_grid(g, lists, fs)
        return k, lw, gut, boxes_h(fs, k, lw)

    leaf_fs = FLOW_LEAF_SIZES[-1]
    leaf_k, leaf_w, sub_gap, boxes = fitted(leaf_fs)
    for fs in FLOW_LEAF_SIZES:
        k, lw, gut, bh = fitted(fs)
        if (max(stacks(bh), default=0) <= avail_leaf_h
                and no_word_breaks(fs, lw)):
            leaf_fs, leaf_k, leaf_w, sub_gap, boxes = fs, k, lw, gut, bh
            break
        # keep last (floor) fit
        leaf_fs, leaf_k, leaf_w, sub_gap, boxes = fs, k, lw, gut, bh

    # breathe: grow the leaf gap so the fullest column reaches the bottom
    # margin (never overflowing any column), capped so it stays tidy
    gap_caps = [(avail_leaf_h - sum(c)) / (len(c) - 1)
                for v in boxes.values() for c in v if len(c) >= 2]
    leaf_gap = max(FLOW_LG, min(36.0, min(gap_caps))) if gap_caps else FLOW_LG

    # a leaf block narrower than its column is CENTRED under the branch, and
    # the branch header shrinks to sit over it rather than spanning the canvas
    block_w = leaf_k * leaf_w + (leaf_k - 1) * sub_gap
    narrowed = block_w + FLOW_IND < col_w - 0.5

    nodes, edges = [], []
    nodes.append({"x": int(root_x), "y": FLOW_Y0, "w": rw, "h": root_h,
                  "text": root_label, "level": 0, "fs": root_fs,
                  "align": "center"})

    branch_cxs = []
    for i, (b, v) in enumerate(zip(chunk, lists)):
        col_left = FLOW_X0 + i * (col_w + FLOW_CG)
        bcx = col_left + col_w / 2
        branch_cxs.append(bcx)
        if narrowed:
            # header sits over its own leaf block, not over the whole canvas
            bw = min(col_w, max(block_w, FLOW_MINBRANCH))
            block_left = bcx - block_w / 2
        else:
            bw = col_w
            block_left = col_left + FLOW_IND
        branch_node = {"x": int(bcx - bw / 2), "y": int(branch_top),
                       "w": int(bw), "h": branch_h, "text": b,
                       "level": 1, "fs": branch_fs, "align": "center"}
        nodes.append(branch_node)
        # branch -> leaves : connectors face INWARD (see flow_edges)
        col_nodes = []
        for j, col in enumerate(split_cols(v, leaf_k)):
            leaf_x = block_left + j * (leaf_w + sub_gap)
            y = leaf_top0
            made = []
            for leaf, lh in zip(col, boxes[b][j]):
                n = {"x": int(leaf_x), "y": int(y),
                     "w": int(leaf_w), "h": lh, "text": leaf,
                     "level": 2, "fs": leaf_fs, "align": "left"}
                nodes.append(n)
                made.append(n)
                y += lh + leaf_gap
            col_nodes.append(made)
        edges.extend(flow_edges(branch_node, col_nodes))

    # root -> branches trunk (down, across the bus, down to each branch)
    root_bottom = FLOW_Y0 + root_h
    bus_y = root_bottom + FLOW_GAP1 * 0.5
    edges.append({"pts": [(root_cx, root_bottom), (root_cx, bus_y)]})
    if len(branch_cxs) > 1:
        edges.append({"pts": [(min(branch_cxs), bus_y),
                              (max(branch_cxs), bus_y)]})
    for bcx in branch_cxs:
        edges.append({"pts": [(bcx, bus_y), (bcx, branch_top)]})

    # The diagram is TOP-ALIGNED: it starts just under the title band and any
    # leftover height is left at the bottom.  It used to be pushed down by 45%
    # of the slack to "sit balanced", but on a sparse slide (a few regions with
    # one country each) that floated the chart in the middle of an empty page
    # and read as a centrepiece rather than a taxonomy.  Reading starts at the
    # top, so the diagram starts at the top.

    # round + int the edge points for clean rendering
    for e in edges:
        e["pts"] = [(int(round(x)), int(round(y))) for x, y in e["pts"]]
    return {"nodes": nodes, "edges": edges}

# ── Contents (button-grid navigation) geometry ──────────────────────
# Shared by the dashboard page and the PPT slide so both place the section
# buttons identically: a balanced 1- or 2-column grid, filled column by column,
# sized to the section count so nothing overlaps or clips.
def contents_buttons_layout(buttons: list[dict]) -> list[dict]:
    X0, top, W = 15, 100, CONTENT_W
    colgap, rowgap = 26, 12
    n = max(1, len(buttons))
    ncols = 1 if n <= 6 else 2
    per = -(-n // ncols)                          # ceil division -> rows/column
    col_w = (W - (ncols - 1) * colgap) / ncols
    avail = CONTENT_BOTTOM - top
    bh = min(66.0, (avail - (per - 1) * rowgap) / per)
    items = []
    for i, b in enumerate(buttons):
        c, r = i // per, i % per
        items.append({
            "x": int(round(X0 + c * (col_w + colgap))),
            "y": int(round(top + r * (bh + rowgap))),
            "w": int(round(col_w)), "h": int(round(bh)),
            "label": b["label"], "slug": b["slug"]})
    return items

# ── Contents (formal Table of Contents) geometry ─────────────────────
# Shared by the dashboard page and the PPT slide, so both lay the numbered
# sections out identically.  Returns a flat list of positioned draw items
# (subtitle / accent / number / title / sub / divider); the renderers turn
# them into HTML (with dcc.Link) or native PPT text boxes (with hyperlinks).
def contents_layout(sections: list[dict]) -> list[dict]:
    X0, W = 44, 1280 - 88
    top0, bottom, colgap = 152, 696, 56
    NUM_FS, TITLE_FS, SUB_FS = 19.0, 15.0, 11.5
    NUM_H, TITLE_H, SUB_H = 28, 25, 22
    DIV_GAP, SEC_GAP, SUB_PAD, SUB_IND = 12, 22, 4, 22

    def sec_h(s):
        h = NUM_H + TITLE_H
        if s["subs"]:
            h += SUB_PAD + len(s["subs"]) * SUB_H
        return h + DIV_GAP + SEC_GAP

    heights = [sec_h(s) for s in sections]
    total = sum(heights) or 1
    avail = bottom - top0

    def split(ncols):
        cols = [[] for _ in range(ncols)]
        colsum = [0.0] * ncols
        target = total / ncols
        col = 0
        for s, h in zip(sections, heights):
            if col < ncols - 1 and colsum[col] > 0 and colsum[col] + h > target:
                col += 1
            cols[col].append((s, h))
            colsum[col] += h
        return cols, (max(colsum) if colsum else 0)

    ncols, cols, mx = 1, *split(1)
    for nc in (1, 2, 3):
        c, m = split(nc)
        ncols, cols, mx = nc, c, m
        if m <= avail or nc == 3:
            break
    scale = min(1.0, avail / mx) if mx > avail else 1.0
    nh, th, sh = NUM_H * scale, TITLE_H * scale, SUB_H * scale
    nfs, tfs, sfs = NUM_FS * scale, TITLE_FS * scale, SUB_FS * scale
    spad, dgap, sgap = SUB_PAD * scale, DIV_GAP * scale, SEC_GAP * scale
    col_w = (W - (ncols - 1) * colgap) / ncols

    items = [
        {"kind": "subtitle", "x": X0, "y": 112, "w": W,
         "text": "Click a section to navigate"},
        {"kind": "accent", "x": X0, "y": 138, "w": 210},
    ]
    for c in range(ncols):
        cx = X0 + c * (col_w + colgap)
        y = float(top0)
        for s, _ in cols[c]:
            items.append({"kind": "number", "x": int(cx), "y": int(y),
                          "text": s["no"], "fs": nfs})
            y += nh
            items.append({"kind": "title", "x": int(cx), "y": int(y),
                          "w": int(col_w), "text": s["title"],
                          "slug": s["slug"], "fs": tfs})
            y += th
            if s["subs"]:
                y += spad
                for sub in s["subs"]:
                    items.append({"kind": "sub", "x": int(cx + SUB_IND),
                                  "y": int(y), "w": int(col_w - SUB_IND),
                                  "text": sub["label"], "slug": sub["slug"],
                                  "fs": sfs})
                    y += sh
            y += dgap
            items.append({"kind": "divider", "x": int(cx), "y": int(y),
                          "w": int(col_w)})
            y += sgap
    return items

def taxonomy_flow(root_label: str, groups: dict) -> list[dict]:
    """Build the taxonomy org chart, paginating in BOTH directions so nothing
    is squeezed past legibility.  ``groups`` is an ordered
    {branch: [leaf, ...]} mapping taken straight from the data.

    - *Horizontally*: branches spill to another slide once they cannot hold a
      readable column width (``FLOW_MINCOL``).
    - *Vertically*: a branch with more leaves than fit at ``FLOW_MINLEAF``
      continues on a further slide instead of shrinking the labels - every
      slide repeats the root and the branch headers, and the caller's existing
      "(continued j/k)" heading identifies it.

    Both surfaces get this for free: ``paginate``/``page_count`` already return
    one page per flow slide, and ``export_native._taxonomy_flow_slides`` emits
    one PPT slide per flow slide."""
    branches = list(groups.keys())
    maxcols = max(1, int((FLOW_W + FLOW_CG) // (FLOW_MINCOL + FLOW_CG)))
    per = min(len(branches), maxcols) or 1
    chunks = [branches[i:i + per] for i in range(0, len(branches), per)] \
        or [[]]
    slides = []
    for ch in chunks:
        mx = max((len(groups.get(b, [])) for b in ch), default=0)
        cap = _leaf_capacity(root_label, groups, ch)
        if mx <= cap:                               # fits readably in one page
            slides.append(_flow_slide(root_label, groups, ch))
            continue
        # spread the leaves EVENLY over the slides they need, rather than
        # filling each to capacity and leaving a stub (14 over a cap of 13
        # reads as 7 + 7, not 13 + 1)
        k = -(-mx // cap)                           # ceil: slides required
        per_slide = -(-mx // k)
        for i in range(0, mx, per_slide):
            sub = {b: list(groups.get(b, []))[i:i + per_slide] for b in ch}
            slides.append(_flow_slide(root_label, sub, ch))
    return slides

# ── Taxonomy segment grid (reference "By Segment" boxes) ─────────────
# The Market Segments page as the reference slide draws it: one boxed list per
# segment axis, a filled header bar ("By Product Type") over bordered rows.  A
# segment is a ❑ row - bold when its axis goes a level deeper, plain in a flat
# list - and each of its sub-segments a ➢ row indented beneath it.
#
# * A box holds at most TAX_ROWS rows.  A longer axis continues in a duplicate
#   box headed "(1/2)", "(2/2)" ..., repeating the parent segment at the top of
#   the continuation so a ➢ row is never cut off from its ❑.
# * Boxes go into at most TAX_LANES columns per slide, each into the shortest
#   column that still has room (so short lists stack, as in the reference);
#   a box that fits no column starts the next slide.
# * The columns a slide actually uses share its width equally: four axes give
#   four columns, two give two - nothing is drawn for axes the workbook lacks.
#
# Row height is the band height over TAX_ROWS, so a full box exactly fills a
# column and nothing can run off the slide.  A label that wraps counts as the
# rows it occupies.  Pagination measures at the narrowest width (TAX_LANES
# columns); a slide with fewer columns draws wider, so it only ever needs
# fewer lines than were budgeted.
TAX_Y0 = CONTENT_TOP + 5                  # first box top (under the header)
TAX_YB = 690                              # last usable y (above the footer)
TAX_LANES = 4                             # most columns on one slide
TAX_ROWS = 30                             # most rows in one box
TAX_GAP_X = 14                            # gutter between columns
TAX_GAP_Y = 12                            # gap between stacked boxes
TAX_HEAD_H = 24                           # one-line header bar
TAX_FS = 11.0                             # the deck's content size
TAX_ROW_H = (TAX_YB - TAX_Y0 - TAX_HEAD_H) / TAX_ROWS
TAX_MARK_W = 20                           # marker glyph cell
TAX_IND = 18                              # sub-segment indent
TAX_EDGE = "#C9D2D9"                      # box border (as Company Coverage)
TAX_RULE = "#E4EAEF"                      # row separators
TAX_MAIN, TAX_SUB = "❑", "➢"
# Wrap factor: pt -> the width model's px.  Calibrated against PowerPoint's own
# line breaks (TextRange.BoundHeight) on this page in both templates: regular
# text never wraps later than predicted from 1.19 (Calibri, FMI) / 1.12 (Segoe
# UI, Fact.MR) up, so 1.25 keeps a margin for labels not yet seen while no
# longer reserving a second row for labels that sit on one line - the general
# 1.35 did that to 8 of 142.  Bold text (headers, parent rows) had no borderline
# case to calibrate on, so it keeps the general factor.
TAX_K = 1.25
TAX_K_BOLD = FLOW_PPT_K

def _tax_lane_w(n: int) -> float:
    return (CONTENT_W - (n - 1) * TAX_GAP_X) / max(1, n)

def _tax_text_x(level: int) -> float:
    """Left offset of a row's label inside its box."""
    return 6 + TAX_MARK_W + level * TAX_IND

def _tax_lines(r: dict, w: float) -> int:
    """Rows a row's label occupies in a box ``w`` wide (measured for the deck,
    less the 4px right pad and the text box's own 3px margins)."""
    usable = w - _tax_text_x(r["level"]) - 4 - 6
    k = TAX_K_BOLD if r["bold"] else TAX_K
    return _wrap_count(r["text"], usable, TAX_FS * k)

def _tax_head_h(title: str, w: float) -> float:
    n = _wrap_count(title, w - 16 - 6, TAX_FS * TAX_K_BOLD)
    return TAX_HEAD_H if n <= 1 else n * TAX_ROW_H + 6

def _tax_text(x, y, w, h, s, bold, color, n=1) -> dict:
    p = {"t": "text", "x": x, "y": y, "w": w, "h": h, "s": s,
         "size": TAX_FS, "bold": bold, "color": color, "align": "l"}
    if n > 1:
        # wrapped: one row per line, PowerPoint single-spaced and anchored at
        # the top, so the first line stays level with the row's marker
        p.update({"wrap": True, "lh": TAX_ROW_H, "line": 1.0, "anchor": "t"})
    return p

def _tax_split(rows: list, cap: int, w: float) -> list:
    """Cut one axis's rows into boxes of at most ``cap`` rows.  A ❑ row never
    ends a box its ➢ rows continue past, and a box that opens on ➢ rows
    repeats their ❑ first."""
    pieces, cur, used, parent = [], [], 0, None
    for r in rows:
        n = _tax_lines(r, w)
        if r["level"] == 0:
            parent = r
        if cur and used + n > cap:
            if cur[-1]["level"] == 0 and r["level"] == 1 and len(cur) > 1:
                carry = [cur.pop()]
            else:
                carry = [parent] if r["level"] == 1 and parent else []
            pieces.append(cur)
            cur = carry
            used = sum(_tax_lines(c, w) for c in cur)
        cur.append(r)
        used += n
    if cur:
        pieces.append(cur)
    return pieces

def _tax_box(P: list, x, y, w, box, head_fill, head_ink, ink) -> float:
    """Draw one box at (x, y), ``w`` wide, into ``P``; returns its bottom."""
    hh = _tax_head_h(box["title"], w)
    lines = [_tax_lines(r, w) for r in box["rows"]]
    h = hh + sum(lines) * TAX_ROW_H
    P.append({"t": "rect", "x": x, "y": y, "w": w, "h": h,
              "fill": "#FFFFFF", "line": TAX_EDGE, "lw": 1})
    P.append({"t": "rect", "x": x, "y": y, "w": w, "h": hh,
              "fill": head_fill, "line": head_fill, "lw": 1})
    P.append(_tax_text(x + 8, y, w - 16, hh, box["title"], True, head_ink,
                       1 if hh <= TAX_HEAD_H else 2))
    ry = y + hh
    for i, (r, n) in enumerate(zip(box["rows"], lines)):
        if i:
            P.append({"t": "line", "x1": x + 1, "y1": ry, "x2": x + w - 1,
                      "y2": ry, "color": TAX_RULE, "width": 1})
        lvl = r["level"]
        P.append(_tax_text(x + 6 + lvl * TAX_IND, ry, TAX_MARK_W, TAX_ROW_H,
                           TAX_SUB if lvl else TAX_MAIN, False, ink))
        tx = x + _tax_text_x(lvl)
        P.append(_tax_text(tx, ry, w - (tx - x) - 4, n * TAX_ROW_H, r["text"],
                           r["bold"], ink, n))
        ry += n * TAX_ROW_H
    return y + h

def taxonomy_columns(groups: list, head_fill: str, head_ink: str,
                     ink: str) -> list[list[dict]]:
    """Primitive pages - one per slide - for the segment grid.

    ``groups`` is an ordered [(header, [(segment, [sub-segment, ...]), ...])]
    list straight from the data, one entry per axis the workbook has."""
    band = TAX_YB - TAX_Y0
    w4 = _tax_lane_w(TAX_LANES)                  # narrowest column: budget here
    boxes = []
    for title, segs in groups:
        deep = any(kids for _m, kids in segs)
        rows = []
        for m, kids in segs:
            rows.append({"text": str(m), "level": 0, "bold": deep})
            rows.extend({"text": str(k), "level": 1, "bold": False}
                        for k in kids)
        if not rows:
            continue
        # a header that wraps (long axis name plus its "(j/k)") costs rows
        head = _tax_head_h(f"{title} (1/2)", w4)
        cap = min(TAX_ROWS, int((band - head) / TAX_ROW_H + 1e-6))
        pieces = _tax_split(rows, max(1, cap), w4)
        k = len(pieces)
        for j, piece in enumerate(pieces, start=1):
            t = title if k == 1 else f"{title} ({j}/{k})"
            h = _tax_head_h(t, w4) + TAX_ROW_H * sum(
                _tax_lines(r, w4) for r in piece)
            boxes.append({"title": t, "rows": piece, "h": h})

    # shortest column first, in axis order; no room anywhere -> next slide
    pages, used = [], []
    for b in boxes:
        fit = [i for i, u in enumerate(used)
               if u + (TAX_GAP_Y if u else 0) + b["h"] <= band + 0.5]
        if not fit:
            pages.append([[] for _ in range(TAX_LANES)])
            used = [0.0] * TAX_LANES
            fit = list(range(TAX_LANES))
        i = min(fit, key=lambda c: (used[c], c))
        pages[-1][i].append(b)
        used[i] += (TAX_GAP_Y if used[i] else 0) + b["h"]

    out = []
    for lanes in pages:
        lanes = [ln for ln in lanes if ln]      # empty columns only trail
        w = _tax_lane_w(len(lanes))
        P = []
        for li, lane in enumerate(lanes):
            x = PAGE_LEFT + li * (w + TAX_GAP_X)
            y = TAX_Y0
            for b in lane:
                y = _tax_box(P, x, y, w, b, head_fill, head_ink, ink) \
                    + TAX_GAP_Y
        out.append(P)
    return out or [[]]

def continuation_canvas(page, title, j, k):
    from dash import html
    band = components.static_band(f"{title} (continued {j}/{k})")
    return html.Div([band, *_components(page["blocks"])], style={
        "position": "relative", "width": CANVAS_W, "height": page["height"],
        "background": PAGE_BG, "flexShrink": 0,
        "boxShadow": "0 2px 10px rgba(0,40,80,0.18)",
        "fontFamily": "'Segoe UI', system-ui, sans-serif"})

def render_stack(title, spec):
    """Full standalone render (no-slicer path): stacked canvases."""
    from dash import html
    pages = paginate(spec)
    k = len(pages)
    kids = []
    for i, pg in enumerate(pages):
        if i == 0:
            band = components.static_band(title)
            kids.append(html.Div([band, *_components(pg["blocks"])], style={
                "position": "relative", "width": CANVAS_W,
                "height": pg["height"], "background": PAGE_BG,
                "flexShrink": 0,
                "boxShadow": "0 2px 10px rgba(0,40,80,0.18)",
                "fontFamily": "'Segoe UI', system-ui, sans-serif"}))
        else:
            kids.append(continuation_canvas(pg, title, i + 1, k))
    return html.Div(kids, style={"display": "flex",
                                 "flexDirection": "column",
                                 "alignItems": "center", "gap": 18})


# ── Template switch support ──────────────────────────────────────────
def _retheme() -> None:
    """Re-derive the constants this module bakes from theme tokens.

    Only the panel-title size and three tints depend on the template; the page
    box (CONTENT_TOP / CONTENT_BOTTOM) is deliberately SHARED by both templates
    so a page paginates identically either way and the deck never gains or
    loses a slide when the template changes.
    """
    global TITLE_FS, _PP_BAND_FILL, _PP_BAND_LINE, FLOW_LINE
    TITLE_FS = FS_PANEL_TITLE
    _PP_BAND_FILL, _PP_BAND_LINE = _T_PP_BAND_FILL, _T_PP_BAND_LINE
    FLOW_LINE = _T_FLOW_LINE
