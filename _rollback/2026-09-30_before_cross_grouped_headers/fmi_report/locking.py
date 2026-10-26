"""Locked exports - the deck with the values a client must not read withheld.

A thin adapter over the standalone `pptlock.py` at the project root.  It adds
nothing to the locking rules; it only decides, from THIS report's own model,
what the policy should be:

* **open years** - the last historical year, the estimate year and the
  forecast horizon (`config.HIST_END` / `EST_YEAR` / `END_YEAR`), so a locked
  deck still tells the story its teaser is meant to tell.  They are read from
  `config`, not hardcoded, so they follow the workbook like everything else.
* **protected phrases** - the market's own name, so a study called
  "Prescription 5000-ppm Fluoride Toothpaste" does not withhold its subject.

The unlocked exports are untouched: `export_pptx_native` /
`export_regions_zip` build the deck exactly as before and the lock is applied
afterwards, to the finished file.
"""
from __future__ import annotations

import re
import shutil
import sys
import tempfile
import zipfile
from contextlib import contextmanager
from copy import deepcopy
from pathlib import Path

from dash import dcc, html

from . import config, data

_PROJ_ROOT = Path(__file__).resolve().parents[1]


def _pptlock():
    """Import the standalone module, wherever it is sitting.

    Normally `pptlock.py` is at the project root, which is on `sys.path` for
    both `python run.py` and `gunicorn run:server`.  The fallback loads it by
    path so a checkout started from another working directory still locks.
    """
    try:
        import pptlock                                    # noqa: PLC0415
        return pptlock
    except ImportError:
        pass
    import importlib.util
    src = _PROJ_ROOT / "pptlock.py"
    if not src.exists():
        raise ImportError(
            f"pptlock.py not found at {src} - the locked exports need it.")
    spec = importlib.util.spec_from_file_location("pptlock", src)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _pptblur():
    """The blur companion, loaded the same way as `pptlock`."""
    try:
        import pptblur                                    # noqa: PLC0415
        return pptblur
    except ImportError:
        pass
    import importlib.util
    src = _PROJ_ROOT / "pptblur.py"
    if not src.exists():
        return None
    spec = importlib.util.spec_from_file_location("pptblur", src)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def default_policy(**over):
    """The report's own locking policy.  Keyword arguments override it.

    **NO year is open, no CAGR, and no page may hand a figure back.**  The
    deck used to leave the last historical year, the estimate year and the
    forecast horizon readable, plus every CAGR, and the page rulebook could
    keep a few figures besides; all of that is now withheld like any other
    value.  `guess_years=False` matters as much as the empty set: `scan`
    fills an empty `open_years` in from the deck's own headers unless told
    not to.  Pass `open_years=`/`page_exemptions=True` to get the teaser back.

    `totals_only=True` is the reference deck's rule - it narrows an open
    column to the TOTAL row - and is kept as it stands.  With no open column
    left it is inert rather than removed.
    """
    pl = _pptlock()
    # The market name may be phrased either way in the deck ("Amylases AGM" in
    # a chart title, "Amylases AGM Market" in a heading), so protect both.
    #
    # Segment and sub-segment names go in beside it, because a segment is
    # often named BY a figure - a "Pricing Tier" axis splits into "Below 80
    # USD" and "Above 100 USD", a pack size into "500 ml".  Those figures name
    # the category; they are not the category's value, and withholding them
    # left rows and axis ticks reading "Below 🔒 USD", which tells the reader
    # nothing and withholds nothing.  Only names carrying a DIGIT are listed:
    # a name without one can never shield a figure, and every phrase here
    # costs a scan of every sentence in the deck.
    named = [str(s) for s in data.SEGMENTS]
    for subs in data.SUBSEGMENTS.values():
        named += [str(s) for s in subs]
    # ...and the level beneath, which the taxonomy page lists by name
    for kids in getattr(data, "SEGMENT_CHILDREN", {}).values():
        named += [str(k) for sub, ks in kids.items() for k in (sub, *ks)]
    # The segment pages' 'Section 7.1' bars read like a decimal figure; they
    # are page numbering, not market data (2026-09-28)
    sub_sections = [f"Section {n}.{k}" for n in range(1, 31)
                    for k in range(1, len(data.SEGMENTS) + 1)]
    # ...and the Industry Structure page's HHI class bands, which are fixed
    # thresholds rather than figures (2026-09-29)
    from .templates_industry import HHI_BANDS
    protect = tuple(dict.fromkeys(
        [p for p in (data.MARKET_LABEL, data.MARKET) if p]
        + [p for p in named if any(c.isdigit() for c in p)]
        + sub_sections + list(HHI_BANDS)))
    kw = dict(open_years=set(), open_cagr=False, guess_years=False,
              page_exemptions=False, protect=protect, totals_only=True)
    kw.update(over)
    return pl.LockPolicy(**kw)


@contextmanager
def _deep_rels(limit: int = 10_000):
    """Room on the stack for a big deck's relationship graph.

    python-pptx walks package relationships by recursion, once per related
    part, both when opening a deck and when saving one.  A report with many
    countries puts a four-figure number of parts in a single deck, and the
    walk then runs past CPython's default 1,000-frame limit and dies with a
    RecursionError after the whole deck has been built.  Raised for the
    duration of the call only.
    """
    was = sys.getrecursionlimit()
    try:
        sys.setrecursionlimit(max(was, limit))
        yield
    finally:
        sys.setrecursionlimit(was)


def lock_deck(src, dst=None, policy=None, blur: bool = True):
    """Withhold this deck's protected values.  Returns the written path.

    `pptlock` now carries the whole treatment itself: the blur under the
    padlock, the axis labels a chart cannot be read without, the inline
    padlock inside a sentence, and a PAGE RULEBOOK that says how each page of
    this report is marked - a padlock per bar where there are three of them,
    one centred mark where there are sixteen, a few scattered figures left
    readable in a ranking table.  `pptblur.py` was this project's own,
    smaller answer to the same problem and is superseded; it stays on disk
    for reference and is no longer in the export path.

    `dst=None` locks the file in place (via a temporary copy, so a failure
    part-way through cannot leave a half-written deck behind).
    """
    pl = _pptlock()
    src = Path(src)
    in_place = dst is None
    out = Path(tempfile.mkdtemp()) / src.name if in_place else Path(dst)
    out.parent.mkdir(parents=True, exist_ok=True)
    with _deep_rels():
        pl.apply_locks(src, out, policy=policy or default_policy())
    mask_chart_categories(out)
    scrub_chart_data(out)
    shield_charts(out)
    if in_place:
        shutil.move(str(out), str(src))
        return src
    return out


# ───────────────────────── the hover tooltip ─────────────────────────────
# PowerPoint reports a chart point on hover - "Series 'Europe' Point '2031' /
# Value: 45403.856" - reading it straight out of `<c:numCache>`.  That is the
# withheld figure at FULL precision, which is more than the slide ever showed.
#
# The cache cannot simply be emptied: a bar's height IS its cached value, so
# clearing it clears the bar.  And the bars cannot be redrawn as plain shapes
# either - 69 of this deck's 80 charts lay their plot area out AUTOmatically
# and 65 scale their axis automatically, so PowerPoint decides the geometry at
# render time and there is no reliable rectangle to copy.
#
# So the tooltip is stopped at the mouse instead of at the data: a shape laid
# over the chart takes the hover, and a chart that never sees the pointer
# never draws a tooltip.  The shape has to carry a FILL to capture the mouse -
# a "no fill" shape in PowerPoint is only live on its outline - so it is
# filled white at 1% opacity, which is invisible on any ground this deck uses.
#
# ⚠ Be clear about what this is and is not.  It stops a READER: hover and
# "Edit Data" (see `scrub_chart_data`) are how a client would actually meet
# these numbers.  It is NOT cryptographic - the values remain in the chart XML
# and anyone who unzips the .pptx, or deletes the overlay in PowerPoint, can
# still reach them.  Withholding a figure from a sample deck is presentation,
# not encryption; if a number must be genuinely unrecoverable it has to not be
# exported at all.

# ─────────────────── a figure hidden inside a CATEGORY label ──────────────
# `composition_bars` names each row "North America  ·  70.7K" - the group and
# its absolute size - and the same string becomes the chart's CATEGORY label
# on both surfaces.  Neither locking pass reached it: `pptlock` deliberately
# leaves axis text alone (a chart cannot be read without its scale) and
# `_mask_figure` only rewrote the labels a trace DRAWS.  So a padlocked page
# still printed every region's size down its own axis - which on these pages
# is the very figure the rest of the slide is withholding.
#
# The value is REMOVED, not covered: what is written back carries no number,
# so unzipping the file does not give it back.  Only the NAME is kept - "U.S.",
# not "U.S.  ·  ▒▒▒" - the same rule `pptlock.mask_category_value` follows, so
# whichever pass reaches a label first, the axis reads the same.
_CAT_MASK = "▒▒▒"                        # ▒▒▒ (no longer written)

# NAME · VALUE, value last: 70.7K / 1.2M / 3.4B / 842 / 1,234
_CAT_VALUE_RE = re.compile(
    r"^(?P<name>.+?)\s*[·•]\s*"
    r"(?P<val>[$€£¥₹]?\d[\d,]*(?:\.\d+)?\s*(?:[KMB]n?)?%?)\s*$")

# A bare four-digit number reads as a YEAR, not a size - "Europe · 2036" is a
# period label and must survive.  A grouped or suffixed figure never can be.
_YEARLIKE_RE = re.compile(r"^(?:19|20)\d{2}$")


def mask_cat_label(text):
    """The category label with its embedded figure masked, or None.

    None means "nothing to withhold here" so every caller can leave the label
    exactly as it found it.
    """
    m = _CAT_VALUE_RE.match(str(text))
    if not m:
        return None
    val = m.group("val").strip()
    if _YEARLIKE_RE.match(val):
        return None
    return m.group("name").rstrip()


_SHIELD_TAG = "lockshield"


def shield_charts(path, alpha_pct: float = 1.0) -> dict:
    """Lay an invisible shape over every chart so hover reveals nothing.

    Returns a count of what was covered.  Idempotent: a deck already shielded
    is left alone, so locking a locked deck does not stack overlays.
    """
    from pptx import Presentation                          # noqa: PLC0415
    from pptx.util import Emu                              # noqa: PLC0415
    from pptx.enum.shapes import MSO_SHAPE                 # noqa: PLC0415

    prs = Presentation(str(path))
    n_charts = n_shielded = 0
    for slide in prs.slides:
        frames = [sh for sh in slide.shapes if getattr(sh, "has_chart", False)]
        already = {sh.name for sh in slide.shapes if sh.name.startswith(_SHIELD_TAG)}
        for gf in frames:
            n_charts += 1
            tag = f"{_SHIELD_TAG}_{gf.shape_id}"
            if tag in already:
                continue
            box = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Emu(gf.left), Emu(gf.top), Emu(gf.width), Emu(gf.height))
            box.name = tag
            box.shadow.inherit = False
            box.line.fill.background()
            fill = box.fill
            fill.solid()
            fill.fore_color.rgb = _white()
            _set_alpha(fill.fore_color, alpha_pct)
            # a freshly added shape carries no click action, which is what we
            # want - `click_action.action` is read-only, so there is nothing
            # to clear here.
            n_shielded += 1
    prs.save(str(path))
    return {"charts": n_charts, "shielded": n_shielded}


def _white():
    from pptx.dml.color import RGBColor                    # noqa: PLC0415
    return RGBColor(0xFF, 0xFF, 0xFF)


def _set_alpha(color, pct: float) -> None:
    """`<a:alpha val>` is in thousandths of a percent: 1% -> 1000."""
    from pptx.oxml.ns import qn                            # noqa: PLC0415
    from pptx.oxml import parse_xml                        # noqa: PLC0415
    from pptx.oxml.ns import nsdecls                       # noqa: PLC0415
    srgb = color._xFill.find(qn("a:srgbClr"))
    if srgb is None:
        return
    for old in srgb.findall(qn("a:alpha")):
        srgb.remove(old)
    srgb.append(parse_xml(
        f'<a:alpha {nsdecls("a")} val="{int(max(0.0, pct) * 1000)}"/>'))


# ─────────────────────── the SECOND way a chart leaks ────────────────────
# `pptlock` withholds what a slide DRAWS.  A native chart also carries the
# numbers that produced it, in two places the padlock never reaches:
#
#   1. `ppt/embeddings/Microsoft_Excel_SheetN.xlsx` - the chart's own data
#      sheet.  Right-click a locked chart -> "Edit Data" and the whole series
#      opens in Excel, every withheld year at full precision.  This is the
#      worse of the two: it is the entire dataset for that chart, not one
#      figure.
#   2. `<c:numCache>` inside `ppt/charts/chartN.xml` - what PowerPoint reads
#      to draw the bar, and what it reports in the HOVER tooltip
#      ("Series 'Europe' Point '2031' / Value: 45403.856").
#
# (1) is pure loss with no visual cost, so it is removed unconditionally.
# (2) cannot be removed outright: a bar's height IS its cached value, so
# deleting it deletes the bar.  See `blank_locked_points` below.
#
# This lives HERE rather than in `pptlock.py` on purpose - that file is
# vendored from the analyst's Locker folder and is kept pristine so a future
# drop-in still applies cleanly.

_CHART_RE = re.compile(r"ppt/charts/chart\d+\.xml$")
_EXTERNAL_DATA = re.compile(rb"<c:externalData[^>]*?(?:/>|>.*?</c:externalData>)",
                            re.S)
_FORMULA = re.compile(rb"<c:f>.*?</c:f>", re.S)


def scrub_chart_data(path) -> dict:
    """Strip every chart's embedded workbook from a finished deck.

    Removes the `ppt/embeddings/*.xlsx` parts, the `<c:externalData>` element
    that points at them, the relationship entries and the `[Content_Types]`
    overrides - so "Edit Data" has nothing to open.  The `<c:f>` formula
    references go too: they name cells in a workbook that is no longer there.

    The cached values are left alone, so every chart still draws exactly as it
    did.  Returns a small report of what was taken out.
    """
    path = Path(path)
    zin = zipfile.ZipFile(path)
    names = zin.namelist()
    embeds = {n for n in names if n.startswith("ppt/embeddings/")}
    report = {"embedded_workbooks": len(embeds), "charts": 0, "bytes_before":
              path.stat().st_size}

    tmp = Path(tempfile.mkdtemp()) / path.name
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            n = item.filename
            if n in embeds:
                continue                       # drop the data sheet entirely
            data = zin.read(n)
            if _CHART_RE.match(n):
                data = _FORMULA.sub(b"", _EXTERNAL_DATA.sub(b"", data))
                report["charts"] += 1
            elif n.startswith("ppt/charts/_rels/") and n.endswith(".rels"):
                data = _drop_rels(data, "package")
            elif n == "[Content_Types].xml":
                data = _drop_overrides(data, embeds)
            zout.writestr(item, data)
    zin.close()
    shutil.move(str(tmp), str(path))
    report["bytes_after"] = path.stat().st_size
    return report


_CAT_BLOCK_RE = re.compile(rb"<c:cat>.*?</c:cat>", re.S)
_PT_V_RE = re.compile(rb"(<c:pt\b[^>]*>\s*<c:v>)(.*?)(</c:v>)", re.S)


def mask_chart_categories(path) -> dict:
    """Withhold a figure carried inside a chart's CATEGORY labels.

    These labels live in the chart part's `<c:cat>` string cache, not in any
    shape, which is why no padlock or blur ever reached them.  Only the cache
    is rewritten - the numeric `<c:val>` series that give the bars their
    lengths are untouched, so every chart still draws exactly as it did.

    Runs over every chart in the deck and masks only what
    `mask_cat_label` recognises, so a chart whose categories are plain names
    (or years) is left alone.
    """
    path = Path(path)
    zin = zipfile.ZipFile(path)
    report = {"charts": 0, "charts_masked": 0, "labels_masked": 0}

    def _mask_cat_block(block: bytes) -> bytes:
        def _one(m):
            raw = m.group(2)
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:
                return m.group(0)
            new = mask_cat_label(_unescape(text))
            if new is None:
                return m.group(0)
            report["labels_masked"] += 1
            return m.group(1) + _escape(new).encode("utf-8") + m.group(3)
        return _PT_V_RE.sub(_one, block)

    tmp = Path(tempfile.mkdtemp()) / path.name
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if _CHART_RE.match(item.filename):
                report["charts"] += 1
                before = report["labels_masked"]
                data = _CAT_BLOCK_RE.sub(
                    lambda m: _mask_cat_block(m.group(0)), data)
                if report["labels_masked"] > before:
                    report["charts_masked"] += 1
            zout.writestr(item, data)
    zin.close()
    shutil.move(str(tmp), str(path))
    return report


def _unescape(s: str) -> str:
    return (s.replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&apos;", "'")
             .replace("&amp;", "&"))


def _escape(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;"))


def _drop_rels(xml: bytes, kind: str) -> bytes:
    """Remove the relationship a chart uses to reach its embedded workbook."""
    out, changed = [], False
    for rel in re.split(rb"(?=<Relationship\b)", xml):
        if b"<Relationship" in rel and b"/package" in rel and b"embeddings/" in rel:
            head = rel.split(b"/>", 1)
            if len(head) == 2:
                rel = head[1]
                changed = True
            else:
                continue
        out.append(rel)
    return b"".join(out) if changed else xml


def _drop_overrides(xml: bytes, parts) -> bytes:
    """Drop the [Content_Types] Override lines for parts we removed."""
    for p in parts:
        xml = re.sub(
            rb'<Override PartName="/' + re.escape(p.encode()) + rb'"[^>]*/>',
            b"", xml)
    return xml


def export_pptx_locked(out_path, policy=None, **kw) -> Path:
    """`export_native.export_pptx_native`, then the padlock.

    Every keyword the unlocked export takes (region, country, reports,
    template, ...) is passed straight through, so the two produce the same
    deck and differ only in what can be read.
    """
    from .export_native import export_pptx_native, fit_to_slide
    out_path = Path(out_path)
    tmp = Path(tempfile.mkdtemp()) / "unlocked.pptx"
    # a locked deck never runs off a slide: a page too tall for one is split
    # over numbered slides, "(1/2)", "(2/2)" (see export_native.fit_to_slide)
    with fit_to_slide():
        export_pptx_native(tmp, **kw)
    return lock_deck(tmp, out_path, policy=policy)


def export_regions_zip_locked(out_zip, policy=None, **kw) -> Path:
    """`export_native.export_regions_zip`, then the padlock on every deck.

    The ZIP is built by the existing function first and re-written with each
    member locked, rather than by re-implementing its per-region loop - so the
    locked ZIP always holds exactly the decks the unlocked one does.
    """
    from .export_native import export_regions_zip, fit_to_slide
    out_zip = Path(out_zip)
    policy = policy or default_policy()
    work = Path(tempfile.mkdtemp())
    plain = work / "plain.zip"
    with fit_to_slide():                 # see export_pptx_locked
        export_regions_zip(plain, **kw)

    log = kw.get("progress") or (lambda msg: None)
    out_zip.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(plain) as src, \
            zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as dst:
        names = src.namelist()
        for i, name in enumerate(names, 1):
            member = work / f"{i}.pptx"
            with src.open(name) as fh, open(member, "wb") as fo:
                shutil.copyfileobj(fh, fo)
            if name.lower().endswith(".pptx"):
                log(f"  [{i}/{len(names)}] Locking {name}")
                lock_deck(member, policy=policy)
            dst.write(member, name)
    return out_zip


# ═══════════════ the dashboard preview of the locked deck ════════════
# The deck is locked by reading what it SHOWS.  The dashboard cannot be read
# that way - it is a component tree, not shapes - so the preview applies
# pptlock's OWN reading rules (`is_value`, `is_open_label`, `unit_of`,
# `prose_figures`, imported, never re-implemented) to that tree.  The four
# kinds of value it finds are the four the deck locks:
#
#     table cell    a <Td> in a column whose header is not an open label
#     chart label   a data label whose category is not an open label
#     shape value   a KPI card: one child is a lone figure, the rest its label
#     prose figure  a figure inside a sentence -> the deck's own "xx"
#
# Nothing here writes to the export path: a preview only rewrites the tree
# Dash is about to render, and the deck is locked from the finished .pptx.
_LOCK_ASSET = "lock.png"


def ensure_lock_assets() -> bool:
    """Put the padlock artwork where BOTH surfaces find it.

    `pptlock` carries the artwork embedded and unpacks it to a temp dir; Dash
    can only serve what is under `fmi_report/assets/`.  Copying it there gives
    the preview the very same PNG the deck uses - and `LockArt` looks in that
    folder BEFORE its own cache, so the deck now takes it from there too.
    Returns False if the artwork could not be found at all, in which case the
    preview falls back to the padlock character.
    """
    pl = _pptlock()
    dest = Path(__file__).resolve().parent / "assets"
    if (dest / _LOCK_ASSET).exists():
        return True
    art = pl.LockArt()
    if not art.available:
        return False
    try:
        dest.mkdir(parents=True, exist_ok=True)
        for p in Path(art.dir).glob("lock*.png"):
            shutil.copyfile(p, dest / p.name)
    except OSError:
        return False
    return (dest / _LOCK_ASSET).exists()


def _padlock(size: float, unit: str = ""):
    """The padlock as it stands in for a figure of ``size`` px, keeping its
    unit - the deck's rule, drawn with the deck's artwork."""
    pl = _pptlock()
    h = max(11.0, float(size) * pl.LOCK_IMG_K / 1.34)
    if ensure_lock_assets():
        mark = html.Img(src=f"/assets/{_LOCK_ASSET}", alt="locked",
                        style={"height": h, "width": h * pl.LOCK_ART_ASPECT,
                               "verticalAlign": "middle"})
    else:
        mark = html.Span("\U0001F512", style={"fontSize": h})
    kids = [mark]
    if unit:
        kids.append(html.Span(f" {unit}", style={"fontSize": max(9.0, h * .62),
                                                 "verticalAlign": "middle"}))
    return html.Span(kids, className="locked-value",
                     style={"display": "inline-flex", "alignItems": "center",
                            "gap": 1, "lineHeight": 1})


def _kids(node):
    """A component's children as a list, whatever shape they arrived in."""
    ch = getattr(node, "children", None)
    if ch is None:
        return []
    return list(ch) if isinstance(ch, (list, tuple)) else [ch]


def _text_of(node) -> str:
    """The plain text a node renders, or "" when it holds anything else."""
    if isinstance(node, str):
        return node
    if isinstance(node, (int, float)):
        return str(node)
    parts = [k for k in _kids(node) if isinstance(k, (str, int, float))]
    return " ".join(str(p) for p in parts) if parts else ""


def _tag(node) -> str:
    return type(node).__name__


def _mask_table(tbl, policy) -> None:
    """A <Table>: the header decides which columns stay readable.

    The same rule as `pptlock._scan_table` - column 0 is the row label and
    never a value, a column whose header `is_open_label` stays, and
    `totals_only` narrows even those to the total row.
    """
    rows = []
    for part in _kids(tbl):                       # Thead / Tbody, or bare Trs
        rows.extend([r for r in _kids(part) if _tag(r) == "Tr"]
                    if _tag(part) in ("Thead", "Tbody") else
                    ([part] if _tag(part) == "Tr" else []))
    if len(rows) < 2:
        return
    header = [_text_of(c).strip() for c in _kids(rows[0])]
    open_col = [policy.is_open_label(h) for h in header]
    for r, row in enumerate(rows[1:], start=1):
        cells = _kids(row)
        is_total = bool(cells) and policy.is_total_label(_text_of(cells[0]))
        for c, cell in enumerate(cells):
            if c == 0 or _tag(cell) != "Td":
                continue
            text = _text_of(cell).strip()
            if policy.is_protected(text) or not pl_is_value(text):
                continue
            if c < len(open_col) and open_col[c] and \
                    (not policy.totals_only or is_total):
                continue
            size = (cell.style or {}).get("fontSize", 10.5) \
                if getattr(cell, "style", None) else 10.5
            cell.children = _padlock(size, pl_unit_of(text))


def _mask_figure(fig, policy) -> None:
    """A Plotly figure: the data labels the chart already draws.

    `pptlock` only touches labels the deck ALREADY shows and keeps a point
    whose CATEGORY is an open label; both rules are kept here.  Hover is
    switched off with them - a preview that gives the figure back on hover
    would not be a preview of a locked report.
    """
    for tr in (fig.get("data") or []) if isinstance(fig, dict) else \
            list(getattr(fig, "data", ()) or ()):
        get = tr.get if isinstance(tr, dict) else \
            (lambda k, d=None, _t=tr: getattr(_t, k, d))
        horiz = str(get("orientation") or "") == "h"
        cats = get("labels") if get("labels") is not None else \
            (get("y") if horiz else get("x"))
        cats = list(cats) if cats is not None else []
        # a figure carried INSIDE the category label - see mask_cat_label
        if cats and all(isinstance(c, str) for c in cats):
            masked = [mask_cat_label(c) for c in cats]
            if any(m is not None for m in masked):
                cats = [m if m is not None else c
                        for c, m in zip(cats, masked)]
                key = "labels" if get("labels") is not None else \
                    ("y" if horiz else "x")
                if isinstance(tr, dict):
                    tr[key] = cats
                else:
                    setattr(tr, key, cats)
        texts = get("text")
        from_template = False
        if texts is None or isinstance(texts, str):
            # A label drawn from a TEMPLATE ("%{y}", "%{value}") carries no
            # text to rewrite, so the values are read out of the trace and put
            # back as an explicit label list - the open ones unchanged, the
            # rest padlocked.  Without this the chart would still print every
            # figure while its table twin was locked.
            if get("texttemplate") is None:
                continue
            vals = get("values") if get("values") is not None else \
                (get("x") if horiz else get("y"))
            if vals is None:
                continue
            texts = [None if v is None else f"{v:,.1f}"
                     if isinstance(v, float) else f"{v:,}"
                     if isinstance(v, int) else str(v) for v in vals]
            from_template = True
        out, hit = [], False
        for i, t in enumerate(texts):
            cat = str(cats[i]) if i < len(cats) else ""
            if t in (None, "") or policy.is_open_label(cat) \
                    or not pl_is_value(t):
                out.append(t)
                continue
            out.append("\U0001F512" + (f" {pl_unit_of(t)}"
                                       if pl_unit_of(t) else ""))
            hit = True
        if not hit:
            continue
        # a template-drawn label is re-pointed at the list just built; one that
        # already had text keeps drawing it, so the template is dropped
        tmpl = "%{text}" if from_template else None
        if isinstance(tr, dict):
            tr["text"] = out
            tr["hoverinfo"] = "skip"
            tr.pop("hovertemplate", None)
            if tmpl:
                tr["texttemplate"] = tmpl
            else:
                tr.pop("texttemplate", None)
        else:
            tr.text = out
            tr.hoverinfo = "skip"
            if hasattr(tr, "hovertemplate"):
                tr.hovertemplate = None
            if hasattr(tr, "texttemplate"):
                tr.texttemplate = tmpl


def _mask_kpi(node, policy) -> bool:
    """A KPI-style card: label plus ONE lone figure.

    `pptlock._shape_value_para`, in the tree: among at most three text
    children exactly one must read as a figure on its own, and the rest are
    its label - which is what says whether the figure is open.
    """
    pl = _pptlock()
    kids = _kids(node)
    texts = [(i, k) for i, k in enumerate(kids)
             if not isinstance(k, str) and _text_of(k).strip()]
    if not 2 <= len(texts) <= 3:
        return False
    hits = [i for i, (_j, k) in enumerate(texts)
            if pl._is_lone_figure(_text_of(k).strip())
            and not policy.is_protected(_text_of(k).strip())]
    if len(hits) != 1:
        return False
    vi = hits[0]
    value = texts[vi][1]
    label = " ".join(_text_of(k).strip()
                     for j, (_x, k) in enumerate(texts) if j != vi)
    if policy.is_open_label(label):
        return True                      # an open KPI: nothing to withhold
    txt = _text_of(value).strip()
    size = (value.style or {}).get("fontSize", 19) \
        if getattr(value, "style", None) else 19
    value.children = _padlock(size, pl_unit_of(txt))
    return True


def _mask_prose(s: str, policy) -> str:
    """A figure inside a sentence becomes the deck's own placeholder."""
    pl = _pptlock()
    marks = pl.prose_figures(s, policy.protect, policy.min_bare_int)
    for a, b, _whole, _num, unit in reversed(marks):
        s = s[:a] + policy.placeholder + (f" {unit}" if unit else "") + s[b:]
    return s


def pl_is_value(text) -> bool:
    return _pptlock().is_value(text)


def pl_unit_of(text) -> str:
    return _pptlock().unit_of(text)


def mask_tree(node, policy=None, copy: bool = True):
    """Withhold, in a Dash tree, exactly what the deck withholds.

    The tree is DEEP-COPIED first (only when previewing, which is off by
    default): the renderers build their components fresh per request today,
    but one cached figure or shared block would otherwise be padlocked for
    every viewer, unlocked ones included. Pass `copy=False` when the caller
    already owns the tree.
    """
    policy = policy or default_policy()
    if copy:
        node = deepcopy(node)
    if isinstance(node, (list, tuple)):
        return [mask_tree(k, policy, copy=False) for k in node]
    if isinstance(node, str):
        return _mask_prose(node, policy) if policy.lock_prose else node
    if not hasattr(node, "children") and not isinstance(node, dcc.Graph):
        return node

    tag = _tag(node)
    if tag == "Table":
        if policy.lock_tables:
            _mask_table(node, policy)
        return node
    if tag == "Graph":
        if policy.lock_charts and getattr(node, "figure", None) is not None:
            _mask_figure(node.figure, policy)
        return node
    if policy.lock_shape_values and _mask_kpi(node, policy):
        return node

    kids = getattr(node, "children", None)
    if kids is None:
        return node
    if isinstance(kids, (list, tuple)):
        node.children = [mask_tree(k, policy, copy=False)
                         for k in kids]
    else:
        node.children = mask_tree(kids, policy, copy=False)
    return node
