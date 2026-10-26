"""FMI report THEME - the whole visual design system in one swappable module.

Everything the report looks like lives here: brand colours, the chart palette,
KPI accents, typography, page-chrome geometry, the spacing scale and the
component-metric helpers that the layout engine and both renderers share.

Extracted from the FMI B2B report template
  "New_Global XYZ Market_Report_Sample_2026_Future market Insights_B2B.pptx"
(theme colour scheme, slide master, layouts, and pixel samples of rendered
slides).

WHY THIS IS A SEPARATE MODULE
    `config.py` also handles dataset discovery, report years and geo mappings -
    concerns that have nothing to do with how the report LOOKS.  Keeping the
    design system here makes the template a single portable unit: to re-skin
    the report for another brand, copy this file, change the values, and point
    `config.py`'s re-export at it.  Nothing else in the codebase needs to move.

CONTRACT - the reason this works
    The layout engine (`layout.py`) reserves a box for every block; the two
    renderers (`components.py` for Dash, `export_native.py` for PowerPoint)
    draw into it.  **No renderer may invent a padding, a line height, a font
    size or a column width** - it calls a helper here, so the drawn size always
    equals the reserved size.  Each of the three used to carry its own copy of
    these numbers, which is exactly how components ended up overlapping.

UNITS
    The report canvas is 1280x720 PX; the deck is 960x540 PT - the same 16:9
    page at 1px = 0.75pt (`PT_PER_PX`).  Geometry tokens are in PX; the export
    converts.  Font tokens are given in px (CSS) and converted for python-pptx.

WHAT LIVES WHERE
    colours          FMI_* , DATA_COLORS, KPI_ACCENTS, semantic GOOD/BAD/...
    typography       FONT_* , FS_*
    page chrome      HEADER_H, RULE_*, TITLE_*, LOGO_*, FOOTER_*, PAGE_NO_X
    spacing scale    GAP, GAP_X, PAGE_MARGIN, CARD_PAD_*, PANEL_PAD, *_CLEAR
    component metrics  line_px, wrap_lines, kpi_height, kpi_value_size,
                       panel_title_height, summary_height, table_row_h,
                       table_head_h
    growth policy    STRETCH / stretch_factor / MAX_ROW_GAP
"""
from __future__ import annotations

import datetime as _dt
import math

# ═══════════════════════════════════════════════════════════════════════
#  FMI DESIGN SYSTEM - the single source of truth for the report's look
# ═══════════════════════════════════════════════════════════════════════
# Every value below was extracted from the FMI B2B report template
#   "New_Global XYZ Market_Report_Sample_2026_Future market Insights_B2B.pptx"
# (theme colour scheme, slide master, layouts, and pixel samples of rendered
# slides).  Both surfaces - the Dash dashboard (components.py) and the native
# PowerPoint export (export_native.py) - read these tokens, so the two cannot
# drift apart.  Change the brand here, not in the renderers.
#
# UNITS: the report canvas is 1280x720 PX, the deck is 960x540 PT - the same
# 16:9 page at 1 px = 0.75 pt.  Geometry tokens below are in PX (the canvas
# unit); export_native converts.  Font tokens are given in BOTH (px for CSS,
# pt for python-pptx) because that conversion is not automatic in text.

# ── Canvas geometry ──────────────────────────────────────────────────
CANVAS_W   = 1280
CANVAS_H   = 720
PT_PER_PX  = 0.75            # 960pt / 1280px

# ── Brand colours ────────────────────────────────────────────────────
# Read off the DROT template's own slide master, which parks its palette in
# swatch strips just outside the canvas: a CATEGORICAL column at x=-113, a
# five-step BLUE ramp and a five-step GREY ramp at x=1297.  That is the
# template stating its design system in its own file, so it is taken verbatim
# rather than sampled from rendered pixels.
# THE WORKING PALETTE, counted across all 69 reference slides - the colours a
# reader actually sees, not just the ones parked in the master.  Frequencies
# in brackets are how often each appears as a fill or a run colour.
# The DROT theme's own six accent slots ("Custom 70"), declared first because
# the working palette below aliases them.
FMI_THEME_1   = "#3D85C6"    # accent1 - the softer blue its charts use
FMI_TEAL_DROT = "#008189"    # accent4
FMI_GREY_MID  = "#808080"    # DROT leans on grey, not on colour
FMI_GREY_LT   = "#B2B2B2"
FMI_PAPER     = "#F8F8F8"    # theme lt2 - its pale ground

FMI_INK     = "#0B2E4A"      # [550] the deck's ink AND its darkest accent
FMI_BLUE    = "#0E72AE"      # [124] the accent blue
# The reference does carry a green [201 uses], but it is the one accent that
# sits outside its theme's own six slots, and beside teal it read as a second
# "good" colour.  Replaced by accent5, rust/brown, which keeps every category
# visually distinct and stays on the house theme.
FMI_GREEN   = "#A35313"      # theme accent5 - rust/brown (was #2E7D32)
FMI_OCHRE   = "#A86F00"      # [234] caution / second category
FMI_RUST    = "#A35313"      # [177]
FMI_RED     = "#C00000"      # [180] decline / warning
FMI_PURPLE  = "#6F3CC2"      # [61]
FMI_SLATE   = "#6C7A89"      # [60] secondary text
FMI_CORAL   = "#E44D3A"      # [12] the sparing hot accent
FMI_TEAL    = FMI_TEAL_DROT  # accent4 - the real teal, not the old green

# Each accent has a PALE CARD TINT it is set on - that pairing is the
# reference's signature, and the tints are its own values.
TINT_OCHRE  = "#FFF6D9"      # [163]
TINT_RUST   = "#FFF0E0"      # [134]
TINT_GREEN  = TINT_RUST      # follows its accent: rust's pale tint
TINT_RED    = "#FDECEC"      # [86]
TINT_PURPLE = "#F3EEFB"      # [64]
TINT_BLUE   = "#EAF6FB"      # [49]
TINT_SLATE  = "#EEF2F5"      # [45]
TINT_PANEL  = "#F5F8FA"      # [49] the quietest card ground
CARD_TINTS  = (TINT_BLUE, TINT_OCHRE, TINT_RUST, TINT_PURPLE,
               TINT_RED, TINT_SLATE)      # no repeat now that green is rust

# The slide master's parked ramps, kept for SEQUENTIAL scales (maps, heat
# meters) where a hue has to run dark to light.
FMI_BLUE_1  = "#2774B5"
FMI_BLUE_2  = "#3B82BF"
FMI_BLUE_3  = "#5AA0D4"
FMI_BLUE_4  = "#69B1E0"
FMI_BLUE_5  = "#79C3EC"
FMI_GREY_1  = "#58595B"
FMI_GREY_2  = "#808285"
FMI_GREY_3  = "#939598"
FMI_GREY_4  = "#BCBEC0"
FMI_GREY_5  = "#D1D3D4"

FMI_NAVY    = FMI_INK        # the deck's navy IS its ink
FMI_CHROME  = "#0667AC"      # the master's footer band and its rule
FMI_CYAN    = FMI_BLUE_5
FMI_AMBER   = FMI_OCHRE
FMI_MAROON  = FMI_RED
FMI_OLIVE   = "#7B8422"
FMI_ORANGE  = "#C55A11"
FMI_NAVY_DK = "#082238"      # one step under the ink, for deep grounds
FMI_SKY     = FMI_BLUE_4
FMI_GOLD    = "#FCD150"      # cover band
FMI_RULE    = FMI_BLUE_4
FMI_EDGE    = "#C6D5E1"      # [76] card hairline, 0.75pt
FMI_NOTE    = FMI_RED
# The master splits the rule under the page title into three coloured
# segments, one per content column.  These are its colours.
# The master literally draws these three rules in #92C73E / #FF4500 /
# #7030A0, but at 101 slides they were the loudest thing in the deck and read
# as leftover Office brights beside DROT's muted accents.  Re-cut to the
# template's OWN accent slots, which is what the rest of the page uses.
FMI_SEG_1   = "#008189"      # theme accent4, teal
FMI_SEG_2   = "#9E7E0A"      # theme accent2, dark gold
FMI_SEG_3   = "#40315E"      # theme accent3, deep purple

PAGE_BG    = "#FFFFFF"       # FMI pages are white
HEADER_BG  = FMI_NAVY        # brand blue: table headers, panel rules, chips
PANEL_BG   = "#FFFFFF"
PANEL_EDGE = FMI_EDGE

# ── Chart palette ────────────────────────────────────────────────────
# Ordered as the template's own charts use them (blue-led, then the theme
# accents), so a multi-series chart reads as an FMI chart from the first
# series.  figures._layout() feeds this to Plotly's `colorway` and
# export_native._palette() to the native charts - one list, both surfaces.
# Blue-led, then the master's categorical column in its own order, then the
# rest of the blue ramp - so a two-series chart is the template's two blues
# and a ten-series chart still never repeats a hue.
# The DROT theme's own accent slots ("Custom 70"), in its order: softer blue,
# dark gold, deep purple, teal, rust, grey.  Deliberately muted - no bright
# primary blue, and RED IS NOT IN THE SERIES SET because the reference keeps
# it for genuine negatives.
FMI_THEME_ACCENTS = ("#3D85C6", "#9E7E0A", "#40315E", "#008189",
                     "#A35313", "#5F5F5F")
DATA_COLORS = [FMI_THEME_1, "#9E7E0A", FMI_TEAL_DROT, "#40315E", FMI_RUST,
               FMI_GREY_MID, FMI_BLUE, FMI_OCHRE, FMI_INK, FMI_SLATE,
               FMI_GREY_LT, "#5F5F5F"]
SERIES_PALETTE = DATA_COLORS

# KPI card accent cycle - taken verbatim from the template's "At a Glance"
# cards (Market Size / CAGR / Forecast / Absolute Opportunity).  This is the
# ONE place the FMI template uses a green; set every entry to FMI_NAVY to go
# back to a mono-blue KPI band.
KPI_ACCENTS = (FMI_INK, FMI_BLUE, FMI_GREEN, FMI_OCHRE)
# DROT sets a dark accent on a PALE TINT rather than on white - that pairing
# is the reference's signature, so the KPI band takes the quiet blue tint.
KPI_TINT    = TINT_BLUE

ACCENT   = FMI_BLUE          # primary accent - first data colour
# Semantic growth/decline colours.  The template carries no YoY-classification
# legend, so the project's standing "no green for semantics" decision stands:
# blue = growth, amber = below pace, pink = decline.
# The reference is explicit about these: green for what is doing well, red
# for decline, ochre for caution - 201 / 180 / 234 runs.  This overturns the
# project's earlier "no green" decision, on the user's instruction to match
# the reference exactly.
GOOD     = FMI_TEAL_DROT     # the reference's teal carries "doing well"
BAD      = FMI_CORAL         # ...and its softer red carries decline
NEUTRAL  = FMI_OCHRE
SEQ_LOW  = TINT_BLUE         # sequential 'minimum' - the reference's pale blue
SEQ_HIGH = FMI_BLUE          # sequential 'maximum' - its accent blue

# ── Typography ───────────────────────────────────────────────────────
# The template's theme is Calibri (minor) / Calibri Light (major); its display
# face is Bell MT, used ONLY on the cover and the section dividers.
FONT_UI            = "Calibri"
FONT_DISPLAY       = "Bell MT"
FONT_STACK         = "Calibri, 'Segoe UI', system-ui, sans-serif"
FONT_DISPLAY_STACK = "'Bell MT', 'Times New Roman', Georgia, serif"

# Font scale - PX for the dashboard (canvas units); pt = px * PT_PER_PX.
FS_PAGE_TITLE  = 27          # 20pt - the DROT master's title size
# The template's panel headings are 11-12pt (15-16px).  The generator's panels
# are far denser than the template's 2-3-per-slide layout, so the heading is
# rendered at 13px: clearly larger and heavier than the old 11px, without
# re-flowing every page.  layout.TITLE_FS tracks this value.
FS_PANEL_TITLE = 13
FS_BODY        = 16          # 12pt - card bullets / body copy
FS_TABLE       = 14          # 10.5pt - table cells
FS_TABLE_HEAD  = 14          # 10.5pt
FS_KPI_LABEL   = 18          # 14pt bold - KPI card label
FS_KPI_VALUE   = 40          # 30pt bold - KPI card number (36pt at full size)
FS_KPI_SUB     = 15          # 11pt - KPI caption
FS_NOTE        = 13          # 10pt - the red indicative note
FS_SOURCE      = 12          # 9pt  - "Source@ Future Market Insights"
FS_PAGE_NO     = 13          # 10pt - page number, bottom-right

# ═══════════════════════════════════════════════════════════════════════
#  SPACING SCALE + COMPONENT METRICS - ONE source of truth for geometry
# ═══════════════════════════════════════════════════════════════════════
# The layout engine (layout.py) reserves a box for every block; the two
# renderers (components.py for Dash, export_native.py for PowerPoint) then draw
# into it.  Historically each of the three carried its OWN copy of the padding
# and line-height numbers, so a change to the type scale silently made the
# reserved boxes too small and components overlapped.
#
# Everything below is therefore DERIVED from the font scale above and consumed
# by all three.  Rule: no renderer may invent a padding or line height - it
# calls one of these helpers, so the drawn size always equals the reserved size.

LINE_H       = 1.20          # text line box as a multiple of the font size
BULLET_LINE  = 1.40          # looser leading for bullet/body copy

GAP          = 14            # standard gap BETWEEN stacked components
GAP_X        = 16            # standard gap between side-by-side components
PAGE_MARGIN  = 15            # left/right page margin
CARD_PAD_X   = 14            # card inner padding, horizontal
CARD_PAD_Y   = 10            # card inner padding, vertical
PANEL_PAD    = 8             # panel body padding
TITLE_PAD_T  = 5             # panel heading strip: padding above the text
TITLE_PAD_B  = 4             # ...and below it (the bracket rule sits under this)
TITLE_RULE   = 3             # thickness of the heading bracket rules
FOOTER_CLEAR = 6             # clearance between content and the footer rule
HEADER_CLEAR = 10            # clearance between the brand rule and content

# Titled panels carry a hairline card border.  The FMI template's
# "Title and Content" LAYOUTS use a bare bracket rule, and the migration
# applied that to every panel - but the template's own CONTENT slides (the
# At-a-Glance cards, the data tables) are bordered cards, and without a box
# a chart has nothing to sit in: plots float, axis labels run to the edge
# and neighbouring charts read as merged.  The border is what makes the
# grid read as a grid.  The blue centred heading and its bracket stay -
# inside the card.
PANEL_BORDER = True
# DROT draws its content cards as ROUNDED rectangles with a hairline and no
# drop shadow - 1,072 roundRects against 2,011 plain rects, and the plain ones
# are its bands and rules.  Both renderers read these two tokens, so the card
# looks the same on the dashboard and in the deck.
PANEL_RADIUS  = 8            # px corner radius on the dashboard card
PANEL_ROUNDED = True         # ...and the DECK draws the same rounded card
PANEL_SHADOW  = "none"       # DROT cards are outlined, never shadowed


def line_px(font_px: float, leading: float = LINE_H) -> int:
    """Height of one rendered line of text, rounded up."""
    return int(math.ceil(font_px * leading))


def wrap_lines(text, width_px: float, font_px: float,
               bold: bool = False) -> int:
    """How many lines `text` needs at `font_px` inside `width_px`.

    Uses the same average-glyph-width model as layout.text_width (Calibri is
    ~0.50 em bold / ~0.48 em regular).  Deliberately slightly pessimistic: an
    over-estimate costs a little whitespace, an under-estimate costs an
    overlap."""
    n = len(str(text or ""))
    if not n or width_px <= 0:
        return 1
    per_line = max(1.0, width_px / (font_px * (0.50 if bold else 0.48)))
    return max(1, int(math.ceil(n / per_line)))


# ── How much a block may GROW when a page has spare height ───────────
# A page rarely fills exactly.  The engine used to hand the leftover to every
# component equally, which inflated tables and text cards far beyond their
# content and left big dead zones inside their borders (the source caption
# ended up floating 250px under its own table).  A component may now only
# absorb spare height in proportion to how well it can USE it:
#   * charts and maps scale, so they take the lion's share
#   * tables, cards and text have a fixed content height - a little breathing
#     room reads as padding, more reads as a mistake
# Whatever is still left over is distributed as WHITESPACE (inter-row gaps and
# a centring offset), which is what a designed page does with slack.
STRETCH = {
    "figure":   1.00,        # charts absorb spare height freely
    "flowchart": 1.00,
    "writeup":  0.35,        # commentary can breathe a little
    "sumpanel": 0.25,        # editorial cards: a touch of padding only
    "table":    0.18,        # a table is its rows; padding beyond that is dead
    "kpi":      0.15,
}
STRETCH_DEFAULT = 0.30
MAX_ROW_GAP = 46             # cap on the gap leftover space may open up


def stretch_factor(kind: str) -> float:
    """Fraction of its own height a block may grow by when space is spare."""
    return STRETCH.get(kind, STRETCH_DEFAULT)


def table_row_h(fs: float = FS_TABLE) -> int:
    """Rendered height of one table body row at font size ``fs``.
    layout.table_size reserves this and components.df_table pins it, so the
    reserved height and the drawn height are the same number."""
    return max(22, line_px(fs) + 8)


def table_head_h(fs: float = FS_TABLE) -> int:
    """Rendered height of the table header row."""
    return max(24, line_px(fs) + 10)


def kpi_value_size(width_px: float = 300, value: str = "") -> float:
    """The value font size a KPI card uses, from its WIDTH and the value's own
    length.

    Two rules, in order:
      1. The size must NOT depend on the card's height - the height is computed
         FROM the size, so keying it off the height made the two disagree (the
         card budgeted a 28.8px number while the renderer drew a 40px one,
         pushing the caption out of the box).  Width is known first.
      2. The number must FIT its card.  The template's 36pt suits short values
         like "3.6%"; a workbook whose figures read "2,595 US$ Mn" needs less,
         or the number swamps the page.  The size steps down until the string
         fits the card's inner width - derived, so any market's number
         formatting is handled without a per-dataset rule.
    """
    base = FS_KPI_VALUE if width_px >= 240 else FS_KPI_VALUE * 0.72
    text = str(value or "")
    if not text:
        return base
    inner = max(40.0, width_px - 2 * CARD_PAD_X)
    fs = base
    # 0.68 em/char measured off the rendered card: KPI values are bold and
    # mix digits, separators and unit words ('2,595 US$ Mn'), which run
    # wider than plain digits
    while fs > 20 and len(text) * fs * 0.68 > inner:
        fs -= 1.0
    return fs


def kpi_height(label="", sub="", width_px: float = 300,
               big: bool = False, value: str = "") -> int:
    """Height a KPI card needs for its actual content.

    Was a fixed 57px, which the FMI type scale (18px label / 40px value / 15px
    caption) overflows by ~15px the 'cards merging' defect.  Now derived, and
    wrap-aware so a long label from another workbook grows the card instead of
    spilling out of it."""
    inner = max(40.0, width_px - 2 * CARD_PAD_X)
    h = 2 * CARD_PAD_Y
    h += wrap_lines(label, inner, FS_KPI_LABEL, bold=True) * line_px(
        FS_KPI_LABEL)
    h += 4 + line_px(FS_KPI_VALUE if big
                     else kpi_value_size(width_px, value), 1.15)
    if sub:
        h += 3 + wrap_lines(sub, inner, FS_KPI_SUB) * line_px(FS_KPI_SUB)
    return int(math.ceil(h))


def kpi_value_px(card_h: float, card_w: float = 300,
                 value: str = "") -> float:
    """Back-compatible wrapper - the size comes from the card WIDTH and the
    value's own length (see kpi_value_size); ``card_h`` is ignored."""
    return kpi_value_size(card_w, value)


def panel_title_height(title="", width_px: float = 300) -> int:
    """Height of a panel's heading strip, including its bracket rule.

    Wrap-aware: a title that needs two lines reserves two lines, so the chart
    or table below can never be pushed under it."""
    inner = max(40.0, width_px - 2 * PANEL_PAD - 18)      # 18 = focus icon
    lines = min(3, wrap_lines(title, inner, FS_PANEL_TITLE, bold=True))
    return int(TITLE_PAD_T + lines * line_px(FS_PANEL_TITLE)
               + TITLE_PAD_B + TITLE_RULE)


def summary_height(title="", bullets=(), width_px: float = 300) -> int:
    """Height an editorial summary card needs for its actual bullets."""
    inner = max(40.0, width_px - 2 * CARD_PAD_X)
    h = 2 * CARD_PAD_Y
    h += wrap_lines(title, inner, FS_PANEL_TITLE, bold=True) * line_px(
        FS_PANEL_TITLE) + 5 + 2 + 8            # text + pad + rule + margin
    for b in bullets:
        h += wrap_lines(b, inner - 16, FS_BODY - 4) * line_px(
            FS_BODY - 4, BULLET_LINE) + 6
    return int(math.ceil(h))

# ── Page chrome geometry (PX on the 1280x720 canvas) ─────────────────
# Mirrors the template's slide master:
#   white header 0,0 960x74.1pt · title 12,12 785.5x57pt · rule under it ·
#   logo 845.9,13 98.1x55.4pt · footer rule y=505.6pt · page no 914.4,505.8pt
# The template's own header is 74.1pt (99px) and its footer rule sits at
# 505.6pt (674px), leaving a 521px body.  The generator's pages are built for a
# 616px body - far denser than the template's 2-3-panels-per-slide sample - so
# adopting those numbers verbatim would re-paginate the whole report and push
# the bespoke full-page slides under the footer.  The header is therefore
# rendered at 84px (63pt) and the footer rule at 706px, which keeps the FMI
# elements, proportions and order intact while leaving the content box exactly
# 614px: byte-identical pagination, no page or slide gained or lost.
# Raise HEADER_H / lower FOOTER_RULE_Y here (layout.CONTENT_TOP and
# CONTENT_BOTTOM follow automatically) to go fully literal at the cost of
# re-flowing pages.
# 2026-09-11: the header grew from 84 to 96px so it can hold the client's
# heading standard - a bold title (up to two lines; 21pt since 2026-09-15)
# over a 14pt sub header.  layout.CONTENT_TOP moved down with it.
HEADER_H       = 96          # white header zone - no coloured band
TITLE_X        = 16          # 12pt
TITLE_Y        = 10
TITLE_W        = 1040
TITLE_H        = 68
RULE_Y         = 96          # the rule sits directly under the header
RULE_H         = 3           # 2.25pt - the DROT master's own rule weight
# THE DROT MASTER'S SIGNATURE: the rule under the page title is not one line
# but THREE coloured segments, one over each content column (green / orange /
# purple, at x=17 w=353, x=378 w=368, x=753 w=353 of a 1280 page).  Given as
# (start, width) FRACTIONS of the header band so the rule follows the band
# rather than assuming a page size.  An empty tuple means "one plain rule",
# which is what Fact.MR uses.
# 2026-09-11: the segments used to stop at ~87% of the band and left an empty
# stretch on the right; they now run edge to edge of the content area.
HEADER_RULE_SEGMENTS = ((0.000, 0.330, FMI_SEG_1),
                        (0.335, 0.330, FMI_SEG_2),
                        (0.670, 0.330, FMI_SEG_3))
LOGO_W         = 111         # 1.795:1 - the wordmark's true aspect ratio
LOGO_H         = 62
LOGO_X         = CANVAS_W - LOGO_W - 20
LOGO_Y         = 12
FOOTER_RULE_Y  = 706         # blue rule across the foot of the page
FOOTER_RULE_H  = 3           # 2.25pt
FOOTER_TEXT_Y  = 707
FOOTER_MARGIN  = 16          # 12pt - the rule's left/right inset
PAGE_NO_X      = 1232
# DROT closes every page with a SOLID blue bar carrying the copyright and the
# page number in white, with its own rule sitting on top of it.  The master's
# bar is 36px deep (y=684); this one starts at FOOTER_RULE_Y instead, because
# `layout.CONTENT_BOTTOM` is derived from that line and both templates share
# it - moving it would re-paginate the whole report and break the invariant
# that Fact.MR and FMI have identical page counts (HANDOVER §5a).
# 2026-09-11: the solid blue bar at the foot of every slide was removed on the
# client's request; the copyright line now sits on the white page, in the
# master's own dark grey so it stays readable.
FOOTER_BAND       = None            # None on a template with no footer bar
FOOTER_BAND_H     = CANVAS_H - FOOTER_RULE_Y
FOOTER_TEXT_COLOR = FMI_GREY_1      # dark grey on the white page

SOURCE_ORG = "Future Market Insights"

# ── YoY performance classification - SINGLE SOURCE OF TRUTH ───────────
# The same green / orange / red for the coloured cells AND their legend
# keywords (▲ outpacing / ▶ below pace / ▼ decline).  Dashboard + PPT.
#
# ▲ is NOT GOOD (the DROT teal) for the same reason Fact.MR's ▲ is not its
# blue: the report's prose says "green ▲" / "red ▼" in so many words, so the
# marks are literally green and red.  GOOD/BAD still colour the chart series.
# The green is muted to sit beside DROT's other accents rather than shout.
FMI_MARK_GREEN = "#2E7D32"    # deep, low-chroma green - reads green, stays calm
OUTPACE_COLOR = FMI_MARK_GREEN  # ▲ outpacing / at-or-above market   (green)
BELOW_COLOR   = FMI_OCHRE     # ▶ below market pace  - the reference's own
DECLINE_COLOR = FMI_CORAL     # ▼ decline - the reference's softer red
CLASS_COLORS  = (OUTPACE_COLOR, BELOW_COLOR, DECLINE_COLOR)

# ── 'Bar inside a bar' nested-bar chart ──────────────────────────────
NEST_OUTER = FMI_SKY          # wide background bar (secondary metric)
NEST_INNER = FMI_ORANGE       # narrow foreground bar (primary metric)

# DROT's body copy is NEAR-BLACK on white [3,365 uses], not navy.  Setting
# TEXT to the navy made it the most-used colour in the whole deck (4,455
# uses, more than white) and is most of why our pages read heavy.
TEXT     = "#2E2E2E"         # [3365] body copy
MUTED    = FMI_SLATE         # [60] secondary text
FAINT    = "#99A6B0"
GRID     = "#D7E0EA"        # [68] the reference's own hairline grey

# legacy aliases used by chart builders
BRAND, BRAND_LT, BRAND_DK = HEADER_BG, ACCENT, FMI_NAVY_DK
POS, NEG, ORANGE, CARD, BORDER = GOOD, BAD, FMI_ORANGE, PANEL_BG, PANEL_EDGE
BG = PAGE_BG

# ═══════════════════════════════════════════════════════════════════════
#  DUAL-TEMPLATE CONTRACT  (added 2026-08-12)
# ═══════════════════════════════════════════════════════════════════════
# The report ships TWO templates, Fact.MR and FMI, selectable at runtime.
# `theme_factmr.py` is this file's twin and defines the SAME names, so a
# renderer never asks which template is active - it just reads the token.
# Everything below either names this theme or gives an FMI value for a colour
# that the renderers used to hardcode inline.

NAME  = "fmi"
LABEL = "FMI"

BRAND_DEEP = FMI_NAVY_DK     # deep brand shade, top stop of the map scale

# ── Component tints (were literal hexes inside the renderers) ────────
# DROT builds its content out of ROUNDED cards: a white or pale-tinted fill,
# a 0.75pt blue-grey hairline, no drop shadow.  The tints below are its own
# (measured across 69 slides: #EDF3FA, #F5FAFF, #F7FAFC for the blue family).
ZEBRA          = TINT_PANEL   # [49] alternating table row fill
# DROT's tables carry NO dark header block: the header row is a pale band
# with dark type, which is most of why its pages read light where ours read
# heavy.  Two tokens, so the renderer never asks which template is active.
TABLE_HEAD_BG   = "#E0E8EF"   # [62] pale band, not a navy block
TABLE_HEAD_TEXT = FMI_INK     # dark type on the pale band
TABLE_TOTAL_BG = "#E0E8EF"    # [62] 'total' row / emphasised cell

FLOW_LINE      = FMI_SKY      # taxonomy connector lines
FLOW_LEAF_EDGE = FMI_SKY      # taxonomy leaf box border
FLOW_LEAF_TEXT = FMI_INK      # taxonomy leaf box text

SIG_HEAD_BG    = "#EAF1F8"    # Market Signals heading strip
SIG_EDGE       = FMI_EDGE     # Market Signals rules
SIG_ZEBRA      = "#F6F9FC"    # Market Signals alternating rows

# Impact arrows (competitor dashboard, market dynamics)
ARROW_HIGH     = FMI_BLUE
ARROW_MED      = FMI_OCHRE
ARROW_LOW      = "#9AA3A8"

# Competitor visuals
STAR_COLORS    = [FMI_NAVY, FMI_SKY, FMI_AMBER, FMI_CYAN]
BOX_COLORS     = [FMI_AMBER, FMI_CYAN, FMI_NAVY]
NAVY_DEEP      = FMI_NAVY_DK

# Value-chain / supply-chain stage colours
VC_PURPLE      = FMI_PURPLE
VC_NAVY        = FMI_NAVY

# Diverging bars (BPS / price variation): positive vs negative
# Diverging bars (BPS / price variation): the reference's teal against its
# softer red - never bright blue against pure red, which read as a warning
# on every page.
DIV_POS        = FMI_TEAL_DROT
DIV_NEG        = FMI_CORAL

# Waterfall / candlestick up-bar fill + line
UP_FILL        = "#DCEAEA"    # pale teal, the up-bar body
UP_LINE        = FMI_TEAL_DROT

# Price-strip / boxplot chrome
STRIP_START    = FMI_SKY
STRIP_LINE     = "#CBD8E1"

# Reference lines and their annotations
REF_LINE       = FMI_GREY_MID  # the dashed baseline recedes

# Radar, 3-D funnel, target-vs-actual
RADAR_LINE     = FMI_NAVY                        # accent1 outline
RADAR_FILL     = "rgba(42,95,143,0.62)"          # accent1 at 62%
F3_CONE        = "#DCE9F5"
F3_BODY        = FMI_BLUE
F3_TOP         = FMI_SKY
F3_RIM         = FMI_NAVY_DK
F3_LBL         = FMI_INK
TA_TARGET      = FMI_NAVY_DK
TA_ACTUAL      = FMI_SKY

# Trend / YoY line + marker accents
TREND_UP       = FMI_TEAL_DROT
TREND_DOWN     = FMI_CORAL
OK_COLOR       = FMI_TEAL_DROT
NOT_OK_COLOR   = FMI_CORAL

# Candlestick / price-range down-bars and the neutral outline
DN_FILL        = "#F7E3DF"    # pale coral, the down-bar body
DN_LINE        = FMI_CORAL
NEUTRAL_LINE   = "#7C8B96"

# Heatmap 'meter' gradient - single-hue sequential, running from the
# reference's own pale blue tint up through its accent blue to its ink, so a
# heat scale introduces no colour the reference does not use.
HEAT_STOPS     = [
    [0.00, "#FFFFFF"], [0.14, TINT_BLUE], [0.32, "#CBDDEE"],
    [0.52, "#8FBBDA"], [0.70, "#3E8FC4"], [0.86, FMI_BLUE],
    [1.00, FMI_INK],
]

# Text drawn INSIDE a coloured bar / segment
ON_FILL_TEXT   = "#FFFFFF"

AMBER_MARK     = FMI_AMBER    # palette gold for marks / bars in the deck

# Small stateful marks on the competitor dashboard
STAR_EMPTY     = "#D8DDE1"    # unearned ✦
GAUGE_TRACK    = "#E4E9ED"    # unfilled arc of the share gauge
WARN_TEXT      = FMI_RUST     # cautionary caption

# Pricing-pattern band + its three series
PP_BAND_FILL   = "#C9DCEE"
PP_BAND_LINE   = FMI_SKY
PP_COLORS      = [FMI_ORANGE, FMI_NAVY, FMI_PURPLE]

# Sequential map / heatmap scale stops (low → mid → high)
SEQ_STOPS      = [SEQ_LOW, FMI_BLUE, FMI_NAVY_DK]


# Dashboard chrome (CSS)
TOC_HOVER      = FMI_BLUE     # Contents button hover fill
SIDEBAR_DEEP   = FMI_NAVY_DK  # 'Export ALL regions' button

# -- Dashboard chrome tints (the left menu, buttons, nav bar) ----------
# Same roles, in the navy family.
UI_TINT        = "#EAF1F8"    # hover wash on menu links / nav buttons
UI_TINT_STRONG = "#D9E6F2"    # the ACTIVE menu link
UI_SUB_BG      = "#F2F6FA"    # sub-section heading strip
UI_SUB_RAIL    = FMI_RULE     # its left rail
UI_SUB_OPEN    = "#E6EEF6"    # sub-section heading when open
UI_SUB_LINE    = "#DEE6ED"    # hairline down the sub-section group

# -- Categorical palettes (multi-hue sets, no two neighbours alike) ----
# The same roles as Fact.MR's sets, drawn from the template's own accents.
SECTION_PALETTE  = [FMI_INK, FMI_BLUE, FMI_TEAL, FMI_PURPLE, FMI_OCHRE,
                    FMI_RUST, FMI_SLATE]
COUNTRY_PALETTE  = [FMI_BLUE, FMI_INK, FMI_PURPLE, FMI_GREEN, FMI_OCHRE,
                    FMI_RED, FMI_SLATE]
CARD_PALETTE     = [FMI_NAVY_DK, FMI_BLUE, FMI_SKY, FMI_OCHRE]
TIMELINE_PALETTE = [FMI_BLUE, FMI_RUST, FMI_SKY, FMI_PURPLE, FMI_OCHRE]
SWOT_PALETTE     = {"S": FMI_SKY, "W": FMI_NAVY,
                    "O": FMI_AMBER, "T": FMI_RUST}
# The same four roles as SWOT_PALETTE, but dark enough to READ as text
# on the white page.  The palette itself still fills the pinwheel arms;
# only the heading, its rule and its bullets use these.
SWOT_TEXT        = {"S": FMI_NAVY, "W": FMI_NAVY_DK,
                    "O": FMI_OCHRE, "T": FMI_RUST}

DYN_HEAD_BG      = FMI_NAVY    # Market Dynamics table header
RULE_TINT        = FMI_RULE    # hairline rules on the dynamics pages

# ── PPT cover + section divider ──────────────────────────────────────
# Same drawn cover, in the template's navy + gold.
COVER_BG       = FMI_PAPER    # DROT covers are LIGHT, not navy
COVER_GLOW     = TINT_BLUE    # soft wash, bottom-left (light cover)
COVER_RING     = FMI_GREY_LT  # ring outlines, quiet on white
COVER_DOT      = TINT_BLUE    # filled disc inside the rings
COVER_BAR      = FMI_THEME_1  # ascending bar motif
COVER_ACCENT   = FMI_BLUE     # accent rule beside the title
# DROT has no bright yellow bar: its section pages are light with dark text,
# so the band becomes a pale rule-like strip.
SECTION_BAND   = "#EEF2F5"
# The section divider, as roles rather than as literals in the renderer.
# DROT's section pages are LIGHT: a pale panel, dark title, grey body, and a
# quiet strip carrying "Section N" - no navy block, no yellow bar.
DIVIDER_BG     = "#D9ECF8"  # section divider page ground - light tint of the FMI blues
DIVIDER_PANEL  = FMI_PAPER      # the big panel behind the section title
DIVIDER_TITLE  = FMI_INK        # dark title on the pale panel
DIVIDER_KICKER = FMI_BLUE
DIVIDER_BODY   = TEXT
DIVIDER_MARK   = FMI_OCHRE      # the bullet marker
SECTION_BAND_TEXT = FMI_INK
COVER_TEXT     = FMI_INK        # a light cover needs dark type
COVER_ON_LIGHT = True
COVER_SUB      = FMI_SLATE    # subtitle - grey on the LIGHT cover
ON_DARK_MUTED  = "#C9DCEE"    # secondary text on any dark slide
CHIP_EDGE      = FMI_SKY      # outline of the slicer chips in the title band

# ── Brand identity ───────────────────────────────────────────────────
# The wordmark cropped to the header box from the template's slide master
# (298x166, ratio 1.795 - exactly the master's 98x55pt placement).  It sits on
# a white card because the artwork carries a white ground.
LOGO_FILE      = "fmi_logo_fmi.png"
LOGO_FALLBACK  = "fmi_logo_template.png"   # raw square master image
LOGO_AR        = 298 / 166
LOGO_ON_CARD   = True

# Both lines are the DROT template's own wording.  The year is taken from the
# clock rather than typed, so a deck built next year does not claim this one.
_YEAR = _dt.date.today().year
SOURCE_TEXT    = f"Source: FMI Analysis, {_YEAR}"
COPYRIGHT      = f"©{_YEAR} Future Market Insights, All Rights Reserved"

# The deck is built natively (python-pptx blank default) rather than on top of
# the sample file - that carries 180 finished slides and 16 MB of media.  Kept
# here for provenance and for anyone who wants to diff against it.
PPT_TEMPLATE   = None
PPT_TEMPLATE_REF = (r"C:\Users\Shweta\Downloads\FMI Report Template"
                    r"\New_Global XYZ Market_Report_Sample_2026"
                    r"_Future market Insights_B2B.pptx")

PPT_FONT       = FONT_UI      # Calibri, the template's minor font

# ── The deck's OWN theme part ────────────────────────────────────────
# python-pptx builds from a default package carrying Microsoft's "Office"
# scheme (accent1 4F81BD, accent2 C0504D, ...).  Every colour this project
# draws is explicit, so that scheme is invisible - until someone adds a shape
# or a chart in PowerPoint afterwards, which then inherits the OLD brights.
# `export_native._install_theme()` rewrites the theme part with the values
# below, so the file a client opens is on the template's real house theme.
PPT_THEME_NAME = "Custom 70"
PPT_THEME_COLORS = {          # DROT's own clrScheme, verbatim
    "dk2": "000000", "lt2": "F8F8F8",
    "accent1": "3D85C6", "accent2": "9E7E0A", "accent3": "40315E",
    "accent4": "008189", "accent5": "A35313", "accent6": "5F5F5F",
    "hlink": "5F5F5F", "folHlink": "919191",
}
PPT_THEME_FONTS = ("Calibri Light", "Calibri")     # major, minor

# ── Page chrome style ────────────────────────────────────────────────
# FMI pages open with a WHITE header zone: dark title text, light-blue rule
# underneath, logo top-right.  Fact.MR fills a coloured band with white title
# text instead.  One token decides which, read in exactly two places (the Dash
# page header and the PPT slide header).
HEADER_STYLE   = "rule"       # "band" (Fact.MR) | "rule" (FMI)
HEADER_FILL    = PAGE_BG      # the header zone is white
TITLE_COLOR    = "#000000"    # DROT sets the page title in black
HEADER_LABEL   = MUTED        # slicer labels read on white
HEADER_RULE    = FMI_RULE     # light-blue rule under the title
HEADER_SHADOW  = "none"       # the rule does the separating, not a shadow
# The slicer chips and the ⌂ / ‹ / › buttons were filled with the brand
# colour outright - the last solid dark blocks left in a DROT-light header.
CHIP_BG        = UI_TINT_STRONG   # slicer chip inside the title zone
# `CHIP_BG` is a CSS value (Fact.MR uses rgba()), so the DECK needs its own
# hex: same role, a form python-pptx can take.
# The Agenda's navigation tiles were solid brand blocks - a whole page of
# dark on both surfaces.  DROT would set these as pale cards with dark type.
TOC_TILE_BG    = "#E0E8EF"
TOC_TILE_TEXT  = FMI_INK
CHIP_FILL      = "#D9E6F2"        # the chip in the deck
CHIP_TEXT      = FMI_INK          # dark type on the pale chip
NAV_BTN_BG     = TINT_SLATE       # ⌂ / ‹ / › buttons, quiet on white
NAV_BTN_TEXT   = FMI_INK
NAV_BTN_EDGE   = FMI_EDGE
CHIP_BORDER    = FMI_RULE
BAND_X         = 13           # band geometry kept so both themes match
BAND_W         = 1254

