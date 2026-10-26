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

# ── Brand colours (theme colour scheme of the template) ──────────────
FMI_NAVY    = "#2A5F8F"      # accent1 - THE primary brand blue
FMI_BLUE    = "#0070C0"      # footer rule + dominant chart series
FMI_CYAN    = "#00B0F0"
FMI_AMBER   = "#FFC000"
FMI_MAROON  = "#8E1416"      # accent2
FMI_OCHRE   = "#9E7E0A"      # accent3
FMI_PURPLE  = "#40315E"      # accent4
FMI_TEAL    = "#008189"      # accent5
FMI_RUST    = "#A35313"      # accent6
FMI_NAVY_DK = "#1F476B"      # deep navy used in chart series
FMI_ORANGE  = "#FC8E00"
FMI_SKY     = "#7AABD8"      # light accent1 (master's tint strip)
FMI_GOLD    = "#FCD150"      # cover band
FMI_RULE    = "#9BC0E1"      # light rule under the page title (sampled)
FMI_EDGE    = "#D5DBDB"      # card / table border
FMI_INK     = "#1A2634"      # headings and card labels
FMI_NOTE    = "#C00000"      # red "Note: ... indicative" footnote

PAGE_BG    = "#FFFFFF"       # FMI pages are white
HEADER_BG  = FMI_NAVY        # brand blue: table headers, panel rules, chips
PANEL_BG   = "#FFFFFF"
PANEL_EDGE = FMI_EDGE

# ── Chart palette ────────────────────────────────────────────────────
# Ordered as the template's own charts use them (blue-led, then the theme
# accents), so a multi-series chart reads as an FMI chart from the first
# series.  figures._layout() feeds this to Plotly's `colorway` and
# export_native._palette() to the native charts - one list, both surfaces.
DATA_COLORS = [FMI_BLUE, FMI_NAVY, FMI_CYAN, FMI_AMBER, FMI_TEAL,
               FMI_RUST, FMI_PURPLE, FMI_SKY, FMI_OCHRE, FMI_ORANGE,
               FMI_MAROON, FMI_NAVY_DK, "#A08EC4", "#E894D0"]
SERIES_PALETTE = DATA_COLORS

# KPI card accent cycle - taken verbatim from the template's "At a Glance"
# cards (Market Size / CAGR / Forecast / Absolute Opportunity).  This is the
# ONE place the FMI template uses a green; set every entry to FMI_NAVY to go
# back to a mono-blue KPI band.
KPI_ACCENTS = ("#1A3A5C", "#1E6091", "#148E8E", "#1E7B4B")

ACCENT   = FMI_BLUE          # primary accent - first data colour
# Semantic growth/decline colours.  The template carries no YoY-classification
# legend, so the project's standing "no green for semantics" decision stands:
# blue = growth, amber = below pace, pink = decline.
GOOD     = FMI_BLUE
BAD      = "#E572A9"
NEUTRAL  = FMI_AMBER
SEQ_LOW  = "#E8F1F9"         # sequential 'minimum' - pale brand tint
SEQ_HIGH = FMI_BLUE          # sequential 'maximum'

# ── Typography ───────────────────────────────────────────────────────
# The template's theme is Calibri (minor) / Calibri Light (major); its display
# face is Bell MT, used ONLY on the cover and the section dividers.
FONT_UI            = "Calibri"
FONT_DISPLAY       = "Bell MT"
FONT_STACK         = "Calibri, 'Segoe UI', system-ui, sans-serif"
FONT_DISPLAY_STACK = "'Bell MT', 'Times New Roman', Georgia, serif"

# Font scale - PX for the dashboard (canvas units); pt = px * PT_PER_PX.
FS_PAGE_TITLE  = 32          # 24pt - the black page title, top-left
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
HEADER_H       = 84          # white header zone - no coloured band
TITLE_X        = 16          # 12pt
TITLE_Y        = 10
TITLE_W        = 1040
TITLE_H        = 68
RULE_Y         = 88          # light-blue rule sits directly under the header
RULE_H         = 8           # ~6pt
LOGO_W         = 111         # 1.795:1 - the wordmark's true aspect ratio
LOGO_H         = 62
LOGO_X         = CANVAS_W - LOGO_W - 20
LOGO_Y         = 12
FOOTER_RULE_Y  = 706         # blue rule across the foot of the page
FOOTER_RULE_H  = 3           # 2.25pt
FOOTER_TEXT_Y  = 707
FOOTER_MARGIN  = 16          # 12pt - the rule's left/right inset
PAGE_NO_X      = 1232

SOURCE_ORG = "Future Market Insights"

# ── YoY performance classification - SINGLE SOURCE OF TRUTH ───────────
# The same green / orange / red for the coloured cells AND their legend
# keywords (▲ outpacing / ▶ below pace / ▼ decline).  Dashboard + PPT.
OUTPACE_COLOR = GOOD          # ▲ outpacing / at-or-above market   (green)
BELOW_COLOR   = "#A87F1F"     # ▶ below market pace                (orange/amber)
DECLINE_COLOR = BAD           # ▼ decline                          (red)
CLASS_COLORS  = (OUTPACE_COLOR, BELOW_COLOR, DECLINE_COLOR)

# ── 'Bar inside a bar' nested-bar chart ──────────────────────────────
NEST_OUTER = FMI_SKY          # wide background bar (secondary metric)
NEST_INNER = FMI_ORANGE       # narrow foreground bar (primary metric)

TEXT     = "#252423"         # body copy
MUTED    = "#596369"         # secondary text (template's muted grey)
FAINT    = "#9AA3A8"
GRID     = "#E6E6E6"

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
ZEBRA          = "#EEF4FA"    # alternating table row fill
TABLE_HEAD_BG  = HEADER_BG    # table header row
TABLE_TOTAL_BG = "#DCE7F2"    # 'total' row / emphasised cell

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
DIV_POS        = FMI_BLUE
DIV_NEG        = BAD

# Waterfall / candlestick up-bar fill + line
UP_FILL        = "#DCE9F5"
UP_LINE        = FMI_BLUE

# Price-strip / boxplot chrome
STRIP_START    = FMI_SKY
STRIP_LINE     = "#CBD8E1"

# Reference lines and their annotations
REF_LINE       = FMI_ORANGE

# Radar, 3-D funnel, target-vs-actual
RADAR_LINE     = ACCENT
F3_CONE        = FMI_SKY
F3_BODY        = ACCENT
F3_TOP         = "#D5E5F3"
F3_RIM         = FMI_SKY
F3_LBL         = FMI_INK
TA_TARGET      = "#D5E5F3"
TA_ACTUAL      = ACCENT

# Trend / YoY line + marker accents
TREND_UP       = FMI_BLUE
TREND_DOWN     = FMI_MAROON
OK_COLOR       = FMI_BLUE
NOT_OK_COLOR   = FMI_MAROON

# Pricing-pattern band + its three series
PP_BAND_FILL   = "#F2F7FB"
PP_BAND_LINE   = "#DCE7F2"
PP_COLORS      = [FMI_ORANGE, FMI_NAVY, FMI_PURPLE]

# Sequential map / heatmap scale stops (low → mid → high)
SEQ_STOPS      = [SEQ_LOW, FMI_BLUE, FMI_NAVY_DK]

# PPT section divider band
SECTION_BAND   = FMI_NAVY

# Dashboard chrome (CSS)
TOC_HOVER      = FMI_BLUE     # Contents button hover fill
SIDEBAR_DEEP   = FMI_NAVY_DK  # 'Export ALL regions' button

# ── Brand identity ───────────────────────────────────────────────────
# The wordmark cropped to the header box from the template's slide master
# (298x166, ratio 1.795 - exactly the master's 98x55pt placement).  It sits on
# a white card because the artwork carries a white ground.
LOGO_FILE      = "fmi_logo_fmi.png"
LOGO_FALLBACK  = "fmi_logo_template.png"   # raw square master image
LOGO_AR        = 298 / 166
LOGO_ON_CARD   = True

SOURCE_TEXT    = "Source: Future Market Insights"
COPYRIGHT      = ("© Future Market Insights  ·  "
                  "www.futuremarketinsights.com  ·  Confidential")

# The deck is built natively (python-pptx blank default) rather than on top of
# the sample file - that carries 180 finished slides and 16 MB of media.  Kept
# here for provenance and for anyone who wants to diff against it.
PPT_TEMPLATE   = None
PPT_TEMPLATE_REF = (r"C:\Users\Shweta\Downloads\FMI Report Template"
                    r"\New_Global XYZ Market_Report_Sample_2026"
                    r"_Future market Insights_B2B.pptx")

PPT_FONT       = FONT_UI      # Calibri, the template's minor font

# ── Page chrome style ────────────────────────────────────────────────
# FMI pages open with a WHITE header zone: dark title text, light-blue rule
# underneath, logo top-right.  Fact.MR fills a coloured band with white title
# text instead.  One token decides which, read in exactly two places (the Dash
# page header and the PPT slide header).
HEADER_STYLE   = "rule"       # "band" (Fact.MR) | "rule" (FMI)
HEADER_FILL    = PAGE_BG      # the header zone is white
TITLE_COLOR    = FMI_INK      # dark title text on white
BAND_X         = 13           # band geometry kept so both themes match
BAND_W         = 1254

