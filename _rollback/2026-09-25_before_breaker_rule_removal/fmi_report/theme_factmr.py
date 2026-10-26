"""Fact.MR report THEME - the Fact.MR visual design system as one swappable
module, the twin of `theme_fmi.py`.

WHERE THESE VALUES COME FROM
    Every colour here is the Fact.MR palette that used to live inline in
    `config.py` (commit 45412a9, tag `pre-fmi-theme-install-2026-08-12`),
    copied verbatim - lifted originally out of the PBIX `CY23SU08` theme and
    then hand-tuned by the user on 2026-07-20.  Nothing was re-picked or
    "improved" during the move: the Fact.MR report renders exactly as it did
    before this file existed.

WHY IT EXISTS
    The report now ships TWO templates, Fact.MR and FMI, selectable at runtime
    (see `theming.py`).  Both expose the SAME token names, so a renderer never
    asks which template is active - it just reads the token.  That is what keeps
    the "no `if fmi:` scattered through the codebase" rule enforceable.

    `config.py` re-exports whichever theme is active; `theming.apply()` rebinds
    these names into the consumer modules when the user flips the switch.

WHAT LIVES WHERE
    colours          brand + DATA_COLORS + semantic GOOD/BAD/NEUTRAL
    component tints  the values renderers used to hardcode inline
    brand identity   LOGO_FILE / SOURCE_TEXT / COPYRIGHT / PPT_FONT
    page chrome      HEADER_STYLE and the band/rule geometry
    typography       FONT_* , FS_*
    spacing scale    GAP, PAGE_MARGIN, CARD_PAD_*, PANEL_PAD, ...
    component metrics  line_px, wrap_lines, kpi_height, table_row_h, ...

UNITS
    The report canvas is 1280x720 PX; the deck is 960x540 PT - the same 16:9
    page at 1 px = 0.75 pt (`PT_PER_PX`).  Geometry tokens are in PX; the
    export converts.  Font tokens are in px (CSS) and converted for python-pptx.
"""
from __future__ import annotations

import math

NAME  = "factmr"
LABEL = "Fact.MR"

# ── Canvas geometry (shared by both templates, so pagination cannot drift) ──
CANVAS_W   = 1280
CANVAS_H   = 720
PT_PER_PX  = 0.75            # 960pt / 1280px

# ── PBIX canvas geometry (from the background SVGs) ──────────────────
PAGE_BG    = "#FFFFFF"       # page background - white (was light blue #C4E2F4)
HEADER_BG  = "#007BA5"       # title band - darkened brand cyan (white text)
PANEL_BG   = "#FFFFFF"
PANEL_EDGE = "#BFBFBF"

# ── Custom brand palette (user-supplied, 2026-07-20) ─────────────────
# No green: the old olive green is replaced with a medium blue (#2F6FB0) that
# stays in the palette's blue family and reads as growth/positive.
DATA_COLORS = ["#00A4DC", "#6ECFF2", "#D6CAC4", "#EBC46B", "#EDA28B",
               "#C4E2F4", "#9AB3C7", "#7D72E5", "#CF72E5", "#E572A9",
               "#ADADAD", "#2F6FB0", "#7CAFDD", "#F1975A"]
SERIES_PALETTE = DATA_COLORS

ACCENT   = "#00A4DC"         # tableAccent / first data colour
GOOD     = "#2F6FB0"         # growth / positive - blue (no green), vs pink decline
BAD      = "#E572A9"
NEUTRAL  = "#EBC46B"
SEQ_LOW  = "#DFF4FC"         # sequential 'minimum' - pale brand tint
SEQ_HIGH = "#00A4DC"         # sequential 'maximum'

# Deep brand shade, the top stop of the sequential map scale.
BRAND_DEEP = "#00526E"

# KPI card accent cycle.  Fact.MR cards were drawn in the flat accent; the
# cycle repeats it so the token exists for both themes with no visual change.
KPI_ACCENTS = (ACCENT, ACCENT, ACCENT, ACCENT)
KPI_TINT    = PANEL_BG        # Fact.MR KPI cards are white, as they always were

# ── YoY performance classification - SINGLE SOURCE OF TRUTH ───────────
# The same green / orange / red for the coloured cells AND their legend
# keywords (▲ outpacing / ▶ below pace / ▼ decline).  Dashboard + PPT.
#
# These are deliberately NOT the GOOD/BAD chart tokens.  GOOD is blue and BAD
# is pink here - a Power BI-matching choice for the chart SERIES - but the
# report's own prose names the marks out loud ("gains the most share ... green ▲",
# "cedes the most ... red ▼"), so a mark drawn in blue or pink contradicts the
# sentence it sits in.  The marks therefore carry a literal green and red,
# while GOOD/BAD keep colouring the series they were tuned for.
OUTPACE_COLOR = "#1E8E3E"     # ▲ outpacing / at-or-above market   (green)
BELOW_COLOR   = "#A87F1F"     # ▶ below market pace                (orange/amber)
DECLINE_COLOR = "#D64545"     # ▼ decline  - the same red as TREND_DOWN below
CLASS_COLORS  = (OUTPACE_COLOR, BELOW_COLOR, DECLINE_COLOR)

# ── 'Bar inside a bar' nested-bar chart ──────────────────────────────
NEST_OUTER = "#A9D8EF"        # wide background bar (secondary metric)
NEST_INNER = "#F1975A"        # narrow foreground bar (primary metric)

TEXT     = "#252423"         # theme foreground
MUTED    = "#605E5C"         # foregroundNeutralSecondary
FAINT    = "#ADADAD"
GRID     = "#E6E6E6"

# legacy aliases used by chart builders
BRAND, BRAND_LT, BRAND_DK = HEADER_BG, ACCENT, BRAND_DEEP
POS, NEG, ORANGE, CARD, BORDER = GOOD, BAD, "#F1975A", PANEL_BG, PANEL_EDGE
BG = PAGE_BG

# ═══════════════════════════════════════════════════════════════════════
#  COMPONENT TINTS - values the renderers used to hardcode inline
# ═══════════════════════════════════════════════════════════════════════
# Each of these was a literal hex sitting in figures.py / components.py /
# export_native.py / templates_*.py.  Moving them here is what makes the
# template switchable: the renderer reads the token, the theme decides the
# colour.  Every Fact.MR value below is the exact hex that was in the code.

ZEBRA          = "#F2FAFD"    # alternating table row fill
TABLE_HEAD_BG   = HEADER_BG   # table header row (brand blue)
TABLE_HEAD_TEXT = "#FFFFFF"   # white on the blue band, as before
TABLE_TOTAL_BG = "#E0F2FA"    # 'total' row / emphasised cell

FLOW_LINE      = "#8FB9CC"    # taxonomy connector lines
FLOW_LEAF_EDGE = "#6ECFF2"    # taxonomy leaf box border
FLOW_LEAF_TEXT = "#0F3D4C"    # taxonomy leaf box text

SIG_HEAD_BG    = "#CFEEFA"    # Market Signals heading strip
SIG_EDGE       = "#BFD9E6"    # Market Signals rules
SIG_ZEBRA      = "#F3FAFD"    # Market Signals alternating rows

# Impact arrows (competitor dashboard, market dynamics)
ARROW_HIGH     = "#2F6FB0"
ARROW_MED      = "#E2A828"
ARROW_LOW      = "#9AA0A6"

# Competitor visuals
STAR_COLORS    = ["#12558C", "#5B7C99", "#F2B739", "#6ECFF2"]
BOX_COLORS     = ["#F2B739", "#6ECFF2", "#12558C"]
NAVY_DEEP      = "#12558C"    # deep navy used by the competitor pages

# Value-chain / supply-chain stage colours
VC_PURPLE      = "#7A3FA0"
VC_NAVY        = "#0E3A5F"

# Diverging bars (BPS / price variation): positive vs negative
DIV_POS        = "#2F6FB0"
DIV_NEG        = "#E572A9"

# Waterfall / candlestick up-bar fill + line
UP_FILL        = "#DCE9F5"
UP_LINE        = "#2F6FB0"

# Price-strip / boxplot chrome
STRIP_START    = "#A9D8EF"
STRIP_LINE     = "#CBD8E1"

# Reference lines and their annotations
REF_LINE       = "#F1975A"

# Radar, 3-D funnel, target-vs-actual
RADAR_LINE     = "#2B5F8E"                  # brand navy-blue outline
RADAR_FILL     = "rgba(43,95,142,0.62)"     # translucent blue area
F3_CONE        = "#DCE9F5"
F3_BODY        = "#1E86C9"
F3_TOP         = "#5CB3E6"
F3_RIM         = "#0C5A8F"
F3_LBL         = "#1F3B52"
TA_TARGET      = "#0E2E52"
TA_ACTUAL      = "#9CC3E6"

# Trend / YoY line + marker accents
TREND_UP       = "#2F6FB0"
TREND_DOWN     = "#D64545"
OK_COLOR       = "#2F6FB0"
NOT_OK_COLOR   = "#D24B3E"

# Candlestick / price-range down-bars and the neutral outline
DN_FILL        = "#F6D2CE"
DN_LINE        = "#C0392B"
NEUTRAL_LINE   = "#5B7C99"

# Heatmap 'meter' gradient - single-hue sequential, white → deep.
# Perceptually ordered and colour-blind-safe (ColorBrewer 'Blues' family).
HEAT_STOPS     = [
    [0.00, "#FFFFFF"], [0.14, "#EAF2FB"], [0.32, "#CBE0F1"],
    [0.52, "#93C0E3"], [0.70, "#5798CE"], [0.86, "#25709F"],
    [1.00, "#003E58"],
]

# Text drawn INSIDE a coloured bar / segment
ON_FILL_TEXT   = "#FFFFFF"

AMBER_MARK     = "#F2B739"    # palette gold for marks / bars in the deck

# Small stateful marks on the competitor dashboard
STAR_EMPTY     = "#D4D7DA"    # unearned ✦
GAUGE_TRACK    = "#E4E7EA"    # unfilled arc of the share gauge
WARN_TEXT      = "#B06A2C"    # cautionary caption

# Pricing-pattern band + its three series
PP_BAND_FILL   = "#BFE0F2"
PP_BAND_LINE   = "#7FBEDD"
PP_COLORS      = ["#F2994A", "#2F6FB0", "#7D3C98"]

# Sequential map / heatmap scale stops (low → mid → high)
SEQ_STOPS      = ["#C4E2F4", ACCENT, BRAND_DEEP]


# Dashboard chrome (CSS)
TOC_HOVER      = "#00A4DC"    # Contents button hover fill
SIDEBAR_DEEP   = "#00526E"    # 'Export ALL regions' button

# -- Dashboard chrome tints (the left menu, buttons, nav bar) ----------
# These reach the browser as CSS custom properties, so flipping the template
# restyles the whole shell without a reload (see app._theme_vars).
UI_TINT        = "#E3F4FB"    # hover wash on menu links / nav buttons
UI_TINT_STRONG = "#CFEEFA"    # the ACTIVE menu link
UI_SUB_BG      = "#F3F8FB"    # sub-section heading strip
UI_SUB_RAIL    = "#BFDCEB"    # its left rail
UI_SUB_OPEN    = "#E8F3FA"    # sub-section heading when open
UI_SUB_LINE    = "#DCE7EE"    # hairline down the sub-section group

# -- Categorical palettes (multi-hue sets, no two neighbours alike) ----
# Used for segment sections, closing cards, the "Who Shaped the Year" timeline,
# the SWOT pinwheel and the value-chain stage ramp.  Every value is the exact
# set that was inline in the template modules.
SECTION_PALETTE  = ["#0E3A5F", "#1668A8", "#5B7C99", "#8E44AD", "#C58A21",
                    "#B5532F", "#2C6E7F"]
COUNTRY_PALETTE  = ["#1F6FB0", "#0E3A5F", "#7A4FB0", "#5B7C99", "#D98A2B",
                    "#C0504D", "#2C6E7F"]
CARD_PALETTE     = ["#0E3A5F", "#1668A8", "#5B7C99", "#C58A21"]
TIMELINE_PALETTE = ["#0E5AA7", "#E8833A", "#5B7C99", "#8A55C8", "#D9982A"]
SWOT_PALETTE     = {"S": "#5B7C99", "W": "#3E6E9E",
                    "O": "#F2C230", "T": "#E8833A"}
# The same four roles as SWOT_PALETTE, but dark enough to READ as text
# on the white page.  The palette itself still fills the pinwheel arms;
# only the heading, its rule and its bullets use these.
SWOT_TEXT        = {"S": "#4A6A85", "W": "#3E6E9E",
                    "O": "#A87F1F", "T": "#C4611F"}

DYN_HEAD_BG      = "#1F5C8B"   # Market Dynamics table header
RULE_TINT        = "#D6E4EE"   # hairline rules on the dynamics pages

# ── PPT cover + section divider ──────────────────────────────────────
# The cover is drawn, not imported: a layered brand background, concentric
# rings, an ascending-bar motif and an accent rule under the title.
COVER_BG       = "#003E58"    # deep brand ground
COVER_GLOW     = "#00688F"    # soft glow, bottom-left
COVER_RING     = "#4FB7DE"    # concentric ring outlines
COVER_DOT      = ACCENT       # filled disc inside the rings
COVER_BAR      = "#3E93BC"    # ascending bar motif
COVER_ACCENT   = ACCENT       # thin accent rule beside/under the title
COVER_SUB      = "#C4E7F5"    # subtitle text on the dark cover
ON_DARK_MUTED  = "#C4E2F4"    # secondary text on any dark slide
SECTION_BAND   = "#F2B739"    # 'Section N' bar on the divider slides
# Fact.MR's divider is unchanged: brand-blue panel, white title, gold bar.
# The names exist here only so the renderer never asks which template is on.
DIVIDER_BG     = "#DFF4FC"  # section divider page ground - pale brand cyan (SEQ_LOW)
DIVIDER_PANEL  = HEADER_BG
DIVIDER_TITLE  = "#FFFFFF"
DIVIDER_KICKER = "#BFE3F5"
DIVIDER_BODY   = "#EAF4FB"
DIVIDER_MARK   = "#F2B739"
SECTION_BAND_TEXT = "#333333"
COVER_TEXT     = "#FFFFFF"
COVER_ON_LIGHT = False
CHIP_EDGE      = "#6ECFF2"    # outline of the slicer chips in the title band

# ═══════════════════════════════════════════════════════════════════════
#  BRAND IDENTITY
# ═══════════════════════════════════════════════════════════════════════
# Asset file name inside fmi_report/assets/.  Fact.MR's wordmark is a wide
# lockup on a white ground, so on the coloured header band it sits on a small
# white card (LOGO_ON_CARD) rather than directly on the fill.
LOGO_FILE      = "factmr_logo.png"
LOGO_FALLBACK  = "fmi_logo_white.png"
LOGO_AR        = 1236 / 394            # Fact.MR wordmark w / h
LOGO_ON_CARD   = True

SOURCE_ORG     = "Fact.MR"
SOURCE_TEXT    = "Source: Fact.MR"     # caption under every data table
COPYRIGHT      = "© Fact.MR  ·  www.factmr.com  ·  Confidential"

# Optional .pptx to build the deck on top of.  Fact.MR decks are built from
# python-pptx's blank default, which is what produced every deck to date.
PPT_TEMPLATE   = None

# ── Typography ───────────────────────────────────────────────────────
# Fact.MR renders in Segoe UI on both surfaces (export_native.FONT).
FONT_UI            = "Segoe UI"
PPT_FONT           = "Segoe UI"

# Fact.MR's deck keeps the package's default theme part: its own colours are
# all explicit, and nothing about the PBIX report asks for a house scheme.
# `None` means "leave the theme alone" - the same token name as the FMI
# theme, so `export_native` never asks which template is active.
PPT_THEME_NAME   = None
PPT_THEME_COLORS = None
PPT_THEME_FONTS  = None
FONT_DISPLAY       = "Segoe UI Semibold"
FONT_STACK         = "'Segoe UI', system-ui, -apple-system, sans-serif"
FONT_DISPLAY_STACK = "'Segoe UI Semibold', 'Segoe UI', system-ui, sans-serif"

# Font scale - PX for the dashboard (canvas units); pt = px * PT_PER_PX.
FS_PAGE_TITLE  = 24          # the page title inside the coloured band
FS_PANEL_TITLE = 11          # matches layout.TITLE_FS as shipped
FS_BODY        = 16
FS_TABLE       = 14
FS_TABLE_HEAD  = 14
FS_KPI_LABEL   = 18
FS_KPI_VALUE   = 40
FS_KPI_SUB     = 15
FS_NOTE        = 13
FS_SOURCE      = 12
FS_PAGE_NO     = 13

# ── Page chrome ──────────────────────────────────────────────────────
# Fact.MR pages open with a filled colour BAND carrying white title text
# (components.page_header), where FMI uses a white zone with dark title text
# and a light rule underneath.  One token decides which, read in exactly two
# places: the Dash page header and the PPT slide header.
HEADER_STYLE   = "band"      # "band" (Fact.MR) | "rule" (FMI)
HEADER_FILL    = HEADER_BG   # the band is filled with the brand colour
TITLE_COLOR    = "#FFFFFF"   # white title text on the coloured band
HEADER_LABEL   = "rgba(255,255,255,0.85)"   # slicer labels inside the band
HEADER_RULE    = ""          # no rule: the fill itself separates header/body
HEADER_SHADOW  = "0 2px 5px rgba(0,0,0,0.25)"
TOC_TILE_BG    = HEADER_BG    # Fact.MR agenda tiles stay brand-filled
TOC_TILE_TEXT  = "#FFFFFF"
CHIP_FILL      = HEADER_BG    # the chip in the deck, exactly as before
CHIP_TEXT      = "#FFFFFF"    # Fact.MR chips keep white type on brand blue
NAV_BTN_BG     = HEADER_BG
NAV_BTN_TEXT   = "#FFFFFF"
NAV_BTN_EDGE   = "#FFFFFF"
CHIP_BG        = "rgba(255,255,255,0.15)"   # slicer chip inside the title zone
CHIP_BORDER    = "rgba(255,255,255,0.4)"
HEADER_H       = 75          # height of the filled band
BAND_X         = 13          # band inset from the page edge
BAND_W         = 1254
TITLE_X        = 28
TITLE_Y        = 10
TITLE_W        = 1040
TITLE_H        = 68
RULE_Y         = 88          # unused with a band; kept so both themes match
RULE_H         = 0
# The FMI template splits the rule under its title into coloured segments;
# Fact.MR fills the header band instead and draws no rule at all.  The token
# exists in both themes because the renderers read it unconditionally - empty
# means "no segments", which with RULE_H = 0 means "no rule".
HEADER_RULE_SEGMENTS = ()
LOGO_W         = 111
LOGO_H         = 62
LOGO_X         = CANVAS_W - LOGO_W - 20
LOGO_Y         = 12
FOOTER_RULE_Y  = 706
FOOTER_RULE_H  = 3
FOOTER_TEXT_Y  = 707
FOOTER_MARGIN  = 16
PAGE_NO_X      = 1232
# Fact.MR closes a page with a rule and a muted line of text, not with the
# FMI template's solid bar.  Same three token names, so no renderer branches.
FOOTER_BAND       = None
FOOTER_BAND_H     = 0
FOOTER_TEXT_COLOR = None      # None = the renderer's own muted colour

# ═══════════════════════════════════════════════════════════════════════
#  SPACING SCALE + COMPONENT METRICS
# ═══════════════════════════════════════════════════════════════════════
# The same contract `theme_fmi.py` documents: the layout engine reserves a box,
# the two renderers draw into it, and neither invents a padding or a line
# height - both call the helpers below, so drawn size == reserved size.
# The maths is deliberately identical in both themes; only the FS_* inputs
# differ, which is what makes the two templates read differently.

LINE_H       = 1.20          # text line box as a multiple of the font size
BULLET_LINE  = 1.40          # looser leading for bullet/body copy

GAP          = 14            # standard gap BETWEEN stacked components
GAP_X        = 16            # standard gap between side-by-side components
PAGE_MARGIN  = 15            # left/right page margin
CARD_PAD_X   = 14            # card inner padding, horizontal
CARD_PAD_Y   = 10            # card inner padding, vertical
PANEL_PAD    = 8             # panel body padding
TITLE_PAD_T  = 5             # panel heading strip: padding above the text
TITLE_PAD_B  = 4             # ...and below it
TITLE_RULE   = 3             # thickness of the heading bracket rules
FOOTER_CLEAR = 6             # clearance between content and the footer rule
HEADER_CLEAR = 10            # clearance between the brand rule and content
PANEL_BORDER = True
# Fact.MR keeps the PBIX report's own card: barely-rounded corners and a soft
# drop shadow.  Same token names as the FMI theme so no renderer branches.
# The 3px is the dashboard card's own softening, which the PBIX deck never
# had: Fact.MR slides draw SQUARE panels, so `PANEL_ROUNDED` is False and the
# deck is byte-for-byte the shape it always was.  Two tokens rather than one
# threshold, because "3px looks square enough" is a guess and this is a fact.
PANEL_RADIUS  = 3
PANEL_ROUNDED = False
PANEL_SHADOW  = "1px 2px 5px rgba(0,0,0,0.12)"          # titled panels carry a hairline card border


def line_px(font_px: float, leading: float = LINE_H) -> int:
    """Height of one rendered line of text, rounded up."""
    return int(math.ceil(font_px * leading))


def wrap_lines(text, width_px: float, font_px: float,
               bold: bool = False) -> int:
    """How many lines `text` needs at `font_px` inside `width_px`.

    Same average-glyph-width model as layout.text_width.  Deliberately
    slightly pessimistic: an over-estimate costs a little whitespace, an
    under-estimate costs an overlap."""
    n = len(str(text or ""))
    if not n or width_px <= 0:
        return 1
    per_line = max(1.0, width_px / (font_px * (0.50 if bold else 0.48)))
    return max(1, int(math.ceil(n / per_line)))


# ── How much a block may GROW when a page has spare height ───────────
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
    """Rendered height of one table body row at font size ``fs``."""
    return max(22, line_px(fs) + 8)


def table_head_h(fs: float = FS_TABLE) -> int:
    """Rendered height of the table header row."""
    return max(24, line_px(fs) + 10)


def kpi_value_size(width_px: float = 300, value: str = "") -> float:
    """The value font size a KPI card uses, from its WIDTH and the value's own
    length.  Never keyed off the card height - the height is computed FROM the
    size, so keying it off the height made the two disagree."""
    base = FS_KPI_VALUE if width_px >= 240 else FS_KPI_VALUE * 0.72
    text = str(value or "")
    if not text:
        return base
    inner = max(40.0, width_px - 2 * CARD_PAD_X)
    fs = base
    while fs > 20 and len(text) * fs * 0.68 > inner:
        fs -= 1.0
    return fs


def kpi_height(label="", sub="", width_px: float = 300,
               big: bool = False, value: str = "") -> int:
    """Height a KPI card needs for its actual content, wrap-aware."""
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
    """Back-compatible wrapper - size comes from the card WIDTH and the value's
    own length; ``card_h`` is ignored."""
    return kpi_value_size(card_w, value)


def panel_title_height(title="", width_px: float = 300) -> int:
    """Height of a panel's heading strip, including its bracket rule."""
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
