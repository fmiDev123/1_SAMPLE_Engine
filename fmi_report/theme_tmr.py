"""TMR (Transparency Market Research) THEME.

The FMI design system re-skinned to the TMR deck template
("Input/Changes/TMR/TMR Template.pptx"): the same page structure and
geometry as FMI, with TMR's navy / blue palette, its wordmark top-RIGHT,
the blue rule and copyright line along the foot, and its own front and last
pages (see front_last.py - they are the template's slides 1 and 294).

Selected with ``--template tmr``.  Only values that differ from FMI are set
here; everything else (geometry, component metrics, spacing) is FMI's.
"""
from __future__ import annotations

import datetime as _dt

from .theme_fmi import *            # noqa: F401,F403  (the FMI token set)
from . import theme_fmi as _f

NAME = "tmr"
LABEL = "TMR"

# ── Palette, as it occurs in the TMR template's slides ───────────────
TMR_NAVY   = "#0E2A5C"     # header chips / section bars
TMR_BLUE   = "#2A5F8F"     # theme accent1 - the working blue
TMR_BLUE_2 = "#26527E"
TMR_BLUE_3 = "#1F476B"
TMR_BRIGHT = "#0070C0"     # the rule along the foot, 'Thank You'
TMR_SKY    = "#0081E2"
TMR_SKY_LT = "#79C3EC"
TMR_TEAL   = "#00ACAC"
TMR_ORANGE = "#F47620"
TMR_GREY   = "#BCBEC0"
TMR_SLATE  = "#586272"
TMR_TINT   = "#EBF2F9"     # the pale card ground it uses throughout
TMR_EDGE   = "#C9D6E4"

PAGE_BG    = "#FFFFFF"
HEADER_BG  = TMR_NAVY
ACCENT     = TMR_BLUE
GOOD       = TMR_TEAL
BAD        = TMR_ORANGE
SEQ_LOW    = TMR_TINT
SEQ_HIGH   = TMR_BLUE
SEQ_STOPS  = [TMR_TINT, TMR_BLUE, TMR_NAVY]
HEAT_STOPS = [[0.00, "#FFFFFF"], [0.14, TMR_TINT], [0.32, "#CFE0F0"],
              [0.52, "#8FB7DC"], [0.70, TMR_SKY], [0.86, TMR_BLUE],
              [1.00, TMR_NAVY]]
DATA_COLORS = [TMR_BLUE, TMR_SKY, TMR_NAVY, TMR_TEAL, TMR_ORANGE, TMR_SKY_LT,
               TMR_SLATE, TMR_BLUE_3, TMR_GREY, "#3DC6C4", "#5B9BD5",
               "#A09F9D"]
SERIES_PALETTE = DATA_COLORS
KPI_ACCENTS = (TMR_NAVY, TMR_BLUE, TMR_TEAL, TMR_ORANGE)
KPI_TINT = TMR_TINT

PANEL_EDGE = TMR_EDGE
ZEBRA = "#F5F9FD"
TABLE_HEAD_BG = TMR_TINT
TABLE_HEAD_TEXT = TMR_NAVY
TABLE_TOTAL_BG = TMR_TINT
FLOW_LINE = TMR_SKY_LT
FLOW_LEAF_EDGE = TMR_SKY_LT
FLOW_LEAF_TEXT = TMR_NAVY
SIG_HEAD_BG = TMR_TINT
SIG_EDGE = TMR_EDGE
UI_TINT = TMR_TINT
UI_TINT_STRONG = "#D6E4F2"
TOC_TILE_BG = TMR_TINT
TOC_TILE_TEXT = TMR_NAVY
CHIP_FILL = "#D6E4F2"
CHIP_TEXT = TMR_NAVY
CHIP_EDGE = TMR_SKY_LT
CHIP_BORDER = TMR_SKY_LT
NAV_BTN_BG = TMR_TINT
NAV_BTN_TEXT = TMR_NAVY
NAV_BTN_EDGE = TMR_EDGE
TOC_HOVER = TMR_BLUE
SIDEBAR_DEEP = TMR_NAVY
BRAND_DEEP = TMR_NAVY

# ── Page chrome: white header, bold black title, NO rule under it ────
HEADER_STYLE = "rule"
HEADER_FILL = PAGE_BG
TITLE_COLOR = "#000000"
HEADER_RULE = None
HEADER_RULE_SEGMENTS = ()
HEADER_LABEL = TMR_SLATE
LOGO_RIGHT = True               # wordmark top-right, as in the TMR master
FOOTER_LINE = TMR_BRIGHT        # the blue rule along the foot of each page
FOOTER_TEXT_COLOR = "#000000"

# ── Brand identity ───────────────────────────────────────────────────
LOGO_FILE = "tmr_logo.png"
LOGO_FALLBACK = "tmr_logo.png"
BIG_LOGO_FILE = None            # no square logo: the wordmark as is
BIG_LOGO_DARK = None            # (on dark pages it sits on a white card)
LOGO_AR = 252 / 188
LOGO_ON_CARD = False
SOURCE_ORG = "Transparency Market Research"
_YEAR = _dt.date.today().year
SOURCE_TEXT = f"Source: TMR Analysis, {_YEAR}"
COPYRIGHT = f"©{_YEAR} Transparency Market Research, All Rights Reserved"

# ── Cover / section dividers ─────────────────────────────────────────
COVER_BG = "#FFFFFF"
COVER_GLOW = TMR_TINT
COVER_RING = TMR_SKY_LT
COVER_DOT = TMR_TINT
COVER_BAR = TMR_BLUE
COVER_ACCENT = TMR_BRIGHT
COVER_TEXT = "#000000"
COVER_SUB = TMR_SLATE
COVER_ON_LIGHT = True
TITLE_BG = ("#F3F8FD", "#DCE9F6")
TITLE_WAVES = ("#CFE0F0", "#A9C8E6", "#82AFDB")
TITLE_NUMBER = TMR_SKY
TITLE_KICKER = TMR_BLUE
TITLE_TEXT = TMR_NAVY
TITLE_LINE = TMR_SKY
TITLE_BAR = ("#6FA6D8", TMR_BLUE)
TITLE_STRIP = "#B7D1EA"
DIVIDER_PANEL = TMR_TINT
DIVIDER_TITLE = TMR_NAVY
DIVIDER_KICKER = TMR_BLUE
DIVIDER_MARK = TMR_BLUE
SECTION_BAND_TEXT = TMR_NAVY

# ── The deck's own theme part (the template's clrScheme) ─────────────
PPT_THEME_NAME = "Custom 63"
PPT_THEME_COLORS = {
    "dk2": "000000", "lt2": "F8F8F8",
    "accent1": "2A5F8F", "accent2": "9E7E0A", "accent3": "40315E",
    "accent4": "008189", "accent5": "A35313", "accent6": "5F5F5F",
    "hlink": "5F5F5F", "folHlink": "919191",
}
PPT_THEME_FONTS = ("Calibri Light", "Calibri")
