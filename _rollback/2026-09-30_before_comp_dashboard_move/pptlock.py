"""pptlock - put the padlock on the values a client must not read.

A SINGLE, SELF-CONTAINED file.  It imports nothing from this project, so it
can be dropped straight into whichever engine produces your decks:

    from pptlock import LockPolicy, lock_pptx
    lock_pptx("report.pptx", "report_locked.pptx",
              LockPolicy(open_years={2025, 2026, 2036}))

or from the command line:

    python pptlock.py report.pptx report_locked.pptx --open-years 2025,2026,2036
    python pptlock.py report.pptx --dry-run          # list what WOULD be locked

Its only hard dependency is ``python-pptx``.


WHAT IT DOES
------------
The deck that generates a report knows what a number MEANS - that a cell is
"2029 x North America" - so it can decide what to withhold.  A finished
``.pptx`` carries no such model: it is shapes, tables, charts and runs of
text.  So this module re-derives the decision from what the deck itself
SHOWS, which is why it works in two steps you can put a human between:

    scan(deck, policy)   -> [Candidate, ...]   what it proposes to withhold
    apply(deck, out, decisions=...)            what you told it to withhold

``lock_pptx`` is the two of them in one call, with no review.

Four kinds of value are found, and each is withheld in the way its container
allows:

    table cell    a PowerPoint cell cannot hold a picture, so the text is
                  cleared and the padlock is placed OVER the cell, whose
                  rectangle follows from the table's own column widths and
                  row heights
    chart label   a native chart draws its labels from a number format, so
                  the padlock arrives as the label box's PICTURE FILL (or,
                  with no artwork available, as an "xx" number format)
    shape value   a KPI-style shape whose whole text is one figure: text
                  cleared, padlock placed over the shape
    prose figure  a figure INSIDE a sentence.  DrawingML has no inline
                  picture - a lock in running text can only be a character,
                  and PowerPoint draws the emoji as a different padlock from
                  the artwork.  So these become "xx" / "xx%" placeholders,
                  which is the convention the reference decks already use.

WHAT IT CANNOT DO
-----------------
* A chart or table pasted as an IMAGE is opaque; it is reported and skipped.
* Row geometry is read from the table's own ``<a:tr h="...">``.  PowerPoint
  grows a row when its text needs more room, so on a deck whose rows are
  auto-sized the padlock can sit a pixel or two off centre.  Clearing the
  text can only make a row shorter, never taller, so it never drifts down.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.presentation import Presentation as _PrsClass
from pptx.oxml.ns import nsdecls, qn
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt
from pptx.oxml import parse_xml

__all__ = ["LockPolicy", "Candidate", "LockArt", "scan", "apply_locks",
           "lock_pptx",
           # the reading rules and the padlock's geometry, which a report
           # ENGINE can import so the two never drift (see below)
           "is_value", "unit_of", "is_open_label", "prose_figures",
           "LOCK_ART", "LOCK_STOPS", "LOCK_KEYHOLE", "LOCK_EDGE", "LOCK_RIM",
           "LOCK_SHADOW", "LOCK_ART_SCALE", "LOCK_ART_ASPECT", "LOCK_IMG_K"]

EMU_PER_PX = 9525                      # 1280 px canvas == 13.333 in slide


# ═════════════════════════ what counts as a value ════════════════════
# These are the report's own reading rules, kept here so the module stands
# alone.  A value STARTS with a number - after any YoY marker, sign or
# currency symbol - which keeps every real figure ("1,969", "▲ 4.5%",
# "5-10%", "1.29x") and excludes the descriptive cells that merely mention
# one ("Tier 1", "1.1% Sodium Fluoride").
_VALUE_RE = re.compile(r"^[▲▶▼△▽↑↓+\-±~<>≈≥≤\s]*[$€£¥₹]?\s*\d")
_YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")
_CAGR_RE = re.compile(r"cagr", re.I)
# "CAGR Index" is an index BUILT from CAGR, not a rate - it is a value like
# its sibling "Absolute Index" and locks the same way.  Indexing beats the
# CAGR exemption.
_INDEX_RE = re.compile(r"\bindex\b", re.I)
_BLANKS = {"", "-", "–", "—", "n/a", "na", "nan", "none", "·"}
_UNIT_PCT_RE = re.compile(r"%\s*$")
_UNIT_K_RE = re.compile(r"(?:\d|^)\s?([KMB]n?)\s*$")

# A figure inside a sentence: an optional currency symbol, digits, and
# whatever unit trails them.  The thousands separator must be FOLLOWED by
# three digits - "[\d,]*" would swallow the comma after "2036," and leave a
# year looking like "2036,", which no longer reads as a year and would be
# withheld (taking the sentence's comma with it).
_NUM = r"[$€£¥₹]?\d+(?:,\d{3})*(?:\.\d+)?"
# A RANGE reads as one figure - "51-57%" is a single share, and withholding
# only its right-hand half ("51-🔒%") both looks broken and gives the left
# half away.
_PROSE_FIG = re.compile(
    rf"(?<![\w.])({_NUM}(?:\s*[-–—]\s*{_NUM})?)(\s*[%×x])?")
_NUM_PART = re.compile(r"\d+(?:,\d{3})*(?:\.\d+)?")
_QUOTED = re.compile(r"'[^']*'")
# a figure that is part of a hyphenated word is part of a NAME, not a
# measurement - "5000-ppm", "24-hour", "3-in-1"
_HYPHENATED = re.compile(r"[-‑–]\s*[A-Za-z]")
# A UNIT standing after the figure.  `min_bare_int` calls a small bare
# integer a count - "6 regions", "5 product type categories" - but a unit
# says it is a measurement however small it is, and share shifts are quoted
# in basis points: "gains the most share +86 BPS" was published in full on
# every share-shift page because 86 is under the floor.  Counts never carry
# one of these ("regions", "categories", "countries", "players" are not
# units), so the two stay told apart.
_UNIT_WORD = re.compile(
    r"\s*(?:(?:bps|bp|pps|ppts?|percent|per\s+cent|units?|tonnes?|tons?"
    r"|mt|kt|kg|lbs?|mn|bn|tn|k|million|billion|thousand|trillion"
    r"|usd|eur|gbp|jpy|inr)\b"
    r"|us\$|[$\u20ac\u00a3\u00a5\u20b9])", re.I)
# money written in front of the figure says the same thing
_CURRENCY_LEAD = re.compile(r"^[$€£¥₹]")


def _norm_phrase(s) -> str:
    """A name compared the way a slide may have re-typeset it: case, runs of
    whitespace and en/em/minus dashes do not matter."""
    s = str(s or "").lower().replace("–", "-").replace("—", "-") \
        .replace("−", "-").replace(" ", " ")
    return " ".join(s.split())


def _protect_pattern(phrase) -> str:
    """``phrase`` as a regex that tolerates the same re-typesetting."""
    def _word(w):
        return "".join("[-–—−]" if ch == "-" else re.escape(ch)
                       for ch in w)
    return r"\s+".join(_word(w) for w in _norm_phrase(phrase).split())


def is_value(text) -> bool:
    """True when a rendered string READS as a number rather than as a label."""
    s = (str(text) if text is not None else "").strip()
    return bool(s) and s.lower() not in _BLANKS and bool(_VALUE_RE.match(s))


def unit_of(text) -> str:
    """The unit a figure carries, for the padlock that replaces it - so a
    withheld "53.6%" still says it was a share and "108 K" a size."""
    s = str(text or "").strip()
    if _UNIT_PCT_RE.search(s):
        return "%"
    m = _UNIT_K_RE.search(s)
    return m.group(1) if m else ""


def is_open_label(text, open_years=(), open_cagr: bool = True) -> bool:
    """True when a column header / KPI label / panel title names a value that
    must stay readable.

    A label is open when it names CAGR, or when EVERY year it mentions is an
    open year - so "2036", "Share 2026" and "F-CAGR (2026-2036)" are all
    open, while "2029" and "Incremental $ Opportunity 2026-2031" are not.  A
    label naming no year at all is NOT open: locked is the default.
    "CAGR Index" is an index BUILT from CAGR rather than a rate, so indexing
    beats the CAGR exemption.

    Kept as a free function, not just a ``LockPolicy`` method, because a
    report engine has its own idea of where the open years come from and
    only wants the RULE.
    """
    s = str(text) if text is not None else ""
    if not s:
        return False
    if _INDEX_RE.search(s):
        return False
    if _CAGR_RE.search(s):
        return open_cagr
    years = {int(y) for y in _YEAR_RE.findall(s)}
    return bool(years) and years <= {int(y) for y in open_years}


# ═════════════════════════════ the policy ════════════════════════════
# ═══════════════════════ the page rulebook ═══════════════════════════
# A finished deck carries no page model - apply_locks sees shapes, not "the
# Regional Outlook page".  But the engine that wrote the deck puts the same
# HEADING on the same page in every market: only the market's name changes
# around it.  So a page is recognised by the stable part of its heading, and
# each one may state its own treatment.  Anything not listed keeps the
# general behaviour, which is what most pages want.
#
# NOTE ON OPEN YEARS.  The report's policy (fmi_report/locking.default_policy)
# no longer leaves ANY year readable, and no CAGR either - what used to be
# the teaser (last historical year / estimate year / forecast horizon) is now
# withheld like every other value, with the same padlock in the same place.
# Rules below that only existed to widen or narrow that exemption are
# therefore inert; they are kept so a caller passing `open_years=` back in
# gets the old behaviour unchanged.
#
# A rule is a dict; every key is optional:
#     charts   "per_value"  a padlock standing in each withheld figure's place
#              "one_lock"   the figures blurred under a single padlock
#     place    where that single padlock goes: "centre" (default),
#              "whitespace" (the paper a horizontal ranking's bars leave),
#              "top_centre" (lifted above the columns)
#     tables   "per_value" (default) or "one_lock"
# Rules that follow a CHART's own shape rather than a page's heading.  These
# hold everywhere in the deck and in any market, and a page rule does not
# override them: a chart is what it is wherever it is drawn.
CHART_RULES = (
    ("waterfall - a stacked pair with one invisible spacer series",
     "The steps are left blank and the chart takes ONE padlock, centred.  "
     "A blurred figure on every step added a smudge without adding "
     "meaning; the single mark already carries it.  A build-up "
     "reads as a single story, what happens between two totals, so a mark "
     "on every step repeated it.  64 in this deck: both Summary pages, all "
     "twelve Absolute $ Opportunity pages, the country overviews and every "
     "segmental analysis page."),
)


PAGE_RULES = (
    # ══════════════ CROSS-SECTIONAL ASSESSMENT ══════════════
    # The mix chart: 100%-stacked columns, one per sub-segment.  Each column
    # takes ONE medium padlock at its centre and its segment labels go quiet
    # (client, 2026-09-30).
    ("Market Share Analysis:", {
        "charts": "per_column_centre", "lock_fy": 0.34,
        "note": "Cross-sectional mix: a medium padlock at the centre of "
                "each stacked column."}),
    # ══════════════ SECTION 1 - how the report opens ══════════════
    ("Market Signals", {
        "note": "Signals table.  The question column is prose and stays; the "
                "two read-out columns carry figures INSIDE their sentences, "
                "which the general rule finds and withholds, unit and all."}),

    # ══════════════ SECTION 2 - scope ══════════════
    # Taxonomy and coverage: names, not numbers.  Listed so a reader can see
    # they were considered and found to hold nothing to withhold.
    ("Market Taxonomy", {"note": "Region / segment names only - no values."}),
    ("Company Coverage", {"note": "Company names only - no values."}),

    # ══════════════ SECTION 3 - Global Overview ══════════════
    # The five pages every report opens its assessment with, in order.

    # 1 & 3.  Volume and Value summaries: a CAGR chart of three bars and an
    # incremental waterfall of five.  Few enough that a padlock on each reads
    # as a legend, where a blur would say less.  The waterfall stacks its
    # steps on an invisible base, so PowerPoint will not put a label outside
    # the bar - those labels are pinned to the bar's top edge and lifted
    # clear (see _lift_label); the clustered CAGR chart already sits outside
    # and is left alone.
    ("Market Summary, By Volume", {
        "charts": "per_value", "labels": "above",
        "note": "CAGR chart keeps a padlock per bar; the incremental "
                "waterfall follows the deck-wide waterfall rule - steps "
                "blurred under one centred mark.  Nothing is left readable: "
                "the CAGR bars and the 2026 / 2036 figures lock with the "
                "rest."}),
    ("Market Summary, By Value", {
        "charts": "per_value", "labels": "above",
        "note": "As the Volume twin."}),

    # 2.  Pricing: a box plot and a diverging bar chart, both drawn as
    # shapes, so the padlock already stands where each figure did.
    # "Market Pricing", not just "Pricing": section 4 has a "Product Pricing
    # & Regional Price Index Trends" page whose heading would otherwise match
    # this one and take a rule written for a page it is not.
    ("Market Pricing & Regional Price Index", {
        "charts": "per_value",
        "note": "Box plot value and one padlock per regional bar, each "
                "keeping its % ."}),

    # 4.  Volume & Value: sixteen years across two charts - far too many for
    # a mark apiece, so the figures blur under ONE padlock.  All four cards
    # lock: `cards_only` used to name the ten-year opportunity because the
    # open-year rule would otherwise have left the other three readable, and
    # with no open year left there is nothing to override.
    ("Market Volume & Value Analysis", {
        "charts": "one_lock", "place": "centre",
        "note": "Every figure blurred under one centred padlock per chart; "
                "all four cards lock."}),

    # 5.  Country Indices: two rankings whose bars leave paper at their ends,
    # so the one padlock sits in that paper rather than over a bar.  The
    # tables keep a few figures each so the reader sees what the numbers look
    # like without being able to read a column down.
    ("Summary Country Indices", {
        "charts": "one_lock", "place": "whitespace",
        "show_per_col": 3,
        "note": "Rankings blur under one padlock in the white space; heat "
                "map takes a padlock under each country name; tables keep 3 "
                "figures per column, values and padlocks centred."}),

    # The executive-summary series that closes the section.
    ("Market Overview", {
        "note": "Prose only - every figure in the sentences is withheld, "
                "keeping its % or K."}),
    ("Supply Side Trends", {
        "charts": "one_lock", "place": "hole",
        "note": "Four evidence boxes locked in prose; the ring's shares blur "
                "and ONE padlock sits in the hole, sized to it - placed from "
                "the centre read-out's own position, since the ring is not "
                "centred in a frame that carries a legend beside it."}),
    ("Analysis and Recommendation", {
        "note": "Prose only, as Market Overview."}),
    ("Scenario Outlook", {
        "charts": {"LINE": "per_value", "COLUMN": "one_lock"},
        "place": "top_centre",
        "open_cats": ("Likely",),
        "note": "Line chart keeps the base forecast readable and locks the "
                "two flanking cases; the trajectory chart blurs every year, "
                "its ends included, with one padlock in the air above the "
                "bars.  '2026' / '2036' were dropped from open_cats when the "
                "open years closed."}),

    # ══════════════ COUNTRY OVERVIEW ══════════════
    ("Key Segments Overview", {
        "charts": "per_value",
        "note": "One ranking per segment: a padlock in front of every "
                "withheld bar, on every chart - the narrow ones took a "
                "single mark only because they carry fewer bars."}),

    # ═══ COUNTRY SEGMENT SECTIONS - one per segment, per country ═══
    # Three pages each, repeating for every segment the workbook holds, so
    # each rule is written against the part of the heading that carries no
    # segment name.
    ("Segmental Analysis", {
        "open_rows": (1,),
        "note": "Two build-up charts (one centred padlock each, by the "
                "deck-wide waterfall rule) and two tables whose FIRST row "
                "is shown in full, the rest withheld."}),
    ("Market Tables", {
        "note": "Correct as it stands: with no open year left, the Total "
                "row is withheld along with the breakdown beneath it."}),
    ("YoY Analysis", {
        "totals_only": False,
        "note": "Every year column withheld cell by cell, 2025 / 2026 / "
                "2036 included; `totals_only: False` is kept but inert now "
                "that no column is open.  Matched AFTER the two 'YoY "
                "Analysis of Global Market ...' entries above, so this one "
                "picks up the country pages only - both want the same "
                "treatment in any case."}),

    # ══════════════ GLOBAL DRIVING FACTORS ══════════════
    # Three pages, reviewed and correct as the general rules leave them.
    ("Market Dynamics by Region", {
        "note": "Every cell is a SENTENCE carrying figures, so each one is "
                "withheld where it stands and the words around it stay.  "
                "The table is redrawn line by line, which is why its row "
                "heights are rewritten - see _row_heights."}),
    ("Global Driving Factors", {
        "note": "Web chart: each spoke is its own label, 'Factor  <share>', "
                "and the padlock stands where the share was, keeping its %. "
                " The 0-10 axis scale stays readable (it is a ruler, not a "
                "value) and the polygon carries the picture.  The panel "
                "beside it withholds the same figures in prose."}),

    # ══════════════ SUPPLY CHAIN ══════════════
    ("Value Chain & Stakeholder Mapping", {
        "lock": False,
        "note": "Published as it stands - nothing on this page is "
                "withheld, in this report or any future one."}),

    # ══════════════ PRICING ══════════════
    # Reviewed and approved as the general rules leave it, so these entries
    # RECORD the treatment rather than change it: the pages are the same in
    # every market and only the data moves, so what each one does is written
    # down here instead of being rediscovered.  If a general default is ever
    # changed, these notes are what says whether this section still reads the
    # way it was signed off.
    ("Product Pricing & Regional Price Index", {
        "note": "Box plot and diverging bars, both drawn as shapes: a "
                "padlock stands in each withheld figure's place, keeping "
                "its unit.  Axis scales stay readable."}),
    ("Global Pricing Assessment", {
        "note": "Four pages of shape-drawn price panels - a padlock per "
                "withheld figure, the open years readable."}),
    ("Price Positioning & Key Findings", {
        "note": "Prose only: every figure in the sentences is withheld, "
                "keeping its % or unit."}),
    ("Key Pricing Drivers", {
        "note": "Radar of pricing factors: a padlock beside each spoke's "
                "share, the spoke lengths left to carry the picture."}),
    ("Factors Impacting Pricing", {
        "note": "Table of drivers: a padlock in each withheld cell, with "
                "the sentences beneath withheld to match."}),
    ("Regional Price Assessment", {
        "note": "Two price tables - a padlock per withheld cell, open "
                "years narrowed to the Total row."}),
    ("Global Pricing Benchmarking", {
        "note": "Benchmark table: a padlock per withheld cell."}),
    ("Price Forecast & Dynamics", {
        "note": "Forecast tables with the open years readable; the "
                "commentary withholds the same figures."}),
    ("Competitor Pricing Strategies", {
        "note": "Strategy table: a padlock per withheld cell."}),
    ("Industry-Wide Operating Margin", {
        "note": "Margin table: a padlock per withheld cell."}),
    ("Buyers Assessment", {
        "note": "Buyer table: a padlock per withheld cell."}),

    # ═════════ SEGMENT SECTIONS - one per segment in the dataset ═════════
    # These repeat verbatim for every segment the workbook holds, so each
    # rule is written once against the part of the heading that does NOT
    # carry the segment's name.  Twelve segments here; five or twenty in
    # another market changes nothing.

    # The opening page of each segment: a coloured box per sub-segment, three
    # figures in each.  A padlock per row said it three times over, so the
    # figures blur and each box takes ONE mark, centred.
    ("Segment Outlook", {
        "shapes": "per_column",
        "note": "Sub-segment boxes: figures blurred, one centred padlock "
                "per box."}),

    # The two YoY tables.  Every column is a year, and the open ones are the
    # point of the page - but the deck-wide totals_only narrows open columns
    # to a Total row these tables do not have, so it withheld them too.
    ("YoY Analysis of Global Market Volume", {
        "totals_only": False,
        "note": "Every year withheld cell by cell, 2025 / 2026 / 2036 "
                "included."}),
    ("YoY Analysis of Global Market Value", {
        "totals_only": False,
        "note": "As the Volume twin."}),

    # ══════════════ SECTION 4 - By Region ══════════════
    # 1.  Regional Outlook: a grid of one box per region, three figures each.
    # Three padlocks per box repeated the message thirty-nine times, so the
    # figures blur and each region takes ONE mark, centred on the run it
    # withheld.  Three of the grid's cells stay readable - one row shown on
    # the first region, another on the fourth, the last on the final region -
    # so the reader sees what each row holds without being able to read a
    # region across or a row along.  Columns count from the right where the
    # index is negative, so this keeps working when a market has a different
    # number of regions.  The sentences below the grid show those same three
    # figures and withhold the rest.
    ("Regional Outlook", {
        "shapes": "per_column",
        # the first and last regions show everything; the rest show nothing
        "grid_open": ((0, "*"), (-1, "*")),
        "prose_open_regions": ("first", "last"),
        "note": "The first and last region show all three figures and take "
                "no padlock; every other region blurs all three under one "
                "centred mark.  The Sub-Segment Insights sentences for "
                "those two regions stay whole; the rest are withheld."}),

    # ══════════════ SECTION 4 - By Region ══════════════
    # 2.  Volume / Value by region: two stacked columns drawn as RECTANGLES,
    # not as a chart.  A padlock on every slice repeated the message ten
    # times, so the shares blur and one padlock sits in the paper between the
    # two stacks.
    ("Market Analysis, By Region", {
        "shapes": "one_lock", "lock_scale": 1.25,
        "note": "Slice shares blurred; one medium padlock centred between "
                "the two stacks."}),

    # 3.  Share tables: already right - a padlock in each withheld cell,
    # open years on the total row.  Recorded so the formatting is on the
    # record for future reports rather than rediscovered.
    ("Market Share Analysis, By Region", {
        "note": "Table formatting is correct as it stands: padlock per "
                "withheld cell, open years narrowed to the Total row."}),

    # 4.  Regional trends: two stacked charts whose columns fill the plot, so
    # the padlock goes in the clear air above them.
    ("Market Trends, By Region", {
        "charts": "per_column_top",
        "strip_values": ("By segment",),
        "note": "Each column is left blank and takes ONE padlock on its top, "
                "2025 / 2026 / 2036 included now that no year is open.  The "
                "two by-segment read-outs drop their numbers and keep the "
                "country list."}),

    # 5.  Top 10 countries: a treemap.  The padlock belongs under the country
    # name and has to be big enough to read against a large tile.
    ("Top 10 Countries", {
        "card_lock_scale": 2.4,
        "note": "One padlock per tile, under the country name, sized to the "
                "tile rather than to the figure it replaced."}),

    # ══════════════ MARKET DYNAMICS & FORCES ══════════════
    # Five pages of argument.  Every figure on them lives INSIDE a sentence,
    # so there is nothing to place a padlock beside - each one is withheld
    # where it stands and the words around it carry on.  Recorded because
    # every report has this section and it needs no page-specific handling.
    ("Market Drivers", {
        "note": "Prose only.  Each driver quantifies itself mid-sentence; "
                "the figure is withheld in place, keeping its % or K."}),
    ("Market Restraints", {
        "note": "Prose only, as Market Drivers."}),
    ("Market Opportunities", {
        "note": "Prose only, as Market Drivers."}),
    ("Recent Developments", {
        "note": "Prose only - deal sizes and dates inside the milestone "
                "sentences are withheld where they stand."}),
    ("Porter's Five Forces", {
        "note": "Prose only.  The five force panels read their scores in "
                "sentences; the verdict words stay and the scores go."}),

    # ══════════════ MARKET SHARE BY LEADING PLAYERS ══════════════
    # 1.  Industry structure: a funnel of tiers.  The tier NAMES are the
    # point of the page and stay; the share against each tier is withheld
    # where it sits, on the band itself, so the funnel keeps its shape.
    ("Industry Structure Landscape", {
        # The four headline cards on page 2 (new entrants, exits, net change,
        # share moving between tiers) each withhold their figure - counts
        # included - under a MEDIUM padlock centred where the figure stood.
        # The cards are found by shape name, so the table cells on the same
        # pages keep their own, cell-sized padlocks (client, 2026-09-30).
        "lock_counts_in": ("FMI KPI Value",),
        "card_lock_scale_by_name": {"FMI KPI Value": 2.0},
        "note": "Tier ladder, tier profile matrix and structural shift.  "
                "Every figure is withheld where it stands, its unit and sign "
                "kept beside the padlock; a Total of 100% is a definition, "
                "not a figure, and stays.  The four headline cards lock "
                "their figure, counts included, under a medium padlock."}),
    ("Advice for Market Vendors", {
        "note": "Prose only."}),

    # 2-4.  The share rings.  A doughnut's slice shares blur and ONE padlock
    # goes in the hole - but the hole is NOT the frame's centre: every ring
    # here pins its plot area to the left 68% and puts the legend in the
    # rest.  See _ring_geom, which reads that pinned layout back.
    ("Company Value Share", {
        "charts": "one_lock", "place": "hole",
        "note": "Ring shares blurred; ONE padlock in the hole, sized to the "
                "hole and centred on the RING, not on the chart frame."}),
    ("Company Share by Product Type", {
        "charts": "one_lock", "place": "hole",
        "note": "As Company Value Share."}),
    ("Company Share by Region", {
        "charts": "one_lock", "place": "hole",
        "note": "As Company Value Share, once per ring across the row."}),
    ("Who Shaped Growth", {
        "note": "Prose only."}),

    # 5.  The dashboard.  Its figures are laid out in rows of eight, which is
    # the same shape as an axis scale - so the axis guard was exempting the
    # whole page and leaving every revenue readable.  This page says so, and
    # a run is then judged on whether its numbers climb in even steps.
    ("Competitive Dashboard View", {
        "rows_of_values": True,
        "note": "Overall Rev and Target Rev are withheld across all eight "
                "players, and the share ring under each player takes a "
                "padlock in its centre.  Names, regions and the strategy "
                "and portfolio marks stay."}),
)


@dataclass
class LockPolicy:
    """Which of a deck's values stay readable.

    ``open_years`` are the years a client may read - normally the last
    historical year, the estimate year and the forecast horizon.  Leave it
    empty and every year locks.  ``guess_years`` lets ``scan`` fill it in
    from the deck's own table headers when you have not said.
    """

    open_years: set = field(default_factory=set)
    open_cagr: bool = True            # CAGR columns / labels stay readable
    # In a table, narrow EVERY open column - the open years and the CAGR
    # alike - to the TOTAL row: the body locks line by line and only the
    # total keeps its figures.  Use where the breakdown IS the product.
    totals_only: bool = False
    lock_tables: bool = True
    lock_charts: bool = True
    lock_shape_values: bool = True
    lock_prose: bool = True
    # Figures inside a SENTENCE that sits in a table cell - the read-out
    # columns of a signals page ("What we're observing").
    lock_table_prose: bool = True
    # Show a chart's withheld figures BLURRED under ONE padlock rather than a
    # padlock on each, once the chart has this many of them.  Set
    # ``blur_charts=False`` for a padlock on every figure, whatever the count.
    blur_charts: bool = True
    blur_min_labels: int = 4
    blur_ink: str = "#605E5C"
    # Per-page treatments, matched on the slide's heading (see PAGE_RULES).
    page_rules: tuple = PAGE_RULES
    # Obey the page rules' "what stays readable" keys - the teaser a page is
    # allowed to give away: a few scattered figures per column, the first row
    # of a table in full, the first and last region of a grid, a named chart
    # category, a page published whole.  Set False and a rule may still say
    # HOW a page is marked (`charts`, `place`, `labels`, `tables`,
    # `strip_values`) but may no longer keep a value readable, so every figure
    # the deck shows is withheld.  Axis scales are NOT affected: they are
    # never candidates in the first place (see `_drop_axis_labels`).
    page_exemptions: bool = True
    # what a withheld figure inside a sentence becomes when it cannot be
    # drawn as artwork (no padlock PNGs, or prose_padlocks turned off)
    placeholder: str = "xx"
    # Draw the padlock INSIDE sentences as artwork rather than writing the
    # placeholder.  The paragraph is re-laid-out to do it (see
    # _redraw_prose_shape); set False to keep "xx" and leave the text alone.
    prose_padlocks: bool = True
    # A private-use character: legal in XML (a control character is not - the
    # file would carry a literal "_x0001_" where the padlock should go), and
    # nothing in a market report ever contains it.  It is removed again when
    # the sentence is redrawn, so it never reaches the client.
    prose_mark: str = ""
    # a row whose label starts with one of these is the table's total row
    total_labels: tuple = ("total", "grand total", "all ", "overall")
    # never withhold a bare integer below this: it is a COUNT ("6 regions"),
    # not a figure.  Money is in thousands and shares carry a %, so neither
    # is caught.
    min_bare_int: int = 100
    guess_years: bool = True
    # Phrases whose figures are part of a NAME and must survive - above all
    # the market's own: a deck titled "Prescription 5000-ppm Fluoride
    # Toothpaste" would otherwise withhold its own subject.  Matched
    # case-insensitively, anywhere in the text.
    protect: tuple = ()

    def is_open_label(self, text) -> bool:
        """True when a column header / KPI label names a value that stays.

        A label is open when it names CAGR, or when EVERY year it mentions
        is an open year - so "2036", "Share 2026" and "F-CAGR (2026-2036)"
        are open while "2029" and "Variance (BPS)" are not.  A label naming
        no year at all is NOT open: locked is the default.
        """
        return is_open_label(text, self.open_years, self.open_cagr)

    def is_protected(self, text) -> bool:
        """True when a WHOLE text is one of the ``protect`` phrases.

        A segment named by a figure - "5% to <10%", "80 to 100 USD", "500
        ml" - reads as a lone figure, so a taxonomy box or a row label holding
        only that name would be taken for a KPI value and padlocked.  The
        prose rule already skips a protected phrase INSIDE a sentence; this is
        the same rule for a shape or cell that holds nothing but the name.
        """
        t = _norm_phrase(text)
        return bool(t) and any(t == _norm_phrase(p) for p in self.protect or ())

    def is_total_label(self, text) -> bool:
        s = str(text or "").strip().lower()
        return any(s.startswith(t.strip()) for t in self.total_labels)


# ═══════════════════════════ the candidates ══════════════════════════
@dataclass
class Candidate:
    """One value the policy proposes to withhold.

    ``key`` is stable across runs of the same deck, so a UI can hand back a
    set of keys as the decision and get exactly what the reviewer ticked.
    """

    slide: int                 # 1-based
    kind: str                  # table-cell | chart-label | shape-value | prose
    where: str                 # human-readable location
    text: str                  # what it reads now
    unit: str                  # "%", "K", "" - kept by the padlock
    suggested: bool            # what the policy proposes
    shape_path: tuple = ()     # index path to the shape, groups included
    ref: tuple = ()            # within the shape: (row, col) / (series,
    #                            point) / (paragraph,) / (paragraph, mark)

    @property
    def key(self) -> str:
        sp = "-".join(str(i) for i in self.shape_path)
        rf = "-".join(str(i) for i in self.ref)
        return f"{self.slide}|{self.kind}|{sp}|{rf}"


@dataclass
class LockReport:
    locked: int = 0
    skipped_images: int = 0
    candidates: list = field(default_factory=list)
    notes: list = field(default_factory=list)

    def __str__(self):
        out = [f"locked {self.locked} value(s)"]
        if self.skipped_images:
            out.append(f"{self.skipped_images} large picture(s) left alone - "
                       f"check whether any is a chart or table pasted as an "
                       f"image, which cannot be edited")
        return "; ".join(out + self.notes)


# ═════════════════════════════ the padlock ═══════════════════════════
# ONE design.  Geometry lives here rather than only inside a PNG so the
# artwork can be rebuilt anywhere; the figures are in a 24x24 viewBox and are
# the same numbers the report's own padlock is drawn from.
LOCK_ART = {
    "body":    {"x": 6.2, "y": 10.2, "w": 11.6, "h": 11.1, "r": 0.9},
    "shackle": {"cx": 12.0, "cy": 6.93, "r": 3.05, "sw": 2.5, "foot": 11.7},
    "panel":   {"inset": 0.85, "y": 11.95, "h": 2.35, "r": 0.6},
    "keyhole": {"cx": 12.0, "cy": 15.53, "r": 1.01, "sw": 1.0, "sy": 19.0},
    "top": 2.63, "bottom": 21.3,
}
LOCK_STOPS = {
    "body":    [(0, "#F3CD60"), (.33, "#EEC04F"), (.45, "#F1B93A"),
                (.67, "#F5B42D"), (.86, "#F0A920"), (1, "#F3A517")],
    "panel":   [(0, "#FDDC7E"), (1, "#F9D26C")],
    "shackle": [(0, "#969CA7"), (.42, "#BABFCA"), (.72, "#C7CCD6"),
                (1, "#B6BCC8")],
}
LOCK_KEYHOLE = "#6F4305"
LOCK_EDGE = ("#FFFFFF", 0.14)
LOCK_RIM = ("#6B4A05", 0.38)
LOCK_SHADOW = {"dx": 0.55, "dy": 0.75, "blur": 0.62,
               "color": "#2E2205", "opacity": 0.42}

_A = LOCK_ART
# The artwork fills ~0.83 of its canvas HEIGHT.  Sizing the picture box to a
# text run's point size x this factor lands the padlock at the weight the
# figure it replaces had.
LOCK_IMG_K = 1.34
# The visible artwork fills ~0.78 of the 24x24 viewBox's HEIGHT, so a picture
# sized to the font size alone would read SMALLER than the glyph it stands
# in for.  These two keep it at the right weight, and let a surface size the
# padlock by height and let the width fall out of its own proportions.
LOCK_ART_SCALE = 24.0 / (_A["bottom"] - _A["top"])               # ~1.29
LOCK_ART_ASPECT = _A["body"]["w"] / (_A["bottom"] - _A["top"])   # ~0.62
LOCK_IMG_AR = LOCK_ART_ASPECT                                    # older name

# A chart label is ONE box, so a label that must read "lock %" can only
# arrive as one image with the unit already drawn in it.  These are the
# variants; a deck built by this project already ships them.
_ASSET_STEM = {"": "lock_label", "%": "lock_pct", "K": "lock_k"}


class LockArt:
    """Finds - or draws - the padlock artwork.

    Looked for in this order, so a project that already HAS the artwork uses
    its own files - and one that has nothing still gets the padlock:

        1. the directory you pass in / ``$PPTLOCK_ASSETS``
        2. ``lock_assets/`` or ``fmi_report/assets/`` beside this file
        3. the copy EMBEDDED in this file (see ``_EMBEDDED_ART`` at the
           bottom), unpacked into a cache directory on first use - which is
           what makes a single .py enough to hand to someone else
        4. drawn from ``LOCK_ART`` with Pillow, if it happens to be installed
        5. nothing - the caller falls back to "xx" placeholders

    Only the PLAIN padlock can be DRAWN at step 4: the unit variants have
    text baked into them and there is no font we can count on finding.  The
    embedded set has all of them, so step 3 is the one that matters.
    """

    _CACHE = None
    _UNPACKED = None

    def __init__(self, assets_dir=None):
        self.dir = None
        cands = []
        if assets_dir:
            cands.append(Path(assets_dir))
        if os.environ.get("PPTLOCK_ASSETS"):
            cands.append(Path(os.environ["PPTLOCK_ASSETS"]))
        here = Path(__file__).resolve().parent
        cands += [here / "lock_assets", here / "fmi_report" / "assets",
                  here / "assets"]
        for c in cands:
            if (c / "lock.png").exists():
                self.dir = c
                break
        if self.dir is None:
            self.dir = self._unpack() or self._draw()

    # -- the copy that travels with this file -----------------------
    @classmethod
    def _unpack(cls):
        """Write the embedded artwork to a cache dir, once per machine."""
        if cls._UNPACKED is not None:
            return cls._UNPACKED
        art = globals().get("_EMBEDDED_ART") or {}
        if not art:
            cls._UNPACKED = False
            return None
        import base64
        import tempfile
        out = Path(tempfile.gettempdir()) / "pptlock-art-embedded"
        try:
            out.mkdir(parents=True, exist_ok=True)
            for name, b64 in art.items():
                p = out / name
                if not p.exists():
                    p.write_bytes(base64.b64decode("".join(b64.split())))
            cls._UNPACKED = out
        except Exception:
            cls._UNPACKED = False
            return None
        return cls._UNPACKED

    # -- lookup ------------------------------------------------------
    def plain(self):
        p = self.dir / "lock.png" if self.dir else None
        return p if p and p.exists() else None

    def dark(self):
        """The padlock darkened, for a mark that sits ON a chart.

        The gold artwork reads well on paper and disappears over a filled
        bar; this is the same shape carried down to about half its
        lightness, so it stays the house mark and still stands out.
        """
        base = self.plain()
        if base is None:
            return None
        if getattr(self, "_dark", None) is not None:
            return self._dark
        self._dark = None
        try:
            from PIL import Image
            import tempfile
            out = Path(tempfile.gettempdir()) / "pptlock-lock-dark.png"
            if not out.exists():
                im = Image.open(base).convert("RGBA")
                px = im.load()
                for yy in range(im.height):
                    for xx in range(im.width):
                        r, g, b, a = px[xx, yy]
                        px[xx, yy] = (int(r * 0.55), int(g * 0.5),
                                      int(b * 0.42), a)
                im.save(out)
            self._dark = out
        except Exception:
            self._dark = base
        return self._dark

    def labelled(self, unit="", white=False):
        """The artwork with ``unit`` baked in, for a chart data label."""
        if not self.dir:
            return None
        stem = _ASSET_STEM.get(unit or "", "lock_label")
        name = f"{stem}{'_white' if white and unit else ''}.png"
        p = self.dir / name
        if p.exists():
            return p
        p = self.dir / "lock_label.png"
        return p if p.exists() else self.plain()

    @property
    def available(self) -> bool:
        return self.plain() is not None

    # -- fallback: draw it ------------------------------------------
    @classmethod
    def _draw(cls):
        """Render ``lock.png`` from LOCK_ART into a cache dir.  Pillow only;
        no font is needed because only the plain padlock is drawn."""
        if cls._CACHE is not None:
            return cls._CACHE
        try:
            from PIL import Image, ImageChops, ImageDraw, ImageFilter
        except Exception:
            cls._CACHE = None
            return None
        import tempfile
        out_dir = Path(tempfile.gettempdir()) / "pptlock-art"
        out_dir.mkdir(parents=True, exist_ok=True)
        out = out_dir / "lock.png"
        if out.exists():
            cls._CACHE = out_dir
            return out_dir

        a, SS, H0, FILL = LOCK_ART, 4, 480, 0.83
        body, sh, pan, key = a["body"], a["shackle"], a["panel"], a["keyhole"]
        H = H0 * SS
        k = H * FILL / (a["bottom"] - a["top"])
        pad = H * (1 - FILL) / 2
        W = round(body["w"] * k + 2 * pad)

        def X(u):
            return pad + (u - body["x"]) * k

        def Y(u):
            return pad + (u - a["top"]) * k

        def ramp(stops, n):
            pts = [(o, tuple(int(c[i:i + 2], 16) for i in (1, 3, 5)))
                   for o, c in stops]
            res = []
            for i in range(n):
                t = i / max(1, n - 1)
                for (o0, c0), (o1, c1) in zip(pts, pts[1:]):
                    if t <= o1 or (o1, c1) == pts[-1]:
                        f = 0 if o1 == o0 else (min(max(t, o0), o1) - o0) / (o1 - o0)
                        res.append(tuple(round(p + (q - p) * f)
                                         for p, q in zip(c0, c1)))
                        break
                else:
                    res.append(pts[-1][1])
            return res

        def grad(box, stops, horizontal=False):
            x0, y0, x1, y1 = box
            w, h = max(1, round(x1 - x0)), max(1, round(y1 - y0))
            n = w if horizontal else h
            strip = Image.new("RGB", (n, 1) if horizontal else (1, n))
            strip.putdata(ramp(stops, n))
            return strip.resize((w, h), Image.BILINEAR)

        img = Image.new("RGBA", (W, H), (0, 0, 0, 0))

        def blank():
            return Image.new("L", (W, H), 0)

        def stamp(mask, stops=None, box=None, solid=None, alpha=None,
                  horizontal=False):
            if solid is not None:
                layer = Image.new("RGB", (W, H), solid)
            else:
                layer = Image.new("RGB", (W, H))
                layer.paste(grad(box, stops, horizontal),
                            (round(box[0]), round(box[1])))
            m = mask
            if alpha is not None:
                r = grad(box, [(o, "#%02x%02x%02x" % (v, v, v))
                               for o, v in alpha], horizontal=True)
                m = Image.new("L", (W, H))
                m.paste(r.convert("L"), (round(box[0]), round(box[1])))
                m = Image.composite(m, Image.new("L", (W, H)), mask)
            img.paste(layer, (0, 0), m)

        m = blank()
        d = ImageDraw.Draw(m)
        ro = sh["r"] + sh["sw"] / 2
        d.arc([X(sh["cx"] - ro), Y(sh["cy"] - ro), X(sh["cx"] + ro),
               Y(sh["cy"] + ro)], 180, 360, fill=255,
              width=round(sh["sw"] * k))
        for side in (-1, 1):
            cx = sh["cx"] + side * sh["r"]
            d.rectangle([X(cx - sh["sw"] / 2), Y(sh["cy"]),
                         X(cx + sh["sw"] / 2), Y(sh["foot"])], fill=255)
        stamp(m, LOCK_STOPS["shackle"],
              (X(sh["cx"] - ro), Y(sh["cy"] - ro), X(sh["cx"] + ro),
               Y(sh["foot"])), horizontal=True)

        body_box = (X(body["x"]), Y(body["y"]),
                    X(body["x"] + body["w"]), Y(body["y"] + body["h"]))
        m = blank()
        ImageDraw.Draw(m).rounded_rectangle(body_box, radius=body["r"] * k,
                                            fill=255)
        body_mask = m
        stamp(m, LOCK_STOPS["body"], body_box)

        pan_box = (X(body["x"] + pan["inset"]), Y(pan["y"]),
                   X(body["x"] + body["w"] - pan["inset"]),
                   Y(pan["y"] + pan["h"]))
        m = blank()
        ImageDraw.Draw(m).rounded_rectangle(pan_box, radius=pan["r"] * k,
                                            fill=255)
        stamp(m, LOCK_STOPS["panel"], pan_box)

        m = blank()
        d = ImageDraw.Draw(m)
        d.ellipse([X(key["cx"] - key["r"]), Y(key["cy"] - key["r"]),
                   X(key["cx"] + key["r"]), Y(key["cy"] + key["r"])], fill=255)
        d.rectangle([X(key["cx"] - key["sw"] / 2), Y(key["cy"]),
                     X(key["cx"] + key["sw"] / 2), Y(key["sy"])], fill=255)
        stamp(m, solid=LOCK_KEYHOLE)

        stamp(body_mask, solid=LOCK_EDGE[0], box=body_box,
              alpha=[(0, 0), (.55, 0), (1, round(LOCK_EDGE[1] * 255))])

        rim = Image.new("L", (W, H))
        r = grad(body_box, [(0, "#000000"), (.72, "#000000"),
                            (1, "#%02x%02x%02x"
                                % ((round(LOCK_RIM[1] * 255),) * 3))])
        rim.paste(r.convert("L"), (round(body_box[0]), round(body_box[1])))
        img.paste(Image.new("RGB", (W, H), LOCK_RIM[0]), (0, 0),
                  Image.composite(rim, Image.new("L", (W, H)), body_mask))

        blur = max(1.0, LOCK_SHADOW["blur"] * k / SS) * SS
        sil = img.getchannel("A").filter(ImageFilter.GaussianBlur(blur / 2))
        shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        tint = tuple(int(LOCK_SHADOW["color"][i:i + 2], 16) for i in (1, 3, 5))
        shadow.paste(Image.new("RGB", (W, H), tint), (0, 0),
                     sil.point(lambda v: round(v * LOCK_SHADOW["opacity"])))
        shadow = ImageChops.offset(shadow, round(LOCK_SHADOW["dx"] * k),
                                   round(LOCK_SHADOW["dy"] * k))
        shadow.alpha_composite(img)
        shadow.resize((max(1, round(W / SS)), H0),
                      Image.LANCZOS).save(out)
        cls._CACHE = out_dir
        return out_dir


# ═══════════════════════════ shape walking ═══════════════════════════
def _iter_shapes(shapes, prefix=()):
    """Every shape on a slide, groups flattened, with its index path."""
    for i, shp in enumerate(shapes):
        path = prefix + (i,)
        if shp.shape_type == 6 or getattr(shp, "shapes", None) is not None:
            try:
                yield from _iter_shapes(shp.shapes, path)
                continue
            except (AttributeError, TypeError):
                pass
        yield path, shp


def _abs_rect(shp):
    """(left, top, width, height) in px, or None where the shape has none."""
    try:
        if None in (shp.left, shp.top, shp.width, shp.height):
            return None
        return (shp.left / EMU_PER_PX, shp.top / EMU_PER_PX,
                shp.width / EMU_PER_PX, shp.height / EMU_PER_PX)
    except (AttributeError, TypeError):
        return None


def _para_text(para) -> str:
    return "".join(r.text for r in para.runs)


def _rewrite_paragraph(para, transform) -> bool:
    """Apply ``transform(whole_paragraph_text) -> new_text`` in place, keeping
    every run's own formatting.

    PowerPoint splits a number across runs constantly ("1", "43", ",", "010"),
    so a regex run by run misses most figures.  This joins the paragraph,
    transforms it whole, then writes the result BACK across the same runs by
    position - the runs keep their fonts, colours and bold, and only their
    text changes.  The last run absorbs any change in length.
    """
    runs = list(para.runs)
    if not runs:
        return False
    old = "".join(r.text for r in runs)
    new = transform(old)
    if new == old:
        return False
    # walk the ORIGINAL run boundaries against the new string, keeping each
    # run's share of the text; where the transform shortened a stretch the
    # boundaries slide, which is exactly what we want (formatting follows the
    # surrounding words, not the removed digits)
    if len(runs) == 1:
        runs[0].text = new
        return True
    scale = len(new) / len(old) if old else 1.0
    pos, cut = 0, 0
    for r in runs[:-1]:
        pos += len(r.text)
        nxt = min(len(new), max(cut, round(pos * scale)))
        r.text = new[cut:nxt]
        cut = nxt
    runs[-1].text = new[cut:]
    return True


# ═════════════════════════════ scanning ══════════════════════════════
def _guess_open_years(prs, policy) -> set:
    """The three open years, read off the deck's own table headers.

    A report's size tables run the whole horizon, so the LAST year in the
    grid is the forecast horizon and the first is the history start.  The
    estimate year is the one the deck's headings call out - failing that,
    the horizon, the year before the mid-point, and the last historical
    year cannot be told apart, so only the horizon is assumed open.
    """
    years = set()
    for slide in prs.slides:
        for _p, shp in _iter_shapes(slide.shapes):
            if not getattr(shp, "has_table", False):
                continue
            hdr = [c.text.strip() for c in shp.table.rows[0].cells]
            found = [int(y) for h in hdr for y in _YEAR_RE.findall(h)]
            if len(found) >= 4:
                years.update(found)
    if not years:
        return set()
    lo, hi = min(years), max(years)
    # history end / estimate year: the report's own convention is that the
    # estimate year follows the last historical one, and the two sit a
    # decade before the horizon
    est = hi - 10 if (hi - 10) > lo else lo + 1
    return {est - 1, est, hi}


def _scattered(n_rows, n_show, seed):
    """``n_show`` row indices out of ``n_rows``, chosen the same way every
    run so two exports of one deck withhold the same cells."""
    import random
    rnd = random.Random(f"{seed}|{n_rows}|{n_show}")
    return set(rnd.sample(range(1, max(2, n_rows)),
                          min(n_show, max(1, n_rows - 1))))


def _is_small_count(txt) -> bool:
    """True for a bare integer below the policy's floor - a count, not money."""
    t = str(txt).strip()
    if not re.fullmatch(r"\d{1,3}", t):
        return False
    try:
        return int(t) < LockPolicy.min_bare_int
    except ValueError:
        return False


def _value_grid(slide, policy):
    """The grid of value cells a page like Regional Outlook is built from.

    Returns rows of (x, path, text), each row sorted left to right, the rows
    themselves top to bottom.  The page draws one small shape per cell, so
    the grid is recovered by clustering those shapes on their y and x.
    """
    cells = []
    for path, shp in _iter_shapes(slide.shapes):
        if not getattr(shp, "has_text_frame", False):
            continue
        vi = _shape_value_para(shp)
        if vi is None:
            continue
        txt = _para_text(shp.text_frame.paragraphs[vi]).strip()
        if (not is_value(txt) or _is_small_count(txt)
                or re.fullmatch(r"(19|20)\d{2}", txt)):
            continue
        r = _abs_rect(shp)
        if r is not None:
            cells.append((r[1], r[0], path, txt))
    if not cells:
        return []
    rows, cur, last_y = [], [], None
    for y, x, path, txt in sorted(cells):
        if last_y is not None and abs(y - last_y) > 10:
            rows.append(sorted(cur)); cur = []
        cur.append((x, path, txt)); last_y = y
    if cur:
        rows.append(sorted(cur))
    return rows


def _grid_open(slide, policy, spec):
    """(paths, texts) of the cells a grid page leaves readable.

    ``spec`` names them as (column, row) pairs; a negative column counts from
    the right, so "the last region" keeps working when a market has a
    different number of them.
    """
    rows = _value_grid(slide, policy)
    paths, texts = set(), set()
    if not rows:
        return paths, texts
    for col, row in spec:
        # row "*" opens the WHOLE column - every figure that region carries
        want = range(len(rows)) if row == "*" else (
            [row] if row < len(rows) else [])
        for ri in want:
            line = rows[ri]
            if not line:
                continue
            c = col if col >= 0 else len(line) + col
            if 0 <= c < len(line):
                _x, path, txt = line[c]
                paths.add(path)
                texts.add(txt)
    return paths, texts


# A read-out panel whose lines are "<figure>  <name>".  Where a page asks
# for it, the figure is taken out and the name kept - which reads better
# than a padlock in a legend, and gives nothing away.
_STRIP_LINE = re.compile(r"^(\s*[■▪●•·-]*\s*)"
                         r"[$€£¥₹]?\d[\d,.]*\s*"
                         r"(?:[KMB]n?|%)?\s+")


def _is_strip_panel(shp, rule) -> bool:
    pats = (rule or {}).get("strip_values")
    if not pats or not getattr(shp, "has_text_frame", False):
        return False
    head = " ".join(shp.text_frame.text.split())[:60].lower()
    return any(pat.lower() in head for pat in pats)


_LONE_FIGURE = re.compile(r"^\s*[$€£¥₹]?"
                          r"\d[\d,.]*\s*(?:[KMB]n?|%)?\s*$")


def _strip_panel(shp):
    """Drop the leading figure from every line of a read-out panel.

    Run by run, not by rewriting the paragraph: these lines are built as
    "<swatch><figure><name>", each its own run, and redistributing the
    shortened text across them by position handed the name's first letter to
    the swatch's run - so every country came out with a coloured initial.
    Clearing the figure's own run leaves every other run exactly as it was.
    """
    done = 0
    for para in shp.text_frame.paragraphs:
        runs = list(para.runs)
        for i, r in enumerate(runs):
            if not _LONE_FIGURE.match(r.text or ""):
                continue
            r.text = ""
            # the gap that separated figure from name goes with it
            if i + 1 < len(runs):
                nxt = runs[i + 1]
                nxt.text = (nxt.text or "").lstrip()
            done += 1
            break
    return done


def _grid_headers(slide, rows):
    """The column headings above a value grid, left to right.

    On the Regional Outlook page these are the region names in their coloured
    boxes.  They are found by their COLUMN: a heading sits above the first
    value row and shares that column's x.
    """
    if not rows:
        return []
    top = min(_shape_y(slide, path) for _x, path, _t in rows[0])
    out = []
    for x, _path, _t in rows[0]:
        best, best_y = None, None
        for _p, shp in _iter_shapes(slide.shapes):
            if not getattr(shp, "has_text_frame", False):
                continue
            t = shp.text_frame.text.strip()
            if not t or len(t) > 30:
                continue
            r = _abs_rect(shp)
            if r is None or r[1] >= top or abs(r[0] - x) > 24:
                continue
            if best_y is None or r[1] > best_y:
                best, best_y = " ".join(t.split()), r[1]
        out.append(best)
    return out


def _shape_y(slide, path):
    r = _abs_rect(_shape_at(slide, path))
    return r[1] if r else 0.0


def _scan_table(slide_no, shp, path, policy, out, rule=None):
    tbl = shp.table
    rows = list(tbl.rows)
    if len(rows) < 2:
        return
    # A page may ask that a FEW of each column's figures stay readable -
    # enough to show what the numbers look like, not enough to read the
    # column down.  Which ones is seeded off the header, so the gaps do not
    # line up across columns and two runs agree.
    # Rows a page keeps readable whatever the column rule says - "show the
    # client the first line in full, withhold the rest".
    open_rows = set((rule or {}).get("open_rows") or ())
    show_n = (rule or {}).get("show_per_col")
    # ``totals_only`` narrows every open column to the TOTAL row.  A YoY
    # table has no total row, so that switch withholds the open years too -
    # a page may say it wants them in full.
    tot_only = (rule or {}).get("totals_only", policy.totals_only)
    open_cells = {}
    if show_n:
        for ci in range(1, len(rows[0].cells)):
            head = rows[0].cells[ci].text.strip()
            open_cells[ci] = _scattered(len(rows), int(show_n),
                                        f"{path}|{head}")
    header = [c.text.strip() for c in rows[0].cells]
    open_col = [policy.is_open_label(h) for h in header]
    total_rows = {r for r, row in enumerate(rows)
                  if r and policy.is_total_label(row.cells[0].text)}
    for r, row in enumerate(rows):
        if r == 0:
            continue
        cells = list(row.cells)
        for c, cell in enumerate(cells):
            if c == 0:
                continue                       # the row label is not a value
            text = cell.text.strip()
            if policy.is_protected(text):
                continue                       # a segment name, not a value
            if not is_value(text):
                # Not a lone figure - but a cell can hold a SENTENCE with
                # figures in it ("Knee leads at 46.4% of 2036 value"), and
                # those are values too.  They are withheld the way a
                # sentence anywhere else is: the padlock stands where the
                # figure did and the words around it stay.
                if policy.lock_prose and policy.lock_table_prose:
                    for mi, (_a, _b, whole, unit) in enumerate(
                            _prose_marks(text, policy)):
                        out.append(Candidate(
                            slide=slide_no, kind="table-prose",
                            where=f"{header[c] if c < len(header) else c}: "
                                  f"{text[:44]}",
                            text=whole, unit=unit, suggested=True,
                            shape_path=path, ref=(r, c, mi)))
                continue
            if r in open_rows:
                continue                       # this page shows this row
            if r in open_cells.get(c, ()):
                continue                       # one of this column's few
            if c < len(open_col) and open_col[c]:
                # an open column stays open - unless totals_only narrows it
                # to the total row, which is the whole point of that switch
                if not tot_only or r in total_rows:
                    continue
            elif r in total_rows and c < len(open_col) and open_col[c]:
                continue
            out.append(Candidate(
                slide=slide_no, kind="table-cell",
                where=f"table {'/'.join(map(str, path))} "
                      f"[{rows[r].cells[0].text.strip() or r}, "
                      f"{header[c] if c < len(header) else c}]",
                text=text, unit=unit_of(text), suggested=True,
                shape_path=path, ref=(r, c)))


# ── a value written INTO a category label ────────────────────────────
# A category can name a row and measure it at once - the structure chart writes
# each region as "North America  ·  7.0K", the size BEING the row label - so the
# figure is a value like any other and is withheld wherever it is written.  It
# cannot be blurred in place: a category label is one text item and PowerPoint
# gives no way to draw a picture over part of one, so the figure is replaced by
# the padlock the dashboard shows in the same spot.
LOCK_GLYPH = "🔒"
_CAT_VALUE_RE = re.compile(r"-?\d[\d,]*(?:\.\d+)?\s*[KMB]?%?")
# A withheld category label keeps only its NAME - "U.S.", not "U.S.  ·  🔒".
# The bars beside it are already blurred and padlocked, so a padlock on every
# axis label only repeated that.  False restores the name-plus-padlock label.
CATEGORY_NAME_ONLY = True
_CAT_SEP_RE = re.compile(r"(?:\s*[·|:\-–—])+\s*$")


def mask_category_value(cat, glyph=LOCK_GLYPH):
    """`"North America · 7.0K"` -> `"North America"`; None if the label
    carries no size.  (With CATEGORY_NAME_ONLY off: `"North America · 🔒"`.)

    The label has to NAME something - two letters or more, so "U.S." counts as
    readily as "North America" - and then set the figure apart from the name
    with a separator.  A tick that is ONLY a figure (2026, 40%) is a SCALE, the
    ruler the chart is read against, and stays; so does a period written as
    "Q1 2026" or "FY26", and so does a segment named "Type 2".
    """
    text = str(cat)
    m = _CAT_VALUE_RE.search(text)
    if m is None:
        return None
    head = text[:m.start()]
    if len(re.findall(r"[A-Za-z]", head)) < 2:
        return None
    if not re.search(r"(?:[·|:\-–—]|\s{2,})\s*$", head):
        return None
    if CATEGORY_NAME_ONLY:
        return (_CAT_SEP_RE.sub("", head) + text[m.end():]).strip()
    return (head + glyph + text[m.end():]).rstrip()


def _mask_chart_categories(chart, policy) -> int:
    """Withhold the size each category label carries.  Returns how many went.

    Every series keeps its own copy of the categories, so all of them are
    rewritten - leaving one behind would let PowerPoint redraw the axis from
    the copy that still held the figures.
    """
    n = 0
    for cat in chart._chartSpace.iter(qn("c:cat")):
        for v in cat.iter(qn("c:v")):
            if not v.text or policy.is_open_label(v.text):
                continue
            masked = mask_category_value(v.text, policy.placeholder
                                         if not policy.prose_padlocks
                                         else LOCK_GLYPH)
            if masked is not None:
                v.text = masked
                n += 1
    return n


def _chart_categories(chart):
    try:
        return [str(c) for c in chart.plots[0].categories]
    except Exception:
        return []


def _chart_shows_values(chart) -> bool:
    """Only labels the deck ALREADY draws are touched - giving a point a
    label it never had would make PowerPoint start drawing one."""
    try:
        for plot in chart.plots:
            if plot.has_data_labels and plot.data_labels.show_value:
                return True
        for s in chart.series:
            dl = s._element.find(qn("c:dLbls"))
            if dl is not None:
                return True
    except Exception:
        pass
    return False


def _chart_label_unit(chart) -> str:
    """The unit the chart's own labels print, so the padlock keeps it.

    Read off what the chart SAYS rather than guessed from its type: a
    doughnut labelled with values is not showing percentages, and a bar
    chart formatted ``0"%"`` is.
    """
    try:
        for plot in chart.plots:
            if not plot.has_data_labels:
                continue
            dls = plot.data_labels
            if getattr(dls, "show_percentage", False):
                return "%"
            fmt = str(dls.number_format or "")
            if "%" in fmt:
                return "%"
            if re.search(r'"\s*[KMB]n?\s*"', fmt):
                return re.sub(r'\W', "", re.search(r'"\s*([KMB]n?)\s*"',
                                                   fmt).group(1))
    except Exception:
        pass
    return ""


def _series_is_invisible(series) -> bool:
    """True when a series is drawn with no fill - a spacer, not a bar."""
    try:
        from pptx.enum.dml import MSO_FILL
        return series.format.fill.type == MSO_FILL.BACKGROUND
    except Exception:
        return False


def _scan_chart(slide_no, shp, path, policy, out, rule=None):
    chart = shp.chart
    if not _chart_shows_values(chart):
        return
    cats = _chart_categories(chart)
    unit = _chart_label_unit(chart)
    for si, series in enumerate(chart.series):
        if _series_is_invisible(series):
            # A waterfall stands its steps on an INVISIBLE base series: the
            # bars are there to lift the visible ones, and they carry labels
            # nobody sees.  Withholding those put a second padlock under
            # every step and left blur smudges along the floor of the plot.
            continue
        vals = list(series.values)
        for pi, v in enumerate(vals):
            if v is None:
                continue
            cat = cats[pi] if pi < len(cats) else ""
            keep = (rule or {}).get("open_cats")
            if keep is not None:
                # this page names the categories that stay readable; the year
                # rule does not get a say ("Likely" carries no year, so the
                # year rule would withhold it)
                if any(_norm_dash(k) == _norm_dash(cat) for k in keep):
                    continue
            elif policy.is_open_label(cat):
                continue
            out.append(Candidate(
                slide=slide_no, kind="chart-label",
                where=f"chart {'/'.join(map(str, path))} "
                      f"[{series.name or si}, {cat or pi}]",
                text=f"{v:,.1f}" if isinstance(v, float) else str(v),
                unit=unit, suggested=True,
                shape_path=path, ref=(si, pi)))


# A paragraph is a LONE figure when almost nothing survives stripping its
# digits, punctuation and unit - "143,010", "4.0%", "24,041 US$ Th" do;
# "2026 MARKET SIZE (US$ TH)" does not, even though it opens with digits.
_FIG_STRIP = re.compile(r"[\d,.\s%×x$€£¥₹+\-–—()]")


def _is_lone_figure(s: str) -> bool:
    s = str(s).strip()
    return is_value(s) and len(_FIG_STRIP.sub("", s)) <= 8


def _shape_value_para(shp):
    """Index of the ONE paragraph that is a lone figure, or None.

    A KPI card reads "2026 MARKET SIZE\\n24,041": a label and a value.  When
    exactly one paragraph is a value on its own we know its box, so it can
    take the artwork rather than a placeholder.  The LABEL often opens with
    a year, which reads as a value on its own - hence ``_is_lone_figure``
    rather than ``is_value`` here.
    """
    try:
        paras = list(shp.text_frame.paragraphs)
    except Exception:
        return None
    if len(paras) > 3:
        return None
    hits = [i for i, p in enumerate(paras) if _is_lone_figure(_para_text(p))]
    return hits[0] if len(hits) == 1 else None


def prose_figures(text, protect=(), min_bare_int: int = 100):
    """``[(start, end, whole, number, unit)]`` for every figure in a sentence
    that a report may withhold.

    THE finder - a report engine that withholds figures in its own prose
    should call this rather than keep a second regex, because the two drift:
    the thousands-separator bug that ate the comma after "2036," lived in
    both copies and had to be fixed twice.

    Skipped:

    * years, which are structure rather than a figure
    * anything inside single quotes - a category is quoted BY NAME and those
      names carry figures of their own ('1.1% Sodium Fluoride' must not lose
      its 1.1)
    * anything inside a ``protect`` phrase, above all the market's own name
    * a figure inside a hyphenated token ("5000-ppm", "24-hour"), which is a
      name rather than a measurement
    * a bare integer under ``min_bare_int`` - that is a COUNT ("6 regions"),
      and money in these reports is in thousands while shares carry a %, so
      neither is caught
    """
    s = str(text)
    skip = [m.span() for m in _QUOTED.finditer(s)]
    for phrase in protect or ():
        if phrase:
            skip += [m.span() for m in
                     re.finditer(_protect_pattern(phrase), s, re.I)]
    out = []
    for m in _PROSE_FIG.finditer(s):
        num, tail = m.group(1), m.group(2) or ""
        if any(a <= m.start() < b for a, b in skip):
            continue
        parts = _NUM_PART.findall(num)
        # a bare year, or a span opening on one ("2021-2025", "2026-28"), is
        # structure rather than a figure
        if not tail and parts and _YEAR_RE.fullmatch(parts[0]):
            continue
        # a unit after the figure, or a currency in front of it, makes it a
        # measurement rather than a count, whatever its size
        measured = (bool(tail) or bool(_CURRENCY_LEAD.match(num))
                    or bool(_UNIT_WORD.match(s[m.end():])))
        if not measured and all(n.isdigit() and int(n) < min_bare_int
                                for n in parts):
            continue
        if _HYPHENATED.match(s[m.end():m.end() + 3]):
            continue
        out.append((m.start(), m.end(), m.group(0), num, tail.strip()))
    return out


def _prose_marks(s, policy):
    """``prose_figures`` in the shape this module applies: with the unit the
    padlock keeps."""
    return [(a, b, whole, unit) for a, b, whole, _num, unit
            in prose_figures(s, policy.protect, policy.min_bare_int)]


def scan(deck, policy: LockPolicy = None) -> list:
    """Every value ``policy`` proposes to withhold, in slide order.

    ``deck`` is a path or an open ``Presentation``.  Nothing is modified.
    """
    policy = policy or LockPolicy()
    prs = deck if isinstance(deck, Presentation().__class__) else \
        Presentation(str(deck))
    if policy.guess_years and not policy.open_years:
        policy.open_years = _guess_open_years(prs, policy)
    out = []
    for n, slide in enumerate(prs.slides, start=1):
        rule = _rule_for(_slide_heading(slide), policy)
        if rule.get("lock") is False:
            continue                  # this page is published as it stands
        grid_paths, grid_texts = set(), set()
        open_leads = ()
        if rule.get("grid_open"):
            grid_paths, grid_texts = _grid_open(slide, policy,
                                                rule["grid_open"])
        if rule.get("prose_open_regions"):
            # The read-out bullets name their region and then quote its
            # figures ("Germany: 20% of 2026 value ...").  Where the grid
            # shows a region in full, its sentence may show it too.
            names = _grid_headers(slide, _value_grid(slide, policy))
            pick = []
            for which in rule["prose_open_regions"]:
                if names:
                    pick.append(names[0] if which == "first" else names[-1])
            open_leads = tuple(n.lower() for n in pick if n)
        for path, shp in _iter_shapes(slide.shapes):
            if policy.lock_tables and getattr(shp, "has_table", False):
                _scan_table(n, shp, path, policy, out, rule)
                continue
            if policy.lock_charts and getattr(shp, "has_chart", False):
                _scan_chart(n, shp, path, policy, out, rule)
                continue
            if not getattr(shp, "has_text_frame", False):
                continue
            vi = _shape_value_para(shp) if policy.lock_shape_values else None
            if vi is not None:
                txt = _para_text(shp.text_frame.paragraphs[vi]).strip()
                if policy.is_protected(txt):
                    # a segment NAMED by a figure ("5% to <10%", "Below 80
                    # USD") in a box of its own - a name, not a value
                    continue
                if _is_small_count(txt) and getattr(shp, "name", "") not in (
                        rule.get("lock_counts_in") or ()):
                    # A bare integer under min_bare_int is a COUNT, not a
                    # figure - the numbered badge on a region box, "6
                    # regions", a rank.  The prose rule has always known
                    # this; a value standing alone in a shape did not, so the
                    # badges on a grid page were being withheld.  A page may
                    # name the shapes whose counts ARE its headline figures
                    # ("lock_counts_in") - those lock like any other value.
                    continue
                if path in grid_paths:
                    continue            # this page keeps this cell readable
                if re.fullmatch(r"(19|20)\d{2}", txt):
                    # A year standing alone in its own shape is the category
                    # label under a column, not a figure - and it has no
                    # label of its own for the open-year rule to read, so it
                    # would otherwise be withheld and the axis lose its years.
                    continue
                label = " ".join(_para_text(p).strip()
                                 for i, p in
                                 enumerate(shp.text_frame.paragraphs)
                                 if i != vi)
                only = rule.get("cards_only")
                if only is not None:
                    # this page names the cards that lock; the year rule does
                    # not get a say, either way
                    lab = " ".join(label.lower().split())
                    if any(_norm_dash(o) in _norm_dash(lab) for o in only):
                        out.append(Candidate(
                            slide=n, kind="shape-value",
                            where=label[:40] or "shape", text=txt,
                            unit=unit_of(txt), suggested=True,
                            shape_path=path, ref=(vi,)))
                    continue
                if not policy.is_open_label(label):
                    out.append(Candidate(
                        slide=n, kind="shape-value",
                        where=f"{label[:40] or 'shape'}",
                        text=txt, unit=unit_of(txt), suggested=True,
                        shape_path=path, ref=(vi,)))
                continue
            if not policy.lock_prose:
                continue
            if _is_strip_panel(shp, rule):
                continue        # rewritten wholesale in apply, not withheld
            for pi, para in enumerate(shp.text_frame.paragraphs):
                s = _para_text(para)
                for mi, (_a, _b, whole, unit) in enumerate(
                        _prose_marks(s, policy)):
                    lead = s.strip().lstrip("●▪•·-  ").lower()
                    if open_leads and lead.startswith(open_leads):
                        continue        # this region's sentence stays whole
                    if whole.strip() in grid_texts:
                        # the page shows this figure in its own grid, so the
                        # sentence that reads it back may show it too
                        continue
                    out.append(Candidate(
                        slide=n, kind="prose",
                        where=s.strip()[:60], text=whole, unit=unit,
                        suggested=True, shape_path=path, ref=(pi, mi)))
    loose = set()
    for n, slide in enumerate(prs.slides, start=1):
        if _rule_for(_slide_heading(slide), policy).get("rows_of_values"):
            loose.add(n)
    return _drop_axis_labels(prs, out, loose)


# An axis SCALE is not a value: it is the ruler a chart is read against, and
# the decks this locks leave it showing.  A shape-drawn chart writes each tick
# as its own little text box holding one figure, which is exactly the shape a
# KPI card has - so they are told apart by COMPANY.  Three or more identical
# little boxes lined up in a column (a value axis) or a row (a category axis)
# are a scale; a card stands alone.
_AXIS_MIN_TICKS = 3
_AXIS_ALIGN_PX = 6.0            # how far apart two ticks may sit and still
                                # count as sharing an edge
_AXIS_MAX_PX = 46.0             # a tick box is small; a KPI card is not


def _reads_as_number(t):
    """The numeric value of a tick label, or None."""
    m = re.search(r"-?\d[\d,]*(?:\.\d+)?", str(t or ""))
    if not m:
        return None
    try:
        v = float(m.group(0).replace(",", ""))
    except ValueError:
        return None
    u = str(t)[m.end():m.end() + 1].upper()
    return v * {"K": 1e3, "M": 1e6, "B": 1e9}.get(u, 1.0)


def _split_run(members, rects, axis):
    """Break one aligned run into the separate ladders it is made of."""
    i = 1 if axis == "x" else 0          # measure along the run, not across it
    pos = sorted((rects[k][i], k) for k in members if k in rects)
    if len(pos) < 3:
        return [set(members)]
    gaps = [b[0] - a[0] for a, b in zip(pos, pos[1:])]
    typical = sorted(gaps)[len(gaps) // 2]
    runs, cur = [], {pos[0][1]}
    for (_, key), gap in zip(pos[1:], gaps):
        if typical > 0 and gap > typical * 2.5:
            runs.append(cur); cur = set()
        cur.add(key)
    runs.append(cur)
    return runs


_AXIS_SIZE_TOL = 1.5            # a scale's ticks are all drawn the same size


def _same_size(run, rects) -> bool:
    """True when every box in a run is drawn at the same height.

    One ruler is set in one type size, so a run whose boxes differ in height
    is not a scale but an accident of alignment - the pricing sparkline's
    "+8%" (a 12px scale label) happening to share a top edge with a row of
    read-outs (38px), which together would otherwise read as a two-tick
    scale and leave the read-outs published.
    """
    hs = [rects[k][3] for k in run if k in rects]
    return bool(hs) and max(hs) <= min(hs) * _AXIS_SIZE_TOL


def _is_axis_run(members, texts) -> bool:
    """True when a row or column of lone figures is a SCALE, not data.

    A scale counts in even steps - 0, 2, 4, 6 - while a row of figures on a
    dashboard does not.  Without this test any row of eight company revenues
    looked exactly like an axis and was left readable.
    """
    # A scale may carry a word among its numbers - a diverging chart labels
    # its centre "Global" - so entries that are not numbers are passed over
    # rather than disqualifying the run.
    vals = [v for v in (_reads_as_number(texts.get(k)) for k in members)
            if v is not None]
    # A scale is a set of TICKS, and one ruler drawn beside four stacked
    # panels repeats those ticks four times: the pricing-pattern sparklines
    # label their top and bottom "+8%" / "-108%", which is a scale of two
    # read once per panel.  Judge the distinct ticks, so the repetition does
    # not read as a step of zero - and so a row of figures that are all the
    # SAME value ("0.00  0.00  0.00", a read-out) collapses to one tick and
    # is correctly not a scale.
    vals = sorted(set(vals))
    if len(vals) < 2:
        return False
    order = vals
    steps = [b - a for a, b in zip(order, order[1:])]
    if not steps or min(steps) <= 0:
        return False
    return (max(steps) - min(steps)) <= max(steps) * 0.02


def _drop_axis_labels(prs, cands, loose=()):
    """Take the axis scales back out of ``cands``.

    A run of aligned little figure boxes is spared only when it actually
    READS as a scale - its numbers climb in even steps (`_is_axis_run`).
    That test used to run on the pages the rulebook tagged `rows_of_values`
    and nowhere else, so on every other page ANY three aligned figures were
    assumed to be an axis and left readable: the Global Pricing Assessment
    read-outs sat in rows of three and were spared wholesale.  The tagged
    page already worked this way; it is now simply what every page does, so
    a real ruler is still never withheld.  ``loose`` is kept for callers and
    no longer changes the outcome.
    """
    loose = set(loose)
    rects = {}
    for n, slide in enumerate(prs.slides, start=1):
        for path, shp in _iter_shapes(slide.shapes):
            r = _abs_rect(shp)
            if r is not None:
                rects[(n, path)] = r
    buckets, texts = {}, {}
    for c in cands:
        if c.kind != "shape-value":
            continue
        key = (c.slide, c.shape_path)
        texts[key] = c.text
        r = rects.get(key)
        if r is None or r[3] > _AXIS_MAX_PX:
            continue
        # a column of ticks shares a left edge; a row of them shares a top one
        buckets.setdefault((c.slide, "x", round(r[0] / _AXIS_ALIGN_PX)),
                           set()).add(key)
        buckets.setdefault((c.slide, "y", round(r[1] / _AXIS_ALIGN_PX)),
                           set()).add(key)
    drop = set()
    for (slide_no, axis, _), members in buckets.items():
        # Figures laid out in a row or a column are the same SHAPE as an axis,
        # so shape alone cannot tell a ruler from data: the run is judged on
        # whether its numbers actually climb in even steps.
        for run in _split_run(members, rects, axis):
            if (len(run) >= _AXIS_MIN_TICKS and _same_size(run, rects)
                    and _is_axis_run(run, texts)):
                drop |= run
    return [c for c in cands
            if not (c.kind == "shape-value"
                    and (c.slide, c.shape_path) in drop)]


# ═════════════════════════════ applying ══════════════════════════════
def _col_x(tbl, c):
    return sum(tbl.columns[i].width for i in range(c)) / EMU_PER_PX


# PowerPoint breaks a table cell earlier than the column's bare width
# suggests - measured off its own rendering, at about 80% of the space
# between the cell margins.
_TABLE_FILL = 0.80
# A cell's own margins are read from the file, so no further allowance is
# needed on top of the line box.  Measured against PowerPoint's rendering of
# both a dense table (13 rows, wrapping notes) and a sparse one (5 rows of
# short figures): the pitch it draws is the line box plus those margins.
_CELL_CHROME_PX = 0.0
_ROW_SAFETY = 1.15


def _row_heights(tbl):
    """Each row's height as PowerPoint will actually DRAW it, in px.

    ``<a:tr h="...">`` is a minimum, not a measurement: PowerPoint grows a
    row whose text needs more than one line, and a header like "$ Opp (USD
    Mn)" in a narrow column routinely takes two.  Trusting the declared
    heights put every padlock a row too high on such a table - one landing
    in the header, the rest covering the very figures the page meant to
    leave readable - so the wrapped line count is measured here and the row
    grown to match.
    """
    out = []
    for row in tbl.rows:
        h = row.height / EMU_PER_PX
        need = 0.0
        for c, cell in enumerate(row.cells):
            txt = cell.text.strip()
            if not txt:
                continue
            pt = _cell_font_pt(cell, 9.0)
            # a merged cell (a group heading over several columns) sets its
            # text across every column it spans, not just its first
            span = max(1, int(getattr(cell, "span_width", 1) or 1))
            cw = sum(tbl.columns[k].width
                     for k in range(c, min(c + span, len(tbl.columns))))                 / EMU_PER_PX
            avail = max(16.0, cw - _cell_pad(cell, "left", 7.0)
                        - _cell_pad(cell, "right", 7.0))
            # PowerPoint breaks a table cell earlier than the column's bare
            # width suggests - measured off its own rendering, at about 80%
            # of the space between the margins.  Predicting the wrap matters
            # because every row's height feeds the NEXT row's position: a
            # row under-measured here walks every padlock below it down the
            # table, and the last rows' marks fall off the end entirely.
            if cell.text_frame.word_wrap is False:
                # A cell set to wrap="none" still breaks inside a table, but
                # only once its text is wider than the FULL space between the
                # margins, measured in its own face and weight - a label at
                # 98% of that stays on one line, one at 104% takes two.
                # Taking 80% here doubled rows PowerPoint draws single, and
                # the grown table ran off the bottom of the slide.
                run = next((r for p in cell.text_frame.paragraphs
                            for r in p.runs if r.text.strip()), None)
                face = run.font.name if run is not None else None
                bold = bool(run.font.bold) if run is not None else False
                n = sum(len(_wrap_lines(ln, avail, pt * PT_TO_PX, face, bold))
                        for ln in txt.split("\n"))
            else:
                avail *= _TABLE_FILL
                n = len(_wrap_lines(txt, avail, pt * PT_TO_PX))
            pad_v = (_cell_pad(cell, "top", 3.0)
                     + _cell_pad(cell, "bottom", 3.0) + _CELL_CHROME_PX)
            need = max(need, n * pt * LINE_LEAD * PT_TO_PX + pad_v)
        # Rounded UP: a written row height is only honoured if it is at
        # least what PowerPoint needs, and a row it grows is a row that
        # walks every padlock beneath it down the table.  Erring high costs
        # a little air; erring low costs the alignment.
        out.append(max(h, need * _ROW_SAFETY))
    return out


def _row_y(tbl, r, heights=None):
    hs = _row_heights(tbl) if heights is None else heights
    return sum(hs[:r])


def _cell_font_pt(cell, default=9.0):
    for p in cell.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size:
                return r.font.size.pt
        if p.font.size:
            return p.font.size.pt
    return default


def _place_blur(slide, cx, cy, pt_size, text, colour_hex="#605E5C"):
    """A blurred figure placed on the slide, centred on (cx, cy).

    For the charts a deck draws as SHAPES - a stacked column built from
    rectangles - where there is no data label to fill with a picture.
    """
    png = _blurred_value_png(text, colour_hex)
    if png is None:
        return None
    try:
        from PIL import Image
        with Image.open(png) as im:
            ar = im.width / im.height
    except Exception:
        return None
    h = pt_size * PT_TO_PX * 1.9          # the padded canvas, not the glyph
    try:
        return slide.shapes.add_picture(
            str(png), Emu(int((cx - h * ar / 2) * EMU_PER_PX)),
            Emu(int((cy - h / 2) * EMU_PER_PX)), height=Emu(int(h * EMU_PER_PX)))
    except Exception:
        return None


def _run_ink(para, default="#605E5C"):
    for r in para.runs:
        try:
            if r.font.color is not None and r.font.color.rgb is not None:
                return "#%s" % r.font.color.rgb
        except Exception:
            pass
    return default


def _place_lock(slide, art, cx, cy, pt_size, unit="", colour=None):
    """The padlock centred on (cx, cy), sized to a ``pt_size`` text run.

    Where the artwork with the unit baked in is available it is used whole;
    otherwise the plain padlock is placed and the unit written beside it.
    Returns the picture, or None when there is no artwork at all.
    """
    png = art.labelled(unit) if unit else art.plain()
    baked = bool(unit) and png is not None and png.name != "lock.png"
    if png is None:
        return None
    h = pt_size * (4.0 / 3.0) * LOCK_IMG_K
    pic = slide.shapes.add_picture(str(png), Emu(int(cx * EMU_PER_PX)),
                                   Emu(int((cy - h / 2) * EMU_PER_PX)),
                                   height=Emu(int(h * EMU_PER_PX)))
    pic.left = Emu(int((cx - pic.width / EMU_PER_PX / 2) * EMU_PER_PX))
    if unit and not baked:
        # no baked variant - write the unit beside the plain padlock
        w = pic.width / EMU_PER_PX
        tb = slide.shapes.add_textbox(
            Emu(int((cx + w / 2 + 1) * EMU_PER_PX)),
            Emu(int((cy - 8) * EMU_PER_PX)),
            Emu(int(18 * EMU_PER_PX)), Emu(int(16 * EMU_PER_PX)))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = Emu(0)
        tf.margin_top = tf.margin_bottom = Emu(0)
        run = tf.paragraphs[0].add_run()
        run.text = unit
        run.font.size = Pt(pt_size)
        if colour is not None:
            run.font.color.rgb = colour
    return pic


def _cell_pad(cell, side, default):
    try:
        v = getattr(cell, f"margin_{side}")
        return default if v is None else v / EMU_PER_PX
    except Exception:
        return default


def _blank_cell(cell):
    """Empty a cell but keep its run's font size.

    PowerPoint ignores an EMPTY run's formatting, falls back to its 18 pt
    default and GROWS the row to fit it - which would drag every padlock off
    the row it was placed on.  A single space keeps the size.
    """
    size = _cell_font_pt(cell)
    tf = cell.text_frame
    p = tf.paragraphs[0]
    for extra in list(tf.paragraphs)[1:]:
        extra._p.getparent().remove(extra._p)
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    run = p.add_run()
    run.text = " "
    run.font.size = Pt(size)
    return size


_A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
_R_NS = ('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
         'relationships"')
# what an invisible spacer of this shape does to the label's box, measured
# off PowerPoint's own rendering
_LABEL_SPACER = {"": "M", "%": "MM", "K": "MM"}


def _chart_image_rid(chart, path: Path):
    cache = getattr(chart.part, "_pptlock_rids", None)
    if cache is None:
        cache = {}
        chart.part._pptlock_rids = cache
    if path.name not in cache:
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        from pptx.parts.image import Image as PPTXImage, ImagePart
        part = ImagePart.new(chart.part.package, PPTXImage.from_file(str(path)))
        cache[path.name] = chart.part.relate_to(part, RT.IMAGE)
    return cache[path.name]


_SER_AFTER_DLBLS = ("c:trendline", "c:errBars", "c:cat", "c:val", "c:xVal",
                    "c:yVal", "c:shape", "c:smooth", "c:bubbleSize",
                    "c:bubble3D", "c:extLst")


def _ensure_series_dlbls(plot, series) -> None:
    """Give a series its own ``c:dLbls`` seeded from the PLOT's.

    Load-bearing.  Most decks switch labels on for a whole plot, so the
    settings ("show the value", the font) live on ``c:barChart/c:dLbls`` and
    the series has none.  Asking python-pptx for a single point's label makes
    it create a series-level ``c:dLbls`` from ITS defaults - every ``show*``
    off - and a series-level block overrides the plot's, so every label the
    lock did not touch silently disappears.  Copying the plot's settings down
    first is what keeps the open labels on the chart.
    """
    import copy
    ser = series._element
    if ser.find(qn("c:dLbls")) is not None:
        return
    src = plot._element.find(qn("c:dLbls"))
    if src is None:
        return
    new = copy.deepcopy(src)
    for d in new.findall(qn("c:dLbl")):      # per-point overrides are the
        new.remove(d)                        # plot's own, not this series'
    ser.insert_element_before(new, *_SER_AFTER_DLBLS)


# ═══════════════════ blurred figures + ONE padlock ═══════════════════
# A chart with a dozen withheld labels reads badly with a padlock on every
# one: the marks crowd out the picture they are protecting.  So the figure
# is SHOWN, blurred past reading, and the chart takes a single padlock.  The
# digits are not merely covered - the label's text is replaced by a picture,
# so the numbers are not in the file at all.
_BLUR_DIR = None
_BLUR_FONTS = ("C:/Windows/Fonts/segoeui.ttf", "C:/Windows/Fonts/arial.ttf",
               "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
               "/Library/Fonts/Arial.ttf")
# Enough that neighbouring digits run into one another and no figure can be
# read, but not so much that the number stops looking like a number: the
# reader should still see its LENGTH and shape.  Paired with an opacity a
# little over half, which is what makes it read as protected rather than as
# a smudge or a censor bar.
_BLUR_STRENGTH = 0.20          # blur radius as a fraction of the text height
_BLUR_ALPHA = 0.62             # ink opacity of the blurred figure
_BLUR_BASE_PX = 34             # rendered at this height, scaled down in place
# Measured off PowerPoint's own rendering: one "M" gives a label box ~0.96x
# its height and each further "M" adds ~0.59.  The picture is padded to that
# shape so the blurred figure is not stretched inside its box.
_SPACER_BOX_A, _SPACER_BOX_B = 0.37, 0.59


def _blurred_value_png(text, colour_hex="#605E5C"):
    """A PNG of ``text`` rendered and blurred past reading, padded to the
    shape of a data label's box.  Cached per (text, colour).  None if Pillow
    is not installed - the caller then falls back to the padlock."""
    global _BLUR_DIR
    try:
        from PIL import Image, ImageDraw, ImageFilter, ImageFont
    except Exception:
        return None
    if _BLUR_DIR is None:
        import tempfile
        _BLUR_DIR = Path(tempfile.gettempdir()) / "pptlock-blur"
        _BLUR_DIR.mkdir(parents=True, exist_ok=True)
    safe = "".join(c if c.isalnum() else "_" for c in str(text))[:48]
    out = _BLUR_DIR / f"{safe}_{colour_hex.lstrip('#')}.png"
    if out.exists():
        return out
    font = None
    for f in _BLUR_FONTS:
        if Path(f).exists():
            try:
                font = ImageFont.truetype(f, _BLUR_BASE_PX)
                break
            except Exception:
                pass
    if font is None:
        return None
    t = str(text)
    pad = int(_BLUR_BASE_PX * 0.35)       # room for the blur to fall off
    tw = int(font.getlength(t))
    H = _BLUR_BASE_PX + 2 * pad
    W = max(int((_SPACER_BOX_A + _SPACER_BOX_B * max(1, len(t))) * H),
            tw + 2 * pad)
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.text((W / 2, H / 2), t, font=font, anchor="mm",
           fill=tuple(int(colour_hex.lstrip("#")[i:i + 2], 16)
                      for i in (0, 2, 4)) + (int(255 * _BLUR_ALPHA),))
    img = img.filter(ImageFilter.GaussianBlur(_BLUR_BASE_PX * _BLUR_STRENGTH))
    img.save(out, optimize=True)
    return out


def _hide_chart_label(point) -> bool:
    """Leave a data label showing nothing at all.

    Not the same as withholding it: this is for the segments of a stacked
    column whose padlock is carried by the segment at the TOP, so the ones
    below must simply go quiet.
    """
    after = ("c:spPr", "c:txPr", "c:dLblPos", "c:showLegendKey", "c:showVal",
             "c:showCatName", "c:showSerName", "c:showPercent",
             "c:showBubbleSize", "c:separator")
    try:
        dLbl = point.data_label._get_or_add_dLbl()
        for tag in ("c:tx", "c:numFmt", "c:spPr"):
            for old_ in dLbl.findall(qn(tag)):
                dLbl.remove(old_)
        dLbl.insert_element_before(parse_xml(
            f'<c:numFmt {nsdecls("c")} formatCode="&quot;&quot;" '
            f'sourceLinked="0"/>'), *after)
        return True
    except Exception:
        return False


def _lift_label(point, dy_frac):
    """Nudge ONE data label up by ``dy_frac`` of the chart's height.

    A waterfall stacks its steps on an invisible base, so its labels can only
    sit INSIDE the bar - PowerPoint refuses "outside end" on a stacked plot.
    DrawingML will however carry a manual offset from wherever the label
    would otherwise land, which is how the padlock gets to sit ON TOP of the
    step instead of over it.
    """
    after = ("c:showLegendKey", "c:showVal", "c:showCatName", "c:showSerName",
             "c:showPercent", "c:showBubbleSize", "c:separator")
    try:
        dLbl = point.data_label._get_or_add_dLbl()
        # Pin the label to the TOP of its own bar first.  A step's label
        # otherwise sits at the bar's middle, and lifting a fixed amount from
        # there leaves it inside a tall bar and above a short one.
        for old_ in dLbl.findall(qn("c:dLblPos")):
            dLbl.remove(old_)
        dLbl.insert_element_before(parse_xml(
            f'<c:dLblPos {nsdecls("c")} val="inEnd"/>'), *after)
        for old_ in dLbl.findall(qn("c:layout")):
            dLbl.remove(old_)
        dLbl.insert(0, parse_xml(
            f'<c:layout {nsdecls("c")}><c:manualLayout>'
            f'<c:x val="0"/><c:y val="{dy_frac:.4f}"/>'
            f'</c:manualLayout></c:layout>'))
        return True
    except Exception:
        return False


def _blur_chart_label(chart, point, text, colour_hex="#605E5C") -> bool:
    """Show ONE data label's figure, blurred, in place of the figure."""
    png = _blurred_value_png(text, colour_hex)
    if png is None:
        return False
    dLbl = point.data_label._get_or_add_dLbl()
    after = ("c:spPr", "c:txPr", "c:dLblPos", "c:showLegendKey", "c:showVal",
             "c:showCatName", "c:showSerName", "c:showPercent",
             "c:showBubbleSize", "c:separator")
    for tag in ("c:tx", "c:numFmt", "c:spPr", "c:txPr"):
        for old in dLbl.findall(qn(tag)):
            dLbl.remove(old)
    rid = _chart_image_rid(chart, png)
    spacer = "M" * max(1, len(str(text)))
    dLbl.insert_element_before(parse_xml(
        f'<c:numFmt {nsdecls("c")} formatCode="&quot;{spacer}&quot;;'
        f'&quot;{spacer}&quot;;&quot;{spacer}&quot;" sourceLinked="0"/>'
    ), *after)
    dLbl.insert_element_before(parse_xml(
        f'<c:spPr {nsdecls("c")}><a:blipFill {_A_NS}>'
        f'<a:blip {_R_NS} r:embed="{rid}"/>'
        f'<a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr>'
    ), *after[1:])
    dLbl.insert_element_before(parse_xml(
        f'<c:txPr {nsdecls("c")}><a:bodyPr {_A_NS}/><a:lstStyle {_A_NS}/>'
        f'<a:p {_A_NS}><a:pPr><a:defRPr><a:solidFill>'
        f'<a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr>'
        f'</a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p>'
        f'</c:txPr>'
    ), *after[2:])
    return True


# Where a native chart actually plots, inside its frame.  Measured off
# PowerPoint's rendering of these panels: the value axis and the category
# labels take a fixed bite out of each edge.
_PLOT_INSETS = {
    "col":   {"l": 45.5, "r": 14.5, "t": 8.0, "b": 17.0},
    "col_lg": {"l": 45.5, "r": 14.5, "t": 8.0, "b": 46.0},
    "bar_h": {"l": 94.0, "r": 24.0, "t": 19.0, "b": 37.0},
}
# The padlock's height as a share of the plot's size - the geometric mean of
# width and height, so one ratio reads at the same weight on a wide, shallow
# chart and a tall, narrow one.
_CHART_LOCK_FRAC = 0.115
_CHART_LOCK_MIN_PX, _CHART_LOCK_MAX_PX = 28.0, 64.0


def _plot_box(rect, kind="col"):
    x, y, w, h = rect
    i = _PLOT_INSETS.get(kind, _PLOT_INSETS["col"])
    return (x + i["l"], y + i["t"],
            max(10.0, w - i["l"] - i["r"]), max(10.0, h - i["t"] - i["b"]))


def _chart_lock_pt(box):
    import math
    _x, _y, w, h = box
    px = math.sqrt(max(1.0, w) * max(1.0, h)) * _CHART_LOCK_FRAC
    px = max(_CHART_LOCK_MIN_PX, min(_CHART_LOCK_MAX_PX, px))
    return min(px, h * 0.55)


def _chart_lock(slide, art, box, fx=0.5, fy=0.5, dark=False, scale=1.0):
    """THE padlock for one chart, at (fx, fy) of its plot box and kept wholly
    inside it, so a mark never straddles a chart's border."""
    x, y, w, h = box
    ph = _chart_lock_pt(box) * scale
    pw = ph * 0.7
    cx = min(max(x + w * fx, x + pw / 2 + 2), x + w - pw / 2 - 2)
    cy = min(max(y + h * fy, y + ph / 2 + 2), y + h - ph / 2 - 2)
    png = art.dark() if dark else art.plain()
    if png is None:
        return False
    try:
        slide.shapes.add_picture(str(png), Emu(int((cx - pw / 2) * EMU_PER_PX)),
                                 Emu(int((cy - ph / 2) * EMU_PER_PX)),
                                 height=Emu(int(ph * EMU_PER_PX)))
    except Exception:
        return False
    return True


def _lock_chart_label(chart, point, art, unit="") -> bool:
    """Turn ONE data label into the padlock.

    A native chart draws a label from a number format, so the padlock can
    only reach it as the label box's PICTURE FILL, with the text reduced to
    an invisible spacer that sets the box's shape.  With no artwork the
    format alone is used, and the label reads "xx".
    """
    png = art.labelled(unit) if art.available else None
    dLbl = point.data_label._get_or_add_dLbl()
    after = ("c:spPr", "c:txPr", "c:dLblPos", "c:showLegendKey", "c:showVal",
             "c:showCatName", "c:showSerName", "c:showPercent",
             "c:showBubbleSize", "c:separator")
    for tag in ("c:tx", "c:numFmt", "c:spPr", "c:txPr"):
        for old in dLbl.findall(qn(tag)):
            dLbl.remove(old)
    if png is None:
        code = f'"xx{unit}";"xx{unit}";"xx{unit}"'.replace('"', "&quot;")
        dLbl.insert_element_before(parse_xml(
            f'<c:numFmt {nsdecls("c")} formatCode="{code}" sourceLinked="0"/>'
        ), *after)
        return True
    rid = _chart_image_rid(chart, png)
    spacer = _LABEL_SPACER.get(unit or "", "M")
    dLbl.insert_element_before(parse_xml(
        f'<c:numFmt {nsdecls("c")} formatCode="&quot;{spacer}&quot;;'
        f'&quot;{spacer}&quot;;&quot;{spacer}&quot;" sourceLinked="0"/>'
    ), *after)
    dLbl.insert_element_before(parse_xml(
        f'<c:spPr {nsdecls("c")}><a:blipFill {_A_NS}>'
        f'<a:blip {_R_NS} r:embed="{rid}"/>'
        f'<a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr>'
    ), *after[1:])
    dLbl.insert_element_before(parse_xml(
        f'<c:txPr {nsdecls("c")}><a:bodyPr {_A_NS}/><a:lstStyle {_A_NS}/>'
        f'<a:p {_A_NS}><a:pPr><a:defRPr><a:solidFill>'
        f'<a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr>'
        f'</a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p>'
        f'</c:txPr>'
    ), *after[2:])
    return True


def _norm_dash(s: str) -> str:
    """Year spans are written with en/em dashes in one deck and hyphens in
    another; a rule should not have to know which."""
    return (str(s).lower().replace("–", "-").replace("—", "-")
            .replace("−", "-"))


def _slide_heading(slide) -> str:
    """The page's own heading - the text of the title band across the top."""
    best, best_y = "", None
    for _p, shp in _iter_shapes(slide.shapes):
        if not getattr(shp, "has_text_frame", False):
            continue
        t = shp.text_frame.text.strip()
        if not t:
            continue
        r = _abs_rect(shp)
        if r is None or r[1] > 100 or r[2] < 180:
            continue                       # not in the title band
        if best_y is None or r[1] < best_y:
            best, best_y = t.replace("\n", " "), r[1]
    return " ".join(best.split())


def rulebook(deck=None, policy: LockPolicy = None):
    """Print the page rulebook, and - given a deck - which slides it reaches.

    The engine that writes these decks puts the same heading on the same page
    in every market, so a rule written once keeps working.  Run this against
    a new report before locking it: a page whose heading was renamed shows up
    as unmatched here rather than as a surprise in a client deck.
    """
    policy = policy or LockPolicy()
    for what, why in CHART_RULES:
        print()
        print(f"  BY CHART SHAPE: {what}")
        for i, ln in enumerate(_wrap_lines(why, 430, 7.0)):
            print(f"      {'note   : ' if i == 0 else '         '}{ln}")
    hits = {}
    if deck is not None:
        prs = deck if isinstance(deck, Presentation().__class__) else             Presentation(str(deck))
        for n, slide in enumerate(prs.slides, start=1):
            head = _slide_heading(slide)
            for pat, _r in policy.page_rules:
                if pat.lower() in head.lower():
                    hits.setdefault(pat, []).append(n)
                    break
            else:
                if head:
                    hits.setdefault(None, []).append(n)
    for pat, rule in policy.page_rules:
        keys = {k: v for k, v in rule.items() if k != "note"}
        where = hits.get(pat)
        print()
        print(f"  {pat}")
        if where is not None:
            print(f"      slides : {', '.join(map(str, where))}")
        if keys:
            print(f"      rule   : {keys}")
        else:
            print(f"      rule   : (general behaviour)")
        if rule.get("note"):
            for i, ln in enumerate(_wrap_lines(rule["note"], 430, 7.0)):
                print(f"      {'note   : ' if i == 0 else '         '}{ln}")
    if hits.get(None):
        u = hits[None]
        print()
        print(f"  {len(u)} slides used the general rules "
              f"(first few: {', '.join(map(str, u[:12]))})")


# The rule keys that KEEP a value readable, as opposed to saying how a page
# is marked.  `policy.page_exemptions=False` drops them, and dropping them
# here - in the one place every consumer reads its rule - is what makes the
# switch reach `scan` and `apply_locks` alike.
_EXEMPTION_KEYS = ("show_per_col", "open_rows", "grid_open", "open_cats",
                   "cards_only", "prose_open_regions")


def _rule_for(heading, policy) -> dict:
    """The rule a heading matches, or {} for the general behaviour."""
    h = (heading or "").lower()
    for pattern, rule in (policy.page_rules or ()):
        if pattern.lower() in h:
            if not policy.page_exemptions:
                rule = {k: v for k, v in rule.items()
                        if k not in _EXEMPTION_KEYS
                        and not (k == "lock" and v is False)}
            return rule
    return {}


def _is_waterfall(chart) -> bool:
    """True for a build-up chart: two stacked series, one of them invisible.

    A waterfall is drawn by standing each visible step on a transparent
    spacer, so it is recognisable by its SHAPE rather than by the page it
    sits on - which is what lets one rule cover all 64 of them in this deck,
    and however many the next market happens to have.
    """
    try:
        if "STACKED" not in str(chart.chart_type):
            return False
        ser = list(chart.series)
        return len(ser) == 2 and any(_series_is_invisible(x) for x in ser)
    except Exception:
        return False


def _is_ring(chart) -> bool:
    """True for a pie or a doughnut - a chart whose mark belongs in its
    MIDDLE.

    A ring says one thing, that its split is withheld, and the padlock in the
    hole says it once.  How many slices it happens to have does not change
    that: a two-line ring reads exactly like a ten-line one, so the count
    that decides between blurring and a mark per figure elsewhere is not
    allowed to decide it here.
    """
    try:
        t = str(chart.chart_type)
    except Exception:
        return False
    return "DOUGHNUT" in t or "PIE" in t


def _chart_kind(chart) -> str:
    """Which plot-inset set this chart is measured with."""
    try:
        t = str(chart.chart_type)
    except Exception:
        return "col"
    if "BAR_" in t:                       # horizontal bars
        return "bar_h"
    try:
        if chart.has_legend:
            return "col_lg"
    except Exception:
        pass
    return "col"


# ══════════════ the padlock INSIDE a running sentence ════════════════
# DrawingML has no inline picture: a padlock in running text can only be the
# character, and PowerPoint draws the emoji as a flat black lock, nothing
# like the artwork on the charts.  So a sentence that withholds a figure is
# CUT at that figure, each piece is drawn in its own little box, and the
# artwork is placed in the gap between them.  The line breaks are computed
# here rather than left to PowerPoint, because the padlock's position has to
# be known before the text is laid out.
PT_TO_PX = 4.0 / 3.0
# a paragraph's line spacing multiplies the FONT's line height, not its
# point size - carry the 1.2 the type designer built in
LINE_LEAD = 1.2
# The width model runs a little ahead of PowerPoint's own metrics; measured
# on rendered slides at 8, 9.6 and 10.5 pt.
TEXT_K = 0.975

_CW = {'0': 0.5552, '1': 0.4023, '2': 0.5552, '3': 0.5552, '4': 0.5762, '5': 0.5552, '6': 0.5581, '7': 0.5361, '8': 0.5552, '9': 0.5581, 'a': 0.522, 'b': 0.6025, 'c': 0.4702, 'd': 0.6025, 'e': 0.5312, 'f': 0.3447, 'g': 0.6025, 'h': 0.582, 'i': 0.2612, 'j': 0.2612, 'k': 0.5249, 'l': 0.2612, 'm': 0.8857, 'n': 0.5835, 'o': 0.5972, 'p': 0.6025, 'q': 0.6025, 'r': 0.3701, 's': 0.4312, 't': 0.3613, 'u': 0.5835, 'v': 0.5073, 'w': 0.7563, 'x': 0.501, 'y': 0.5083, 'z': 0.4644, 'A': 0.6714, 'B': 0.604, 'C': 0.6211, 'D': 0.7173, 'E': 0.5176, 'F': 0.5024, 'G': 0.6973, 'H': 0.7354, 'I': 0.2915, 'J': 0.3965, 'K': 0.6108, 'L': 0.4888, 'M': 0.9243, 'N': 0.7671, 'O': 0.7559, 'P': 0.5845, 'Q': 0.7559, 'R': 0.6226, 'S': 0.5444, 'T': 0.5518, 'U': 0.7031, 'V': 0.6416, 'W': 0.9658, 'X': 0.6191, 'Y': 0.5771, 'Z': 0.5869, '!': 0.3037, '"': 0.4375, '#': 0.5913, '$': 0.5552, '%': 0.8403, '&': 0.7153, "'": 0.2583, '(': 0.332, ')': 0.332, '*': 0.4341, '+': 0.6943, ',': 0.2412, '-': 0.4019, '.': 0.2412, '/': 0.4136, ':': 0.2412, ';': 0.2412, '<': 0.6943, '=': 0.6943, '>': 0.6943, '?': 0.4438, '@': 0.9546, '[': 0.332, '\\': 0.4048, ']': 0.332, '^': 0.6943, '_': 0.415, '`': 0.2886, '{': 0.332, '|': 0.2783, '}': 0.332, '~': 0.6943, ' ': 0.2749}
_DEF_CW = 0.56


# ── whose face the width model is measuring ──────────────────────────
# `_CW` is Segoe UI - the Fact.MR template's face, which is what it was
# calibrated on.  The FMI template sets its prose in Calibri, ~8% narrower,
# so every fragment measured ~8% wide; `_draw_marked_line` only ever moves
# the cursor ON by the model, so that surplus was paid out as blank space in
# FRONT of the padlock, and it accumulated with the length of the sentence:
# 45px before the lock on a long line, 8px on a short one.  Measure the real
# face wherever its file can be found, and fall back to the table - scaled
# per family - where it cannot, so a machine with no font files still lays
# out no worse than before.
_FONT_DIRS = (r"C:\Windows\Fonts", "/usr/share/fonts/truetype/msttcorefonts",
              "/usr/share/fonts/truetype", "/usr/share/fonts",
              "/usr/local/share/fonts", "/Library/Fonts",
              "/System/Library/Fonts")
_FONT_FILES = {
    ("calibri", False): ("calibri.ttf", "Carlito-Regular.ttf"),
    ("calibri", True): ("calibrib.ttf", "Carlito-Bold.ttf"),
    ("segoe ui", False): ("segoeui.ttf",),
    ("segoe ui", True): ("segoeuib.ttf",),
    ("arial", False): ("arial.ttf", "Arimo-Regular.ttf"),
    ("arial", True): ("arialbd.ttf", "Arimo-Bold.ttf"),
}
# how wide a face runs against the Segoe-calibrated table, for the fallback
_FACE_K = {"calibri": 0.92, "segoe ui": 1.0, "arial": 1.0}
# A font object can only be built at a WHOLE number of pixels, and these
# sentences are set at 11.5pt = 15.33px.  Measuring at 15 came out 2% short,
# which on a long line put the padlock inside the last letter.  So measure
# once at a big reference size and scale: an advance is linear in the size.
_FACE_REF_PX = 512
_FACE_FONT = {}


def _face_font(face, bold):
    """A measurable font object for ``face``, at the reference size."""
    if not face:
        return None
    key = (str(face).strip().lower(), bool(bold))
    if key in _FACE_FONT:
        return _FACE_FONT[key]
    fn = None
    try:
        from PIL import ImageFont
        for stem in _FONT_FILES.get((key[0], key[1]), ()):
            for folder in _FONT_DIRS:
                f = Path(folder) / stem
                if f.exists():
                    fn = ImageFont.truetype(str(f), _FACE_REF_PX)
                    break
            if fn is not None:
                break
    except Exception:
        fn = None
    _FACE_FONT[key] = fn
    return fn


def text_width(s, font_px: float, face=None) -> float:
    """Width of ``s`` at ``font_px``, in px, before the calibration.

    The table is Segoe UI; ``face`` scales it to the family actually being
    set.  Used only when the face's own file cannot be measured.
    """
    w = sum(_CW.get(c, _DEF_CW) for c in str(s)) * font_px
    return w * _FACE_K.get(str(face or "").strip().lower(), 1.0)


def _tw(text, font_px: float, bold=False, face=None) -> float:
    """Width of ``text`` as PowerPoint will actually set it, in px.

    The face's own metrics where they can be read - that is the same advance
    PowerPoint lays the run out on, so the cursor lands where the text really
    ends.  Otherwise the table, calibrated (TEXT_K) and widened for bold the
    way it always was.
    """
    fn = _face_font(face, bold)
    if fn is not None:
        try:
            return fn.getlength(str(text)) * font_px / _FACE_REF_PX
        except Exception:
            pass
    return text_width(text, font_px, face) * TEXT_K * (1.06 if bold else 1.0)


def _lock_slot_px(pt_size: float) -> float:
    """The room ONE padlock takes on a line: the picture plus the clear air
    either side of it."""
    fs_px = pt_size * PT_TO_PX
    return (fs_px * LOCK_IMG_K * LOCK_ART_ASPECT
            + 2 * max(3.0, fs_px * 0.32))


def _wrap_lines(text, max_w, font_px, face=None, bold=False):
    """Greedy word wrap at the width the deck actually renders at."""
    lines, cur = [], ""
    for word in str(text).split():
        trial = f"{cur} {word}" if cur else word
        if cur and _tw(trial, font_px, bold, face) > max_w:
            lines.append(cur)
            cur = word
        else:
            cur = trial
    if cur or not lines:
        lines.append(cur)
    return lines


def _wrap_marked(text, max_w, pt_size, mark, face=None, bold=False):
    """Wrap a sentence whose padlocks will be DRAWN rather than typed.

    The marker is one character to the width model and a good deal wider
    once it is a picture, so wrapping the raw text puts too much on a line.
    Each marker stands in as a run of "M"s as wide as the artwork will be -
    a run, not spaces, so the wrapper cannot break the padlock away from the
    unit that follows it - and the markers go back once the breaks are known.
    """
    s = str(text)
    fs_px = pt_size * PT_TO_PX
    if mark not in s:
        return _wrap_lines(s, max_w, fs_px, face, bold)
    em = _tw("M", fs_px, bold, face) or 1.0
    filler = "M" * max(1, int(round(_lock_slot_px(pt_size) / em)))
    s = re.sub(re.escape(mark) + r"\s*([%×x]|[KMB]n?(?![A-Za-z]))",
               lambda m: filler + m.group(1), s)
    lines = _wrap_lines(s.replace(mark, filler), max_w, fs_px, face, bold)
    return [ln.replace(filler, mark) for ln in lines]


def _plain_box(slide, x, y, w, h):
    tb = slide.shapes.add_textbox(Emu(int(x * EMU_PER_PX)),
                                  Emu(int(y * EMU_PER_PX)),
                                  Emu(int(max(4.0, w) * EMU_PER_PX)),
                                  Emu(int(max(4.0, h) * EMU_PER_PX)))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.word_wrap = False
    return tb


# The padlock PNG carries transparent margin around the drawing - the file is
# 329x480 but the ink occupies (37,38)-(316,471).  A picture is positioned by
# its FILE box, so placing it "next to" a word by that box leaves a gap of
# invisible pixels: the lock reads as floating away from the value it belongs
# to.  These are the margins as fractions of the file, measured once, so a
# lock can be placed and advanced by its INK instead.
_ART_INK = {}


def _art_ink(png):
    """(left, right, width, height) of the ink inside ``png``, 0..1."""
    key = str(png)
    if key in _ART_INK:
        return _ART_INK[key]
    box = (0.0, 0.0, 1.0, 1.0)
    try:
        from PIL import Image
        with Image.open(png) as im:
            a = im.convert("RGBA").getchannel("A").getbbox()
        if a:
            box = (a[0] / im.width, 1 - a[2] / im.width,
                   (a[2] - a[0]) / im.width, (a[3] - a[1]) / im.height)
    except Exception:
        pass
    _ART_INK[key] = box
    return box


# One gap, used on both sides of every inline padlock in the deck, so the
# spacing is the same on every line at every size.  A shade under a space.
# Matched to the width of a space at the same size: the padlock then reads as
# a word in the sentence rather than as something wedged against the last
# letter.  Tighter than this and "2036" and the lock touch.
_INLINE_GAP = 0.27            # of the text's own height
# How far a lifted label rises above the step it belongs to, in px.  One
# label height plus a little air.
_LABEL_LIFT_PX = 21.0
# The padlock in a doughnut's hole, as a share of the chart's height - sized
# so it fills the hole the way the reference artwork does.
_DONUT_LOCK_FILL = 0.62      # of the hole's diameter


def _ring_geom(chart, rect):
    """Where a doughnut's ring actually sits: (cx, cy, hole_diameter).

    The frame's centre is NOT the ring's.  Every ring in these decks pins its
    plot area to the left ~68% and puts the legend in the rest, so a padlock
    centred on the frame lands off the hole entirely - by 57px on the company
    share pages.  The pinned layout is read back here instead.
    """
    x, y, w, h = rect
    ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    try:
        lay = chart._chartSpace.find(
            ".//c:plotArea/c:layout/c:manualLayout", ns)
    except Exception:
        lay = None
    if lay is not None:
        g = {}
        for child in lay:
            tag = child.tag.split("}")[1]
            if tag in ("x", "y", "w", "h"):
                try:
                    g[tag] = float(child.get("val"))
                except (TypeError, ValueError):
                    pass
        if {"x", "y", "w", "h"} <= set(g):
            px, py = x + w * g["x"], y + h * g["y"]
            pw, ph = w * g["w"], h * g["h"]
        else:
            px, py, pw, ph = _plot_box(rect, "pie")
    else:
        px, py, pw, ph = _plot_box(rect, "pie")
    ring = min(pw, ph)                    # the ring is a circle in that box
    hole = ring * _hole_frac(chart)
    return px + pw / 2, py + ph / 2, hole


def _hole_frac(chart):
    """The hole as a fraction of the ring's diameter."""
    ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    try:
        el = chart._chartSpace.find(".//c:doughnutChart/c:holeSize", ns)
        return max(0.2, min(0.9, float(el.get("val")) / 100.0))
    except Exception:
        return 0.5


_INK_CENTRE = {}


def _ink_centre(png):
    """Where the drawn mark sits inside its own PNG, 0..1.

    The padlock art carries uneven transparent margins - its ink centres at
    53% across and 52% down - so a picture centred by its BOX hangs low and
    right of the point it was meant to mark.
    """
    key = str(png)
    if key not in _INK_CENTRE:
        c = (0.5, 0.5)
        try:
            from PIL import Image
            with Image.open(png) as im:
                a = im.convert("RGBA").getchannel("A").getbbox()
                if a:
                    c = ((a[0] + a[2]) / 2 / im.width,
                         (a[1] + a[3]) / 2 / im.height)
        except Exception:
            pass
        _INK_CENTRE[key] = c
    return _INK_CENTRE[key]


def _centre_art(shape, png, cx, cy, height):
    """Sit ``shape`` so the padlock's INK centres on (cx, cy)."""
    aspect = (shape.width / shape.height) if shape.height else 0.7
    shape.height = Emu(int(height * EMU_PER_PX))
    shape.width = Emu(int(height * aspect * EMU_PER_PX))
    w = shape.width / EMU_PER_PX
    ix, iy = _ink_centre(png)
    shape.left = Emu(int((cx - w * ix) * EMU_PER_PX))
    shape.top = Emu(int((cy - height * iy) * EMU_PER_PX))


# ── performance marks (▲ ▶ ▼) inside a redrawn sentence ─────────────
# The engine sets each mark in its own colour - ▲ green, ▼ red - as a run of
# its own.  A redrawn line is set in ONE colour, the sentence's, so without
# this every ▲/▼ in a locked read-out came out in body ink.  The span is the
# engine's own (components._MARK_SPAN): the colour word the sentence names the
# mark with, the mark, and the figure it labels.
_MARK_CHARS = "▲▶▼"
_MARK_WORD = {"green": "▲", "red": "▼", "amber": "▶", "orange": "▶"}
_MARK_SPAN_RE = re.compile(
    r"(?:(?P<word>green|red|amber|orange)\s+)?(?P<mark>[▲▶▼])"
    r"(?:\s*[+\-−]?[\d,]+(?:\.\d+)?\s*%?)?", re.IGNORECASE)
# used only when the source paragraph set a mark in body ink
_MARK_FALLBACK = {"▲": RGBColor(0x2E, 0x7D, 0x32),
                  "▶": RGBColor(0xA8, 0x6F, 0x00),
                  "▼": RGBColor(0xE4, 0x4D, 0x3A)}


def _mark_colours(para) -> dict:
    """{mark: colour} for the marks a paragraph sets in a colour of their own.

    Read BEFORE the paragraph is rewritten: `_rewrite_paragraph` keeps the run
    boundaries by position, so once a figure is swapped for the padlock marker
    the coloured run no longer sits on its mark.  A mark in the sentence's own
    ink is left out, so the fallback colour can take it.
    """
    runs = list(para.runs)
    if not runs:
        return {}
    body = None
    try:
        main = max(runs, key=lambda r: len(r.text or ""))
        if main.font.color is not None and main.font.color.type is not None:
            body = main.font.color.rgb
    except Exception:
        pass
    out = {}
    for r in runs:
        t = r.text or ""
        found = [m for m in _MARK_CHARS if m in t]
        if not found:
            continue
        try:
            if r.font.color is None or r.font.color.type is None:
                continue
            rgb = r.font.color.rgb
        except Exception:
            continue
        if rgb == body:
            continue
        for m in found:
            out.setdefault(m, rgb)
    return out


def _mark_pieces(text, colour, marks):
    """``text`` as [(piece, colour)] - each ▲/▶/▼ span in its own colour."""
    if not any(m in text for m in _MARK_CHARS):
        return [(text, colour)]
    marks = marks or {}
    out, i = [], 0
    for m in _MARK_SPAN_RE.finditer(text):
        if m.start() > i:
            out.append((text[i:m.start()], colour))
        mk = _MARK_WORD.get((m.group("word") or "").lower()) or m.group("mark")
        out.append((m.group(0), marks.get(mk) or _MARK_FALLBACK.get(mk, colour)))
        i = m.end()
    if i < len(text):
        out.append((text[i:], colour))
    return [p for p in out if p[0]]


def _draw_marked_line(slide, art, x, y, w, line_px, text, pt_size, colour,
                      mark, bold=False, font=None, marks=None):
    """Draw ONE wrapped line whose padlocks are drawn as artwork.

    Cutting the line at each marker is what keeps it clean.  Reserving a gap
    of spaces instead would put the padlock at a position taken from a width
    MODEL of the text before it, and the model and PowerPoint's own metrics
    disagree by a pixel per word - enough, by the middle of a sentence, to
    drop the lock on top of the next word.  Here the model only ever decides
    how far to move ON, so the disagreement shows as spacing, never overlap.
    """
    fs_px = pt_size * PT_TO_PX
    gap = max(1.5, fs_px * _INLINE_GAP)
    png = art.plain()
    ink_l, _ink_r, ink_w, _ink_h = _art_ink(png) if png else (0, 0, 1.0, 1.0)
    pic_h = pt_size * PT_TO_PX * LOCK_IMG_K
    pic_w = pic_h * (LOCK_ART_ASPECT / max(0.01, ink_w))
    lock_w = pic_w * ink_w                     # the INK's width
    parts = str(text).split(mark)
    last = len(parts) - 1
    cx = x
    for i, part in enumerate(parts):
        if i:
            # the padlock's INK starts one gap past the word, and the next
            # word starts one gap past the ink - the same on every line
            pic = _place_lock(slide, art, cx + pic_w / 2,
                              y + line_px / 2, pt_size, "", colour)
            if pic is not None:
                pic.left = Emu(int((cx - pic_w * ink_l) * EMU_PER_PX))
            cx += lock_w + gap
            part = part.lstrip()
        if i < last:
            # the gap before the padlock IS the gap - a trailing space would
            # add a second one and the lock would drift right
            part = part.rstrip()
        if not part.strip():
            continue
        room = x + w - cx
        if room <= 4:
            break
        tw = _tw(part, fs_px, bold, font)
        tb = _plain_box(slide, cx, y, min(tw + 8, room), line_px + 2)
        # one run per colour - a ▲/▼ keeps the colour the engine gave it; the
        # face and size are the same throughout, so the width does not change
        for piece, ink in _mark_pieces(part, colour, marks):
            run = tb.text_frame.paragraphs[0].add_run()
            run.text = piece
            run.font.size = Pt(pt_size)
            run.font.bold = bold
            if font:
                run.font.name = font      # the model is calibrated on the
                                          # deck's own face, not the theme's
            if ink is not None:
                run.font.color.rgb = ink
        cx += tw
        if i < last:
            cx += gap                     # the gap that opens the padlock


def _marked_line_w(text, pt_size, mark, bold=False, face=None) -> float:
    """How wide ``_draw_marked_line`` will lay one line out.

    The same walk that function makes, without drawing anything: a word
    advances by the width model, a padlock by its INK plus the gap either
    side.  A line has to be measured BEFORE it is drawn for a centred or
    right-aligned paragraph to know where to start it, which is why the walk
    is counted here rather than reported back by the drawing.
    """
    fs_px = pt_size * PT_TO_PX
    gap = max(1.5, fs_px * _INLINE_GAP)
    # the ink's width, which is what _draw_marked_line advances by: the
    # picture is scaled so its INK measures LOCK_ART_ASPECT of the type's
    # height, whatever transparent margin the file carries around it
    lock_w = fs_px * LOCK_IMG_K * LOCK_ART_ASPECT
    parts = str(text).split(mark)
    last = len(parts) - 1
    cx = end = 0.0
    for i, part in enumerate(parts):
        if i:
            end = cx + lock_w           # where the padlock's ink stops
            cx += lock_w + gap
            part = part.lstrip()
        if i < last:
            part = part.rstrip()
        if not part.strip():
            continue                    # no word, so no gap after it either
        cx += _tw(part, fs_px, bold, face)
        end = cx
        if i < last:
            cx += gap
    # the CURSOR runs one gap past the last mark; the line stops at its ink
    return end


def _align_dx(align, line_w, avail_w) -> float:
    """How far into its box a line starts, for the paragraph's own alignment.

    A redrawn line is placed by hand, so PowerPoint's alignment no longer
    applies to it - the shape it came from is what states it.  Miss this and
    a KPI card's centred read-out, cut into pieces so the padlock can be
    drawn, is re-laid hard against the card's left edge while every card
    beside it stays centred.
    """
    slack = max(0.0, avail_w - line_w)
    if align == PP_ALIGN.CENTER:
        return slack / 2
    if align == PP_ALIGN.RIGHT:
        return slack
    return 0.0


def _para_style(para, default_pt=10.5):
    """(font size in pt, colour, bold) for a paragraph.

    Read off the run carrying the most text, not the first one: these
    paragraphs open with a coloured bullet glyph ("▪ ") in its own run, and
    taking the colour from that would repaint the whole sentence with it.
    """
    runs = list(para.runs)
    if not runs:
        return default_pt, None, False, None
    main = max(runs, key=lambda r: len(r.text or ""))
    pt = main.font.size.pt if main.font.size is not None else default_pt
    if pt == default_pt:
        for r in runs:
            if r.font.size is not None:
                pt = r.font.size.pt
                break
    col = None
    try:
        if main.font.color is not None and main.font.color.type is not None:
            col = main.font.color.rgb
    except Exception:
        pass
    # The width model is calibrated on the deck's own face.  Re-drawing a
    # fragment without naming it let PowerPoint fall back to the theme font,
    # so every piece was set in a face the model had never measured and the
    # error grew along the line.
    name = next((r.font.name for r in runs if r.font.name), None)
    return pt, col, bool(main.font.bold), name


def _redraw_prose_shape(slide, shp, art, policy, mark_cols=None):
    """Re-lay a text shape whose sentences carry the marker, so each padlock
    is placed as ARTWORK where its figure stood.

    The shape's own text is cleared and every line is drawn into its own
    little box.  The line breaks have to be decided here - the padlock's
    position is only knowable once they are - so the wrap is computed with
    the same width model the report engine draws its own panels with.

    The type shrinks a notch at a time if the re-laid text would run past
    the box it came from: a sentence with a padlock in it is a little wider
    than the same sentence with the figure, and a panel that overflowed
    would be worse than one set a half-point smaller.

    ``mark_cols`` maps a paragraph index to the ▲/▼ colours read off it
    BEFORE its figures were swapped (see `_mark_colours`); a paragraph not
    in it is read now, its runs being untouched.
    """
    rect = _abs_rect(shp)
    if rect is None:
        return False
    tf = shp.text_frame
    mark = policy.prose_mark
    mark_cols = mark_cols or {}
    paras = []
    for pi, para in enumerate(tf.paragraphs):
        txt = _para_text(para)
        if not txt.strip():
            paras.append(None)          # a blank line still takes its space
            continue
        marks = mark_cols[pi] if pi in mark_cols else _mark_colours(para)
        # the paragraph's own alignment travels with it: the lines are placed
        # by hand from here on, so nothing else can carry it (see _align_dx)
        paras.append((txt,) + _para_style(para) + (para.alignment, marks))
    if not any(p for p in paras):
        return False

    x, y, w, h = rect
    try:
        ml = tf.margin_left / EMU_PER_PX
        mr = tf.margin_right / EMU_PER_PX
        mt = tf.margin_top / EMU_PER_PX
        mb = tf.margin_bottom / EMU_PER_PX
    except Exception:
        ml = mr = mt = mb = 0.0
    try:
        anchor = tf.vertical_anchor
    except Exception:
        anchor = None
    avail_w = max(24.0, w - ml - mr)

    def lay(scale):
        """Wrapped lines and total height at ``scale`` x the original size."""
        out, total = [], 0.0
        for item in paras:
            if item is None:
                out.append(None)
                total += 6.0
                continue
            txt, pt, col, bold, font, align, marks = item
            pt = pt * scale
            lines = _wrap_marked(txt, avail_w, pt, mark, font, bold)
            lp = pt * LINE_LEAD * PT_TO_PX
            out.append((lines, pt, col, bold, lp, font, align, marks))
            total += lp * len(lines) + lp * 0.30      # space after
        return out, total

    scale = 1.0
    laid, total = lay(scale)
    while total > h - mt - 2 and scale > 0.72:
        scale -= 0.04
        laid, total = lay(scale)
    # Past that floor the text may outgrow its box, but never the SLIDE: a
    # line drawn below the page's bottom edge is simply lost to the reader.
    try:
        page_h = slide.part.package.presentation_part.presentation \
            .slide_height / EMU_PER_PX
    except Exception:
        page_h = None
    while page_h and y + mt + total > page_h - 2 and scale > 0.5:
        scale -= 0.04
        laid, total = lay(scale)

    for para in list(tf.paragraphs):
        for r in list(para.runs):
            r.text = ""

    # ``total`` counts a space AFTER the last paragraph, which no renderer
    # puts inside the block - leave it out or a middle-anchored block sits
    # half that space high.
    block = total - 0.30 * next((it[4] for it in reversed(laid) if it), 0.0)
    # The frame's own anchoring, which the hand-placed lines no longer get for
    # free.  A KPI card's text is MIDDLE-anchored; laid from the top it stands
    # a few pixels above every card beside it.
    cy = y + mt
    if anchor == MSO_ANCHOR.MIDDLE:
        cy = y + max(mt, (h - block) / 2)
    elif anchor == MSO_ANCHOR.BOTTOM:
        cy = y + max(mt, h - mb - block)
    for item in laid:
        if item is None:
            cy += 6.0
            continue
        lines, pt, col, bold, lp, font, align, marks = item
        for ln in lines:
            dx = _align_dx(align, _marked_line_w(ln, pt, mark, bold, font),
                           avail_w)
            _draw_marked_line(slide, art, x + ml + dx, cy, avail_w - dx, lp,
                              ln, pt, col, mark, bold, font, marks)
            cy += lp
        cy += lp * 0.30
    return True


def _shape_at(slide, path):
    """The shape a candidate's index path points at."""
    shapes = slide.shapes
    for i in path[:-1]:
        shapes = shapes[i].shapes
    return shapes[path[-1]]


def apply_locks(deck, out_path=None, policy: LockPolicy = None,
                decisions=None, assets_dir=None) -> LockReport:
    """Withhold every value ``decisions`` names, and save to ``out_path``.

    ``decisions`` is a set of ``Candidate.key`` strings - what a reviewer
    ticked.  Pass None to take the policy's own proposal, which is what
    ``lock_pptx`` does.
    """
    policy = policy or LockPolicy()
    prs = deck if isinstance(deck, Presentation().__class__) else \
        Presentation(str(deck))
    art = LockArt(assets_dir)
    rep = LockReport()
    cands = scan(prs, policy)
    rep.candidates = cands
    if decisions is None:
        chosen = [c for c in cands if c.suggested]
    else:
        keys = set(decisions)
        chosen = [c for c in cands if c.key in keys]
    if not art.available:
        rep.notes.append("no padlock artwork found - used 'xx' placeholders "
                         "(install Pillow, or point assets_dir at a folder "
                         "holding lock.png)")

    slides = list(prs.slides)

    # A size written into a category label is withheld for EVERY chart, not
    # only the ones holding candidates: the figure is in the axis text rather
    # than in a data label, so nothing else in this pass would reach it.
    if policy.lock_charts:
        for slide in slides:
            for _p, shp in _iter_shapes(slide.shapes):
                if getattr(shp, "has_chart", False):
                    rep.locked += _mask_chart_categories(shp.chart, policy)

    # Group by the thing that has to be edited AS A WHOLE: a table is emptied
    # once and its padlocks placed together from one set of row heights, and
    # a paragraph is rewritten once with all of its figures at once.
    by_shape = {}
    for c in chosen:
        para = c.ref[0] if c.kind == "prose" else None
        if c.kind == "table-prose":
            para = (c.ref[0], c.ref[1])          # one group per cell
        by_shape.setdefault((c.slide, c.shape_path, c.kind, para),
                            []).append(c)

    # Pictures are opaque, so a chart or table pasted as one cannot be
    # touched.  Only BIG ones are worth reporting: a deck is full of small
    # pictures - logos, icons, the padlocks themselves - and counting those
    # would turn the report into noise.
    for slide in slides:
        for _p, shp in _iter_shapes(slide.shapes):
            r = _abs_rect(shp) if shp.shape_type == 13 else None
            if r and r[2] >= 160 and r[3] >= 110:
                rep.skipped_images += 1

    # shapes whose sentences now carry the marker and must be REDRAWN, so the
    # padlock arrives as artwork instead of a character PowerPoint would draw
    # from the emoji font (see _redraw_prose_shape)
    prose_shapes = set()
    prose_marks = {}                    # shape -> {paragraph: ▲/▼ colours}
    hole_locks = []                     # doughnuts, resolved after the pass
    shape_blurs = {}                    # shape-drawn charts: one lock each

    # Each slide's own treatment, looked up once from its heading.  Pages the
    # rulebook does not name are listed in the report, so a heading that was
    # renamed in the engine shows up here rather than in a client deck.
    rules, unmatched = {}, {}
    for i, sl in enumerate(slides, start=1):
        head = _slide_heading(sl)
        r = _rule_for(head, policy)
        rules[i] = r
        if not r and head:
            unmatched.setdefault(head, i)
    if policy.page_rules:
        rep.notes.append(f"{len(rules) - sum(1 for r in rules.values() if r)}"
                         f" of {len(rules)} slides used the general rules")

    # A padlock is placed in the MIDDLE of its cell, so a column whose
    # figures are left-aligned reads as two columns - numbers down one edge,
    # marks down the centre.  Every column that HOLDS figures is centred so
    # the two line up; columns of names or notes are left alone.  This runs
    # before any cell is emptied, which is what keeps a numeric column
    # recognisable as one.
    for si_, sl_ in enumerate(slides, start=1):
        if rules.get(si_, {}).get("table_align") == "left":
            continue
        for _p_, sh_ in _iter_shapes(sl_.shapes):
            if not getattr(sh_, "has_table", False):
                continue
            tbl_ = sh_.table
            rows_ = list(tbl_.rows)
            for ci in range(1, len(tbl_.columns)):
                vals = txts = 0
                for row_ in rows_[1:]:
                    t_ = row_.cells[ci].text.strip()
                    if not t_:
                        continue
                    if is_value(t_):
                        vals += 1
                    else:
                        txts += 1
                if vals and vals >= txts:
                    for row_ in rows_:
                        for para_ in row_.cells[ci].text_frame.paragraphs:
                            para_.alignment = PP_ALIGN.CENTER


    for (slide_no, shape_path, kind, _para), group in by_shape.items():
        slide = slides[slide_no - 1]
        rule = rules.get(slide_no, {})
        try:
            shp = _shape_at(slide, shape_path)
        except (IndexError, AttributeError):
            continue
        if kind == "table-cell":
            tbl = shp.table
            rect = _abs_rect(shp)
            if rect is None:
                continue
            x0, y0 = rect[0], rect[1]
            heights = _row_heights(tbl)
            # The heights in the file are MINIMUMS: PowerPoint grows any row
            # whose text needs more, the table renders taller than it
            # declares, and every padlock below a grown row sits low - by a
            # whole row once a few have accumulated.  Estimating the growth
            # was not enough on its own, so the estimate is also WRITTEN
            # BACK: a row set to the height its own content needs is a row
            # PowerPoint has no reason to grow, which makes the geometry
            # computed here the geometry it draws.
            for ri, hh in enumerate(heights):
                try:
                    tbl.rows[ri].height = Emu(int(hh * EMU_PER_PX))
                except Exception:
                    pass
            for c in group:
                r, col = c.ref
                cell = tbl.cell(r, col)
                size = _blank_cell(cell)
                cx = x0 + _col_x(tbl, col) + \
                    tbl.columns[col].width / EMU_PER_PX / 2
                cy = (y0 + _row_y(tbl, r, heights)
                      + heights[r] / 2)
                _place_lock(slide, art, cx, cy, size, c.unit)
                rep.locked += 1
        elif kind == "chart-label":
            chart = shp.chart
            # ``chart.series`` walks the plots in order, so the global series
            # index addresses this list exactly.  (Pairing them by identity
            # would not work: lxml hands out a fresh proxy object for the
            # same node on every lookup.)
            pairs = [(p, s) for p in chart.plots for s in p.series]
            # ONE padlock for the chart, with every withheld figure shown
            # blurred, once there are enough of them that a mark on each
            # would crowd out the picture.  Below that, a padlock per figure
            # still reads as a legend.
            waterfall = _is_waterfall(chart)
            ring = _is_ring(chart)
            want = rule.get("charts")
            if isinstance(want, dict):
                # a page whose charts differ - a line chart beside a column
                # chart - states each one separately
                ct = str(chart.chart_type)
                want = next((v for k, v in want.items() if k in ct),
                            want.get("*"))
            if waterfall:
                # A build-up reads as ONE story - what happens between two
                # totals - so it carries one mark, wherever it appears and
                # whatever the page around it does.  A padlock per step said
                # the same thing five times over.
                blur = True
            elif want == "per_value":
                blur = False
            elif want == "one_lock":
                blur = True
            elif ring and policy.blur_charts:
                # A ring is withheld as ONE thing, however few slices it has
                # - its shares blur and the hole takes the mark (see
                # _is_ring).  Left to the count, a three-slice ring took a
                # padlock per slice and nothing in the middle.
                blur = True
            else:
                blur = (policy.blur_charts
                        and len(group) >= policy.blur_min_labels)
            if want == "per_column_centre":
                # A stacked column carries ONE medium padlock at the centre
                # of the column; every segment label goes quiet.  The marks
                # are drawn as pictures over the plot, one per category.
                rect_ = _abs_rect(shp)
                cats = set()
                for c in group:
                    si, pi = c.ref
                    if si >= len(pairs):
                        continue
                    plot, series = pairs[si]
                    try:
                        point = series.points[pi]
                    except (IndexError, AttributeError):
                        continue
                    _ensure_series_dlbls(plot, series)
                    _hide_chart_label(point)
                    cats.add(pi)
                    rep.locked += 1
                if rect_ and cats:
                    n_cat = max(len(_chart_categories(chart)), max(cats) + 1)
                    bx, by, bw, bh = _plot_box(rect_, _chart_kind(chart))
                    slot = bw / max(n_cat, 1)
                    ph = max(18.0, min(34.0, slot * 0.34, bh * 0.3))
                    # ``lock_fy``: where the column's middle falls in the
                    # estimated plot box - a legend of several rows under the
                    # chart lifts the real plot above the estimate
                    fy = float(rule.get("lock_fy", 0.5))
                    for pi in sorted(cats):
                        _chart_lock(slide, art,
                                    (bx + pi * slot, by, slot, bh),
                                    fx=0.5, fy=fy,
                                    scale=ph / max(1.0, _chart_lock_pt(
                                        (bx + pi * slot, by, slot, bh))))
                continue
            if want == "per_column_top":
                # A stacked column carries ONE padlock, on top of the whole
                # column: the segment at the top of the stack takes the mark
                # (pinned to its own top edge and lifted clear), and every
                # segment below it goes quiet.  A mark per segment said the
                # same thing thirteen times over.
                rect_ = _abs_rect(shp)
                by_cat = {}
                for c in group:
                    by_cat.setdefault(c.ref[1], []).append(c)
                for pi, cs in by_cat.items():
                    top_c = max(cs, key=lambda c: c.ref[0])
                    for c in cs:
                        si = c.ref[0]
                        if si >= len(pairs):
                            continue
                        plot, series = pairs[si]
                        try:
                            point = series.points[pi]
                        except (IndexError, AttributeError):
                            continue
                        _ensure_series_dlbls(plot, series)
                        if c is top_c:
                            if _lock_chart_label(chart, point, art, c.unit):
                                if rect_:
                                    _lift_label(point,
                                                -_LABEL_LIFT_PX / rect_[3])
                                rep.locked += 1
                        else:
                            _hide_chart_label(point)
                            rep.locked += 1
                continue

            blurred = marked = False
            for c in group:
                si, pi = c.ref
                if si >= len(pairs):
                    continue
                plot, series = pairs[si]
                try:
                    point = series.points[pi]
                except (IndexError, AttributeError):
                    continue
                _ensure_series_dlbls(plot, series)
                done = False
                if waterfall:
                    # The steps say nothing at all: a blurred figure on a
                    # build-up adds a smudge to every bar without adding
                    # meaning, and the single padlock already carries it.
                    done = _hide_chart_label(point)
                    blurred |= done
                elif blur:
                    done = _blur_chart_label(chart, point, str(c.text),
                                             policy.blur_ink)
                    blurred |= done
                if not done:
                    done = _lock_chart_label(chart, point, art, c.unit)
                if (done and rule.get("labels") == "above"
                        and not _is_waterfall(chart)
                        and "STACKED" in str(chart.chart_type)):
                    # Only a STACKED plot needs this.  A clustered chart
                    # already puts its labels outside the bar; pinning those
                    # to the bar's top edge would drag them back ONTO it.
                    rect_ = _abs_rect(shp)
                    if rect_:
                        _lift_label(point, -_LABEL_LIFT_PX / rect_[3])
                if done:
                    rep.locked += 1
                    marked = True
            place = ("centre" if waterfall else rule.get("place", "centre"))
            if ring or place == "hole":
                # The ring's mark goes in the MIDDLE of the plate - its hole
                # where it has one, the centre of the pie where it does not.
                # Placed from the ring's own geometry, not the frame's: these
                # frames carry the legend beside or below the plot, so a mark
                # centred on the frame lands off the ring entirely
                # (_ring_geom).  A read-out already sitting in the hole is
                # re-sized and re-centred there rather than doubled.
                rect = _abs_rect(shp)
                if rect is not None and marked:
                    hole_locks.append((slide_no, rect, chart))
                continue
            if blurred:
                rect = _abs_rect(shp)
                if rect is not None:
                    kind_ = _chart_kind(chart)
                    box = _plot_box(rect, kind_)
                    dark = rule.get("shade") == "dark"
                    sc = float(rule.get("lock_scale", 1.0))
                    if place == "top_centre":
                        _chart_lock(slide, art, box, fx=0.5, fy=0.16,
                                    dark=dark, scale=sc)
                    elif place == "whitespace":
                        _chart_lock(slide, art, box, fx=0.82, fy=0.5,
                                    dark=dark, scale=sc)
                    else:
                        _chart_lock(slide, art, box, dark=dark, scale=sc)
        elif kind == "shape-value":
            rect = _abs_rect(shp)
            c = group[0]
            vi = c.ref[0]
            para = shp.text_frame.paragraphs[vi]
            if rule.get("shapes") in ("one_lock", "per_column")                     and rect is not None:
                # a chart drawn as rectangles: the figure blurs where it
                # stood and the page takes ONE padlock (placed after the
                # pass, once every such figure's position is known)
                pt_ = next((r.font.size.pt for r in para.runs if r.font.size),
                           9.0)
                ink = _run_ink(para)
                if _place_blur(slide, rect[0] + rect[2] / 2,
                               rect[1] + rect[3] / 2, pt_, c.text, ink):
                    for r in list(para.runs):
                        r.text = ""
                    shape_blurs.setdefault(slide_no, []).append(rect)
                    rep.locked += 1
                    continue
            size = next((r.font.size.pt for r in para.runs if r.font.size),
                        14.0)
            # A space, not nothing: emptying the line collapses it, the card
            # re-centres, and the locked card's heading no longer lines up
            # with its neighbours'.
            for j, r in enumerate(list(para.runs)):
                r.text = " " if j == 0 else ""
            if rect is not None:
                # The text block is anchored to the shape's MIDDLE, so the
                # value line is just below the shape's centre - not at
                # (vi+0.5)/n of its height.  On a tall treemap tile that old
                # guess dropped the padlock a third of the way down the tile,
                # nowhere near the country name it belongs to.
                heights = []
                for para_ in shp.text_frame.paragraphs:
                    pt_ = next((r.font.size.pt for r in para_.runs
                                if r.font.size), size)
                    heights.append(pt_ * LINE_LEAD * PT_TO_PX)
                block = sum(heights) or rect[3]
                top = rect[1] + rect[3] / 2 - block / 2
                cy = top + sum(heights[:vi]) + heights[vi] / 2
                # a page may size the padlock for particular shapes
                # ("card_lock_scale_by_name") and leave the rest as they are
                sc = float((rule.get("card_lock_scale_by_name") or {}).get(
                    getattr(shp, "name", ""),
                    rule.get("card_lock_scale", 1.0)))
                if sc == 1.0 and len(heights) > 1:
                    # The padlock is drawn a little taller than the type it
                    # replaces, which on a shape carrying ONLY a value is
                    # fine - but here the value sits above its own captions,
                    # and the overflow lands on the words beneath.  Hold it
                    # to its own line.
                    room = heights[vi] / (PT_TO_PX * LOCK_IMG_K)
                    size = min(size, room)
                if sc > 1.0:
                    # a padlock grown past the line it replaces must grow
                    # DOWNWARD - its top stays on the line, so it never
                    # reaches up into the name above it
                    h_ = size * sc * PT_TO_PX * LOCK_IMG_K
                    cy += (h_ - heights[vi]) / 2
                _place_lock(slide, art, rect[0] + rect[2] / 2, cy,
                            size * sc, c.unit)
            rep.locked += 1
        elif kind == "table-prose":
            tbl = shp.table
            rect = _abs_rect(shp)
            if rect is None:
                continue
            r, c = group[0].ref[0], group[0].ref[1]
            hide = {g.ref[2] for g in group}
            cell = tbl.cell(r, c)
            para = cell.text_frame.paragraphs[0]

            def swap_c(t, _hide=hide, _m=policy.prose_mark):
                marks = _prose_marks(t, policy)
                for i in range(len(marks) - 1, -1, -1):
                    if i not in _hide:
                        continue
                    a, b, _whole, unit = marks[i]
                    t = t[:a] + _m + unit + t[b:]
                return t

            whole = "".join(x.text for pp in cell.text_frame.paragraphs
                            for x in pp.runs)
            marked = swap_c(whole)
            if marked == whole:
                continue
            rep.locked += len(hide)
            if not (art.available and policy.prose_padlocks):
                _rewrite_paragraph(para, lambda _t, _n=marked: _n)
                continue
            pt, col, bold, font = _para_style(para, 9.5)
            cell_marks = _mark_colours(para)      # before the cell is cleared
            cw = tbl.columns[c].width / EMU_PER_PX
            cx0 = rect[0] + _col_x(tbl, c)
            heights = _row_heights(tbl)
            # Write them back, exactly as the table-cell branch does: these
            # sentences are redrawn at computed positions, so if PowerPoint
            # is left free to grow the rows the text drifts out of them and
            # the last rows' lines land past the end of the table.
            for ri, hh in enumerate(heights):
                try:
                    tbl.rows[ri].height = Emu(int(hh * EMU_PER_PX))
                except Exception:
                    pass
            cy0 = rect[1] + _row_y(tbl, r, heights)
            pad_l = _cell_pad(cell, "left", 7.0)
            avail = max(24.0, cw - pad_l - _cell_pad(cell, "right", 7.0))
            lines = _wrap_marked(marked, avail, pt, policy.prose_mark,
                                 font, bold)
            lp = pt * LINE_LEAD * PT_TO_PX
            for pp in cell.text_frame.paragraphs:
                for x in list(pp.runs):
                    x.text = ""
            ty = cy0 + max(2.0, (heights[r] - lp * len(lines)) / 2)
            for ln in lines:
                _draw_marked_line(slide, art, cx0 + pad_l, ty, avail, lp,
                                  ln, pt, col, policy.prose_mark, bold, font,
                                  cell_marks)
                ty += lp

        elif kind == "prose":
            hide = {c.ref[1] for c in group}
            pi = group[0].ref[0]
            para = shp.text_frame.paragraphs[pi]

            def swap(s, _hide=hide, _m=policy.prose_mark):
                marks = _prose_marks(s, policy)
                for i in range(len(marks) - 1, -1, -1):
                    if i not in _hide:
                        continue
                    a, b, _whole, unit = marks[i]
                    s = s[:a] + _m + unit + s[b:]
                return s

            # the ▲/▼ colours have to be read before the rewrite moves the
            # run boundaries (see _mark_colours)
            prose_marks.setdefault((slide_no, shape_path), {}) \
                .setdefault(pi, _mark_colours(para))
            if _rewrite_paragraph(para, swap):
                rep.locked += len(hide)
                if art.available and policy.prose_padlocks:
                    prose_shapes.add((slide_no, shape_path))

    for i, sl in enumerate(slides, start=1):
        r = rules.get(i, {})
        if not r.get("strip_values"):
            continue
        for path_, shp in _iter_shapes(sl.shapes):
            if _is_strip_panel(shp, r):
                rep.locked += _strip_panel(shp)
                # The panel is set to "resize shape to fit text", so
                # PowerPoint regrows it on open and the list runs past the
                # box it sits in.  Hand it to the same redraw that lays out
                # a locked sentence: that clears the frame and places each
                # line itself, shrinking the type a notch at a time until
                # the whole list fits.
                prose_shapes.add((i, path_))

    for slide_no, boxes in shape_blurs.items():
        r = rules.get(slide_no, {})
        sc = float(r.get("lock_scale", 1.0))
        slide = slides[slide_no - 1]
        if r.get("shapes") == "per_column":
            # One padlock per COLUMN - a region's box - centred on the run of
            # cells it withheld.  Three withheld rows put it in the middle
            # one; two put it on the line between them.
            cols = {}
            for b in boxes:
                cols.setdefault(round(b[0] / 12), []).append(b)
            for run in cols.values():
                run.sort(key=lambda b: b[1])
                # Split the column into runs of ADJACENT withheld cells.
                runs, cur = [], [run[0]]
                for b in run[1:]:
                    prev = cur[-1]
                    if b[1] - (prev[1] + prev[3]) <= prev[3] * 0.6:
                        cur.append(b)
                    else:
                        runs.append(cur); cur = [b]
                runs.append(cur)
                if len(runs) > 1:
                    # The readable row sits BETWEEN two withheld ones, so the
                    # column already reads as blurred / shown / blurred.  A
                    # padlock would have to land either on the figure the
                    # page meant to show or on one half of a split - and it
                    # would be saying what the blur has already said.
                    continue
                best = runs[0]
                x0 = min(b[0] for b in best)
                x1 = max(b[0] + b[2] for b in best)
                y0 = best[0][1]; y1 = best[-1][1] + best[-1][3]
                _chart_lock(slide, art, (x0, y0, x1 - x0, y1 - y0), scale=sc)
            continue
        # ONE padlock for the shape-drawn chart, in the middle of everything
        # it just blurred - which on a pair of stacked columns is the paper
        # between them.
        x0 = min(b[0] for b in boxes); x1 = max(b[0] + b[2] for b in boxes)
        y0 = min(b[1] for b in boxes); y1 = max(b[1] + b[3] for b in boxes)
        _chart_lock(slide, art, (x0, y0, x1 - x0, y1 - y0), scale=sc)

    for slide_no, rect, chart in hole_locks:
        # The ring's own centre read-out has already been withheld, so a
        # padlock is sitting in the hole - but sized to that little caption
        # and placed where the caption's BOX was, which is not where the ring
        # is.  Both are corrected here against the ring's real geometry.
        slide = slides[slide_no - 1]
        cx0, cy0, hole = _ring_geom(chart, rect)
        want_h = max(20.0, min(80.0, hole * _DONUT_LOCK_FILL))
        png = art.plain()
        found = None
        for _p, sh in _iter_shapes(slide.shapes):
            if sh.shape_type != 13:
                continue
            r = _abs_rect(sh)
            if r is None or r[2] > 60:
                continue
            if (rect[0] <= r[0] + r[2] / 2 <= rect[0] + rect[2]
                    and rect[1] <= r[1] + r[3] / 2 <= rect[1] + rect[3]):
                found = sh
                break
        if found is None:
            # No padlock in the hole - the ring's centre read-out was a bare
            # year, which is a caption, not a figure, so nothing withheld it.
            # Clear the caption and put the mark in the hole regardless.
            for _p2, sh2 in _iter_shapes(slide.shapes):
                if not getattr(sh2, "has_text_frame", False):
                    continue
                t2 = sh2.text_frame.text.strip()
                if not t2 or len(t2) > 14:
                    continue
                r2 = _abs_rect(sh2)
                if r2 is None or r2[2] > rect[2] * 0.35:
                    continue
                mx, my = r2[0] + r2[2] / 2, r2[1] + r2[3] / 2
                if (rect[0] <= mx <= rect[0] + rect[2]
                        and rect[1] <= my <= rect[1] + rect[3]):
                    for para2 in sh2.text_frame.paragraphs:
                        for run2 in para2.runs:
                            run2.text = ""
                    break
            if png is None:
                continue
            try:
                found = slide.shapes.add_picture(
                    str(png), Emu(0), Emu(0),
                    height=Emu(int(want_h * EMU_PER_PX)))
            except Exception:
                continue
        if png is not None:
            _centre_art(found, png, cx0, cy0, want_h)

    for slide_no, shape_path in sorted(prose_shapes):
        slide = slides[slide_no - 1]
        try:
            shp = _shape_at(slide, shape_path)
        except (IndexError, AttributeError):
            continue
        _redraw_prose_shape(slide, shp, art, policy,
                            prose_marks.get((slide_no, shape_path)))

    if out_path:
        prs.save(str(out_path))
    return rep


def lock_pptx(src, dst, policy: LockPolicy = None, assets_dir=None):
    """Lock a deck end to end: scan it, take the policy's proposal, save."""
    return apply_locks(src, dst, policy=policy, assets_dir=assets_dir)


# ═════════════════════════════════ CLI ═══════════════════════════════
def _main(argv=None):
    ap = argparse.ArgumentParser(
        description="Put the padlock on the values a client must not read.")
    ap.add_argument("src")
    ap.add_argument("dst", nargs="?")
    ap.add_argument("--open-years", default="",
                    help="comma-separated years that stay readable, "
                         "e.g. 2025,2026,2036 (default: guessed from the deck)")
    ap.add_argument("--lock-cagr", action="store_true",
                    help="withhold CAGR values too (default: they stay)")
    ap.add_argument("--totals-only", action="store_true",
                    help="in tables, keep the open years on the TOTAL row only")
    ap.add_argument("--no-prose", action="store_true",
                    help="leave figures inside sentences alone")
    ap.add_argument("--placeholder", default="xx")
    ap.add_argument("--protect", default="",
                    help="';'-separated phrases whose figures must survive - "
                         "above all the market's own name, e.g. "
                         "\"Prescription 5000-ppm Fluoride Toothpaste\"")
    ap.add_argument("--assets", default=None,
                    help="folder holding lock.png / lock_pct.png / lock_k.png")
    ap.add_argument("--dry-run", action="store_true",
                    help="list what would be withheld and stop")
    a = ap.parse_args(argv)

    policy = LockPolicy(
        open_years={int(y) for y in a.open_years.replace(" ", "").split(",")
                    if y},
        open_cagr=not a.lock_cagr,
        totals_only=a.totals_only,
        lock_prose=not a.no_prose,
        placeholder=a.placeholder,
        protect=tuple(p.strip() for p in a.protect.split(";") if p.strip()))

    if a.dry_run:
        cands = scan(a.src, policy)
        print(f"open years: {sorted(policy.open_years) or '(none)'}")
        print(f"{len(cands)} value(s) proposed for withholding\n")
        for c in cands:
            print(f"  s{c.slide:>3}  {c.kind:<12} {c.text:<14} {c.where[:64]}")
        return 0

    dst = a.dst or str(Path(a.src).with_name(
        Path(a.src).stem + " (locked).pptx"))
    rep = apply_locks(a.src, dst, policy=policy, assets_dir=a.assets)
    print(f"open years: {sorted(policy.open_years) or '(none)'}")
    print(rep)
    print(f"saved: {dst}")
    return 0


# ══════════════════════ the padlock, embedded ════════════════════════
# The artwork travels INSIDE this file so a single .py is all anyone needs.
# Drawn from LOCK_ART above and stored at 160 px tall - five times the size
# it is ever placed at, so it stays crisp - and unpacked to a cache
# directory the first time it is used (LockArt._unpack).  A project that
# ships its own PNGs never touches this: those win, and the output matches
# whatever that project already draws.
_EMBEDDED_ART = {
    "lock.png": """
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
TkSuQmCC""",
    "lock_label.png": """
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""",
    "lock_pct.png": """
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
RK5CYII=""",
    "lock_pct_white.png": """
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""",
    "lock_k.png": """
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
Pfx/FmypxIAMy2EAAAAASUVORK5CYII=""",
    "lock_k_white.png": """
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
cxgc9W8AAAAASUVORK5CYII=""",
}


if __name__ == "__main__":
    sys.exit(_main())
