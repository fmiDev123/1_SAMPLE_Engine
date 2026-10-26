"""Central configuration: file locations, years, PBIX theme, geo mappings.

Colours and canvas geometry are lifted straight out of the PBIX files
(`CY23SU08` theme + the 'New Country Wise' background SVGs) so the app
matches the Power BI report's look.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# ── Data file discovery ──────────────────────────────────────────────
_PROJ_ROOT = Path(__file__).resolve().parents[1]

def _find_data_file() -> Path:
    """Locate the dataset workbook.

    Priority: MARKET_DATA / MATCHA_DATA env var, then 'Dataset_*.xlsx',
    then ANY Excel workbook next to (or inside) the project folder -
    just drop a .xlsx in the folder and the report picks it up.

    MARKET_DATA may be an absolute path OR just a filename living in the
    project folder - the latter lets several Render services share one
    repo, each pinned to its own market via MARKET_DATA=<file name>.
    """
    for var in ("MARKET_DATA", "MATCHA_DATA"):
        env = os.environ.get(var)
        if not env:
            continue
        p = Path(env)
        if p.exists():                       # absolute or cwd-relative path
            return p
        cand = _PROJ_ROOT / env              # bare filename in the repo
        if cand.exists():
            return cand
        # Named a workbook that is not there.  Falling through to the glob
        # below would load a DIFFERENT market without saying so - which is how
        # a deck for the wrong study gets built and never questioned.
        print(f"WARNING: {var}={env!r} does not exist; falling back to "
              f"whatever workbook is in {_PROJ_ROOT}", file=sys.stderr)
    patterns = ("Dataset_*.xlsx", "Dataset_*.XLSX",   # FMI convention first
                "*.xlsx", "*.XLSX", "*.xlsm", "*.xls")
    for pattern in patterns:
        for base in (_PROJ_ROOT, _PROJ_ROOT.parent, Path.cwd()):
            hits = [p for p in sorted(base.glob(pattern))
                    if not p.name.startswith("~$")]      # skip Excel locks
            if hits:
                return hits[0]
    # No workbook anywhere. That is normal for the PORTAL, which uploads its
    # datasets at runtime and binds them per session - it never renders from
    # this import-time default. The standalone report (run.py) does need one
    # and will fail later, with a clear message, when it tries to use None.
    return None

DATA_FILE = _find_data_file()

# ── Report years ─────────────────────────────────────────────────────
# Defaults follow the FMI convention (5 historical + estimate + forecast
# horizon).  data.py calls set_years() with the years actually found in
# the workbook, so any year range works.
YEARS        = list(range(2021, 2037))
HIST_START   = 2021
HIST_END     = 2025          # estimate year (country overview cards)
EST_YEAR     = 2026          # first forecast year
MID_YEAR     = 2031
END_YEAR     = 2036          # forecast horizon
HIST_YEARS   = list(range(HIST_START, HIST_END + 1))
FORE_YEARS   = list(range(EST_YEAR, END_YEAR + 1))

def set_years(years: list[int]) -> None:
    """Re-derive every year constant from the workbook's year columns.

    Called by data.py at load time (before the template modules import
    these constants).  Keeps the FMI 5-year-history convention where the
    range allows it.
    """
    global HIST_START, HIST_END, EST_YEAR, MID_YEAR, END_YEAR
    years = sorted(int(y) for y in years)
    YEARS[:] = years                          # mutate in place - shared object
    HIST_START = years[0]
    END_YEAR   = years[-1]
    HIST_END   = years[4] if len(years) > 6 else years[max(len(years) // 3, 1) - 1]
    EST_YEAR   = HIST_END + 1 if HIST_END + 1 in years else years[min(5, len(years) - 1)]
    MID_YEAR   = EST_YEAR + 5 if EST_YEAR + 5 < END_YEAR else \
        years[(years.index(EST_YEAR) + len(years) - 1) // 2]
    HIST_YEARS[:] = [y for y in years if y <= HIST_END]
    FORE_YEARS[:] = [y for y in years if y >= EST_YEAR]

UNIT_LABELS  = {"Volume": "Tons", "Value": "US$ Mn"}

def plural(word: str, n: int | None = None) -> str:
    """English plural of a geography/segment noun.

    The deck used to write f"{child.lower()}s", which is right for "region"
    and wrong for every "country": a region-scoped deck read "across 2
    countrys".  ``n`` of 1 keeps the singular."""
    w = str(word)
    if n == 1:
        return w
    if w.endswith("y") and not w.endswith(("ay", "ey", "iy", "oy", "uy")):
        return w[:-1] + "ies"
    if w.endswith(("s", "x", "z", "ch", "sh")):
        return w + "es"
    return w + "s"


def set_unit_labels(volume: str | None = None, value: str | None = None) -> None:
    """Override the display units (mutates in place so importers see it).

    FMI workbooks state the units in the sheet title, e.g. 'Volume (Kg)'
    or 'Value (US$ Thousand)'; data.py detects and applies them here."""
    if volume:
        UNIT_LABELS["Volume"] = volume.strip()
    if value:
        UNIT_LABELS["Value"] = value.strip()


# A Value figure is quoted in a MAGNITUDE - 'US$ Mn' means the number 240 is
# 240 million dollars.  Price is derived as Value / Volume, so leaving that
# magnitude out makes every price too small by exactly this factor.
_MAGNITUDES = (
    (("bn", "billion", "billions"), 1e9),
    (("mn", "mm", "million", "millions"), 1e6),
    (("th", "thousand", "thousands", "k"), 1e3),
)


def _split_magnitude(label: str) -> tuple[float, str]:
    """Separate a unit label into its MAGNITUDE and the thing being counted.

    'US$ Mn' -> (1e6, 'US$'), 'Mn Doses' -> (1e6, 'Doses'), 'MT' -> (1, 'MT').
    The magnitude word can sit on either side of the noun, which is why both
    'US$ Mn' and 'Mn Doses' have to be understood.
    """
    scale, rest = 1.0, []
    for word in str(label).split():
        w = word.strip("().,").casefold()
        hit = next((m for names, m in _MAGNITUDES if w in names), None)
        if hit is not None and scale == 1.0:
            scale = hit
        else:
            rest.append(word.strip("(),"))
    return scale, " ".join(rest).strip()


def value_scale() -> float:
    """How many currency units one Value figure stands for.

    'US$ Mn' -> 1e6, 'US$ Thousand' -> 1e3, a bare 'US$' -> 1.  Read off the
    label the workbook itself states, so it follows the dataset rather than
    assuming millions.
    """
    return _split_magnitude(UNIT_LABELS.get("Value", ""))[0]


def volume_scale() -> float:
    """How many physical units one Volume figure stands for.

    Volume labels carry a magnitude too - the Humira workbook is sized in
    'Mn Doses'.  Ignoring it made a price per DOSE come out as a price per
    million doses, i.e. wrong by 1e6 in the opposite direction to the Value
    side, so both have to be read for the quotient to mean anything.
    """
    return _split_magnitude(UNIT_LABELS.get("Volume", ""))[0]


def price_scale() -> float:
    """Multiplier turning raw Value/Volume into currency per physical unit."""
    vs = volume_scale()
    return value_scale() / vs if vs else value_scale()


def currency() -> str:
    """The currency out of the Value label - 'US$' from 'US$ Mn'."""
    noun = _split_magnitude(UNIT_LABELS.get("Value", ""))[1]
    return noun or "US$"


def volume_noun() -> str:
    """What the Volume figures count, without the magnitude - 'Doses', 'MT'."""
    return _split_magnitude(UNIT_LABELS.get("Volume", ""))[1]


def price_unit() -> str:
    """The unit a DERIVED price is quoted in - 'US$/MT', 'US$/Doses'.

    The deck used to hardcode 'US$/unit', which said nothing about what a unit
    was and hid the magnitude bug: the number beside it was Value/Volume with
    the 'Mn' left in, so a market at US$ 3,840 per MT printed '0.00 US$/unit'.
    Naming the volume unit makes the figure checkable against the workbook.
    """
    vol = volume_noun()
    return f"{currency()}/{vol}" if vol else f"{currency()}/unit"


def abbrev_floor() -> float:
    """Smallest magnitude a chart number may be abbreviated (K/M/B) at.

    A unit label that already states a magnitude states it ONCE.  Abbreviating
    the number underneath it states a second one: a market of 50,531 US$ Mn
    drawn as '50.5K' reads as thousands beside a caption that says millions,
    and every exhibit picks its own tier, so one deck showed '32K' and '5,520'
    in the same six-item legend.  The same shorthand is what turns a real bar
    into a zero - a 15.5 US$ Mn step in a waterfall whose peak is 15,797
    rounds to '0.0K' and the reader sees nothing.

    So: where the workbook's own label carries the magnitude, print the
    numbers in that unit verbatim.  Where it does not ('MT', 'Units', 'Tons')
    the K/M/B shorthand adds the only magnitude on show and stays truthful, so
    it is kept - but never at the thousands tier, which is the one that both
    collides with 'Mn'/'Thousand' labels and rounds small bars away.

    inf  - every stated unit carries its own magnitude; never abbreviate.
    1e6  - some do (e.g. Value 'US$ Mn' with Volume 'MT'); M/B only.
    1e3  - none do; the original K/M/B behaviour.
    """
    stated = [l for l in (UNIT_LABELS.get("Value"), UNIT_LABELS.get("Volume"))
              if str(l or "").strip()]
    if not stated:
        return 1e3
    with_mag = sum(1 for l in stated if _split_magnitude(l)[0] > 1.0)
    if with_mag == len(stated):
        return float("inf")
    return 1e6 if with_mag else 1e3

# ═══════════════════════════════════════════════════════════════════════
#  VISUAL DESIGN SYSTEM - re-exported from the ACTIVE template's theme
# ═══════════════════════════════════════════════════════════════════════
# Canvas geometry, the brand palette, typography, the spacing scale and the
# shared component-metric helpers live in ONE swappable module per template:
#
#     theme_factmr.py   Fact.MR   (the default - the report as originally shipped)
#     theme_fmi.py      FMI       (derived from the FMI B2B sample template)
#
# Both define the SAME token names, so nothing downstream branches on which is
# active.  This file keeps only the concerns that are not about how the report
# LOOKS: dataset discovery, report years, unit labels and the geo map.
#
# The choice below is only the START-UP default (override with REPORT_TEMPLATE=
# factmr|fmi).  At runtime the dashboard switch and the PowerPoint export both
# go through `theming.apply()` / `theming.use()`, which rebind these names here
# and in every renderer.  This is the ONE place a template name is tested.
TEMPLATE = (os.environ.get("REPORT_TEMPLATE") or "factmr").strip().lower()
if TEMPLATE in ("fmi", "future market insights"):
    TEMPLATE = "fmi"
    from .theme_fmi import *                                      # noqa: F401,F403
else:
    TEMPLATE = "factmr"
    from .theme_factmr import *                                   # noqa: F401,F403

# ── Map support: report geographies -> ISO-3 country codes ───────────
GEO_ISO3 = {
    "USA": ["USA"], "Canada": ["CAN"], "Brazil": ["BRA"],
    "Germany": ["DEU"], "United Kingdom": ["GBR"], "France": ["FRA"],
    "Italy": ["ITA"], "Spain": ["ESP"],
    "Benelux": ["BEL", "NLD", "LUX"],
    "Nordics": ["SWE", "NOR", "DNK", "FIN", "ISL"],
    "China": ["CHN"], "Japan": ["JPN"], "South Korea": ["KOR"],
    "India": ["IND"],
    "ASEAN Countries": ["IDN", "THA", "VNM", "PHL", "MYS", "SGP",
                        "MMR", "KHM", "LAO", "BRN"],
    "Australia & New Zealand": ["AUS", "NZL"],
    "Turkiye": ["TUR"], "South Africa": ["ZAF"],
    # generic names seen in other FMI workbooks
    "U.K.": ["GBR"], "UK": ["GBR"], "U.S.": ["USA"], "US": ["USA"],
    "United States": ["USA"], "Mexico": ["MEX"], "Argentina": ["ARG"],
    "Poland": ["POL"], "Russia": ["RUS"], "Switzerland": ["CHE"],
    "Netherlands": ["NLD"], "Belgium": ["BEL"], "Sweden": ["SWE"],
    "Norway": ["NOR"], "Denmark": ["DNK"], "Finland": ["FIN"],
    "Denmark & Finland": ["DNK", "FIN"], "Austria": ["AUT"],
    "Portugal": ["PRT"], "Greece": ["GRC"], "Ireland": ["IRL"],
    "Czech Republic": ["CZE"], "Hungary": ["HUN"], "Romania": ["ROU"],
    "Rest of Europe": ["AUT", "PRT", "GRC", "IRL", "CZE", "HUN", "ROU",
                       "BGR", "HRV", "SVK", "SVN", "LTU", "LVA", "EST"],
    "Indonesia": ["IDN"], "Thailand": ["THA"], "Vietnam": ["VNM"],
    "Malaysia": ["MYS"], "Singapore": ["SGP"], "Philippines": ["PHL"],
    "Australia": ["AUS"], "New Zealand": ["NZL"], "Egypt": ["EGY"],
    "Saudi Arabia": ["SAU"], "UAE": ["ARE"], "Israel": ["ISR"],
    "GCC Countries": ["SAU", "ARE", "QAT", "KWT", "OMN", "BHR"],
    "Nigeria": ["NGA"], "Chile": ["CHL"], "Colombia": ["COL"],
    "Peru": ["PER"], "Turkey": ["TUR"],
    # remaining EU / European geographies (e.g. the Decorations workbook)
    "Luxembourg": ["LUX"], "Malta": ["MLT"], "Slovakia": ["SVK"],
    "Slovenia": ["SVN"], "Croatia": ["HRV"], "Bulgaria": ["BGR"],
    "Lithuania": ["LTU"], "Latvia": ["LVA"], "Estonia": ["EST"],
    # 'Rest of ...' buckets and singular spellings.  A geography with no
    # entry here is silently dropped from the bubble maps, so a Top-10 map
    # quietly plots 8 bubbles - these keep the map and its table in step.
    "Nordic": ["SWE", "NOR", "DNK", "FIN", "ISL"],
    "Rest of LATAM": ["ARG", "COL", "PER", "CHL", "ECU", "BOL", "URY",
                      "PRY", "VEN"],
    "Rest of Latin America": ["ARG", "COL", "PER", "CHL", "ECU", "BOL",
                              "URY", "PRY", "VEN"],
    "Rest of Southeast Asia": ["VNM", "MMR", "KHM", "LAO", "BRN"],
    "Rest of South Asia & Pacific": ["PAK", "BGD", "LKA", "NPL", "PNG"],
    "Rest of GCC": ["QAT", "KWT", "OMN", "BHR"],
    "Other GCC Countries": ["QAT", "KWT", "OMN", "BHR"],
    "Rest of MEA": ["NGA", "EGY", "KEN", "MAR", "DZA", "GHA", "ETH", "TZA"],
    "Rest of Middle East & Africa": ["NGA", "EGY", "KEN", "MAR", "DZA",
                                     "GHA", "ETH", "TZA"],
    "Rest of North America": ["MEX"],
    "Rest of East Asia": ["TWN", "HKG", "MNG"],
    "Cyprus": ["CYP"], "Iceland": ["ISL"], "Eastern Europe": ["POL",
        "CZE", "ROU", "GRC", "HUN", "SVK", "HRV", "SVN", "BGR", "LTU",
        "LVA", "EST", "CYP"],
    "Western Europe": ["DEU", "FRA", "ITA", "ESP", "NLD", "BEL", "AUT",
        "SWE", "DNK", "IRL", "PRT", "FIN", "LUX", "MLT"],
}
