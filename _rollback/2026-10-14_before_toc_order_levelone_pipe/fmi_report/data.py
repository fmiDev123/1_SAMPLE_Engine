"""Load ANY FMI market workbook into a tidy in-memory model.

Two workbook layouts are auto-detected:

1. **Master Sheet** format (e.g. ``Dataset_Global Matcha Market_FMI.XLSX``)
   - a single "Master Sheet" tab:

       Region | Segment | Sub-segment | Unit | 2021 ... 2036

   ``Region`` holds Global, the world regions AND the countries;
   ``Segment`` holds the product segments plus the two hierarchy
   segments ``Region`` (children of Global) and ``Country`` (children
   of a region).

2. **Per-geography sheets** format (e.g. ``Tortilla Market.xlsx``)
   - one tab per geography, each holding a Volume section, a Price
   section and a Value section of segment blocks.  The parser converts
   every sheet into the same tidy model above, detects the aggregate
   sheet (the one that equals the sum of the others) and synthesises
   the Global → region → country hierarchy from it.

Everything downstream (segments, years, hierarchy) is derived from the
data, so dropping a different workbook into the project folder is all
that is needed.
"""
from __future__ import annotations

import os
import re
from functools import lru_cache

import numpy as np
import pandas as pd

from . import config
from .config import DATA_FILE, YEARS

_TIDY_COLS = ["Region", "Segment", "Sub-segment", "Unit"]
# "Sub-region" is a geography LEVEL, not a way of cutting the market.  Left
# out of this tuple it was offered as a product segment - and since
# `total_series` sums SEGMENTS[0], a Europe study whose first segment was
# "Sub-region" reported zero for every country beneath it.
_HIERARCHY_SEGS = ("Region", "Sub-region", "Country")

# Set when a super-region aggregate (e.g. 'Europe') is relabeled to
# 'Global' during load - kept so the market-name detector can still strip
# that leading geography from page titles.
_RELABELED_ROOT: str | None = None


def _title_cell(sheet: str) -> str:
    """Row-0 / col-0 title of a sheet (used to detect market & units).
    Returns '' if the sheet or cell is missing."""
    for eng in ("calamine", None):
        try:
            kw = {"engine": eng} if eng else {}
            cell = pd.read_excel(DATA_FILE, sheet_name=sheet, header=None,
                                 nrows=1, usecols=[0], **kw).iloc[0, 0]
            return str(cell) if cell is not None else ""
        except Exception:
            continue
    return ""


def _year_like(v) -> int | None:
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if np.isnan(f) or f != int(f):
        return None
    return int(f) if 1990 <= f <= 2100 else None


def _first_year_run(df: pd.DataFrame) -> pd.DataFrame:
    """Keep only the FIRST contiguous block of year columns.

    A Master Sheet does not stop at its size series.  Workbooks routinely
    carry analysis blocks to its right - H-CAGR / F-CAGR, then a SECOND run
    of the same years holding YoY growth rates.  Those repeat the year
    headers, and because pandas de-duplicates them ('2022', '2022.1') while
    this loader converts both back to the int 2022, the two blocks collide:
    `df[years]` then returns TWO columns per year, `total_series` comes back
    with 241 entries for 16 years, and `g[2026]` is a Series where every
    caller expects a number.  Downstream that surfaces as pages which are
    silently blank - every builder on the Market Signals page raised
    TypeError inside its own `except: pass`.

    The size series is always the first run, so that is what is kept; every
    later year column is renamed out of the way (never dropped, so a reader
    of the frame can still see it was there).  A sheet with one clean run is
    returned untouched.
    """
    seen, run_over, renamed = set(), False, {}
    for pos, col in enumerate(df.columns):
        if not isinstance(col, int):
            if seen:
                run_over = True          # the run has ended; years after it
            continue                     # belong to another block
        if run_over or col in seen:
            renamed[pos] = f"_extra_{col}_{pos}"
        else:
            seen.add(col)
    if not renamed:
        return df
    cols = list(df.columns)
    for pos, name in renamed.items():
        cols[pos] = name
    df.columns = cols
    return df.loc[:, [c for c in df.columns
                      if not str(c).startswith("_extra_")]]


# ── Format 1: Master Sheet ───────────────────────────────────────────
def _apply_sheet_unit_labels(*sheets: str) -> None:
    """Set the Volume/Value display units to exactly what the datasheet
    states, e.g. a title 'Value (US$ Th)' / 'Volume (Units Th)' -> those
    strings verbatim.  Scans each sheet's title cell for the first
    '(Volume|Value) (<unit>)' it finds and stops once both are known.

    Records WHICH roles the title supplied in `_TITLE_UNITS`, because
    `config.UNIT_LABELS` ships with defaults ('Tons', 'US$ Mn') and so is never
    empty - without this, a label read off the Unit column could not tell a
    real title unit from an untouched default, and would either overwrite the
    title or never fill a gap."""
    global _TITLE_UNITS
    found: dict[str, str] = {}
    for sheet in sheets:
        for role, unit in re.findall(r"(Volume|Value)\s*\(([^)]+)\)",
                                     _title_cell(sheet), flags=re.I):
            found.setdefault(role.capitalize(), unit.strip())
    _TITLE_UNITS = {k for k, v in found.items() if v}
    if found.get("Volume") or found.get("Value"):
        config.set_unit_labels(volume=found.get("Volume"),
                               value=found.get("Value"))


# The Unit column names the MEASURE, and the whole report looks a measure up by
# that exact word - `data.segment_table(geo, "Volume", seg)`, `has_unit`,
# `has_price`.  Some workbooks put the measurement unit in the column as well:
# 'Volume (MT)' instead of 'Volume'.  That is the same fact the sheet title
# already carries, and it is only ever written on the volume side, so a
# workbook can look completely healthy while every volume lookup in the deck
# quietly misses.
_UNIT_WITH_LABEL_RE = re.compile(r"^\s*(volume|value)\s*[\(\[]([^\)\]]*)[\)\]]\s*$",
                                 re.IGNORECASE)

# Which unit labels the SHEET TITLE stated, set by `_apply_sheet_unit_labels`.
_TITLE_UNITS: set[str] = set()


def _normalise_units(col: pd.Series) -> pd.Series:
    """Reduce the Unit column to the bare measure, keeping what it said.

    'Volume (MT)' -> 'Volume', and 'MT' becomes the display label unless the
    sheet title already gave one (the title is the better source: it is where
    FMI workbooks state units, and it survives an analyst tidying the column).

    A unit that is already bare, or that is neither Volume nor Value, is passed
    through untouched - so this cannot rename a measure the report does not
    know, it can only strip a decoration off one it does.
    """
    seen: dict[str, str] = {}

    def one(u):
        m = _UNIT_WITH_LABEL_RE.match(str(u))
        if not m:
            return str(u)
        measure = m.group(1).capitalize()
        label = m.group(2).strip()
        if label:
            seen.setdefault(measure, label)
        return measure

    out = col.map(one)
    # Only fill a label the title did not supply; never overwrite one it did.
    missing = {k: v for k, v in seen.items() if k not in _TITLE_UNITS}
    if missing:
        config.set_unit_labels(volume=missing.get("Volume"),
                               value=missing.get("Value"))
    return out


def _self_referential_segment(df: pd.DataFrame) -> str | None:
    """The name of a segment whose every row is a geography's own grand total
    (``Region == Sub-segment``), or None.

    Some workbooks state each country's total that way - `Region='U.S.',
    Segment='Country Total', Sub-segment='U.S.'` - instead of parenting the
    country under its region the way `Segment='Country'` rows do.  The test is
    STRUCTURAL, not a name match, so a workbook that calls it 'Country Total',
    'Total' or anything else is still recognised."""
    for seg in df["Segment"].drop_duplicates():
        if seg in _HIERARCHY_SEGS:
            continue
        rows = df[df["Segment"] == seg]
        if len(rows) and (rows["Region"] == rows["Sub-segment"]).all():
            return str(seg)
    return None


def _reparent_selfref_geographies(df: pd.DataFrame) -> pd.DataFrame:
    """Re-point hierarchy rows that name THEMSELVES at their actual parent.

    Two workbook dialects state the Global → region → country tree, and they
    disagree about what the ``Region`` column means on a hierarchy row:

    * the parented dialect puts the PARENT there -
      ``Region='Global', Segment='Region', Sub-segment='North America'`` and
      ``Region='North America', Segment='Country', Sub-segment='U.S.'``.
      The edge is stated outright, and everything downstream reads it;
    * the self-referential dialect repeats the geography -
      ``Region='North America', Segment='Region', Sub-segment='North America'``
      and ``Region='U.S.', Segment='Country', Sub-segment='U.S.'``.  Each row
      declares what LEVEL a geography sits at, but never whose child it is.

    The second dialect used to slip through unnoticed, and that was the whole
    failure: `_region_hierarchy` and `_country_hierarchy` both open by asking
    whether ``Segment == 'Region'`` / ``'Country'`` rows exist at all, find
    that they do, and return the frame untouched as "already stated" - while
    the query those levels are actually read with, ``Region == 'Global' and
    Segment == 'Region'``, matches nothing, because no such row names Global
    as the parent.  REGIONS and COUNTRIES both came out empty on a workbook
    that had every geography in it, and the deck degraded to one bar per
    regional chart, "Global is the largest region, at 100.0%", "Top 0
    Countries" and a country section that repeated Global.

    The missing edge is recoverable, because these sheets are laid out
    depth-first - a region, then its own countries, then the next region - so
    a country belongs to the region most recently declared above it.  That
    order is read off the hierarchy rows themselves, which is where the
    workbook states it; nothing here matches a list of known region names, so
    it carries to any market.

    Only the ``Region`` column is rewritten.  The row's own figures are the
    geography's stated totals and are kept exactly as the workbook gives them,
    so this recovers the tree without inventing a single number.

    A workbook of the first dialect is returned untouched: the test demands
    that EVERY hierarchy row be self-referential, so one genuine parent link
    is enough to leave the frame alone.  A sheet whose first geography is a
    country - no region open yet to hang it on - is also returned untouched,
    since a wrong hierarchy is worse than none.
    """
    hier = df[df["Segment"].isin(_HIERARCHY_SEGS)]
    if hier.empty:
        return df
    if not (hier["Region"].astype(str) == hier["Sub-segment"].astype(str)).all():
        return df                                # already parented properly

    # what level each geography was declared at, in the sheet's own order
    level: dict[str, str] = {}
    for geo, seg in zip(hier["Sub-segment"].astype(str),
                        hier["Segment"].astype(str)):
        level.setdefault(geo, seg)

    parent: dict[str, str] = {}
    region = subregion = None
    for geo, seg in level.items():
        if seg == "Region":
            parent[geo], region, subregion = "Global", geo, None
        elif seg == "Sub-region":
            if region is None:
                return df                        # a sub-region with no region
            parent[geo], subregion = region, geo
        else:                                    # Country
            if region is None:
                return df                        # a country before any region
            parent[geo] = subregion or region

    df = df.copy()
    rows = df["Segment"].isin(_HIERARCHY_SEGS)
    df.loc[rows, "Region"] = df.loc[rows, "Sub-segment"].map(parent)
    return df.dropna(subset=["Region"])


def _hierarchy_reconciles(df: pd.DataFrame, rows: list, years: list[int],
                          tol: float = 0.02) -> bool:
    """Do the countries just parented actually ADD UP to their region?

    The parent is read off sheet ORDER, which is right for a sheet laid out
    region → its countries → next region.  A sheet that listed every region
    first and every country afterwards would parent the lot to the last
    region - silently, and a wrong hierarchy is worse than none.  So the
    result is checked against the region totals the workbook already states
    (the `Segment == 'Region'` rows under Global) before it is trusted; if it
    does not reconcile, the loader leaves the model exactly as it was."""
    stated: dict[tuple, float] = {}
    for _, r in df[(df["Region"] == "Global")
                   & (df["Segment"] == "Region")].iterrows():
        stated[(r["Sub-segment"], r["Unit"])] = sum(float(r[y]) for y in years)
    built: dict[tuple, float] = {}
    for row in rows:                             # [parent, 'Country', sub, unit, *vals]
        key = (row[0], row[3])
        built[key] = built.get(key, 0.0) + sum(float(v) for v in row[4:])
    checked = 0
    for key, total in stated.items():
        if key not in built or not total:
            continue
        checked += 1
        if abs(built[key] - total) > abs(total) * tol:
            return False
    return checked > 0


def _country_hierarchy(df: pd.DataFrame, years: list[int]) -> pd.DataFrame:
    """Give a workbook a Region→Country hierarchy when it states country
    totals but never parents them.

    Without `Segment == 'Country'` rows, `COUNTRIES` comes out empty,
    `ALL_COUNTRIES` falls back to standing the REGIONS in for countries, and
    every country-level exhibit is wrong or blank: 'Top 10 Countries' lists
    regions, and the country indices (which need `region_of(country)`) have
    nothing to rank.  The same rows are meanwhile promoted to a product
    SEGMENT, so the registry builds a full page block for a segment whose only
    sub-segment is the geography itself.

    Both problems are the same missing link, so both are fixed by adding it.

    The parent of each country is read off the Master Sheet's own ROW ORDER -
    the sheet lists a region and then its countries, then the next region -
    which is the only place the workbook states the relationship.  Nothing is
    inferred from the names, so it carries to any market.

    The country's TOTAL is taken from a self-referential segment when the
    workbook states one, and otherwise SUMMED from the country's own product
    rows - the same way the Global row is synthesised above.  Depending on a
    'Country Total' line would make the country level vanish the moment an
    analyst tidied those rows away, even though every country's data is still
    right there in the sheet."""
    if (df["Segment"] == "Country").any():
        return df                                # already parented properly
    regions = list(df.loc[(df["Region"] == "Global")
                          & (df["Segment"] == "Region"),
                          "Sub-segment"].drop_duplicates())
    if not regions:
        return df                                # no region level to hang off

    # walk the sheet in order: every geography that is not a region and not
    # the root belongs to the region most recently seen above it
    parent: dict[str, str] = {}
    current: str | None = None
    for geo in df["Region"]:
        if geo in regions:
            current = geo
        elif geo != "Global" and current is not None:
            parent.setdefault(geo, current)
    if not parent:
        return df                                # nothing below the regions

    seg = _self_referential_segment(df) or _total_marker_segment(df, years)
    # The stated total (when there is one) is preferred per (geography, unit),
    # but it may not cover every unit - a workbook can state 'Market Total'
    # in Value only while its Volume lives in the ordinary segment rows.  A
    # unit the total rows skip is derived from the first segment that has
    # rows in it, the same way a workbook with no stated total is handled -
    # otherwise that unit simply has no country level at all.
    rows = []
    body = df[~df["Segment"].isin(_HIERARCHY_SEGS)]
    for (geo, unit), grp in body[body["Region"].isin(parent)].groupby(
            ["Region", "Unit"], sort=False):
        stated = grp[grp["Segment"] == seg] if seg is not None else grp.iloc[:0]
        if len(stated):                          # the workbook states the total
            vals = stated[years].sum()
        else:
            # every complete segment sums to the same geography total; take
            # the first one this geography actually has rows in
            first = grp["Segment"].iloc[0]
            vals = grp[grp["Segment"] == first][years].sum()
        rows.append([parent[geo], "Country", geo, unit,
                     *vals.tolist()])
    if not rows:
        return df
    if not _LAYOUT_HIERARCHY and not _hierarchy_reconciles(df, rows, years):
        return df                                # order rule does not hold here

    if seg is not None:
        # the self-referential rows have become the hierarchy, so drop them as
        # a product segment - otherwise the registry pages them a second time
        df = df[df["Segment"] != seg]
    return pd.concat([df, pd.DataFrame(rows, columns=_TIDY_COLS + years)],
                     ignore_index=True)


def _is_facet(label: str) -> bool:
    """True for a level-3/4 label that names a CUT rather than a part.

    Fact.MR workbooks break one sub-segment down several independent ways at
    once - 'By Additive', 'By Closure', 'By Material' - and each block
    re-spends the SAME money.  They are alternative partitions, not siblings,
    so only one of them may be counted.  A part ('Microcollection Tubes',
    '20 mg') is counted alongside its siblings as usual.
    """
    return str(label).strip().lower().startswith("by ")


def _collapse_deeper(df: pd.DataFrame, years: list[int]) -> pd.DataFrame:
    """Aggregate the deeper taxonomy columns onto 'Sub-segment'.

    A plain sum where those columns hold parts; where they hold facets, one
    facet stands for all of them.  Summing the facets multiplies the market by
    the number of cuts - about 3x in the Capillary Blood workbook, whose own
    'Total Market' rows the corrected figures reconcile against exactly.

    A workbook with no facet labels takes the original groupby path, so
    markets that never had this shape are untouched.
    """
    deeper = [c for c in ("Sub-segment 1", "Sub-segment 2") if c in df.columns]
    plain = df.groupby(_TIDY_COLS, sort=False)[years].sum().reset_index()
    if not deeper:
        return plain
    if not any(df[c].map(_is_facet).any() for c in deeper):
        return plain

    def _walk(rows, cols):
        if not cols:
            return rows[years].sum()
        col = cols[0]
        vals = list(dict.fromkeys(rows[col].tolist()))     # first-seen order
        facets = [v for v in vals if _is_facet(v)]
        parts = [v for v in vals if not _is_facet(v)]
        if facets and not parts:
            return _walk(rows[rows[col] == facets[0]], cols[1:])
        total = None
        for v in parts:                       # a stray facet beside real parts
            sub = _walk(rows[rows[col] == v], cols[1:])    # is dropped, never
            total = sub if total is None else total + sub  # added on top
        return total

    out = []
    for key, g in df.groupby(_TIDY_COLS, sort=False):
        vals = _walk(g, deeper)
        out.append(list(key) + [float(vals[y]) for y in years])
    return pd.DataFrame(out, columns=_TIDY_COLS + list(years))


def _total_marker_segment(df: pd.DataFrame, years: list[int]) -> str | None:
    """The name of a segment that is really each geography's grand total.

    `_self_referential_segment` catches the form that repeats the geography
    ('Region=U.S., Sub-segment=U.S.').  Fact.MR workbooks write the same idea
    with a constant label instead - every row 'Total Market' - which is not
    self-referential, and which would otherwise be paged as a product segment
    whose one sub-segment carries no breakdown at all.

    The test stays structural: ONE distinct sub-segment, and a per-geography
    total that matches what the real segments already say.  A genuine
    single-sub-segment product segment does not add up to the whole market, so
    it is not mistaken for one.
    """
    body = df[~df["Segment"].isin(_HIERARCHY_SEGS)]
    cand = [g for g in body["Segment"].drop_duplicates()
            if body[body["Segment"] == g]["Sub-segment"].nunique() == 1]
    for seg in cand:
        others = body[body["Segment"] != seg]
        if others.empty:
            continue
        checked = agreed = 0
        for (geo, unit), grp in body[body["Segment"] == seg].groupby(
                ["Region", "Unit"], sort=False):
            peer = others[(others["Region"] == geo) & (others["Unit"] == unit)]
            if peer.empty:
                continue
            first = peer["Segment"].iloc[0]
            ref = float(peer[peer["Segment"] == first][years].sum().sum())
            got = float(grp[years].sum().sum())
            if not ref:
                continue
            checked += 1
            if abs(got - ref) <= abs(ref) * 0.02:
                agreed += 1
        if checked and agreed >= checked * 0.9:
            return str(seg)
    # No candidate reconciled arithmetically.  A segment the sheet itself
    # LABELS as the market total ('Market Total' / 'Total Market'), with a
    # single constant sub-segment, is still that - the label states the
    # layout the way a 'Rest of <region>' line states the hierarchy.  Trusting
    # it matters exactly when the numbers do NOT agree (a reference-format
    # workbook with placeholder figures): the alternative is paging the total
    # as a product segment whose one sub-segment carries no breakdown, and
    # summing an EMPTY table for any unit the total rows do not cover.
    for seg in cand:
        if re.fullmatch(r"(?:market\s+total|total\s+market)",
                        str(seg).strip(), flags=re.I):
            return str(seg)
    return None


# True once a hierarchy has been recovered from the sheet's LAYOUT rather than
# from its arithmetic - i.e. the workbook's geographies are known NOT to add
# up.  The country step reads it: re-testing the same broken sums one level
# down would throw away the hierarchy the sheet plainly states.
_LAYOUT_HIERARCHY = False


def _region_blocks_by_sums(order, totals):
    """Region -> its countries, decided by ARITHMETIC: walking the sheet's own
    order, a candidate parent closes as soon as the run beneath it adds up to
    it, and the geography after that starts a new one.  None when the order
    rule does not hold."""
    groups, i = {}, 0
    while i < len(order):
        parent = order[i]
        target = float(totals[parent].sum())
        kids, run, j = [], 0.0, i + 1
        while j < len(order):
            run += float(totals[order[j]].sum())
            kids.append(order[j])
            if target and abs(run - target) <= abs(target) * 0.02:
                break
            if target and run > target * 1.02:
                kids = []
                break
            j += 1
        if not kids or j >= len(order):
            return None
        groups[parent] = kids
        i = j + 1
    return groups or None


def _region_blocks_by_layout(order):
    """Region -> its countries, read off the sheet's LAYOUT instead of its
    arithmetic.

    Arithmetic is preferred and stays first: it cannot invent a hierarchy the
    numbers do not support.  But a workbook whose geographies do NOT add up is
    then left with no region level at all, and the deck degrades to "Global is
    the largest region, at 100.0%", "Top 0 Countries" and a country section
    that renders Global over and over - which is worse than a hierarchy the
    sheet plainly states in another way.

    The layout states it.  These sheets list a region and then its countries,
    and close most blocks with a "Rest of <region>" line.  Those lines are the
    anchors: each names its own region, and where it sits fixes that block's
    last row.  Whatever is left between two anchored blocks is one more block,
    headed by its own first geography - which is how a region with no "Rest of"
    line (North America over U.S. and Canada) is still found.

    Nothing here reads a list of known region names, so it carries to any
    market.  Returns None when the layout does not support a hierarchy either.
    """
    pos = {g: i for i, g in enumerate(order)}
    spans = []
    for i, g in enumerate(order):
        head = str(g).strip()
        if not head.lower().startswith("rest of "):
            continue
        j = pos.get(head[len("rest of "):].strip())
        if j is not None and j < i:
            spans.append((j, i))
    if not spans:
        return None                              # no anchor to divide on
    spans.sort()
    last = -1
    for a, b in spans:                           # anchors must not overlap
        if a <= last:
            return None
        last = b
    blocks, cur = [], 0
    for a, b in spans:
        if a > cur:                              # an un-anchored block before it
            blocks.append((cur, a - 1))
        blocks.append((a, b))
        cur = b + 1
    if cur < len(order):
        blocks.append((cur, len(order) - 1))
    groups = {order[a]: order[a + 1:b + 1] for a, b in blocks}
    if len(groups) < 2 or not any(groups.values()):
        return None
    return groups


def _region_hierarchy(df: pd.DataFrame, years: list[int]) -> pd.DataFrame:
    """Give a workbook a Global -> Region level when it never states one.

    `_country_hierarchy` can parent countries off sheet order, but only onto
    regions the workbook already declares.  A workbook that declares neither
    (the Capillary Blood sheet lists all 40 geographies flat) leaves REGIONS
    empty, and then every regional exhibit reads 'Global is the largest
    region, at 100%', 'Top 0 Countries' comes out blank, and the country
    section renders Global over and over again.

    Which geographies are regions is decided by ARITHMETIC, not by their
    names: walking the sheet's own order, a geography is a region when the
    geographies listed under it add up to it.  That carries to any market, and
    it cannot invent a hierarchy the numbers do not support.
    """
    if (df["Segment"] == "Region").any():
        return df                                # already stated

    total_seg = (_self_referential_segment(df)
                 or _total_marker_segment(df, years))

    def _sum_years(rows):
        # Coerce to numbers first, so a placeholder cell ('XXXX', a dash, a
        # stray note) becomes NaN and is skipped, rather than making the sum
        # a string that later blows up float(). Real data is already numeric,
        # so this is a no-op for a valid workbook.
        return rows[years].apply(pd.to_numeric, errors="coerce").sum()

    def _total(geo, unit):
        rows = df[(df["Region"] == geo) & (df["Unit"] == unit)]
        if total_seg is not None:
            t = rows[rows["Segment"] == total_seg]
            if len(t):
                return _sum_years(t)
        body = rows[~rows["Segment"].isin(_HIERARCHY_SEGS)]
        if body.empty:
            return None
        first = body["Segment"].iloc[0]
        return _sum_years(body[body["Segment"] == first])

    units = list(df["Unit"].drop_duplicates())
    order = [g for g in dict.fromkeys(df["Region"].tolist()) if g != "Global"]
    if len(order) < 3:
        return df

    totals = {g: _total(g, units[0]) for g in order}
    if any(v is None for v in totals.values()):
        return df

    groups = _region_blocks_by_sums(order, totals)
    if groups is None:
        # the numbers do not nest - fall back to what the sheet's LAYOUT says
        groups = _region_blocks_by_layout(order)
        if groups:
            global _LAYOUT_HIERARCHY
            _LAYOUT_HIERARCHY = True
    if not groups:
        return df

    rows = []
    for reg in groups:
        for unit in units:
            vals = _total(reg, unit)
            if vals is None:
                return df
            rows.append(["Global", "Region", reg, unit, *vals.tolist()])
    return pd.concat([df, pd.DataFrame(rows, columns=_TIDY_COLS + years)],
                     ignore_index=True)


# Segment -> main sub-segment -> the names one level beneath it.  Filled by
# `_load_master_sheet` from the deeper columns before they are aggregated
# away; only the taxonomy page reads it, every figure still comes from the
# collapsed rows.  Empty for a workbook with no deeper level.
SEGMENT_CHILDREN: dict[str, dict[str, list[str]]] = {}

# Cross-section cells, kept before the deeper columns are folded away: one row
# per (Region, Segment, Sub-segment, Child, Unit) of every segment named
# 'A x B', with its yearly values - the Cross-sectional Assessment pages read
# their A x B matrices from it (2026-09-29).  Empty for a workbook without
# that deeper level.
CROSS_CELLS: pd.DataFrame = pd.DataFrame()


# 'A x B' and 'A × B' (the multiplication sign) both name a cross-tab segment
CROSS_SPLIT = re.compile(r"\s+[x×]\s+", re.I)


def is_cross(segment) -> bool:
    """True for a cross-tab segment named 'A x B' / 'A × B'."""
    return len(CROSS_SPLIT.split(str(segment))) > 1


def cross_axes(segment) -> tuple[str, str]:
    """('Material Type', 'End Use Application') for
    'Material Type x End Use Application'."""
    parts = CROSS_SPLIT.split(str(segment), maxsplit=1)
    return (parts[0].strip(), parts[1].strip()) if len(parts) == 2         else (str(segment), "")


def _cross_b_column(rows: pd.DataFrame, deeper: list[str]):
    """The deeper column that holds the B axis of a cross-tab.

    A cross lists, for every A sub-segment, the same B sub-segments - so the
    B column is the one whose set of values is the SAME under every A.  In a
    two-level cross (A x B) 'Sub-segment 1' and 'Sub-segment 2' both name B;
    in a three-level one (A main > A sub x B) 'Sub-segment 1' is A's own
    child, which differs under each A, and 'Sub-segment 2' is B.  Falls back
    to the first column that names something other than A.
    """
    def same_under_every_a(col):
        sets = [frozenset(g[col].dropna().astype(str).str.strip())
                for _a, g in rows.groupby("Sub-segment", sort=False)]
        return bool(sets) and all(x == sets[0] for x in sets) and len(sets[0]) > 0
    for col in deeper:
        if same_under_every_a(col):
            return col
    return None


def _cross_cells(df: pd.DataFrame, years: list[int]) -> pd.DataFrame:
    deeper = [c for c in ("Sub-segment 1", "Sub-segment 2") if c in df.columns]
    rows = df[df["Segment"].map(is_cross)]
    if rows.empty or not deeper:
        return pd.DataFrame()
    frames = []
    for _seg, srows in rows.groupby("Segment", sort=False):
        ref = srows[srows["Region"] == srows["Region"].iloc[0]]
        bcol = _cross_b_column(ref, deeper)
        cols = [bcol] if bcol else deeper

        def child(r, _cols=cols):
            for c in _cols:
                v = r[c]
                if pd.notna(v) and str(v).strip() and                         str(v).strip().casefold() != str(r["Sub-segment"]).casefold():
                    return str(v).strip()
            return None
        frames.append(srows.assign(Child=srows.apply(child, axis=1)))
    rows = pd.concat(frames).dropna(subset=["Child"])
    if rows.empty:
        return pd.DataFrame()
    return (rows.groupby(["Region", "Segment", "Sub-segment", "Child", "Unit"],
                         sort=False)[years].sum().reset_index())


def cross_matrix(region: str, unit: str, segment: str, year: int):
    """Sub-segment x child matrix of `segment` for `region` at `year`, in the
    workbook's own order; an empty frame when the workbook has no cells."""
    c = CROSS_CELLS
    if c is None or c.empty or year not in c.columns:
        return pd.DataFrame()
    g = c[(c["Region"] == region) & (c["Segment"] == segment)
          & (c["Unit"] == unit)]
    if g.empty:
        return pd.DataFrame()
    rows = list(dict.fromkeys(g["Sub-segment"]))
    cols = list(dict.fromkeys(g["Child"]))
    m = g.pivot_table(index="Sub-segment", columns="Child", values=year,
                      aggfunc="sum")
    return m.reindex(index=rows, columns=cols).fillna(0.0).astype(float)


def cross_b_groups(segment: str) -> list[tuple[str, list[str]]]:
    """[(main, [sub, ...])] for the B axis of a cross-tab, in the workbook's
    order - 'Cleansing & General Hygiene' over 'Dirt & Sebum Removal' and
    'Routine Body & Coat Hygiene'.  The grouping comes from B's own segment
    (its sub-segment > child hierarchy).  A B axis with no such level gives
    one group per sub-segment, each named for itself."""
    c = CROSS_CELLS
    if c is None or c.empty:
        return []
    subs = list(dict.fromkeys(c.loc[c["Segment"] == segment, "Child"]))
    b_name = cross_axes(segment)[1].casefold()
    kids_of = next((v for k, v in SEGMENT_CHILDREN.items()
                    if k.casefold() == b_name), {})
    main_of = {kid: main for main, kids in kids_of.items() for kid in kids}
    out: dict[str, list[str]] = {}
    for sub in subs:
        out.setdefault(main_of.get(sub, sub), []).append(sub)
    return list(out.items())


def cross_matrix_main(region: str, unit: str, segment: str, year: int):
    """`cross_matrix` with the B axis rolled up to its main sub-segments."""
    m = cross_matrix(region, unit, segment, year)
    if m.empty:
        return m
    out = {main: m[[s for s in subs if s in m.columns]].sum(axis=1)
           for main, subs in cross_b_groups(segment)}
    return pd.DataFrame(out, index=m.index)


def _segment_children(df: pd.DataFrame) -> dict[str, dict[str, list[str]]]:
    """The taxonomy's second level, read off 'Sub-segment 1/2'.

    The first deeper column that names something other than the parent
    supplies the children.  A column that only repeats the parent (a promoted
    segment, or a leaf restated to fill the row) is skipped, and one holding
    facets ('By Closure') follows the first facet, the way `_collapse_deeper`
    counts it."""
    deeper = [c for c in ("Sub-segment 1", "Sub-segment 2") if c in df.columns]
    out: dict[str, dict[str, list[str]]] = {}
    if not deeper:
        return out
    body = df[~df["Segment"].isin(_HIERARCHY_SEGS)]
    for (seg, sub), rows in body.groupby(["Segment", "Sub-segment"],
                                         sort=False):
        kids: list[str] = []
        for col in deeper:
            vals = rows[col].dropna().astype(str).str.strip()
            vals = [v for v in dict.fromkeys(vals)
                    if v and v.casefold() != str(sub).casefold()]
            kids = [v for v in vals if not _is_facet(v)]
            if kids:
                break
            facets = [v for v in vals if _is_facet(v)]
            if facets:
                rows = rows[rows[col].astype(str).str.strip() == facets[0]]
        if kids:
            out.setdefault(str(seg), {})[str(sub)] = kids
    return out


# Second-level rows (segment > sub-segment > child) with their yearly values,
# kept before `_collapse_deeper` folds them into the sub-segment: the size and
# YoY TABLES list them beneath their parent (client, 2026-10-07).  Empty for a
# workbook with no deeper level.
LEAF_CELLS: pd.DataFrame = pd.DataFrame()


def _leaf_cells(df: pd.DataFrame, years: list) -> pd.DataFrame:
    """One row per (Region, Segment, Sub-segment, Child, Unit) with its values,
    the child column chosen exactly as `_segment_children` chooses it."""
    deeper = [c for c in ("Sub-segment 1", "Sub-segment 2") if c in df.columns]
    if not deeper:
        return pd.DataFrame()
    body = df[~df["Segment"].isin(_HIERARCHY_SEGS) & ~df["Segment"].map(is_cross)]
    frames = []
    for (seg, sub), rows in body.groupby(["Segment", "Sub-segment"],
                                         sort=False):
        for col in deeper:
            vals = rows[col].dropna().astype(str).str.strip()
            vals = [v for v in dict.fromkeys(vals)
                    if v and v.casefold() != str(sub).casefold()]
            kids = [v for v in vals if not _is_facet(v)]
            if kids:
                part = rows[rows[col].astype(str).str.strip().isin(kids)]
                part = part.assign(Child=part[col].astype(str).str.strip())
                frames.append(part)
                break
            facets = [v for v in vals if _is_facet(v)]
            if facets:
                rows = rows[rows[col].astype(str).str.strip() == facets[0]]
    if not frames:
        return pd.DataFrame()
    allr = pd.concat(frames)
    return (allr.groupby(["Region", "Segment", "Sub-segment", "Child", "Unit"],
                         sort=False)[years].sum().reset_index())


def deep_table(geo: str, unit: str, segment: str):
    """`segment_table` with each sub-segment's own children listed beneath it
    ('– Fruit-based'), for the size and YoY tables.  Returns (table, mains):
    `mains` are the parent rows - the only ones a total or a market-pace
    figure is built from.  A workbook (or geography) with no second level
    gives the plain table and every row a main."""
    main = segment_table(geo, unit, segment)
    c = LEAF_CELLS
    if c is None or c.empty or main.empty:
        return main, list(main.index)
    g = c[(c["Region"] == geo) & (c["Segment"] == segment)
          & (c["Unit"] == unit)]
    if g.empty:
        return main, list(main.index)
    rows, idx, used = [], [], set()
    for sub in main.index:
        rows.append(main.loc[sub])
        idx.append(sub)
        used.add(sub)
        kids = g[g["Sub-segment"] == sub]
        if len(kids) < 2:               # one child only restates its parent
            continue
        for _, r in kids.iterrows():
            vals = r[YEARS].astype(float)
            if vals.fillna(0.0).abs().sum() <= 0:
                continue
            name = f"– {r['Child']}"
            while name in used:
                name += " "
            used.add(name)
            rows.append(vals)
            idx.append(name)
    full = pd.DataFrame(rows, index=idx)[YEARS].astype(float)
    return full, list(main.index)


def _load_master_sheet(sheet: str = "Master Sheet") -> pd.DataFrame:
    # units come from the datasheet's own titles, kept verbatim
    _apply_sheet_unit_labels(sheet, "Value", "Volume")
    df = pd.read_excel(DATA_FILE, sheet_name=sheet, header=5)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df.columns = [
        int(float(c)) if str(c).replace(".", "").isdigit() else str(c).strip()
        for c in df.columns
    ]
    df = df.dropna(subset=["Region", "Segment", "Sub-segment", "Unit"])
    for c in _TIDY_COLS:
        df[c] = df[c].astype(str).str.strip()
    # Some workbooks state the segment the way the slide reads it - 'By
    # Region', 'By Product Type'.  The report model names a segment WITHOUT
    # that prefix (`templates_main._by` puts it back for display), and the
    # Region/Country hierarchy rows are found by name, so 'By Region' would
    # otherwise be read as an ordinary product segment and the geography
    # hierarchy would vanish.  A hyphen is safe: 'By-product Type' has no
    # space after 'By'.
    df["Segment"] = df["Segment"].str.replace(r"^By\s+", "", regex=True,
                                              case=False)
    df["Unit"] = _normalise_units(df["Unit"])
    df = _first_year_run(df)
    years = [c for c in df.columns if isinstance(c, int)]

    # A segment whose 'Sub-segment' column says NOTHING - one value, and that
    # value is the segment's own name - carries its real breakdown one column
    # deeper.  The Denture Care workbook heads four of its six segments that
    # way: 'Denture Type / Denture Type', with Complete / Partial / Removable
    # Implant-Supported sitting in 'Sub-segment 1'.  Promote that column for
    # those segments, or the aggregation below sums the whole segment into ONE
    # row and every page for it shows a single bar named after the segment.
    # A segment whose top column DOES discriminate is left alone, so the
    # report still leads with the level the workbook leads with.
    _deeper = next((c for c in ("Sub-segment 1", "Sub-segment 2")
                    if c in df.columns), None)
    if _deeper is not None:
        for _seg, _rows in df.groupby("Segment", sort=False):
            _subs = _rows["Sub-segment"].dropna().unique()
            if len(_subs) != 1:
                continue                      # a real breakdown already
            if str(_subs[0]).strip().casefold() != str(_seg).strip().casefold():
                continue                      # one genuine sub-segment
            _promoted = _rows[_deeper]
            if _promoted.notna().any() and _promoted.nunique() > 1:
                df.loc[_rows.index, "Sub-segment"] = _promoted

    # The taxonomy page lists that deeper level by name, so record it before
    # the aggregation below folds it into its parent.
    global SEGMENT_CHILDREN, CROSS_CELLS, LEAF_CELLS
    SEGMENT_CHILDREN = _segment_children(df)
    CROSS_CELLS = _cross_cells(df, years)
    LEAF_CELLS = _leaf_cells(df, years)

    # Some workbooks carry deeper leaf rows in 'Sub-segment 1/2' columns -
    # aggregate to the top Sub-segment level the report model uses.
    df = _collapse_deeper(df, years)

    # Region-rooted workbooks (e.g. a Europe-only study) have no 'Global'
    # rows - synthesise them so the drill hierarchy is always
    # Global → region → country.
    if not (df["Region"] == "Global").any():
        child_subs = set(df.loc[df["Segment"].isin(_HIERARCHY_SEGS),
                                "Sub-segment"])
        tops = [r for r in df["Region"].drop_duplicates()
                if r not in child_subs]

        # A single aggregate root that already carries its own 'Region'
        # breakdown (super-region → sub-regions → countries, e.g. Europe →
        # Eastern/Western Europe → countries) IS the global aggregate.
        # Relabel it 'Global' rather than nesting a new level above it -
        # otherwise the real regions/countries get orphaned and the drill
        # (and every child chart) comes back empty.
        region_parents = set(df.loc[df["Segment"] == "Region", "Region"])
        if len(tops) == 1 and tops[0] in region_parents:
            global _RELABELED_ROOT
            _RELABELED_ROOT = tops[0]
            df = df.copy()
            df.loc[df["Region"] == tops[0], "Region"] = "Global"
            df = _reparent_selfref_geographies(df)
            return _country_hierarchy(_region_hierarchy(df, years),
                                      years).reset_index(drop=True)

        seg0 = next(s for s in df["Segment"].drop_duplicates()
                    if s not in _HIERARCHY_SEGS)
        extra = []
        for top in tops:                       # Global → region link rows
            rows = df[(df["Region"] == top) & (df["Segment"] == seg0)]
            for unit, grp in rows.groupby("Unit"):
                extra.append(["Global", "Region", top, unit,
                              *grp[years].sum().tolist()])
        # Global segment tables = sum over the top regions
        top_rows = df[df["Region"].isin(tops)
                      & ~df["Segment"].isin(_HIERARCHY_SEGS)]
        glob = (top_rows.groupby(["Segment", "Sub-segment", "Unit"],
                                 sort=False)[years].sum().reset_index())
        glob.insert(0, "Region", "Global")
        df = pd.concat([pd.DataFrame(extra, columns=_TIDY_COLS + years),
                        glob, df], ignore_index=True)
    # A workbook that declares each geography's LEVEL against itself rather
    # than against its parent states the whole tree, but states no edge; give
    # it the edges before the two hierarchy builders decide it needs none.
    df = _reparent_selfref_geographies(df)
    return _country_hierarchy(_region_hierarchy(df, years),
                              years).reset_index(drop=True)


# ── Format 2: one sheet per geography ────────────────────────────────
def _find_year_header(df: pd.DataFrame) -> tuple[int, list[int], list[int]]:
    """Locate the header row and the first contiguous run of year columns.

    Returns (header_row, year_col_positions, years)."""
    for r in range(min(15, len(df))):
        row = df.iloc[r]
        cols, years = [], []
        for c in range(len(row)):
            y = _year_like(row.iloc[c])
            if y is not None:
                if cols and c != cols[-1] + 1:
                    break                       # first contiguous block only
                cols.append(c)
                years.append(y)
        if len(years) >= 4 and years == sorted(years):
            return r, cols, years
    raise ValueError("no year header row found")


def _prune_to_top_level(rows: list[tuple[str, np.ndarray]],
                        block_total: np.ndarray | None,
                        ) -> list[tuple[str, np.ndarray]]:
    """Sheets nest parents above their children (parent = sum of the
    following rows).  Keep only top-level rows so column sums equal the
    segment total; fall back to the raw rows if pruning fits worse."""
    def total(rs):
        return np.nansum([v.sum() for _, v in rs])

    pruned, i = [], 0
    while i < len(rows):
        name, vec = rows[i]
        target = np.nansum(vec)
        acc, j, matched = 0.0, i + 1, None
        while j < len(rows) and acc < target * 1.02:
            acc += np.nansum(rows[j][1])
            if j > i + 1 and abs(acc - target) <= abs(target) * 0.015:
                matched = j
                break
            j += 1
        pruned.append((name, vec))
        i = (matched + 1) if matched is not None else i + 1

    if block_total is not None and np.nansum(block_total) > 0:
        tgt = np.nansum(block_total)
        if abs(total(pruned) - tgt) > abs(total(rows) - tgt):
            return rows                          # pruning made the fit worse
    return pruned


def _parse_geo_sheet(name: str, df: pd.DataFrame, years: list[int],
                     year_cols: list[int], header_row: int):
    """Parse one geography tab → (totals per unit, tidy segment rows)."""
    seg_col, sub_col = year_cols[0] - 2, year_cols[0] - 1
    totals: dict[str, np.ndarray] = {}
    # segment blocks: {(unit, segment): [(sub, values), ...]}
    blocks: dict[tuple[str, str], list[tuple[str, np.ndarray]]] = {}
    block_totals: dict[tuple[str, str], np.ndarray] = {}
    unit = "Volume"                              # sheets start with volume

    for r in range(header_row + 1, len(df)):
        seg = df.iat[r, seg_col]
        sub = df.iat[r, sub_col]
        seg = str(seg).strip() if pd.notna(seg) else ""
        sub = str(sub).strip() if pd.notna(sub) else ""
        if seg in ("Price", "Value"):            # section markers
            unit = seg
            continue
        vals = pd.to_numeric(df.iloc[r, year_cols], errors="coerce"
                             ).to_numpy(dtype=float)
        if not seg or np.all(np.isnan(vals)):
            continue
        if seg == name:                          # geography total rows
            for u in ("Volume", "Value"):
                if sub == f"{name} {u}" and u not in totals:
                    totals[u] = vals
            continue
        if unit == "Price":                      # $/unit tables - skip
            continue
        if sub == seg or sub in ("", "0", "0.0"):  # block-total / filler
            if sub == seg:
                block_totals[(unit, seg)] = vals
            continue
        blocks.setdefault((unit, seg), []).append((sub, vals))

    tidy = []
    for (u, seg), rows in blocks.items():
        for sub, vals in _prune_to_top_level(rows, block_totals.get((u, seg))):
            tidy.append([name, seg, sub, u, *vals])
    return totals, tidy


def _load_geo_sheets() -> pd.DataFrame:
    xl = pd.ExcelFile(DATA_FILE)
    sheets: dict[str, dict] = {}
    years_ref: list[int] | None = None

    for sn in xl.sheet_names:
        raw = pd.read_excel(xl, sheet_name=sn, header=None)
        try:
            hdr, ycols, years = _find_year_header(raw)
        except ValueError:
            continue                             # not a data tab
        if years_ref is None:
            years_ref = years
        totals, tidy = _parse_geo_sheet(sn.strip(), raw, years, ycols, hdr)
        if totals and tidy:
            sheets[sn.strip()] = {"totals": totals, "tidy": tidy}

    if not sheets:
        raise ValueError("no per-geography sheets recognised")

    config.set_years(years_ref)

    # aggregate sheet = the one whose total ≈ the sum of all the others
    names = list(sheets)
    parent = None
    for cand in names:
        others = [n for n in names if n != cand]
        if not others:
            continue
        cand_t = np.nansum(sheets[cand]["totals"].get("Volume",
                           sheets[cand]["totals"].get("Value")))
        rest_t = sum(np.nansum(sheets[n]["totals"].get("Volume",
                     sheets[n]["totals"].get("Value"))) for n in others)
        if rest_t and abs(cand_t - rest_t) <= cand_t * 0.02:
            parent = cand
            break

    rows: list[list] = []

    def add_totals(region, segment, subsegment, totals):
        for u, vals in totals.items():
            rows.append([region, segment, subsegment, u, *vals])

    if parent:                                   # Global → region → countries
        children = [n for n in names if n != parent]
        add_totals("Global", "Region", parent, sheets[parent]["totals"])
        for row in sheets[parent]["tidy"]:       # Global mirrors the aggregate
            rows.append(["Global", *row[1:]])
        for child in children:
            add_totals(parent, "Country", child, sheets[child]["totals"])
        for n in names:
            rows.extend(sheets[n]["tidy"])
    else:                                        # flat: all sheets ⇒ regions
        for n in names:
            add_totals("Global", "Region", n, sheets[n]["totals"])
            rows.extend(sheets[n]["tidy"])
        # synthesise the Global segment tables as the sum over sheets
        flat = pd.DataFrame(rows, columns=_TIDY_COLS + years_ref)
        seg_rows = flat[~flat["Segment"].isin(["Region", "Country"])]
        glob = (seg_rows.groupby(["Segment", "Sub-segment", "Unit"],
                                 sort=False)[years_ref].sum().reset_index())
        glob.insert(0, "Region", "Global")
        rows.extend(glob.values.tolist())

    df = pd.DataFrame(rows, columns=_TIDY_COLS + years_ref)
    # Collapse any duplicate (Region, Segment, Sub-segment, Unit) rows - some
    # per-geography sheets repeat a segment block, which would otherwise give
    # segment tables a duplicate index and crash the report.  Mirrors the
    # Master-Sheet loader's final groupby.
    df = df.groupby(_TIDY_COLS, sort=False)[years_ref].sum().reset_index()
    return df


# ── Format 3: split 'Volume Data' / 'Value Data' sheets ──────────────
def _load_volume_value_sheets(sheet_names: list[str]) -> pd.DataFrame:
    """Workbooks that split Volume and Value into their own sheets, with a
    richer dimension block (e.g. the Food Color dataset):

        Region | Color Type | Color Type Sub Category 1 | Segment |
        Sub-Segment 1 | Sub-Segment 2 | 2020 ... 2035

    (header on the 3rd row).  Maps to the standard tidy model - the
    ``Segment`` column becomes the segment, ``Sub-Segment 1`` the
    sub-segment, and everything finer (the Color-Type dimension and the
    ``Sub-Segment 2`` leaves) is summed away.  Unit comes from the sheet.
    These files are large (100k+ rows) so we read with the fast
    ``calamine`` engine, falling back to the default if unavailable.
    """
    unit_of = {}
    for sn in sheet_names:
        low = sn.strip().lower()
        if low in ("volume data", "volume"):
            unit_of[sn] = "Volume"
        elif low in ("value data", "value"):
            unit_of[sn] = "Value"
    if len(unit_of) < 2:
        raise ValueError("no Volume/Value data sheets recognised")

    def _keep(col):                          # only the columns we model -
        s = str(col).strip()                 # keeps memory low on big files
        if s in ("Region", "Segment", "Sub-Segment 1"):
            return True
        try:
            return 1990 <= int(float(s)) <= 2100
        except (TypeError, ValueError):
            return False

    def _read(sheet):
        try:
            return pd.read_excel(DATA_FILE, sheet_name=sheet, header=2,
                                 engine="calamine", usecols=_keep)
        except Exception:
            return pd.read_excel(DATA_FILE, sheet_name=sheet, header=2,
                                 usecols=_keep)

    frames, years_ref = [], None
    for sheet, unit in unit_of.items():
        # units live in the sheet title, e.g. 'Volume (Kg)' / 'Value (US$
        # Thousand)' - pull them out for the display labels
        title = _title_cell(sheet)
        m = re.search(rf"{unit}\s*\(([^)]+)\)", title, flags=re.I)
        if m:
            config.set_unit_labels(**{unit.lower(): m.group(1)})
        raw = _read(sheet)
        raw = raw.loc[:, ~raw.columns.astype(str).str.startswith("Unnamed")]
        raw.columns = [
            int(c) if isinstance(c, (int, float)) and float(c).is_integer()
            and 1990 <= c <= 2100 else str(c).strip()
            for c in raw.columns
        ]
        years = [c for c in raw.columns if isinstance(c, int)]
        if years_ref is None:
            years_ref = years
        if not {"Region", "Segment", "Sub-Segment 1"}.issubset(raw.columns):
            raise ValueError(f"{sheet}: missing Region/Segment/Sub-Segment 1")
        tidy = raw[["Region", "Segment", "Sub-Segment 1"] + years].copy()
        tidy.columns = ["Region", "Segment", "Sub-segment"] + years
        tidy = tidy.dropna(subset=["Region", "Segment", "Sub-segment"])
        for c in ("Region", "Segment", "Sub-segment"):
            tidy[c] = tidy[c].astype(str).str.strip()
        for y in years:
            tidy[y] = pd.to_numeric(tidy[y], errors="coerce")
        tidy["Unit"] = unit
        tidy = (tidy.groupby(_TIDY_COLS, sort=False)[years]
                    .sum().reset_index())
        frames.append(tidy)

    df = pd.concat(frames, ignore_index=True)
    config.set_years(years_ref)
    return df.reset_index(drop=True)


# ── Loader dispatch ──────────────────────────────────────────────────
_MASTER_COLS = {"Region", "Segment", "Sub-segment", "Unit"}
# the master sheet actually used, so the unit labels and the market name are
# read off the sheet that carried the data rather than off a guessed name
_MASTER_SHEET: str | None = None


def _master_sheet_name(sheet_names) -> str | None:
    """The master-data sheet, found by its SHAPE rather than its name.

    'Master Sheet' is the FMI convention and is still preferred, but the same
    layout arrives under other names ('Market Data'), and a workbook that is
    otherwise identical should not need an edit to load.  A sheet qualifies
    when its header row carries Region / Segment / Sub-segment / Unit -
    which no other sheet in these workbooks does.
    """
    if "Master Sheet" in sheet_names:
        return "Master Sheet"
    for name in sheet_names:
        try:
            head = pd.read_excel(DATA_FILE, sheet_name=name, header=5, nrows=0)
        except Exception:
            continue
        cols = {str(c).strip() for c in head.columns}
        if _MASTER_COLS.issubset(cols):
            return name
    return None


def _load_master() -> pd.DataFrame:
    if DATA_FILE is None:
        # No bundled workbook (the portal case). Return a well-formed but
        # EMPTY frame so the module-level derivations below produce empty
        # lists instead of raising; the portal rebinds real values per
        # session, so this import-time default is never rendered.
        return pd.DataFrame(columns=_TIDY_COLS + list(config.YEARS))
    try:
        sheet_names = pd.ExcelFile(DATA_FILE, engine="calamine").sheet_names
    except Exception:
        sheet_names = pd.ExcelFile(DATA_FILE).sheet_names
    master = _master_sheet_name(sheet_names)
    if master:
        global _MASTER_SHEET
        _MASTER_SHEET = master
        df = _load_master_sheet(master)
        config.set_years([c for c in df.columns if isinstance(c, int)])
        return df
    lower = {s.strip().lower() for s in sheet_names}
    if {"volume data", "value data"} & lower and \
            len([s for s in lower if s in
                 ("volume data", "value data", "volume", "value")]) >= 2:
        return _load_volume_value_sheets(sheet_names)
    return _load_geo_sheets()


MASTER = _load_master()

SEGMENTS: list[str] = [
    s for s in MASTER["Segment"].drop_duplicates().tolist()
    if s not in _HIERARCHY_SEGS
]


def taxonomy_axes() -> list[str]:
    """The segments a taxonomy lists: every product segment except a cross-tab
    of two others ('Product Type x Therapeutic Application'), which re-lists
    their combinations rather than naming a new axis.  Recognised by its
    parts, not its wording: each side of the ' x ' must be a segment itself."""
    names = {s.casefold() for s in SEGMENTS}

    def cross(s: str) -> bool:
        parts = CROSS_SPLIT.split(str(s))
        return len(parts) > 1 and all(p.strip().casefold() in names
                                      for p in parts)
    return [s for s in SEGMENTS if not cross(s)]


# ── Unit auto-scaling ────────────────────────────────────────────────
def _bump_label(label: str) -> str:
    """Next-higher display unit for one ÷1000 step of the data."""
    l = label.strip()
    low = l.lower()
    if "us$" in low or "usd" in low or "$" in l:
        for a, b in (("thousand", "Mn"), ("'000", "Mn"), ("million", "Bn"),
                     ("mn", "Bn"), ("billion", "Tn"), ("bn", "Tn")):
            if a in low:
                return re.sub(a, b, l, flags=re.IGNORECASE)
        return f"{l} '000"
    if low == "kg":
        return "Tons"
    if low.startswith("'000 "):
        return "Mn " + l[5:]
    if low.startswith("mn "):
        return "Bn " + l[3:]
    return f"'000 {l}"


def _autoscale_units() -> None:
    """DISABLED - the report now uses the datasheet's own unit verbatim.

    Previously each Unit was divided by a global factor so exhibits showed
    short numbers (e.g. 3,496,222 US$ Thousand -> 3,496 US$ Mn), but that
    changed the unit label away from what the workbook states.  Per request
    the raw sheet unit ('US$ Thousand', 'Kg', ...) and raw values are kept
    everywhere; chart axes still abbreviate the *numbers* to K/M/B while the
    unit label stays exactly as the datasheet.  Left as a no-op so the call
    site and the _bump_label helper remain for reference."""
    return


_autoscale_units()


def seg(i: int) -> str:
    """i-th product segment (wraps around for workbooks with fewer)."""
    return SEGMENTS[i % len(SEGMENTS)]


# ── Market name (configurable) ───────────────────────────────────────
def _detect_market_name() -> str:
    """MARKET_NAME env var wins; otherwise parse the Master Sheet title
    cell, e.g. 'Global Matcha Market Value (US$ Mn) and Volume (Tons)
    Analysis and Forecast, 2021-2036' -> 'Matcha'; else derive it from
    the workbook file name, e.g. 'Tortilla Market.xlsx' -> 'Tortilla'."""
    env = os.environ.get("MARKET_NAME")
    if env:
        return env.strip()
    if DATA_FILE is None:
        # No bundled workbook (the portal). Never rendered - the portal
        # rebinds the real name per session - so any placeholder is fine.
        return "Report"

    def strip_geo(name: str) -> str:
        # drop a leading geography ("Europe Tortilla" -> "Tortilla") -
        # every page title already prefixes the selected geography
        geos = set(MASTER["Region"]) | {"Global"}
        if _RELABELED_ROOT:                    # e.g. 'Europe' -> 'Global'
            geos.add(_RELABELED_ROOT)
        for g in sorted(geos, key=len, reverse=True):
            if name.lower().startswith(g.lower() + " "):
                return name[len(g):].strip()
        return name

    for sheet in (_MASTER_SHEET or "Master Sheet", "Master Sheet",
                  "Volume Data", "Value Data"):
        cell = _title_cell(sheet)
        m = re.match(r"(.+?)\s+Market\b", cell)
        if m:
            return strip_geo(m.group(1).strip())
    stem = DATA_FILE.stem
    # underscores first: '\bMarket\b' cannot see a boundary inside
    # 'Market_Filled', so the word survived and the label came out as
    # '... Market Filled Market'
    stem = stem.replace("_", " ")
    stem = re.sub(r"^Dataset[_\s]*", "", stem, flags=re.I)
    stem = re.sub(r"[_\s]*FMI[_\s]*", "", stem, flags=re.I)
    stem = re.sub(r"\bGlobal\b", "", stem, flags=re.I)
    stem = re.sub(r"\bMarket\b.*$", "", stem, flags=re.I)
    stem = stem.strip(" -_")
    return strip_geo(stem) if stem else "Market"


MARKET = _detect_market_name()          # e.g. "Matcha", "Tortilla"
MARKET_LABEL = f"{MARKET} Market"       # e.g. "Matcha Market"

# ── Geography hierarchy ──────────────────────────────────────────────
REGIONS: list[str] = (
    MASTER.query("Region == 'Global' and Segment == 'Region'")["Sub-segment"]
    .drop_duplicates()
    .tolist()
)

# Geographies that carry their own segment breakdown rows.  Some
# workbooks list a country in the Region→Country hierarchy but provide
# no segment data for it (e.g. Estonia/Latvia/Cyprus in the Decorations
# workbook) - every country-report page for it would be blank, so such
# countries are hidden from the country slicer, drill-through and
# summary indices.  They still appear in region-level tables/donuts,
# which read the hierarchy rows directly.
_SEG_GEOS: set[str] = set(
    MASTER.loc[~MASTER["Segment"].isin(_HIERARCHY_SEGS), "Region"]
)

# child -> parent, off the hierarchy rows themselves
_PARENT: dict[str, str] = {}
for _hseg in _HIERARCHY_SEGS:
    for _p, _c in (MASTER.loc[MASTER["Segment"] == _hseg,
                              ["Region", "Sub-segment"]]
                   .drop_duplicates().itertuples(index=False)):
        _PARENT.setdefault(str(_c), str(_p))


def _region_of(geo: str):
    """The REGION a geography sits under, however many levels down.

    A three-level workbook parents its countries on a SUB-REGION
    (Europe -> Western Europe -> Germany), so grouping the Country rows by
    their own parent keyed them on "Western Europe" and the region's country
    list came back empty.  Climbing to the region flattens that level for the
    report's Global -> region -> country model; the sub-regions are still
    there as their own hierarchy rows.
    """
    seen, cur = set(), geo
    while cur and cur not in REGIONS and cur not in seen:
        seen.add(cur)
        cur = _PARENT.get(cur)
    return cur if cur in REGIONS else None


COUNTRIES: dict[str, list[str]] = {}
for _region, _rows in MASTER.query("Segment == 'Country'").groupby(
        "Region", sort=False):
    _key = _region if _region in REGIONS else (_region_of(_region) or _region)
    _bucket = COUNTRIES.setdefault(_key, [])
    for _c in _rows["Sub-segment"].drop_duplicates().tolist():
        if _c in _SEG_GEOS and _c not in _bucket:
            _bucket.append(_c)

ALL_COUNTRIES: list[str] = [c for r in REGIONS for c in COUNTRIES.get(r, [])]
if not ALL_COUNTRIES:                   # flat workbooks: regions stand in
    ALL_COUNTRIES = list(REGIONS)
# ── which units this workbook actually carries ────────────────────────
# A study is normally sized in both Volume and Value, and the report is built
# for both.  Some workbooks carry ONE of them - a value-only study - and every
# Volume page then has nothing to draw and every price nothing to divide by
# (price is Value / Volume, derived, never read).  `registry` reads this to
# leave those pages out rather than emit them empty; nothing else changes, so
# a workbook with both is unaffected.
UNITS_PRESENT: set[str] = {str(u) for u in MASTER["Unit"].dropna().unique()}


def has_unit(unit: str) -> bool:
    """True when the workbook is sized in ``unit`` ("Volume" / "Value")."""
    return str(unit) in UNITS_PRESENT


def has_price() -> bool:
    """True when a price can be derived: it needs BOTH units."""
    return {"Volume", "Value"} <= UNITS_PRESENT


GEOS_TOP = ["Global"] + REGIONS               # slicer on most pages


def is_bucket(name) -> bool:
    """True for a catch-all that is not a country ('Rest of LATAM', 'Other
    European Countries'): it sums several, so a ranking of countries leaves it
    out (client, 2026-10-07)."""
    return str(name).strip().lower().startswith(("rest of", "other "))


def region_of(country: str) -> str | None:
    for region, countries in COUNTRIES.items():
        if country in countries:
            return region
    return None


def children(geo: str) -> list[str]:
    """Sub-geographies shown when *geo* is selected (mirrors the PBIX drill)."""
    if geo == "Global":
        return REGIONS
    return COUNTRIES.get(geo, [])


# ── Segment taxonomy ─────────────────────────────────────────────────
SUBSEGMENTS: dict[str, list[str]] = {
    s: MASTER.query("Region == 'Global' and Segment == @s")["Sub-segment"]
    .drop_duplicates()
    .tolist()
    for s in SEGMENTS
}


# ── Core accessors ───────────────────────────────────────────────────
@lru_cache(maxsize=None)
@lru_cache(maxsize=4096)
def unit_for(segment: str, want: str = "Value", geo: str = "Global") -> str:
    """The unit a segment is actually sized in.

    A workbook may size one breakdown in value and another in volume (the
    robotics study sizes joint type in procedures and revenue stream in value).
    Pages that assume a unit call this first: it returns `want` when the segment
    has rows in it, the other unit when it does not, and `want` when it has
    neither so callers still get a well-formed - if empty - frame.
    """
    other = "Volume" if want == "Value" else "Value"
    for u in (want, other):
        try:
            t = segment_table(geo, u, segment)
        except Exception:
            continue
        if len(t) and float(t.abs().to_numpy().sum()):
            return u
    return want


def segment_table(geo: str, unit: str, segment: str) -> pd.DataFrame:
    """Sub-segment x year matrix for one geography / unit / segment.

    Sub-segments with no business in this geography (all-zero rows, e.g.
    a colour type not sold in a country) are dropped so tables and
    charts only show applicable rows."""
    rows = MASTER[
        (MASTER["Region"] == geo)
        & (MASTER["Unit"] == unit)
        & (MASTER["Segment"] == segment)
    ]
    out = rows.set_index("Sub-segment")[YEARS].astype(float)
    out = out.loc[out.fillna(0.0).abs().sum(axis=1) > 0]
    return out


@lru_cache(maxsize=None)
def total_series(geo: str, unit: str) -> pd.Series:
    """Total market per year for a geography (sum of one full segment).

    The first segment is preferred, but a segment sized in one unit only
    (e.g. a value-only 'Market Total') must not zero out the other unit -
    every Volume exhibit, price (Value ÷ Volume) and growth figure downstream
    would read 0 / nan.  The first segment that actually has rows in the
    requested unit stands in for it."""
    for s in SEGMENTS:
        t = segment_table(geo, unit, s)
        if len(t):
            return t.sum(axis=0)
    return segment_table(geo, unit, SEGMENTS[0]).sum(axis=0)


@lru_cache(maxsize=None)
def children_table(geo: str, unit: str) -> pd.DataFrame:
    """Child-geography x year matrix (regions under Global, countries
    under a region) taken from the hierarchy rows of the Master Sheet."""
    child_seg = "Region" if geo == "Global" else "Country"
    rows = MASTER[
        (MASTER["Region"] == geo)
        & (MASTER["Unit"] == unit)
        & (MASTER["Segment"] == child_seg)
    ]
    out = rows.set_index("Sub-segment")[YEARS].astype(float)
    if out.empty and child_seg == "Country" and COUNTRIES.get(geo):
        # a three-level workbook parents its countries on a SUB-region, so the
        # region itself has no Country rows of its own - take the ones that
        # climb back to it (see `_region_of`), in the workbook's order
        kids = COUNTRIES[geo]
        rows = MASTER[(MASTER["Unit"] == unit)
                      & (MASTER["Segment"] == "Country")
                      & (MASTER["Sub-segment"].isin(kids))]
        out = (rows.drop_duplicates("Sub-segment")
               .set_index("Sub-segment")[YEARS].astype(float)
               .reindex([k for k in kids if k in set(rows["Sub-segment"])]))
    if out.empty:                    # leaf geo / no sub-hierarchy:
        s = total_series(geo, unit)  # show the geography itself as sole child
        out = pd.DataFrame([s.reindex(YEARS).to_numpy(dtype=float)],
                           columns=YEARS, index=[geo])
    return out


def default_country(region: str | None = None) -> str:
    """Largest market by Value, over one region or over the whole dataset.

    With no region this is the country slicer's default, unchanged. With a
    region it is that region's lead market - the U.S. for North America,
    Germany for Europe - which is what a deck built "for North America"
    should open on.

    Derived from the workbook rather than a table of region-to-capital,
    because the lead market is a fact about the data and moves with it: a
    table would be one more thing to remember to edit when a market shifts,
    and it would be wrong for exactly the workbooks nobody checked.
    """
    pool = ALL_COUNTRIES if not region or region == "Global" else         COUNTRIES.get(region) or []
    if not pool:
        # A region with no countries of its own falls back to the dataset
        # default rather than to nothing, so an export still builds.
        pool = ALL_COUNTRIES
    if not pool:
        return "Global"
    try:
        return max(pool,
                   key=lambda c: float(total_series(c, "Value").iloc[-1]))
    except Exception:
        return pool[0]


def valid_geos(region: str | None, country: str | None) -> tuple[str, str]:
    """Coerce a (region, country) selection to values that actually exist
    in the loaded dataset.

    The exporters use this so an out-of-dataset default (e.g. the CLI's
    old ``--country USA`` against a Europe-only workbook) can never blank
    out a report - it falls back to Global / the largest market instead.
    """
    if region not in GEOS_TOP:
        region = "Global"
    if not country or country not in ALL_COUNTRIES:
        country = default_country()
    return region, country
