"""Check a workbook BEFORE exporting from it.

The report can parse a workbook perfectly and still build a deck full of
nonsense, because parsing only asks "is this the right shape?" and never asks
"do these numbers agree with each other?".  Every defect that cost real time on
the Triglyceride study was of the second kind: the sheets, the header row and
the columns were exactly the FMI convention, but five complete segmentations of
the same market gave five different market sizes, and no geography added up to
its parent.  The engine picked one of the five (the first one) and rendered a
confident 158-slide deck on top of it.

This says so instead, in a few seconds:

    python check_workbook.py --data Global_Some_Market.xlsx

Exit status is 0 when the workbook can carry a report, 1 when it cannot.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

TOL = 0.02                      # 2% - the tolerance the loader itself uses

# A CROSS cut - 'Molecular Target x Technology' - re-spends the same money as
# the segmentations it is built from, so it is excluded from the agreement
# test.  Workbooks write the join either way: an ASCII 'x' surrounded by
# spaces, or a real multiplication sign (U+00D7 '×', which Excel produces when
# an analyst types the symbol).  Matching only the ASCII form let the CF-PEEK
# workbook's 'CF-PEEK Product Type × Medical Application' through as though it
# were an ordinary segmentation.
_CROSS_RE = re.compile(r"\s[x×✕✖]\s", re.IGNORECASE)


def _is_cross(seg) -> bool:
    """True when a segmentation name joins two other segmentations."""
    return bool(_CROSS_RE.search(str(seg)))


def _fmt(v: float) -> str:
    return f"{v:,.0f}"


def _spread(vals: list[float]) -> float:
    """How far apart the biggest and smallest readings are, as a fraction."""
    lo, hi = min(vals), max(vals)
    return (hi - lo) / hi if hi else 0.0


def brief() -> str | None:
    """One line for the export log, or None when the workbook is sound.

    Deliberately cheap and quiet: the export already has the package loaded,
    so this only re-reads what is in memory.  Run check_workbook.py itself for
    the full picture."""
    try:
        import pandas as pd
        from fmi_report import data
        from fmi_report.config import YEARS
        body = data.MASTER[~data.MASTER["Segment"].isin(
            ("Region", "Sub-region", "Country"))]
        segs = [s for s in body["Segment"].dropna().unique()
                if not _is_cross(s)]
        notes = []
        for unit in body["Unit"].dropna().unique():
            vals = []
            for s in segs:
                r = body[(body.Region == "Global") & (body.Segment == s)
                         & (body.Unit == unit)]
                if len(r):
                    v = float(r[YEARS].apply(pd.to_numeric,
                                             errors="coerce").sum().sum())
                    if v:
                        vals.append(v)
            if len(vals) > 1 and _spread(vals) > TOL:
                notes.append(f"{unit} segmentations {_spread(vals):.0%} apart")
        if getattr(data, "_LAYOUT_HIERARCHY", False):
            notes.append("geographies do not add up (hierarchy taken from the "
                         "sheet layout)")
        if notes:
            return ("WORKBOOK: " + "; ".join(notes)
                    + "  ->  run check_workbook.py")
    except Exception:
        pass
    return None


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--data", metavar="XLSX",
                    help="workbook (default: MARKET_DATA, else the one in the "
                         "project folder)")
    ap.add_argument("--unit", default=None,
                    help="which unit to test (default: every unit present)")
    args = ap.parse_args(argv)
    if args.data:
        os.environ["MARKET_DATA"] = args.data

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import pandas as pd
    from fmi_report import data
    from fmi_report.config import UNIT_LABELS, YEARS

    m = data.MASTER
    if m is None or not len(m):
        print("FAIL  the workbook loaded no rows at all.")
        return 1

    print(f"workbook : {data.DATA_FILE.name}")
    print(f"market   : {data.MARKET_LABEL}")
    print(f"years    : {YEARS[0]}-{YEARS[-1]}")
    print(f"units    : Value '{UNIT_LABELS['Value']}'  |  "
          f"Volume '{UNIT_LABELS['Volume']}'")

    hier = ("Region", "Sub-region", "Country")
    body = m[~m["Segment"].isin(hier)]
    units = ([args.unit] if args.unit
             else [u for u in body["Unit"].dropna().unique()])
    problems, warnings = [], []

    def total(geo, seg, unit):
        r = body[(body.Region == geo) & (body.Segment == seg)
                 & (body.Unit == unit)]
        if not len(r):
            return None
        return float(r[YEARS].apply(pd.to_numeric, errors="coerce").sum().sum())

    # ── 1. do the segmentations agree? ────────────────────────────────
    # Every complete segmentation of one market must come to the same total.
    # A CROSS cut ('A x B') re-spends the same money and is excluded.
    print("\n1. Do the segmentations agree with each other?")
    segs = [s for s in body["Segment"].dropna().unique() if not _is_cross(s)]
    crossed = [s for s in body["Segment"].dropna().unique() if _is_cross(s)]
    if crossed:
        print(f"   (excluded as cross-cuts: {', '.join(map(str, crossed))})")
    for unit in units:
        vals = {s: total("Global", s, unit) for s in segs}
        vals = {k: v for k, v in vals.items() if v}
        if len(vals) < 2:
            continue
        sp = _spread(list(vals.values()))
        if sp <= TOL:
            print(f"   OK   {unit}: all {len(vals)} segmentations agree at "
                  f"{_fmt(max(vals.values()))}")
        else:
            problems.append(
                f"the {unit} segmentations disagree by {sp:.0%} - there is no "
                f"single market size to report")
            print(f"   FAIL {unit}: {len(vals)} segmentations, {sp:.0%} apart")
            for s, v in sorted(vals.items(), key=lambda kv: -kv[1]):
                print(f"          {str(s):<34}{_fmt(v):>14}")

    # ── 2. is there a stated total? ───────────────────────────────────
    print("\n2. Is each geography's own total stated?")
    # the loader CONSUMES a total-marker segment into the hierarchy and drops
    # it, so look in the sheet itself rather than in the loaded frame
    marker = []
    try:
        raw = pd.read_excel(data.DATA_FILE, sheet_name=data._MASTER_SHEET,
                            header=5, usecols=[1])
        marker = [s for s in raw.iloc[:, 0].dropna().unique()
                  if str(s).strip().lower() in ("total market", "total",
                                                "market")]
    except Exception:
        pass
    if marker:
        print(f"   OK   '{marker[0]}' rows state it")
    else:
        warnings.append("no 'Total Market' segment - each geography's total "
                        "has to be derived by summing a segmentation, which "
                        "is only safe when they all agree (check 1)")
        print("   WARN no 'Total Market' segment; totals are derived")

    # ── 3. do the geographies nest? ───────────────────────────────────
    print("\n3. Do the geographies add up to their parents?")
    if not data.REGIONS:
        problems.append("no region level could be built at all")
        print("   FAIL no regions were found")
    else:
        src = "sheet LAYOUT (the numbers did not support it)" \
            if getattr(data, "_LAYOUT_HIERARCHY", False) else "the numbers"
        print(f"   hierarchy recovered from: {src}")
        seg0 = segs[0] if segs else None
        for unit in units:
            g = total("Global", seg0, unit)
            r = sum(total(x, seg0, unit) or 0.0 for x in data.REGIONS)
            if g and abs(r - g) <= TOL * g:
                print(f"   OK   {unit}: Global {_fmt(g)} = its regions")
            elif g:
                problems.append(
                    f"{unit}: Global ({_fmt(g)}) is not the sum of its regions "
                    f"({_fmt(r)}) - {r / g:.1f}x apart")
                print(f"   FAIL {unit}: Global {_fmt(g)}  vs regions "
                      f"{_fmt(r)}   ({r / g:.1f}x)")
            bad = []
            for reg in data.REGIONS:
                kids = data.COUNTRIES.get(reg) or []
                if not kids:
                    continue
                p = total(reg, seg0, unit)
                k = sum(total(c, seg0, unit) or 0.0 for c in kids)
                if p and abs(k - p) > TOL * p:
                    bad.append(f"{reg} {_fmt(p)} vs {_fmt(k)}")
            if bad:
                problems.append(f"{unit}: {len(bad)} region(s) do not equal "
                                f"their countries")
                for b in bad[:4]:
                    print(f"          {b}")

    # ── verdict ───────────────────────────────────────────────────────
    print("\n" + "=" * 68)
    for w in warnings:
        print(f"WARNING  {w}")
    if not problems:
        print("PASS  this workbook can carry a full report, "
              "Global and per region/country.")
        return 0
    print("FAIL  this workbook cannot carry a trustworthy report:")
    for p in problems:
        print(f"   - {p}")
    print("\nA deck will still BUILD - every page renders and every chart "
          "draws.\nWhat it cannot do is agree with itself, because the source "
          "numbers do not.\nThis is a data fix, not an engine fix.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
