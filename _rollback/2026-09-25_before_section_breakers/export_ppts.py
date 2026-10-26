"""Export the report decks from the command line.
The dashboard's four export buttons, without the dashboard:

    1  this view          the region currently selected, one deck
    2  all regions        one deck per region, zipped
    3  this view locked    (1) with the protected values withheld
    4  all regions locked  (2) with the protected values withheld

Each option calls exactly what the matching button calls, with the same
arguments, so a deck built here is the deck the browser would have sent.

    python export_ppts.py --out "C:\\Users\\Me\\Downloads\\exports"
    python export_ppts.py --out DIR --region "North America" --only 3,4
    python export_ppts.py --data Dataset_Other.xlsx --out DIR --only 1

The workbook is chosen the way the dashboard chooses it - ``--data``, else
MARKET_DATA, else the one Dataset_*.xlsx in the project folder - and must be
set before the package is imported, which is why the import sits inside main().
"""
from __future__ import annotations

import argparse
import os
import sys
import time
import traceback
from pathlib import Path

OPTIONS = {
    "1": "this view",
    "2": "all regions",
    "3": "this view (locked)",
    "4": "all regions (locked)",
}


def _mb(p: Path) -> float:
    return p.stat().st_size / 1e6 if p.exists() else 0.0


def _parse_only(raw: str) -> list[str]:
    """"1,3" / "13" / "all" -> ['1', '3'].  Order is always 1..4."""
    if raw.strip().lower() in ("", "all"):
        return list(OPTIONS)
    picked = {c for c in raw if c in OPTIONS}
    if not picked:
        raise SystemExit(f"--only: nothing recognised in {raw!r}; "
                         f"use digits from {''.join(OPTIONS)} or 'all'")
    return [k for k in OPTIONS if k in picked]


def _geo_help(data) -> str:
    """Why a workbook offers no regions, when it plainly names some.

    The report models Global -> region -> country, and `data` only builds that
    hierarchy when the workbook's own numbers support it: a region has to be
    the sum of the geographies under it.  A sheet that LISTS geographies but
    gives each one an independent set of figures has no hierarchy to find, so
    REGIONS comes back empty - and a --region asked for on top of that used to
    be dropped in silence, which is how a run that said "North America" put a
    GLOBAL deck on disk."""
    named = [g for g in dict.fromkeys(data.MASTER["Region"].dropna().astype(str))
             if g != "Global"]
    if not named:
        return "This workbook carries no geography other than Global."
    return (
        f"This workbook NAMES {len(named)} geographies in its Region column "
        f"({', '.join(named[:4])} ...), but no hierarchy could be built from"
        " them: a region has to equal the sum of the geographies under it,"
        " and these do not add up.  Only 'Global' can be exported until the"
        " workbook's geography figures are made to nest.")


def _resolve_region(data, wanted: str) -> str:
    """The workbook's own spelling of ``wanted``, or exit saying what it has."""
    choices = ["Global"] + list(data.REGIONS)
    for r in choices:
        if str(r).strip().lower() == wanted.strip().lower():
            return r
    raise SystemExit('\n'.join([
        f'--region: "{wanted}" is not a region in this workbook.',
        "  available: " + ", ".join(map(str, choices)),
        "  " + _geo_help(data)]))


def _resolve_country(data, wanted):
    """Same for --country; unset means the dashboard's own default."""
    if not wanted:
        return data.default_country()
    for c in data.ALL_COUNTRIES:
        if str(c).strip().lower() == wanted.strip().lower():
            return c
    have = (", ".join(map(str, data.ALL_COUNTRIES[:12])) + " ..."
            if data.ALL_COUNTRIES else "(none)")
    raise SystemExit('\n'.join([
        f'--country: "{wanted}" is not a country in this workbook.',
        "  available: " + have,
        "  " + _geo_help(data)]))


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Export report decks (the dashboard's four buttons).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__.split("\n\n", 1)[1])
    ap.add_argument("--out", required=True, metavar="DIR",
                    help="folder to write into (created if absent)")
    ap.add_argument("--data", metavar="XLSX",
                    help="workbook: a file name in the project folder or a "
                         "full path (default: MARKET_DATA, else Dataset_*)")
    ap.add_argument("--toc", metavar="DOCX",
                    help="the research team's Word table of contents: a file "
                         "name in the project folder or a full path "
                         "(default: REPORT_TOC, else TOC_*.docx)")
    ap.add_argument("--region", default="Global",
                    help="the region the single-deck options export "
                         "(default: Global)")
    ap.add_argument("--country", default=None,
                    help="country report to include (default: the "
                         "dashboard's own default)")
    ap.add_argument("--template", default="fmi", choices=("fmi", "factmr"),
                    help="visual template (default: fmi)")
    ap.add_argument("--only", default="all", metavar="LIST",
                    help="which options to run, e.g. 1,3 or 34 "
                         "(default: all four)")
    ap.add_argument("--zip-regions", default=None, metavar="LIST",
                    help="restrict the all-regions options (2 and 4) to "
                         "these regions, semicolon-separated, e.g. "
                         "\"North America;Europe\" (default: every region)")
    args = ap.parse_args()

    if args.data:
        os.environ["MARKET_DATA"] = args.data
    if args.toc:
        os.environ["REPORT_TOC"] = args.toc
    wanted = _parse_only(args.only)

    # Imported only now: config reads MARKET_DATA at import time.
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from fmi_report import config, data, theming
    from fmi_report.export_native import (export_pptx_native,
                                          export_regions_zip)
    from fmi_report.locking import (export_pptx_locked,
                                    export_regions_zip_locked)

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    tpl = theming.normalise(args.template)
    label = theming.LABELS[tpl]
    region = _resolve_region(data, args.region)
    country = _resolve_country(data, args.country)
    zip_regions = ([r.strip() for r in args.zip_regions.split(";") if r.strip()]
                   if args.zip_regions else None)
    market = data.MARKET_LABEL
    log_path = out_dir / "_export_log.txt"

    def say(msg: str) -> None:
        line = f"[{time.strftime('%H:%M:%S')}] {msg}"
        print(line, flush=True)
        with log_path.open("a", encoding="utf-8") as fh:
            fh.write(line + "\n")

    try:                       # a workbook whose numbers contradict each
        import check_workbook  # other builds a deck that cannot agree with
        warn = check_workbook.brief()   # itself - say so before, not after
        if warn:
            say(warn)
    except Exception:
        pass

    say(f"START  {market}  |  {config.DATA_FILE.name}  |  {label}"
        f"  |  region={region}  country={country}"
        + (f"  zip-regions={zip_regions}" if zip_regions else ""))

    # A single-deck export is named for its region, so exporting several
    # regions in turn fills a folder instead of overwriting one file.
    geo = "" if region == "Global" else f" - {region}"
    zip_geo = (" - All Regions" if not zip_regions
               else " - " + ", ".join(zip_regions))

    jobs = {
        "1": (f"{market} Report{geo} ({label}).pptx",
              lambda p: export_pptx_native(p, region=region,
                                           country=country, template=tpl)),
        "2": (f"{market}{zip_geo} ({label}).zip",
              lambda p: export_regions_zip(p, country=country, template=tpl,
                                           progress=say,
                                           regions=zip_regions)),
        "3": (f"{market} Report{geo} ({label}) (locked).pptx",
              lambda p: export_pptx_locked(p, region=region,
                                           country=country, template=tpl)),
        # The multi-deck options report each region as it lands: they run for
        # long enough that silence is indistinguishable from a hang.
        "4": (f"{market}{zip_geo} ({label}) (locked).zip",
              lambda p: export_regions_zip_locked(p, country=country,
                                                  template=tpl,
                                                  progress=say,
                                                  regions=zip_regions)),
    }

    failed = []
    for key in wanted:
        name, run = jobs[key]
        target = out_dir / f"{key} - {name}"
        say(f"{key}) {OPTIONS[key]} -> {target.name}")
        t0 = time.time()
        try:
            run(target)
        except Exception:
            failed.append(key)
            say(f"   FAILED after {time.time() - t0:.0f}s")
            traceback.print_exc()
            continue
        say(f"   done  {_mb(target):.2f} MB  {time.time() - t0:.0f}s")

    done = [k for k in wanted if k not in failed]
    say(f"FINISHED  {len(done)}/{len(wanted)} ok"
        + (f"  |  FAILED: {', '.join(failed)}" if failed else "")
        + f"  |  {out_dir}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
