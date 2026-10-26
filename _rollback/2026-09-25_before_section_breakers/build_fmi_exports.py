"""Build report exports from the command line - no browser, no clicking.

Every job here is exactly what the matching sidebar button produces; only the
way you ask for it differs.  With no options it keeps its original behaviour:
all four FMI exports for Global.

    set MARKET_DATA=DataSet_Commercial Microwave Ovens Market.xlsx
    python build_fmi_exports.py                     all four, Global, FMI
    python build_fmi_exports.py --what view         just "export this view"
    python build_fmi_exports.py --what view --region Europe
    python build_fmi_exports.py --what locked --region Europe
    python build_fmi_exports.py --what regions      one deck per region, ZIP
    python build_fmi_exports.py --what locked-regions
    python build_fmi_exports.py --what view locked  two of them, one run
    python build_fmi_exports.py --template factmr   the other corporate look
    python build_fmi_exports.py --list              what may be typed where

Files land in ./exports (override with --out).  They are written one at a
time, so a later failure still leaves the earlier files on disk.
"""
import argparse
import sys
import time
import traceback
from pathlib import Path

from fmi_report import data, theming
from fmi_report.export_native import export_pptx_native, export_regions_zip
from fmi_report.locking import export_pptx_locked, export_regions_zip_locked

# The four jobs, under the names the sidebar buttons use.  "all" runs them in
# this order.
WHATS = ["view", "regions", "locked", "locked-regions"]
TITLES = {
    "view":           "this view, unlocked",
    "regions":        "all regions, unlocked",
    "locked":         "this view, LOCKED",
    "locked-regions": "all regions, LOCKED",
}
HERE = Path(__file__).resolve().parent


def parse_args(argv=None):
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--what", nargs="+", default=["all"],
                    choices=["all"] + WHATS,
                    help="which export(s) to build (default: all)")
    ap.add_argument("--region", default="Global",
                    help='region for the single-deck exports, e.g. "Europe" '
                         '(default: Global).  Ignored by the ZIP exports, '
                         'which always cover every region.')
    ap.add_argument("--country", default=None,
                    help="country for the country-report pages "
                         "(default: the market's largest, as in the app)")
    ap.add_argument("--template", default="fmi",
                    help='corporate look: "fmi" or "factmr" (default: fmi)')
    ap.add_argument("--out", default=None,
                    help="folder to write into (default: ./exports)")
    ap.add_argument("--list", action="store_true",
                    help="print the regions and countries this workbook "
                         "offers, then exit without building anything")
    return ap.parse_args(argv)


def resolve(args):
    """Turn the typed options into validated values, or exit with advice."""
    # theming.normalise() falls back to the default for anything it does not
    # recognise, so a typo would silently build the WRONG look -- match here
    # instead, ignoring case and punctuation ("Fact.MR" == "factmr").
    def key(s):
        return "".join(ch for ch in s.lower() if ch.isalnum())
    wanted = key(args.template)
    tpl = next((name for name in theming.TEMPLATES
                if wanted in (key(name), key(theming.LABELS[name]))), None)
    if tpl is None:
        sys.exit(f'--template: "{args.template}" is not a template.  '
                 f'Choose one of: {", ".join(theming.TEMPLATES)}')

    regions = ["Global"] + list(data.REGIONS)
    region = next((r for r in regions
                   if r.lower() == args.region.strip().lower()), None)
    if region is None:
        sys.exit(f'--region: "{args.region}" is not in this workbook.  '
                 f'Choose one of: {", ".join(regions)}')

    if args.country:
        country = next((c for c in data.ALL_COUNTRIES
                        if c.lower() == args.country.strip().lower()), None)
        if country is None:
            sys.exit(f'--country: "{args.country}" is not in this workbook.  '
                     f'Run  python build_fmi_exports.py --list  to see them.')
    else:
        country = data.default_country()

    whats = WHATS if "all" in args.what else \
        [w for w in WHATS if w in args.what]      # keep the canonical order
    out = Path(args.out).expanduser() if args.out else HERE / "exports"
    return whats, region, country, tpl, out


def show_list():
    print(f"market   : {data.MARKET_LABEL}")
    print(f"templates: {', '.join(theming.TEMPLATES)}")
    print(f"regions  : Global, {', '.join(data.REGIONS)}")
    print(f"default country: {data.default_country()}")
    print("countries:")
    for r in data.REGIONS:
        names = data.COUNTRIES.get(r, [])
        if names:
            print(f"  {r}: {', '.join(names)}")


def jobs_for(whats, region, country, tpl, out):
    """(label, build-function, output path) per requested export."""
    label = theming.LABELS[tpl]
    M = data.MARKET_LABEL
    # Global keeps the original file names; a region export says which one it
    # is, so two regions cannot overwrite each other.
    tag = "" if region == "Global" else f" - {region}"
    built = {
        "view": (lambda p: export_pptx_native(p, region=region,
                                              country=country, template=tpl),
                 f"{M} Report ({label}){tag}.pptx"),
        "regions": (lambda p: export_regions_zip(p, country=country,
                                                 template=tpl),
                    f"{M} - All Regions ({label}).zip"),
        "locked": (lambda p: export_pptx_locked(p, region=region,
                                                country=country,
                                                template=tpl),
                   f"{M} Report ({label}){tag} (locked).pptx"),
        "locked-regions": (lambda p: export_regions_zip_locked(
                               p, country=country, template=tpl),
                           f"{M} - All Regions ({label}) (locked).zip"),
    }
    n = len(whats)
    return [(f"{i}/{n}  {TITLES[w]}", built[w][0], out / built[w][1])
            for i, w in enumerate(whats, 1)]


def main(argv=None) -> int:
    args = parse_args(argv)
    if args.list:
        show_list()
        return 0
    whats, region, country, tpl, out = resolve(args)
    out.mkdir(parents=True, exist_ok=True)

    print(f"market  : {data.MARKET_LABEL}")
    print(f"template: {theming.LABELS[tpl]} | region: {region} | "
          f"country: {country}")
    print(f"folder  : {out}\n", flush=True)

    done, failed = [], []
    t_all = time.time()
    for label, fn, path in jobs_for(whats, region, country, tpl, out):
        print(f"[{label}] building {path.name} ...", flush=True)
        t = time.time()
        try:
            fn(path)
            mb = path.stat().st_size / 1e6
            print(f"[{label}] OK  {time.time() - t:6.0f}s  {mb:6.2f} MB  "
                  f"{path.name}\n", flush=True)
            done.append((path.name, mb))
        except Exception as exc:
            print(f"[{label}] FAILED after {time.time() - t:.0f}s: "
                  f"{type(exc).__name__}: {exc}", flush=True)
            traceback.print_exc()
            failed.append((path.name, f"{type(exc).__name__}: {exc}"))

    print("=" * 66)
    print(f"finished in {(time.time() - t_all) / 60:.1f} min")
    for name, mb in done:
        print(f"  OK      {mb:6.2f} MB  {name}")
    for name, err in failed:
        print(f"  FAILED            {name}  <- {err}")
    print(f"\nfolder: {out}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
