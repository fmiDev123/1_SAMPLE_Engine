"""Export the report decks from the command line.
The dashboard's four export buttons, without the dashboard:

    1  this view          the geography selected, one deck per geography
    2  all regions        one deck per region, zipped
    3  this view locked    (1) with the protected values withheld
    4  all regions locked  (2) with the protected values withheld

Each option calls exactly what the matching button calls, with the same
arguments, so a deck built here is the deck the browser would have sent.

    python export_ppts.py --out "C:\\Users\\Me\\Downloads\\exports"
    python export_ppts.py --out DIR --region "North America" --only 3,4
    python export_ppts.py --data Dataset_Other.xlsx --out DIR --only 1

Samples for any mix of geographies (options 1 and 3).  --region and --country
take one name, a ';'-separated list, or ``all``; --level picks the depth:

    --level region    the region's own report only, no country pages
    --level country   country pages only, no region report
    --level both      region report + country pages (default)

    # one global deck, no country pages
    python export_ppts.py --out DIR --only 1 --region Global --level region
    # several regions, region level only: one deck each
    python export_ppts.py --out DIR --only 1 --region "Europe;Asia Pacific" --level region
    # several countries in one deck, country level only
    python export_ppts.py --out DIR --only 1 --country "Germany;France;Japan" --level country
    # the same countries, one deck per country
    python export_ppts.py --out DIR --only 1 --country "Germany;France" --level country --per-country
    # every country in Europe, country level only
    python export_ppts.py --out DIR --only 1 --region Europe --country all --level country
    # what the workbook offers / what would be built, without building
    python export_ppts.py --data X.xlsx --list
    python export_ppts.py --data X.xlsx --out DIR --only 1 --region all --level region --dry-run

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

# --level -> the `reports` argument of the exporters.  "both" is None so the
# exporters keep their own defaults (single deck: all four reports; ZIP: main
# + country), i.e. exactly what a bare run always built.
LEVEL_REPORTS = {
    "both": None,
    "region": ("main", "dynamics", "competitor"),
    "country": ("country",),
}
ZIP_LEVEL_REPORTS = {
    "both": None,
    "region": ("main",),
    "country": ("country",),
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


def _resolve_country(data, wanted: str):
    """Same for --country."""
    for c in data.ALL_COUNTRIES:
        if str(c).strip().lower() == wanted.strip().lower():
            return c
    have = (", ".join(map(str, data.ALL_COUNTRIES[:12])) + " ..."
            if data.ALL_COUNTRIES else "(none)")
    raise SystemExit('\n'.join([
        f'--country: "{wanted}" is not a country in this workbook.',
        "  available: " + have,
        "  (run with --list to see every region and country)",
        "  " + _geo_help(data)]))


# ── multi-geography selection ─────────────────────────────────────────
def _split(raw) -> list[str]:
    """"A;B" / "A|B" -> ['A', 'B'].  With no ';' or '|' it splits on commas,
    so "Germany, France" works and a single name is left whole."""
    if raw is None:
        return []
    raw = str(raw)
    if ";" in raw or "|" in raw:
        parts = raw.replace("|", ";").split(";")
    else:
        parts = raw.split(",")
    return [t.strip() for t in parts if t.strip()]


def _resolve_regions(data, raw) -> list[str]:
    """--region -> the workbook's spellings, in the order given.  ``all`` is
    Global plus every region."""
    toks = _split(raw) or ["Global"]
    if any(t.lower() == "all" for t in toks):
        return ["Global"] + list(data.REGIONS)
    out = []
    for t in toks:
        r = _resolve_region(data, t)
        if r not in out:
            out.append(r)
    return out


ALL = ["*"]            # --country all: decided per region in _plan


def _resolve_countries(data, raw):
    """--country -> workbook spellings; ``all`` -> ALL; unset -> None."""
    toks = _split(raw)
    if not toks:
        return None
    if any(t.lower() == "all" for t in toks):
        return ALL
    out = []
    for t in toks:
        c = _resolve_country(data, t)
        if c not in out:
            out.append(c)
    return out


def _countries_of(data, region: str) -> list[str]:
    if region == "Global":
        return list(data.ALL_COUNTRIES)
    return [c for c in data.COUNTRIES.get(region, [])
            if c in data.ALL_COUNTRIES]


def _plan(data, regions, countries, level, per_country):
    """The decks to build: ([(region, [countries]), ...], [unplaced countries]).

    A deck is a region's report plus country pages for some countries.  A
    country goes into the deck of the region it sits in (under Global, every
    chosen country is in).  With no --country, ``both`` keeps the long-standing
    behaviour - the region's lead market - and ``country`` takes every country
    in the region.  Countries that fit none of the selected regions come back
    as the second value so the caller can say so, never silently drop them.
    """
    decks: list[tuple[str, list[str]]] = []
    placed: set[str] = set()
    for reg in regions:
        pool = _countries_of(data, reg)
        if countries is None:
            ctys = (pool if level == "country"
                    else [data.default_country(reg)])
        elif countries == ALL:
            ctys = pool
        else:
            ctys = [c for c in countries if c in pool]
        if level == "region":
            ctys = []
        placed.update(ctys)
        if level == "country" and not ctys:
            continue                       # nothing to show for this region
        if per_country and ctys:
            # under Global a lone country's own region heads its deck, so
            # "Germany" is a Europe/Germany sample, not a global one
            decks += [(reg if reg != "Global" else (data.region_of(c) or reg),
                       [c]) for c in ctys]
        else:
            decks.append((reg, ctys))
    unplaced = ([c for c in countries if c not in placed]
                if countries and countries != ALL and level != "region" else [])
    return decks, unplaced


def _deck_geo(region: str, ctys: list[str], level: str) -> str:
    """File-name part: '', 'Europe', 'Germany', 'Europe - Germany, France'."""
    shown = ""
    if ctys and level != "region":
        shown = (", ".join(ctys) if len(ctys) <= 3
                 else f"{len(ctys)} Countries")
    if level == "country" and shown:
        return shown if len(ctys) == 1 else f"{region} - {shown}"
    parts = [p for p in ("" if region == "Global" else region, shown) if p]
    return " - ".join(parts)


def _print_geographies(data) -> None:
    print(f"{data.MARKET_LABEL}")
    print("Regions : Global" + "".join(f", {r}" for r in data.REGIONS))
    if not data.REGIONS:
        print("          " + _geo_help(data))
    for r in data.REGIONS:
        cs = [c for c in data.COUNTRIES.get(r, []) if c in data.ALL_COUNTRIES]
        print(f"  {r} ({len(cs)}): {', '.join(cs) if cs else '-'}")


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Export report decks (the dashboard's four buttons).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__.split("\n\n", 1)[1])
    ap.add_argument("--out", default=None, metavar="DIR",
                    help="folder to write into (created if absent); "
                         "required unless --list")
    ap.add_argument("--data", metavar="XLSX",
                    help="workbook: a file name in the project folder or a "
                         "full path (default: MARKET_DATA, else Dataset_*)")
    ap.add_argument("--segmentation", metavar="XLSX",
                    help="a Segmentation brief (Parameter | Value | "
                         "Segmentation | Companies): the dataset is generated "
                         "from it first (indicative figures pinned to its "
                         "2026 / 2036 value and CAGR, standard regional "
                         "taxonomy when it names no regions), then the deck "
                         "is built from that")
    ap.add_argument("--toc", metavar="DOCX",
                    help="the research team's Word table of contents: a file "
                         "name in the project folder or a full path "
                         "(default: REPORT_TOC, else TOC_*.docx)")
    ap.add_argument("--companies", metavar="XLSX",
                    help="workbook whose 'Competition' sheet lists the "
                         "companies for the Company Coverage page, used when "
                         "the dataset and the TOC name none (default: any "
                         "other .xlsx beside the dataset)")
    ap.add_argument("--region", default="Global",
                    help="region(s) for the single-deck options: one name, a "
                         "';'-separated list, or 'all' (Global + every "
                         "region).  One deck per region.  (default: Global)")
    ap.add_argument("--country", default=None,
                    help="country(ies) for the country pages: one name, a "
                         "';'-separated list, or 'all'.  Each goes into the "
                         "deck of its own region.  (default: each region's "
                         "lead market)")
    ap.add_argument("--level", default="both",
                    choices=("region", "country", "both"),
                    help="region = the region/global report only; country = "
                         "country pages only; both = region report + country "
                         "pages (default)")
    ap.add_argument("--indicative", action="store_true",
                    help="the data is example data: a locked deck withholds "
                         "EVERY figure (CAGRs, indices, shares too).  Set "
                         "automatically for a dataset made by "
                         "fmi_indicative_numbers.py")
    ap.add_argument("--one-deck", action="store_true",
                    help="with several regions (--region \"Global;North "
                         "America\"): ONE file - the common opening, each "
                         "region's market outlook in turn, then the country "
                         "outlook (--country), then pricing, dynamics, supply "
                         "chain and competition once")
    ap.add_argument("--per-country", action="store_true",
                    help="one deck per country instead of one per region")
    ap.add_argument("--list", action="store_true",
                    help="print the regions and countries in the workbook "
                         "and exit")
    ap.add_argument("--dry-run", action="store_true",
                    help="print the decks that would be built and exit")
    ap.add_argument("--writeup", default="classic", choices=("classic", "asis"),
                    help="commentary wording: 'classic' = the 2015-style "
                         "templates (fmi_templates, FMI_Classic_Report_"
                         "Templates.xlsx), 'asis' = the earlier wording "
                         "(default: classic)")
    ap.add_argument("--industry", default=None,
                    choices=("food", "consumer", "healthcare", "pharma",
                             "industrial"),
                    help="industry of the market; lets the classic text use "
                         "its driver phrases (left out when not given)")
    ap.add_argument("--template", default="fmi",
                    choices=("fmi", "factmr", "tmr"),
                    help="visual template: fmi, factmr or tmr "
                         "(Transparency Market Research) (default: fmi)")
    ap.add_argument("--only", default="all", metavar="LIST",
                    help="which options to run, e.g. 1,3 or 34 "
                         "(default: all four)")
    ap.add_argument("--zip-regions", default=None, metavar="LIST",
                    help="restrict the all-regions options (2 and 4) to "
                         "these regions, semicolon-separated, e.g. "
                         "\"North America;Europe\" (default: every region)")
    args = ap.parse_args()

    if not args.out and not args.list:
        ap.error("--out is required (unless --list)")
    if args.segmentation:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        import fmi_indicative_numbers as _ind
        seg = Path(args.segmentation)
        if not seg.exists():
            seg = Path(__file__).resolve().parent / args.segmentation
        built = _ind.build_from_segmentation(seg)
        print(f"segmentation -> dataset: {built}", flush=True)
        args.data = str(built)
        comp = built.with_name(built.name.replace(
            "_Formatted_Dataset_ind.xlsx", "_Companies.xlsx"))
        if comp.exists() and not args.companies:
            args.companies = str(comp)
    if args.data:
        os.environ["MARKET_DATA"] = args.data
    if args.toc:
        os.environ["REPORT_TOC"] = args.toc
    if args.companies:
        os.environ["REPORT_COMPANIES"] = args.companies
    if args.indicative:
        os.environ["FMI_INDICATIVE"] = "1"
    os.environ["FMI_WRITEUP"] = args.writeup
    if args.industry:
        os.environ["FMI_INDUSTRY"] = args.industry
    wanted = _parse_only(args.only)

    # Imported only now: config reads MARKET_DATA at import time.
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from fmi_report import config, data, theming
    from fmi_report.export_native import (export_pptx_native,
                                          export_regions_zip)
    from fmi_report.locking import (export_pptx_locked,
                                    export_regions_zip_locked)

    if args.list:
        _print_geographies(data)
        return 0

    out_dir = Path(args.out)
    tpl = theming.normalise(args.template)
    label = theming.LABELS[tpl]
    regions = _resolve_regions(data, args.region)
    countries = _resolve_countries(data, args.country)
    level = args.level
    decks, unplaced = _plan(data, regions, countries, level, args.per_country)
    if unplaced:
        print(f"note: not in the selected region(s), left out: "
              f"{', '.join(unplaced)}", file=sys.stderr)
    single = [k for k in wanted if k in ("1", "3")]
    if single and not decks:
        raise SystemExit("nothing to build: no selected country sits in the "
                         "selected region(s).  Use --list to see the "
                         "geography, or --region all.")
    zip_regions = ([r.strip() for r in args.zip_regions.split(";") if r.strip()]
                   if args.zip_regions else None)
    market = data.MARKET_LABEL
    reports = LEVEL_REPORTS[level]
    zip_kw = ({} if ZIP_LEVEL_REPORTS[level] is None
              else {"reports": ZIP_LEVEL_REPORTS[level]})
    zip_geo = (" - All Regions" if not zip_regions
               else " - " + ", ".join(zip_regions))

    if args.dry_run:
        print(f"{market}  level={level}  template={label}")
        for key in single:
            for reg, ctys in decks:
                print(f"  {key}) {market} Report"
                      f"{' - ' + g if (g := _deck_geo(reg, ctys, level)) else ''}"
                      f"  [region={reg}  countries={', '.join(ctys) or '-'}]")
        for key in (k for k in wanted if k in ("2", "4")):
            print(f"  {key}) {market}{zip_geo}.zip  [{OPTIONS[key]}]")
        return 0

    out_dir.mkdir(parents=True, exist_ok=True)
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
        f"  |  level={level}  regions={', '.join(regions)}"
        f"  countries={'all' if countries == ALL else ', '.join(countries or ['(default)'])}"
        + (f"  zip-regions={zip_regions}" if zip_regions else ""))

    def single_job(key, reg, ctys):
        """One deck: (file name, builder) for option 1 or 3 and one plan row."""
        geo = _deck_geo(reg, ctys, level)
        geo = f" - {geo}" if geo else ""
        kw = dict(region=reg, country=(ctys[0] if ctys else None),
                  countries=ctys or None, template=tpl)
        if reports is not None:
            kw["reports"] = reports
        if key == "1":
            return (f"{market} Report{geo} ({label}).pptx",
                    lambda p: export_pptx_native(p, **kw))
        return (f"{market} Report{geo} ({label}) (locked).pptx",
                lambda p: export_pptx_locked(p, **kw))

    # (output file name, builder) in the order they run
    def one_deck_job(key):
        """ONE deck holding every listed region's outlook, then the
        countries'."""
        regs = [r for r in regions if r != "Global" or len(regions) > 1] \
            or regions
        regs = list(regions)
        if countries == ALL:
            ctys = []
            for r_ in regs:
                for c_ in _countries_of(data, r_):
                    if c_ not in ctys:
                        ctys.append(c_)
        elif countries:
            ctys = list(countries)
        else:
            ctys = _countries_of(data, regs[-1])
        label_ = ", ".join(regs + list(ctys))
        if len(label_) > 90:
            label_ = (f"{regs[0]} + {len(regs) - 1} Regions + "
                      f"{len(ctys)} Countries" if regs[0] == "Global"
                      else f"{len(regs)} Regions + {len(ctys)} Countries")
        kw = dict(region=regs[0], country=(ctys[0] if ctys else None),
                  countries=ctys or None, template=tpl, outlooks=regs)
        if key == "1":
            return (f"{market} Report - {label_} ({label}).pptx",
                    lambda p: export_pptx_native(p, **kw))
        return (f"{market} Report - {label_} ({label}) (locked).pptx",
                lambda p: export_pptx_locked(p, **kw))

    jobs = []
    for key in wanted:
        if key in ("1", "3") and args.one_deck and len(regions) > 1:
            name, run = one_deck_job(key)
            jobs.append((key, name, run))
        elif key in ("1", "3"):
            for reg, ctys in decks:
                name, run = single_job(key, reg, ctys)
                jobs.append((key, name, run))
        elif key == "2":
            jobs.append((key, f"{market}{zip_geo} ({label}).zip",
                         lambda p: export_regions_zip(
                             p, country=None, template=tpl, progress=say,
                             regions=zip_regions, **zip_kw)))
        else:
            # The multi-deck options report each region as it lands: they run
            # for long enough that silence is indistinguishable from a hang.
            jobs.append((key, f"{market}{zip_geo} ({label}) (locked).zip",
                         lambda p: export_regions_zip_locked(
                             p, country=None, template=tpl, progress=say,
                             regions=zip_regions, **zip_kw)))

    failed = []
    for n, (key, name, run) in enumerate(jobs, 1):
        target = out_dir / f"{key} - {_safe_name(name)}"
        say(f"[{n}/{len(jobs)}] {key}) {OPTIONS[key]} -> {target.name}")
        t0 = time.time()
        try:
            run(target)
        except Exception:
            failed.append(f"{key}:{target.name}")
            say(f"   FAILED after {time.time() - t0:.0f}s")
            traceback.print_exc()
            continue
        say(f"   done  {_mb(target):.2f} MB  {time.time() - t0:.0f}s")

    say(f"FINISHED  {len(jobs) - len(failed)}/{len(jobs)} ok"
        + (f"  |  FAILED: {', '.join(failed)}" if failed else "")
        + f"  |  {out_dir}")
    return 1 if failed else 0


def _safe_name(name: str) -> str:
    """A file name Windows accepts."""
    return "".join("_" if ch in '<>:"/\\|?*' else ch for ch in name).strip()


if __name__ == "__main__":
    raise SystemExit(main())
