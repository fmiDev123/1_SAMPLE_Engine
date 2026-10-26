"""One-command PowerPoint export.

    python export_ppt.py
    python export_ppt.py --region Europe --country Germany
    python export_ppt.py --report main            (or: country, both)
    python export_ppt.py --out "C:/path/My Report.pptx"

Starts its own temporary server, captures every report page at high
resolution, and writes one 16:9 slide per page - like Power BI's own
'Export to PowerPoint'.
"""
from __future__ import annotations

import argparse
from pathlib import Path

from fmi_report import data
from fmi_report.export import export_standalone

def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--region", default="Global",
                    help="Region slicer value for the main report")
    ap.add_argument("--country", default=None,
                    help="Country slicer value for the country report "
                         "(defaults to the largest market in the dataset)")
    ap.add_argument("--report", default="both",
                    choices=["main", "country", "both"],
                    help="Which report(s) to export")
    ap.add_argument("--mode", default="native",
                    choices=["native", "image"],
                    help="native = editable shapes/charts/tables "
                         "(default); image = page screenshots")
    ap.add_argument("--out", default=None, help="Output .pptx path")
    args = ap.parse_args()

    reports = ("main", "country") if args.report == "both" \
        else (args.report,)
    out = Path(args.out) if args.out else \
        Path(__file__).parent / f"{data.MARKET_LABEL} Report.pptx"
    if args.mode == "native":
        from fmi_report.export_native import export_pptx_native
        export_pptx_native(out, region=args.region, country=args.country,
                           reports=reports, progress=print)
    else:
        export_standalone(out, region=args.region, country=args.country,
                          reports=reports)

if __name__ == "__main__":
    main()
