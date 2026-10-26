"""Flexible Plastic Packaging - real report, unlocked: one deck per region
(region + all its countries) and one for Global.  No competitor section (the
workbook carries no company data).  Run: python export_flexible_plastic.py"""
import os, sys, time
from pathlib import Path
OUT = Path(__file__).resolve().parent / "Output" / "Flexible_Plastic_Packaging_kMT"
DEST = Path(os.environ.get("FPP_DEST", OUT))
DEST.mkdir(parents=True, exist_ok=True)
os.environ["MARKET_DATA"] = str(OUT / "Flexible_Plastic_Packaging_Market_Formatted_Dataset.xlsx")
os.environ["FMI_WRITEUP"] = "classic"
os.environ["FMI_INDUSTRY"] = "consumer"
sys.path.insert(0, str(Path(__file__).resolve().parent))
from fmi_report import data, theming
from fmi_report.export_native import export_pptx_native
from fmi_report import registry
# placeholder companies (no company data in the workbook): leave the slide out
registry.PAGES[:] = [p for p in registry.PAGES if p["name"] != "Competitor Pricing Strategies"]

REPORTS = ("main", "country", "dynamics")
which = sys.argv[1:] or ["Global"] + list(data.REGIONS)
for reg in which:
    ctys = ([] if reg == "Global" else
            [c for c in data.COUNTRIES.get(reg, []) if c in data.ALL_COUNTRIES])
    name = f"Flexible Plastic Packaging Market Report - {reg}.pptx"
    t = time.time()
    export_pptx_native(DEST / name, region=reg, country=(ctys[0] if ctys else None),
                       reports=REPORTS if ctys else ("main", "dynamics"),
                       countries=ctys or None, template="fmi")
    print(f"{name}  {(DEST/name).stat().st_size/1e6:.1f} MB  {time.time()-t:.0f}s", flush=True)
