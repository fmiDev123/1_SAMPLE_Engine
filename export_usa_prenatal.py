"""USA Prenatal Vitamin Supplements - real report, unlocked.  One deck (USA is
the only geography).  Reasons come from reasons_usa.json (FMI_REASONS)."""
import os, sys, time
from pathlib import Path
OUT = Path(__file__).resolve().parent / "Output" / "USA_Prenatal"
DEST = Path(os.environ.get("FPP_DEST", OUT)); DEST.mkdir(parents=True, exist_ok=True)
os.environ["MARKET_DATA"] = str(OUT / "USA_Prenatal_Vitamin_Supplements_Market_Formatted_Dataset.xlsx")
os.environ["FMI_WRITEUP"] = "classic"
sys.path.insert(0, str(Path(__file__).resolve().parent))
from fmi_report import registry
registry.PAGES[:] = [p for p in registry.PAGES if p["name"] != "Competitor Pricing Strategies"]
from fmi_report.export_native import export_pptx_native
name = "USA Prenatal Vitamin Supplements Market Report.pptx"
t = time.time()
export_pptx_native(DEST / name, region="USA", country="USA", reports=("main", "dynamics"), template="fmi")
print(f"{name} {(DEST/name).stat().st_size/1e6:.1f} MB {time.time()-t:.0f}s", flush=True)
