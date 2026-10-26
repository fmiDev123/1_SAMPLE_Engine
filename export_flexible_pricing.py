"""Pricing assessment slides only: Global, then each region (its countries are
benchmarked inside the region's pages).  One file."""
import os, sys
from pathlib import Path
OUT = Path(__file__).resolve().parent / "Output" / "Flexible_Plastic_Packaging_kMT"
DEST = Path(os.environ.get("FPP_DEST", OUT))
DEST.mkdir(parents=True, exist_ok=True)
os.environ["MARKET_DATA"] = str(OUT / "Flexible_Plastic_Packaging_Market_Formatted_Dataset.xlsx")
os.environ["FMI_WRITEUP"] = "classic"
os.environ["FMI_INDUSTRY"] = "consumer"
sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx import Presentation
from fmi_report import data, geo, theming
from fmi_report import export_native as X
from fmi_report.registry import PAGES, MAIN_REPORT, SEC_PRICING

pages = [p for p in PAGES if p['name'] != 'Competitor Pricing Strategies' and p["report"] == MAIN_REPORT and p.get("section") == SEC_PRICING]
with theming.use("fmi"):
    prs = Presentation(); prs.slide_width, prs.slide_height = X.SLIDE_W, X.SLIDE_H
    with geo.active("Global"):
        X._title_slide(prs, "Global", None)
    for n, g in enumerate(["Global"] + list(data.REGIONS), 1):
        with geo.active(g):
            if g != "Global":
                prs.slides  # dividers: reuse the engine's own
            X._SUBHEADER["text"] = ""
            slides = [X._section_divider_slide(prs, n, f"{g} Pricing Assessment",
                      [p["name"] for p in pages], kicker=geo.label(MAIN_REPORT))]
            for p in pages:
                from fmi_report import subheaders
                X._SUBHEADER["text"] = subheaders.for_page(p, g, None)
                try:
                    X._flow_slides(prs, p, g, None)
                finally:
                    X._SUBHEADER["text"] = ""
    X._renumber_sections(prs)
    X._tidy_deck(prs); X._enforce_content_floor(prs); X._install_theme(prs)
    sys.setrecursionlimit(10000)
    out = DEST / "Flexible Plastic Packaging Market - Pricing Assessment (Global and Regions).pptx"
    prs.save(str(out)); print(out, len(prs.slides._sldIdLst), "slides")
