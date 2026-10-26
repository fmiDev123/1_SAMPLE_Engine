"""Run export_ppts.py for a VALUE-ONLY workbook (no volume rows).

    python build_value_only.py --data d_ind.xlsx --toc toc.docx ... (the usual export_ppts.py flags)

The engine itself is untouched.  This wrapper only takes the pages that are
about volume out of the page list in memory before the export starts, and only
when the workbook really has no volume - otherwise it runs export_ppts.py
exactly as it is.  Without it a value-only workbook stops on the country
"Volume Overview" page (a CAGR chart with nothing to draw).
"""
import os
import runpy
import sys
from functools import partial
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))


def _arg(flag):
    a = sys.argv
    return a[a.index(flag) + 1] if flag in a else None


for flag, env in (("--data", "MARKET_DATA"), ("--toc", "REPORT_TOC"),
                  ("--companies", "REPORT_COMPANIES")):
    v = _arg(flag)
    if v:
        os.environ[env] = v

from fmi_report import data, registry            # noqa: E402


def _has_volume():
    try:
        for g in data.REGIONS:
            s = data.total_series(g, "Volume")
            if float(s.fillna(0).abs().sum()) > 0:
                return True
    except Exception:
        return True                  # cannot tell: leave the pages alone
    return False


def _about_volume(page):
    fn = page.get("fn") or page.get("func") or page.get("render")
    if isinstance(fn, partial) and "Volume" in fn.args:
        return True
    return "volume" in str(page.get("name", "")).lower()


if not _has_volume():
    keep = [p for p in registry.PAGES if not _about_volume(p)]
    print(f"value-only workbook: {len(registry.PAGES) - len(keep)} volume "
          "page(s) left out", flush=True)
    registry.PAGES[:] = keep

sys.argv = ["export_ppts.py"] + sys.argv[1:]
runpy.run_path(str(HERE / "export_ppts.py"), run_name="__main__")
