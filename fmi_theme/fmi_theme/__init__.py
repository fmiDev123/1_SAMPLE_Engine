"""fmi_theme - the FMI visual design system as a portable package.

This is everything that makes a generated report *look* like an FMI report,
packaged so it can be dropped into any checkout of the report generator:

    theme_fmi.py            colours, typography, page chrome, spacing scale,
                            component metrics, growth policy
    assets/fmi_logo_fmi.png the FMI wordmark, lifted from the template's
                            slide master (transparent PNG, 1.795:1)

WHY THIS EXISTS
    The generator has several checkouts on this machine.  Only one of them
    carried the FMI look, because the design lived inside that copy's
    `config.py`.  Extracting it here makes the template a unit you can install
    into any copy (see `install.py`) or version on its own.

USING IT
    Either install it into a checkout:

        python install.py "C:/path/to/matcha_report"

    ...which copies `theme_fmi.py` into `fmi_report/` and the logo into
    `fmi_report/assets/`, or import the tokens directly:

        from fmi_theme import theme_fmi as T
        T.FMI_NAVY, T.FS_PAGE_TITLE, T.kpi_height(...)

THE CONTRACT
    The layout engine reserves a box for every block; the renderers draw into
    it.  No renderer may invent a padding, a line height, a font size or a
    column width - it calls a helper here, so the drawn size always equals the
    reserved size.  That is what keeps the web page and the PowerPoint slide
    in agreement.
"""
from __future__ import annotations

from pathlib import Path

from . import theme_fmi
from .theme_fmi import *                                    # noqa: F401,F403

__all__ = ["theme_fmi", "ASSETS", "LOGO", "version"]

ASSETS = Path(__file__).resolve().parent / "assets"
LOGO = ASSETS / "fmi_logo_fmi.png"


def version() -> str:
    """Short description of what this theme is and where it came from."""
    return ("FMI theme - extracted from 'New_Global XYZ Market_Report_Sample_"
            "2026_Future market Insights_B2B.pptx'")
