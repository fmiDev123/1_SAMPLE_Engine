"""Install the FMI theme into a checkout of the report generator.

    python install.py "C:/Users/Shweta/Downloads/trials/matcha_report"

Copies:
    fmi_theme/theme_fmi.py            -> <target>/fmi_report/theme_fmi.py
    fmi_theme/assets/fmi_logo_fmi.png -> <target>/fmi_report/assets/

...then tells you the one edit needed in that checkout's `config.py` to switch
it over.  Existing files are backed up next to themselves with a .bak suffix,
so a bad install is reversible without git.

Nothing is overwritten silently and nothing is deleted.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SRC = HERE / "fmi_theme"

RE_EXPORT = '''# ═══════════════════════════════════════════════════════════════════════
#  VISUAL DESIGN SYSTEM - re-exported from theme_fmi.py
# ═══════════════════════════════════════════════════════════════════════
from .theme_fmi import *                                          # noqa: F401,F403
'''


def install(target: Path) -> int:
    pkg = target / "fmi_report"
    if not pkg.is_dir():
        print(f"!! {target} does not look like a checkout "
              f"(no fmi_report/ inside)")
        return 1
    assets = pkg / "assets"
    assets.mkdir(exist_ok=True)

    jobs = [(SRC / "theme_fmi.py", pkg / "theme_fmi.py"),
            (SRC / "assets" / "fmi_logo_fmi.png",
             assets / "fmi_logo_fmi.png")]
    for src, dst in jobs:
        if dst.exists():
            backup = dst.with_suffix(dst.suffix + ".bak")
            shutil.copy2(dst, backup)
            print(f"   backed up {dst.name} -> {backup.name}")
        shutil.copy2(src, dst)
        print(f"   installed {dst.relative_to(target)}")

    cfg = pkg / "config.py"
    already = cfg.exists() and "theme_fmi" in cfg.read_text(encoding="utf8")
    print(f"\nInstalled into {target}")
    if already:
        print("config.py already imports theme_fmi - nothing else to do.")
    else:
        print("ONE edit left, in fmi_report/config.py: delete that file's own\n"
              "colour / typography / geometry block and put this in its place\n"
              "(keep the dataset discovery, years, unit labels and geo map):\n")
        print(RE_EXPORT)
        print("Then check it loads:  python -c \"from fmi_report.app import "
              "create_app; create_app()\"")
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(install(Path(sys.argv[1]).resolve()))
