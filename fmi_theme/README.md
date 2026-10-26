# fmi_theme

The FMI visual design system, packaged on its own so it can be installed into
any checkout of the report generator.

These are the only two files that `FMI trials report\matcha_report` has and
`Downloads\trials\matcha_report` does not - i.e. the whole of the FMI look:

```
fmi_theme/
    __init__.py                  re-exports the tokens; ASSETS / LOGO paths
    theme_fmi.py                 THE TEMPLATE (372 lines)
    assets/fmi_logo_fmi.png      FMI wordmark, from the template's slide master
install.py                       drop it into a checkout
README.md
```

## What's inside `theme_fmi.py`

| Group | Tokens |
|---|---|
| Colours | `FMI_NAVY` `FMI_BLUE` `FMI_CYAN` `FMI_AMBER` `FMI_TEAL` `FMI_RUST` `FMI_PURPLE` `FMI_MAROON` `FMI_OCHRE` `FMI_GOLD` `FMI_RULE` `FMI_EDGE` `FMI_INK`, `DATA_COLORS`, `KPI_ACCENTS`, semantic `GOOD`/`BAD`/`NEUTRAL` |
| Typography | `FONT_UI` (Calibri), `FONT_DISPLAY` (Bell MT), `FONT_STACK`, the `FS_*` scale |
| Page chrome | `HEADER_H` `RULE_Y/H` `TITLE_*` `LOGO_*` `FOOTER_*` `PAGE_NO_X` |
| Spacing scale | `GAP` `GAP_X` `PAGE_MARGIN` `CARD_PAD_*` `PANEL_PAD` `HEADER_CLEAR` `FOOTER_CLEAR` `PANEL_BORDER` |
| Component metrics | `line_px` `wrap_lines` `kpi_height` `kpi_value_size` `panel_title_height` `summary_height` `table_row_h` `table_head_h` |
| Growth policy | `STRETCH` `stretch_factor` `MAX_ROW_GAP` |

## The contract

The layout engine reserves a box for every block; the two renderers
(`components.py` for Dash, `export_native.py` for PowerPoint) draw into it.

> **No renderer may invent a padding, a line height, a font size or a column
> width** - it calls a helper here, so the drawn size always equals the
> reserved size.

That is what keeps the web page and the PowerPoint slide in agreement. Each of
the three files used to carry its own copy of these numbers, which is exactly
how components ended up overlapping.

**Units:** the canvas is 1280×720 **px**, the deck 960×540 **pt** - the same
16:9 page at 1 px = 0.75 pt (`PT_PER_PX`). Geometry tokens are in px; the
export converts.

## Install into another checkout

```bash
python install.py "C:/Users/Shweta/Downloads/trials/matcha_report"
```

It copies `theme_fmi.py` into `fmi_report/` and the logo into
`fmi_report/assets/` (backing up anything it replaces as `.bak`), then prints
the one edit needed in that checkout's `config.py`: replace its own
colour/typography/geometry block with

```python
from .theme_fmi import *
```

keeping that file's dataset discovery, years, unit labels and geo map.

Verify with:

```bash
python -c "from fmi_report.app import create_app; create_app()"
```

## Use the tokens directly

```python
from fmi_theme import theme_fmi as T
T.FMI_NAVY            # '#2A5F8F'
T.kpi_height("Market Size (2026)", "estimate year", 300)
```

## Re-skinning for a different brand

Copy `theme_fmi.py`, change the values, point `config.py`'s re-export at the
new module. Nothing else in the generator moves - that is the entire reason
this is a separate module.

## Provenance

Extracted from `New_Global XYZ Market_Report_Sample_2026_Future market
Insights_B2B.pptx` - its theme colour scheme, slide master, layouts, and pixel
samples of rendered slides.
