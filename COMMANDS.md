# FMI_PPT - Command Reference (CMD / PowerShell)

All commands are run from the project folder.

```
cd "C:\Users\Shweta\OneDrive - Future Market Insights Pvt.Ltd\Documents\FMI_PPT"
```

`DIR` = output folder (created if missing). `X.xlsx` = your `Dataset_*.xlsx` workbook.

---

## 1. Main exporter - `export_ppts.py`

Builds the four exports:

| # | Export |
|---|---|
| 1 | This view (one deck per region / country) |
| 2 | All regions (ZIP, one deck per region) |
| 3 | This view, locked (numbers blurred, padlocks) |
| 4 | All regions, locked (ZIP) |

### Basic

```
python export_ppts.py --out "DIR"
python export_ppts.py --data "X.xlsx" --out "DIR"
python export_ppts.py --help
```

### Choose which exports

```
python export_ppts.py --out "DIR" --only 1
python export_ppts.py --out "DIR" --only 3
python export_ppts.py --out "DIR" --only 1,3
python export_ppts.py --out "DIR" --only 3,4
python export_ppts.py --out "DIR" --only 4
```

### Regions and countries

```
python export_ppts.py --out "DIR" --region "North America" --only 3
python export_ppts.py --out "DIR" --region "Europe;East Asia" --level region --only 1
python export_ppts.py --out "DIR" --region all --level region --only 1
python export_ppts.py --out "DIR" --country "Germany;France;Japan" --level country --only 1
python export_ppts.py --out "DIR" --country "Germany;France" --level country --per-country --only 1
python export_ppts.py --out "DIR" --region Europe --country all --level country --only 3
python export_ppts.py --out "DIR" --region "Global;North America" --one-deck --only 1
```

Common region names: `Global`, `North America`, `Latin America`, `Europe`, `East Asia`, `South Asia & Pacific`, `Middle East & Africa`.

### Limit the ZIPs (exports 2 and 4)

```
python export_ppts.py --out "DIR" --zip-regions "North America"
python export_ppts.py --out "DIR" --zip-regions "North America;Europe"
```

### Full run for one region, all four exports

```
python export_ppts.py --out "DIR" --region "North America" --country "U.S." --zip-regions "North America"
```

### Inspect without building

```
python export_ppts.py --data "X.xlsx" --list
python export_ppts.py --data "X.xlsx" --out "DIR" --only 1 --region all --level region --dry-run
```

### Template, wording, industry

```
python export_ppts.py --out "DIR" --template fmi
python export_ppts.py --out "DIR" --template factmr
python export_ppts.py --out "DIR" --template tmr
python export_ppts.py --out "DIR" --writeup classic
python export_ppts.py --out "DIR" --writeup asis
python export_ppts.py --out "DIR" --industry food
python export_ppts.py --out "DIR" --industry consumer
python export_ppts.py --out "DIR" --industry healthcare
python export_ppts.py --out "DIR" --industry pharma
python export_ppts.py --out "DIR" --industry industrial
```

### Extra inputs

```
python export_ppts.py --out "DIR" --toc "TOC_Market.docx"
python export_ppts.py --out "DIR" --companies "Companies.xlsx"
python export_ppts.py --out "DIR" --segmentation "Segmentation_Brief.xlsx"
python export_ppts.py --out "DIR" --indicative --only 3
```

### All options

| Flag | Meaning |
|---|---|
| `--out DIR` | Output folder (required unless `--list`) |
| `--data XLSX` | Workbook (else `MARKET_DATA`, else `Dataset_*.xlsx`) |
| `--segmentation XLSX` | Generate dataset from a segmentation brief first |
| `--toc DOCX` | Research team's table of contents (else `REPORT_TOC`, else `TOC_*.docx`) |
| `--companies XLSX` | Workbook with a `Competition` sheet for company coverage |
| `--region` | One name, `;` list, or `all` |
| `--country` | One name, `;` list, or `all` |
| `--level` | `region` / `country` / `both` (default) |
| `--indicative` | Example data: locked deck withholds every figure |
| `--one-deck` | Several regions into one file |
| `--per-country` | One deck per country instead of per region |
| `--list` | Print regions and countries in the workbook |
| `--dry-run` | Print decks that would be built |
| `--writeup` | `classic` (default) / `asis` |
| `--industry` | `food` / `consumer` / `healthcare` / `pharma` / `industrial` |
| `--template` | `fmi` (default) / `factmr` / `tmr` |
| `--only LIST` | Options to run, e.g. `1,3` (default: all four) |
| `--zip-regions LIST` | Regions for exports 2 and 4, `;` separated |

---

## 2. Batch exporter - `build_fmi_exports.py`

Uses the `MARKET_DATA` environment variable. Output goes to `.\exports` by default.

```
set MARKET_DATA=Dataset_Your Market.xlsx
python build_fmi_exports.py
python build_fmi_exports.py --what view
python build_fmi_exports.py --what view --region Europe
python build_fmi_exports.py --what locked --region Europe
python build_fmi_exports.py --what regions
python build_fmi_exports.py --what locked-regions
python build_fmi_exports.py --what view locked
python build_fmi_exports.py --template factmr
python build_fmi_exports.py --country Germany
python build_fmi_exports.py --out "DIR"
python build_fmi_exports.py --list
```

PowerShell equivalent of `set`:

```
$env:MARKET_DATA = "Dataset_Your Market.xlsx"
```

| Flag | Meaning |
|---|---|
| `--what` | `all` (default), `view`, `regions`, `locked`, `locked-regions` (one or more) |
| `--region` | Region for single-deck exports (default `Global`) |
| `--country` | Country for country pages |
| `--template` | `fmi` / `factmr` |
| `--out` | Output folder (default `.\exports`) |
| `--list` | Show regions and countries, then exit |

---

## 3. Single deck - `export_ppt.py`

```
python export_ppt.py
python export_ppt.py --region Europe --country Germany
python export_ppt.py --report main
python export_ppt.py --report country
python export_ppt.py --report both
python export_ppt.py --out "C:/path/My Report.pptx"
```

> Note: this script imports `fmi_report.export`, which was removed from this copy
> (see `READ ME FIRST.md`). Use `export_ppts.py` instead.

---

## 4. Checks

Is the workbook shaped the way the engine needs?

```
python check_workbook.py --data "X.xlsx"
python check_workbook.py --data "X.xlsx" --unit "USD Million"
```

Did the built deck get all its numbers? (exit code 1 if anything is blank)

```
python check_deck.py "1 - Some Market Report (FMI).pptx"
python check_deck.py "1 - Some Market Report (FMI).pptx" 10
```

---

## 5. Market-specific scripts

Write into `Output\<market>\` (override with the `FPP_DEST` environment variable).

```
python export_flexible_plastic.py
python export_flexible_plastic.py Global
python export_flexible_plastic.py Europe
python export_flexible_pricing.py
python export_usa_prenatal.py
```

---

## 6. Environment variables

| Variable | Purpose |
|---|---|
| `MARKET_DATA` | Workbook file name or full path (also `MATCHA_DATA`) |
| `REPORT_TOC` | Table-of-contents `.docx` |
| `REPORT_COMPANIES` | Companies workbook |
| `FMI_WRITEUP` | `classic` or `asis` |
| `FMI_INDUSTRY` | Industry for classic text |
| `FMI_INDICATIVE` | `1` = example data |
| `FPP_DEST` | Output folder for the market-specific scripts |

CMD:

```
set MARKET_DATA=Dataset_Your Market.xlsx
```

PowerShell:

```
$env:MARKET_DATA = "Dataset_Your Market.xlsx"
```

---

## 7. Setup

```
pip install -r requirements.txt
python --version
```

Run times (approx): unlocked single deck ~30 s, all regions ~10 min, locked single ~2-4 min, locked all regions ~20+ min.
The run log is written to `_export_log.txt` in the output folder.
