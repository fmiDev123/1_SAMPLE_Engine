# Report engine, command line only

A copy of the report project with the web side removed. It builds the same
PowerPoint decks the dashboard's four export buttons build, from a command
prompt, with no server to start and no browser involved.

Taken 2026-09-01 from `C:\Users\Shweta\Downloads\all lock\matcha_report`
at commit `49edd3f`.

## What was left out, and why nothing here needs it

Removed: the whole `portal/` web application (Studio, the client viewer, its
database and uploads), and the four modules that exist only to serve a
dashboard - `fmi_report/app.py`, `export.py`, `markets.py`, `serving.py`.
Also the audits, the deploy configuration, the test suite and the git
history, none of which an export needs.

That list was not guessed. The import graph of the export path was walked
from `export_native`, `locking`, `config`, `data`, `theming` and `registry`,
and it reaches 24 modules; the four above are not among them. Then both
exports were run from THIS folder to prove it, before it was handed over.

## Run an export

Open a command prompt in this folder and use the same Python that runs the
main project:

```
cd "C:\Users\Shweta\Downloads\report engine backup 2026-09-01"

python export_ppts.py --data "YOUR_WORKBOOK.xlsx" --out "C:\Users\Shweta\Downloads\my exports"
```

That builds all four decks. To pick which ones, and which geography:

```
:: one deck, for one region
python export_ppts.py --data "YOUR_WORKBOOK.xlsx" --out DIR --region "North America" --only 1

:: the locked deck as well, for a client who has not paid yet
python export_ppts.py --data "YOUR_WORKBOOK.xlsx" --out DIR --region "Europe" --only 1,3

:: every region, zipped
python export_ppts.py --data "YOUR_WORKBOOK.xlsx" --out DIR --only 2
```

The four numbers are: **1** this view, **2** all regions (ZIP), **3** this
view locked, **4** all regions locked. Locked means every figure is withheld
behind a padlock, which is the version to show before payment.

A locked build is slow - four minutes or so for one region, and about twenty
for a locked all-regions ZIP - because every number is masked one at a time.
An unlocked deck takes about thirty seconds.

## Samples for any mix of geographies

`--region` and `--country` take one name, a `;`-separated list, or `all`.
`--level` sets the depth: `region` (region/global report only), `country`
(country pages only) or `both` (default). One deck is written per region; add
`--per-country` for one deck per country.

```
python export_ppts.py --data X.xlsx --list                      :: regions and countries available
python export_ppts.py --data X.xlsx --out DIR --only 1 --region Global --level region
python export_ppts.py --data X.xlsx --out DIR --only 1 --region "Europe;East Asia" --level region
python export_ppts.py --data X.xlsx --out DIR --only 1 --country "Germany;France;Japan" --level country
python export_ppts.py --data X.xlsx --out DIR --only 1 --country "Germany;France" --level country --per-country
python export_ppts.py --data X.xlsx --out DIR --only 3 --region Europe --country all --level country
python export_ppts.py --data X.xlsx --out DIR --only 1 --region all --level region --dry-run
```

`--dry-run` prints the decks that would be built without building them.
Options 1 and 3 follow these flags; 2 and 4 (ZIPs) follow `--level` and
`--zip-regions`. A bare `--region Europe` now opens on Europe's lead market
rather than the dataset-wide one.

## Check a workbook before you build from it

```
python check_workbook.py "YOUR_WORKBOOK.xlsx"
```

## What is in here

```
export_ppts.py        the four exports, from the command line
export_ppt.py         a thinner single-deck entry point
build_fmi_exports.py  a batch of exports in one run
check_workbook.py     is this workbook shaped the way the engine needs
pptlock.py            the padlock and blur pass a locked deck goes through
pptblur.py            the blur it uses
fmi_report/           the engine: data, metrics, figures, page templates
TOC_<market>.docx     the research team's table of contents (optional)
fmi_theme/            the FMI and Fact.MR slide masters
sample_workbook/      one dataset, so this can be tested without a real one
HOW_TO_EXPORT.md      the long guide, written for someone non-technical
SWAP_MARKET.md        how to point the engine at a different workbook
MULTI_MARKET.md       running several markets
```

## Proved working, not just copied

Both paths were run inside this folder before it was handed over:

```
1 - Humira Biosimilars Market Report - North America (FMI).pptx    1.61 MB, 28s
```

The locked path was run too, because it is the one that reaches
`pptlock.py` at the project root - a previous backup of this project copied
only the package folders and missed exactly those root files, and the gap
was not noticed until someone needed a locked deck.

Delete `_smoketest/` whenever you like; it only holds those test decks.

## If you need the dashboard or the portal back

This copy cannot serve them. Use the full project at
`C:\Users\Shweta\Downloads\all lock\matcha_report`.
