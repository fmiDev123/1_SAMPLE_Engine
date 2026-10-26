# How to switch the dashboard to a different market

The dashboard is fully data-driven. Changing markets means **swapping one
Excel file** - no code changes. The market name, page titles, regions,
countries, segments and years are all read from the workbook at startup.

## 1. Preview locally first

1. Copy the new workbook (e.g. `Dataset_Global Green Tea Market.xlsx`)
   into this project folder.
2. **Delete the previous** `Dataset_*.xlsx` - keep exactly ONE in the
   folder (if there are several, the app just uses the first one
   alphabetically).
3. Run the app and check it:
   ```powershell
   python run.py
   ```
   Open <http://127.0.0.1:8050>. Every title / slicer / number now
   reflects the new market automatically.

## 2. Push it to the live link

Once it looks right, from this folder run:

```powershell
& "C:\Users\Shweta\AppData\Local\Programs\Git\cmd\git.exe" add -A
& "C:\Users\Shweta\AppData\Local\Programs\Git\cmd\git.exe" commit -m "Switch dataset to <market> market"
& "C:\Users\Shweta\AppData\Local\Programs\Git\cmd\git.exe" push
```

Render auto-redeploys in ~2 minutes and the same public link shows the
new market. (Git is not on PATH on this PC, hence the full path.)

## Workbook requirement

Must have a **Master Sheet** tab structured as:

```
Region | Segment | Sub-segment | Unit | 2021 ... 2036
```

including the Region/Country hierarchy rows. Any FMI dataset in this format
works.

### The Unit column - `Volume` and `Value`

The `Unit` column names the **measure**, and the report looks each one up by
that word. Write it bare:

```
Unit
Volume
Value
```

A unit written with its measurement baked in - `Volume (MT)`, `Value (US$ Th)`
- is also accepted now: the loader strips the bracket to `Volume` / `Value` and
keeps `MT` as the display label if the sheet title did not already give one.

That decoration used to be silently destructive, and it is worth knowing why,
because the workbook looks perfectly healthy when it happens. Only the volume
side is usually written that way, so `Value` kept matching and `Volume` stopped:
every volume lookup in the deck missed, `has_unit("Volume")` went false, and the
engine concluded the workbook had no volume at all. In one Coffee Extract deck
that silently removed the whole volume section and the whole pricing section
(pricing needs both measures), left 608 country volume tables empty, and drew
one blank chart - from a workbook whose volume figures were complete and
correct throughout.

### Two ways of writing the hierarchy rows - both accepted

Those hierarchy rows say where each geography sits in the Global → region →
country tree, and workbooks write them one of two ways. The engine reads both,
so you do not have to reshape a sheet that uses the second.

**Parented** - the `Region` column holds the PARENT:

```
Region         | Segment | Sub-segment
Global         | Region  | North America
North America  | Country | U.S.
North America  | Country | Canada
```

**Self-referential** - the `Region` column repeats the geography itself, so the
row states what LEVEL the geography sits at but not whose child it is:

```
Region         | Segment | Sub-segment
North America  | Region  | North America
U.S.           | Country | U.S.
Canada         | Country | Canada
```

For the second form the parent is recovered from the sheet's own **row order**,
which is depth-first - a region, then that region's countries, then the next
region. So the rows above are read as U.S. and Canada belonging to North
America. Keep that order and the tree comes out right; nothing is inferred
from the geography NAMES, so it works for any market.

Only the `Region` column is rewritten during the load. Every figure is used
exactly as the workbook states it.

If a sheet's geographies are shuffled out of depth-first order, the loader
leaves the hierarchy alone rather than guess - `check_workbook.py` then
reports `no regions were found`, which is the signal to re-sort the sheet.

## The table of contents

The deck carries the research team's own table of contents, straight from the
Word file they write it in - the pages that used to be pasted in by hand,
thirty rows at a time. Drop that file in this folder beside the workbook:

```
Dataset_Global Green Tea Market.xlsx
TOC_Global Green Tea Market.docx        <- picked up automatically
```

The name only has to start with `TOC_` and end in `.docx`; any `.docx` whose
name mentions "toc" is used as a fallback. To point at one somewhere else,
set `REPORT_TOC` to its path.

Nothing in the file needs preparing. The numbering ("2.", "2.1.",
"2.2.1.1.") is Word's own automatic list numbering, and the engine works it
out the same way Word does, so the contents come out exactly as they read on
screen - indent, colour, the red `Note:` lines and all. **No TOC file is not
an error**: the deck is built without those pages and the log says so.

The **List of Figures** and **List of Tables** that follow are NOT read from
that file. They list the FULL study - 200 figures and 200 tables - planned
by `fmi_report/exhibits.py` from the workbook's own regions, countries,
segments, units and years, and split over the analysis blocks by a weightage:

```
Analysis block                          Figures  Tables
Executive Summary and Market Overview      10       8
Global Market Size and Forecast            12      15
Segment Analysis                           35      45
Cross-Segment Analysis                     15      20
Regional Analysis                          30      28
Country-Level Analysis                     40      38
Pricing, Value Chain and Trade             15      18
Competitive Landscape                      25      20
Trends, Opportunities and Scenarios        12       5
Client-Specific USP Analysis                6       3
```

The split adjusts itself to the study: more segment axes means more segment
exhibits, one axis means no cross-segment exhibits, and a study with no
regions moves the regional share to the country and segment blocks. Priority
countries (the largest by forecast-year value) get 8-12 exhibits each and
secondary countries 2-4. A block the workbook cannot fill (no volume, few
countries) hands the rest to the others, so the totals always hold.

Every title states geography, indicator, segmentation level, unit and period:
`Figure 12: Global XYZ Market Value by Product Type, US$ Mn, 2021–2036`.

To change the split, edit `BLOCKS` at the top of `fmi_report/exhibits.py`, or
for one build point `REPORT_EXHIBIT_WEIGHTS` at a JSON file:

```
{"totals": {"figures": 180, "tables": 220},
 "weights": {"segment": [40, 50], "country": [30, 30]}}
```

All three sit after `Market Taxonomy Segments`, each behind its own section
front page, and are listed on the Agenda.

## Gotchas

- **GitHub browser upload caps at 25 MB.** Files larger than that (some
  workbooks are 40 MB+) must go through the `git push` commands above,
  not drag-and-drop on github.com.
- Keep only ONE `TOC_*.docx` in the folder, for the same reason as the
  workbook: several, and the engine takes the first one alphabetically.
- **Thin-data geographies auto-hide.** Countries that appear in the
  hierarchy but have no segment rows are dropped from the country picker
  automatically, so they never show blank pages.
- Keep only ONE `Dataset_*.xlsx` in the folder at a time.
