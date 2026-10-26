# Prompt: add the FMI Formatted Dataset output to this project

Paste everything below the line into the project's AI assistant (Claude Code / Copilot Chat in
VS Code), with `fmi_formatted_dataset.py` already copied into the repo.

---

## Task

This project generates FMI Excel Reports (for example
`Global_Packaged_Muffins_Market_Excel_Report_Future_Market_Insights_2026.xlsx`, with sheets
`Market Data` and `Cross-sectional`). Every time a report is generated, the project must also
generate the matching **Formatted Dataset** workbook, saved next to the report as
`<Geo>_<Market>_Market_Formatted_Dataset.xlsx`.

The converter already exists: `fmi_formatted_dataset.py` (single file; needs `openpyxl` and
`xlsxwriter`). **Do not rewrite it or reimplement its logic.** Integrate it.

## Steps

1. Put `fmi_formatted_dataset.py` in the project's tools or src package (wherever the report
   writer lives) and add `openpyxl` and `xlsxwriter` to `requirements.txt` / `pyproject.toml`.
2. Find the function that saves the Excel Report. Directly after the save, call the converter
   using whichever source is easiest in this codebase, in this order of preference:

   **A. From the report file just written (simplest, always works):**
   ```python
   from fmi_formatted_dataset import read_excel_report, run_qc, write_formatted_dataset, default_out

   ds = read_excel_report(report_path, market=market_name)   # e.g. "Packaged Muffins"
   issues = run_qc(ds)
   out = default_out(ds, report_path)
   write_formatted_dataset(ds, out)
   ```

   **B. From in-memory data (use this if the project already holds the numbers as rows):**
   build one dict per leaf row and per metric, then
   ```python
   from fmi_formatted_dataset import from_records, write_formatted_dataset
   records = [
       {"Region": "Global", "Segment": "Product Type", "Sub-segment": "Sweet muffins",
        "Sub-segment 1": "Fruit muffins", "Sub-segment 2": "Fruit muffins",
        "Unit": "Value", "values": [...16 numbers or "XX"...]},
       ...
   ]
   ds = from_records(records, years=range(2021, 2037), market="Packaged Muffins",
                     value_unit="US$ Mn", volume_unit="MT")
   write_formatted_dataset(ds, out_path)
   ```
   A pandas long table works too: `from_records(df.to_dict("records"), ...)`.

   **C. From the segmentation only (no numbers yet, sample decks):**
   ```python
   from fmi_formatted_dataset import from_taxonomy, write_formatted_dataset
   ds = from_taxonomy(
       taxonomy={"Product Type": {"Sweet muffins": ["Fruit muffins", "..."],
                                  "Savoury muffins": ["Cheese muffins", "..."]},
                 "Storage Format": ["Ambient", "Chilled", "Frozen"]},
       geographies={"North America": ["U.S.", "Canada"], "Europe": ["Germany", "..."]},
       years=range(2021, 2037), market="Packaged Muffins",
       cross=[("Product Type", "Flavour")])        # optional
   write_formatted_dataset(ds, out_path)          # every cell = "XX"
   ```

   If the project's own Raw Data workbook (Region | Segment | Sub-segment | Sub-segment 1 |
   Sub-segment 2 | Unit | years) is what exists, use `read_raw_data(path)` instead of A.

3. Log the QC result: print every `ERROR` and `WARN` from `run_qc(ds)`. Do not block report
   generation on `WARN`; on `ERROR` still write the file but surface the messages clearly.
4. Add a CLI hook or script entry so a user can also run it by hand:
   `python fmi_formatted_dataset.py <report or raw data.xlsx> [-o OUT.xlsx] [--market "..."]`.

## Output format (fixed; do not change)

| Sheet | Content |
|---|---|
| Home | Navigation buttons: Value Data, Master Sheet, Volume Data |
| Master Sheet | Title merged A1:V4 ("Global <Market> Market Value (US$ Mn) and Volume (MT) Analysis and Forecast, 2021-2036"); header row 6: `Region, Segment, Sub-segment, Sub-segment 1, Sub-segment 2, Unit, 2021…2036`; leaf rows only |
| Value | Pivot-style outline pasted as values: `Unit | Value` on row 19, `Row Labels | 2021…` on row 21, data from row 22, Grand Total at the end |
| Volume | Same as Value for Volume |

Rules the converter already applies (keep them if you touch anything):

- **Unit column holds `Value` or `Volume`**, not `US$ Mn` / `MT`. The units go in the sheet
  titles only. (`--unit-style unit` / `write_formatted_dataset(..., unit_style="unit")` switches
  the column back to US$ Mn / MT if a client asks.)
- Leaf rows only. `Total Market` rows and parent rows (e.g. "Sweet muffins" when its children are
  listed) are dropped; the pivot sheets rebuild them as sums.
- Hierarchy mapping from `└` indents: 1 level → name repeated in all three sub-segment columns;
  2 levels → `Sub-segment` = parent, `Sub-segment 1` = `Sub-segment 2` = leaf; 3 levels → one per
  column.
- Geography breakdowns: `Region / Country` under Global becomes Segment `Region`; under a region
  it becomes Segment `Country`; the geography name is repeated across the sub-segment columns.
- Cross-sectional pairs `A × B` become Segment `"<Dim 1> × <Dim 2>"`, `Sub-segment` = A's top
  parent, `Sub-segment 1` = A, `Sub-segment 2` = B. Pairs whose A or B has child pairs listed are
  dropped so nothing double counts.
- Master order: geography in source order → all Value rows, then all Volume rows → geography
  segment, plain segments, cross segments. Pivot sheets: aggregates first, then countries A to Z
  with "Rest of…/Other…" last; everything inside a geography A to Z like an Excel pivot.
- Values rounded to 3 decimals; `XX` (or any text) carries through and makes its parent sums `XX`.

## Acceptance checks (run them before you finish)

1. Generate one report end to end; confirm the Formatted Dataset file appears next to it and
   opens in Excel with 4 sheets in the order Home, Master Sheet, Value, Volume.
2. Master Sheet row 6 headers match exactly; the Unit column contains only `Value` and `Volume`.
3. For a numeric run, `run_qc` reports 0 `ERROR`s: every segment of a geography sums to the same
   total, every region equals the sum of its countries, Global equals the sum of regions.
4. Value sheet `Grand Total` equals the top geography's total.
5. No changes were made inside `fmi_formatted_dataset.py` other than the config constants at the
   top (titles, colours, column widths), unless a test above failed.
