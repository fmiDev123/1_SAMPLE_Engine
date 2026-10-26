# FMI_PPT session notes, 14-16 Oct 2026

Everything done in the long session about building locked SAMPLE decks. Read this first
when continuing. Project folder: `C:\Users\Shweta\OneDrive - Future Market Insights Pvt.Ltd\Documents\FMI_PPT`.
Output decks are kept in `C:\Users\Shweta\Downloads\My exports TEST new` (the user renames/moves
files there, so list the folder before assuming a name).

## 1. How a sample is generated (the recipe that worked every time)

Inputs: an Excel dataset (all cells `XX`) + a Word table of contents + the headline figures the
user pastes (2026 value, 2036 value, CAGR) + the geographies (Global, a region, a country) + template.

1. Copy the files to a temp work folder with short names (some input paths are over 260
   characters; Windows then says the file "does not exist". Use `\\?\` prefix to copy).
2. Fill the placeholders with indicative numbers pinned to the headline figures:
   `python fmi_indicative_numbers.py WORK\d.xlsx -o WORK\d_ind.xlsx --market "<Market>" --global-value <2026 in workbook unit> --value-2036 <2036 in same unit>`
   - The unit is what the workbook title says (almost always US$ Mn). 1.8 Bn = 1800.
   - Pass `--global-value` + `--value-2036` (CAGR is derived). If the user's CAGR does not match
     their two end values, say so and ask; do not silently pick (Coal and Shrinkable Paper had
     this: 1,665.2 -> 2,674.2 is 4.85% over 10 years, not 4.4%).
   - This writes `d_ind.xlsx.indicative` with a `headline:` line (needs the headline patch, see 3).
3. Build the locked sample deck:
   `python export_ppts.py --data WORK\d_ind.xlsx --toc WORK\toc.docx --companies WORK\d.xlsx --out WORK\out --only 3 --template fmi|tmr --region "Global;<Region>" --country <Country> --one-deck`
   - `--only 3` = locked deck. `--companies` = the ORIGINAL workbook (its Competition sheet; the
     generated one drops it and the build crashes in `share_takeaways`).
   - Default template FMI unless the user says TMR (or Fact.MR).
   - Typical "key country" for a region when the user does not name one: the first listed (Germany for Europe).
4. Verify (see section 5), copy the .pptx to `My exports TEST new`.
5. Tell the user: figures that are readable, figures that are generated, any mismatch.

Run times: country-only unlocked about 25 s; full locked 3-9 min. For two builds at once use
`Start-Process python -ArgumentList` with the whole argument string quoted (a region name with a
space like "Latin America" must be wrapped in backtick-quoted quotes or it splits).

## 2. Decks produced (all in `My exports TEST new`)

| Market | Geographies | Template | Notes |
|---|---|---|---|
| Organ Transport Devices | Global, North America, U.S. | FMI | 2026 values only indicative |
| Senior Dog Cognitive Supplement | Global, NA, U.S. | FMI | |
| Excavator | Global, Europe, Germany | TMR | "Western Europe" asked for; workbook had only Europe then |
| Breaded Plant-based Protein | Global, Europe, Germany | TMR | 1,602.4 -> 3,205.3, 7.2% |
| Metal Powders for AM | Global, East Asia, China | TMR | 578.9 -> 1,226.8 (prints 1,227) |
| Shrinkable Paper | Global, Europe, Italy | FMI | 1,500 + 5.9% -> 2,661 |
| Hand Tools | Global, Western Europe, Germany | FMI | 18,900 -> 32,900, 5.7% |
| Printed Rigid Packaging | Global, Europe, Germany | FMI | 183,900 -> 293,900, 4.8%; built 3 times |
| Coal | Global, Western Europe, Germany | TMR | 1,665,200 -> 2,674,200; CAGR printed 4.9% (user said 4.4%, inconsistent) |
| Dried Berries | Global, Latin America, Brazil | TMR | user asked Chile; Chile is not a country in that workbook, Brazil used |
| Beauty Device (user's own run) | Global, MEA, South Africa | TMR | repaired, see 4 |

## 3. Engine changes made (all in the project folder; originals zipped beside it)

Each change has a backup zip named `backup_before_<change>_<date>.zip` in the project folder.

1. **Headline figures stay readable** (patch `headline_figures_open_2026-10-15.patch`, applied with
   `patch --binary` after converting the patch to CRLF): `fmi_indicative_numbers.py`,
   `fmi_report/config.py`, `fmi_report/locking.py`, `pptlock.py`. The locked sample keeps readable
   only the given 2025/2026/2036 values, the 2026-2036 opportunity, the multiple and the forecast
   CAGR on Global pages. Everything else is withheld. Backup: `backup_before_headline_figures_open_2026-10-15.zip`.
2. **TOC inside a table cell** (`fmi_report/toc.py`): the Breaded Plant-based Protein Word file kept
   its chapter list in a table, so 0 rows were read and the deck had no TOC pages. Now falls back to
   reading all paragraphs (minus text boxes) when the body has none. Backup: `backup_before_toc_in_table_fix_2026-10-15.zip`.
3. **Geography / segment readers** (`fmi_formatted_dataset.py`):
   - raw reader picks the sheet naming the most geographies (it used a Global-only `Segment Data` sheet);
   - region hierarchy read from the workbook's own layout when no Region/Country rows exist
     (`_geographies_from_layout`, `_regions_first_groups`): handles "region then its countries" and
     "all regions, then all countries"; handles Western/Eastern Europe split and names the standard
     list lacks (Chile, ASEAN, Kingdom of Saudi Arabia, Turkiye, ...);
   - header `Geography | Dimension | Segment | Metric`: Dimension is the main segment, Segment its members.
   Backups: `backup_before_geo_layout_fix_2026-10-15.zip`, `backup_before_dimension_layout_fix_2026-10-16.zip`,
   `backup_before_western_europe_layout_2026-10-16.zip`.
4. **Pricing slide duplicate "Others"** (`fmi_report/pricing_slides.py`): a source sub-segment named
   "Others" clashed with the engine's aggregate "Others" and crashed the build; now folded in.
5. **Blank table pages** (`fmi_report/deck_style.py`), see section 4. Backups:
   `backup_before_blank_table_pages_fix_2026-10-16.zip` (and an earlier `backup_before_blank_page_fix_2026-10-16.zip`
   in the output folder holding the broken decks).

New helper scripts in the project folder: `check_blank_slides.py`, `scan_empty_bodies.py` (lists titled
slides with no body), `restore_table_pages.py` (fills such slides from a side build),
`fix_blank_table_pages.py` (OLD delete version, do not use; the user wants slides filled, not deleted).

## 4. The blank-slide mistake and how it was handled

What the user saw: Printed Rigid Packaging slides 173-176 (and 15 more) and Beauty Device (South
Africa) slides 176, 177, 183, 184 had a title and subtitle but no table.

Cause: the house-style step `deck_style.country_volume_off` removes Volume tables from country pages
but assumed a table is always two slides ("(1/2)" Volume, "(2/2)" Value). A long Volume table spills
over many slides ("(1/6)".."(4/6)"); the panel was removed from each and the slide stayed. The data
was fine. Same flaw on paged "Germany YoY Analysis ... (k/N)" pages (found later in a re-test).

First (wrong) action: I deleted the empty slides. The user said they wanted them REPLACED WITH THE
CORRECT DATA, not deleted, and not a whole regeneration.

Correct handling used:
1. Restored the originals from the backup zip.
2. Built the same pages in a side build with `FMI_DECK_STYLE=off` (country-only, locked) so every page exists.
3. `restore_table_pages.py deck.pptx source.pptx out.pptx` copied the missing shapes (panel, title,
   table, source line, padlock picture) into each empty slide, matched by slide title. Locked tables
   stay locked (cells blank with the padlock overlay).
4. Decks fixed in place: Printed Rigid Packaging (two identical copies, 19 slides each) and Beauty
   Device (4 slides). Slide counts 267 and 247 restored.
5. Model fix: paged "Market Tables" and "YoY Analysis" sets of more than 2 slides now keep all pages
   including Volume. Every build also prints `WARNING: N page(s) have a title but no content` if any
   paged page is empty (added last, compiled but not exercised on a defective deck).

Still open: `Meat Packaging Market Report.pptx` has 4 empty UAE table pages (127, 128, 133, 147);
offered to fill them, not done. Not tested: a full-size locked build with the final rule.

## 5. Checks to run on every new deck

- `python scan_empty_bodies.py "deck.pptx"` (0 expected; section dividers/breakers can show up as
  untitled false positives).
- TOC slides exist: titles "Table of Contents (n/N)".
- Headline figures present in text: search the slide text for the 2026/2036 values and the CAGR.
- The earlier blank-slide check looked only at text length, which a title+subtitle defeats;
  use `scan_empty_bodies.py`.

## 6. Open questions / things the user may revisit

- Coal: 4.4% vs the 4.85% implied by the two end values (deck shows 4.9%).
- Dried Berries: Chile not in workbook; Brazil used. The Excel the user named
  (`TOC_..._TMR.xlsx`) did not exist; used `Global_Dried_Berries_Market_Excel_Report_2026_2036.xlsx`.
- 2036 values of 1,000 or more print as whole numbers (1,226.8 shows as 1,227).
- The sheet-choice change in `read_raw_data` applies to every workbook; earlier samples were not re-run.
