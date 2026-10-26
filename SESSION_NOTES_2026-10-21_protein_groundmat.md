# Session 21 Oct 2026: two more sample decks

Recipe is the one in `SESSION_NOTES_2026-10-14_to_16.md` section 1. Rebuild the same decks with the commands below.

## Gotchas found this session
- Output path over 260 chars makes the build run fully (3-4 min) and then fail at save with FileNotFoundError.
  Fix: copy inputs to a SHORT work folder (`C:\fmitmp\<name>`) and use a short `--out`; copy the result to
  `C:\Users\Shweta\Downloads\My exports TEST new`. Do not use the Claude scratchpad as `--out`.
- `rm -rf C:/fmitmp` is blocked by the safety check; the user must delete `C:\fmitmp` themselves.
- 2036 values of 1,000+ print as whole numbers (4,503.2 shows 4,503).
- Region/country names must match the workbook: "Middle East & Africa", "U.S." (not "USA").

## Deck 1: Protein Ingredients, Global / Middle East & Africa / UAE, FMI
Inputs: `C:\Users\Shweta\Downloads\5\Global_Protein_Ingredients_Market_Excel_Report_Future_Market_Insights_2026.xlsx`
and `...\5\Global Protein Ingredients Market_Table Of Content_Future Market Insights_2026.docx`.
Figures (US$ Mn): 2026 = 4503.2, 2036 = 8706.3, CAGR 6.8% (consistent).
```
python fmi_indicative_numbers.py W\d.xlsx -o W\d_ind.xlsx --market "Protein Ingredients" --global-value 4503.2 --value-2036 8706.3
python export_ppts.py --data W\d_ind.xlsx --toc W\toc.docx --companies W\d.xlsx --out W\out --only 3 --template fmi --region "Global;Middle East & Africa" --country UAE --one-deck
```
Result: 292 slides, 0 empty, 13 TOC slides, 191 s. File:
`3 - Protein Ingredients Market Report - Global, Middle East & Africa, UAE (FMI) (locked).pptx`

## Deck 2: Ground Protection Mat, Global / North America / U.S., TMR
Inputs: folder `C:\Users\Shweta\Downloads\Ground_Protection_Mat_Market_FACTMR_Style_Dataset\` holding
`Ground_Protection_Mat_Market_FACTMR_Style_Dataset.xlsx` and `Global Ground Protection Mat Market_2026–36_FMI.docx`.
Figures (US$ Mn): 2026 = 1900 (USD 1.9 Bn), 2036 = 3500, CAGR 6.3% (consistent).
```
python fmi_indicative_numbers.py W\d.xlsx -o W\d_ind.xlsx --market "Ground Protection Mat" --global-value 1900 --value-2036 3500
python export_ppts.py --data W\d_ind.xlsx --toc W\toc.docx --companies W\d.xlsx --out W\out --only 3 --template tmr --region "Global;North America" --country "U.S." --one-deck
```
Result: 201 slides, 0 empty, 11 TOC slides, 94 s. File:
`3 - Ground Protection Mat Market Report - Global, North America, U.S. (TMR) (locked).pptx`

Both were built with PYTHONHASHSEED unset, so a rebuild may differ in tiny ways; set `PYTHONHASHSEED=0` for identical output.
Verification used: `python scan_empty_bodies.py deck.pptx`, TOC slide count, headline figures found in slide text.
