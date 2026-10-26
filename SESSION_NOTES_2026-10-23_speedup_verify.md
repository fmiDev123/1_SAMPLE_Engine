# Speed-up: re-verification, no new changes (23 Oct 2026)

Client asked (again) to research and implement multi-threading/multi-processing for faster PPT
generation, with a rollback point first, a warning before using extra CPU, and before/after tests
against known-good decks in `C:\Users\Sayali\Downloads\Test`. This is the third time this exact
brief has been worked: see `SESSION_NOTES_2026-10-19_speedup.md` (wrote the speed-up) and
`SESSION_NOTES_2026-10-22_speedup_wireup.md` (found it was only half wired in, fixed the import).
Both conclusions still hold today; nothing in the engine was changed in this session.

## Rollback point (made BEFORE checking anything)
`backup_before_multiproc_speedup_2026-10-23.zip` (+ matching `.manifest.json`, SHA-1 per file).
To go back: extract the zip over the project folder.

## What's already in place (confirmed still active today)
- `fmi_report/__init__.py` imports `speedup` (the 22 Oct fix) - python-pptx image-partname /
  sha1-lookup / shape-lookup caching is live.
- `pptlock.py` has its own `lru_cache`'d phrase/regex/heading lookups.
- `export_ppts.py --jobs N|auto` builds several DECKS at once, one process per deck, and prints the
  CPU/memory warning before using more than one core. Default stays `--jobs 1`.

## New research this session: is there anything left to parallelise inside ONE deck?
Checked whether chart drawing could be split out to worker processes (render charts independently,
assemble sequentially). Finding: no - every chart in `fmi_report/export_native.py` is built as a
native, editable PowerPoint chart (`pptx.chart.data.CategoryChartData` / `BubbleChartData`) written
directly into the shared `Presentation` package; `fmi_report/figures.py`'s matplotlib charts are a
separate, unused-by-export leftover (not the Dash portal this copy doesn't carry, see "READ ME
FIRST.md"). There is no independent, picture-only chart-rendering step to offload to another process
the way the 19/22 Oct notes already ruled out splitting the whole deck - same conclusion, now also
checked from the charts angle specifically. No safe additional parallelism found.

## Tests today (same machine, same moment, PYTHONHASHSEED=0 both runs)
Deck: Protein Ingredients, FMI, Global/Middle East & Africa/UAE, one-deck, locked (`--only 3`) -
inputs from `C:\Users\Sayali\Downloads\Test` (figures 2026=4503.2, 2036=8706.3, CAGR 6.8%, matching
the 21 Oct recipe in `SESSION_NOTES_2026-10-21_protein_groundmat.md`).

| | `FMI_NO_SPEEDUP=1` | fixed (default) |
|---|---|---|
| time | 306 s | 220 s (~1.4x) |

`compare_pptx.py` on the two: **0 of 2756 parts differ** - the speed-up changes nothing.

Multi-deck `--jobs`: Ground Protection Mat, TMR, Global/North America/U.S., `--only 1,3` (builds the
unlocked and locked deck together).

| | `--jobs 1` (sequential) | `--jobs 2` |
|---|---|---|
| time | 160 s | 94 s (~1.7x) |

Warning printed as designed: `NOTE  --jobs: building 2 decks, 2 at the same time - this uses 2 of
this computer's 4 CPU cores and roughly 3.0 GB of memory until it finishes`. Locked decks from both
runs: 0 parts differ. Unlocked decks: only the embedded chart workbook's own save-timestamp
(`docProps/core.xml` inside each `ppt/embeddings/Microsoft_Excel_SheetN.xlsx`) differs - same
harmless, pre-existing artifact the 19/22 Oct notes already documented.

## Against the stored "good" reference decks in `Downloads\Test`
Raw byte/part comparison against `Protein Ingredients Market Report.pptx` is NOT meaningful: that
deck was built 21 Oct, before the 22 Oct fix that restored Volume onto every segment page (see
[[feedback_volume_on_every_segment]] in Claude's memory / `deck_style.py`), so today's rebuild has 29
more slides (321 vs 292) by design, not by regression. Checked structurally instead: 0 of 321 slides
with no body (`scan_empty_bodies.py`), headline figures (4,503.2 / 8,706.3 / 6.8%) present in the
locked deck's open text. No speed-up-caused defect found.

## How to re-check after any future change
`set PYTHONHASHSEED=0`, build with and without `FMI_NO_SPEEDUP=1`, diff with `python compare_pptx.py
A.pptx B.pptx` (exit 0 = identical; embedded-chart-workbook timestamp diffs are expected and
harmless). For a structural check against an older stored reference deck, compare slide count /
headings / `scan_empty_bodies.py` rather than raw zip parts, since unrelated content fixes made
between builds will legitimately change slide count.
