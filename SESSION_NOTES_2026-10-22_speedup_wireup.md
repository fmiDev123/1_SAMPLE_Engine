# Finishing the 19 Oct speed-up (22 Oct 2026)

Goal (client request): build decks faster, change nothing in the decks. Create a rollback point
first; warn before using more than one CPU; test that output is unchanged.

## Finding: the 19 Oct speed-up was only half wired in
`fmi_report/speedup.py` (python-pptx image-partname / sha1-lookup / shape-lookup caching) was
fully written and documented in `SESSION_NOTES_2026-10-19_speedup.md`, but **nothing imported it** -
`fmi_report/__init__.py` was still just a docstring. So that half of the documented "2x on one
core" gain was dead code; only the inline `pptlock.py` caching (`lru_cache`, `_heading_fast`) and
the `--jobs N/auto` multi-deck parallelism (in `export_ppts.py`) were actually active.

## Rollback point (made BEFORE this change)
`backup_before_speedup_wireup_2026-10-22.zip` (+ matching `.manifest.json`, SHA-1 per file), same
convention as the 19 Oct backup. To go back: extract the zip over the project folder.

## The fix
One line added to `fmi_report/__init__.py`:
```python
from . import speedup as _speedup  # noqa: F401  (import installs its patches)
```
`speedup.py` already calls `install()` at module load time, so importing it is all that was needed.

## Multi-core: already implemented, untouched by this change
`export_ppts.py --jobs N|auto` builds several DECKS at once, each a separate process - the warning
it already prints before using more than one core:
> `NOTE  --jobs: building N decks, M at the same time - this uses M of this computer's C CPU
> cores and roughly M*1.5 GB of memory until it finishes (--jobs 1 = one at a time)`

Default stays `--jobs 1` - no extra cores used unless asked. Parallelising a SINGLE deck (threads or
processes) is not attempted: confirmed (again) that python-pptx/lxml serialise through the GIL and
the builder keeps running state (numbering, section counters, relationship ids, the locking
policy's closures) in module globals, so one deck cannot be split across processes and merged back
without changing the deck. Same conclusion as the 19 Oct notes.

## Tests (PYTHONHASHSEED=0, same command run with the fix off then on)
Dataset: `sample_workbook/Global_Humira_Biosimilars_Market_FMI_Dataset_Formatted.xlsx` (the
project's clean reference workbook - see note below on why the Downloads\Test datasets couldn't be
used directly).

`python export_ppts.py --data ... --out ... --only 1 --template fmi` (unlocked single deck,
Global + default country):

| | FMI_NO_SPEEDUP=1 (pre-fix behaviour) | fixed (default) |
|---|---|---|
| time | 71 s | 48 s (~1.5x) |

`compare_pptx.py` on the two decks: 1113/1113 parts present both sides; 112 differ, all
`ppt/embeddings/Microsoft_Excel_SheetN.xlsx` (the tiny workbook embedded behind each chart).
Unzipped one and diffed its own parts: every actual chart-data part was byte-identical; the only
difference was that embedded workbook's own `docProps/core.xml` (its save timestamp) - the same
harmless, pre-existing clock artifact the 19 Oct notes already documented ("the original does this
too"). No content or formatting difference.

Locked decks (which add a picture per padlock/blurred label - the case `speedup.py`'s image-lookup
caching targets most) were not re-measured today; the 19 Oct notes already measured ~1.9-2.1x on
four such decks and nothing in today's fix changes that code path, only whether it runs at all.

## Side finding (not fixed - out of scope for this task, flagging for the client)
The raw sample workbooks in `C:\Users\Sayali\Downloads\Test` (e.g.
`Cement_Market_FMI_Sample_Dataset_2021_2036.xlsx`) and the project's own
`Coal_Market_TMR_Style_Sample_Dataset.xlsx` contain literal `"XX"` / `"XXXXXXXX"` placeholder text
in Value/Volume cells, across every region in the Cement workbook. `fmi_report/data.py`'s
`segment_table()` does `.astype(float)` on those cells and crashes the moment a page needs that
geo/unit - so `export_ppts.py --data <one of these>` currently fails outright, independent of the
speed-up and of anything touched today. Could not do a byte-for-byte compare against the "good"
decks in that folder for this reason; the Humira workbook above was used instead because it has no
placeholder cells.

## How to re-check after any change
`set PYTHONHASHSEED=0` then build with and without `FMI_NO_SPEEDUP=1`, diff with
`python compare_pptx.py A.pptx B.pptx` (exit 0 = identical; embedded-chart-workbook timestamp diffs
are expected and harmless).
