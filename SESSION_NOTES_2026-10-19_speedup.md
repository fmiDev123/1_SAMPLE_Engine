# Speed-up of deck generation (19 Oct 2026)

Goal: build decks faster, change nothing in the decks.  Result: about 2x faster on one core, and
an optional `--jobs N` to build several decks at the same time on several cores.

## Rollback point (made BEFORE any change)
`backup_before_speedup_2026-10-19.zip` (+ `backup_before_speedup_2026-10-19.manifest.json`, a SHA-1 of every file).
To go back: extract the zip over the project folder.  (Quick off-switch that needs no restore:
set `FMI_NO_SPEEDUP=1`; it disables `fmi_report/speedup.py` only.)

## What was slow (cProfile of a 222-slide locked deck, 147 s)
build of the unlocked deck ~35 %, locking ~60 %.  Inside the locking: the same few hundred phrases
turned into regexes 1.5 million times; python-pptx re-walking the whole package to name every new
picture; every slide's heading re-found 4-5 times by creating a shape object per shape; XPath for
"is this a placeholder" ~750,000 times.

## What changed (files)
| file | change |
|---|---|
| `pptlock.py` | phrase normaliser / protect pattern / compiled regex remembered (`functools.lru_cache`); `scan()` remembers each slide's heading while it runs; `_slide_heading` reads the title band straight from the slide XML (`_heading_fast`), falling back to the old code when a text shape has no position of its own |
| `fmi_report/speedup.py` (new) | python-pptx: next free `imageN` name and "same image already in the package" found from a cache that is dropped whenever a relationship is removed; placeholder / non-visual-properties lookups without XPath; `qn()` remembered |
| `fmi_report/__init__.py` | imports `speedup` |
| `export_ppts.py` | new `--jobs N|auto` (default 1) and internal `--run-job`; prints a notice about CPU / memory before using more than one core |
| `compare_pptx.py` (new) | compares two decks part by part (see below) |

Nothing that decides WHAT goes on a slide was touched.

## Multi-threading / multi-processing: what is possible here, honestly
* Threads: no gain.  python-pptx / lxml hold the Python GIL and a deck is one shared object.
* One deck across several processes: not safe.  The builder keeps running state (numbering, contents,
  section counters, hyperlinks) in module globals and python-pptx cannot merge two decks, so splitting
  a deck would change it.  The locking step's policy also cannot be sent to another process (it is a
  class made on the fly and reads its caller's local variables).
* Several decks at once: safe and effective, because each deck is a separate run.  `--jobs 2`/`auto`
  starts one process per deck.  It uses that many CPU cores and about 0.7-1.2 GB of memory each (measured
  peak 709 MB for a 196-slide deck).  Default stays `--jobs 1` so nothing uses extra cores unless asked.
  `--jobs auto` leaves one core free and never exceeds free memory / 1.5 GB.

## Tests (all with `PYTHONHASHSEED=0`, which makes the original code byte-reproducible; two runs of the
## ORIGINAL code were identical part for part, so a byte comparison is a fair test)
| deck | original | new | files differing |
|---|---|---|---|
| Air to Air Heat Pump, FMI, Global/Western Europe/Germany (222 slides) | 147 s | 74 s | 0 of 1769 |
| Coal, TMR, Global/Western Europe/Germany (196) | 120 s | 58 s | 0 of 1590 |
| Beauty Device, TMR, Global/MEA/South Africa (247) | 238 s | 123 s | 0 of 2167 |
| Printed Rigid Packaging, FMI, Global/Europe/Germany (356, volume kept) | 315 s | 170 s | 0 of 2467 |
| Coal, plain `export_ppts.py` | 124 s | 60 s | 0 of 1546 |
| Heat Pump `--only 1,3` one after the other | - | 110 s | - |
| same with `--jobs 2` | - | 76 s | locked deck 0 differ; unlocked deck differs only in the clock written inside each embedded chart workbook (the original does this too) |
| Locking step alone on the same deck | 98 s | 61 s | 0 |
Fast heading vs old heading: same text and same XML side effects on all slides of four decks.

Against the reference decks in `Attachments\Desktop\to pratham` (Coal, Beauty Device, Printed Rigid Packaging
built with today's engine through plain `export_ppts.py`): same slide count and the same heading on every slide
(196/196, 247/247, 267/267).  About 10-12 slides per deck carry 2 more padlock pictures than the 16 Oct
references - the same on the ORIGINAL code, so it comes from the earlier 19 Oct change (only the given figures
stay readable), not from the speed-up.

## How to re-check after any change
`python compare_pptx.py A.pptx B.pptx` (exit 0 = identical).  Build with `set PYTHONHASHSEED=0` first.
