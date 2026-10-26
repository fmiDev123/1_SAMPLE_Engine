# How to export the PowerPoints - plain-English guide

You do not need to know anything about programming to use this. You type one
line, press Enter, and PowerPoint files appear in a folder you choose.

---

## What you can make

There are **four** exports. One command makes all four.

| # | Name | What it is |
|---|---|---|
| 1 | **This view** | One deck for the region you pick (Global unless you say otherwise) |
| 2 | **All regions** | A ZIP holding one deck **per region** - Global, North America, Europe ... or just the regions you name |
| 3 | **This view, locked** | Same as 1, with the numbers withheld (blurred, with padlocks) |
| 4 | **All regions, locked** | Same as 2, with the numbers withheld in every deck |

"Locked" is the version you can show to someone **before** they have paid: the
report looks complete, but the figures are covered up.

### What "locked" covers

Every number the report shows is withheld. In place of each figure you get a
**padlock**, sitting exactly where the number was, over a blurred smudge. That
includes the years the deck used to give away for free - **2025, 2026, 2036 -
and every CAGR**. Nothing is readable any more: not a table cell, not a chart
label, not a KPI card, not a figure inside a sentence.

Three things stay readable **on purpose**, because without them the charts
would be unreadable rather than merely locked:

- the **scales** down the side and along the bottom of a chart (0%, 50%, 100%)
- the **year captions** under the columns (2021, 2026, 2036)
- **counting numbers** that are not market data, such as a "1, 2, 3" rank badge

So a client can see the shape of the market and none of its numbers.

---

## Step 1 - open the black command window

1. Press the **Windows key** on your keyboard.
2. Type `powershell`.
3. Press **Enter**.

A dark window opens. This is where you type the commands. You can make the
text bigger with **Ctrl** and **+**.

---

## Step 2 - go to the project folder

Copy this line, paste it into the window (right-click pastes), press **Enter**:

```
cd "C:\Users\Shweta\Downloads\Locker Folder\matcha_report"
```

Nothing visible happens. That is correct - it just means "work in this
folder from now on". You do this once each time you open the window.

> **Use this folder, not the other one.** There is a second copy of the same
> project at `C:\Users\Shweta\Downloads\all lock\matcha_report`. The two are
> kept identical, but if you work in the wrong one you can get errors that
> make no sense. When in doubt, use the path above.

---

## Step 3 - check your Excel file is in place

The project reads **one Excel workbook**. Put your workbook in the folder
above, and make sure its name starts with `Dataset_`, for example:

```
Dataset_Air_Ambulance_Market_Filled.xlsx
```

If there is only one such file in the folder, it is used automatically and you
do not have to name it. If there are several, see *"Choosing which workbook"*
below.

---

## Step 4 - the one command that makes everything

```
python export_ppts.py --out "C:\Users\Shweta\Downloads\My exports"
```

That is the whole thing. It makes **all four** exports and puts them in a
folder called `My exports` in your Downloads. The folder is created for you if
it does not exist.

You can change `My exports` to any folder name you like. Keep the quotation
marks - they matter when a folder name has spaces in it.

### What you will see while it runs

Lines appear as each export finishes:

```
[16:22:20] START  Dental X-ray Market  |  Dataset_Dental_Xray_Market_Filled.xlsx  |  FMI
[16:22:20] 1) this view -> 1 - Dental X-ray Market Report (FMI).pptx
[16:22:43]    done  1.51 MB  24s
[16:25:47] 2) all regions -> 2 - Dental X-ray Market - All Regions (FMI).zip
[16:35:20]    done  28.27 MB  573s
```

**Do not close the window while it is working.** The last line says
`FINISHED` when everything is done.

### What lands in the folder

```
1 - <Market> Report (FMI).pptx                    one deck
2 - <Market> - All Regions (FMI).zip              one deck per region
3 - <Market> Report (FMI) (locked).pptx           locked
4 - <Market> - All Regions (FMI) (locked).zip     locked, per region
_export_log.txt                                   a record of the run
```

To get at the decks inside a `.zip`: right-click it → **Extract All**.

---

## How long does it take?

Roughly, on this machine:

| Export | Time |
|---|---|
| 1 - this view | about **half a minute** |
| 2 - all regions | about **10 minutes** |
| 3 - this view, locked | about **2 minutes** |
| 4 - all regions, locked | **20 minutes or more** |

All four together: **around half an hour**. Locking is slow because every
figure has to be blurred one at a time. A market with many countries takes
longer. It is normal for the window to sit quiet for minutes at a stretch.

---

## Making just one of them

Add `--only` and the number:

```
python export_ppts.py --out "C:\Users\Shweta\Downloads\My exports" --only 3
```

Both locked ones only:

```
python export_ppts.py --out "C:\Users\Shweta\Downloads\My exports" --only 3,4
```

---

## Just one region

Add `--region` with the region's name in quotation marks:

```
python export_ppts.py --out "C:\Users\Shweta\Downloads\My exports" --region "North America" --only 3
```

That gives you the locked deck for North America on its own, without waiting
for all the others.

`--region` changes exports **1** and **3** (the single decks). On its own it
does **not** touch exports 2 and 4 - those cover every region, which is what
they are for. To narrow them too, see the next section.

The region name must be spelt exactly as the report spells it. The usual ones:

```
Global
North America
Latin America
Europe
East Asia
South Asia & Pacific
Middle East & Africa
```

---

## Cutting the ZIPs down to certain regions

By default exports **2** and **4** hold one deck for *every* region. If you
only want some of them, add `--zip-regions` with the region names in
quotation marks, separated by semicolons:

```
python export_ppts.py --out "C:\Users\Shweta\Downloads\My exports" --zip-regions "North America"
```

That ZIP now holds the North America deck and nothing else. For more than one
region, separate them with a semicolon:

```
--zip-regions "North America;Europe"
```

**This is the setting that saves the most time.** Export 4 is slow because
every deck in it has to be blurred one at a time, so cutting seven regions
down to one takes it from twenty-odd minutes to about four.

Two things to know:

- The file name changes to say what is inside - you get
  `2 - <Market> - North America (FMI).zip` instead of `... - All Regions ...`.
  So a narrowed ZIP will never quietly overwrite a full one.
- Spell the region exactly as the list above spells it. A name the report does
  not know **stops the export with a message** naming the ones it does know,
  rather than handing you an empty ZIP you would not spot until you opened it.

To get every region again, simply leave `--zip-regions` out.

A full run for a single region - all four exports, nothing wasted - looks
like this:

```
python export_ppts.py --out "DIR" --region "North America" --country "U.S." --zip-regions "North America"
```

---

## Choosing which workbook

If several `Dataset_*.xlsx` files sit in the folder, say which one:

```
python export_ppts.py --data "Dataset_Air_Ambulance_Market_Filled.xlsx" --out "C:\Users\Shweta\Downloads\My exports"
```

---

## The Fact.MR look instead of FMI

```
python export_ppts.py --out "C:\Users\Shweta\Downloads\My exports" --template factmr
```

Leave it out and you get FMI.

---

## Seeing the report in your browser first

To look through the report before exporting:

```
python run.py --port 8123
```

Wait for it to say it is running, then open **http://127.0.0.1:8123** in your
browser. It takes a minute or two to start on a big market.

To stop it, click the black window and press **Ctrl** and **C** together.

If that port is busy, use another number - `--port 8124`, `--port 8125`, and
so on. You can run two markets at once on two different port numbers.

---

## When something goes wrong

**"python is not recognised"**
Python is not installed, or not on the PATH. Install it from python.org and
tick *"Add Python to PATH"* during setup.

**"Could not find an Excel workbook"**
There is no `Dataset_*.xlsx` in the project folder. Put yours there, or point
at it with `--data`.

**"no per-geography sheets recognised"**
The workbook's layout is not one the reader understands - usually a sheet or a
column heading named differently from the other markets. This needs a
developer to look at the file.

**Nothing happens for a long time**
That is usually normal - see the timings above. The `_export_log.txt` file in
your output folder shows how far it got.

**It stopped part-way**
Whatever finished is still good. Re-run with `--only` and just the numbers you
are missing, so you do not redo the slow ones.

**A locked deck opens but looks unlocked**
Check you opened the file with `(locked)` in its name. Files 1 and 2 are the
open versions, on purpose.

---

## Quick reference

| I want ... | Type this |
|---|---|
| Everything | `python export_ppts.py --out "DIR"` |
| Only the locked ones | `python export_ppts.py --out "DIR" --only 3,4` |
| Only all-regions, locked | `python export_ppts.py --out "DIR" --only 4` |
| One region, locked | `python export_ppts.py --out "DIR" --region "Europe" --only 3` |
| ZIPs for one region only | `python export_ppts.py --out "DIR" --zip-regions "North America"` |
| ZIPs for a few regions | `python export_ppts.py --out "DIR" --zip-regions "North America;Europe"` |
| One region, all four | `python export_ppts.py --out "DIR" --region "Europe" --zip-regions "Europe"` |
| A particular workbook | `python export_ppts.py --data "Dataset_X.xlsx" --out "DIR"` |
| The Fact.MR template | `python export_ppts.py --out "DIR" --template factmr` |
| See the reminder list | `python export_ppts.py --help` |
| Open the dashboard | `python run.py --port 8123` |

Replace `DIR` with where you want the files, in quotation marks, for example
`"C:\Users\Shweta\Downloads\My exports"`.

## Faster builds
Decks now build about twice as fast on one core (see SESSION_NOTES_2026-10-19_speedup.md).
Several decks in one command (for example `--only 1,3`, or a deck per region) can also be built at the
same time: add `--jobs 2` (or `--jobs auto`).  This uses that many CPU cores and about 1 GB of memory
per deck while it runs; without `--jobs` nothing changes.
