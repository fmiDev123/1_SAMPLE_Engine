# Prompt: Redesign the Pricing Assessment slides in the FMI deck generator

Paste everything below the line into Copilot Chat / Claude Code in VS Code with the deck-generator project open. Put `build_pricing_slides.py` (the reference build) in the project root first so the assistant can read it.

---

## Role and goal

You are editing the existing python-pptx deck generator in this project. Replace the current "Pricing Assessment" slide builder with three data-driven slide layouts that match the reference implementation in `build_pricing_slides.py`:

1. **Single-segment pricing slide**: one per single segment (Product Type, Product Format, Pet Type, Sales Channels, Ingredient Type, Primary Product Function / Benefit, Price Tier, Pack Size).
2. **Cross-sectional price matrix slide**: one per cross segment that exists in the data (any `Segment` value containing `×`).
3. **Country price benchmarking slide**: one per report, built on the primary segment (default: Product Type).

Do not rewrite unrelated slides. Keep the existing header (logo, title, italic subtitle, teal/gold/purple separator lines) and footer helpers if the project already has them; only reuse the reference helpers where the project has no equivalent.

## Step 0: Locate and remove the current implementation

- Search the codebase for `Pricing Pattern`, `Price Point Assessment`, `Average Pricing by`, `Price Range`, `Price By Product` and `Price By Application`. These strings belong to the current pricing slide builder.
- Remove the following from that builder:
  - the locked sparkline chart with no Y-axis values;
  - the single-row range chart labelled `A`;
  - the legend that truncates series names (`Shampoos & Cleansin...`);
  - the "Price By Product" / "Price By Application" hyperlink text at the bottom;
  - the right-hand table that pairs product types with sales channels.
- **Check the data source for that right-hand table.** The Master Sheet has no Product Type × Sales Channel cross segment, and its values (for example 15,949.66 for E-commerce) do not match any value/volume ratio in the Master Sheet. If the code builds these numbers by multiplying a base price by channel factors or by any other synthetic method, delete that logic. Never put a cross-sectional price on a slide unless the cross segment exists in the data.

## Step 1: Data layer (add to the data module, not inside slide code)

Source: `Formatted_Dataset.xlsx`, sheet `Master Sheet`, header on row 6. The columns are `Region, Segment, Sub-segment, Sub-segment 1, Sub-segment 2, Unit, 2021 ... 2036`, where `Unit` is `Value` (US$ Mn) or `Volume` (MT).

1. Split rows into Value and Volume frames indexed on `[Region, Segment, Sub-segment, Sub-segment 1, Sub-segment 2]`.
2. **Price formula:** `price_usd_per_mt = value_usd_mn * 1_000_000 / volume_mt`. Replace zero volume with NA before dividing and skip NA cells in charts and tables.
3. **Aggregation rule:** to get a parent, region or "average" price, sum value and sum volume first, then divide. Never take the mean of child prices.
4. Implement these functions (signatures as in the reference file):
   - `segment_prices(v, q, region, segment, level)` returns a DataFrame of prices by year at `level` = `"Sub-segment"` (parent) or `"Sub-segment 1"` (child).
   - `region_total(v, q, region, segment)` returns the volume-weighted average price for the region.
   - `cagr(start, end, n)`.
5. **Availability detection** (drives which slides get built):
   - `single_segments` = unique `Segment` values for the region that do not contain `×` and are not `Country`.
   - `cross_segments` = unique `Segment` values that contain `×`. Split on `×` to get the row dimension and column dimension names.
   - `countries` = unique `Region` values other than the regional total; drop `Other European Countries` from rankings.
   - A segment is chartable only if at least 2 sub-segments have non-NA prices for every year from 2021 to 2036.

## Step 2: Slide routing rules

| Data situation | Slide to build | Chart level |
|---|---|---|
| Single segment with 3 to 8 parents | Single-segment slide | Parent (`Sub-segment`) |
| Single segment with fewer than 3 parents and up to 8 children (for example Pet Type) | Single-segment slide | Child (`Sub-segment 1`) |
| Flat segment where parent equals child (Product Format, Price Tier) | Single-segment slide | Parent |
| More than 8 series at the chosen level | Single-segment slide | Top 7 by value plus an aggregated "Others" (sum value / sum volume) |
| Cross segment present (`A × B`) | Cross matrix slide | Rows = parents of A; columns = parents of B, mapped from B's own single-segment table |
| Cross segment with more than 9 column parents | Cross matrix slide | Top 8 columns by value plus "Others" |
| Cross segment with fewer than 3 column parents and up to 9 column children | Cross matrix slide | Columns = children of B |
| Cross segment missing | No matrix slide; never synthesise one | |
| Country rows exist for the primary segment | Country benchmarking slide | Parent level of the primary segment |
| Pack Size | Single-segment slide, but chart within each parent | Do not compare ml-based, g-based and piece-based packs on one index |

## Step 3: Shared frame (16:9, 13.333 in × 7.5 in)

- Logo at (0.30, 0.12), height 0.72 in. The title text starts at x = 1.75 in so it never overlaps the logo.
- Title: 22 pt bold, black, at y 0.14. Subtitle: 12.5 pt italic at y 0.53.
- Separator lines at y 0.95: teal `#008189` (0.3 to 4.5 in), gold `#9E7E0A` (4.6 to 8.8 in), plum `#40315E` (8.9 to 13.03 in).
- **Headline box** at (0.30, 1.05, 12.73 × 0.50): navy `#0B2E4A` fill, 13 pt bold white text. It holds one complete sentence that states the main finding (generated, see Step 6).
- Body zone from y 1.65 to 5.50. Insight cards from y 5.60 to 6.95. Footer at y 7.02: the source and method note on the left (8.5 pt, grey `#59636E`), "©2026 Future Market Insights, All Rights Reserved" on the right.
- Panels use a flat `#F3F6F9` fill with no outline, no shadow and no edge strips. Remove the `p:style` element from every auto-shape so LibreOffice and PowerPoint do not add a theme shadow.
- Fonts are Calibri throughout. The minimum size on the slide is 8.5 pt.
- **No em dashes or en dashes anywhere.** Write ranges as `2021 to 2036`. Add a final `assert '–' not in text and '—' not in text` over every run before saving.
- Keep white space at or below 20 percent: the panels and cards must fill the body zone edge to edge with 0.15 in gutters.

## Step 4: Single-segment slide layout

| Zone | Position (x, y, w, h in inches) | Content |
|---|---|---|
| Left panel | 0.30, 1.65, 7.05, 3.85 | Panel title: `Price Trend by {Segment}, 2021 to 2036 (US$/MT)`. Native LINE chart at 0.38, 1.98, 6.9, 3.48. One series per sub-segment, ordered by 2026 price from high to low. 2 pt lines, no markers, series colours in order `C55A11, 0B2E4A, 7030A0, 008189, 9E7E0A, 5B7DB1, A23B5C, 6E8B3D`. Legend at the bottom with full names (no truncation). Visible Y axis with `#,##0` format, min and max rounded to 500 around the data, major unit 500, light gridlines `#D9DEE4`. Italic note at the top right: `Forecast period: 2026 to 2036`. |
| Right top | 7.50, 1.65, 5.53, 1.75 | Panel title: `2026 Price Index vs {Region} Average (= 100)`. Native BAR_CLUSTERED chart of `price_2026 / region_avg_2026 * 100`, sorted with the highest at the top. Bars with an index of 100 or more are orange `C55A11`; bars below 100 are blue `5B7DB1`. Data labels `0.0` at the outside end. The value axis is hidden. Set `has_title = False` (otherwise the series name "Index" prints as a title). Force `c:tickLblSkip val=1` on the category axis so every label prints. |
| Right bottom | 7.50, 3.48, 5.53, 2.02 | Table with columns `{Segment} | 2021 | 2026 | 2036 | CAGR 2026 to 2036 | Index`. Column widths are 2.33, 0.6, 0.6, 0.6, 0.85 and 0.55 in. The last row is `{Region} Average`. The Index cell is shaded by band (Step 7). The header is navy with white 8.5 pt text; rows alternate white and `#F3F6F9`. |
| Insight cards | 3 cards across, y 5.60, h 1.35 | Price Hierarchy, Price Momentum, Outlook (Step 6). |

## Step 5: Cross-sectional matrix slide layout

| Zone | Position | Content |
|---|---|---|
| Left panel | 0.30, 1.65, 8.25, 3.85 | Panel title: `2026 Price Matrix (US$/MT), Shaded by Index vs {Region} Average`. Native table at 0.40, 2.02, 8.05 × 3.10. Rows = row-dimension parents sorted by row average price. Columns = column-dimension parents in source order, then `Row Avg.`. The last row is `Column Avg.`. Every cell = sum value / sum volume for that pair in 2026, shaded by its index against the overall cross-segment average. Column headers use full names at 8 pt and wrap; never truncate them. The first column is 1.7 in, the last is 0.66 in, and the remaining width is split equally. Below the table, a 5-chip colour legend: `Index below 97.5`, `97.5 to 99`, `Parity: 99 to 101`, `101 to 102.5`, `Index above 102.5`. |
| Right panel | 8.70, 1.65, 4.33, 3.85 | Panel title: `Price Range Across All {Col} Sub-Types, 2026`, with the italic note `Bar spans lowest to highest price (US$/MT) at the finest data level.` Native BAR_STACKED floating range chart: series 1 = min price (no fill), series 2 = max minus min (teal `008189`), overlap 100, gap 55. The value axis min is set to the floor (rounded down to 2,000), so the axis reads in real US$/MT. Each range bar carries a custom data label `"{min} to {max}"` in white bold 8.5 pt at the inside base. The min and max are computed at the finest level (`Sub-segment 1` × `Sub-segment 2`). |
| Insight cards | as above | Premium Pockets, Value Pockets, Within-Type Spread. |

## Step 6: Generated text (headline and insight cards)

All text must be complete sentences built only from computed numbers. Do not include market speculation or causal claims that the data cannot show. Use these templates, as the reference file does:

- **Single-segment headline:** `{top} command the highest {BASE} price point at US$ {p}/MT, a {spread}% premium over {bottom}.`
- **Price Hierarchy:** `{top} sit at the top of the {BASE} price ladder at an index of {i_top}, while {bottom} trade at {i_bot}, which keeps the full {segment} spread within {spread}%.`
- **Price Momentum:** `{fastest} post the fastest price growth at a {g}% CAGR between {BASE} and {END}, compared with {g_slow}% for {slowest}.`
- **Outlook:** `The {Region} average price rises from US$ {a26}/MT in {BASE} to US$ {a36}/MT in {END}, and the gap between the highest and lowest priced {segment} moves from {d21}% in {START} to {d36}% in {END}.`
- **Cross headline:** `{row} for {col} is the highest priced combination in {BASE} at US$ {p}/MT, {+x}% above the {Region} average.`
- **Premium Pockets / Value Pockets:** use the max and min cells of the index matrix and state the index and the percentage gap to the average.
- **Within-Type Spread:** the row with the largest max/min ratio, with its min and max US$/MT.
- **Country headline:** `{top} is the most premium {Region} market at US$ {p}/MT in {BASE}, {+x}% above the regional average, while {bottom} is the lowest priced.`
- **Country cards:** Country Spread (max/min average across ranked countries), Growth Leader (highest price CAGR 2026 to 2036 versus the regional CAGR), Mix Exposure (the country with the widest gap between its cheapest and most expensive sub-segment).
- Subject-verb agreement: use "command/sit/post" for plural segment names (most names end in "Products"). Add a helper `is_plural(name)` that returns True when the name ends in `s` and pick the verb form from it.

## Step 7: Index colour bands (shared by all three slides)

| Index vs average | Fill |
|---|---|
| 102.5 or more | `#F4B183` |
| 101.0 to 102.5 | `#FBE3D3` |
| 99.0 to 101.0 | `#F7F7F7` |
| 97.5 to 99.0 | `#D9E6F2` |
| below 97.5 | `#9DC0E0` |

## Step 8: Country benchmarking slide layout

| Zone | Position | Content |
|---|---|---|
| Left panel | 0.30, 1.65, 8.25, 3.85 | Panel title: `2026 Price Range Across {Segment plural} by Country (US$/MT), Ranked by Average`. Native COLUMN_STACKED floating range chart at 0.38, 2.18, 8.1 × 3.3: series 1 = the country's min parent price (no fill), series 2 = max minus min. Countries are sorted by average price from high to low. Bars are teal when the index is 100 or more and blue `5B7DB1` when it is below. The value axis min is the floor rounded to 1,000, with major unit 1,000 and format `#,##0`. Each bar carries a custom data label with the country index (`113`), in white bold 8 pt at the inside end. Force `tickLblSkip = 1`. |
| Right panel | 8.70, 1.65, 4.33, 3.85 | Table `Premium and Discount Markets`: the top 5 and bottom 5 countries with `Avg. 2026 | Index | CAGR 2026 to 2036`. The Index cell is shaded by band. |
| Insight cards | as above | Country Spread, Growth Leader, Mix Exposure. |

## Step 9: Config switches to add

- `PRICING_BASE_YEAR = 2026`, `PRICING_START = 2021`, `PRICING_END = 2036`. The historical and forecast split is read from the base year.
- `PRICING_PRIMARY_SEGMENT = "Product Type"`, which drives the country slide.
- `PRICING_MASK = False`. When True, every US$/MT figure in tables, labels and text prints as `XX` and axis number formats become `"XX"`. Indices and CAGR stay visible, so the visuals still read. Use True for sample decks.
- `PRICING_COUNTRY_SLIDES = False`. When True, also build single-segment slides for each country (same layout, with region = country and the index against the country's own average).

## Step 10: QA the assistant must run after the change

1. Recompute three sample cells by hand from the Master Sheet: one parent price, one cross cell, and one country average. They must match the slide to the dollar.
2. Confirm that no chart category label is skipped and no legend entry or column header is truncated.
3. Confirm that every series in a line chart spans 2021 to 2036 with no gaps.
4. Render the deck to PDF with LibreOffice and check each pricing slide for text overflow in the table rows and cards (the target is a maximum of 4 lines per card).
5. Run the dash assertion and confirm every insight bullet is a complete sentence ending with a full stop.
6. Confirm the deck opens in PowerPoint without a repair prompt.

## Expected output for this dataset (Europe Pet Cosmetic Products)

| Slide | Built from | Count |
|---|---|---|
| Single-segment | Product Type, Product Format, Pet Type, Sales Channels, Ingredient Type, Primary Product Function / Benefit, Price Tier, Pack Size | 8 |
| Cross matrix | Product Type × Primary Product Function / Benefit (6 × 8), Product Type × Ingredient Type (6 × 3) | 2 |
| Country benchmarking | Product Type across 24 ranked countries | 1 |
