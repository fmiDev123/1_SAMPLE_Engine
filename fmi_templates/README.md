# fmi_templates: slide content library for the FMI Python deck generator

Industry-neutral, placeholder-driven content for all 76 slide types in the FMI auto-generated
report deck. These 76 types cover every one of the 597 slides in the Packaged Muffins sample. Data-dependent wording is chosen by deterministic rules, and industry terms
come from a lexicon. The library needs no AI and no manual rewrite.

## Files
| File | What it holds |
|---|---|
| `slide_templates.py` | The content: title, subtitle and every text element for each slide type (IDs A01 to N05). |
| `rules.py` | Word-choice rules, such as accelerating vs moderating, leader-equals-fastest variants, concentration bands and price direction. All thresholds are in `THRESHOLDS`. |
| `lexicon.py` | Industry wording for `food`, `consumer`, `healthcare`, `pharma` and `industrial`. |
| `engine.py` | `build_ctx`, `render_slide`, masking (`v_*` becomes `XX`), dash clean-up, and `qc_slide` with per-industry banned terms. |
| `placeholders.py` | Placeholder naming convention and definitions. |
| `demo_muffins.py` | A worked example that re-renders the muffins sample in masked mode. Run `python -m fmi_templates.demo_muffins`. |
| `export_workbook.py` | Regenerates the review workbook from the code. |

## Use in the generator
```python
from fmi_templates import build_ctx, render_slide, qc_slide, BANNED_BY_INDUSTRY

ctx = build_ctx({... names, n_ counts, v_ formatted figures, r_ raw numbers, c_ company data ...})
out = render_slide("G03", ctx, mask=True)          # mask=False for full reports
issues = qc_slide(out, mask=True, banned_terms=BANNED_BY_INDUSTRY[ctx["industry"]])
assert not issues and "_missing" not in out          # block the save if anything fails
```
`out` maps each element key, such as `title`, `insights` or `drivers`, to text or a list of bullet texts. You map each key once to the text box in your layout.

## Placeholder prefixes
* `v_` is a market figure. It is pre-formatted (for example "5.2%") and masked as XX in samples.
* `c_` is real company data and is never masked.
* `n_` is a count.
* `s_` is a model score.
* `p_` is a rule phrase.
* `lx_` is lexicon wording.
* `r_` is a raw number that drives the rules. It is never printed.

BPS values are unsigned; the template adds the + or - sign.

## Classic style (2015 report format)
`classic.py` holds the same content written in the voice of the 2015 FMI reports (reference: REP-GB-717). Each data slide has a title, an italic headline with the key number, auto-numbered Figure and Table captions, and a Market Outlook box.

```python
from fmi_templates.classic import classic_ctx, render_deck
slides, list_of_figures, list_of_tables = render_deck([("R14", classic_ctx(ctx)), ("R23", classic_ctx(item_ctx)), ...], mask=True)
```
* Run `python -m fmi_templates.demo_classic` to regenerate pages of the 2015 report from its own numbers and compare them with the original text.
* Run `python -m fmi_templates.export_classic` to rebuild the classic review workbook.
