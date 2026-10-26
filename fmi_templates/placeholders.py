"""
Placeholder dictionary. describe(name) returns a plain-English definition for any placeholder used
in slide_templates.py, from an explicit entry or from the naming convention below.

Naming convention
  prefix   v_ market figure (masked as XX in samples) | c_ company data (real, never masked)
           n_ count | s_ model score | p_ rule phrase (rules.py) | lx_ industry wording (lexicon.py)
           r_ raw number feeding a rule (never printed)
  scope    seg1 = primary segment (first in taxonomy, e.g. Product Type)
           seg2 = route-to-market segment (sales channel, end use or distribution channel)
           sub  = the segment being reported in the current chapter
           child = child geography (regions under Global, countries under a region)
           cty  = country (any level)   x / xa / xb = cross-section cells   pm = price-matrix cell
           m    = the metric of the slide (volume or value)
  pick     lead = largest by value | fast = highest F-CAGR | slow = lowest F-CAGR
           gain = largest share gain (BPS) | lose = largest share loss (BPS) | top / high = highest | low = lowest
  measure  share_est / share_end = share in est_year / fc_end | cagr = F-CAGR | h_cagr = historical CAGR
           opp = absolute $ opportunity est_year to fc_end | bps = basis points | idx = index (=100 base)
           gap_bps = difference vs benchmark in BPS (unsigned) | chg = signed % change | mult = end / start multiple

Format conventions for v_ values (supplied pre-formatted by the pipeline)
  CAGR and shares: "5.2%" | BPS: "120" (unsigned, the template adds + or -) | money: "1,234.5" | multiple: "1.6"
"""

EXPLICIT = {
    # identity and scope
    "market": ("name", "Market name in title case without the word 'Market', e.g. Packaged Muffins."),
    "market_lc": ("name", "Market name for mid-sentence use, e.g. packaged muffins. Keep acronyms in capitals."),
    "industry": ("name", "Lexicon profile: food, consumer, healthcare, pharma or industrial."),
    "publisher": ("name", "Future Market Insights."),
    "publisher_short": ("name", "FMI."),
    "source": ("name", "Source line text, e.g. FMI Analysis, 2026."),
    "pub_year": ("name", "Copyright year."),
    "report_date": ("name", "Publication date, e.g. 07 Oct 2026."),
    "geo": ("name", "Geography of the current slide: Global, a region or a country."),
    "geo_ref": ("name", "Geography for mid-sentence use: 'global' for Global, otherwise the name."),
    "geo_ref_cap": ("name", "geo_ref with a capital first letter."),
    "geo_top": ("name", "Top geography of the report, normally Global."),
    "geo_top_ref": ("name", "geo_top for mid-sentence use, normally 'global'."),
    "parent_geo": ("name", "Benchmark geography: Global for a region, the region for a country."),
    "focus_region": ("name", "Region with a deep-dive chapter, e.g. North America."),
    "focus_country": ("name", "Country with a deep-dive chapter, e.g. U.S."),
    "child_label": ("name", "'region' when geo is Global, 'country' when geo is a region."),
    "child_label_pl": ("name", "Plural of child_label."),
    "child_label_cap": ("name", "Capitalised child_label."),
    "child_label_adj": ("auto", "Adjective form: 'regional' or 'country'. Built automatically."),
    "child_label_adj_cap": ("auto", "Capitalised child_label_adj. Built automatically."),
    "hist_start": ("name", "First historical year, e.g. 2021."), "hist_end": ("name", "Last historical year, e.g. 2025."),
    "est_year": ("name", "Estimate (base) year, e.g. 2026."), "mid_year": ("name", "Mid-forecast year, e.g. 2031."),
    "fc_end": ("name", "Forecast end year, e.g. 2036."),
    "hist_period": ("name", "e.g. 2021-2025 (hyphen, never an en dash)."), "fc_period": ("name", "e.g. 2026-2036."),
    "full_period": ("name", "e.g. 2021-2036."), "mid_period": ("name", "e.g. 2026-2031."), "yoy_period": ("name", "e.g. 2022-2036."),
    "val_unit": ("name", "Value unit, e.g. US$ Mn."), "vol_unit": ("name", "Volume unit, e.g. MT, Units, '000 Units."),
    "price_unit": ("name", "Price unit, e.g. US$/MT."), "page_tag": ("name", "' (1/2)' on split slides, otherwise empty. Never add a tag to a single slide."),
    "metric": ("name", "'volume' or 'value' on slides built for both."), "metric_cap": ("name", "Capitalised metric."),
    "m_unit": ("name", "vol_unit or val_unit, matching metric."),
    "n_top": ("count", "min(10, number of countries in geo)."),
    "segment": ("name", "Segment of the current chapter, e.g. Product Type."), "segment_lc": ("name", "Lower-case segment."),
    "seg1": ("name", "Primary segment, e.g. Product Type."), "seg2": ("name", "Route-to-market segment, e.g. Sales Channel."),
    "item": ("name", "Current item inside a repeat block (sub-segment, region, company ...)."),
    "rank": ("count", "Position of the item in a repeat block, starting at 1 (auto)."),
    "company": ("name", "Company name on profile slides."),
    "n_prof": ("count", "Number of profile slides per company."),
    "tl1": ("name", "Timeline bucket labels for restraint and opportunity slides, e.g. 2021-2022."),
    # core figures
    "v_val_est": ("figure", "Market value in est_year."), "v_val_end": ("figure", "Market value in fc_end."),
    "v_val_mid": ("figure", "Market value in mid_year."), "v_val_hist_start": ("figure", "Market value in hist_start."),
    "v_vol_est": ("figure", "Market volume in est_year."), "v_vol_end": ("figure", "Market volume in fc_end."),
    "v_f_cagr": ("figure", "Value F-CAGR, est_year to fc_end."), "v_h_cagr": ("figure", "Value historical CAGR, hist_start to hist_end."),
    "v_vol_f_cagr": ("figure", "Volume F-CAGR."), "v_vol_h_cagr": ("figure", "Volume historical CAGR."),
    "v_abs_opp": ("figure", "Absolute $ opportunity = value fc_end minus value est_year."),
    "v_abs_hist": ("figure", "Value added hist_start to est_year."),
    "v_abs_opp_mid_share": ("figure", "Share of v_abs_opp realised by mid_year."),
    "v_mult": ("figure", "Value fc_end divided by value est_year, one decimal, without the x (template adds x)."),
    "v_val_vol_gap_bps": ("figure", "Value F-CAGR minus volume F-CAGR, in BPS, unsigned."),
    "v_cr3": ("figure", "Combined value share of the top three companies."),
    # rule inputs
    "r_f_cagr": ("raw", "Value F-CAGR as a fraction (0.052). Drives p_trend and p_price_mix."),
    "r_h_cagr": ("raw", "Value historical CAGR as a fraction."),
    "r_cr3": ("raw", "Top-three company share as a fraction. Drives p_concentration."),
}

PREFIX = {"v_": "figure", "c_": "company", "n_": "count", "s_": "score", "p_": "rule phrase", "lx_": "lexicon", "r_": "raw"}

TOKENS = [
    ("seg1", "primary segment"), ("seg2", "route-to-market segment"), ("sub", "current-chapter segment"),
    ("child", "child geography"), ("cty", "country"), ("xa", "cross-section row segment"), ("xb", "cross-section column segment"),
    ("x", "cross-section cell"), ("pm", "price-matrix cell"), ("price_seg", "primary-segment sub-segment by price"),
    ("lead", "largest by value"), ("fast", "highest F-CAGR"), ("slow", "lowest F-CAGR"), ("gain", "largest share gain"),
    ("lose", "largest share loss"), ("top", "highest"), ("high", "highest"), ("low", "lowest"),
    ("share_est", "share in est_year"), ("share_end", "share in fc_end"), ("cagr", "F-CAGR"), ("opp", "absolute $ opportunity"),
    ("bps", "basis points"), ("idx", "index (base = 100)"), ("gap", "gap vs benchmark"), ("chg", "signed % change"),
    ("mult", "multiple"), ("vol", "volume basis"), ("m", "slide metric basis"), ("item", "repeat item"),
]


def describe(name):
    if name in EXPLICIT:
        return EXPLICIT[name]
    base = name
    for suf in ("_The", "_the"):
        if name.endswith(suf):
            return ("auto", f"{name[:-4]} with 'the' added for countries such as the U.S. or the Netherlands ({'sentence start' if suf == '_The' else 'mid-sentence'}). Built automatically.")
    kind = "name"
    for pre, k in PREFIX.items():
        if name.startswith(pre):
            kind, base = k, name[len(pre):]
            break
    if kind == "rule phrase":
        return (kind, "Phrase chosen by rules.py; see the Rules sheet.")
    if kind == "lexicon":
        return (kind, "Industry wording from lexicon.py; see the Lexicon sheet.")
    parts = [desc for tok, desc in TOKENS if f"_{tok}_" in f"_{base}_" or base.startswith(tok + "_") or base == tok]
    if base.endswith("_lc"):
        parts.append("lower-case form")
    return (kind, ("Pipeline value: " + ", ".join(parts) + ".") if parts else "Pipeline value; see slide context.")
