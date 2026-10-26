"""
FMI auto-generated deck: slide content templates (one entry per slide TYPE).

How to read a template
----------------------
* {placeholder}   value supplied by the pipeline (see placeholders.py for the full dictionary).
* {v_*}           market figure. Masked as "XX" when the deck is rendered in sample mode.
* {c_*}           real company data from secondary sources. Never masked.
* {s_*}           model score (0-10 or 1-5). Never masked.
* {n_*}           count (regions, sub-segments, players). Never masked.
* {p_*}           phrase chosen by a rule in rules.py from the raw numbers (r_*). This is where the
                  "no AI" logic lives: direction words, comparisons and variants are picked by rule,
                  never hard-coded.
* {lx_*}          industry wording from lexicon.py.
* R(over, text)   repeat the text once per item in ctx[over] (each item adds its own keys).

Style rules baked in: no em or en dashes, every bullet is a complete sentence, years written as
ranges with hyphens ({fc_period} = "2026-2036").

COMMON elements (geo pill, source line, copyright footer) are added to every slide automatically.
"""

COMMON = {
    "geo_tag": "{geo}",
    "source": "Source: {source}",
    "footer": "©{pub_year} {publisher}, All Rights Reserved",
}


def R(over, text):
    """Repeat block: render `text` once per item in ctx[over]."""
    return {"repeat": over, "text": text}


SLIDES = []


def S(sid, section, name, sample, applies, title, subtitle, **elements):
    SLIDES.append(dict(id=sid, section=section, name=name, sample=sample, applies=applies,
                       title=title, subtitle=subtitle, elements=elements))


# =============================================================================================
# A. FRONT MATTER
# =============================================================================================

S("A01", "Front matter", "Cover", "1", "deck",
  title="{geo_top} {market} Market",
  subtitle="Value ({val_unit}) and Volume ({vol_unit}) Analysis and Forecast, {full_period}",
  subtitle_value_only="Value ({val_unit}) Analysis and Forecast, {full_period}",
  meta="Region focus: {focus_region} | Country focus: {focus_country} | Published {report_date} | Source: {source}",
)

S("A02", "Front matter", "Agenda", "2", "deck",
  title="Agenda",
  subtitle="Storyline of the report by section; select a section card to jump to it.",
  group_1="Market Overview",
  group_1_items=["{sec_signals}. Market Signals", "{sec_taxonomy}. Market Taxonomy",
                 "{sec_toc}. Table of Contents", "{sec_lof}. List of Figures and List of Tables"],
  group_2="{geo_top} Market Outlook",
  group_2_items=["{sec_geo_overview}. {geo_top} {market} Market Overview",
                 "{sec_geo_region}. {geo_top} Market Analysis by Region",
                 "{sec_geo_segment}. {geo_top} Market Analysis by Segment",
                 "{sec_cross}. Cross-Sectional Assessment"],
  group_3="{focus_region} Market Outlook",
  group_3_items=["{sec_reg_overview}. {focus_region} Market Overview",
                 "{sec_reg_country}. {focus_region} Market Analysis by Country",
                 "{sec_reg_segment}. {focus_region} Market Analysis by Segment"],
  group_4="{focus_country} Market Outlook",
  group_4_items=["{sec_cty_overview}. {focus_country} Market Overview (Volume and Value)",
                 "{sec_cty_segment}. {focus_country} Segment Analysis (Volume and Value)"],
  group_5="Pricing and Supply Chain Assessment",
  group_5_items=["{sec_pricing}. {market} Market Pricing Analysis", "{sec_supply}. Supply Chain Analysis"],
  group_6="Market Dynamics",
  group_6_items=["{sec_dynamics}. Market Dynamics and Forces"],
  group_7="Competition Assessment",
  group_7_items=["{sec_comp}. Competitive Landscape", "{sec_share}. Market Share of Leading Players",
                 "{sec_profiles}. Company Profiles and Strategic Benchmarking"],
)

S("A03", "Front matter", "Section divider", "3, 6, 11, 24, 34 ...", "deck",
  title="{section_title}",
  subtitle="",
  section_bar="Section {sec_no}",
  kicker="{market} Market",
  contents=R("sub_sections", "➤ {item}"),
)

S("A04", "Front matter", "Part breaker", "5, 33, 250, 443, 534 ...", "deck",
  title="{part_title}",
  subtitle="",
  kicker="{market} Market",
  contents=R("part_sections", "➤ {item}"),
)

S("A05", "Front matter", "Market Signals", "4", "global",
  title="Market Signals Reshaping {market} Demand",
  subtitle="The questions clients are asking, the evidence behind each answer and what it implies for the {geo_ref} market",
  table_headers=["QUESTIONS WE ARE HEARING", "WHAT WE ARE OBSERVING", "SELECT DATA POINTS"],
  row1_q="How fast is {market_lc} demand set to grow through {fc_end}?",
  row1_o="{geo_ref_cap} demand is forecast to expand at {v_f_cagr} CAGR over {fc_period}, {p_trend} the {v_h_cagr} recorded over {hist_period}.",
  row1_d="F-CAGR: {v_f_cagr} | H-CAGR: {v_h_cagr}",
  row2_q="Which {seg1_lc} categories are buyers prioritising?",
  row2_o="'{seg1_lead}' leads the {seg1_lc} segment with {v_seg1_lead_share_end} of {fc_end} value across {n_seg1_sub} tracked sub-segments{p_seg1_fast_clause}.",
  row2_d="Lead: {seg1_lead} | {v_seg1_lead_share_end} share ({fc_end})",
  row3_q="Where is incremental demand concentrating?",
  row3_o="'{seg2_fast}' is the fastest-growing {seg2_lc} sub-segment at {v_seg2_fast_cagr} CAGR, {p_seg2_fast_vs_market}.",
  row3_d="Fastest: {seg2_fast} | {v_seg2_fast_cagr} F-CAGR",
  row4_q="Which {child_label_pl} will drive the next wave of {market_lc} demand?",
  row4_o="{p_child_lead_fast}",
  row4_d="Largest: {child_lead} ({v_child_lead_share_end}) | Fastest: {child_fast} ({v_child_fast_cagr})",
  row5_q="Where should suppliers focus investment?",
  row5_o="The market adds {v_abs_opp} {val_unit} of incremental value over {fc_period}, of which {v_abs_opp_mid_share} is realised by {mid_year}, and the leading categories and {child_label_pl} above capture most of it.",
  row5_d="$ opportunity: {v_abs_opp} {val_unit} ({fc_period}) | {v_abs_opp_mid_share} by {mid_year}",
)

S("A06", "Front matter", "Market Taxonomy: Regions", "7", "deck",
  title="Market Taxonomy: Regional Coverage",
  subtitle="Geographic scope of the study: the global market split into {n_regions} regions and the {n_countries} countries assessed within them",
  root="{market} Market",
  region_box="{item}",
  country_box="{item}",
)

S("A07", "Front matter", "Market Taxonomy: Segments", "8-9", "deck",
  title="Market Taxonomy: Segmentation{page_tag}",
  subtitle="Segmentation framework of the study: {n_segments} segments and the sub-segments sized within each",
  segment_header="By {item}",
  level1="❑ {item}",
  level2="➢ {item}",
)

S("A08", "Front matter", "Company Coverage", "10", "deck",
  title="{market} Market: Company Coverage",
  subtitle="Companies profiled and benchmarked in the Competition Assessment section",
  intro="The {n_companies} companies below are profiled and benchmarked in the competition assessment of the {market} Market, with key brands shown in brackets.",
  column_header="Companies Covered",
  company_line=R("companies", "❑ {item} ({brands})"),
  note="Note: The company list is indicative and can be customised on request.",
)

S("A09", "Front matter", "Table of Contents", "12-23", "deck",
  title="Table of Contents{page_tag}",
  subtitle="",
  toc_patterns=[
      "Executive Summary",
      "{geo_top} Market Snapshot, {est_year} and {fc_end}",
      "Market Overview: Definition, Scope and Taxonomy",
      "Research Methodology and Assumptions",
      "{geo_top} {market} Market Demand (Volume, {vol_unit}) Analysis, {hist_period} and Forecast, {fc_period}",
      "{geo_top} {market} Market Pricing Analysis ({price_unit}), {full_period}",
      "{geo_top} {market} Market Value ({val_unit}) Analysis, {hist_period} and Forecast, {fc_period}",
      "Market Background: Macro-Economic Factors, Value Chain and Market Dynamics",
      "{geo_top} {market} Market Analysis {hist_period} and Forecast {fc_period}, by {item}",
      "{geo_top} {market} Market Analysis {hist_period} and Forecast {fc_period}, by Region",
      "{item} {market} Market Analysis {hist_period} and Forecast {fc_period}",
      "{item} {market} Market Analysis by Country",
      "Market Structure Analysis: Tier Structure and Market Concentration",
      "Competition Analysis: Competition Dashboard and Company Profiles",
      "{item}: Overview, Product Portfolio, Key Financials, SWOT and Strategy",
      "Assumptions and Acronyms Used",
  ],
)

S("A10", "Front matter", "List of Figures", "25-28", "deck",
  title="List of Figures{page_tag}",
  subtitle="",
  figure_patterns=[
      "Figure {n}: {geo} Market Value and YoY Growth, {val_unit} and %, {full_period}",
      "Figure {n}: {geo} Market Volume and YoY Growth, {vol_unit} and %, {full_period}",
      "Figure {n}: {geo} Market Value Snapshot, {val_unit}, {est_year} and {fc_end}",
      "Figure {n}: {geo} Average Price by {item}, {price_unit}, {full_period}",
      "Figure {n}: {geo} Market Value Share and BPS Analysis by {item}, {est_year} and {fc_end}",
      "Figure {n}: {geo} Market Absolute $ Opportunity by {item}, {val_unit}, {fc_period}",
      "Figure {n}: {geo} Market Attractiveness by {item}, {fc_period}",
      "Figure {n}: {geo} Market Value by {item}, {val_unit}, {full_period}",
      "Figure {n}: {geo} Market Value Share by Company, %, {est_year}",
  ],
)

S("A11", "Front matter", "List of Tables", "29-32", "deck",
  title="List of Tables{page_tag}",
  subtitle="",
  table_patterns=[
      "Table {n}: {geo} Market Snapshot: Value, Volume and CAGR, {val_unit}, {est_year} and {fc_end}",
      "Table {n}: {geo} Market Taxonomy by Segment, {est_year}",
      "Table {n}: {geo} Market Volume by {item}, {vol_unit}, {full_period}",
      "Table {n}: {geo} Market Value by {item}, {val_unit}, {full_period}",
      "Table {n}: {geo} Market YoY Growth by {item}, %, {yoy_period}",
      "Table {n}: {geo} Market Value by {item} and {item2}, {val_unit}, {est_year} and {fc_end}",
      "Table {n}: {geo} Average Price by {item}, {price_unit}, {full_period}",
      "Table {n}: {item} Key Financials, {c_currency_unit}, FY{c_y1} to FY{c_y5}",
  ],
)

# =============================================================================================
# G. GEOGRAPHY OVERVIEW (use with geo = Global or a region)
# =============================================================================================

S("G01", "Geo overview", "Market Summary by Volume", "35, 252", "global, region",
  title="{geo} {market} Market Summary, By Volume ({vol_unit})",
  subtitle="Volume growth, incremental demand and the {child_label_pl} driving it, {full_period}",
  kpi_labels=["{est_year} MARKET SIZE ({vol_unit})", "{fc_end} MARKET SIZE ({vol_unit})",
              "H-CAGR ({hist_period})", "F-CAGR ({fc_period})"],
  chart1_title="Volume CAGR by Period",
  chart1_categories=["{hist_start}-{est_year}", "{fc_period}", "{full_period}"],
  chart2_title="Incremental Volume Addition ({vol_unit}), {full_period}",
  insights_head="Key Insights",
  insights=[
      "{geo} volume is forecast to grow at {v_vol_f_cagr} CAGR over {fc_period}, {p_vol_trend} the {v_vol_h_cagr} recorded over {hist_period}.",
      "Volume rises from {v_vol_est} {vol_unit} in {est_year} to {v_vol_end} {vol_unit} in {fc_end}, a {v_vol_mult}x expansion over the forecast decade.",
      "The forecast decade adds {v_vol_abs_fc} {vol_unit}, {p_vol_period_compare} the {v_vol_abs_hist} {vol_unit} added over {hist_start}-{est_year}.",
      "{child_fast} leads growth among {child_label_pl} at {v_child_fast_vol_cagr} CAGR, while {child_lead} remains the largest volume pool.",
  ],
  drivers_head="Growth Drivers",
  drivers=[
      "Demand is concentrating in '{seg1_lead}', which holds {v_seg1_lead_vol_share} of {seg1_lc} volume in {est_year}.",
      "Wider adoption among {lx_end_user} continues to broaden the volume base across {n_children} {child_label_pl}.",
      "Ongoing investment in {lx_innovation} is {lx_demand_effect}.",
  ],
)

S("G02", "Geo overview", "Pricing and Price Index", "36, 253", "global, region",
  title="{geo} {market} Market Pricing and {child_label_adj_cap} Price Index",
  subtitle="Average price by {seg1_lc} and price variation by {child_label} against the {geo} average, {fc_period}",
  chart1_title="Average Price by {seg1}: Range across {full_period} ({price_unit})",
  chart2_title="Price Variation by {child_label_cap} vs {geo} Average, {fc_period}",
  takeaways_head="Key Takeaways",
  takeaways=[
      "Average selling price is derived as value divided by volume, and '{price_seg_top}' commands the highest {seg1_lc} price at {v_price_seg_top_end} {price_unit} in {fc_end}.",
      "'{price_seg_fast}' records the steepest price uplift over {fc_period}, at {v_price_seg_fast_chg}.",
      "Against the {geo} average of {v_price_avg_est} {price_unit}, {n_child_above_price} of {n_children} {child_label_pl} price above the average and {n_child_below_price} price below it.",
      "{child_price_high} is the highest-priced {child_label} at {v_child_price_high_gap} versus the {geo} average, while {child_price_low} is the lowest at {v_child_price_low_gap}, which points to premiumisation headroom in lower-priced markets.",
  ],
)

S("G03", "Geo overview", "Market Summary by Value", "37, 254", "global, region",
  title="{geo} {market} Market Summary, By Value ({val_unit})",
  subtitle="Value growth, incremental opportunity and the {child_label_pl} driving it, {full_period}",
  kpi_labels=["{est_year} MARKET SIZE ({val_unit})", "{fc_end} MARKET SIZE ({val_unit})",
              "H-CAGR ({hist_period})", "F-CAGR ({fc_period})"],
  chart1_title="Value CAGR by Period",
  chart1_categories=["{hist_start}-{est_year}", "{fc_period}", "{full_period}"],
  chart2_title="Incremental Value Addition ({val_unit}), {full_period}",
  insights_head="Key Insights",
  insights=[
      "{geo} value is forecast to grow at {v_f_cagr} CAGR over {fc_period}, {p_trend} the {v_h_cagr} recorded over {hist_period}.",
      "Value rises from {v_val_est} {val_unit} in {est_year} to {v_val_end} {val_unit} in {fc_end}, a {v_mult}x expansion over the forecast decade.",
      "The forecast decade adds {v_abs_opp} {val_unit}, {p_val_period_compare} the {v_abs_hist} {val_unit} added over {hist_start}-{est_year}.",
      "{child_fast} leads growth among {child_label_pl} at {v_child_fast_cagr} CAGR, while {child_lead} remains the largest value pool.",
  ],
  drivers_head="Growth Drivers",
  drivers=[
      "{p_price_mix}",
      "Investment in {lx_premium_driver} is lifting average realisation across the {seg1_lc} mix.",
      "Demand is concentrating in '{seg1_lead}', which holds {v_seg1_lead_share_est} of {seg1_lc} value in {est_year}.",
  ],
)

S("G04", "Geo overview", "Volume and Value Analysis", "38, 255", "global, region",
  title="{geo} {market} Market Volume and Value Analysis",
  subtitle="Historical and forecast value and volume, and the incremental $ opportunity they create, {full_period}",
  chart1_title="{geo} Market Value, {hist_period} (Historical)",
  chart2_title="{geo} Market Value, {fc_period} (Forecast)",
  kpi_labels=["VALUE H-CAGR ({hist_period})", "VALUE F-CAGR ({fc_period})",
              "INCREMENTAL $ OPPORTUNITY ({mid_period})", "INCREMENTAL $ OPPORTUNITY ({fc_period})"],
  readout_head="{geo} Volume and Value Analysis",
  readout=[
      "Value expands from {v_val_est} {val_unit} in {est_year} to {v_val_end} {val_unit} by {fc_end} at {v_f_cagr} CAGR, {p_trend} the {v_h_cagr} recorded over {hist_period}.",
      "Volume rises from {v_vol_est} {vol_unit} to {v_vol_end} {vol_unit} over the same period, at {v_vol_f_cagr} CAGR.",
      "{p_price_mix}",
      "The forecast decade adds {v_abs_opp} {val_unit} of incremental opportunity, {v_abs_opp_mid_share} of which is realised in the first five years, taking the market to {v_mult}x its {est_year} value.",
  ],
)

S("G05", "Geo overview", "Country Indices (Value)", "39-40, 256", "global, region",
  title="{geo} {market} Market Summary: Country Indices (Value){page_tag}",
  subtitle="Countries ranked on {fc_end} market size, growth index and absolute-opportunity index",
  chart1_title="Top {n_top} Countries by CAGR Index",
  chart2_title="{fc_end} Market by Country ({val_unit}): Top {n_top}",
  chart3_title="Top {n_top} Countries by Absolute $ Opportunity Index",
  table1_headers=["Country", "CAGR Index", "F-CAGR", "$ Opp ({val_unit})"],
  table2_headers=["Country", "Absolute Index", "F-CAGR", "$ Opp ({val_unit})"],
  insight="{p_index_leaders}",
  index_note="Note: CAGR Index = country F-CAGR relative to the {geo} F-CAGR (= 100); Absolute Index = country $ opportunity relative to the average country $ opportunity (= 100).",
)

S("G06", "Geo overview", "Market Overview (executive summary)", "41, 257", "global, region",
  title="{geo} {market} Market Overview",
  subtitle="Executive summary of market size, competitive structure and the drivers, restraints and trends shaping demand",
  overview_head="Market Overview",
  overview=[
      "The {geo_ref} {market_lc} market is valued at {v_val_est} {val_unit} in {est_year} and is forecast to reach {v_val_end} {val_unit} by {fc_end}.",
      "This implies a {v_f_cagr} CAGR over {fc_period}, {p_trend} the {v_h_cagr} recorded over {hist_period}.",
      "The forecast period creates {v_abs_opp} {val_unit} of incremental demand, a {v_mult}x expansion on the {est_year} base.",
  ],
  competition_head="Competitive Landscape",
  competition=[
      "{p_competition_line}",
      "Demand is assessed across {n_children} {child_label_pl} and {n_segments} segments, covering {n_seg1_sub} {seg1_lc} categories.",
      "{child_lead} is the largest {child_label} at {v_child_lead_share_end} of {fc_end} value.",
  ],
  drivers_head="Drivers",
  drivers=[
      "Demand for '{seg1_fast}' is expanding fastest within {seg1_lc}, at {v_seg1_fast_cagr} CAGR.",
      "{child_fast} is the fastest-growing {child_label}, expanding at {v_child_fast_cagr} CAGR.",
      "'{seg2_fast}' is gaining momentum within {seg2_lc}, growing at {v_seg2_fast_cagr} CAGR.",
  ],
  restraints_head="Restraints",
  restraints=[
      "'{seg1_slow}' grows at only {v_seg1_slow_cagr} CAGR, which dilutes blended growth within {seg1_lc}.",
      "{child_slow} trails the market at {v_child_slow_cagr} CAGR, which limits upside from part of the established base.",
      "Costs of {lx_input_costs} remain volatile and can compress margins when increases cannot be passed through.",
  ],
  trends_head="Trends",
  trends=[
      "{p_seg1_share_shift}",
      "{child_fast} is outpacing the {geo} average by {v_child_fast_gap_bps} BPS of CAGR.",
      "Demand is broadening across {n_seg1_sub} {seg1_lc} categories, supported by {lx_innovation}.",
  ],
)

S("G07", "Geo overview", "Supply-Side Trends", "42, 258", "global, region",
  title="{geo} {market} Market: Supply-Side Trends",
  subtitle="Recent developments, investment flows and R&D activity shaping the supply side of the market",
  developments_head="Recent Developments and Key Industry Activities",
  developments=[
      "Capacity and product development are concentrating in '{seg1_fast}' and in {child_fast}, the fastest-growing {child_label}.",
      "Portfolio expansion is tracking the shift toward higher-growth {seg1_lc} categories.",
      "{p_comm_focus}",
  ],
  insights_head="Key Insights",
  insights=[
      "Market value rises from {v_val_est} {val_unit} in {est_year} to {v_val_end} {val_unit} in {fc_end}.",
      "The forecast CAGR is {v_f_cagr} over {fc_period}.",
      "'{seg1_lead}' is the lead {seg1_lc} category with {v_seg1_lead_share_end} of {fc_end} value.",
      "{child_lead} is the lead {child_label} with {v_child_lead_share_end} of {fc_end} value.",
      "The incremental opportunity over {fc_period} is {v_abs_opp} {val_unit}.",
  ],
  invest_head="Strategic Acquisitions and New Investments",
  invest=[
      "Investment is flowing toward '{seg1_fast}', the fastest-growing {seg1_lc} category at {v_seg1_fast_cagr} CAGR.",
      "{child_fast} ({v_child_fast_cagr} CAGR) is a priority {child_label} for new capacity and distribution.",
      "Scale players are anchoring their portfolios around '{seg1_lead}', which holds {v_seg1_lead_share_end} share.",
  ],
  rd_head="Substantial Investment in R&D",
  rd=[
      "R&D programmes are focused on {lx_innovation}, in line with the migration toward '{seg1_fast}'.",
      "Development spend is targeting the {seg1_lc} categories that gain share through {fc_end}.",
      "Innovation underpins the {v_f_cagr} forecast CAGR by sustaining both volume growth and price realisation.",
  ],
  chart_title="Market Structure: {seg1} Share ({fc_end})",
  note="Note: Supply-side observations are derived from the FMI market model and company tracking and are indicative.",
)

S("G08", "Geo overview", "Analysis and Recommendations", "43, 259", "global, region",
  title="{geo} {market} Market: Analysis and Recommendations",
  subtitle="Overall market approach, differentiating strategy and {publisher_short}'s recommendations",
  approach_head="Overall Market Approach",
  approach=["{p_approach}",
            "The market adds {v_abs_opp} {val_unit} of demand to {fc_end}, a {v_mult}x expansion on the {est_year} base."],
  differentiate_head="Differentiating Strategy",
  differentiate=["Differentiate through the high-growth '{seg1_fast}' category and the '{seg2_fast}' {seg2_lc} route, where growth outpaces the market."],
  segment_head="Market Segment",
  segment=["{p_segment_priority}"],
  region_head="Market {child_label_cap}",
  region=["{p_region_priority}"],
  reco_head="{publisher_short} Analysis and Recommendation",
  reco=[
      "Concentrate resources on the leading segment and {child_label} while seeding the fastest-growing pockets early.",
      "Defend share in '{seg1_lead}' as the value anchor of the portfolio.",
      "Build {lx_validation} into the offer to protect pricing as competition intensifies.",
  ],
)

S("G09", "Geo overview", "Scenario Outlook", "44, 260", "global, region",
  title="{geo} {market} Market: Scenario Outlook",
  subtitle="Optimistic, likely and conservative outcomes for {fc_end} and the CAGR behind each",
  chart1_title="{geo} {market} Market Value ({val_unit}): Scenario Forecast, {est_year} and {fc_end}",
  chart2_title="{geo} {market} Market Value by Scenario ({val_unit}), {fc_period}",
  legend=["Conservative Scenario", "Likely Scenario", "Optimistic Scenario"],
  cards=["OPTIMISTIC ({fc_end})", "LIKELY ({fc_end})", "CONSERVATIVE ({fc_end})"],
  readout_head="Scenario Assumptions and Outlook",
  readout=[
      "Optimistic: the market reaches {v_scn_opt} {val_unit} by {fc_end} at {v_scn_opt_cagr} CAGR, assuming faster adoption and stronger price realisation.",
      "Likely (base forecast): the market reaches {v_scn_base} {val_unit} by {fc_end} at {v_scn_base_cagr} CAGR.",
      "Conservative: the market reaches {v_scn_con} {val_unit} by {fc_end} at {v_scn_con_cagr} CAGR, assuming slower adoption and tighter pricing.",
      "Scenarios flex the base forecast's year-on-year growth by plus or minus {v_scn_flex} across {fc_period}.",
      "The gap between the optimistic and conservative outcomes is {v_scn_spread} {val_unit} in {fc_end}, which frames the planning range for capacity and investment decisions.",
  ],
)

# =============================================================================================
# B. BREAKDOWN BY CHILD GEOGRAPHY (regions under Global, countries under a region)
# =============================================================================================

S("B01", "Geo breakdown", "Regional / Country Outlook", "46, 262", "global, region",
  title="{geo} {market} Market: {child_label_cap} Outlook",
  subtitle="Size, growth and share across {n_children} {child_label_pl}, highlighting where value concentrates, {fc_period}",
  intro="The by-{child_label} view breaks the {geo} market down geographically. Each {child_label} is assessed on market size, growth outlook and share of value across the forecast period.",
  table_headers=["{child_label_cap}", "{est_year} Value ({val_unit})", "{fc_end} Value ({val_unit})", "CAGR ({fc_period})"],
  highlights_head="Key Highlights",
  highlights=[
      "{child_lead} holds the largest share in {est_year}, at {v_child_lead_share_est} of market value.",
      "{p_child_lead_fast_short}",
      "Market value expands from {v_val_est} {val_unit} to {v_val_end} {val_unit} across {fc_period}.",
      "{p_child_landscape}",
  ],
  insights_head="{child_label_cap} Insights",
  insights=R("children", "● {item}: {v_item_share_est} of {est_year} value and {v_item_cagr} CAGR; {p_item_role}."),
  overview_head="{child_label_cap} Overview",
  overview=[
      "✓ The by-{child_label} split shows where market value concentrates and where growth is fastest.",
      "✓ {p_above_pace}",
  ],
  closing="The {geo} market is projected to grow at a CAGR of {v_f_cagr} from {est_year} to {fc_end}.",
)

S("B02", "Geo breakdown", "Share and BPS by Child Geo (Volume / Value)", "47-48, 263-264", "global, region",
  title="{geo} {market} Market Analysis by {child_label_cap}: {metric_cap} ({m_unit})",
  subtitle="{metric_cap} share by {child_label} in {est_year} and {fc_end}, with share gains and losses in basis points",
  chart_title="Market Share and BPS Assessment by {child_label_cap}, {est_year} and {fc_end}",
  bps_head="BPS Change ({est_year} to {fc_end})",
  kpi_labels=["{est_year} TOTAL ({m_unit})", "{fc_end} TOTAL ({m_unit})", "CAGR ({fc_period})"],
  readout_head="{child_label_cap} Highlights",
  readout=[
      "{child_lead_m} is the largest {child_label}, at {v_child_lead_m_share_end} of {fc_end} {metric}.",
      "{p_child_share_shift_m}",
      "{child_fast_m} grows fastest at {v_child_fast_m_cagr} CAGR over {fc_period}.",
      "Total {metric} rises from {v_m_est} to {v_m_end} {m_unit} across {n_children} {child_label_pl}.",
  ],
)

S("B03", "Geo breakdown", "Size and Share Table by Child Geo", "49-50, 265-266", "global, region",
  title="{geo} {market} Market Size and Share by {child_label_cap}: {metric_cap} ({m_unit})",
  subtitle="Year-by-year {metric} by {child_label}, {full_period}, highlighting the {child_label_pl} gaining and ceding share",
  table_title="{geo} Market Size by {child_label_cap} ({m_unit})",
  readout_head="{geo} {metric_cap} Share Read-out by {child_label_cap}",
  readout=[
      "{child_lead_m} is the largest {child_label} with {v_child_lead_m_share_end} of {fc_end} {metric} ({v_child_lead_m_end} {m_unit}).",
      "{child_fast_m} grows fastest at {v_child_fast_m_cagr} CAGR over {fc_period}.",
      "{p_child_share_shift_m}",
      "Total {metric} rises from {v_m_est} to {v_m_end} {m_unit} across {n_children} {child_label_pl}.",
  ],
)

S("B04", "Geo breakdown", "Trends by Child Geo", "51, 267", "global, region",
  title="{geo} {market} Market Trends by {child_label_cap}",
  subtitle="How {child_label_adj} volume and value shares shift over {full_period}, highlighting who gains ground",
  chart1_title="{geo} Volume by {child_label_cap} ({vol_unit})",
  chart2_title="{geo} Value by {child_label_cap} ({val_unit})",
  takeaway="{child_gain} gains the most value share over {full_period}, while {child_lose} cedes the most, as demand shifts toward faster-growing {child_label_pl}.",
)

S("B05", "Geo breakdown", "Top Countries Treemap", "52, 268", "global, region",
  title="{geo} {market} Market: Top {n_top} Countries by Value ({fc_end})",
  subtitle="Country hierarchy of the {geo_ref} market in {fc_end}, with tile size proportional to market value",
  chart_title="{geo} Market by Country, {fc_end} ({val_unit}); tile size = market value",
  takeaway="The top {n_top} countries account for {v_topn_share} of {geo} value in {fc_end}, led by {cty_lead_the} at {v_cty_lead_share}.",
  rule_note="Use only when the geography has 3 or more countries; otherwise drop the slide (rule: n_countries >= 3).",
)

# =============================================================================================
# D. SEGMENT CHAPTER (repeat per segment, at Global and regional level)
# =============================================================================================

S("D01", "Segment chapter", "Segment Outlook", "55, 67, 79-80 ...", "global, region",
  title="{geo} {market} Market Outlook by {segment}{page_tag}",
  subtitle="Size, growth and share across {n_sub} sub-segments of {segment}, {fc_period}",
  intro="The {segment} segment classifies the market by {segment_lc}. Each sub-segment is assessed on market size, growth outlook and share of value across the forecast period.",
  table_headers=["Sub-Segment", "{est_year} Value ({val_unit})", "{fc_end} Value ({val_unit})", "CAGR ({fc_period})"],
  highlights_head="Key Highlights",
  highlights=[
      "'{sub_lead}' holds the largest share in {est_year}, at {v_sub_lead_share_est} of segment value.",
      "{p_sub_lead_fast}",
      "Segment value expands from {v_val_est} {val_unit} to {v_val_end} {val_unit} across {fc_period}.",
      "{p_sub_landscape}",
  ],
  insights_head="Sub-Segment Insights",
  insights=R("subs", "● {item}: {v_item_share_est} of {est_year} value and {v_item_cagr} CAGR; {p_item_role}."),
  overview_head="Segment Overview",
  overview=[
      "✓ The {segment} split shows how demand is distributed across {segment_lc} categories and where growth concentrates.",
      "✓ {p_sub_share_shift}",
  ],
  closing="'{sub_fast}' is projected to grow at {v_sub_fast_cagr} CAGR from {est_year} to {fc_end}, against {v_f_cagr} for the {geo_ref} market.",
)

S("D02", "Segment chapter", "Market Structure and Attractiveness", "56, 68, 81 ...", "global, region",
  title="{geo} Market Structure and Attractiveness by {segment} (Value)",
  subtitle="{segment} mix within each {child_label} and the attractiveness of each sub-segment on growth and size",
  chart1_title="{segment} Mix within Each {child_label_cap} (Share of {fc_end} Value); row label = {child_label} size ({val_unit}), largest first",
  chart2_title="Market Attractiveness by {segment}; bubble size = {fc_end} value",
  insight=[
      "{p_sub_share_shift}",
      "'{sub_lead}' holds its highest share in {sub_lead_peak_child} at {v_sub_lead_peak_share} of {fc_end} value.",
      "Bubbles to the right of the CAGR Index line grow faster than the segment average.",
  ],
)

S("D03", "Segment chapter", "Absolute $ Opportunity", "57, 69, 82-83, 106 ...", "global, region",
  title="{geo} Absolute $ Opportunity Analysis by {segment}{page_tag}",
  subtitle="Value build-up by {segment_lc} and the absolute $ opportunity each sub-segment adds, {fc_period}",
  chart_title=R("subs_on_page", "{geo} Incremental $ Opportunity ({val_unit}), {fc_period}: {item}"),
  summary_head="Summary: {geo} Absolute $ Opportunity by {segment}",
  summary=[
      "Across its {n_sub} sub-segments, the {segment_lc} segment adds {v_abs_opp} {val_unit} to the {geo_ref} market between {est_year} and {fc_end}.",
      R("subs_top3_opp", "'{item}' adds {v_item_opp} {val_unit}, or {v_item_opp_share} of the increment."),
      "{p_sub_opp_tail}",
  ],
  rule_note="Place the summary box on the last page of the chapter. List the top min(3, n_sub - 1) sub-segments; p_sub_opp_tail names the smallest only when n_sub >= 4.",
)

S("D04", "Segment chapter", "Volume Table", "58-59, 70-71 ...", "global, region",
  title="{geo} Market Volume Analysis by {segment} ({vol_unit}){page_tag}",
  subtitle="Year-by-year volume by {segment_lc}, {full_period}, highlighting the sub-segments driving growth",
  table_title="{geo} Volume by {segment} ({vol_unit})",
  readout_head="{geo} Volume Analysis by {segment}",
  readout=[
      "'{sub_lead_vol}' is the largest {segment_lc} category at {v_sub_lead_vol_share_end} of {fc_end} volume ({v_sub_lead_vol_end} {vol_unit}).",
      "'{sub_fast_vol}' grows fastest at {v_sub_fast_vol_cagr} CAGR over {fc_period}.",
      "{p_sub_vol_share_shift}",
      "Total volume rises from {v_vol_est} to {v_vol_end} {vol_unit} across {n_sub} sub-segments.",
  ],
)

S("D05", "Segment chapter", "Value Table", "60, 72 ...", "global, region",
  title="{geo} Market Value Analysis by {segment} ({val_unit}){page_tag}",
  subtitle="Year-by-year value by {segment_lc}, {full_period}, highlighting the sub-segments driving growth",
  table_title="{geo} Value by {segment} ({val_unit})",
  readout_head="{geo} Value Analysis by {segment}",
  readout=[
      "'{sub_lead}' is the largest {segment_lc} category at {v_sub_lead_share_end} of {fc_end} value ({v_sub_lead_end} {val_unit}).",
      "'{sub_fast}' grows fastest at {v_sub_fast_cagr} CAGR over {fc_period}.",
      "{p_sub_share_shift}",
      "Total value rises from {v_val_est} to {v_val_end} {val_unit} across {n_sub} sub-segments.",
  ],
)

S("D06", "Segment chapter", "Value and Share Shift", "61, 73 ...", "global, region",
  title="{geo} Market Value and Share Shift by {segment}",
  subtitle="Value composition by {segment_lc} in {hist_start}, {est_year} and {fc_end}, with share shifts in basis points",
  chart_title="{geo} Value Composition by {segment}: {hist_start}, {est_year} and {fc_end} ({val_unit}); right axis = BPS share shift",
  bps_head="BPS Change ({est_year} to {fc_end})",
  readout_head="{geo} Read-out by {segment}",
  readout=[
      "{p_sub_share_shift}",
      "'{sub_fast}' is the fastest-growing {segment_lc} category at {v_sub_fast_cagr} CAGR over {fc_period}.",
      "Total value climbs from {v_val_hist_start} {val_unit} in {hist_start} to {v_val_end} {val_unit} in {fc_end}.",
      "Bars show the value mix at each snapshot, and the scale shows how each share moves.",
  ],
)

S("D07", "Segment chapter", "YoY Growth Table (Volume / Value)", "62, 64, 99-100 ...", "global, region",
  title="{geo} Market {metric_cap} YoY Growth by {segment} (%){page_tag}",
  subtitle="Year-on-year {metric} growth by {segment_lc}, {yoy_period}, highlighting acceleration and slowdown",
  table_title="{geo} {metric_cap} YoY Growth by {segment} (%): ▲ outpacing the market | ▶ below market pace | ▼ decline",
  insight="'{sub_yoy_peak}' records the highest single-year growth of {v_sub_yoy_peak} in {sub_yoy_peak_year}, {p_yoy_decline}.",
)

S("D08", "Segment chapter", "YoY Heatmap (Volume / Value)", "63, 65 ...", "global, region",
  title="{geo} Market {metric_cap} YoY Growth by {segment} (%){page_tag}",
  subtitle="Year-on-year {metric} growth by {segment_lc}, {yoy_period}, highlighting acceleration and slowdown",
  chart_title="{geo} {metric_cap} YoY Heatmap by {segment} (%): lighter = lower growth, darker = higher growth",
  insight="Growth across {segment_lc} peaks in {yoy_peak_year} and {p_yoy_trend}.",
)

# =============================================================================================
# X. CROSS-SECTIONAL ASSESSMENT (segment A x segment B)
# =============================================================================================

S("X01", "Cross-section", "Cross-Section Snapshot", "234, 238, 242, 247", "global",
  title="{geo} Market Snapshot: {seg_a} × {seg_b}, {est_year}E vs {fc_end}F",
  subtitle="{geo} market value ({val_unit}) for each {seg_a_lc} sub-segment across {seg_b_lc} sub-segments",
  table_title="{geo} Market Value: {seg_a} × {seg_b}, {est_year}E and {fc_end}F ({val_unit})",
  findings_head="Key Findings",
  findings=[
      "'{xa_lead}' accounts for the largest share of the {geo_ref} market by {seg_a_lc}, at {v_xa_lead_val} {val_unit} ({v_xa_lead_share}) in {est_year}E, and '{xb_lead_in_a}' is its leading {seg_b_lc} sub-segment at {v_xb_lead_in_a_val} {val_unit}.",
      "'{x_lead_cell}' is the largest single cell at {v_x_lead_cell_val} {val_unit} in {est_year}E, while '{x_fast_cell}' records the fastest growth at {v_x_fast_cell_cagr} CAGR over {est_year}E-{fc_end}F.",
      "The top three {seg_a_lc} × {seg_b_lc} cells generate a combined absolute $ opportunity of {v_x_top3_opp} {val_unit} by {fc_end}F.",
  ],
  rule_note="Columns must be leaf sub-segments of {seg_b}. Never use a grouping header as a cell; cells with zero value by definition must be excluded from all 'largest' and 'fastest' picks.",
)

S("X02", "Cross-section", "Cross-Section Share", "235, 239, 243-244, 248", "global",
  title="{geo} Market Share Analysis: {seg_a} × {seg_b}, {est_year}E{page_tag}",
  subtitle="Share of {geo_ref} market value, with each {seg_a_lc} row split across {seg_b_lc} sub-segments",
  table_title="{geo} Value Share: {seg_a} × {seg_b}, {est_year}E (% of each row)",
  chart_title="{seg_b} Mix by {seg_a}, {est_year}E (% share)",
  insights_head="Share Insights",
  insights=[
      "'{x_top_share_b}' holds the highest share of '{x_top_share_a}' at {v_x_top_share} in {est_year}E.",
      "{p_x_concentration}",
      "'{x_gain_b}' gains {v_x_gain_bps} BPS of cross-sectional share by {fc_end}F.",
  ],
)

S("X03", "Cross-section", "Cross-Section Opportunity and Attractiveness", "236, 240, 245, 249", "global",
  title="{geo} Absolute $ Opportunity and Attractiveness: {seg_a} × {seg_b}, {est_year}E-{fc_end}F",
  subtitle="Incremental value ({val_unit}) and value CAGR by cell, with an attractiveness index scored 1 to 5",
  table_title="Incremental $ Opportunity ({val_unit}) and Value CAGR: {seg_a} × {seg_b}, {est_year}E-{fc_end}F",
  top5_head="Top 5 Opportunity Cells ({val_unit} | CAGR | Index)",
  top5=R("x_top_cells", "{rank}. {item}: {v_item_opp} {val_unit} | {v_item_cagr} | {s_item_idx}/5"),
  pool_head="Largest Pool",
  pool=["'{x_opp_lead}' adds {v_x_opp_lead} {val_unit} by {fc_end}F, or {v_x_opp_lead_share} of the total opportunity."],
  fastest_head="Fastest Growth",
  fastest=["'{x_fast_cell}' grows at {v_x_fast_cell_cagr} CAGR, {v_x_fast_gap_bps} BPS above the cross-section average."],
  play_head="Where to Play",
  play=["'{x_idx_top}' ranks highest ({s_x_idx_top}/5) and should be prioritised, while '{x_idx_low}' ranks lowest ({s_x_idx_low}/5) and should be deprioritised."],
)

# =============================================================================================
# C. COUNTRY DEEP-DIVE (geo = focus country, parent_geo = its region)
# =============================================================================================

S("C01", "Country", "Country Overview (Volume / Value)", "445-446", "country",
  title="{geo} {market} Market {metric_cap} Overview ({m_unit})",
  subtitle="{geo} market {metric}: size, growth, forecast build-up and gap to the {parent_geo} benchmark",
  kpi_labels=["{est_year} MARKET SIZE ({m_unit}): estimate year", "{fc_end} MARKET SIZE ({m_unit}): forecast horizon",
              "H-CAGR ({hist_period})", "F-CAGR ({fc_period})"],
  chart1_title="{geo} {metric_cap} ({m_unit}), {full_period}",
  chart2_title="CAGR Comparison: {geo} vs {parent_geo} and {geo_top}",
  benchmark_tag="vs {parent_geo} benchmark: {p_bench_arrow}",
  chart3_title="Forecast Build-up, {fc_period} ({m_unit})",
  chart4_title="{geo} Market Share, {est_year}",
  chart4_categories=["Share of {parent_geo}", "Share of {geo_top}"],
  readout=[
      "The {geo} {market_lc} market expands at {v_m_f_cagr} CAGR through {fc_end}, {p_vs_parent}.",
      "{geo_The} holds {v_share_top_est} of the {geo_top_ref} market and {v_share_parent_est} of the {parent_geo} market in {est_year}, and by {fc_end} these shares move to {v_share_top_end} and {v_share_parent_end}, respectively.",
      "The forecast build-up adds {v_m_abs} {m_unit} over {fc_period}, taking the {geo} market to {v_m_mult}x its {est_year} size.",
  ],
)

S("C02", "Country", "Country Key Segments Overview", "447-449", "country",
  title="{geo} {market} Market: Key Segments by Value, {est_year}{page_tag}",
  subtitle="Leading sub-segment within every segmentation of the {geo} market, by value",
  chart_title=R("segments_on_page", "By {item} ({val_unit})"),
  readout="Across the segments shown, the leading sub-segment accounts for between {v_lead_share_min} and {v_lead_share_max} of its segment's {est_year} value.",
  rule_note="Chart category count must equal the number of sub-segments; never pad with blank labels.",
)

S("C03", "Country", "Country Segmental Analysis", "452, 458 ...", "country",
  title="{geo} {market} Market Segmental Analysis: {segment}",
  subtitle="Growth contribution, share variance and absolute $ opportunity by {segment_lc}, {fc_period}",
  chart1_title="Growth Contribution by {segment}, {est_year} to {fc_end} ({val_unit})",
  table1_title="Value Share and Variance by {segment}",
  chart2_title="Absolute $ Opportunity by {segment} ({val_unit}), {fc_period}",
  table2_title="Volume Share and Variance by {segment}",
  table_headers=["{segment}", "Share {est_year}", "Share {fc_end}", "Variance (BPS)", "F-CAGR"],
  takeaway="{p_c03_takeaway}",
)

S("C04", "Country", "Country Volume and Value Tables", "453-454 ...", "country",
  title="{geo} {market} Market by {segment}: Volume and Value Tables{page_tag}",
  subtitle="Year-by-year volume and value of the {geo} market by {segment_lc}, {full_period}",
  table1_title="{geo} Volume by {segment} ({vol_unit})",
  table2_title="{geo} Value by {segment} ({val_unit})",
  takeaway="{p_c04_takeaway}",
)

S("C05", "Country", "Country YoY Analysis", "455-456 ...", "country",
  title="{geo} {market} Market YoY Analysis: {segment}{page_tag}",
  subtitle="Year-on-year volume and value growth by {segment_lc}, highlighting acceleration and slowdown",
  table_title="{geo} {metric_cap} YoY by {segment} (%): ▲ outpacing the market | ▶ below market pace | ▼ decline",
  insight="'{sub_yoy_peak}' records the highest single-year growth of {v_sub_yoy_peak} in {sub_yoy_peak_year}, {p_yoy_decline}.",
)

# =============================================================================================
# P. PRICING
# =============================================================================================

S("P01", "Pricing", "Product Pricing and Price Index Trends", "536", "global",
  title="{market} Market: Product Pricing and {child_label_adj_cap} Price Index Trends",
  subtitle="Price range by {seg1_lc} and indexed {child_label_adj} price movement, highlighting dispersion and structural shifts",
  banner="Average US$ price by {seg1_lc} and indexed {child_label_adj} price movement, adjusted to the base year to remove inflationary noise.",
  chart1_title="Pricing by {seg1}",
  chart2_title="Price Variation by {child_label_cap}",
  takeaways_head="Key Takeaways",
  takeaways=[
      "The upper chart shows the {geo_ref} average price of each {seg1_lc} category in US$, harmonised across {child_label_pl}, and each candlestick spans the lowest to the highest observed price in the base year.",
      "Wider bodies indicate a larger gap between opening and closing price levels; '{price_vol_high}' shows the most movement ({v_price_vol_high}) and '{price_vol_low}' the least ({v_price_vol_low}), a range buyers should monitor when planning procurement.",
      "The lower chart indexes {child_label_adj} price movement to the base year, applied consistently across the horizon, to remove year-on-year inflationary volatility and isolate structural price shifts.",
      "{child_label_adj_cap} price indices diverge by up to {v_price_div} between {est_year} and {fc_end}, and {p_price_conv}.",
  ],
)

S("P02", "Pricing", "Pricing Assessment by Segment", "537", "global",
  title="{geo} {market} Market Pricing Assessment by {seg1}",
  subtitle="Average price ({price_unit}) by {seg1_lc}, {hist_start} to {fc_end}, with the {est_year} price index and growth outlook",
  chart1_title="Price Trend by {seg1}, {hist_start} to {fc_end} ({price_unit})",
  chart1_note="Forecast period: {fc_period}",
  chart2_title="{est_year} Price Index vs {geo} Average (= 100)",
  table_headers=["{seg1}", "{hist_start}", "{est_year}", "{fc_end}", "CAGR {fc_period}", "Index"],
  card1_head="Price Hierarchy",
  card1="'{price_seg_top}' commands the highest {est_year} price point at {v_price_seg_top_est} {price_unit}, a {v_price_top_prem} premium over '{price_seg_low}'.",
  card1b="'{price_seg_top}' sits at the top of the {est_year} price ladder at an index of {v_price_top_idx}, while '{price_seg_low}' trades at {v_price_low_idx}, which keeps the full {seg1_lc} spread within {v_price_spread_idx} index points.",
  card2_head="Price Momentum",
  card2="'{price_seg_fast}' posts the fastest price growth at {v_price_seg_fast_cagr} CAGR over {fc_period}, compared with {v_price_seg_slow_cagr} for '{price_seg_slow}'.",
  card3_head="Outlook",
  card3="The {geo} average price {p_price_dir_verb} from {v_price_avg_est} {price_unit} in {est_year} to {v_price_avg_end} {price_unit} in {fc_end}, and the gap between the highest- and lowest-priced {seg1_lc} {p_price_spread_dir} from {v_price_gap_start} in {hist_start} to {v_price_gap_end} in {fc_end}.",
  footnote="Source: {source}. Average price = market value (US$) divided by market volume ({vol_unit}); {hist_period} historical and {fc_period} forecast.",
)

S("P03", "Pricing", "Price Matrix (Segment A x Segment B)", "538", "global",
  title="{geo} {market} Market Price Matrix: {seg_a} × {seg_b}",
  subtitle="Average price ({price_unit}) for every {seg_a_lc} and {seg_b_lc} combination, {est_year}",
  matrix_title="{est_year} Price Matrix ({price_unit}), Shaded by Index vs {geo} Average",
  legend=["Index below 97.5", "97.5 to 99", "Parity: 99 to 101", "101 to 102.5", "Index above 102.5"],
  chart_title="Price Range Across All {seg_b} Sub-Types, {est_year}",
  chart_note="Each bar spans the lowest to the highest price ({price_unit}) at the finest data level.",
  card1_head="Premium Pockets",
  card1="'{pm_top_a}' in '{pm_top_b}' is the highest-priced combination in {est_year} at {v_pm_top_price} {price_unit}, an index of {v_pm_top_idx} against the {geo} average, which signals room for value-added positioning.",
  card2_head="Value Pockets",
  card2="'{pm_low_a}' in '{pm_low_b}' is the lowest-priced combination at an index of {v_pm_low_idx}, {v_pm_low_gap} below the {geo} average of {v_price_avg_est} {price_unit}.",
  card3_head="Within-Type Spread",
  card3="'{pm_spread_a}' shows the widest internal price spread at {v_pm_spread}, ranging from {v_pm_spread_min} to {v_pm_spread_max} {price_unit} across sub-types.",
  footnote="Source: {source}. Cell price = combined value (US$) divided by combined volume ({vol_unit}); shading shows the index against the {geo} average of {v_price_avg_est} {price_unit}.",
)

S("P04", "Pricing", "Price Benchmarking by Child Geo", "539", "global",
  title="{geo} {market} Market Pricing Assessment: {child_label_adj_cap} Benchmarking",
  subtitle="Average price and price range across {seg1_lc} categories by {child_label}, {est_year} ({price_unit})",
  chart_title="{est_year} Price Range Across {seg1} Categories by {child_label_cap} ({price_unit}), Ranked by Average",
  table_title="Premium and Discount Markets",
  table_headers=["{child_label_cap}", "Avg. {est_year}", "Index", "CAGR {fc_period}"],
  lead="{child_price_high} is the most premium {child_label} at {v_child_price_high} {price_unit} in {est_year}, {v_child_price_high_gap} above the {geo} average, while {child_price_low} is the lowest priced.",
  chart_note="Each bar spans the lowest to the highest {seg1_lc} price in the {child_label}, read on the price index axis ({geo} average = 100); teal marks above-average and blue below-average {child_label_pl}.",
  card1_head="{child_label_adj_cap} Spread",
  card1="Average prices across {child_label_pl} differ by {v_child_price_spread} in {est_year}, from {v_child_price_low} {price_unit} in {child_price_low} to {v_child_price_high} {price_unit} in {child_price_high}.",
  card2_head="Growth Leader",
  card2="{child_price_fast} records the fastest average price growth at {v_child_price_fast_cagr} CAGR over {fc_period}, against {v_price_avg_cagr} for the {geo_ref} market.",
  card3_head="Mix Exposure",
  card3="{child_mix_wide} shows the widest gap between its lowest- and highest-priced {seg1_lc} at {v_child_mix_gap}, so shifts in product mix move its average price the most.",
  footnote="Source: {source}. Range = lowest to highest {seg1_lc} price in each {child_label}; {child_label_pl} ranked by average price.",
)

S("P05", "Pricing", "Price Positioning and Key Findings", "540", "global",
  title="{market} Market: Price Positioning and Key Findings",
  subtitle="Average price and price range of each {child_label} on a common scale, with {publisher_short}'s key findings",
  banner="Average selling price and price range by {child_label} on a shared price scale ({price_unit}).",
  legend=["Average Price", "Price Range"],
  findings_head="Key Findings",
  findings=[
      "Average prices {p_price_band}, ranging from about {v_child_price_low} {price_unit} in {child_price_low} to about {v_child_price_high} {price_unit} in {child_price_high} in {est_year}.",
      "{child_price_high} commands a premium, which signals stronger demand for higher-value {seg1_lc} categories and less price-led supply, while {child_price_low} remains the most accessible, value-led {child_label}.",
      "The marker shows each {child_label}'s average selling price and the band its full low-to-high range, both derived as value divided by volume.",
      "Narrow bands point to commoditised, price-led competition, while wider bands flag headroom to trade buyers up the mix.",
  ],
  note="Note: The information above is indicative and derived from the FMI market model.",
)

S("P06", "Pricing", "Key Pricing Drivers", "541", "global",
  title="Key Pricing Drivers Influencing the {market} Market",
  subtitle="Input costs, market conditions and competitive forces shaping pricing strategy and profitability",
  banner="Input costs, market conditions and competitive forces that shape pricing strategy and profitability across the value chain.",
  chart_title="Price-Impacting Factors (score 0 to 10)",
  factor_labels=["Input Costs", "Labour Costs", "Process Innovation", "Logistics Costs", "Demand Conditions",
                 "Competitive Intensity", "Compliance Costs", "Currency Movements", "Brand Equity", "{lx_factor10}"],
  takeaways_head="Key Takeaways",
  takeaways=[
      "Pricing in the {market} Market reflects a mix of input costs, operating efficiency and market dynamics, with {drv1_lc} carrying the heaviest weight on selling prices today.",
      "{drv1} and {drv2_lc} are the dominant pressures, and as they fluctuate they feed directly into prices and squeeze margins along the value chain.",
      "{drv3} and {drv4_lc} are the next most influential forces, setting how far players can offset cost inflation and how aggressively they can price.",
  ],
  top_head="Highest-Impact Drivers",
  top=R("drivers_top5", "▪ {item}: {s_item}/10, {p_item_driver_role}."),
)

S("P07", "Pricing", "Factors Impacting Pricing (table)", "542", "global",
  title="{market} Market: Factors Impacting Pricing",
  subtitle="Factors shaping pricing decisions: category, price impact, direction of pressure and rationale",
  banner="Cost, demand, efficiency and regulatory forces that shape achievable selling prices, grouped by category with each factor's direction, impact and score.",
  table_headers=["Factor", "Category", "Price Impact", "Direction", "Score (0 to 10)", "How It Shapes Price"],
  rationale={
      "Input Costs": "Movements in {lx_input_costs} set the cost floor and pass through to selling prices with a lag.",
      "Labour Costs": "Wage inflation and plant utilisation determine conversion cost per unit.",
      "Logistics Costs": "Freight, warehousing and distribution costs add to landed price, especially in import-reliant markets.",
      "Demand Conditions": "Strong demand allows suppliers to hold or raise prices, while weak demand forces promotional pricing.",
      "Competitive Intensity": "{lx_competitive_pressure_cap} limit how far prices can rise.",
      "Brand Equity": "Brand equity, backed by {lx_validation}, sustains a price premium over unbranded alternatives.",
      "Process Innovation": "Process automation and yield gains lower unit costs and allow more competitive pricing.",
      "{lx_factor10}": "{lx_factor10_why}",
      "Compliance Costs": "Compliance with {lx_regulator} adds testing, documentation and labelling cost.",
      "Currency Movements": "Exchange-rate swings and tariffs change the landed cost of imported inputs and finished goods.",
  },
  implications_head="What the Factors Imply",
  implications=[
      "{drv1} and {drv2_lc} carry the heaviest weight on selling prices, at {s_drv1} and {s_drv2} out of 10.",
      "{p_direction_split}",
      "Cost-side factors set the price floor, while demand, brand and competitive forces determine how much of that cost can be passed through.",
      "{drv_last} exerts the least pressure and is a second-order effect to monitor rather than actively manage.",
  ],
)

S("P08", "Pricing", "Regional Price Assessment (value chain)", "543", "global",
  title="{market} Market: {child_label_adj_cap} Price Assessment",
  subtitle="Price build-up across the value chain, from manufacturer to retail, and average price by {child_label}, {fc_period}",
  banner="Price build-up across the value chain ({price_unit}) and price trajectory to {fc_end} by {child_label}.",
  table1_title="{child_label_adj_cap} Price by Value-Chain Level ({price_unit})",
  table1_headers=["{child_label_cap}", "Manufacturer", "Distribution", "Retail", "Total Margin (%)"],
  table2_title="Price by {child_label_cap} ({price_unit})",
  covers_head="What This Section Covers",
  covers=[
      "Manufacturer-level pricing benchmarks the ex-works producer price in each {child_label}.",
      "Distribution and channel pricing layers logistics and channel margin onto the producer price.",
      "Retail or end-market pricing shows the average price paid by the end customer in each {child_label}.",
      "The value-chain margin build-up shows the cumulative margin from manufacturer to retail, with all prices derived as value divided by volume.",
  ],
  takeaway="{child_margin_high} carries the highest total value-chain margin at {v_child_margin_high}, reflecting longer distribution chains and higher end-market mark-ups.",
)

S("P09", "Pricing", "Pricing Benchmarking (child geo x segment)", "544", "global",
  title="{market} Market: {geo} Pricing Benchmarking",
  subtitle="Price range by {child_label} and {seg1_lc}, benchmarked against each {child_label}'s weighted average price",
  banner="Manufacturer price benchmarks by {seg1_lc} and {child_label} ({price_unit}).",
  table_title="Price Benchmarking by {child_label_cap} and {seg1} ({price_unit})",
  insights_head="Benchmarking Insights",
  insights=[
      "On a weighted-average basis, {child_price_high} is the highest-priced {child_label} at about {v_child_price_high} {price_unit}, and {child_price_low} is the most competitive at about {v_child_price_low} {price_unit}.",
      "Price ranges are benchmarked for each {seg1_lc} category in {est_year}, derived as value divided by volume with a trading band of plus or minus {v_trade_band}.",
      "Cross-{child_label} price dispersion of about {v_price_disp}x reflects differences in product mix, logistics and channel structure.",
      "These benchmarks support price-corridor setting and margin-defence decisions across {child_label_pl}.",
  ],
)

S("P10", "Pricing", "Price Forecast and Dynamics", "545", "global",
  title="{market} Market: Price Forecast and Dynamics",
  subtitle="Price forecast by {seg1_lc}, {fc_period}, with the factors shaping the trajectory",
  banner="Forward price trajectory to {fc_end} by {seg1_lc} and the structural factors shaping it ({price_unit}).",
  table1_title="Price Forecast to {fc_end} by {seg1} ({price_unit})",
  table2_title="Factors Impacting Pricing",
  table2_headers=["Pricing Factor", "Direction", "Impact (H/M/L)"],
  outlook_head="Pricing Dynamics and Outlook",
  outlook=[
      "The average {seg1_lc} price is projected to {p_price_dir_inf} from about {v_price_avg_est} to about {v_price_avg_end} {price_unit} by {fc_end}, a {v_price_avg_cagr} CAGR.",
      "'{price_seg_fast}' shows the strongest forecast price appreciation, at {v_price_seg_fast_chg} over the horizon.",
      "Upward pressure is led by {drv1_lc} and {drv2_lc}, while {p_down_list} work in the opposite direction.",
      "{p_price_outlook}",
  ],
)

S("P11", "Pricing", "Competitor Pricing Strategies", "546", "global",
  title="{market} Market: Competitor Pricing Strategies",
  subtitle="How leading players position on price: tier, position versus peers and strategy focus",
  banner="How leading players position on price and the levers they use to sustain it.",
  table_headers=["Company", "Tier", "Price Position", "ASP vs Peers", "Strategy Focus"],
  strategy_focus_bank={
      "Premium": ["Premium, brand-led pricing", "Value-based premium pricing", "Innovation-led price leadership"],
      "Parity": ["Parity pricing with mix upsell", "Penetration and share-led pricing"],
      "Value": ["Competitive, promotion-led pricing", "Cost-leadership pricing"],
  },
  note="Note: Competitor figures are indicative, derived from market modelling and expert validation rather than company disclosures.",
  themes_head="Competitive Pricing Themes",
  themes=[
      "Tier 1 players anchor premium pricing and defend average selling price through brand strength and product development, backed by {lx_validation}, rather than discounting.",
      "Mid-tier players price at or just below parity and use launches and expansion to win share in fast-growing segments.",
      "Emerging players compete on value pricing, trading margin for access and volume in price-sensitive {child_label_pl}.",
      "Across tiers, R&D investment and selective M&A are the primary levers for moving up the price ladder over time.",
  ],
)

# =============================================================================================
# V. SUPPLY CHAIN
# =============================================================================================

S("V01", "Supply chain", "Value Chain and Stakeholder Mapping", "549", "global",
  title="{market} Market: Value Chain and Stakeholder Mapping",
  subtitle="Value chain stages, the stakeholders active at each stage and the value added along the chain",
  intro="End-to-end value chain with the key stakeholders at each stage and their indicative share of value added.",
  stage_labels=["{lx_vc1}", "{lx_vc2}", "{lx_vc3}", "{lx_vc4}", "{lx_vc5}"],
  stage_descriptions=["{lx_vc1_desc}", "{lx_vc2_desc}", "{lx_vc3_desc}", "{lx_vc4_desc}", "{lx_vc5_desc}"],
  stage_year="Base Year: {est_year}",
  readout_head="Value-Added Read-out",
  readout=[
      "The {market} Market value chain runs through five stages: {lx_vc1_lc}; {lx_vc2_lc}; {lx_vc3_lc}; {lx_vc4_lc}; and {lx_vc5_lc}.",
      "The {vc_top_stage_lc} stage captures the largest share of value added ({v_vc_top_share}), reflecting {lx_vc_top_reason}.",
      "Forward and backward integration is common, as some suppliers move into processing and some large buyers produce their own inputs.",
      "Percentages show each stage's indicative share of value added along the chain.",
  ],
)

S("V02", "Supply chain", "Operating Margin Analysis", "550", "global",
  title="{market} Market: Industry-Wide Operating Margin Analysis",
  subtitle="Operating margin by {child_label} and value-chain node, highlighting where value concentrates",
  table_title="{market} Market Operating Margin by {child_label_cap} and Node",
  table_headers=["{child_label_cap}", "Node of Supply Chain", "Operating Margin", "Comparison to Parent Market", "Notes"],
  node_notes={
      "{lx_vc1}": "Commodity-linked with limited pricing power",
      "{lx_vc2}": "Highest margins from value addition and scale",
      "{lx_vc3}": "Volume-led with thin margins",
      "{lx_vc4}": "Channel-dependent point-of-sale margin",
  },
  read_head="How to Read This",
  read=[
      "Node of supply chain shows each major stage, from {lx_vc1_lc} through to {lx_vc4_lc}.",
      "Operating margin shows the indicative margin earned at each node in each {child_label}.",
      "Comparison to parent market shows the margin against the wider {lx_parent_market} value chain.",
      "Margins are shown by both {child_label} and node to reveal where value concentrates along the chain.",
  ],
)

S("V03", "Supply chain", "Buyer Assessment", "551", "global",
  title="{market} Market: Buyer Assessment",
  subtitle="Key buyers and end users by procurement volume, average price and revenue, highlighting purchasing power",
  table_title="{market} Market Buyer and End-User Assessment",
  table_headers=["Buyer", "Country", "{child_label_cap}", "Buyer Category", "Procurement Vol. ({vol_unit})",
                 "Avg Price ({price_unit})", "Spend ({val_unit})", "Key Selection Factor"],
  buyer_category_bank=["{lx_buyer1}", "{lx_buyer2}", "{lx_buyer3}", "{lx_buyer4}"],
  selection_factor_bank=["{lx_buyer_factor1}", "{lx_buyer_factor2}", "{lx_buyer_factor3}", "{lx_buyer_factor4}"],
  notes_head="Notes",
  notes=[
      "This is a representative subset; the full assessment lists 20 to 30 active buyers with brands, specifications, procurement strategies and SWOT.",
      "Value-chain economics are indicative, derived from market modelling and expert validation rather than company disclosures.",
  ],
)

# =============================================================================================
# M. MARKET DYNAMICS
# =============================================================================================

S("M01", "Dynamics", "Dynamics by Child Geo", "554", "global",
  title="{geo} {market} Market Dynamics by {child_label_cap}",
  subtitle="Growth drivers, restraints and priority opportunities in each {child_label}, ranked by {fc_end} value",
  table_headers=["{child_label_cap}", "Growth Drivers", "Restraints", "Priority Opportunities"],
  driver_cell=R("children", "{p_item_drivers}"),
  restraint_cell=R("children", "{p_item_restraints}"),
  opportunity_cell=R("children", "Scale '{item_seg1_fast}' ({v_item_seg1_fast_cagr} CAGR) to capture {v_item_opp} {val_unit} of headroom to {fc_end}."),
  takeaway_head="KEY TAKEAWAY",
  takeaway="{top3_children} anchor scale with a combined {v_top3_share} of {geo} value in {fc_end}, while {child_fast} offers the steepest upside at {v_child_fast_cagr} CAGR against a market pace of {v_f_cagr}. {p_above_pace} Entry strategy should therefore differ by {child_label}: defend scale in the leaders, pursue mix and premiumisation in the mid-tier and lead with availability in the fastest movers.",
)

S("M02", "Dynamics", "Driving Factors: Supply Side", "555", "global",
  title="{geo} {market} Market: Driving Factors, Supply Side",
  subtitle="Economic and supply-side factors scored by their impact on the {geo_ref} market",
  factor_labels=["Production Scale-up", "Input Cost Pressure", "Capacity Addition", "Supply Concentration",
                 "Geographic Spread", "Portfolio Breadth", "Route-to-Market Depth", "Value Realisation"],
  readout_head="Economic and Supply-Side Factors",
  readout=[
      "Supply-side conditions are led by {sf1} ({s_sf1}/10), {sf2} ({s_sf2}/10) and {sf3} ({s_sf3}/10) on a 0 to 10 scale scored from the {geo_ref} dataset.",
      "Volume expands at {v_vol_f_cagr} CAGR over {fc_period} while average realisation moves {v_price_avg_cagr} a year, and the balance between the two sets the cost pass-through picture.",
      "Supply is {p_supply_conc}, as {child_lead} holds {v_child_lead_share_end} of {fc_end} value across {n_children} {child_label_pl}.",
      "Capacity added over {fc_period} is worth {v_abs_opp} {val_unit}, a {v_mult}x expansion on the {est_year} base.",
      "Portfolio breadth spans {n_seg1_sub} {seg1_lc} categories and {n_seg2_sub} {seg2_lc} routes, which gives producers room to adjust mix and channel strategy.",
  ],
  note="Note: Factor scores and impact ratings are indicative and generated from the market dataset.",
)

S("M03", "Dynamics", "Driving Factors: Demand Side", "556", "global",
  title="{geo} {market} Market: Driving Factors, Demand Side",
  subtitle="Demand-side factors scored by their impact on the {geo_ref} market",
  factor_labels=["Demand Growth", "Premiumisation", "Application Breadth", "Emerging-Market Pull",
                 "Volume Momentum", "Category Maturity", "Adoption Depth", "Customer Reach"],
  readout_head="Demand-Side Factors",
  readout=[
      "Demand-side conditions are led by {df1} ({s_df1}/10), {df2} ({s_df2}/10) and {df3} ({s_df3}/10) on a 0 to 10 scale.",
      "Value grows at {v_f_cagr} against volume at {v_vol_f_cagr}, {p_price_mix_short}.",
      "{p_seg1_lead_fast}",
      "Pull is strongest in {child_fast} at {v_child_fast_cagr} CAGR, {v_child_fast_gap_bps} BPS above the market. {p_above_pace}",
      "{p_seg2_reach}",
  ],
  note="Note: Factor scores and impact ratings are indicative and generated from the market dataset.",
)

S("M04", "Dynamics", "Market Restraints", "557", "global",
  title="{geo} {market} Market Restraints",
  subtitle="Key restraints on the {geo_ref} market and the likelihood of their impact across {full_period}",
  box_head="Market Restraints",
  timeline_label="Likelihood of impact over time",
  timeline=["{tl1}", "{tl2}", "{tl3}", "{tl4}"],
  r1_head="'{seg1_slow}' Lags the Category and Dilutes Blended Growth",
  r1_body="'{seg1_slow}' grows at only {v_seg1_slow_cagr} CAGR against the market's {v_f_cagr}, a gap of {v_seg1_slow_gap_bps} BPS, and cedes {v_seg1_slow_bps} BPS of {seg1_lc} share between {est_year} and {fc_end}. Because this category still carries meaningful volume, the drag is absorbed by the whole market and blended growth sits below what the leading categories achieve on their own. Repositioning or rationalising the lagging end of the portfolio is therefore a precondition for delivering the headline forecast.",
  r2_head="Uneven {child_label_adj_cap} Performance Caps the Blended Growth Rate",
  r2_body="{child_slow} grows at just {v_child_slow_cagr} CAGR, {v_child_slow_gap_bps} BPS below the market, and {n_below_pace} of {n_children} {child_label_pl} grow below the market pace. Demand is therefore not uniformly available, and suppliers with a broad footprint fund their presence in slower {child_label_pl} from returns earned in faster ones. With {child_lead} accounting for {v_child_lead_share_end} of value, a shortfall in a single leading {child_label} is difficult to offset elsewhere in the forecast window.",
  r3_head="{p_r3_head}",
  r3_body="{p_r3_body}",
  likelihood_rule="Likelihood per period is set by rules.likelihood(score) from the model score s_r*_t1..t4: below 4 = Low, 4 to 6.9 = Medium, 7 and above = High.",
)

S("M05", "Dynamics", "Market Opportunities", "558", "global",
  title="{geo} {market} Market Opportunities",
  subtitle="Key opportunities in the {geo_ref} market and the likelihood of their impact across {full_period}",
  box_head="Market Opportunities",
  timeline_label="Likelihood of impact over time",
  timeline=["{tl1}", "{tl2}", "{tl3}", "{tl4}"],
  o1_head="{cty_abs_top} Carries the Largest Absolute Opportunity",
  o1_body="{cty_abs_top_The} adds {v_cty_abs_top_opp} {val_unit} over {fc_period}, the largest absolute increment of any country in the dataset. Capturing it is a question of scale execution rather than market creation, because demand already exists and grows at {v_cty_abs_top_cagr} CAGR. Prioritising supply, channel coverage and service levels there converts the largest share of the {v_abs_opp} {val_unit} the market adds over the forecast period.",
  o2_head="{cty_white} Offers Under-Penetrated, Above-Pace Headroom",
  o2_body="{cty_white_The} compounds at {v_cty_white_cagr} CAGR from a below-median base and is one of {n_white} countries that combine above-median growth with below-median size. This profile is where share is cheapest to win, because competitive intensity is lower than in the largest markets while growth runs {v_cty_white_gap_bps} BPS ahead of the market. Early distribution and reference accounts established here compound through {fc_end}.",
  o3_head="'{seg2_fast}' Opens the Fastest Route to Market",
  o3_body="{p_o3_body}",
)

S("M06", "Dynamics", "Milestones and Forecast Checkpoints", "559", "global",
  title="{geo} {market} Market: Milestones and Forecast Checkpoints",
  subtitle="Historical inflection points and forecast checkpoints for the {geo_ref} market, in chronological order",
  table_headers=["Year", "Milestone"],
  row_peak="Strongest historical expansion: {geo} value rose {v_peak_hist_yoy} year on year to {v_peak_hist_val} {val_unit}, the steepest single-year gain of {hist_period}, with '{seg1_lead}' setting the pace.",
  row_base="Base year for the forecast: the {geo_ref} market is estimated at {v_val_est} {val_unit} after compounding at {v_h_cagr} over {hist_period}, and {child_lead} holds {v_child_lead_share_est} of value across {n_children} {child_label_pl}.",
  row_mid="Mid-horizon checkpoint: value reaches {v_val_mid} {val_unit}, {v_mult_mid}x the {est_year} base, and the {seg1_lc} mix has moved {v_seg1_gain_bps_mid} BPS toward '{seg1_gain}'.",
  row_x15="The market passes {v_val_1_5x} {val_unit}, one and a half times the {est_year} base, led by '{seg1_fast}' at {v_seg1_fast_cagr} CAGR and {child_fast} at {v_child_fast_cagr} CAGR.",
  row_end="Forecast outcome: the market reaches {v_val_end} {val_unit}, a {v_mult}x expansion on {est_year} that adds {v_abs_opp} {val_unit} at {v_f_cagr} CAGR, while volume compounds at {v_vol_f_cagr}.",
  row_real_dev="{c_dev_date}: {c_dev_company} {c_dev_text} (Source: {c_dev_source}).",
  rule_note="Sort rows by year. Include row_x15 only when 1.5x the base is reached by {fc_end}. When a sourced developments list exists, use row_real_dev and title the slide 'Recent Developments'.",
)

S("M07", "Dynamics", "Porter's Five Forces", "560", "global",
  title="Porter's Five Forces Analysis: {geo} {market} Market",
  subtitle="Supplier power, buyer power, competitive rivalry, substitutes and entry barriers shaping market profitability",
  intro="The five forces are scored from the {geo_ref} dataset, covering {n_seg1_sub} {seg1_lc} categories, {n_seg2_sub} {seg2_lc} routes and {n_children} {child_label_pl} growing at {v_f_cagr} CAGR.",
  suppliers_head="Bargaining Power of Suppliers",
  suppliers=[
      "Assessed {p_force_suppliers}: suppliers of {lx_input_costs} hold leverage where inputs are specialised or capacity is concentrated.",
      "Average realisation moves {v_price_avg_cagr} a year over {fc_period}, which shows how far input costs are being passed through.",
      "Buyers can rebalance across {n_seg1_sub} {seg1_lc} categories, which caps supplier leverage in standard grades.",
  ],
  buyers_head="Bargaining Power of Buyers",
  buyers=[
      "Assessed {p_force_buyers}: demand routes through {n_seg2_sub} {seg2_lc} routes led by '{seg2_lead}' at {v_seg2_lead_share}, which gives large buyers negotiating scale.",
      "{lx_competitive_pressure_cap} strengthen buyer leverage on price and terms.",
      "Switching is easiest where comparable {seg1_lc} categories compete for the same demand.",
  ],
  rivalry_head="Intensity of Competition",
  rivalry=[
      "Assessed {p_force_rivalry}: {n_named} named players compete, and the top three hold an estimated {v_cr3} of value.",
      "Growth is uneven, with '{seg1_fast}' at {v_seg1_fast_cagr} against '{seg1_slow}' at {v_seg1_slow_cagr}, which concentrates competition in the faster categories.",
      "{p_above_pace}",
  ],
  substitutes_head="Threat of Substitutes",
  substitutes=[
      "Assessed {p_force_substitutes}: {lx_substitutes} compete for the same demand.",
      "{p_seg1_share_shift}",
      "Rising realisation of {v_price_avg_cagr} a year widens the price gap that substitutes must close.",
  ],
  entrants_head="Threat of New Entrants",
  entrants=[
      "Assessed {p_force_entrants}: the market grows at {v_f_cagr} CAGR and adds {v_abs_opp} {val_unit} by {fc_end}, which attracts entry.",
      "{lx_entry_barriers_cap} remain the principal barriers to entry.",
      "{p_entry_spread}",
  ],
)

# =============================================================================================
# K. COMPETITIVE LANDSCAPE
# =============================================================================================

S("K01", "Competition", "Competitive Landscape at a Glance", "563", "global",
  title="{market} Market: Competitive Landscape at a Glance",
  subtitle="Market concentration, segment leadership and the moves that will shift share, {fc_period}, on one page",
  kpi_labels=["Top 3 share, {est_year}", "Top 5 share", "Top 3 {seg1_lc} categories", "Largest {seg2_lc}",
              "Named competitors", "Moves tracked", "Largest player", "CAGR {fc_period}"],
  f1_head="{p_k01_conc_head}",
  f1="The top three hold {v_cr3} of {est_year} value, while {n_leaders} companies lead the {n_tracked_subs} sub-segments tracked and no single player leads more than {n_max_leads}.",
  f2_head="The tier structure is {p_tier_word}",
  f2="Tier 1 ({n_t1} players) holds {v_t1_share}, while Tier 2 and Tier 3 together hold {v_t23_share}.",
  f3_head="Growth is shifting within the {seg1_lc} mix",
  f3="{p_seg1_lead_fast}",
  f4_head="Reach divides the field",
  f4="{n_global} players compete globally, {n_regional} regionally and {n_local} in a single market or niche.",
  f5_head="{p_ma_head}",
  f5="{p_ma_body}",
  f6_head="Success factors, in order",
  f6="The success factors rank as segment breadth, scale and share, geographic reach, innovation pipeline and strategic momentum, and the composite ranking is led by {rank1}, {rank2} and {rank3}.",
  method_head="HOW THIS SECTION IS BUILT",
  method="Pages 2 to 4 show where the market stands and who plays. Pages 5 to 7 rate every named company on twenty criteria covering who they are, what they sell and how they win, and a cell reads n/a where no evidence exists. Page 8 converts the pattern into five success factors and a composite ranking. Pages 9 and 10 list what each player is doing and how each internal function should respond.",
)

S("K02", "Competition", "Segment Leadership Map", "564", "global",
  title="{market} Market: Segment Leadership Map",
  subtitle="Share by player group and the named player that leads or is active in each sub-segment, {est_year}",
  chart_title="Who holds the market, {est_year}",
  groups=["Top 3", "Rest of top 5", "Emerging", "Others: long tail of regional and niche suppliers"],
  matrix_title="Who leads which sub-segment, {est_year} (full = leader, half = active)",
  reading_head="READING",
  reading="Read across a row to see a player's breadth and down a column to see how contested a sub-segment is. '{most_contested}' is the most contested sub-segment with {n_most_contested} active players, while '{most_open}' is the most open with {n_most_open}.",
)

S("K03", "Competition", "Value Chain Positioning of Leading Players", "565", "global",
  title="{market} Market: Value Chain Positioning of Leading Players",
  subtitle="How the market is organised by {seg1_lc} and {seg2_lc}, and who leads at each stage",
  kpi_labels=["Top 3 share, {est_year}", "Top 5 share", "{seg1_sub1}, share", "{seg1_sub2}, share"],
  panel_head=R("seg1_subs", "{rank}. {item}"),
  panel_labels=["of market", "CAGR {fc_period}", "LEADS", "ALSO ACTIVE", "WHAT TO WATCH"],
  panel_watch=R("seg1_subs", "{n_item_active} named players compete here, and '{item}' adds {v_item_opp} {val_unit} by {fc_end}."),
  seg2_head="{n_seg2_panel}. {seg2}",
  seg2_kicker="By {seg2_lc}: where demand is served.",
  seg2_watch="The top two {seg2_lc} routes account for {v_seg2_top2_share} of the market, so vendors typically build their offer around one route at a time.",
  enablers_head="ENABLERS",
  enablers="{p_enablers}",
  constraints_head="CONSTRAINTS",
  constraints="{p_constraints}",
)

S("K04", "Competition", "Player Landscape by Tier and Reach", "566", "global",
  title="{market} Market: Player Landscape by Tier and Reach",
  subtitle="Named players by tier and geographic reach, with how each group competes",
  tier_labels=["Tier 1 ({n_t1})", "Tier 2 ({n_t2})", "Tier 3 ({n_t3})"],
  reach_labels=["Global ({n_global}): sells in most major regions", "Regional ({n_regional}): strong in one or two regions",
                "Local or niche ({n_local}): single market or specialist use"],
  summary="TIER MIX  Tier 1: {n_t1} | Tier 2: {n_t2} | Tier 3: {n_t3}     BY REACH  Global: {n_global} | Regional: {n_regional} | Local or niche: {n_local}",
  how_earns=["Direct and key-account sales with premium or bundled positioning.",
             "Distributor-led sales with mid-market or value-added positioning.",
             "Regional partners with value or entry-level positioning."],
  dominant_strategy=R("tiers", "{item_strategy}, with {v_item_share} combined share."),
  strategy_bank=["Acquisition-led scale-up", "Launch-led growth", "Geographic expansion", "Capacity expansion", "Partnership-led growth"],
  rule_note="Tier must come from company revenue in the input sheet (Tier 1 = top revenue band), never from list order.",
)

S("K05", "Competition", "Corporate Scorecard: Who They Are", "567", "global",
  title="{market} Market Corporate Scorecard: Who They Are",
  subtitle="Eight corporate criteria for every named competitor: ownership, strategy, breadth, presence, scale and performance",
  level="LEVEL 1, CORPORATE: WHO THEY ARE",
  criteria=[("Parent and ownership", "From the company roster"), ("Purpose (strategy)", "Dominant active move"),
            ("Portfolio breadth", "Sub-segments: 6 or more = 3, 4 to 5 = 2, 2 to 3 = 1"),
            ("Presence", "Global = 3, regional = 2, local = 1"), ("Power (scale)", "Tier 1 = 3, Tier 2 = 2, Tier 3 = 1"),
            ("Performance", "Share band"), ("People", "Talent; not assessed"), ("Partnerships", "Named alliances")],
  legend=["Strong, direct evidence", "Moderate or partial", "Weak or indirect", "None found"],
  reading_head="READING",
  reading="{n_global} of {n_named} named players compete globally, and scale and share broadly follow the tier structure.",
  strategy_bank=["Buy to scale", "Launch-led growth", "Expand footprint", "Partner to grow"],
  rule_note="Purpose must come from the company's sourced activity list; if none is recorded, show n/a rather than a default label.",
)

S("K06", "Competition", "Offering Scorecard: What They Sell", "568", "global",
  title="{market} Market Offering Scorecard: What They Sell",
  subtitle="Six offering criteria: product, technology, price and model, verticals, proof and IP",
  level="LEVEL 2, OFFERING: WHAT THEY SELL",
  criteria=[("Product", "Of the top three categories offered"), ("Platform (technology)", "R&D and development: both = 3, one = 2"),
            ("Price and model", "Positioning by tier"), ("Positioning (verticals)", "{seg2} routes served"),
            ("Proof", "Named customer references"), ("Patents and IP", "Where the study covers patents")],
  price_model_bank=["Premium or bundled", "Mid-market or value-added", "Value or entry-level"],
  reading_head="READING",
  reading="Product breadth and technology strength do not always coincide, and the {seg2_lc} routes each player serves show where it competes head to head. The proof criterion requires customer research to complete.",
)

S("K07", "Competition", "Execution Scorecard: How They Win", "569", "global",
  title="{market} Market Execution Scorecard: How They Win",
  subtitle="Six execution criteria: channel, promotion, momentum, pipeline, policy exposure and prospects",
  level="LEVEL 3, EXECUTION: HOW THEY WIN",
  criteria=[("Place (channel)", "Route to market by tier"), ("Promotion", "Go-to-market; not assessed"),
            ("Pace (momentum)", "Active moves: 5 or more = 3, 3 to 4 = 2, else 1"), ("Pipeline", "Launch or development in motion"),
            ("Policy exposure", "Regulation in its favour or against"), ("Prospects", "Gaining = 3, holding = 2, at risk = 1")],
  channel_bank=["Direct and key accounts", "Distributors", "Regional partners"],
  reading_head="READING",
  reading="Momentum sits with the players running the most strategic moves, and prospects combine breadth, reach, scale and pace. The promotion and policy-exposure criteria require market and regulatory research to complete.",
)

S("K08", "Competition", "Success Factors and Composite Ranking", "570", "global",
  title="{market} Market: Success Factors and Composite Ranking",
  subtitle="Five weighted success factors and the composite ranking they produce",
  factors=[("1. Segment breadth", "Leads or is active in several sub-segments"),
           ("2. Scale and share", "Tier position and share of market value"),
           ("3. Geographic reach", "Global, regional or single-market footprint"),
           ("4. Innovation pipeline", "R&D and product development in motion"),
           ("5. Strategic momentum", "Number of active strategic moves")],
  composite="Composite: weighted, 0 to 3",
  groups_head="STRATEGIC GROUPS",
  group_line=R("tiers", "{item}: {item_members}. This group wins on {item_best} and is weakest on {item_worst}."),
)

S("K09", "Competition", "Strategic Activity Tracker", "571", "global",
  title="{market} Market: Strategic Activity Tracker",
  subtitle="Each player's active strategic moves: what is happening, type, source, signal and who should act",
  table_headers=["COMPANY", "DEVELOPMENT (WHAT IS HAPPENING)", "DATE", "TYPE", "SOURCE DEPT", "EVIDENCE", "SIGNAL (WHAT IT INDICATES)", "WHO SHOULD ACT"],
  row=R("developments", "{c_dev_company} | {c_dev_text} | {c_dev_date} | {c_dev_type} | {p_dev_dept} | {c_dev_source} | {p_dev_signal} | {p_dev_action}"),
  empty_row="{item} | No material {market_lc}-related development disclosed in the tracking window | n/a | n/a | n/a | {c_dev_checked_source} | n/a | n/a",
  rule_note="Rows come only from a sourced developments sheet (date and source URL required). Never generate a development for a real company from a template.",
)

S("K10", "Competition", "Strategic Activity by Function", "572", "global",
  title="{market} Market: Strategic Activity by Function",
  subtitle="Where competitor activity concentrates by move type and department, and what it means for each function",
  chart_title="Moves by type (count)",
  type_legend="MA = M&A | PR = Product launch | EX = Expansion | RD = R&D | RG = Regulatory or certification | CP = Partnership",
  insight="{top_type_name} is the most frequent move type ({n_top_type} of {n_moves} tracked moves), so {top_type_function} should treat it as the primary competitive signal.",
)

S("K11", "Competition", "Industry Structure Landscape (1/2)", "573", "global",
  title="Industry Structure Landscape: {geo} {market} Market, {est_year} (1/2)",
  subtitle="Tier ladder, market concentration and the profile of each tier",
  ladder_title="Tier Ladder: Share of Market Value, {est_year}",
  tiers=["Tier 1: top or global players", "Tier 2: mid-sized players", "Tier 3: emerging or local players",
         "Others: long tail of regional and niche suppliers"],
  concentration_head="Concentration Read-out",
  bands=["Fragmented: HHI below 1,500", "Moderate: HHI 1,500 to 2,500", "Concentrated: HHI above 2,500"],
  concentration="With an estimated HHI of {v_hhi}, the market is {p_hhi_band}, and the top four players hold {v_cr4} of {est_year} value.",
  matrix_title="Tier Profile Matrix",
  matrix_subtitle="Size, growth and competitive characteristics of each tier",
  matrix_rows=["Players", "Value {est_year} ({val_unit})", "Share of value {est_year}", "Share of value {fc_end}",
               "Share change (BPS)", "Growth vs market", "Typical positioning"],
)

S("K12", "Competition", "Industry Structure Landscape (2/2)", "574", "global",
  title="Industry Structure Landscape: {geo} {market} Market, {est_year} (2/2)",
  subtitle="Structural shift from {est_year} to {fc_end}: tier migration, entry and exit, and the forces shaping the tier structure",
  shift_title="Structural Shift, {est_year} to {fc_end}: Tier Migration and Entry or Exit",
  tier_moves=R("tiers", "{item}: {n_item_up} up-tier (M&A) | {n_item_down} down-tier or exit"),
  kpis=["New entrants, {fc_period}: scaling up from the long tail", "Exits and acquisitions: absorbed mainly by Tier 1 and Tier 2 acquirers",
        "Net change in player count, {fc_period}", "Share moving between tiers: largest transfer from {tier_from} to {tier_to}"],
  forces_title="Structural Forces Shaping the Tier Structure (1 = weak, 5 = strong)",
  forces_headers=["Force", "Drivers in the market", "Score", "Effect on tier structure"],
  forces={
      "Entry barriers": "{lx_entry_barriers_cap}",
      "Buyer power": "{lx_competitive_pressure_cap}",
      "Switching costs": "{lx_switching_cap}",
      "Substitution": "{lx_substitutes_cap}",
      "Regulation": "{lx_regulator_cap}",
  },
  effect_bank={
      "high": "Protects Tier 1 and Tier 2 and slows new entry",
      "medium": "Moderates share shifts between tiers",
      "low": "Keeps the market open to challengers",
      "buyer_high": "Compresses Tier 3 pricing",
      "switching_high": "Locks in the installed customer base",
  },
)

S("K13", "Competition", "Competitive Dashboard", "575", "global",
  title="{market} Market Competitive Dashboard ({est_year})",
  subtitle="Companies benchmarked on portfolio, revenue, share, regional presence and strategy",
  intro="Leading players benchmarked on portfolio, revenue, market share, regional presence and strategy; all figures relate to {est_year}.",
  table_headers=["Company", "Key Brands", "Product Portfolio", "Overall Revenue ({c_currency_unit})", "Target-Market Revenue ({val_unit})",
                 "Market Share (%)", "Regional Presence", "Strategies"],
  brand_cell="{c_key_brands}",
  strategy_legend="E = Expansion | I = Innovation | F = Facility investment | P = Partnership | A = Acquisition | L = Product launch | C = Certification",
)

# =============================================================================================
# L. MARKET SHARE AND ADVICE
# =============================================================================================

S("L01", "Market share", "Advice for Vendors and Suppliers", "577", "global",
  title="{market} Market: Advice for Vendors and Suppliers",
  subtitle="Strategic guidance on pricing, product, geography, channels and competitive risk",
  banner="Where to compete, how to price and which risks to manage across the {market} Market.",
  pricing_head="Pricing and Margin Strategy",
  pricing=[
      "Hold price discipline: with the top three players at about {v_cr3} share, avoid value-destructive discounting and compete on mix and service rather than headline price.",
      "Premiumise toward higher-value categories such as '{price_seg_top}' to lift blended average selling price and defend margin as input costs move.",
      "Set transparent, corridor-based {child_label_adj} pricing to limit cross-border arbitrage.",
  ],
  product_head="Product and Innovation",
  product=[
      "Prioritise R&D in the fastest-growing categories, led by '{seg1_fast}', since a differentiated portfolio is the clearest route to outgrowing the {v_f_cagr} market CAGR.",
      "Invest in {lx_validation} to build switching costs and support premium positioning.",
      "Localise {lx_localise} for each channel and geography.",
  ],
  geo_head="Geographic and Channel Expansion",
  geo=[
      "Expand selectively into under-penetrated markets such as {cty_white_the}, where leaders are thin, rather than competing head-on in saturated markets.",
      "Balance coverage across {lx_channel_examples}, with '{seg2_fast}' the fastest-growing route to end customers.",
      "Build reliable distribution and logistics partnerships to convert demand into fulfilled volume.",
  ],
  risk_head="Competitive and Risk Positioning",
  risk=[
      "Benchmark against '{comp1}' (about {v_comp1_share} share), matching it on service and leading on innovation to take share.",
      "Diversify the customer base to reduce dependence on a few large accounts and on single-{child_label} exposure.",
      "Stress-test plans for regulatory, currency and supply shocks, and maintain flexible capacity and dual sourcing.",
  ],
)

S("L02", "Market share", "Company Value Share", "578", "global",
  title="{market} Market: Company Value Share ({geo})",
  subtitle="Value share by company in the {geo_ref} market, {est_year}, highlighting leadership and the long tail",
  banner="Value share of key companies in the target market: leadership, concentration and how the competitive mix shifts by {child_label}.",
  chart_title="{geo} Value Share by Key Company ({val_unit}, {est_year})",
  donut_centre="100% = {v_val_est} {val_unit}",
  takeaways_head="Key Takeaways",
  takeaways=[
      "The {market} Market is {p_concentration} at the {geo_ref} level, as the top three players hold about {v_cr3} of value while a long tail of other suppliers accounts for about {v_others_share}.",
      "'{comp1}' leads with about {v_comp1_share} share, ahead of '{comp2}' at about {v_comp2_share}, and together they set the competitive benchmark for pricing and innovation.",
      "{p_t1_barrier}",
      "{child_label_adj_cap} shares differ from the {geo_ref} picture, and the next slide shows how leadership shifts across geographies.",
      "Share gains are being driven by capacity expansion, product development and selective M&A among the leading players.",
  ],
)

S("L03", "Market share", "Company Share by Segment", "579", "global",
  title="{market} Market: Company Share by {seg1} (Top {n_shown})",
  subtitle="Company value share within each leading {seg1_lc} category, highlighting where leadership differs",
  banner="Company value share within each leading {seg1_lc} category: where each player concentrates and how leadership varies across the portfolio.",
  chart_title=R("share_items", "{item}"),
  takeaways_head="Company Share Key Takeaways",
  takeaways=[
      R("share_items", "{item}: '{item_comp1}' leads at about {v_item_comp1_share}, ahead of '{item_comp2}' at about {v_item_comp2_share}; the top three hold about {v_item_cr3} and the long tail about {v_item_others}."),
      "{p_leader_consistency}",
      "Note: Competitor figures are indicative, derived from market modelling and expert validation rather than company disclosures.",
  ],
)

S("L04", "Market share", "Company Share by Child Geo", "580", "global",
  title="{market} Market: Company Share by {child_label_cap} (Top {n_shown})",
  subtitle="Company value share within the largest {child_label_pl}, highlighting where leadership shifts across geographies",
  banner="Company value share within each of the largest {child_label_pl}: where each player concentrates and how leadership varies by geography.",
  chart_title=R("share_items", "{item}"),
  takeaways_head="Company Share Key Takeaways",
  takeaways=[
      R("share_items", "{item}: '{item_comp1}' leads at about {v_item_comp1_share}, ahead of '{item_comp2}' at about {v_item_comp2_share}; the top three hold about {v_item_cr3} and the long tail about {v_item_others}."),
      "{p_leader_consistency}",
      "Note: Competitor figures are indicative, derived from market modelling and expert validation rather than company disclosures.",
  ],
)

S("L05", "Market share", "Who Shaped Growth", "581", "global",
  title="{market} Market: Who Shaped Growth",
  subtitle="Players ranked by contribution to category growth, with value share, growth role and signature move",
  banner="A framework that identifies the companies with the greatest influence on {market} Market growth, ranked by growth-shaping impact.",
  ranking_head="RANKED BY GROWTH-SHAPING IMPACT",
  ranking_note="Ranked by overall contribution to category growth, not by share alone.",
  card=R("growth_shapers", "{rank}. {item} | Value share: {v_item_share} | Growth role: {item_role} | Why it shaped growth: {item_why} | Signature move: {item_move}"),
  role_bank=["Market leader", "Portfolio expander", "Access builder", "Premiumiser", "Disruptor", "Consolidator"],
  criteria_head="HOW TO DECIDE WHO SHAPED GROWTH",
  criteria=["1. SHARE GAIN: the company gained value or volume share.",
            "2. DISTRIBUTION EXPANSION: the company expanded channel, retail or online reach.",
            "3. INNOVATION AND LAUNCHES: the company introduced new products, formats or improvements.",
            "4. BUYER INFLUENCE: the company strengthened pull with specifiers, channels or end users.",
            "5. BRAND AND PRICING POWER: the company executed premiumisation, positioning or pricing well."],
  usage_head="BEST WAY TO USE THIS SLIDE",
  usage=["Rank companies by their impact on category growth.", "Cite the single move that explains each company's impact.",
         "Do not rely on market share alone."],
  rule_note="Role, reason and move must come from the sourced company sheet; leave the card out rather than default it.",
)

# =============================================================================================
# N. COMPANY PROFILES (real data from secondary sources; c_* is never masked)
# =============================================================================================

S("N01", "Company profile", "Executive Market Snapshot", "583, 588, 593", "company",
  title="{company}: Executive Market Snapshot (1/{n_prof})",
  subtitle="Total market, addressable opportunity and {company}'s obtainable share by {seg2_lc}, {child_label} and {seg1_lc}",
  banner="Executive snapshot of {company}: market size, addressable opportunity and share of obtainable market by {seg2_lc}, {child_label} and {seg1_lc}.",
  kpi_labels=["TAM ({val_unit}): total market", "SAM ({val_unit}): addressable market", "SOM ({val_unit}): {company} share",
              "REVENUE FY{c_y5} ({c_currency_unit}): company-reported", "EBITDA MARGIN FY{c_y5}: company-reported"],
  chart1_title="SOM by {seg2} ({val_unit})",
  chart2_title="SOM by {child_label_cap} ({val_unit})",
  chart3_title="SOM by {seg1} ({val_unit})",
  takeaways_head="Key Takeaways",
  takeaways=[
      "Market position: {company} captures about {v_som_share_sam} of a {v_sam} {val_unit} addressable market, with its strongest {seg2_lc} position in '{c_seg2_lead}'.",
      "Geographic balance: {c_geo_lead} is {company}'s largest {child_label}, while {c_geo_pen_top} shows the highest penetration of addressable demand at about {v_c_geo_pen_top}.",
      "Financial footing: {company} reported revenue of {c_revenue} in FY{c_y5} with an EBITDA margin of {c_ebitda_margin}, according to {c_fin_source}.",
      "Portfolio focus: the company's {market_lc} offer is anchored by {c_key_brands}, with '{c_seg1_lead}' its largest {seg1_lc} category.",
  ],
  private_variant="Financial footing: {company} is privately held and does not publish financial statements, so revenue and margin shown are FMI estimates.",
)

S("N02", "Company profile", "Regional Footprint and Competitive Position", "584, 589, 594", "company",
  title="{company}: Regional Footprint and Competitive Position (2/{n_prof})",
  subtitle="Penetration of addressable demand by {child_label}, competitive position versus peers and TAM, SAM and SOM",
  banner="Regional penetration and competitive position of {company}: SOM versus SAM by {child_label}, position against peers and footprint indicators.",
  chart1_title="SOM vs SAM Penetration by {child_label_cap}",
  chart2_title="Competitive Position: {company} vs Industry Average and Market Leader",
  chart2_axes=["Market Share (%)", "Revenue Growth (%)"],
  funnel=["TAM: 100% | {v_tam} {val_unit}", "SAM: {v_sam_share_tam} of TAM | {v_sam} {val_unit}", "SOM: {v_som_share_sam} of SAM | {v_som} {val_unit}"],
  footprint_head="Footprint Indicators",
  footprint=["Manufacturing sites: {c_plants}", "Countries served: {c_countries}", "Employees: {c_employees}"],
  takeaways_head="Key Takeaways",
  takeaways=[
      "Regional opportunity: {c_geo_pen_top} shows the highest penetration of addressable demand, and lower-penetration {child_label_pl} such as {c_geo_pen_low} represent the main growth headroom.",
      "{p_c_position}",
      "Operating footprint: {company} operates {c_plants} manufacturing sites and serves customers in {c_countries} countries, according to {c_fin_source}.",
  ],
)

S("N03", "Company profile", "SWOT Analysis", "585, 590, 595", "company",
  title="{company}: SWOT Analysis (3/{n_prof})",
  subtitle="Strengths, weaknesses, opportunities and threats shaping {company}'s competitive position",
  banner="SWOT assessment of {company} in the {market} Market.",
  strengths=["{c_strength_1}",
             "Broad presence across {child_label_pl} and {seg2_lc} routes, with its strongest position in {c_geo_lead}, gives the company scale advantages.",
             "A recognised brand portfolio, including {c_key_brands}, supports customer loyalty and pricing."],
  weaknesses=["{c_weakness_1}",
              "Penetration of addressable demand is uneven, strong in {c_geo_pen_top} but thin in {c_geo_pen_low}.",
              "Reliance on a few large accounts or {child_label_pl} raises concentration risk."],
  opportunities=["Premiumisation headroom exists in '{seg1_lead}' and adjacent higher-value categories.",
                 "Under-penetrated {child_label_pl} offer a runway for share capture.",
                 "The shift toward '{seg2_fast}' opens new demand pools."],
  threats=["Volatile costs of {lx_input_costs}, together with price pressure, can squeeze margins.",
           "{lx_competitive_pressure_cap} limit pricing power.",
           "Tier 1 incumbents and {lx_regulator} raise barriers and switching friction."],
  note="Note: The SWOT combines company disclosures with the FMI market model; market figures are indicative.",
)

S("N04", "Company profile", "Key Financials", "586, 591, 596", "company",
  title="{company}: Key Financials (4/{n_prof})",
  subtitle="Financial performance, key ratios and pricing of {company}, highlighting margin and momentum",
  banner="Key financials and pricing for {company} ({c_currency_unit} unless noted; fiscal years ending {c_fy_end}).",
  table1_title="{company} Financial Performance",
  table1_rows=["Revenue", "Gross Profit", "Operating Profit (EBIT)", "Profit Before Tax", "Net Profit"],
  table2_title="{company} Key Ratios",
  table2_rows=["Gross Margin", "EBITDA Margin", "Net Margin", "Return on Equity", "Return on Assets"],
  table_years=["FY{c_y1}", "FY{c_y2}", "FY{c_y3}", "FY{c_y4}", "FY{c_y5}"],
  table3_title="{company} Pricing Analysis ({price_unit})",
  takeaway="{company}'s revenue {p_c_rev_trend} at a {c_rev_cagr} CAGR between FY{c_y1} and FY{c_y5}, while net margin moved from {c_nm_start} to {c_nm_end}.",
  source="Source: {c_fin_source}; {source}",
  private_variant="Financial statements are not publicly disclosed for {company}; figures shown are FMI estimates.",
)

S("N05", "Company profile", "Strategy Analysis", "587, 592, 597", "company",
  title="{company}: Strategy Analysis (5/{n_prof})",
  subtitle="{company}'s strategic priorities and the competitive moves that support them",
  banner="Strategic priorities of {company} and the competitive moves that support them.",
  kpi_labels=["MARKET SHARE: of {geo_ref} value", "REVENUE GROWTH: FY{c_y5} vs FY{c_y4}", "EBITDA MARGIN: FY{c_y5}",
              "COUNTRIES SERVED", "KEY BRANDS IN THIS MARKET"],
  priorities_head="Strategic Priorities",
  priorities=[
      "{c_stated_strategy}",
      "Defend and extend leadership in '{c_seg2_lead}', {company}'s strongest {seg2_lc} route, while scaling the next-best routes.",
      "Raise penetration beyond {c_geo_pen_top} into lower-penetration {child_label_pl} such as {c_geo_pen_low}.",
      "Premiumise the portfolio around '{seg1_lead}' and adjacent higher-value categories to lift blended pricing.",
      "Protect margin through corridor-based pricing and cost discipline as input prices move.",
  ],
  moves_head="Competitive Moves",
  moves=[
      R("company_developments", "{c_dev_date}: {c_dev_text} (Source: {c_dev_source})."),
      "Invest in R&D and selective M&A, backed by {lx_validation}, to move up the price ladder and widen the competitive moat.",
      "Diversify the customer base and secure supply to reduce concentration and input-cost risk.",
  ],
  no_moves_variant="No material {market_lc}-related development by {company} was disclosed in the tracking window.",
)
