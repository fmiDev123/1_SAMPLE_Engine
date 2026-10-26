"""
Classic FMI report style (the format of the 2015 syndicated reports, e.g. REP-GB-717 Digital Wound
Measurement Devices). This style is written so that only numbers, segment and region names, and
figure and table headings change between reports.

Slide anatomy (every data slide):
    title        e.g. "Global DWM Devices Market Analysis for Diabetic Ulcer Wound Type"
    headline     italic one-liner with the key number, e.g. "... is projected to increase 1.7X by ..."
    figures      auto-numbered captions: "Figure 09: Global DWM Devices Market Value (US$ '000) & ..."
    tables       auto-numbered captions: "Table 01: Global DWM Devices Market Value Forecast By ..."
    outlook      the "Market Outlook" (or "Key Takeaways") box, opening with "Figures 09 and 10 highlight ..."
    source       "Source: FMI, 2026"

Numbering: render_deck(plan) walks the slide plan in order, assigns Figure and Table numbers,
fills {fig_a}, {fig_b}, {tab_a}, {tab_b} in the text, and returns the List of Figures and List of Tables.

Placeholders follow the same prefixes as slide_templates.py (v_ masked figures, c_ company data,
n_ counts, in_ research inputs, cp_ classic rule phrases). House rule kept: hyphen ranges, no en dashes.
"""
import re

from .engine import render_text, build_ctx, MissingKeys
from .rules import _same

# =============================================================================================
# Industry defaults for the qualitative slides (drivers, restraints, trends), in 2015 voice.
# Titles feed the Executive Summary, the impact-analysis grids and the regional impact slides.
# Replace or extend per report; {in_*_evidence} lines are optional research inputs.
# =============================================================================================
CLASSIC_DYNAMICS = {
    "healthcare": {
        "drivers_headline": "Rising disease burden, an increasing geriatric population and product innovation are expected to boost overall sales revenue of {market_lc}",
        "restraints_headline": "High cost of technologically advanced products and limited availability of trained professionals are key factors affecting rapid adoption of {market_lc}",
        "drivers": [
            ("Increasing Prevalence of Chronic Diseases", "Increasing prevalence of chronic diseases is raising the number of patients who require diagnosis, monitoring and treatment. This, in turn, has led to a surge in demand for efficient care solutions, including {market_lc}, in recent years. Healthcare professionals, especially in developed regions, are increasingly inclined towards adoption of advanced products to improve clinical outcomes."),
            ("Increasing Ageing Population", "The ageing population is more susceptible to chronic conditions and requires more frequent medical care. Increasing geriatric population across the globe is expected to drive demand for {market_lc} over the forecast period."),
            ("Product Innovation", "Manufacturers are focused on introducing innovative and technologically advanced products to capture market share. Continuous product launches are expected to support adoption of {market_lc} across healthcare settings."),
        ],
        "restraints": [
            ("High Cost of Advanced Products", "High cost of technologically advanced {market_lc} is a major barrier for growth of the global market. Healthcare providers, especially in developing regions, prefer conventional alternatives that are available at a fraction of the cost."),
            ("Lack of Skilled Professionals", "Lack of awareness and a shortage of trained professionals to operate advanced products are affecting market growth to a certain extent."),
        ],
        "trends": [
            ("Shift from Conventional to Advanced Products", "Players in the market are continuously developing and introducing newer and advanced products through technological innovation, to cater to the steady shift from conventional products to advanced options."),
            ("Increasing Healthcare Expenditure in Emerging Economies", "Healthcare expenditure in Asia Pacific and other emerging regions is increasing rapidly owing to investment by private as well as public organisations."),
            ("Extensive Research and Development", "Manufacturers are constantly investing in research and development activities to design, develop and introduce technologically advanced products and to upgrade their existing product portfolios."),
        ],
    },
    "pharma": {
        "drivers_headline": "Growing patient pool, expanding access to treatment and a strong product pipeline are expected to boost overall sales revenue of {market_lc}",
        "restraints_headline": "Pricing pressure and stringent regulatory requirements are key factors restraining growth of the {market_lc} market",
        "drivers": [
            ("Growing Patient Population", "Rising incidence of the target conditions is increasing the treated patient pool. This, in turn, is expected to support demand for {market_lc} over the forecast period."),
            ("Improving Access to Treatment", "Expansion of healthcare coverage and reimbursement in developed as well as emerging economies is improving patient access to {market_lc}."),
            ("Strong Product Pipeline", "Manufacturers are investing in new formulations and lifecycle management of existing brands, which is expected to create new growth avenues for {market_lc}."),
        ],
        "restraints": [
            ("Pricing Pressure and Generic Competition", "Patent expiries and payer-led price negotiations are exerting downward pressure on prices, which restrains revenue growth of {market_lc}."),
            ("Stringent Regulatory Requirements", "Lengthy and costly approval processes delay product launches and increase the cost of market entry."),
        ],
        "trends": [
            ("Shift towards Patient-Centric Formulations", "Manufacturers are developing convenient dosage forms to improve patient adherence and differentiate their portfolios."),
            ("Growing Role of Contract Manufacturing", "Companies are increasingly outsourcing manufacturing to CDMOs to reduce cost and improve flexibility."),
            ("Expansion in Emerging Markets", "Leading players are strengthening their presence in Asia Pacific and Latin America through partnerships with local distributors."),
        ],
    },
    "food": {
        "drivers_headline": "Changing consumer lifestyles, product innovation and expansion of modern retail are expected to boost overall sales revenue of {market_lc}",
        "restraints_headline": "Volatility in raw material prices and rising health concerns are key factors affecting growth of the {market_lc} market",
        "drivers": [
            ("Rising Demand for Convenient Food", "Busy lifestyles and an increasing number of working professionals are driving demand for convenient, ready-to-eat food products. This, in turn, is expected to support sales of {market_lc} over the forecast period."),
            ("Product Innovation in Flavours and Formats", "Manufacturers are focused on introducing new flavours, formats and better-for-you variants to attract consumers and capture market share."),
            ("Expansion of Modern Retail", "Growth of supermarkets, hypermarkets, convenience stores and online grocery platforms is improving product visibility and availability of {market_lc}."),
        ],
        "restraints": [
            ("Volatility in Raw Material Prices", "Fluctuating prices of agricultural commodities, packaging and energy increase production cost and squeeze manufacturer margins."),
            ("Rising Health Concerns", "Growing consumer awareness of sugar, fat and additive content is shifting part of the demand towards healthier alternatives."),
        ],
        "trends": [
            ("Clean-Label and Better-for-You Products", "Manufacturers are reformulating products with natural ingredients and reduced sugar to cater to health-conscious consumers."),
            ("Premiumisation", "Consumers are increasingly willing to pay a premium for indulgent, artisanal and premium-ingredient products."),
            ("Growth of Online Retail", "E-commerce and quick-commerce platforms are emerging as important sales channels for packaged food products."),
        ],
    },
    "consumer": {
        "drivers_headline": "Rising disposable income, product innovation and expanding distribution are expected to boost overall sales revenue of {market_lc}",
        "restraints_headline": "Intense price competition and volatility in raw material prices are key factors affecting growth of the {market_lc} market",
        "drivers": [
            ("Rising Disposable Income", "Increasing disposable income, especially in emerging economies, is supporting higher spending on {market_lc}."),
            ("Product Innovation", "Manufacturers are introducing new designs, materials and features to attract consumers and capture market share."),
            ("Expanding Distribution Network", "Growth of organised retail and e-commerce is improving availability of {market_lc} across urban and semi-urban areas."),
        ],
        "restraints": [
            ("Intense Price Competition", "Presence of a large number of regional and private-label players is intensifying price competition and limiting margins."),
            ("Volatility in Raw Material Prices", "Fluctuating raw material and logistics costs increase production cost for manufacturers."),
        ],
        "trends": [
            ("Premiumisation", "Consumers are increasingly inclined towards premium and branded products."),
            ("Sustainability", "Manufacturers are adopting sustainable materials and packaging to meet changing consumer preferences and regulations."),
            ("Direct-to-Consumer Sales", "Brands are increasingly selling through their own online channels to build direct relationships with consumers."),
        ],
    },
    "industrial": {
        "drivers_headline": "Growing industrial output, infrastructure investment and technological advancement are expected to boost overall sales revenue of {market_lc}",
        "restraints_headline": "Volatility in raw material prices and high capital investment are key factors affecting growth of the {market_lc} market",
        "drivers": [
            ("Growth in End-Use Industries", "Increasing production activity in key end-use industries is driving demand for {market_lc} across the globe."),
            ("Infrastructure Investment", "Rising public and private investment in infrastructure, especially in emerging economies, is expected to create favourable market prospects for {market_lc}."),
            ("Technological Advancement", "Manufacturers are introducing higher-performance and more efficient products to cater to evolving customer requirements."),
        ],
        "restraints": [
            ("Volatility in Raw Material Prices", "Fluctuating prices of raw materials and energy increase production cost and affect manufacturer margins."),
            ("High Capital Investment", "High initial investment and long qualification cycles act as a barrier for new entrants."),
        ],
        "trends": [
            ("Automation and Digitalisation", "Manufacturers are adopting automation and digital monitoring to improve efficiency and product quality."),
            ("Sustainable Products", "Stringent environmental regulations are encouraging development of sustainable and energy-efficient products."),
            ("Capacity Expansion in Asia Pacific", "Leading players are expanding production capacity in Asia Pacific to cater to growing regional demand."),
        ],
    },
}


def dynamics_ctx(industry):
    d = CLASSIC_DYNAMICS[industry]
    out = {"drivers_headline": d["drivers_headline"], "restraints_headline": d["restraints_headline"]}
    for key in ("drivers", "restraints", "trends"):
        out[key] = [dict(item=t, item_text=x) for t, x in d[key]]
        out[f"{key}_titles"] = [t for t, _ in d[key]]
    return out


# =============================================================================================
# Classic slide templates
# =============================================================================================
CLASSIC = []


def K(cid, chapter, name, ref, maps_to, title, headline, figures=(), tables=(), **elements):
    CLASSIC.append(dict(id=cid, chapter=chapter, name=name, ref=ref, maps_to=maps_to, title=title, headline=headline,
                        figures=list(figures), tables=list(tables), elements=elements))


def R(over, text):
    return {"repeat": over, "text": text}


SOURCE = "Source: FMI, {pub_year}"
FOOTER = "©{pub_year} {publisher}, All Rights Reserved"

# ------------------------------------------------------------------------------------- front matter
K("R01", "Front matter", "Cover", "1", "A01",
  title="{market} Market: Global Industry Analysis and Opportunity Assessment, {hist_start}-{fc_end}",
  headline="",
  report_code="{report_code}",
)

K("R02", "Front matter", "Table of Contents", "2-6", "A09",
  title="Table of Contents{page_tag}", headline="",
  toc=[
      "Executive Summary",
      "Research Methodology",
      "Assumptions, Terminologies and Acronyms Used",
      "What Makes {market} Market Attractive",
      "Market Taxonomy",
      "Market Definition",
      "Global {market} Market Overview",
      "Global {market} Market Dynamics",
      "    Drivers",
      "    Restraints",
      "    Trends",
      R("segments", "Global {market} Market Analysis and Forecast By {item}"),
      R("segments", "    Introduction | Y-o-Y Growth Comparison By {item} | BPS Analysis By {item} | Market Attractiveness Analysis By {item} | Market Size and Forecast By {item}"),
      "Global {market} Market Analysis and Forecast By Region",
      R("regions", "{item} {market} Market Analysis and Forecast | Impact Analysis | Market Size and Forecast By Country | Market Size and Forecast By Segment"),
      "Competitive Landscape",
      R("companies", "    {item}"),
  ],
)

K("R03", "Front matter", "List of Figures", "7-9", "A10",
  title="List of Figures{page_tag}", headline="",
  rule_note="Generated by render_deck(): one line per figure, 'Figure {n}: {caption}', in deck order.",
)

K("R04", "Front matter", "List of Tables", "10-11", "A11",
  title="List of Tables{page_tag}", headline="",
  rule_note="Generated by render_deck(): one line per table, 'Table {n}: {caption}', in deck order.",
)

K("R05", "Front matter", "Section divider", "12, 15, 18 ...", "A03",
  title="Section {sec_no}", headline="",
  section_title="{section_title}",
)

# ------------------------------------------------------------------------------------- executive summary
K("R06", "Executive summary", "Executive Summary (1/2)", "13", "A05, G06",
  title="Executive Summary (1/2)",
  headline="{cp_es_headline}",
  figures=["Global {mkt_short} Market Y-o-Y Growth By Region, {est_year}-{fc_end} (%)"],
  subhead="Global demand for {market_lc} is expected to gain momentum in {child_fast} by the mid of the forecast period, {est_year}-{fc_end}",
  foreword_head="Foreword",
  foreword=[
      "Global demand for {market_lc} is increasing as a result of {cp_driver1_lc} and {cp_driver2_lc}. FMI presents in-depth analysis on the outlook of the {market_lc} market for the forecast period, {est_year}-{fc_end}.",
      "{child_lead} is expected to remain the largest market for {market_lc}, while demand in {child_fast} is projected to increase at the highest CAGR of {v_child_fast_cagr} throughout the forecast period.",
      "Globally, {cp_driver3_lc} is expected to create favourable market prospects for {market_lc} manufacturers.",
  ],
)

K("R07", "Executive summary", "Executive Summary (2/2)", "14", "G06",
  title="Executive Summary (2/2)",
  headline="Global {market_lc} market is expected to witness {v_mult}X increase in revenue by {fc_end} end",
  overview_head="Overview",
  overview=[
      "Global {market_lc} market accounted for US$ {v_val_base} {val_word} in {base_year}, which is expected to increase to US$ {v_val_end} {val_word} in {fc_end} at a CAGR of {v_f_cagr}.",
      "The global {mkt_short_lc} market has been analysed on the basis of {segments_lc} across the globe.",
  ],
  drivers_head="Drivers",
  drivers=R("drivers", "{item}"),
  challenges_head="Challenges",
  challenges=R("restraints", "{item}"),
  competition_head="Competitive Landscape",
  competition=["{cp_competition}"],
)

# ------------------------------------------------------------------------------------- methodology etc.
K("R08", "Methodology", "Research Methodology (1/2)", "16", "",
  title="Research Methodology (1/2)", headline="",
  approach_head="SYSTEMATIC RESEARCH APPROACH",
  steps=["Market profiling", "Formulating discussion guide", "Developing list of respondents", "Data collection",
         "Data validation", "Data analysis", "Insights"],
  desk_research="In-depth secondary research is used to ascertain overall market size, top industry players, top products, industry associations, etc.",
  primary_research="FMI formulates a detailed discussion guide to conduct expert and industry interviews. FMI develops a list of industry players (manufacturers), distributors, retailers and industry experts, and conducts interviews with each group.",
  company_analysis="Market participants, key strengths, product portfolio, mapping as per value chain and key focus segments are analysed for each company.",
  validation="Data is validated by the triangulation method, wherein secondary research, primary research and FMI analysis contribute to the final data.",
)

K("R09", "Methodology", "Research Methodology (2/2)", "17", "",
  title="Research Methodology (2/2)", headline="",
  secondary_head="Secondary Research",
  secondary=["{mkt_short_lc} market trends and forecast", "{mkt_short_lc} market historical trends",
             "{mkt_short_lc} market region-wise trends", "{mkt_short_lc} market growth projections"],
  sources_head="Key Sources Referred",
  sources=["Industry association publications", "Annual reports, publications and presentations of key players",
           "World Bank, IMF and government websites", "{in_sector_sources}"],
  primary_head="Primary Research",
  primary=["Major drivers, restraints and opportunities in the global {mkt_short_lc} market",
           "Key challenges faced by manufacturers and suppliers",
           "Market segmentation on the basis of {segments_lc}, along with region", "Average selling price of {market_lc}"],
  analysis_head="Data Analysis",
  analysis=["Present and future market value potential", "Anticipated growth projections", "Market concentration",
            "Major trends", "Historical trend analysis", "Statistical analysis"],
  key_sources_head="Key Sources",
  key_sources=["Industry experts", "Manufacturers", "Distributors and suppliers"],
  reporting_head="Data Reporting",
  reporting=["Region-wise {mkt_short_lc} market was analysed to arrive at the global market value.",
             "CAGR and Y-o-Y growth rate were used to arrive at the global market value.",
             "Qualitative data and facts, quantitative data and market dynamics are reported."],
)

K("R10", "Methodology", "Assumptions and Acronyms", "19", "",
  title="Assumptions and Acronyms", headline="",
  assumptions_head="Assumptions",
  assumptions=[
      "Yearly change in inflation rate has not been considered while forecasting market numbers.",
      "Bottom-up approach has been used to assess market numbers for each product category.",
      "Top-down approach has been used to counter-validate the market numbers and segment-wise market numbers.",
      "Product-wise weighted average selling price has been considered to deduce market value.",
      "Macro-economic indicators such as GDP, sector spending, end-use business scenario, trade and prices have been considered to arrive at the indicated market numbers.",
      "Secondary axis represents Y-o-Y growth.",
  ],
  currency="Currency: All the indicated denominations are in US$ (US Dollar).",
  source_line="Source: Company annual and financial reports, paid sources.",
  acronyms=["CAGR: Compound Annual Growth Rate", "Y-o-Y: Year-on-Year", "Mn: Million", "Bn: Billion", "'000: Thousand",
            "BPS: Basis Point Share", "ASP: Average Selling Price", "MEA: Middle East and Africa",
            "{mkt_short_acronym_line}"],
  terminology=["Absolute $ Opportunity: increment in total market revenue between consecutive years.",
               "Market Attractiveness: assessment metric based on incremental growth and CAGR of the target segment or market."],
)

K("R11", "Market background", "What Makes the Market Attractive", "21-23", "",
  title="What Makes {market} Market Attractive{page_tag}",
  headline="",
  figures=["{attr_fig_title}"],
  commentary=R("attr_points", "{item}"),
  parent_market_lines=[
      "Global {parent_market_lc} market is expected to witness {v_parent_mult}X increase in its revenue by {fc_end} end.",
      "Global demand for {parent_market_lc} is expected to create a favourable climate for sales of {market_lc} in the coming years.",
  ],
  rule_note="Three pages: parent market (figure 'Global {parent_market} Market ({parent_unit}), {base_year} A and {fc_end} F' with parent_market_lines), a macro indicator (e.g. GDP vs sector spending, source World Bank) and a demographic or epidemiological indicator (source United Nations or WHO). Inputs per page: attr_fig_title, chart data, attr_points (one or two sentences).",
)

K("R12", "Market background", "Market Introduction: Definition", "25", "",
  title="{market} Market Introduction: Definition (1/2)", headline="",
  definition="{in_definition}",
  unit_definition="One unit refers to one {in_unit_definition}.",
  workflow_head="{in_workflow_title}",
)

K("R13", "Market background", "Market Introduction: Taxonomy", "26", "A06, A07",
  title="{market} Market Introduction: Taxonomy (2/2)", headline="",
  regions_head="Regions",
  regions=R("regions", "{item}"),
  segment_boxes=R("segments", "{item}"),
)

K("R14", "Market overview", "Global Market Overview (value)", "28", "G03, G04",
  title="Global {market} - Overview",
  headline="{cp_overview_head}",
  figures=["Global {market} Market Value Forecast ({val_unit}), {hist_start}-{fc_end}"],
  outlook_head="Market Outlook",
  outlook=[
      "Figure {fig_a} highlights overall value of the {market_lc} market from {hist_start} to {fc_end}.",
      "Global {market_lc} market is expected to be valued at US$ {v_val_end} {val_word} by {fc_end} end.",
      "It is projected to create overall incremental opportunity of US$ {v_abs_opp} {val_word} from {base_year} to {fc_end}.",
      "{cp_driver1_cap} along with {cp_driver2_lc} is fuelling demand for {market_lc}.",
  ],
)

K("R14V", "Market overview", "Global Market Overview (volume)", "new", "G01",
  title="Global {market} - Volume Overview",
  headline="Global {market_lc} market is projected to register {v_vol_mult}X increase in volume by the end of forecast period, {fc_end}",
  figures=["Global {market} Market Volume Forecast ({vol_unit}), {hist_start}-{fc_end}"],
  outlook_head="Market Outlook",
  outlook=[
      "Figure {fig_a} highlights overall volume of the {market_lc} market from {hist_start} to {fc_end}.",
      "Global {market_lc} market volume stood at {v_vol_base} {vol_unit} in {base_year} and is expected to reach {v_vol_end} {vol_unit} by {fc_end} end, at a CAGR of {v_vol_f_cagr}.",
      "{cp_price_mix}",
  ],
)

# ------------------------------------------------------------------------------------- dynamics
K("R15", "Market dynamics", "Drivers and Restraints (overview)", "30", "",
  title="{market} Market: Drivers and Restraints", headline="",
  centre="{market} Market",
  drivers=R("drivers", "{item}"),
  restraints=R("restraints", "{item}"),
)

K("R16", "Market dynamics", "Drivers", "31", "M02, M03",
  title="{market} Market: Drivers",
  headline="{drivers_headline}",
  driver_title=R("drivers", "{item}"),
  driver_text=R("drivers", "{item_text}"),
  rule_note="Add one 'For instance, ...' line per driver from research input (in_driver_evidence) where available; never invent statistics.",
)

K("R17", "Market dynamics", "Restraints", "32", "M04",
  title="{market} Market: Restraints",
  headline="{restraints_headline}",
  restraint_title=R("restraints", "{item}"),
  restraint_text=R("restraints", "{item_text}"),
)

K("R18", "Market dynamics", "Market Dynamics Impact Analysis", "33", "M01",
  title="Market Dynamics Impact Analysis",
  headline="{cp_impact_headline}",
  period="{est_year}-{fc_end}",
  axes=["Impact on {mkt_short} Market: Low to High", "Certainty: Low to High"],
  bubbles=R("impact_items", "{item}"),
)

K("R19", "Market dynamics", "Trends", "34", "G07",
  title="{market} Market: Trends", headline="",
  table_headers=["Trends", "Description"],
  rows=R("trends", "{item} | {item_text}"),
)

# ------------------------------------------------------------------------------------- segment chapter
K("R20", "Segment chapter", "Y-o-Y Growth by Segment", "36, 46, 53", "D07, D08, B04, C05",
  title="{geo} {market} Market Y-o-Y Growth Rate By {segment}",
  headline="In terms of revenue contribution, {sub_lead_lc} segment is expected to dominate the {geo_ref} {mkt_short_lc} market throughout the forecast period",
  figures=["{geo} {mkt_short} Market Y-o-Y Growth By {segment}, {est_year}-{fc_end} (%)"],
  outlook_head="Market Outlook",
  outlook=[
      "Figure {fig_a} highlights Y-o-Y growth for the {geo_ref} {market_lc} market by {segment_lc}.",
      "{sub_lead_cap} segment is projected to exhibit an average Y-o-Y growth rate of {v_sub_lead_avg_yoy} between {est_year} and {fc_end}.",
      "Y-o-Y growth of the {sub_fast_lc} segment is projected to {cp_yoy_move}.",
  ],
)

K("R21", "Segment chapter", "Revenue Proportion (BPS) by Segment", "37, 47, 54", "B01, B02, B03, D06, C03",
  title="{geo} {market} Market Revenue Proportion By {segment}",
  headline="Among {mkt_short_lc} by {segment_lc}, {sub_fast_lc} segment is expected to witness the highest CAGR of {v_sub_fast_cagr} over the forecast period",
  figures=["{geo} {mkt_short} Market Value Share By {segment}, {est_year}-{fc_end} (%)"],
  outlook_head="Market Outlook",
  outlook=[
      "Figure {fig_a} highlights {geo_ref} {market_lc} market revenue proportion by {segment_lc} for {est_year}, {mid_year} and {fc_end}.",
      "In terms of value share by {segment_lc}, the {sub_lead_lc} segment is expected to dominate the market, {cp_lead_bps_move} and accounting for {v_sub_lead_share_end} share by {fc_end}.",
      "{cp_loser_line}",
  ],
)

K("R22", "Segment chapter", "Attractiveness Index by Segment", "38, 48, 55", "D02",
  title="{geo} {market} Market Attractiveness Index By {segment}",
  headline="{cp_attr_headline}",
  figures=["{geo} {mkt_short} Market Value Attractiveness By {segment}, {est_year}-{fc_end}"],
  axes=["Attractiveness Index", "% CAGR ({est_year}-{fc_end})"],
  outlook_head="Market Outlook",
  outlook=[
      "Figure {fig_a} represents attractiveness index of the {geo_ref} {market_lc} market on the basis of {segment_lc}.",
      "{cp_attr_line}",
      "{cp_attr_low_line}",
  ],
)

K("R23", "Segment chapter", "Analysis for each Sub-segment", "39-43, 49-50", "D01, D03, D04, D05",
  title="{geo} {mkt_short} Market Analysis for {item_title} {seg_tag}",
  headline="By {fc_end}, {item_lc} segment is anticipated to {cp_item_mult}",
  figures=["{geo} {mkt_short} Market {metric_cap} ({m_unit}) & Y-o-Y Growth (%) for {item_title} {seg_tag}, {hist_start}-{fc_end}",
           "{geo} {mkt_short} Market Absolute $ Opportunity ({val_unit}) for {item_title} {seg_tag}, {hist_start}-{fc_end}"],
  chart_label="CAGR {v_item_cagr}",
  outlook_head="Market Outlook",
  outlook__value=[
      "Figures {fig_a} and {fig_b} highlight {item_lc} segment revenue contribution and overall incremental opportunity from {base_year} to {fc_end}.",
      "{cp_item_base_line}",
      "The {item_lc} segment is expected to account for US$ {v_item_end} {val_word} by {fc_end} at a CAGR of {v_item_cagr} over the forecast period.",
      "This segment is expected to create an incremental opportunity of US$ {v_item_opp} {val_word} for the overall {mkt_short_lc} market from {base_year} to {fc_end}.",
      "Absolute $ value for {est_year} is estimated to be US$ {v_item_abs_est} {val_word}.",
  ],
  outlook__volume=[
      "Figures {fig_a} and {fig_b} highlight {item_lc} segment volume contribution and overall incremental opportunity from {base_year} to {fc_end}.",
      "In {base_year}, {item_lc} segment volume stood at {v_item_base} {vol_unit} and is estimated to reach {v_item_est} {vol_unit} by the end of {est_year}, reflecting a Y-o-Y growth rate of {v_item_yoy_est}.",
      "The {item_lc} segment is expected to account for {v_item_end} {vol_unit} by {fc_end} at a CAGR of {v_item_cagr} over the forecast period.",
      "This segment is expected to create an incremental opportunity of US$ {v_item_opp} {val_word} for the overall {mkt_short_lc} market from {base_year} to {fc_end}.",
  ],
  rule_note="One slide per sub-segment. Variant chosen by ctx['metric'] ('value' or 'volume').",
)

K("R24", "Segment chapter", "Value Forecast Table by Segment", "44, 51, 56", "D04, D05, B03",
  title="{geo} {mkt_short} Market {metric_cap} Forecast By {segment}",
  headline="{cp_table_headline}",
  tables=["{geo} {mkt_short} Market {metric_cap} Forecast By {segment}, {hist_start}-{fc_end} ({m_unit})"],
  table_header_first="{segment}",
  table_header_last="CAGR ({est_year}-{fc_end})",
  table_total="Total",
)

# ------------------------------------------------------------------------------------- regional chapter
K("R25", "Regional chapter", "Regional Market Introduction", "58, 64, 70 ...", "C02",
  title="{geo} {market} Market Introduction",
  headline="",
  chart1_title="{geo} Market Value By {seg1}, {base_year} ({val_unit})",
  chart1_labels=["Market Value ({val_unit})", "CAGR"],
  chart2_title="{geo} Market Value By Country, {base_year} ({val_unit})",
)

K("R26", "Regional chapter", "Regional Market Analysis", "59, 65, 71 ...", "G03, G04, C01",
  title="{geo} {mkt_short} Market Analysis",
  headline="{mkt_short_lc_cap} market value in {geo_the} is expected to {cp_geo_mult} between {base_year} and {fc_end}",
  figures=["{geo} {mkt_short} Market Value Forecast ({val_unit}) & Y-o-Y Growth (%), {hist_start}-{fc_end}",
           "{geo} {mkt_short} Market Absolute $ Opportunity ({val_unit}), {hist_start}-{fc_end}"],
  chart_label="CAGR {v_f_cagr}",
  outlook_head="Key Takeaways",
  outlook=[
      "Figures {fig_a} and {fig_b} highlight {geo_the} market revenue and overall incremental opportunity from {base_year} to {fc_end}.",
      "In {base_year}, the {mkt_short_lc} market in {geo_the} was valued at US$ {v_val_base} {val_word} and is estimated to account for US$ {v_val_est} {val_word} by the end of {est_year}, reflecting a Y-o-Y growth rate of {v_yoy_est}.",
      "The market in {geo_the} is expected to account for US$ {v_val_end} {val_word} by {fc_end} at a CAGR of {v_f_cagr} over the forecast period.",
      "{geo_The} is expected to create an incremental opportunity of US$ {v_abs_opp} {val_word} for the overall {mkt_short_lc} market from {base_year} to {fc_end}.",
      "{cp_geo_country_line}",
      "Absolute $ value for {est_year} is estimated to be US$ {v_abs_est} {val_word}.",
  ],
)

K("R27", "Regional chapter", "Regional Impact Analysis", "60, 66, 72 ...", "M01",
  title="{geo} {market} Market: Impact Analysis", headline="",
  rows=R("impact_items", "{item} | {item_level}"),
  legend=["Drivers", "Restraints", "Low", "Medium", "High"],
  rule_note="Impact level per region comes from the model score via rules.likelihood(): Low, Medium or High.",
)

K("R28", "Regional chapter", "Regional Value Forecast by Country and Segment", "61, 67, 73-74 ...", "B03, B05, C04, G05",
  title="{geo} {mkt_short} Market Value Forecast By Country & {seg2}",
  headline="{cp_country_headline}",
  tables=["{geo} {mkt_short} Market Value Forecast By Country, {hist_start}-{fc_end} ({val_unit})",
          "{geo} {mkt_short} Market Value Forecast By {seg2}, {hist_start}-{fc_end} ({val_unit})"],
  table_header_last="CAGR ({est_year}-{fc_end})",
  table_total="Total",
)

K("R29", "Regional chapter", "Regional Value Forecast by Primary Segment", "62, 68, 75 ...", "C04",
  title="{geo} {mkt_short} Market Value Forecast By {seg1}",
  headline="{cp_seg1_table_headline}",
  tables=["{geo} {mkt_short} Market Value Forecast By {seg1}, {hist_start}-{fc_end} ({val_unit})"],
  table_header_last="CAGR ({est_year}-{fc_end})",
  table_total="Total",
)

# ------------------------------------------------------------------------------------- additions for the current deck
K("R30", "Pricing", "Average Price Analysis", "new", "G02, P01, P02, P04, P05, P09, P10",
  title="{geo} {mkt_short} Market Average Price Analysis By {seg1}",
  headline="{price_seg_top_cap} segment is expected to command the highest average price of US$ {v_price_seg_top_end} per {vol_unit_single} in {fc_end}",
  figures=["{geo} {mkt_short} Market Average Price ({price_unit}) By {seg1}, {hist_start}-{fc_end}",
           "{geo} {mkt_short} Market Average Price ({price_unit}) By Region, {base_year} and {fc_end}"],
  outlook_head="Market Outlook",
  outlook=[
      "Figures {fig_a} and {fig_b} highlight average selling price of {market_lc} by {seg1_lc} and by region.",
      "Average selling price of {market_lc} is estimated at US$ {v_price_avg_base} per {vol_unit_single} in {base_year} and is projected to {cp_price_dir} US$ {v_price_avg_end} per {vol_unit_single} by {fc_end}.",
      "{child_price_high} is the highest-priced region, while {child_price_low} is the most price-competitive market.",
      "Average price is derived as market value divided by market volume.",
  ],
)

K("R31", "Cross-section", "Value Forecast by Two Segments", "new", "X01, X02, X03",
  title="{geo} {mkt_short} Market Value Forecast By {seg_a} and {seg_b}",
  headline="{xa_lead_cap} segment is projected to account for the largest share of the {geo_ref} {mkt_short_lc} market, led by {xb_lead_in_a_lc} at US$ {v_xb_lead_in_a_end} {val_word} in {fc_end}",
  tables=["{geo} {mkt_short} Market Value Forecast By {seg_a} and {seg_b}, {est_year} and {fc_end} ({val_unit})"],
)

K("R32", "Competition", "Market Share Analysis by Key Players", "new", "L02, L03, L04, K13",
  title="Global {mkt_short} Market Share Analysis By Key Players",
  headline="{comp1} accounted for the largest share of the global {mkt_short_lc} market in {base_year}",
  figures=["Global {mkt_short} Market Share Analysis By Key Players, {base_year} (%)"],
  tables=["Global {mkt_short} Market: Competition Dashboard, {base_year}"],
  outlook_head="Market Outlook",
  outlook=[
      "Figure {fig_a} highlights market share of key players in the global {mkt_short_lc} market in {base_year}.",
      "{cp_competition}",
      "{comp1} accounted for an estimated {v_comp1_share} share of the market, followed by {comp2} with {v_comp2_share}.",
      "Key players are focusing on product innovation, geographic expansion and strategic partnerships to strengthen their market position.",
  ],
)

K("R33", "Competitive landscape", "Company Profile (1/2): Company Details and Overview", "101, 103", "N01, N02, N04",
  title="{company} (1/2)", headline="",
  details_head="Company Details",
  details=["Company Name | {company}", "Established | {c_established}", "Headquartered | {c_hq}",
           "Chairman / CEO / President | {c_ceo}"],
  figures=["{company} Segment Revenue Split, {c_fy} (%)"],
  overview_head="Overview",
  overview=R("c_overview", "{item}"),
  source="Source: {c_source}",
  rule_note="Overview bullets come from the company website or annual report: founding, headquarters, what the company makes, where it operates and key customers.",
)

K("R34", "Competitive landscape", "Company Profile (2/2): Products, Strategy and Key Developments", "102, 104", "N03, N05",
  title="Company Profile - {company} (2/2)", headline="",
  products_head="PRODUCT OFFERING",
  products_intro="Products offered by the company include:",
  products=R("c_products", "{item}: {item_text}"),
  strategy_head="STRATEGY",
  strategy="{c_strategy}",
  developments_head="KEY DEVELOPMENTS",
  developments=R("c_developments", "In {item_year}, {item_text}"),
  no_developments="No major development related to the {market_lc} market was announced by {company} during {c_dev_window}.",
)

CLASSIC_INDEX = {k["id"]: k for k in CLASSIC}


# =============================================================================================
# Classic rules
# =============================================================================================
def _lower_name(s):
    """Lower-case a segment name for mid-sentence use, keeping acronyms and dotted country codes."""
    words = []
    for w in s.split():
        words.append(w if (w.isupper() and len(w) > 1) or "." in w or any(ch.isdigit() for ch in w) else w.lower())
    return " ".join(words)


def _title_name(s):
    small = {"and", "or", "of", "for", "the", "in", "on", "to", "with", "by", "a", "an"}
    out = []
    for i, w in enumerate(s.split()):
        if (w.isupper() and len(w) > 1) or "." in w or any(ch.isdigit() for ch in w):
            out.append(w)
        elif i and w.lower() in small:
            out.append(w.lower())
        else:
            out.append("-".join(p[:1].upper() + p[1:] for p in w.split("-")))
    return " ".join(out)


def _cap(s):
    return s[:1].upper() + s[1:] if s else s


NAME_KEYS_LC = ("item", "sub_lead", "sub_fast", "sub_second", "sub_slow", "sub_lose", "sub_attr1", "sub_attr2", "sub_attr_low",
                "price_seg_top", "seg1_lead", "xa_lead", "xb_lead_in_a", "seg_a", "seg_b", "seg1", "seg2", "segment")


def mult_phrase(m, vkey):
    if m >= 1.05:
        return f"witness an increase of {{{vkey}}}X in terms of {{metric}} as compared to {{base_year}}"
    if m <= 0.95:
        return f"decline to {{{vkey}}}X of its {{base_year}} {{metric}}"
    return "remain broadly stable in terms of {metric} as compared to {base_year}"


def classic_derive(c):
    p = {}
    for k in NAME_KEYS_LC:
        if isinstance(c.get(k), str):
            p[k + "_lc"] = c.get(k + "_lc", _lower_name(c[k]))
            p[k + "_cap"] = _cap(p[k + "_lc"])
    if "segments" in c:
        p["segments_lc"] = " and ".join(_lower_name(s) for s in c["segments"]) if len(c["segments"]) <= 2 else \
            ", ".join(_lower_name(s) for s in c["segments"][:-1]) + " and " + _lower_name(c["segments"][-1])
    dr = c.get("drivers_titles")
    if dr:
        p["cp_driver1_lc"], p["cp_driver2_lc"] = _lower_name(dr[0]), _lower_name(dr[1])
        p["cp_driver3_lc"] = _lower_name(dr[2]) if len(dr) > 2 else _lower_name(dr[0])
        p["cp_driver1_cap"] = _cap(p["cp_driver1_lc"])
        p["cp_es_headline"] = (f"{_cap(_lower_name(dr[0]))} along with {_lower_name(dr[1])} is expected to boost overall growth of "
                               f"the {{market_lc}} market in the near future")
        rs = c.get("restraints_titles", [])
        p["cp_impact_headline"] = (f"{_cap(_lower_name(dr[0]))} along with {_lower_name(dr[1])} is expected to boost demand for "
                                   f"{{market_lc}} in the near future" + (f", while {_lower_name(rs[0])} remains the key challenge" if rs else ""))
    if "r_mult" in c:
        m = c["r_mult"]
        p["cp_overview_head"] = ("{market_lc_cap} market is projected to register {v_mult}X increase in its revenue by the end of forecast period, {fc_end}"
                                 if m >= 1.05 else
                                 "{market_lc_cap} market is projected to remain broadly stable in revenue terms through the end of forecast period, {fc_end}")
        p["cp_geo_mult"] = "increase {v_mult}X" if m >= 1.05 else ("decline to {v_mult}X of its current value" if m <= 0.95 else "remain broadly stable")
    if "r_f_cagr" in c and "r_vol_f_cagr" in c:
        d = (c["r_f_cagr"] - c["r_vol_f_cagr"]) * 10000
        p["cp_price_mix"] = ("Value is projected to grow faster than volume over the forecast period, owing to an increase in average selling price."
                             if d >= 25 else
                             "Value is projected to grow slower than volume over the forecast period, owing to a decline in average selling price."
                             if d <= -25 else
                             "Value and volume are projected to grow at a similar pace over the forecast period, as average selling price remains broadly stable.")
    if "r_sub_fast_yoy_start" in c:
        a, b = c["r_sub_fast_yoy_start"], c["r_sub_fast_yoy_end"]
        p["cp_yoy_move"] = ("increase from {v_sub_fast_yoy_start} in {est_year} to {v_sub_fast_yoy_end} by {fc_end}" if b - a >= 0.002 else
                            "decline from {v_sub_fast_yoy_start} in {est_year} to {v_sub_fast_yoy_end} by {fc_end}" if a - b >= 0.002 else
                            "remain stable at around {v_sub_fast_yoy_end} through {fc_end}")
    if "r_sub_lead_bps" in c:
        x = c["r_sub_lead_bps"]
        p["cp_lead_bps_move"] = ("gaining {v_sub_lead_bps} BPS" if x >= 10 else "losing {v_sub_lead_bps} BPS" if x <= -10
                                 else "maintaining its share")
    if "sub_lose" in c:
        p["cp_loser_line"] = ("The {sub_lose_lc} segment is expected to lose {v_sub_lose_bps} BPS over the forecast period to account for "
                              "{v_sub_lose_share_end} share by {fc_end}.")
    if "sub_attr1" in c:
        two = "sub_attr2" in c and not _same(c["sub_attr1"], c.get("sub_attr2"))
        p["cp_attr_headline"] = ("{sub_attr1_cap} followed by {sub_attr2_lc} are expected to be the two most attractive {segment_lc} segments "
                                 "in the {geo_ref} {mkt_short_lc} market in terms of value over the forecast period" if two else
                                 "{sub_attr1_cap} is expected to be the most attractive {segment_lc} segment in the {geo_ref} {mkt_short_lc} market "
                                 "in terms of value over the forecast period")
        p["cp_attr_line"] = ("{sub_attr1_cap} and {sub_attr2_lc} are expected to be the two most attractive {segment_lc} segments, exhibiting an "
                             "attractiveness index of {s_attr1} and {s_attr2} respectively." if two else
                             "{sub_attr1_cap} is expected to be the most attractive {segment_lc} segment, exhibiting an attractiveness index of {s_attr1}.")
        p["cp_attr_low_line"] = ("The {sub_attr_low_lc} segment is expected to be less attractive as compared to other segments over the forecast "
                                 "period, owing to its lower CAGR and incremental opportunity." if "sub_attr_low" in c else "")
    if "sub_lead" in c and "sub_second" in c:
        p["cp_table_headline"] = ("{sub_lead_cap} segment is projected to register {v_lead_vs_second}X higher value than {sub_second_lc} "
                                  "segment in {fc_end}")
    if "cty_lead" in c and "cty_second" in c:
        p["cp_country_headline"] = "{cty_lead_The} {mkt_short_lc} market is projected to register {v_cty_ratio}X higher value than {cty_second_the} in {fc_end}"
    elif "cty_lead" in c:
        p["cp_country_headline"] = "{cty_lead_The} is projected to dominate the {geo} {mkt_short_lc} market in terms of value throughout the forecast period"
    if "seg1_lead" in c:
        p["cp_seg1_table_headline"] = ("{seg1_lead_cap} segment is projected to dominate the {geo} {mkt_short_lc} market throughout the forecast period")
    if "cty_lead" in c and c.get("geo_level") == "region":
        p["cp_geo_country_line"] = "{mkt_short_lc_cap} market in {geo_the} is largely driven by demand from {cty_lead_the}, followed by {cty_second_the}." \
            if "cty_second" in c else "{mkt_short_lc_cap} market in {geo_the} is largely driven by demand from {cty_lead_the}."
    else:
        p["cp_geo_country_line"] = ""
    if "r_cr3" in c:
        cr = c["r_cr3"]
        word = ("consolidated" if cr >= 0.6 else "moderately consolidated" if cr >= 0.4 else "fragmented")
        p["cp_competition"] = (f"The global {{mkt_short_lc}} market is {word}, with the top three players accounting for an estimated "
                               f"{{v_cr3}} share of the market in {{base_year}}.")
    if "r_price_avg_cagr" in c:
        g = c["r_price_avg_cagr"]
        p["cp_price_dir"] = "increase to" if g > 0.005 else ("decline to" if g < -0.005 else "remain broadly stable at")
    return p


def item_classic(item, ctx):
    p = {}
    if "item" in item:
        p["item_lc"] = item.get("item_lc", _lower_name(item["item"]))
        p["item_cap"] = _cap(p["item_lc"])
        p["item_title"] = item.get("item_title", _title_name(item["item"]))
    if "r_item_mult" in item:
        p["cp_item_mult"] = mult_phrase(item["r_item_mult"], "v_item_mult")
    if item.get("is_lead"):
        p["cp_item_base_line"] = ("In {base_year}, {item_lc} segment was the most dominating segment in the {mkt_short_lc} market, accounting for "
                                  "around {v_item_share_base} share or US$ {v_item_base} {val_word}.")
    else:
        p["cp_item_base_line"] = ("In {base_year}, {item_lc} segment was valued at US$ {v_item_base} {val_word} and is estimated to account for "
                                  "US$ {v_item_est} {val_word} by the end of {est_year}, reflecting a Y-o-Y growth rate of {v_item_yoy_est}.")
    return p


# =============================================================================================
# Rendering with figure and table numbering
# =============================================================================================
from .engine import _render_value  # noqa: E402

SKIP = {"rule_note", "no_developments"}


def classic_ctx(ctx):
    full = build_ctx(ctx)
    full.setdefault("mkt_short", full.get("market"))
    full.setdefault("seg_tag", "Segment")
    full.setdefault("metric", "value")
    full.setdefault("metric_cap", "Value")
    full.setdefault("val_word", "Mn")
    full.setdefault("mkt_short_lc", _lower_name(full["mkt_short"]))
    full["mkt_short_lc_cap"] = _cap(full["mkt_short_lc"])
    full["market_lc_cap"] = _cap(full.get("market_lc", full["market"].lower()))
    if "industry" in full and "drivers" not in full:
        full.update(dynamics_ctx(full["industry"]))
    full.update({k: v for k, v in classic_derive(full).items() if k not in ctx})
    full.update({k: v for k, v in _articles_extra(full).items() if k not in full})
    return full


def _articles_extra(full):
    from .engine import THE_NAMES
    out = {}
    for k in ("geo", "cty_lead", "cty_second"):
        v = full.get(k)
        if isinstance(v, str):
            out[k + "_the"] = ("the " + v) if v in THE_NAMES else v
            out[k + "_The"] = ("The " + v) if v in THE_NAMES else v
    return out


def render_classic(cid, ctx, mask=True, counters=None):
    """Render one classic slide. counters = {'fig': n, 'tab': n} is advanced in place."""
    counters = counters if counters is not None else {"fig": 0, "tab": 0}
    t = CLASSIC_INDEX[cid]
    local = dict(ctx)
    local.update(item_classic(ctx, ctx))
    figs, tabs = [], []
    for i, cap in enumerate(t["figures"]):
        counters["fig"] += 1
        n = counters["fig"]
        local["fig_" + "ab"[i]] = f"{n:02d}"
        figs.append((n, render_text(cap, local, mask)))
    for i, cap in enumerate(t["tables"]):
        counters["tab"] += 1
        n = counters["tab"]
        local["tab_" + "ab"[i]] = f"{n:02d}"
        tabs.append((n, render_text(cap, local, mask)))
    out = {"id": cid, "name": t["name"]}
    errors = {}
    for key in ("title", "headline"):
        try:
            out[key] = render_text(t[key], local, mask)
        except MissingKeys as e:
            errors[key] = e.args[0]
    out["figure_captions"] = [f"Figure {n:02d}: {c}" for n, c in figs]
    out["table_captions"] = [f"Table {n:02d}: {c}" for n, c in tabs]
    metric = local.get("metric", "value")
    for key, val in t["elements"].items():
        if key in SKIP:
            continue
        if "__" in key:
            base, variant = key.split("__")
            if variant != metric:
                continue
            key = base
        try:
            r = _render_value(val, local, mask)
            if isinstance(r, list):
                r = [x for x in r if x]
            out[key] = r
        except MissingKeys as e:
            errors[key] = e.args[0]
    out["source"] = render_text(SOURCE, local, mask)
    out["footer"] = render_text(FOOTER, local, mask)
    if errors:
        out["_missing"] = errors
    return out, figs, tabs


def render_deck(plan, mask=True, start_fig=0, start_tab=0):
    """plan: list of (classic_id, ctx). Returns (slides, list_of_figures, list_of_tables)."""
    counters = {"fig": start_fig, "tab": start_tab}
    slides, lof, lot = [], [], []
    for cid, ctx in plan:
        s, f, t = render_classic(cid, ctx, mask, counters)
        slides.append(s)
        lof += [f"Figure {n}: {c}" for n, c in f]
        lot += [f"Table {n}: {c}" for n, c in t]
    return slides, lof, lot
