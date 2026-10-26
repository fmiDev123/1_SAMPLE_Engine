"""
Classic-style demos.

1) DWM check: re-generates pages of the 2015 report (REP-GB-717) from its own numbers, unmasked,
   and prints the original text beside the generated text.
2) Muffins: renders the full classic plan for the Packaged Muffins sample in masked mode.

Run:  python -m fmi_templates.demo_classic
"""
from .classic import classic_ctx, render_deck, render_classic, CLASSIC
from .engine import qc_slide, BANNED_BY_INDUSTRY

DWM = dict(
    industry="healthcare", market="Digital Wound Measurement Devices", market_lc="digital wound measurement devices",
    mkt_short="DWM Devices", publisher="Future Market Insights", publisher_short="FMI", source="FMI, 2015", pub_year=2015,
    geo="Global", geo_ref="global", geo_ref_cap="Global", geo_top="Global",
    hist_start=2014, base_year=2014, est_year=2015, mid_year=2020, fc_end=2025,
    val_unit="US$ '000", val_word="thousand", seg_tag="Wound Type",
    segments=["Wound Type", "End-users"], seg1="Wound Type", seg2="End-users",
    child_label="region", child_label_pl="regions", child_lead="North America", child_fast="North America",
    r_mult=1.65, r_f_cagr=0.047, r_h_cagr=0.04,
)

DWM_CASES = [
    ("R14", dict(v_mult="1.6", v_val_end="2,479.4", v_abs_opp="979.4"),
     "Digital wound measurement devices market is projected to register 1.6X increase in its revenue by the end of forecast period, 2025 | "
     "Figure 05 highlights overall value of the digital wound measurement devices market from 2014 to 2025 | Global digital wound measurement "
     "devices market is expected to be valued at US$ 2.5 Mn by 2025 end | It is projected to create overall incremental opportunity of US$ 1 Mn from 2014 to 2025"),
    ("R21", dict(segment="Wound Type", sub_lead="Diabetic Ulcers", sub_fast="Burns", sub_lose="Incisional Wounds", v_sub_fast_cagr="5.5%",
                 r_sub_lead_bps=49, v_sub_lead_bps="49", v_sub_lead_share_end="44.1%", v_sub_lose_bps="72", v_sub_lose_share_end="14.1%"),
     "Among DWM devices by wound type, burns segment is expected to witness highest CAGR of 5.5% over the forecast period | In terms of value share by "
     "wound type, the diabetic ulcers segment is expected to dominate the market, gaining 49 BPS and accounting for 44.1% share over the forecast period | "
     "The incisional wounds segment is expected to lose 72 BPS over the forecast period to account for 14.1% share by 2025"),
    ("R22", dict(segment="Wound Type", sub_attr1="Diabetic Ulcer", sub_attr2="Burns", sub_attr_low="Incisional Wounds", s_attr1="3.7", s_attr2="1.8"),
     "Diabetic ulcers followed by chronic wounds are expected to be the two most attractive wound type segments in the global DWM devices market in terms "
     "of value over the forecast period | Diabetic ulcer and burns are expected to be the two most attractive wound type segments, exhibiting an attractiveness index of 3.7 and 1.8 respectively"),
    ("R23", dict(segment="Wound Type", item="Diabetic Ulcer", is_lead=True, r_item_mult=1.69, v_item_mult="1.7", v_item_share_base="54%",
                 v_item_base="645.9", v_item_end="1,092.1", v_item_cagr="5.0%", v_item_opp="446.2", v_item_abs_est="27.3", metric="value",
                 metric_cap="Value", m_unit="US$ '000"),
     "Diabetic ulcer wound type segment is projected to increase 1.7X by the end of 2025 as compared to that in 2014 | Figures 09 and 10 highlight diabetic ulcer "
     "segment revenue contribution and overall incremental opportunity from 2014 to 2025 | In 2014, diabetic ulcer segment was the most dominating segment in the DWM "
     "devices market, accounting for around 54% share or US$ 645.9 thousand | The diabetic ulcer segment is expected to account for US$ 1,092.1 thousand by 2025 at a CAGR "
     "of 5% over the forecast period | This segment is expected to create an incremental opportunity of US$ 446.2 thousand for the overall DWM devices market from 2014 to 2025 | "
     "Absolute $ value for 2015 is estimated to be US$ 27.3 thousand."),
    ("R23", dict(segment="Wound Type", item="Chronic Wounds", r_item_mult=1.66, v_item_mult="1.7", v_item_base="308.2", v_item_est="320.9",
                 v_item_yoy_est="4.1%", v_item_end="513.1", v_item_cagr="4.8%", v_item_opp="204.0", v_item_abs_est="12.6", metric="value",
                 metric_cap="Value", m_unit="US$ '000"),
     "By 2025, chronic wounds segment is anticipated to witness an increase of 1.7X in terms of value as compared to 2014 | In 2014, the chronic wounds segment was "
     "valued at US$ 308.2 thousand, which is estimated to increase to US$ 320.9 thousand by 2015 end, reflecting a Y-o-Y growth rate of 4.1% | The chronic wounds segment "
     "is expected to account for US$ 513.1 thousand by 2025 at a CAGR of 4.8% over the forecast period | This segment is expected to create incremental opportunity of US$ 204 thousand"),
    ("R24", dict(segment="Wound Type", sub_lead="Diabetic Ulcer", sub_second="Chronic Wounds", v_lead_vs_second="2.1", metric_cap="Value", m_unit="US$ '000"),
     "Diabetic ulcer wound type segment is projected to register 2.1X higher value than chronic wounds segment in 2025 | "
     "Table 01: Global DWM Devices Market Value Forecast By Wound Type, 2014-2025 (US$ 000)"),
    ("R26", dict(geo="North America", geo_level="region", r_mult=1.73, v_mult="1.7", v_val_base="653.7", v_val_est="681.8", v_yoy_est="4.3%",
                 v_val_end="1,129.7", v_f_cagr="5.2%", v_abs_opp="476.0", v_abs_est="28.1", cty_lead="U.S.", cty_second="Canada"),
     "DWM devices market value in North America is expected to increase 1.7X between 2014 and 2025 | DWM devices market value in North America was US$ 653.7 thousand "
     "in 2014 and is estimated to increase to US$ 1,129.7 thousand by 2025 at a CAGR of 5.2% | In terms of revenue, North America is forecast to represents an overall "
     "incremental opportunity of US$ 1,129.8 thousand by the end of 2025 [original error: this is the 2025 value, not the opportunity] | Absolute $ value for 2015 is estimated to be US$ 28.1 thousand."),
    ("R28", dict(geo="Latin America", cty_lead="Brazil", cty_second="Mexico", v_cty_ratio="2.1"),
     "Brazil DWM devices market is projected to register 2.1X higher value than Mexico in 2025 | Table 07: Latin America DWM Devices Market Value Forecast By Country, 2014-2025 | "
     "Table 08: Latin America DWM Devices Market Value Forecast By End-users, 2014-2025"),
    ("R34", dict(company="Wound Zoom Inc.", c_products=[dict(item="Wound Zoom Camera", item_text="digital, non-contact wound documentation camera system")],
                 c_strategy="The company emphasises the development of innovative products using advanced technologies in order to adhere to industry standards.",
                 c_developments=[dict(item_year=2015, item_text="the company signed an agreement with Imago Care Ltd. to enhance its advanced wound care strategic marketing and technology consultancy."),
                                 dict(item_year=2014, item_text="the company introduced Wound Zoom 2.0, a digital wound measurement device.")],
                 c_dev_window="2014-2015"),
     "PRODUCT OFFERING: Wound Zoom Camera - Digital, non-contact wound documentation camera system | STRATEGY: Company emphasizes on the development of innovative products "
     "using advanced technologies in order to adhere to industry standards | In early 2015, the company signed an agreement with Imago Care Ltd. ... | In 2014, the company introduced digital wound measurement device namely Wound Zoom 2.0."),
]

START_FIG = {"R14": 4, "R21": 6, "R22": 7, "R23": 8, "R26": 28, "R28": 0, "R24": 0, "R34": 0}
START_TAB = {"R28": 6, "R24": 0}


def flat(s):
    parts = []
    for k, v in s.items():
        if k in ("id", "name", "source", "footer", "_missing"):
            continue
        if isinstance(v, list):
            parts += [str(x) for x in v]
        elif v:
            parts.append(str(v))
    return " | ".join(parts)


def dwm_check():
    rows = []
    seen23 = 0
    for cid, extra, original in DWM_CASES:
        ctx = dict(DWM)
        ctx.update(extra)
        start = START_FIG.get(cid, 0)
        if cid == "R23":
            start += 2 * seen23
            seen23 += 1
        s, _, _ = render_classic(cid, classic_ctx(ctx), mask=False, counters={"fig": start, "tab": START_TAB.get(cid, 0)})
        rows.append((cid, s["name"], original, flat(s), s.get("_missing")))
    return rows


MUFFIN = dict(
    industry="food", market="Packaged Muffins", market_lc="packaged muffins", mkt_short="Packaged Muffins",
    publisher="Future Market Insights", publisher_short="FMI", source="FMI, 2026", pub_year=2026, report_code="REP-GB-XXXX",
    geo="Global", geo_ref="global", geo_ref_cap="Global", geo_top="Global", page_tag="",
    hist_start=2021, base_year=2025, est_year=2026, mid_year=2031, fc_end=2036,
    val_unit="US$ Mn", val_word="Mn", vol_unit="MT", vol_unit_single="metric tonne", price_unit="US$/MT",
    segments=["Product Type", "Flavour", "Sales Channel"], seg1="Product Type", seg2="Sales Channel", seg_tag="Segment",
    regions=["North America", "Latin America", "Europe", "East Asia", "South Asia & Pacific", "Middle East & Africa"],
    companies=["Grupo Bimbo", "McKee Foods Corporation"],
    child_label="region", child_label_pl="regions", child_lead="North America", child_fast="South Asia & Pacific",
    child_price_high="Europe", child_price_low="North America",
    r_mult=1.63, r_f_cagr=0.050, r_h_cagr=0.041, r_vol_f_cagr=0.038, r_cr3=0.42, r_price_avg_cagr=0.012,
    comp1="Grupo Bimbo", comp2="McKee Foods Corporation", seg1_lead="Sweet muffins", price_seg_top="Sweet muffins",
    parent_market="Packaged Bakery Products", parent_market_lc="packaged bakery products", parent_unit="US$ Bn",
    attr_fig_title="Global Packaged Bakery Products Market (US$ Bn), 2025 A and 2036 F",
    attr_points=["Global packaged bakery products market is expected to witness XXX increase in its revenue by 2036 end."],
    in_definition="[Market definition from the research scope document.]", in_unit_definition="[unit, e.g. one metric tonne of packaged muffins]",
    in_workflow_title="[Value chain or product workflow title]", in_sector_sources="[Sector sources, e.g. USDA, Eurostat, trade associations]",
    mkt_short_acronym_line="[Market acronym, if used]",
    section_title="Executive Summary", sec_no="01",
    xa_lead="Sweet muffins", xb_lead_in_a="Fruit-based", seg_a="Product Type", seg_b="Flavour",
)

SEG = dict(segment="Product Type", sub_lead="Sweet muffins", sub_fast="Sweet muffins", sub_second="Savoury muffins", sub_lose="Savoury muffins",
           sub_attr1="Sweet muffins", sub_attr2="Savoury muffins", s_attr1="3.1", s_attr2="1.4", r_sub_lead_bps=60, r_sub_fast_yoy_start=0.055, r_sub_fast_yoy_end=0.046,
           metric="value", metric_cap="Value", m_unit="US$ Mn")


def muffin_plan():
    plan = [("R01", {}), ("R05", {}), ("R06", {}), ("R07", {}), ("R08", {}), ("R09", {}), ("R10", {}), ("R11", {}), ("R12", {}),
            ("R13", {}), ("R14", {}), ("R14V", {}), ("R15", {}), ("R16", {}), ("R17", {}), ("R18", {"impact_items": [
                "Rising Demand for Convenient Food", "Product Innovation in Flavours and Formats", "Volatility in Raw Material Prices"]}),
            ("R19", {}), ("R20", SEG), ("R21", SEG), ("R22", SEG),
            ("R23", {**SEG, "item": "Sweet muffins", "is_lead": True, "r_item_mult": 1.65}),
            ("R23", {**SEG, "item": "Savoury muffins", "r_item_mult": 1.55}),
            ("R23", {**SEG, "item": "Savoury muffins", "r_item_mult": 1.55, "metric": "volume", "metric_cap": "Volume", "m_unit": "MT"}),
            ("R24", SEG), ("R25", {"geo": "North America"}),
            ("R26", {"geo": "North America", "geo_level": "region", "cty_lead": "U.S.", "cty_second": "Canada", "r_mult": 1.55}),
            ("R27", {"geo": "North America", "impact_items": [dict(item="Rising Demand for Convenient Food", item_level="High")]}),
            ("R28", {"geo": "North America", "cty_lead": "U.S.", "cty_second": "Canada"}),
            ("R29", {"geo": "North America"}), ("R26", {"geo": "U.S.", "geo_level": "country", "r_mult": 1.5}),
            ("R30", {}), ("R31", {}), ("R32", {}),
            ("R33", {"company": "Grupo Bimbo", "c_established": "[year]", "c_hq": "Mexico City, Mexico", "c_ceo": "[name (title)]",
                     "c_fy": 2025, "c_overview": ["[Sourced overview line 1.]", "[Sourced overview line 2.]"], "c_source": "Company Annual Report"}),
            ("R34", {"company": "Grupo Bimbo", "c_products": [dict(item="[Brand]", item_text="[sourced product description]")],
                     "c_strategy": "[Sourced strategy statement.]", "c_developments": [dict(item_year="[YYYY]", item_text="[sourced development].")],
                     "c_dev_window": "2024-2026"}),
            ("R02", {}), ("R03", {}), ("R04", {})]
    out = []
    for cid, extra in plan:
        c = dict(MUFFIN)
        c.update(extra)
        out.append((cid, classic_ctx(c)))
    return out


def muffin_render():
    slides, lof, lot = render_deck(muffin_plan(), mask=True)
    issues = []
    for s in slides:
        for i in qc_slide(s, True, BANNED_BY_INDUSTRY["food"]):
            issues.append((s["id"], i))
    return slides, lof, lot, issues


if __name__ == "__main__":
    for cid, name, orig, gen, miss in dwm_check():
        print(f"\n### {cid} {name}\nORIGINAL : {orig}\nGENERATED: {gen}")
        if miss:
            print("MISSING:", miss)
    slides, lof, lot, issues = muffin_render()
    print("\nmuffin slides:", len(slides), "figures:", len(lof), "tables:", len(lot), "issues:", issues[:20], len(issues))
