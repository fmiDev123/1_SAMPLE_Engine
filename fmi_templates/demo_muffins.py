"""
Worked example: the Packaged Muffins sample, rendered in masked sample mode (all v_* = XX).
Raw r_* numbers below are illustrative only; they drive word choice, never printed.

Run:  python -m fmi_templates.demo_muffins
"""
from .engine import build_ctx, render_slide, qc_slide, SLIDE_INDEX, BANNED_BY_INDUSTRY

REGIONS = ["North America", "Europe", "South Asia & Pacific", "East Asia", "Latin America", "Middle East & Africa"]

BASE = dict(
    industry="food", market="Packaged Muffins", market_lc="packaged muffins",
    publisher="Future Market Insights", publisher_short="FMI", source="FMI Analysis, 2026", pub_year=2026,
    report_date="07 Oct 2026", geo_top="Global", geo_top_ref="global",
    geo="Global", geo_ref="global", geo_ref_cap="Global", parent_geo="Global",
    focus_region="North America", focus_country="U.S.",
    hist_start=2021, hist_end=2025, est_year=2026, mid_year=2031, fc_end=2036,
    hist_period="2021-2025", fc_period="2026-2036", full_period="2021-2036", mid_period="2026-2031", yoy_period="2022-2036",
    val_unit="US$ Mn", vol_unit="MT", price_unit="US$/MT",
    page_tag="", n_top=10,
    child_label="region", child_label_pl="regions", child_label_cap="Region", n_children=6,
    n_regions=6, n_countries=36, n_segments=17, n_companies=30, n_named=10,
    seg1="Product Type", seg1_lc="product type", n_seg1_sub=2, seg1_sub1="Sweet muffins", seg1_sub2="Savoury muffins",
    seg1_lead="Sweet muffins", seg1_fast="Sweet muffins", seg1_slow="Savoury muffins",
    seg1_gain="Sweet muffins", seg1_lose="Savoury muffins", r_seg1_gain_bps=60, r_seg1_lead_share_end=0.78,
    seg2="Sales Channel", seg2_lc="sales channel", n_seg2_sub=3, seg2_lead="Retail",
    seg2_fast="Foodservice and institutional", r_seg2_fast_cagr=0.061, n_seg2_panel=4,
    child_lead="North America", child_fast="South Asia & Pacific", child_slow="North America",
    child_gain="South Asia & Pacific", child_lose="North America", n_above_pace=3, n_below_pace=3,
    r_child_lead_share_end=0.302,
    r_f_cagr=0.050, r_h_cagr=0.041, r_vol_f_cagr=0.038, r_vol_h_cagr=0.035,
    r_vol_abs_fc=520, r_vol_abs_hist=210, r_abs_opp=2400, r_abs_hist=900,
    cty_cagr_top="India", cty_abs_top="U.S.", cty_white="Netherlands", n_white=7, cty_lead="U.S.",
    price_seg_top="Sweet muffins", price_seg_low="Savoury muffins", price_seg_fast="Sweet muffins", price_seg_slow="Savoury muffins",
    price_vol_high="Sweet muffins", price_vol_low="Savoury muffins",
    child_price_high="Europe", child_price_low="North America", child_price_fast="South Asia & Pacific",
    child_mix_wide="North America", child_margin_high="Europe", n_child_above_price=4, n_child_below_price=2,
    r_child_price_spread=0.18, r_price_conv=True, r_price_avg_cagr=0.012, r_price_gap_start=850, r_price_gap_end=990,
    drv1="Input costs", drv1_lc="input costs", drv2_lc="demand conditions", drv3="Competitive intensity",
    drv4_lc="logistics costs", drv_last="Trade promotion spend", s_drv1=8.6, s_drv2=7.9,
    down_factors=["Process Innovation", "Competitive Intensity"], n_up_factors=8, n_factors=10,
    vc_top_stage="Manufacturers",
    sf1="Geographic Spread", sf2="Capacity Addition", sf3="Input Cost Pressure", s_sf1=8.1, s_sf2=7.4, s_sf3=6.9,
    df1="Adoption Depth", df2="Volume Momentum", df3="Premiumisation", s_df1=7.8, s_df2=7.1, s_df3=6.6,
    tl1="2021-2022", tl2="2023-2025", tl3="2026-2028", tl4="2029-2036",
    top3_children="North America, Europe and South Asia & Pacific",
    s_force_suppliers=6.0, s_force_buyers=5.0, s_force_rivalry=6.5, s_force_substitutes=5.0, s_force_entrants=4.0,
    r_cr3=0.42, r_t1_share=0.55, r_hhi=1350, n_ma=5,
    n_leaders=2, n_tracked_subs=5, n_max_leads=3, n_t1=4, n_t2=2, n_t3=4, n_global=4, n_regional=2, n_local=4,
    rank1="Grupo Bimbo", rank2="McKee Foods Corporation", rank3="Rich Products Corporation",
    most_contested="Sweet muffins", n_most_contested=8, most_open="Vending", n_most_open=6,
    enabler_heads=["Sweet muffins Compounds Fastest Across the Product Type Mix", "South Asia & Pacific Emerges as the Growth Engine"],
    constraint_heads=["Savoury muffins Lags the Category", "Uneven Region Performance Caps the Blended Rate"],
    tier_from="Others", tier_to="Tier 2",
    top_type_name="M&A", n_top_type=5, n_moves=14, top_type_function="corporate development",
    comp1="Grupo Bimbo", comp2="McKee Foods Corporation", n_shown=2, leaders_same=False, item_leader_all="Grupo Bimbo",
    sub_lead_peak_child="North America",
)

SEG_CHAPTER = dict(
    segment="Product Type", segment_lc="product type", n_sub=2,
    sub_lead="Sweet muffins", sub_fast="Sweet muffins", sub_gain="Sweet muffins", sub_lose="Savoury muffins", r_sub_gain_bps=60,
    sub_lead_vol="Sweet muffins", sub_fast_vol="Sweet muffins", sub_gain_vol="Sweet muffins", sub_lose_vol="Savoury muffins",
    r_sub_vol_gain_bps=45, r_seg_cagr=0.050, sub_opp_low="Savoury muffins",
    sub_yoy_peak="Sweet muffins", sub_yoy_peak_year=2027, n_sub_decline=0, yoy_peak_year=2027,
    r_yoy_early=0.055, r_yoy_late=0.046, metric="value", metric_cap="Value", m_unit="US$ Mn",
    subs=[dict(item="Sweet muffins", is_lead=True, is_fast=True, r_item_share=0.77, r_item_cagr=0.052),
          dict(item="Savoury muffins", r_item_share=0.23, r_item_cagr=0.045)],
    subs_on_page=["Sweet muffins", "Savoury muffins"], subs_top3_opp=["Sweet muffins"],
)

CHILD_ITEMS = [dict(item=r, rank=i + 1, is_lead=(i == 0), is_fast=(r == "South Asia & Pacific"),
                    r_item_share=s, r_item_cagr=g, r_item_h_cagr=h, r_item_price_cagr=0.012, item_seg1_lead="Sweet muffins",
                    item_seg1_fast="Sweet muffins")
               for i, (r, s, g, h) in enumerate(zip(REGIONS, [0.31, 0.24, 0.17, 0.15, 0.08, 0.05],
                                                     [0.041, 0.052, 0.068, 0.047, 0.055, 0.058],
                                                     [0.036, 0.044, 0.051, 0.045, 0.049, 0.050]))]

BREAKDOWN = dict(children=CHILD_ITEMS, child_lead_m="North America", child_fast_m="South Asia & Pacific",
                 child_gain_m="Europe", child_lose_m="North America", r_child_gain_m_bps=80,
                 metric="volume", metric_cap="Volume", m_unit="MT")

CROSS = dict(seg_a="Product Type", seg_a_lc="product type", seg_b="Flavour", seg_b_lc="flavour",
             xa_lead="Sweet muffins", xb_lead_in_a="Fruit-based", x_lead_cell="Sweet muffins × Fruit-based",
             x_fast_cell="Savoury muffins × Cheese-based", x_top_share_a="Sweet muffins", x_top_share_b="Fruit-based",
             x_conc_a="Savoury muffins", x_conc_b="Cheese-based", x_div_a="Sweet muffins", x_gain_b="Cheese-based",
             x_opp_lead="Sweet muffins × Fruit-based", x_idx_top="Sweet muffins × Fruit-based", s_x_idx_top=5,
             x_idx_low="Savoury muffins × Herb-based", s_x_idx_low=1,
             x_top_cells=[dict(item="Sweet muffins × Fruit-based", s_item_idx=5), dict(item="Sweet muffins × Chocolate-based", s_item_idx=5),
                          dict(item="Savoury muffins × Cheese-based", s_item_idx=4)],
             pm_top_a="Sweet muffins", pm_top_b="Nut and seed-based", pm_low_a="Savoury muffins", pm_low_b="Vegetable-based",
             pm_spread_a="Sweet muffins")

COUNTRY = dict(geo="U.S.", geo_ref="U.S.", geo_ref_cap="U.S.", parent_geo="North America", r_m_f_cagr=0.039, r_parent_cagr=0.041,
               metric="value", metric_cap="Value", m_unit="US$ Mn",
               segments_on_page=["Product Type", "Flavour", "Primary Flour Base"], sub_opp_lead="Sweet muffins")

COMPANY = dict(company="Grupo Bimbo", n_prof=5, c_y1=2021, c_y2=2022, c_y3=2023, c_y4=2024, c_y5=2025,
               c_currency_unit="US$ Mn", c_fy_end="December", c_seg2_lead="Retail", c_geo_lead="North America",
               c_geo_pen_top="Latin America", c_geo_pen_low="East Asia", c_revenue="[from annual report]",
               c_ebitda_margin="[from annual report]", c_fin_source="Grupo Bimbo Annual Report 2025",
               c_key_brands="Bimbo, Marinela, Little Bites, Thomas'", c_seg1_lead="Sweet muffins",
               c_plants="[from annual report]", c_countries="[from annual report]", c_employees="[from annual report]",
               c_is_leader=True, r_c_rev_cagr=0.04, c_rev_cagr="[calc]", c_nm_start="[calc]", c_nm_end="[calc]",
               c_strength_1="[Sourced strength, e.g. scale of the North American direct-store-delivery network.]",
               c_weakness_1="[Sourced weakness from annual report or analyst coverage.]",
               c_stated_strategy="[Management-stated priority from the latest annual report, written as a full sentence.]",
               company_developments=[dict(c_dev_date="[Month YYYY]", c_dev_text="[sourced development]", c_dev_source="[URL]")])

OVERLAY = {
    "A03": dict(section_title="Market Signals", sec_no=1, sub_sections=["Market Signals"]),
    "A04": dict(part_title="Market Overview", part_sections=["Market Taxonomy", "Table of Contents", "List of Figures, List of Tables"]),
    "A02": dict(sec_signals=1, sec_taxonomy=2, sec_toc=3, sec_lof=4, sec_geo_overview=5, sec_geo_region=6, sec_geo_segment=7,
                sec_cross=8, sec_reg_overview=9, sec_reg_country=10, sec_reg_segment=11, sec_cty_overview=12, sec_cty_segment=13,
                sec_pricing=14, sec_supply=15, sec_dynamics=16, sec_comp=17, sec_share=18, sec_profiles=19),
    "A06": dict(item="North America"), "A07": dict(item="Product Type", page_tag=" (1/2)"),
    "A08": dict(companies=[dict(item="Grupo Bimbo", brands="Bimbo, Marinela, Little Bites, Thomas'"),
                           dict(item="McKee Foods Corporation", brands="Little Debbie, Sunbelt Bakery, Drake's")]),
    "A09": dict(page_tag=" (1/12)"), "A10": dict(page_tag=" (1/4)"), "A11": dict(page_tag=" (1/4)"),
    "G05": dict(page_tag=" (1/2)"),
    "M01": dict(children=CHILD_ITEMS),
    "K13": dict(c_currency_unit="US$ Mn", c_key_brands="[brands per company from the coverage list]"),
    "P06": dict(drivers_top5=[dict(item="Input Costs", s_item=8.6), dict(item="Demand Conditions", s_item=7.9),
                              dict(item="Competitive Intensity", s_item=7.2), dict(item="Logistics Costs", s_item=6.5),
                              dict(item="Labour Costs", s_item=5.8)]),
    "K03": dict(seg1_subs=[dict(item="Sweet muffins", n_item_active=5), dict(item="Savoury muffins", n_item_active=5)]),
    "K04": dict(tiers=[dict(item="Tier 1", item_strategy="Acquisition-led scale-up"), dict(item="Tier 2", item_strategy="Launch-led growth"),
                       dict(item="Tier 3", item_strategy="Geographic expansion")]),
    "K08": dict(tiers=[dict(item="Tier 1", item_members="Grupo Bimbo, McKee Foods Corporation", item_best="scale and share", item_worst="segment breadth"),
                       dict(item="Tier 2", item_members="Europastry, Aspire Bakeries", item_best="strategic momentum", item_worst="segment breadth")]),
    "K09": dict(developments=[dict(c_dev_company="[Company]", c_dev_text="[sourced development]", c_dev_date="[Month YYYY]",
                                   c_dev_type="MA", c_dev_source="[press release URL]")]),
    "K12": dict(tiers=[dict(item="Tier 1", n_item_up=1, n_item_down=0), dict(item="Tier 2", n_item_up=1, n_item_down=0)]),
    "L03": dict(share_items=[dict(item="Sweet muffins", item_comp1="Grupo Bimbo", item_comp2="Rich Products Corporation"),
                             dict(item="Savoury muffins", item_comp1="Hostess", item_comp2="McKee Foods Corporation")]),
    "L04": dict(share_items=[dict(item="North America", item_comp1="Grupo Bimbo", item_comp2="McKee Foods Corporation"),
                             dict(item="Europe", item_comp1="Grupo Bimbo", item_comp2="Rich Products Corporation")], n_shown=3),
    "L05": dict(growth_shapers=[dict(item="Grupo Bimbo", item_role="Market leader", item_why="Expanded reach across its core segment",
                                     item_move="New launch")]),
}

SECTION_OVERLAY = {"Segment chapter": SEG_CHAPTER, "Geo breakdown": BREAKDOWN, "Cross-section": CROSS,
                   "Country": {**SEG_CHAPTER, **COUNTRY}, "Company profile": COMPANY}


def ctx_for(sid):
    s = SLIDE_INDEX[sid]
    c = dict(BASE)
    if s["id"] in ("P03",):
        c.update(CROSS)
    c.update(SECTION_OVERLAY.get(s["section"], {}))
    c.update(OVERLAY.get(sid, {}))
    return build_ctx(c)


def render_all(mask=True):
    results = []
    for sid in SLIDE_INDEX:
        r = render_slide(sid, ctx_for(sid), mask=mask)
        results.append((r, qc_slide(r, mask, BANNED_BY_INDUSTRY["food"])))
    return results


if __name__ == "__main__":
    total = 0
    for r, issues in render_all():
        if issues:
            total += len(issues)
            print(r["id"], r["name"])
            for i in issues:
                print("   ", i)
    print("issues:", total)
