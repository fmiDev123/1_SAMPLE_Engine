"""
Phrase rules: deterministic wording chosen from the numbers, so no AI is needed to correct text.

Inputs use these conventions:
  r_*  raw numbers (CAGR and shares as fractions, e.g. 0.052; BPS as plain numbers, e.g. 120)
  n_*  integer counts
  names (seg1_lead, child_fast ...) as plain strings
Every rule returns text that may itself contain {v_*} placeholders, so masking still applies.

Thresholds are set once in THRESHOLDS so the team can tune them without touching the text.
"""

THRESHOLDS = {
    "trend_bps": 50,          # F-CAGR vs H-CAGR gap that counts as accelerating / moderating
    "bench_bps": 25,          # gap vs benchmark that counts as outpacing / trailing
    "price_mix_bps": 25,      # value vs volume CAGR gap that counts as price/mix uplift or erosion
    "share_shift_bps": 10,    # below this the mix is called stable
    "role_bps": 25,           # sub-segment CAGR vs segment CAGR for above / below average
    "niche_share": 0.10,      # below this share a sub-segment is a niche pool
    "core_share": 0.25,       # at or above this share a sub-segment is a core pool
    "cr3_high": 0.60, "cr3_mid": 0.40, "cr3_low": 0.20,
    "supply_conc_high": 0.40, "supply_conc_mid": 0.25,
    "price_band_tight": 0.15, "price_band_wide": 0.40,
    "price_flat": 0.005, "price_modest": 0.01, "price_strong": 0.03,
    "spread_change": 0.05,    # relative change in price gap that counts as widening / narrowing
    "yoy_bps": 50,
}

T = THRESHOLDS


def _bps(a, b):
    return (a - b) * 10000


def _same(a, b):
    return (a or "").strip().lower() == (b or "").strip().lower()


# ------------------------------------------------------------------------------- growth
def trend_phrase(f, h):
    d = _bps(f, h)
    if d >= T["trend_bps"]:
        return "accelerating from"
    if d <= -T["trend_bps"]:
        return "moderating from"
    return "broadly in line with"


def vs_benchmark(x, bench, bench_label, v_bench, v_gap):
    d = _bps(x, bench)
    if d >= T["bench_bps"]:
        return f"outpacing the {bench_label} benchmark of {{{v_bench}}} by {{{v_gap}}} BPS"
    if d <= -T["bench_bps"]:
        return f"trailing the {bench_label} benchmark of {{{v_bench}}} by {{{v_gap}}} BPS"
    return f"in line with the {bench_label} benchmark of {{{v_bench}}}"


def bench_arrow(x, bench):
    d = _bps(x, bench)
    if d >= T["bench_bps"]:
        return "▲ above benchmark"
    if d <= -T["bench_bps"]:
        return "▼ below benchmark"
    return "▶ in line with benchmark"


def period_compare(fc_abs, hist_abs):
    if hist_abs <= 0:
        return "against"
    ratio = fc_abs / hist_abs
    if ratio >= 1.1:
        return "exceeding"
    if ratio <= 0.9:
        return "below"
    return "broadly matching"


def price_mix(val_cagr, vol_cagr):
    """Returns (full sentence, short clause)."""
    d = _bps(val_cagr, vol_cagr)
    if d >= T["price_mix_bps"]:
        return ("Value grows {v_val_vol_gap_bps} BPS of CAGR faster than volume, which indicates positive price and mix realisation.",
                "which indicates a premiumising market")
    if d <= -T["price_mix_bps"]:
        return ("Value grows {v_val_vol_gap_bps} BPS of CAGR slower than volume, which indicates price erosion or mix dilution.",
                "which indicates price erosion or mix dilution")
    return ("Value and volume grow broadly in step, which indicates stable pricing.",
            "which indicates broadly stable pricing")


# ------------------------------------------------------------------------------- leader / fastest
def lead_fast(lead, fast, scope, lead_share_key, fast_cagr_key, lead_cagr_key=None, quote=True):
    q = "'" if quote else ""
    if _same(lead, fast):
        return (f"{q}{lead}{q} both leads the {scope} and grows fastest at {{{fast_cagr_key}}} CAGR, "
                f"which reinforces its position through {{fc_end}}.")
    return (f"{q}{lead}{q} leads the {scope} with {{{lead_share_key}}} of value, while {q}{fast}{q} grows fastest "
            f"at {{{fast_cagr_key}}} CAGR.")


def seg1_fast_clause(lead, fast):
    if _same(lead, fast):
        return ", and it is also the fastest-growing category"
    return f", while '{fast}' grows fastest at {{v_seg1_fast_cagr}} CAGR"


def child_lead_fast(lead, fast, child_label):
    if _same(lead, fast):
        return (f"{lead} is both the largest and the fastest-growing {child_label}, holding {{v_child_lead_share_end}} "
                f"of {{fc_end}} value and growing at {{v_child_fast_cagr}} CAGR.")
    return (f"{lead} is the largest {child_label} with {{v_child_lead_share_end}} of {{fc_end}} value, while {fast} "
            f"grows fastest at {{v_child_fast_cagr}} CAGR.")


def child_lead_fast_short(lead, fast, child_label):
    if _same(lead, fast):
        return f"{lead} is also the fastest-growing {child_label}, at {{v_child_fast_cagr}} CAGR through {{fc_end}}."
    return f"{fast} is the fastest-growing {child_label}, at {{v_child_fast_cagr}} CAGR through {{fc_end}}."


# ------------------------------------------------------------------------------- share shift
def share_shift(gain, lose, gain_bps, scope, v_gain="v_sub_gain_bps", v_lose="v_sub_lose_bps", quote=True):
    q = "'" if quote else ""
    if abs(gain_bps) < T["share_shift_bps"]:
        return f"The {scope} mix remains broadly stable, with no sub-segment moving more than {T['share_shift_bps']} BPS between {{est_year}} and {{fc_end}}."
    return (f"Share shifts toward {q}{gain}{q} (+{{{v_gain}}} BPS) and away from {q}{lose}{q} (-{{{v_lose}}} BPS) "
            f"between {{est_year}} and {{fc_end}}.")


# ------------------------------------------------------------------------------- roles in repeat lists
def item_role(is_lead, is_fast, share, cagr, ref_cagr, scope="segment", geo_mode=False):
    if is_lead and is_fast:
        return f"it leads the {scope} and is also the fastest-growing" + (" market" if geo_mode else " sub-segment")
    if is_lead:
        return f"it leads the {scope} by value"
    if is_fast:
        return "it is the fastest-growing" + (" market" if geo_mode else " sub-segment")
    size = ("a core pool" if share >= T["core_share"] else
            "a mid-sized pool" if share >= T["niche_share"] else "a niche pool")
    d = _bps(cagr, ref_cagr)
    ref = "the market average" if geo_mode else "the segment average"
    if d >= T["role_bps"]:
        pace = f"growing above {ref}"
    elif d <= -T["role_bps"]:
        pace = f"growing below {ref}"
    else:
        pace = f"tracking {ref}"
    return f"it is {size} {pace}"


def driver_role(rank):
    return {1: "the dominant pressure", 2: "a major pressure", 3: "a key pricing lever",
            4: "a key pricing lever"}.get(rank, "a supporting factor")


# ------------------------------------------------------------------------------- concentration
def concentration(cr3):
    if cr3 >= T["cr3_high"]:
        return "highly consolidated"
    if cr3 >= T["cr3_mid"]:
        return "moderately consolidated"
    if cr3 >= T["cr3_low"]:
        return "fragmented, with a defined leadership group"
    return "highly fragmented"


def supply_conc(lead_share):
    if lead_share >= T["supply_conc_high"]:
        return "concentrated"
    if lead_share >= T["supply_conc_mid"]:
        return "moderately concentrated"
    return "dispersed"


def hhi_band(hhi):
    return "fragmented" if hhi < 1500 else ("moderately concentrated" if hhi <= 2500 else "concentrated")


def above_pace(n_above, n_total, label_pl, geo):
    if n_above == 0:
        return f"No {label_pl[:-1] if label_pl.endswith('s') else label_pl} grows faster than the {geo} average, which indicates that mix shifts rather than geographic outperformance drive growth."
    if n_above == n_total:
        return f"All {n_total} {label_pl} grow at or above the {geo} average."
    return f"{n_above} of {n_total} {label_pl} grow at or above the {geo} average, so the geographic mix shifts toward the faster movers."


def child_landscape(n, label_pl):
    if n <= 3:
        return f"{n} {label_pl} make up the geographic landscape."
    return f"{n} {label_pl} make up the geographic landscape, and the top three account for {{v_child_top3_share}} of {{est_year}} value."


# ------------------------------------------------------------------------------- pricing
def price_band(spread):
    if spread < T["price_band_tight"]:
        return "sit within a tight band across {child_label_pl}"
    if spread < T["price_band_wide"]:
        return "vary moderately across {child_label_pl}"
    return "vary widely across {child_label_pl}"


def price_conv(low_priced_grow_faster):
    if low_priced_grow_faster:
        return "lower-priced {child_label_pl} converge toward the average, reflecting catch-up growth"
    return "the premium held by higher-priced {child_label_pl} widens, reflecting mix upgrading in mature markets"


def price_direction(cagr):
    """Returns (present-tense verb, infinitive, outlook sentence)."""
    if cagr > T["price_flat"]:
        word = ("modest" if cagr < T["price_modest"] else "moderate" if cagr < T["price_strong"] else "strong")
        return ("rises", "rise",
                f"The net outlook is {word} structural price appreciation, with product mix and geographic mix determining realised pricing.")
    if cagr < -T["price_flat"]:
        return ("declines", "decline",
                "The net outlook is gradual price erosion, so volume growth and cost efficiency must carry value growth.")
    return ("holds broadly flat", "remain broadly flat",
            "The net outlook is broadly flat pricing, so value growth will track volume growth closely.")


def spread_dir(start, end):
    if start <= 0:
        return "moves"
    ch = (end - start) / start
    if ch >= T["spread_change"]:
        return "widens"
    if ch <= -T["spread_change"]:
        return "narrows"
    return "holds broadly steady"


def direction_split(n_up, n_total, down_list):
    if n_up == n_total:
        return f"All {n_total} factors exert upward pressure on price."
    if not down_list:
        return f"{n_up} of {n_total} factors exert upward pressure on price."
    return f"{n_up} of {n_total} factors exert upward pressure on price, while {join_list(down_list)} pull prices down."


def join_list(items):
    items = [i for i in items if i]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + " and " + items[-1]


# ------------------------------------------------------------------------------- yoy
def yoy_trend(early, late):
    d = _bps(late, early)
    if d <= -T["yoy_bps"]:
        return "moderates gradually toward {fc_end} as the category matures"
    if d >= T["yoy_bps"]:
        return "strengthens toward {fc_end} as adoption broadens"
    return "remains broadly steady through {fc_end}"


def yoy_decline(n_decline):
    if n_decline == 0:
        return "and no sub-segment posts a year of decline over the period"
    if n_decline == 1:
        return "while one sub-segment posts at least one year of decline"
    return f"while {n_decline} sub-segments post at least one year of decline"


# ------------------------------------------------------------------------------- forces
def force_rating(score):
    """score on 0-10."""
    if score < 3.5:
        return "low"
    if score < 5.5:
        return "moderate"
    if score < 7.5:
        return "moderate to high"
    return "high"


def likelihood(score):
    return "Low" if score < 4 else ("Medium" if score < 7 else "High")


# ------------------------------------------------------------------------------- dev tracker
DEV_TYPE = {
    "MA": ("Corporate development", "Consolidation: capability is being bought rather than built.", "Corporate development should track targets and deal multiples."),
    "PR": ("Product and R&D", "The company is building capability ahead of demand.", "Product teams should monitor the competitor roadmap."),
    "EX": ("Operations and sales", "The company is adding capacity or reach in growth markets.", "Sales and supply planning should review coverage in the affected markets."),
    "RD": ("Product and R&D", "Innovation spend is rising in this category.", "R&D should benchmark pipeline timing and claims."),
    "RG": ("Regulatory and quality", "A regulatory or certification hurdle has been cleared.", "Regulatory teams should assess the impact on market access."),
    "CP": ("Business development", "Partnerships are being used to extend reach quickly.", "Business development should map partner ecosystems."),
}


# ------------------------------------------------------------------------------- derive all p_* keys
def derive(ctx):
    """Add every p_* phrase that can be computed from the raw inputs present in ctx."""
    c = ctx
    p = {}

    def has(*keys):
        return all(k in c for k in keys)

    if has("r_f_cagr", "r_h_cagr"):
        p["p_trend"] = trend_phrase(c["r_f_cagr"], c["r_h_cagr"])
    if has("r_vol_f_cagr", "r_vol_h_cagr"):
        p["p_vol_trend"] = trend_phrase(c["r_vol_f_cagr"], c["r_vol_h_cagr"])
    if has("r_vol_abs_fc", "r_vol_abs_hist"):
        p["p_vol_period_compare"] = period_compare(c["r_vol_abs_fc"], c["r_vol_abs_hist"])
    if has("r_abs_opp", "r_abs_hist"):
        p["p_val_period_compare"] = period_compare(c["r_abs_opp"], c["r_abs_hist"])
    if has("r_f_cagr", "r_vol_f_cagr"):
        p["p_price_mix"], p["p_price_mix_short"] = price_mix(c["r_f_cagr"], c["r_vol_f_cagr"])

    if has("seg1_lead", "seg1_fast", "seg1_lc"):
        p["p_seg1_fast_clause"] = seg1_fast_clause(c["seg1_lead"], c["seg1_fast"])
        p["p_seg1_lead_fast"] = lead_fast(c["seg1_lead"], c["seg1_fast"], c["seg1_lc"] + " segment",
                                          "v_seg1_lead_share_end", "v_seg1_fast_cagr")
        if _same(c["seg1_lead"], c["seg1_fast"]):
            p["p_approach"] = (f"Build scale in '{c['seg1_lead']}', which is both the largest and the fastest-growing "
                               f"{c['seg1_lc']} category, to capture the {{v_f_cagr}} CAGR trajectory to {{fc_end}}.")
            p["p_segment_priority"] = (f"Prioritise '{c['seg1_lead']}', which anchors {{v_seg1_lead_share_end}} of value and grows at "
                                       f"{{v_seg1_fast_cagr}} CAGR, while seeding adjacent categories for diversification.")
        else:
            p["p_approach"] = (f"Balance scale in '{c['seg1_lead']}' with targeted growth in '{c['seg1_fast']}' to capture the "
                               f"{{v_f_cagr}} CAGR trajectory to {{fc_end}}.")
            p["p_segment_priority"] = (f"Prioritise '{c['seg1_fast']}' ({{v_seg1_fast_cagr}} CAGR) for growth, while '{c['seg1_lead']}' "
                                       f"anchors {{v_seg1_lead_share_end}} of value and funds the portfolio.")
    if has("seg1_gain", "seg1_lose", "r_seg1_gain_bps", "seg1_lc"):
        p["p_seg1_share_shift"] = share_shift(c["seg1_gain"], c["seg1_lose"], c["r_seg1_gain_bps"], c["seg1_lc"],
                                              "v_seg1_gain_bps", "v_seg1_lose_bps")
    if has("r_seg2_fast_cagr", "r_f_cagr"):
        d = _bps(c["r_seg2_fast_cagr"], c["r_f_cagr"])
        p["p_seg2_fast_vs_market"] = ("outpacing the {v_f_cagr} market average by {v_seg2_fast_gap_bps} BPS" if d >= T["bench_bps"]
                                      else "in line with the {v_f_cagr} market average" if d > -T["bench_bps"]
                                      else "below the {v_f_cagr} market average by {v_seg2_fast_gap_bps} BPS")
    if has("seg2_lead", "seg2_fast", "seg2_lc", "n_seg2_sub"):
        if _same(c["seg2_lead"], c["seg2_fast"]):
            p["p_seg2_reach"] = (f"Reach runs through {c['n_seg2_sub']} {c['seg2_lc']} routes, and '{c['seg2_lead']}' is both the largest "
                                 f"at {{v_seg2_lead_share}} and the fastest-growing at {{v_seg2_fast_cagr}}.")
            p["p_o3_body"] = (f"'{c['seg2_lead']}' is both the largest and the fastest-growing of {c['n_seg2_sub']} {c['seg2_lc']} routes, "
                              f"carrying {{v_seg2_lead_share}} of demand and growing at {{v_seg2_fast_cagr}} CAGR. Deepening coverage in this "
                              f"route, and leading with '{{seg1_fast}}' at {{v_seg1_fast_cagr}} CAGR, is the highest-yield combination the data supports.")
        else:
            p["p_seg2_reach"] = (f"Reach runs through {c['n_seg2_sub']} {c['seg2_lc']} routes, led by '{c['seg2_lead']}' at "
                                 f"{{v_seg2_lead_share}}, with '{c['seg2_fast']}' growing fastest at {{v_seg2_fast_cagr}}.")
            p["p_o3_body"] = (f"'{c['seg2_fast']}' grows at {{v_seg2_fast_cagr}} CAGR, the fastest of {c['n_seg2_sub']} {c['seg2_lc']} routes, "
                              f"and adds {{v_seg2_fast_opp}} {{val_unit}} over the forecast period, while '{c['seg2_lead']}' still carries "
                              f"{{v_seg2_lead_share}} of demand today. Pairing the incumbent route for volume with the fastest route for growth, "
                              f"and leading with '{{seg1_fast}}' at {{v_seg1_fast_cagr}} CAGR, is the highest-yield combination the data supports.")

    if has("child_lead", "child_fast", "child_label"):
        p["p_child_lead_fast"] = child_lead_fast(c["child_lead"], c["child_fast"], c["child_label"])
        p["p_child_lead_fast_short"] = child_lead_fast_short(c["child_lead"], c["child_fast"], c["child_label"])
        if _same(c["child_lead"], c["child_fast"]):
            p["p_comm_focus"] = f"Commercial focus is intensifying in {c['child_lead']}, which accounts for {{v_child_lead_share_end}} of {{fc_end}} value."
            p["p_region_priority"] = (f"Lead with {c['child_lead']}, which holds {{v_child_lead_share_end}} of value and grows fastest at "
                                      f"{{v_child_fast_cagr}} CAGR, while building positions in the next-largest {{child_label_pl}}.")
        else:
            p["p_comm_focus"] = (f"Commercial focus is intensifying in {c['child_lead']} and {c['child_fast']}, which together account for "
                                 f"{{v_child_lead_fast_share}} of {{fc_end}} value.")
            p["p_region_priority"] = (f"Lead with {c['child_lead']} ({{v_child_lead_share_end}} of value) and capture upside in "
                                      f"{c['child_fast']} ({{v_child_fast_cagr}} CAGR).")
    if has("n_children", "child_label_pl"):
        p["p_child_landscape"] = child_landscape(c["n_children"], c["child_label_pl"])
    if has("n_above_pace", "n_children", "child_label_pl", "geo"):
        p["p_above_pace"] = above_pace(c["n_above_pace"], c["n_children"], c["child_label_pl"], c["geo"])
    if has("child_gain_m", "child_lose_m", "r_child_gain_m_bps"):
        p["p_child_share_shift_m"] = share_shift(c["child_gain_m"], c["child_lose_m"], c["r_child_gain_m_bps"],
                                                 c.get("child_label", "geographic"), "v_child_gain_m_bps", "v_child_lose_m_bps", quote=False)

    # segment chapter
    if has("sub_lead", "sub_fast", "segment_lc"):
        p["p_sub_lead_fast"] = lead_fast(c["sub_lead"], c["sub_fast"], "segment", "v_sub_lead_share_est", "v_sub_fast_cagr")
    if has("sub_gain", "sub_lose", "r_sub_gain_bps", "segment_lc"):
        p["p_sub_share_shift"] = share_shift(c["sub_gain"], c["sub_lose"], c["r_sub_gain_bps"], c["segment_lc"])
    if has("sub_gain_vol", "sub_lose_vol", "r_sub_vol_gain_bps", "segment_lc"):
        p["p_sub_vol_share_shift"] = share_shift(c["sub_gain_vol"], c["sub_lose_vol"], c["r_sub_vol_gain_bps"], c["segment_lc"],
                                                 "v_sub_vol_gain_bps", "v_sub_vol_lose_bps")
    if has("n_sub"):
        if c["n_sub"] >= 4 and "sub_opp_low" in c:
            p["p_sub_opp_tail"] = (f"'{c['sub_opp_low']}' adds the least, at {{v_sub_opp_low}} {{val_unit}} "
                                   f"({{v_sub_opp_low_share}} of the increment).")
        else:
            p["p_sub_opp_tail"] = "Together, these sub-segments account for the full incremental opportunity of the segment."
    if has("n_sub", "segment_lc"):
        p["p_sub_landscape"] = (f"{c['n_sub']} sub-segments make up the {c['segment_lc']} landscape."
                                if c["n_sub"] <= 2 else
                                f"{c['n_sub']} sub-segments make up the {c['segment_lc']} landscape, and the top two account for {{v_sub_top2_share}} of {{est_year}} value.")
    if has("sub_opp_lead", "sub_gain"):
        p["p_c03_takeaway"] = ("'{sub_opp_lead}' contributes {v_sub_opp_lead_share} of {geo} incremental value and also records the largest share gain, at {v_sub_gain_bps} BPS."
                               if _same(c["sub_opp_lead"], c["sub_gain"]) else
                               "'{sub_opp_lead}' contributes {v_sub_opp_lead_share} of {geo} incremental value, while '{sub_gain}' records the largest share gain at {v_sub_gain_bps} BPS.")
    if has("sub_lead", "sub_fast"):
        p["p_c04_takeaway"] = ("'{sub_lead}' remains the largest {segment_lc} category through {fc_end} and also grows fastest, at {v_sub_fast_cagr} CAGR."
                               if _same(c["sub_lead"], c["sub_fast"]) else
                               "'{sub_lead}' remains the largest {segment_lc} category through {fc_end}, while '{sub_fast}' grows fastest at {v_sub_fast_cagr} CAGR.")
    if has("r_yoy_early", "r_yoy_late"):
        p["p_yoy_trend"] = yoy_trend(c["r_yoy_early"], c["r_yoy_late"])
    if has("n_sub_decline"):
        p["p_yoy_decline"] = yoy_decline(c["n_sub_decline"])

    # cross-section
    if has("x_conc_a", "x_conc_b", "x_top_share_a", "x_top_share_b"):
        if _same(c["x_conc_a"], c["x_top_share_a"]) and _same(c["x_conc_b"], c["x_top_share_b"]):
            p["p_x_concentration"] = (f"'{c['x_div_a']}' is the most diversified {{seg_a_lc}} sub-segment, with no single {{seg_b_lc}} "
                                      f"sub-segment above {{v_x_div_max}} of its value.")
        else:
            p["p_x_concentration"] = (f"'{c['x_conc_a']}' is the most concentrated {{seg_a_lc}} sub-segment, with {{v_x_conc_share}} "
                                      f"of its value in '{c['x_conc_b']}'.")

    # country
    if has("r_m_f_cagr", "r_parent_cagr", "parent_geo"):
        p["p_vs_parent"] = vs_benchmark(c["r_m_f_cagr"], c["r_parent_cagr"], c["parent_geo"], "v_parent_cagr", "v_parent_gap_bps")
        p["p_bench_arrow"] = bench_arrow(c["r_m_f_cagr"], c["r_parent_cagr"])

    # indices
    if has("cty_cagr_top", "cty_abs_top"):
        if _same(c["cty_cagr_top"], c["cty_abs_top"]):
            p["p_index_leaders"] = (f"{{cty_cagr_top_The}} leads on both indices, growing at {{v_cty_cagr_top_cagr}} against the {{geo}} "
                                    f"benchmark of {{v_f_cagr}} and adding {{v_cty_abs_top_opp}} {{val_unit}} over {{fc_period}} "
                                    f"({{v_cty_abs_top_opp_share}} of the {{geo}} increment).")
        else:
            p["p_index_leaders"] = (f"{{cty_cagr_top_The}} posts the highest CAGR Index ({{v_cty_cagr_top_idx}}), growing at "
                                    f"{{v_cty_cagr_top_cagr}} against the {{geo}} benchmark of {{v_f_cagr}}, while {{cty_abs_top_the}} "
                                    f"offers the largest absolute opportunity, adding {{v_cty_abs_top_opp}} {{val_unit}} over {{fc_period}} "
                                    f"({{v_cty_abs_top_opp_share}} of the {{geo}} increment).")

    # pricing
    if has("r_child_price_spread"):
        p["p_price_band"] = price_band(c["r_child_price_spread"])
    if has("r_price_conv"):
        p["p_price_conv"] = price_conv(c["r_price_conv"])
    if has("r_price_avg_cagr"):
        p["p_price_dir_verb"], p["p_price_dir_inf"], p["p_price_outlook"] = price_direction(c["r_price_avg_cagr"])
    if has("r_price_gap_start", "r_price_gap_end"):
        p["p_price_spread_dir"] = spread_dir(c["r_price_gap_start"], c["r_price_gap_end"])
    if has("down_factors", "n_up_factors", "n_factors"):
        p["p_direction_split"] = direction_split(c["n_up_factors"], c["n_factors"], c["down_factors"])
        p["p_down_list"] = join_list([d.lower() for d in c["down_factors"]]) or "efficiency gains"

    # competition and supply
    if has("r_cr3"):
        p["p_concentration"] = concentration(c["r_cr3"])
        p["p_competition_line"] = (f"The market is {concentration(c['r_cr3'])}, with the top three players accounting for an estimated "
                                   f"{{v_cr3}} of {{est_year}} value.")
        p["p_k01_conc_head"] = ("Leadership is concentrated in value" if c["r_cr3"] >= T["cr3_mid"]
                                else "Value is spread across many suppliers")
    elif has("seg1_lead"):
        p["p_competition_line"] = (f"'{c['seg1_lead']}' holds {{v_seg1_lead_share_end}} of {{fc_end}} value, which makes it the anchor "
                                   f"{c['seg1_lc']} category for suppliers.")
    if has("r_child_lead_share_end"):
        p["p_supply_conc"] = supply_conc(c["r_child_lead_share_end"])
        if c["r_child_lead_share_end"] > 0.5:
            p["p_entry_spread"] = f"{c['child_lead']} alone accounts for {{v_child_lead_share_end}} of value, so meaningful entry requires scale in that {{child_label}}."
        else:
            p["p_entry_spread"] = "Demand spreads across {n_children} {child_label_pl}, with no single {child_label} above {v_child_lead_share_end}, so regional entry is feasible."
        conc = c["r_child_lead_share_end"] >= 0.30 or c.get("r_seg1_lead_share_end", 0) >= 0.50
        if conc:
            p["p_r3_head"] = "Concentration Risk Limits Pricing and Sourcing Flexibility"
            p["p_r3_body"] = ("The market is concentrated, as {child_lead} accounts for {v_child_lead_share_end} of {fc_end} value and the "
                              "leading {seg1_lc} category for {v_seg1_lead_share_end} of its segment. Average realisation moves "
                              "{v_price_avg_cagr} a year, so cost movements cannot be passed on evenly across {n_children} {child_label_pl}. "
                              "Commercial terms set for the market as a whole underserve part of it, and margin is lost in the mismatch "
                              "rather than in the headline growth rate.")
        else:
            p["p_r3_head"] = "Input-Cost Volatility Pressures Margins"
            p["p_r3_body"] = ("Exposure to {lx_input_costs} leaves suppliers vulnerable to cost spikes that cannot be passed on immediately, "
                              "as average realisation moves only {v_price_avg_cagr} a year. {lx_competitive_pressure_cap} limit how "
                              "quickly increases reach the selling price. Suppliers that hedge inputs, diversify sourcing and improve "
                              "operating efficiency are best placed to protect margins through {fc_end}.")
    for f in ("suppliers", "buyers", "rivalry", "substitutes", "entrants"):
        if f"s_force_{f}" in c:
            p[f"p_force_{f}"] = force_rating(c[f"s_force_{f}"])
    if has("r_hhi"):
        p["p_hhi_band"] = hhi_band(c["r_hhi"])
    if has("r_t1_share"):
        p["p_tier_word"] = "steep" if c["r_t1_share"] >= 0.5 else "relatively flat"
        p["p_t1_barrier"] = ("Tier 1 players collectively command about {v_t1_share} of the market, which underlines meaningful barriers to entry."
                             if c["r_t1_share"] >= 0.4 else
                             "Tier 1 players collectively command about {v_t1_share} of the market, which leaves room for mid-tier challengers.")
    if has("n_ma", "n_named"):
        if c["n_ma"] >= 3:
            p["p_ma_head"] = "Consolidation is an active lever"
            p["p_ma_body"] = f"Acquisition is an active strategy at {c['n_ma']} of {c['n_named']} named players, and the pool of independent targets is shrinking."
        else:
            p["p_ma_head"] = "Organic growth dominates strategy"
            p["p_ma_body"] = f"Most named players are growing organically through launches and expansion, and acquisitions feature at only {c['n_ma']} of {c['n_named']}."
    if has("enabler_heads"):
        p["p_enablers"] = " | ".join(c["enabler_heads"])
    if has("constraint_heads"):
        p["p_constraints"] = " | ".join(c["constraint_heads"])
    if has("leaders_same"):
        p["p_leader_consistency"] = (
            "'{item_leader_all}' leads in every case shown, which reflects portfolio breadth and scale advantages that challengers must target one by one."
            if c["leaders_same"] else
            "Leadership shifts across the cases shown, so a strong position in one does not carry over to the others and targeted planning is essential.")

    # company
    if has("c_is_leader"):
        p["p_c_position"] = ("Competitive position: {company} is the market leader with about {v_c_share} of the {geo_ref} market, which makes it the benchmark for peers."
                             if c["c_is_leader"] else
                             "Competitive position: {company} holds about {v_c_share} of the {geo_ref} market against {v_leader_share} for the market leader, a scale gap of {v_c_gap_pts} percentage points.")
    if has("r_c_rev_cagr"):
        r = c["r_c_rev_cagr"]
        p["p_c_rev_trend"] = "grew" if r > 0.005 else ("declined" if r < -0.005 else "held broadly flat")
    return p


def item_rules(over, item, ctx):
    """Per-item phrases for repeat blocks."""
    p = {}
    if over in ("subs", "children"):
        geo_mode = over == "children"
        scope = "market" if geo_mode else "segment"
        ref = ctx.get("r_f_cagr", 0) if geo_mode else ctx.get("r_seg_cagr", ctx.get("r_f_cagr", 0))
        p["p_item_role"] = item_role(item.get("is_lead", False), item.get("is_fast", False),
                                     item.get("r_item_share", 0), item.get("r_item_cagr", 0), ref, scope, geo_mode)
        if geo_mode:
            # M01 cells
            d, r = [], []
            if item.get("rank", 99) <= 3:
                d.append("A large installed base holds {v_item_share_end} of value.")
            if _bps(item.get("r_item_cagr", 0), ctx.get("r_f_cagr", 0)) >= T["bench_bps"]:
                d.append("Growth runs ahead of the market at {v_item_cagr} CAGR.")
            else:
                d.append("Growth of {v_item_cagr} CAGR rests on an established demand base.")
            if _bps(item.get("r_item_cagr", 0), item.get("r_item_h_cagr", 0)) >= T["trend_bps"]:
                d.append("Growth accelerates against its historical trend.")
            d.append("Demand is led by '{item_seg1_lead}'.")
            if _bps(item.get("r_item_cagr", 0), ctx.get("r_f_cagr", 0)) <= -T["bench_bps"]:
                r.append("Growth trails the market by {v_item_gap_bps} BPS.")
            if item.get("r_item_price_cagr", 0) > T["price_flat"]:
                r.append("Rising price realisation of {v_item_price_cagr} a year may slow volume uptake.")
            if item.get("r_item_share", 0) >= 0.30:
                r.append("High maturity limits headroom for new entrants.")
            if not r:
                r.append("Competition from established suppliers caps pricing headroom.")
            p["p_item_drivers"] = " ".join(d)
            p["p_item_restraints"] = " ".join(r)
    if over == "drivers_top5":
        p["p_item_driver_role"] = driver_role(item.get("rank", 5))
    if over == "developments":
        dept, signal, action = DEV_TYPE.get(item.get("c_dev_type", ""), ("n/a", "n/a", "n/a"))
        p.update(p_dev_dept=dept, p_dev_signal=signal, p_dev_action=action)
    return p
