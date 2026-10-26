"""Structured data for the Value Chain section (Value Chain Analysis).

Provides a documented, market-agnostic default profile - value-chain stages
with value-addition ranges and stakeholder notes, industry-wide operating
margins by region × node, and a buyers-assessment roster - that renders for
ANY loaded market and is overridable later via the workbook.  Nothing
market-specific is hard-coded in the page templates.
"""
from __future__ import annotations

import math

from . import data
from .config import SECTION_PALETTE, UNIT_LABELS

INDICATIVE_NOTE = ("Note: value-chain economics are indicative derived from "
                   "market modelling and expert validation, not company "
                   "disclosures.")

# The price label the rest of the deck uses (pricing pages, competitor
# pricing).  Read through a function rather than frozen at import time: the
# unit is only known once a workbook has been loaded, and the deck can switch
# markets in one process.  Module `__getattr__` keeps the old
# `VC.PRICE_UNIT` spelling working for the templates that read it.
def _price_unit() -> str:
    from . import config
    return config.price_unit()


def __getattr__(name):
    if name == "PRICE_UNIT":
        return _price_unit()
    raise AttributeError(name)

# ── value-chain stages (generalized 5-stage chain) ──────────────────
# name / short value-addition range / stakeholder note.  Colours ramp from
# upstream (navy) to downstream (amber).
STAGE_COLORS = list(SECTION_PALETTE[:5])

STAGES = [
    {"name": "Raw Material\nSuppliers", "valueadd": "5–10%", "icon": "sprout",
     "title": "Raw Material Suppliers",
     "note": "Growers and primary producers supply core inputs; some "
             "forward-integrate into small-scale processing."},
    {"name": "Manufacturers", "valueadd": "20–40%", "icon": "factory",
     "title": "Manufacturers / Processors",
     "note": "Convert inputs into finished and semi-finished product the "
             "single largest value-addition stage."},
    {"name": "Distributors", "valueadd": "5–25%", "icon": "truck",
     "title": "Distributors & Wholesalers",
     "note": "Move product to market, hold inventory and provide logistics, "
             "reach and credit to downstream channels."},
    {"name": "Retailers", "valueadd": "10–35%", "icon": "shop",
     "title": "Retailers & Channels",
     "note": "Sell to end customers across offline and online channels, "
             "capturing point-of-sale margin and brand visibility."},
    {"name": "End Users", "valueadd": "10–20%", "icon": "person",
     "title": "End Users",
     "note": "Consumers and downstream industries; large end users sometimes "
             "backward-integrate to secure supply and cost."},
]

MARGIN_NODES = ["Raw Material Suppliers", "Manufacturers", "Distributors",
                "Retailers"]

# base operating margin (%) per node - upstream lower, mid-chain higher
_NODE_MARGIN = {"Raw Material Suppliers": 10.0, "Manufacturers": 18.0,
                "Distributors": 8.0, "Retailers": 12.0}
_NODE_NOTE = {
    "Raw Material Suppliers": "Commodity-linked",
    "Manufacturers": "Highest margins (value-add + scale)",
    "Distributors": "Volume-led, thin",
    "Retailers": "Channel-dependent PoS margin",
}


def _regions(n: int = 4):
    regs = list(data.REGIONS)
    return regs[:n] if regs else ["North America", "Europe", "Asia-Pacific"]


def value_added_note(market_label: str) -> list[str]:
    """Read-out bullets for the value-chain map - adapts to the loaded market."""
    top = max(STAGES, key=lambda s: _mid(s["valueadd"]))
    return [
        f"The {market_label} value chain runs from raw-material suppliers "
        f"through manufacturers, distributors and retailers to end users.",
        f"'{top['title']}' captures the largest value addition "
        f"({top['valueadd']}), reflecting processing and scale advantages.",
        "Forward and backward integration is common some suppliers process, "
        "and some large end users produce their own input.",
        "Percentages show each stage's indicative share of value added along "
        "the chain.",
    ]


def _mid(rng: str) -> float:
    nums = [float(x) for x in rng.replace("%", "").replace("–", "-").split("-")
            if x.strip().replace(".", "").isdigit()]
    return sum(nums) / len(nums) if nums else 0.0


def margin_rows(n_regions: int = 4):
    """[(region_label, node, margin%, comparison%, note)] for the operating-
    margin table - region shown once per group.  Deterministic and generalized;
    comparison is vs a parent/associated market benchmark."""
    out = []
    for ri, region in enumerate(_regions(n_regions)):
        for ni, node in enumerate(MARGIN_NODES):
            base = _NODE_MARGIN[node]
            margin = base + 2.0 * math.sin(ri * 0.9 + ni)
            comparison = margin - (2.5 + 0.6 * ni) + 0.8 * math.cos(ri)
            out.append({
                "region": region if ni == 0 else "",
                "node": node,
                "margin": f"{margin:.1f}%",
                "comparison": f"{comparison * 100:+,.0f} BPS",
                "note": _NODE_NOTE[node]})
    return out


# ── buyers assessment roster (indicative) ───────────────────────────
_BUYER_CATEGORY = ["Manufacturer", "Distributor", "Retailer", "Others"]
_BUYER_TYPE = ["B2B / Industrial", "B2B / Institutional", "B2C / Retail",
               "Contract / OEM"]
_FACTORS = ["Price & reliability", "Quality & compliance",
            "Lead time & service", "Capacity & scale", "Brand & support"]


# The listed buyers stand for this much of estimate-year demand, split on a
# declining curve, and pay between these multiples of the market's own average
# realisation.  Indicative - but sized off the workbook, so the roster cannot
# out-buy the market it sits in.
_ROSTER_SHARE = 0.32
_PRICE_LO, _PRICE_HI = 0.88, 1.18


def _price_txt(v) -> str:
    """A price with enough decimals to stay meaningful whatever the workbook's
    unit scale - the same rule the pricing pages use."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "0"
    if not v or math.isnan(v) or math.isinf(v):
        return "0"
    d = 2 - int(math.floor(math.log10(abs(v))))
    return f"{v:,.{max(0, min(d, 8))}f}"


def _market_scale():
    """(estimate-year volume, average realisation) for the loaded market."""
    from . import metrics
    from .config import EST_YEAR
    vol = data.total_series("Global", "Volume")
    val = data.total_series("Global", "Value")
    return float(vol[EST_YEAR]), metrics.avg_price(val, vol, [EST_YEAR])


def buyers(n: int = 10):
    """Representative buyer/end-user roster - indicative, but derived.

    The three numeric columns used to be independent arithmetic on the row
    INDEX (`4000 + (i * 1370) % 9000` and friends), which read plausibly and
    reconciled with nothing: the ten buyers procured twenty-six times the
    whole market, their price sat three orders of magnitude off the report's
    own average realisation, and volume x price did not come to the revenue
    in the next column.  They are now sized off the workbook - each buyer
    takes a slice of estimate-year volume on a declining curve, pays a price
    in a spread around the market's average realisation, and its revenue is
    the product of the two, so a reader can multiply any row out.  Still
    indicative (see INDICATIVE_NOTE), just no longer self-contradictory."""
    regs = _regions(8)
    ctys = list(data.ALL_COUNTRIES) if getattr(data, "ALL_COUNTRIES", None)         else regs
    vol_tot, price = _market_scale()
    if not vol_tot or price != price:              # value-only workbook
        vol_tot, price = float(n), float("nan")
    weights = [float(n - i) for i in range(n)]     # biggest buyer first
    wsum = sum(weights) or 1.0
    span = max(1, n - 1)
    out = []
    for i in range(n):
        vol = vol_tot * _ROSTER_SHARE * weights[i] / wsum
        p = price * (_PRICE_LO + (_PRICE_HI - _PRICE_LO) * i / span)
        country = str(ctys[i % len(ctys)]) if ctys else " "
        out.append({
            "company": f"Company {i + 1}",
            "country": country,
            # the country's OWN region - the two lists used to be indexed
            # independently, pairing "Germany" with "Middle East & Africa"
            "region": data.region_of(country) or regs[i % len(regs)],
            "category": _BUYER_CATEGORY[i % len(_BUYER_CATEGORY)],
            "type": _BUYER_TYPE[i % len(_BUYER_TYPE)],
            # the workbook's own volume unit - "MT" was hard-coded here, so
            # a market sized in Units or Kg still read "4,000 MT"
            "volume": f"{vol:,.0f} {UNIT_LABELS['Volume']}",
            "price": _price_txt(p),
            "revenue": f"{vol * p:,.0f}" if p == p else "n/a",
            "factors": _FACTORS[i % len(_FACTORS)]})
    return out


# -- Template switch support ------------------------------------------
def _retheme() -> None:
    """Re-derive what this module bakes from theme tokens (see theming.py)."""
    global STAGE_COLORS
    STAGE_COLORS = list(SECTION_PALETTE[:5])
