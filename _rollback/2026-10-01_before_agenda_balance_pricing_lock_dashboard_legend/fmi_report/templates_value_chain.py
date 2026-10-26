"""Page templates for the Value Chain section (main report).

Three pages, modelled on the reference Value-Chain deck and adapted to this
project's data-driven, dashboard↔PPT-parity architecture:
  1. Value Chain & Stakeholder Mapping  - a chevron flow (custom block)
  2. Industry-Wide Operating Margin      - Region × Node table
  3. Buyers Assessment                   - buyer/end-user roster table
All content comes from value_chain.py (documented defaults, workbook-
overridable); nothing market-specific is hard-coded here.
"""
from __future__ import annotations

import pandas as pd

from . import config, layout as L, value_chain as VC
from .config import UNIT_LABELS
from .data import MARKET_LABEL as MKT


# ═══════════ 1. Value Chain & Stakeholder Mapping (chevron flow) ═════
def value_chain_map():
    return (f"{MKT} Value Chain & Stakeholder Mapping", {"valuechain": {
        "subtitle": "End-to-end value chain with the key stakeholders at each "
                    "stage and their indicative contribution to value addition.",
        "stages": VC.STAGES, "colors": VC.STAGE_COLORS,
        # the year the stage analysis is based on, under each box title
        # (client, 2026-09-30)
        "base_year": f"Base Year: {config.EST_YEAR}",
        "readout": VC.value_added_note(MKT), "note": VC.INDICATIVE_NOTE}})


# ═══════════ 2. Industry-Wide Operating Margin analysis ══════════════
def operating_margin():
    rows = VC.margin_rows(3)
    idx = [r["region"] for r in rows]
    df = pd.DataFrame({
        "Node of Supply Chain": [r["node"] for r in rows],
        "Operating Margin": [r["margin"] for r in rows],
        "Comparison to Parent Market": [r["comparison"] for r in rows],
        "Notes": [r["note"] for r in rows]}, index=idx)
    desc = [
        "Node of Supply Chain: each major stage, from raw-material suppliers "
        "through to retailers.",
        "Operating Margin: the indicative margin earned at each node, by region.",
        "Comparison to Parent Market: margin versus the parent / associated "
        "market (e.g. the broader food or pharmaceutical value chain).",
        "Margins are broken down by both region and node to surface where "
        "value concentrates along the chain.",
    ]
    return (f"{MKT} Industry-Wide Operating Margin Analysis", {"rows": [
        L.frow(L.tbl(df, index_name="Region",
                     title=f"{MKT} Operating Margin by Region & Node")),
        L.frow(L.summary("How to Read This", desc, h=150)),
    ]})


# ═══════════ 3. Buyers Assessment ═══════════════════════════════════
def buyers_assessment():
    b = VC.buyers(10)
    df = pd.DataFrame({
        "Country": [x["country"] for x in b],
        "Region": [x["region"] for x in b],
        "Buyer Category": [x["category"] for x in b],
        "Buyer / End-User Type": [x["type"] for x in b],
        "Procurement Vol.": [x["volume"] for x in b],
        f"Avg Price ({VC.PRICE_UNIT})": [x["price"] for x in b],
        f"Revenue ({UNIT_LABELS['Value']})": [x["revenue"] for x in b],
        "Key Selection Factor": [x["factors"] for x in b],
    }, index=[x["company"] for x in b])
    note = ("This is a representative subset the full assessment lists 20–30 "
            "active participants (with brands, specifications, procurement "
            "strategies and SWOT). " + VC.INDICATIVE_NOTE)
    return (f"{MKT} Buyers Assessment", {"rows": [
        L.frow(L.tbl(df, index_name="Company",
                     title=f"{MKT} Buyers / End-User Assessment")),
        L.frow(L.summary("Notes", [note], h=96)),
    ]})
