"""Structured data configuration for the *Key Pricing Drivers* page.

The page shows a radar/spider chart of ten canonical pricing-driver factors
(each scored 0-10 for how strongly it influences pricing in the loaded market)
plus a set of data-driven Key Takeaways.

Design intent (see ``templates_main.pricing_drivers``):

* The ten factors and their scores are **configuration, not UI code** - they
  live here so scores can be supplied/edited later WITHOUT touching the page
  or the figure/export code.
* Scores are read from the loaded workbook when it carries them (a sheet whose
  name contains "pricing driver" / "price driver", with a factor column and a
  score column).  When the workbook has no such data - the common case today -
  a documented, market-agnostic **default profile** is used so the page still
  renders for every market.
* Nothing market-specific is hard-coded in the page component: the market name
  comes from ``data.MARKET_LABEL`` and the takeaways adapt to whichever factors
  score highest for the loaded market.
"""
from __future__ import annotations

import pandas as pd

from . import data
from .config import DATA_FILE

# ── the ten canonical factors (order = radar sweep, clockwise from top) ──
FACTORS: list[str] = [
    "Input Cost",
    "Labor & Production Costs",
    "Technology & Innovation",
    "Supply Chain & Logistics",
    "Demand & Market Conditions",
    "Competitive Landscape",
    "Regulatory & Compliance Costs",
    "Currency & Trade Dynamics",
    "Brand & Value Perception",
    "Aftermarket & Service Costs",
]

SCALE_MAX = 10.0                      # radar axis maximum (scores are 0..10)

# Default impact profile (0-10).  Deliberately generic / market-agnostic - it
# gives a sensible, well-shaped radar for any market until real per-market
# scores are supplied via the workbook (see ``load_scores``).  Edit these here
# (or, preferably, provide a workbook sheet) - never inside the page component.
_DEFAULT_SCORES: dict[str, float] = {
    "Input Cost": 9.0,
    "Labor & Production Costs": 7.0,
    "Technology & Innovation": 6.5,
    "Supply Chain & Logistics": 7.5,
    "Demand & Market Conditions": 8.5,
    "Competitive Landscape": 8.0,
    "Regulatory & Compliance Costs": 6.0,
    "Currency & Trade Dynamics": 5.5,
    "Brand & Value Perception": 6.5,
    "Aftermarket & Service Costs": 4.5,
}

# short, lower-case fragments used to weave factors into the takeaway prose
_FACTOR_PHRASE: dict[str, str] = {
    "Input Cost": "raw-material and input costs",
    "Labor & Production Costs": "labour and production expenses",
    "Technology & Innovation": "technology and process innovation",
    "Supply Chain & Logistics": "supply-chain and logistics conditions",
    "Demand & Market Conditions": "demand and broader market conditions",
    "Competitive Landscape": "competitive intensity",
    "Regulatory & Compliance Costs": "regulatory and compliance costs",
    "Currency & Trade Dynamics": "currency and trade dynamics",
    "Brand & Value Perception": "brand strength and value perception",
    "Aftermarket & Service Costs": "aftermarket and service costs",
}


def _from_workbook() -> "dict[str, float] | None":
    """Best-effort read of per-market factor scores from the loaded workbook.

    Looks for a sheet whose name contains 'pricing driver' or 'price driver'
    and returns {factor: score} for any of the canonical FACTORS it lists.
    Returns None when no such sheet/data exists.  Never raises - a malformed
    or missing sheet just falls back to the defaults."""
    try:
        try:
            xl = pd.ExcelFile(DATA_FILE, engine="calamine")
        except Exception:
            xl = pd.ExcelFile(DATA_FILE)
        target = None
        for sn in xl.sheet_names:
            low = str(sn).strip().lower()
            if "pricing driver" in low or "price driver" in low:
                target = sn
                break
        if target is None:
            return None
        raw = pd.read_excel(xl, sheet_name=target, header=None)
        canon = {f.lower(): f for f in FACTORS}
        scores: dict[str, float] = {}
        for _, row in raw.iterrows():
            cells = list(row.values)
            label = None
            for c in cells:
                if isinstance(c, str) and c.strip().lower() in canon:
                    label = canon[c.strip().lower()]
                    break
            if label is None:
                continue
            for c in cells:                       # first numeric cell = score
                try:
                    v = float(c)
                except (TypeError, ValueError):
                    continue
                if v == v:                        # not NaN
                    scores[label] = max(0.0, min(SCALE_MAX, v))
                    break
        return scores or None
    except Exception:
        return None


def load_scores() -> "list[tuple[str, float]]":
    """Ordered [(factor, score)] for the radar - workbook values where present,
    otherwise the default profile.  Order always follows FACTORS."""
    wb = _from_workbook() or {}
    return [(f, float(wb.get(f, _DEFAULT_SCORES[f]))) for f in FACTORS]


def takeaways(market_label: str, scored: "list[tuple[str, float]]") -> list[str]:
    """Concise, presentation-ready Key-Takeaway bullets that adapt to the market
    and to whichever factors score highest.  Not identical across markets: the
    market name and the ranked high-/low-impact factors drive the wording."""
    ranked = sorted(scored, key=lambda kv: kv[1], reverse=True)
    names = [n for n, _ in ranked]

    def ph(name: str) -> str:
        return _FACTOR_PHRASE.get(name, name.lower())

    top1, top2, top3 = names[0], names[1], names[2]
    top4 = names[3]
    low1 = names[-1]
    return [
        f"Pricing across the {market_label} is shaped by a mix of input costs, "
        f"operational efficiency and market dynamics with {ph(top1)} carrying "
        f"the single heaviest weight on selling prices today.",
        f"{ph(top1).capitalize()} and {ph(top2)} are the dominant cost "
        f"pressures; as they fluctuate they feed directly into prices and "
        f"squeeze margins along the value chain.",
        f"{ph(top3).capitalize()} and {ph(top4)} are the next most influential "
        f"forces helping offset cost inflation and setting how aggressively "
        f"players can price without ceding share.",
        f"Demand strength, competitive intensity and supply-chain reliability "
        f"together define the ceiling on pricing power across the "
        f"{market_label}.",
        f"Regulatory, currency and brand factors add second-order effects while "
        f"{ph(low1)} exert the least pressure collectively creating a pricing "
        f"balance that must be monitored to stay competitive and profitable.",
    ]


# role tags for the ranked "Highest-Impact Drivers" list, by rank
_RANK_ROLE = ["dominant cost pressure", "major cost pressure",
              "key pricing lever", "key pricing lever", "supporting factor"]


def top_drivers(scored: "list[tuple[str, float]]", n: int = 5) -> list[str]:
    """Ranked one-liners for the top ``n`` factors - a compact, presentation-
    ready list ('Factor - 9.0 / 10 · role') that quantifies the radar and ties
    the prose takeaways to the chart.  Fully derived from the scores."""
    ranked = sorted(scored, key=lambda kv: kv[1], reverse=True)[:n]
    out = []
    for i, (name, score) in enumerate(ranked):
        role = _RANK_ROLE[min(i, len(_RANK_ROLE) - 1)]
        out.append(f"{name} {score:.1f} / {SCALE_MAX:.0f}  ·  {role}")
    return out
