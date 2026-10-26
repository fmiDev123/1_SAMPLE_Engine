"""Structured competitor data for the '- Competitor' report (Competition
Assessment).

The market workbook carries no competitor data, so this module supplies a
documented, **market-agnostic default profile** - a company roster, tier
structure, value shares (global / by region / by product type), revenue &
KPI figures, five-year financials and strategy notes - that renders for ANY
loaded market.  All monetary figures are pinned to the loaded market's actual
value total (``data.total_series``), so magnitudes stay believable per market;
only the competitor split is synthetic.

Everything here is **configuration, not UI code**: page templates read from
these functions and never hard-code company names/shares.  If a workbook later
adds a sheet whose name contains 'competition'/'competitor', ``_from_workbook``
can be extended to read real shares; until then the defaults are used and every
competitor slide is clearly captioned as indicative.
"""
from __future__ import annotations

import math

from . import data, figures, geo, metrics
from .config import (ARROW_HIGH, ARROW_LOW, ARROW_MED,
                     BOX_COLORS as _T_BOX_COLORS, END_YEAR, EST_YEAR,
                     STAR_COLORS as _T_STAR_COLORS, UNIT_LABELS)
from .data import MARKET_LABEL as MKT

INDICATIVE_NOTE = ("Note: competitor figures are indicative derived from "
                   "market modelling and expert validation, not company "
                   "disclosures.")

# ── company roster (indicative placeholders, overridable via workbook) ──
# Ranked, so index 0 is the largest player.  "Others" is appended by the
# share functions and is never part of this list.
COMPANIES: list[str] = [
    "Company A", "Company B", "Company C", "Company D", "Company E",
    "Company F", "Company G", "Company H",
]
OTHERS = "Others"

# base global value share per company (%), descending; the remainder is Others
# (the last two only come into play when the workbook names ten players)
_BASE_SHARE = [13.5, 11.5, 10.0, 8.5, 7.0, 5.5, 4.5, 3.5, 2.8, 2.2]

_HEADER_WORDS = {"rank", "#", "no", "no.", "s.no", "sr. no", "sl. no",
                 "company", "companies", "company name", "name", "player",
                 "players", "top 10 companies", "top companies",
                 "key players"}


def _top_companies_sheet() -> list[dict]:
    """The players on the workbook's 'Top 10 Companies' sheet, in the order
    given (the first is the largest): [{"name", "share", "tier"}], share in
    % and tier as written ('Tier 1'), either None when the sheet has no such
    column.  [] when there is no sheet.

    Columns are found by their headers - 'Company' (or Name / Player),
    '... Share ...', 'Tier'; a sheet without a recognisable header row falls
    back to its first text column that is not a tier label.  A leading rank
    ('1.', '2)') is dropped from a name."""
    import re
    import pandas as pd
    try:
        xl = pd.ExcelFile(data.DATA_FILE)
        sheet = next((n for n in xl.sheet_names
                      if " ".join(str(n).lower().split())
                      in ("top 10 companies", "top companies")), None)
        if sheet is None:
            return []
        raw = pd.read_excel(xl, sheet_name=sheet, header=None, dtype=object)
    except Exception:
        return []

    def blank(v):
        return v is None or (isinstance(v, float) and v != v) or \
            not str(v).strip()

    def num(v):
        try:
            return float(str(v).replace("%", "").strip())
        except (TypeError, ValueError):
            return None

    # the header row: the first row naming a company / name / player column
    head_row, cols = None, {}
    for i in range(min(len(raw), 10)):
        cells = [str(v).strip().lower() if not blank(v) else ""
                 for v in raw.iloc[i].tolist()]
        name_c = next((j for j, c in enumerate(cells)
                       if c in _HEADER_WORDS - {"rank", "#", "no", "no.",
                                                "s.no", "sr. no", "sl. no"}
                       or c.startswith("company name")), None)
        if name_c is not None:
            head_row = i
            cols["name"] = name_c
            cols["share"] = next((j for j, c in enumerate(cells)
                                  if "share" in c), None)
            cols["tier"] = next((j for j, c in enumerate(cells)
                                 if c == "tier" or c.startswith("tier ")
                                 and "tier" in c.split()[0]), None)
            break
    out = []
    if head_row is not None:
        for _, r in raw.iloc[head_row + 1:].iterrows():
            v = r.iloc[cols["name"]]
            if blank(v):
                continue
            name = re.sub(r"^\s*\d+\s*[.)\-:]\s*", "", str(v).strip())
            if not name or name.lower() in _HEADER_WORDS:
                continue
            share = num(r.iloc[cols["share"]]) if cols["share"] is not None \
                else None
            tier = (str(r.iloc[cols["tier"]]).strip()
                    if cols["tier"] is not None
                    and not blank(r.iloc[cols["tier"]]) else None)
            out.append({"name": name, "share": share, "tier": tier})
    else:
        best = []
        for col in raw.columns:
            names = []
            for v in raw[col].tolist():
                if blank(v):
                    continue
                t = re.sub(r"^\s*\d+\s*[.)\-:]\s*", "", str(v).strip())
                if (not t or t.lower() in _HEADER_WORDS
                        or t.replace(".", "").isdigit()
                        or re.fullmatch(r"tier\s*\d*", t, flags=re.I)):
                    continue
                names.append(t)
            if len(names) > len(best):
                best = names
        out = [{"name": n, "share": None, "tier": None} for n in best]
    seen, rows = set(), []
    for r in out:
        if r["name"] not in seen:
            seen.add(r["name"])
            rows.append(r)
    return rows[:10]


# 'Company A' is the workbook's first-named player, 'Company B' its second,
# and so on - when the workbook carries the sheet (2026-09-29).  Its shares
# and tiers, where given, replace the indicative ones.
_SHEET = _top_companies_sheet()
_NAMED = [r["name"] for r in _SHEET]
if _NAMED:
    COMPANIES = _NAMED
if _SHEET and all(r["share"] is not None for r in _SHEET):
    _BASE_SHARE = [float(r["share"]) for r in _SHEET]
_SHEET_TIERS = ({r["name"]: r["tier"] for r in _SHEET}
                if _SHEET and all(r["tier"] for r in _SHEET) else {})

# tier structure (share of the whole market) - a concentrated top tier, matching
# a typical FMI 'industry structure' funnel.  Overridable later via workbook.
TIER_SHARE = {"Tier 1": 0.52, "Tier 2": 0.22, "Tier 3": 0.10, "Others": 0.16}
TIER_LABEL = {"Tier 1": "Top / global players",
              "Tier 2": "Mid-sized players",
              "Tier 3": "Emerging / local players",
              "Others": "Long tail"}


def _region_index(region: str) -> int:
    regs = list(data.REGIONS)
    return regs.index(region) if region in regs else -1


def market_value(region: str = "Global", year: int = None) -> float:
    """Loaded market's value total for ``region`` at ``year`` (END_YEAR by
    default; 0.0 if absent).  The industry-structure funnel passes EST_YEAR so
    it reads as a picture of the CURRENT competitive landscape."""
    yr = END_YEAR if year is None else year
    try:
        return float(data.total_series(region, "Value")[yr])
    except Exception:
        try:
            return float(data.total_series("Global", "Value")[yr])
        except Exception:
            return 0.0


def companies() -> list[str]:
    return list(COMPANIES)


def _from_workbook():
    """Placeholder hook for real per-market competitor shares.  Returns None
    today (no such sheet in the standard workbook); documented so real data can
    be wired in later without touching the page templates."""
    return None


def company_shares(region: str = "Global"):
    """Ordered [(company, share%)] for ``region`` including a trailing
    'Others', summing to ~100.  Deterministic per-region variation so the
    region dropdown produces a genuinely different donut for each region."""
    wb = _from_workbook()
    if wb and region in wb:
        return wb[region]
    names = companies()
    if region == "Global" and _SHEET and \
            all(r["share"] is not None for r in _SHEET):
        # the workbook's own shares, as given; Others is the remainder
        out = [(r["name"], float(r["share"])) for r in _SHEET]
        out.append((OTHERS, max(0.0, 100.0 - sum(v for _, v in out))))
        return out
    ridx = _region_index(region) + 1              # Global=0, then each region
    adj = []
    for i, base in enumerate(_BASE_SHARE[:len(names)]):
        wobble = 1.0 + 0.22 * math.sin(i * 1.3 + (ridx + 1) * 0.8)
        adj.append(max(0.6, base * wobble))
    others = 26.0 + 5.0 * math.cos((ridx + 1) * 0.7)      # ~21-31%
    scale = (100.0 - others) / (sum(adj) or 1.0)
    out = [(names[i], adj[i] * scale) for i in range(len(names))]
    out.append((OTHERS, others))
    return out


def company_colors(n: int):
    """Distinct brand-palette colours for ``n`` companies + Others (grey)."""
    cols = figures._pal(n - 1) if n > 1 else figures._pal(1)
    return list(cols) + ["#B9B3AC"]


def tiers():
    """Industry-structure tiers: {tier: {share, size (value units), range,
    players[]}}.  Sizes are pinned to the EST_YEAR market total (the current
    landscape, matching the other competitor-share pages); players are drawn
    from the roster so the funnel adapts to any market."""
    mv = market_value("Global", EST_YEAR)
    names = companies()
    split = {"Tier 1": names[:4], "Tier 2": names[4:6],
             "Tier 3": names[6:], "Others": []}
    tier_share = dict(TIER_SHARE)
    if _SHEET_TIERS:
        # the workbook names each player's tier; a tier beyond Tier 3 joins
        # the long tail
        split = {t: [] for t in TIER_SHARE}
        for n in names:
            t = _SHEET_TIERS.get(n, "Others")
            split[t if t in ("Tier 1", "Tier 2", "Tier 3") else "Others"] \
                .append(n)
        if all(r["share"] is not None for r in _SHEET):
            sh = {r["name"]: float(r["share"]) / 100 for r in _SHEET}
            for t in ("Tier 1", "Tier 2", "Tier 3"):
                tier_share[t] = sum(sh[n] for n in split[t])
            tier_share["Others"] = max(0.0, 1.0 - sum(
                tier_share[t] for t in ("Tier 1", "Tier 2", "Tier 3")))
    out = {}
    for t, sh in tier_share.items():
        lo = int(round(sh * 100)) - 1
        hi = int(round(sh * 100)) + 1
        out[t] = {"share": sh, "size": mv * sh, "range": f"~{lo}-{hi}%",
                  "players": split.get(t, []), "label": TIER_LABEL[t]}
    return out


def _top_subsegments(n: int = 3):
    """Top-``n`` sub-segments of the first segment by END_YEAR value (the
    market's leading 'product types').  Falls back to generic labels."""
    try:
        seg0 = data.seg(0)
        tbl = data.segment_table("Global", "Value", seg0)
        order = list(tbl[END_YEAR].astype(float).sort_values(
            ascending=False).index)
        got = [str(x) for x in order[:n]]
        if got:
            return got
    except Exception:
        pass
    return [f"Product Type {chr(65 + i)}" for i in range(n)]


def product_types(n: int = 3):
    return _top_subsegments(n)


def end_markets(n: int = 4):
    """Top-``n`` end-markets/applications - from a distribution-channel / end-use
    / application segment when the workbook has one, so it reads as a DIFFERENT
    dimension from the product types.  Falls back to generic labels."""
    kws = ("distribution", "channel", "application", "end use", "end-use",
           "end user", "vertical")
    try:
        for s in data.SEGMENTS:
            if any(k in s.lower() for k in kws):
                subs = list(data.SUBSEGMENTS.get(s, []))[:n]
                if subs:
                    return [str(x) for x in subs]
        # else: use the LAST segment (usually a channel/use split), not seg(0)
        if len(data.SEGMENTS) > 1:
            subs = list(data.SUBSEGMENTS.get(data.SEGMENTS[-1], []))[:n]
            if subs:
                return [str(x) for x in subs]
    except Exception:
        pass
    return [f"End-use {chr(65 + i)}" for i in range(n)]


def shares_by_product(product: str, seed: int):
    """[(company, share%)] for one product type - deterministic variation keyed
    by ``seed`` so each product's donut differs but is stable."""
    names = companies()
    adj = []
    for i, base in enumerate(_BASE_SHARE[:len(names)]):
        wobble = 1.0 + 0.25 * math.cos(i * 1.1 + (seed + 2) * 1.3)
        adj.append(max(0.6, base * wobble))
    others = 24.0 + 4.0 * math.sin((seed + 2) * 1.1)
    scale = (100.0 - others) / (sum(adj) or 1.0)
    out = [(names[i], adj[i] * scale) for i in range(len(names))]
    out.append((OTHERS, others))
    return out


# ── competitor dashboard matrix ──────────────────────────────────────
_STRATEGY_KEYS = ["E", "I", "F", "P", "A", "L", "C"]
STRATEGY_LEGEND = {"E": "Expansion", "I": "R&D Investment", "F": "Financial Aid",
                   "P": "Product Development", "A": "Acquisition",
                   "L": "Product Launch", "C": "Certification"}


def _company_tier(i: int) -> str:
    return "Tier 1" if i < 4 else ("Tier 2" if i < 6 else "Tier 3")


def dashboard_rows(year: int = None):
    """One row per company for the competitor dashboard matrix - brand, revenue,
    share, key regions and strategy codes.  Deterministic and market-pinned.
    ``year`` picks the value base (END_YEAR by default; the Competitive
    Dashboard passes EST_YEAR so it reads as the CURRENT landscape)."""
    mv = market_value("Global", year)
    shs = dict(company_shares("Global"))
    regs = list(data.REGIONS) or ["Global"]
    rows = []
    for i, name in enumerate(companies()):
        share = shs.get(name, 0.0)
        target = mv * share / 100.0
        overall = target * (3.0 + i * 0.45)
        r1 = regs[i % len(regs)]
        r2 = regs[(i + 1) % len(regs)]
        strat = [_STRATEGY_KEYS[j] for j in range(len(_STRATEGY_KEYS))
                 if (i + j) % 2 == 0 or j < 3]
        rows.append({
            "company": name, "tier": _company_tier(i),
            "brand": f"Brand {chr(65 + i)}",
            "overall_rev": overall, "target_rev": target, "share": share,
            "regions": [(r1, "▲"), (r2, "▬")],
            "strategies": strat})
    return rows


# ── creative dashboard matrix (rich visual cells) ────────────────────
STAR_COLORS = list(_T_STAR_COLORS)            # sparkle palette
BOX_COLORS = list(_T_BOX_COLORS)              # product boxes
_ARROW_COLOR = {"high": ARROW_HIGH, "med": ARROW_MED, "low": ARROW_LOW}


def matrix_rows(tier: str | None = None, year: int = None):
    """Rich per-company rows for the creative dashboard matrix.  Adds to
    ``dashboard_rows`` a set of visual encodings: sparkle count, application
    check/cross triple, product colour-box triple, regional presence with
    strengths, a market-share value and active-strategy flags.  Deterministic;
    ``tier`` optionally filters to one tier, ``year`` picks the value base."""
    base = dashboard_rows(year)
    regs = list(data.REGIONS) or ["Global"]
    out = []
    for i, r in enumerate(base):
        if tier and r["tier"] != tier:
            continue
        apps = [((i + k) % 3) != 2 for k in range(3)]     # ✓/✗ triple
        boxes = [((i + k) % 4) != 3 for k in range(3)]    # product boxes on/off
        stars = 2 + (i % 3)                                # 2-4 lit sparkles
        strengths = ["high", "med", "low"]
        pres = [(regs[(i + k) % len(regs)], strengths[k])
                for k in range(min(3, len(regs)))]
        active = set(r["strategies"])
        out.append({
            "company": r["company"], "tier": r["tier"], "brand": r["brand"],
            "overall_rev": r["overall_rev"], "target_rev": r["target_rev"],
            "share": r["share"], "stars": stars, "apps": apps, "boxes": boxes,
            "presence": pres,
            "strategies": [(k, k in active) for k in _STRATEGY_KEYS]})
    return out


def arrow_color(strength: str) -> str:
    return _ARROW_COLOR.get(strength, ARROW_LOW)


# ── per-company strategic profile ────────────────────────────────────
def profile_companies(n: int = 3):
    """Companies that get a full profile (Snapshot / Penetration / Financials).
    Defaults to the top ``n`` by global share."""
    return companies()[:n]


def _idx(company: str) -> int:
    cs = companies()
    return cs.index(company) if company in cs else 0


def profile(company: str) -> dict:
    """A deterministic, market-pinned strategic profile for ``company`` - TAM/
    SAM/SOM, KPIs, SOM splits, regional penetration, win metrics, five-year
    financials and pricing.  All indicative (see INDICATIVE_NOTE)."""
    i = _idx(company)
    mv = market_value("Global")
    share = dict(company_shares("Global")).get(company, 5.0)
    tam = mv
    sam = mv * (0.52 + 0.03 * i)
    som = mv * share / 100.0
    regs = list(data.REGIONS) or ["Global"]
    prods = product_types(4)
    ends = end_markets(4)
    # SOM splits (descending, deterministic)
    def split(keys, seed):
        w = [max(0.4, 1.0 + 0.6 * math.cos(k * 0.9 + seed)) for k in
             range(len(keys))]
        s = sum(w) or 1.0
        return {keys[k]: som * w[k] / s for k in range(len(keys))}
    som_region = split(regs, i + 1)
    som_end = split(ends, i + 2)
    som_product = split(prods, i + 3)
    # regional penetration: SOM vs SAM by region
    penetration = {}
    for k, r in enumerate(regs):
        sam_r = sam * (1.0 + 0.5 * math.cos(k * 0.8)) / len(regs)
        som_r = som_region.get(r, 0.0)
        penetration[r] = {"sam": sam_r, "som": som_r,
                          "pct": (som_r / sam_r * 100.0) if sam_r else 0.0}
    win = {"win_rate": 18.0 + 3.0 * i, "asp_premium": 6.0 - i,
           "pipeline_x": round(2.9 - 0.2 * i, 1),
           "conversion": 30.0 + 4.0 * i, "sales_cycle": 95 - 5 * i}
    # competitive win-position bubbles (Win Rate x vs ASP-vs-Peer y, 0..3.5/0..4)
    win_position = [
        {"name": company, "x": round(0.6 + 0.18 * i, 2), "y": 2.9,
         "color": "#29ABE2", "size": 26},
        {"name": "Industry Avg.", "x": 1.9, "y": 3.0, "color": "#8C8C8C",
         "size": 32},
        {"name": "Market Leader", "x": 2.7, "y": 1.0, "color": ARROW_HIGH,
         "size": 26}]
    # go-to-market benchmarks - target vs actual, on a shared 0..5 index scale
    benchmarks = [
        {"name": "Conversion (%)", "target": round(1.8 + 0.1 * i, 1),
         "actual": round(3.4 + 0.1 * i, 1)},
        {"name": "Sales Cycle (Days)", "target": 4.3, "actual": 2.5},
        {"name": "Pipeline Coverage", "target": 2.4,
         "actual": round(min(4.6, win["pipeline_x"] + 1.5), 1)}]
    # five-year financials (USD Mn), growing
    years = [2020, 2021, 2022, 2023, 2024]
    rev0 = overall = som * (3.0 + i * 0.45)
    rev = [rev0 * (0.78 + 0.06 * y) for y in range(5)]
    gross = [r * (0.34 + 0.005 * i) for r in rev]
    pretax = [r * (0.16 + 0.004 * i) for r in rev]
    net = [r * (0.12 + 0.003 * i) for r in rev]
    fin = {"years": years, "Revenue": rev, "Gross Profit": gross,
           "Pre-Tax Profit": pretax, "Profit After Tax / Net Profit": net}
    ratios = {
        "Gross Margin (%)": [g / r * 100 for g, r in zip(gross, rev)],
        "Pre-Tax Margin (%)": [p / r * 100 for p, r in zip(pretax, rev)],
        "Net Margin (%)": [nn / r * 100 for nn, r in zip(net, rev)],
        "Return on Equity (%)": [12 + i + y for y in range(5)],
        "Return on Assets (%)": [7 + 0.5 * i + 0.4 * y for y in range(5)],
    }
    pricing = {p: (round(4.0 + 0.6 * k + 0.2 * i, 2),
                   round(5.2 + 0.7 * k + 0.25 * i, 2))
               for k, p in enumerate(prods)}
    return {"company": company, "tier": _company_tier(i), "tam": tam,
            "sam": sam, "som": som, "book_to_bill": round(1.0 + 0.04 * i, 2),
            "ebitda": 12.0 + 1.5 * i, "overall_rev": overall,
            "som_region": som_region, "som_end": som_end,
            "som_product": som_product, "penetration": penetration,
            "win": win, "win_position": win_position,
            "benchmarks": benchmarks, "financials": fin, "ratios": ratios,
            "pricing": pricing, "regions": regs, "products": prods,
            "ends": ends, "share": share}


def profile_takeaways(company: str, p: dict) -> list[str]:
    reg_lead = max(p["som_region"], key=p["som_region"].get)
    end_lead = max(p["som_end"], key=p["som_end"].get)
    pen_lead = max(p["penetration"], key=lambda r: p["penetration"][r]["pct"])
    return [
        f"Market Position: {company} captures ~{p['som'] / p['sam'] * 100:.0f}% "
        f"of a {UNIT_LABELS['Value']} SAM, with its strongest end-market in "
        f"'{end_lead}'.",
        f"Geographic Balance: '{reg_lead}' is the largest SOM region; "
        f"'{pen_lead}' shows the highest SAM penetration "
        f"(~{p['penetration'][pen_lead]['pct']:.0f}%).",
        f"Growth Trajectory: a {p['book_to_bill']:.2f} book-to-bill and "
        f"~{p['ebitda']:.0f}% EBITDA margin point to a path to capture more "
        f"of the addressable market.",
        f"Competitive Edge: ~{p['win']['win_rate']:.0f}% win rate with a "
        f"+{p['win']['asp_premium']:.0f}% ASP premium and "
        f"{p['win']['pipeline_x']:.1f}x pipeline coverage.",
    ]


def share_takeaways(region: str):
    """Data-driven Key-Takeaway bullets for the value-share slide, adapting to
    the market name, region and the leading players/concentration."""
    shs = company_shares(region)
    named = [(n, s) for n, s in shs if n != OTHERS]
    named.sort(key=lambda kv: kv[1], reverse=True)
    lead, lead_s = named[0]
    second, second_s = named[1]
    top3 = sum(s for _, s in named[:3])
    others = next((s for n, s in shs if n == OTHERS), 0.0)
    scope = "globally" if region == "Global" else f"in {region}"
    return [
        f"The {MKT} is moderately consolidated {scope}: the top three players "
        f"hold ~{top3:.0f}% of value, while a long tail ('Others') accounts for "
        f"~{others:.0f}%.",
        f"'{lead}' leads with ~{lead_s:.1f}% share, ahead of '{second}' at "
        f"~{second_s:.1f}% the two set the competitive benchmark for pricing "
        f"and innovation.",
        f"Tier 1 players collectively command ~{TIER_SHARE['Tier 1'] * 100:.0f}% "
        f"of the market, underlining meaningful barriers to entry.",
        # The competitive figures are whole-market (the workbook carries no
        # per-region company data), so on a region-scoped deck this goes
        # geography-NEUTRAL rather than claiming to be that region's picture -
        # and the global deck keeps the word it always had.
        f"Regional share differs from the "
        f"{'global' if geo.is_global() else 'overall'} picture use the region "
        f"selector to see how leadership shifts across geographies.",
        f"Share gains {scope} are being driven by capacity expansion, product "
        f"development and selective M&A among the leading players.",
    ]


# -- Template switch support ------------------------------------------
def _retheme() -> None:
    """Re-derive what this module bakes from theme tokens (see theming.py)."""
    global STAR_COLORS, BOX_COLORS, _ARROW_COLOR
    STAR_COLORS = list(_T_STAR_COLORS)
    BOX_COLORS = list(_T_BOX_COLORS)
    _ARROW_COLOR = {"high": ARROW_HIGH, "med": ARROW_MED,
                    "low": ARROW_LOW}
