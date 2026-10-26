"""Competitive-intelligence data for the Competitive Landscape pages.

ONE place that gathers everything the ten competitive-intelligence pages
(templates_ci) read - so when the research API arrives it plugs in HERE
(`_from_api`), and no page template changes.

Today every field comes from the report's own sources:

* the workbook - segment and sub-segment values, the end-use split, regions
* its 'Top 10 Companies' sheet - names, shares, tiers, company type, HQ
* the engine's competitor model (competitors.py) - presence, active
  strategies, portfolio and per-product shares, all indicative
* Market Dynamics (dynamics.py) - drivers and restraints read off the data
* the table of contents - which topics the full study covers

Anything only research can supply - dated events, deal values, people,
patents, named partnerships - is None here and reads 'n/a' on the page, as
the client's reference deck does, until the API fills it.

Scores use one 0-3 scale throughout: 3 strong, 2 moderate, 1 weak, 0 none,
None = not assessed.
"""
from __future__ import annotations

import functools

from . import competitors as C
from . import data
from .config import END_YEAR, EST_YEAR

# strategy code -> (what it reads as, the department that owns it,
#                   the event-type code, what it signals, who should act)
STRATEGY = {
    "A": ("Acquisition to add scale or capability", "Corporate development",
          "MA", "Consolidation: capability is being bought, not built",
          "Corporate development: track targets and deal multiples"),
    "L": ("New product launch", "Product and R&D", "PR",
          "Portfolio widening into adjacent formats",
          "Product: benchmark the launch; Sales: prepare counter-offers"),
    "P": ("Product development programme", "Product and R&D", "PR",
          "Capability build ahead of demand", "Product: watch the roadmap"),
    "E": ("Capacity or geographic expansion", "Operations", "EX",
          "Share push in new regions or volume tiers",
          "Sales: defend key accounts in the regions named"),
    "I": ("R&D investment", "Product and R&D", "RD",
          "Technology bet on next-generation formats",
          "Product and R&D: assess the technology gap"),
    "C": ("Certification or compliance", "Legal and regulatory", "RG",
          "Regulation used as a trust signal",
          "Legal and regulatory: audit own certifications"),
    "F": ("Financial backing / funding", "Finance", "CP",
          "Balance-sheet capacity to invest or acquire",
          "Finance: watch the funding runway"),
}
TYPE_NAMES = {"MA": "M&A", "PR": "Product / portfolio", "EX": "Expansion",
              "RD": "R&D / technology", "RG": "Regulation / compliance",
              "CP": "Capital / funding"}
DEPARTMENTS = ("Sales", "Marketing", "Product and R&D",
               "Corporate development", "Operations", "Legal and regulatory")
# the dominant strategy, in the order it wins when several are active
_PURPOSE = (("A", "Buy to scale"), ("L", "Launch-led growth"),
            ("P", "Build the portfolio"), ("E", "Expand footprint"),
            ("I", "Invest in technology"), ("C", "Certify and comply"),
            ("F", "Funding-backed growth"))
_ROUTE = {"Tier 1": "Direct + key accounts", "Tier 2": "Distributors",
          "Tier 3": "Regional partners", "Others": "Local / online"}
_PRICE = {"Tier 1": "Premium / bundled", "Tier 2": "Mid / value-added",
          "Tier 3": "Value / entry", "Others": "Mixed"}


def _from_api():
    """The research API's competitive-intelligence payload, when it is wired
    in; None today.  Whatever it returns overrides the matching fields."""
    return None


def _main_segments():
    """(product segment, end-use segment) - the first segment of the
    workbook, and the one that names end uses / applications / channels."""
    segs = [s for s in data.SEGMENTS if not data.is_cross(s)]
    if not segs:
        return None, None
    kws = ("end use", "end-use", "application", "channel", "vertical",
           "end user", "distribution")
    end = next((s for s in segs if any(k in s.lower() for k in kws)), None)
    prod = next((s for s in segs if s != end), segs[0])
    return prod, end or (segs[-1] if len(segs) > 1 else None)


def _sub_values(seg):
    """[(sub-segment, EST value, END value)] for `seg`, largest first."""
    if not seg:
        return []
    t = data.segment_table("Global", "Value", seg)
    rows = [(str(i), float(t.loc[i, EST_YEAR]), float(t.loc[i, END_YEAR]))
            for i in t.index]
    return sorted(rows, key=lambda r: r[1], reverse=True)


def _reach(tier, ctype):
    t = (ctype or "").lower()
    if "global" in t or "diversified" in t or "multinational" in t:
        return "Global"
    if "regional" in t:
        return "Regional"
    if "emerging" in t or "local" in t or "niche" in t:
        return "Local or niche"
    return {"Tier 1": "Global", "Tier 2": "Regional"}.get(tier,
                                                          "Local or niche")


def _band(v, cuts):
    """0-3 from a value and three ascending cut-offs."""
    return sum(1 for c in cuts if v >= c)


@functools.lru_cache(maxsize=4)
def build(region: str = "Global") -> dict:
    """Everything the ten pages read (see the module doc)."""
    names = C.companies()
    shares = dict(C.company_shares("Global"))
    tiers = C.tiers()
    tier_of = {n: t for t, v in tiers.items() for n in v["players"]}
    info = C.company_info()
    mrows = {r["company"]: r for r in C.matrix_rows(year=EST_YEAR)}
    prod_seg, end_seg = _main_segments()
    prod_subs, end_subs = _sub_values(prod_seg), _sub_values(end_seg)
    mv0 = sum(r[1] for r in prod_subs) or C.market_value("Global", EST_YEAR)
    mv1 = sum(r[2] for r in prod_subs) or C.market_value("Global", END_YEAR)

    # leadership: the top four sub-segments of each of the two segments; in
    # each, the largest player leads and the next two are active
    cols = [("prod", r[0]) for r in prod_subs[:4]] + \
           [("end", r[0]) for r in end_subs[:4]]
    # the product boxes on the Competitive Dashboard stand for
    # C.product_types(3) and its ticks for the three largest end uses
    box_names = C.product_types(3)
    app_names = [r[0] for r in end_subs[:3]]
    lead = {n: {} for n in names}
    for j, (kind, col) in enumerate(cols):
        sh = sorted(C.shares_by_product(col, j + 1), key=lambda kv: kv[1],
                    reverse=True)
        ranked = [n for n, _v in sh if n in lead]
        for n in lead:
            mr = mrows.get(n, {})
            if kind == "prod" and col in box_names:
                on = mr.get("boxes", [False] * 3)[box_names.index(col)]
            elif kind == "end" and col in app_names:
                on = mr.get("apps", [False] * 3)[app_names.index(col)]
            else:
                on = n in ranked[:3]
            lead[n][col] = 3 if ranked and n == ranked[0] else 2 if on else 0

    companies = []
    for i, n in enumerate(names):
        t = tier_of.get(n, "Others")
        ctype = (info.get(n) or {}).get("type")
        hq = (info.get(n) or {}).get("hq")
        mr = mrows.get(n, {})
        active = [k for k, on in mr.get("strategies", []) if on]
        led = sum(1 for v in lead[n].values() if v == 3)
        act = sum(1 for v in lead[n].values() if v >= 2)
        reach = _reach(t, ctype)
        boxes = mr.get("boxes", [])
        apps = mr.get("apps", [])
        c = {
            "name": n, "tier": t, "share": shares.get(n, 0.0),
            "type": ctype, "hq": hq, "reach": reach,
            "active": active, "led": led, "active_in": act,
            "presence": [r for r, _s in mr.get("presence", [])],
            "purpose": next((p for k, p in _PURPOSE if k in active),
                            "Hold position"),
            "route": _ROUTE.get(t, "n/a"), "price": _PRICE.get(t, "n/a"),
            "verticals": [e for e, ok in zip([r[0] for r in end_subs[:3]],
                                             apps) if ok],
            "products": [p for p, on in zip(C.product_types(3), boxes)
                         if on],
        }
        # ── scores, 0-3 (None = not assessed) ──
        # breadth over the eight sub-segments tracked
        c["s_portfolio"] = 3 if act >= 6 else 2 if act >= 4 else 1 if act >= 2 else 0
        c["s_presence"] = {"Global": 3, "Regional": 2}.get(reach, 1)
        c["s_power"] = {"Tier 1": 3, "Tier 2": 2, "Tier 3": 1}.get(t, 1)
        c["s_performance"] = _band(c["share"], (0.5, 5.0, 10.0))
        c["s_product"] = max(0, min(3, sum(1 for b in boxes if b)))
        c["s_platform"] = (3 if {"I", "P"} <= set(active) else
                           2 if {"I", "P"} & set(active) else 1)
        c["s_pace"] = 3 if len(active) >= 5 else 2 if len(active) >= 3 else 1
        c["pipeline"] = ("Launch pipeline" if "L" in active else
                         "Portfolio development" if "P" in active else None)
        base = (c["s_portfolio"] + c["s_presence"] + c["s_power"]
                + c["s_pace"]) / 4.0
        c["s_prospects"] = 3 if base >= 2.4 else 2 if base >= 1.6 else 1
        companies.append(c)

    # success factors: weights and the score each reads
    factors = [
        ("Segment breadth", 0.30, "s_portfolio",
         "Lead or be active in several sub-segments"),
        ("Scale and share", 0.25, "s_power",
         "Tier position and share of market value"),
        ("Geographic reach", 0.15, "s_presence",
         "Global, regional or single-market footprint"),
        ("Innovation pipeline", 0.15, "s_platform",
         "R&D and product development in motion"),
        ("Strategic momentum", 0.15, "s_pace",
         "Number of active strategic moves"),
    ]
    for c in companies:
        c["composite"] = round(sum(w * c[k] for _n, w, k, _d in factors), 2)

    # strategic activity: one line per active strategy, company by company
    activity = []
    for c in companies:
        top_prod = c["products"][0] if c["products"] else (
            prod_subs[0][0] if prod_subs else "its core line")
        top_reg = c["presence"][0] if c["presence"] else "its home region"
        for k in c["active"]:
            what, dept, code, signal, act = STRATEGY[k]
            where = (f" in {top_reg}" if k in ("E", "A")
                     else f" in {top_prod}" if k in ("L", "P", "I") else "")
            activity.append({"company": c["name"], "tier": c["tier"],
                             "what": what + where, "dept": dept,
                             "code": code, "signal": signal, "act": act,
                             "date": None, "evidence": None})

    toc_topics = []
    try:
        from . import toc
        src = toc.find_toc_file()
        if src is not None:
            toc_topics = [e.text for e in toc.parse_toc(src)]
    except Exception:
        toc_topics = []

    out = {"companies": companies, "leadership": lead, "lead_cols": cols,
           "prod_seg": prod_seg, "end_seg": end_seg,
           "prod_subs": prod_subs, "end_subs": end_subs,
           "mv0": mv0, "mv1": mv1, "tiers": tiers, "factors": factors,
           "activity": activity, "toc_topics": toc_topics}
    api = _from_api()
    if api:
        out.update(api)
    return out


def covered(topic_words, ci=None) -> bool:
    """True when the study's table of contents names the topic."""
    ci = ci or build()
    words = [w.lower() for w in topic_words]
    return any(all(w in t.lower() for w in words) for t in ci["toc_topics"])
